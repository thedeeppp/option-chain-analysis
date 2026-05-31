"""
NSE Option-Chain OI Tracker
---------------------------
Polls NSE option-chain every 3 minutes (with a +15-minute summary row every
6th cycle) for NIFTY, BANKNIFTY, FINNIFTY at the *nearest* expiry. Classifies
each snapshot as Long Buildup / Short Covering / Short Buildup / Long
Unwinding based on (spot direction) + (PCR direction).

Primary data source: nselib (handles cookies/headers).
Fallback:            direct requests against /api/option-chain-v3.

Run:   python app.py
Open:  http://127.0.0.1:5000
"""

import csv
import io
import logging
import os
import threading
import time
from collections import deque
from datetime import datetime, time as dtime
from zoneinfo import ZoneInfo

import requests
from flask import Flask, jsonify, render_template, send_file

try:
    from nselib.derivatives.get_func import get_nse_option_chain as _nselib_fetch
    HAS_NSELIB = True
except Exception:
    HAS_NSELIB = False

import openpyxl

# -----------------------------------------------------------------------------
# Config
# -----------------------------------------------------------------------------
IST = ZoneInfo("Asia/Kolkata")
INDICES = ["NIFTY", "BANKNIFTY", "FINNIFTY"]

# Strike-price step per index (used to derive the ATM strike from spot on the
# aggregate pages). NIFTY/FINNIFTY trade in 50-pt strikes, BANKNIFTY in 100.
STRIKE_STEP = {"NIFTY": 50, "BANKNIFTY": 100, "FINNIFTY": 50}

MARKET_OPEN = dtime(9, 15)
MARKET_CLOSE = dtime(15, 30)
FIRST_FETCH = dtime(9, 18, 30)

POLL_INTERVAL_SEC = 180
BIG_WINDOW_CYCLE = 5  # every 6th cycle (0-indexed 5) also emit a 15m row
MAX_ROWS_PER_INDEX = 300

DATA_DIR = "data"
os.makedirs(DATA_DIR, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
log = logging.getLogger("oi_tracker")

# -----------------------------------------------------------------------------
# Direct-fetch fallback (used if nselib breaks)
# -----------------------------------------------------------------------------
NSE_HOME = "https://www.nseindia.com"
NSE_OPT_PAGE = "https://www.nseindia.com/option-chain"
NSE_OPT_API_V3 = (
    "https://www.nseindia.com/api/option-chain-v3"
    "?type=Indices&symbol={symbol}&expiry={expiry}"
)
NSE_OPT_API_LEGACY = (
    "https://www.nseindia.com/api/option-chain-indices?symbol={symbol}"
)

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": NSE_OPT_PAGE,
}


class DirectClient:
    """Used only if nselib's fetcher fails."""

    def __init__(self):
        self.session = requests.Session()
        self.session.headers.update(HEADERS)
        self._last_refresh = 0.0

    def _refresh(self):
        try:
            self.session.get(NSE_HOME, timeout=10)
            self.session.get(NSE_OPT_PAGE, timeout=10)
            self._last_refresh = time.time()
        except Exception as e:
            log.warning(f"Direct cookie refresh failed: {e}")

    def fetch(self, symbol: str, expiry: str = "") -> dict | None:
        if time.time() - self._last_refresh > 600:
            self._refresh()

        # Prefer v3 if we know expiry; otherwise legacy (returns all expiries)
        import urllib.parse
        if expiry:
            url = (
                f"https://www.nseindia.com/api/option-chain-v3"
                f"?type=Indices&symbol={symbol}"
                f"&expiry={urllib.parse.quote(expiry)}"
            )
        else:
            url = (
                f"https://www.nseindia.com/api/option-chain-v3"
                f"?type=Indices&symbol={symbol}"
            )

        for attempt in (1, 2):
            try:
                r = self.session.get(url, timeout=12)
                if r.status_code == 200 and r.text.strip().startswith("{"):
                    return r.json()
                log.warning(
                    f"Direct {symbol}: HTTP {r.status_code} attempt {attempt}"
                )
            except Exception as e:
                log.warning(f"Direct {symbol} attempt {attempt}: {e}")
            self._refresh()
        return None


DIRECT = DirectClient()


def fetch_option_chain(symbol: str, expiry: str = "") -> dict | None:
    """
    Try nselib first; if it fails, fall back to direct requests.
    Returns the raw JSON dict.
    """
    # Try nselib (only useful if we have an expiry to pass — v3 requires it)
    if HAS_NSELIB and expiry:
        try:
            resp = _nselib_fetch(symbol, expiry)
            if resp is not None and resp.status_code == 200:
                return resp.json()
        except Exception as e:
            log.warning(f"nselib fetch failed for {symbol}: {e}")

    # Fallback
    return DIRECT.fetch(symbol, expiry)


def _compute_nearest_tuesday(today: datetime, monthly_only: bool) -> str:
    """
    Fallback: compute the nearest NSE expiry by date math.
    All NSE index expiries are on Tuesday (since Sept 1, 2025).
    Weekly  -> next Tuesday on/after today
    Monthly -> last Tuesday of the current month (or next month if already past)
    """
    from datetime import timedelta
    import calendar

    def last_tuesday_of(year: int, month: int) -> datetime:
        # Find last Tuesday: walk back from the last day
        last_day = calendar.monthrange(year, month)[1]
        d = today.replace(year=year, month=month, day=last_day)
        # weekday(): Mon=0, Tue=1
        while d.weekday() != 1:
            d -= timedelta(days=1)
        return d

    if monthly_only:
        cand = last_tuesday_of(today.year, today.month)
        if cand.date() < today.date():
            # Roll to next month
            nm = today.month + 1
            ny = today.year + (1 if nm > 12 else 0)
            nm = 1 if nm > 12 else nm
            cand = last_tuesday_of(ny, nm)
    else:
        # Next Tuesday >= today
        days_ahead = (1 - today.weekday()) % 7
        cand = today + timedelta(days=days_ahead)

    return cand.strftime("%d-%b-%Y")


def fetch_nearest_expiry(symbol: str) -> str | None:
    """
    Try several strategies in order:
      1. v3 with empty expiry — sometimes returns full expiryDates list
      2. v3 with a synthetic next-Tuesday — payload has expiryDates regardless
      3. Pure date math (always works, no network)
    """
    today = datetime.now(IST)
    monthly_only = (symbol != "NIFTY")  # BANKNIFTY/FINNIFTY: monthly only

    # Strategy 1: try v3 with no expiry filter
    try:
        url = (
            f"https://www.nseindia.com/api/option-chain-v3"
            f"?type=Indices&symbol={symbol}"
        )
        if time.time() - DIRECT._last_refresh > 600:
            DIRECT._refresh()
        r = DIRECT.session.get(url, timeout=12)
        if r.status_code == 200 and r.text.strip().startswith("{"):
            payload = r.json()
            exps = payload.get("records", {}).get("expiryDates") or []
            if exps:
                log.info(f"{symbol} expiry discovered via v3: {exps[0]}")
                return exps[0]
    except Exception as e:
        log.warning(f"v3 discovery failed for {symbol}: {e}")

    # Strategy 2: date math fallback
    computed = _compute_nearest_tuesday(today, monthly_only)
    log.info(f"{symbol} expiry computed via date math: {computed}")
    return computed


# -----------------------------------------------------------------------------
# Parse: extract totals across all strikes of the chosen expiry
# -----------------------------------------------------------------------------
def parse_chain(payload: dict, expiry: str) -> dict | None:
    """
    Walk records.data, sum CE and PE OI / volume across all strikes whose
    expiryDate == expiry. Also pulls the spot (underlyingValue).
    """

    if not payload:
        return None
    records = payload.get("records") or {}
    data = records.get("data") or []
    spot = records.get("underlyingValue")
    # v3 sometimes nests the active-expiry data under "filtered" instead
    if not data:
        filtered = payload.get("filtered") or {}
        data = filtered.get("data") or []
        if spot is None:
            spot = filtered.get("underlyingValue")
    if spot is None:
        return None

    ce_oi = ce_chg = ce_vol = 0
    pe_oi = pe_chg = pe_vol = 0
    ce_price_chg = pe_price_chg = 0.0
    strike_count = 0

    for row in data:
        if row.get("expiryDates") != expiry:
            continue
        strike_count += 1
        ce = row.get("CE") or {}
        pe = row.get("PE") or {}
        ce_oi += int(ce.get("openInterest") or 0)
        ce_chg += int(ce.get("changeinOpenInterest") or 0)
        ce_vol += int(ce.get("totalTradedVolume") or 0)
        ce_price_chg += float(ce.get("change") or 0)
        pe_oi += int(pe.get("openInterest") or 0)
        pe_chg += int(pe.get("changeinOpenInterest") or 0)
        pe_vol += int(pe.get("totalTradedVolume") or 0)
        pe_price_chg += float(pe.get("change") or 0)

    if strike_count == 0:
        return None

    pcr = (pe_oi / ce_oi) if ce_oi > 0 else 0.0
    return {
        "expiry": expiry,
        "spot": float(spot),
        "ce_oi": ce_oi,
        "ce_oi_chg_session": ce_chg,  # NSE's day-cumulative OI change (summed)
        "ce_price_chg": ce_price_chg,  # NSE's day-cumulative premium change (summed)
        "ce_vol_total": ce_vol,
        "pe_oi": pe_oi,
        "pe_oi_chg_session": pe_chg,
        "pe_price_chg": pe_price_chg,
        "pe_vol_total": pe_vol,
        "pcr": round(pcr, 3),
        "strike_count": strike_count,
    }


# Number of strikes shown above and below the ATM strike on the chain pages.
CHAIN_STRIKES_EACH_SIDE = 10


def _leg(side: dict) -> dict:
    """Flatten one NSE CE/PE leg into the 10 display fields (raw values)."""
    def num(v):
        return v if isinstance(v, (int, float)) else 0
    return {
        "oi": int(num(side.get("openInterest"))),
        "chng_oi": int(num(side.get("changeinOpenInterest"))),
        "vol": int(num(side.get("totalTradedVolume"))),
        "iv": float(num(side.get("impliedVolatility"))),
        "ltp": float(num(side.get("lastPrice"))),
        "chng": float(num(side.get("change"))),
        "bid_qty": int(num(side.get("buyQuantity1"))),
        "bid": float(num(side.get("buyPrice1"))),
        "ask": float(num(side.get("sellPrice1"))),
        "ask_qty": int(num(side.get("sellQuantity1"))),
    }


def parse_chain_strikes(payload: dict, expiry: str) -> dict | None:
    """
    Build the per-strike option-chain snapshot for the chain pages.

    Reuses the same payload `parse_chain` consumes (fetched once per cycle) and
    keeps every per-leg field NSE returns. Windows the strike ladder to
    ±CHAIN_STRIKES_EACH_SIDE rows around the ATM strike (strike nearest spot),
    so the window adapts to each index's strike step (50 vs 100) automatically.
    """
    if not payload:
        return None
    records = payload.get("records") or {}
    data = records.get("data") or []
    spot = records.get("underlyingValue")
    if not data:
        filtered = payload.get("filtered") or {}
        data = filtered.get("data") or []
        if spot is None:
            spot = filtered.get("underlyingValue")
    if spot is None:
        return None
    spot = float(spot)

    by_strike: dict[float, dict] = {}
    for row in data:
        if row.get("expiryDates") != expiry:
            continue
        strike = row.get("strikePrice")
        if strike is None:
            continue
        ce = _leg(row.get("CE") or {})
        pe = _leg(row.get("PE") or {})
        by_strike[float(strike)] = {
            "strike": float(strike),
            "ce": ce,
            "pe": pe,
            "ce_obs": classify_strike(ce["chng"], ce["chng_oi"]),
            "pe_obs": classify_strike(pe["chng"], pe["chng_oi"]),
        }

    if not by_strike:
        return None

    strikes = sorted(by_strike)
    atm = min(strikes, key=lambda s: abs(s - spot))
    atm_idx = strikes.index(atm)
    lo = max(0, atm_idx - CHAIN_STRIKES_EACH_SIDE)
    hi = min(len(strikes), atm_idx + CHAIN_STRIKES_EACH_SIDE + 1)
    window = [by_strike[s] for s in strikes[lo:hi]]

    return {
        "expiry": expiry,
        "spot": spot,
        "atm_strike": atm,
        "timestamp": datetime.now(IST).strftime("%H:%M:%S"),
        "rows": window,
    }


# -----------------------------------------------------------------------------
# Observation logic
# -----------------------------------------------------------------------------
def classify(spot_pct: float, pcr_change: float) -> str:
    """
    Map (spot direction, PCR direction) -> trader sentiment label.

       Spot ↑  +  PCR ↑   -> Long Buildup
       Spot ↑  +  PCR ↓   -> Short Covering
       Spot ↓  +  PCR ↑   -> Short Buildup
       Spot ↓  +  PCR ↓   -> Long Unwinding
    """
    EPS_SPOT = 0.03    # %
    EPS_PCR = 0.01     # absolute change in PCR ratio

    if abs(spot_pct) < EPS_SPOT and abs(pcr_change) < EPS_PCR:
        return "Sideways / No clear bias"

    spot_up = spot_pct > 0
    pcr_up = pcr_change > 0

    if spot_up and pcr_up:
        return "Long Buildup"
    if spot_up and not pcr_up:
        return "Short Covering"
    if not spot_up and pcr_up:
        return "Short Buildup"
    return "Long Unwinding"


def classify_strike(price_chg: float, oi_chg: float) -> str:
    """
    Per-strike (per-leg) sentiment from that option's price change + OI change.
    This is the standard intraday OI interpretation applied to a single
    instrument (works identically for a CE leg or a PE leg):

       Price ↑  +  OI ↑   -> Long Buildup     (fresh longs)
       Price ↓  +  OI ↑   -> Short Buildup    (fresh shorts)
       Price ↑  +  OI ↓   -> Short Covering   (shorts exiting)
       Price ↓  +  OI ↓   -> Long Unwinding   (longs exiting)

    `price_chg` is NSE's `change` (LTP change vs prev close); `oi_chg` is
    NSE's `changeinOpenInterest` (OI change vs prev close).
    """
    EPS_PRICE = 0.01   # rupees
    EPS_OI = 1         # contracts/lots

    if abs(price_chg) < EPS_PRICE and abs(oi_chg) < EPS_OI:
        return "—"

    price_up = price_chg > 0
    oi_up = oi_chg > 0

    if price_up and oi_up:
        return "Long Buildup"
    if not price_up and oi_up:
        return "Short Buildup"
    if price_up and not oi_up:
        return "Short Covering"
    return "Long Unwinding"


# -----------------------------------------------------------------------------
# State store
# -----------------------------------------------------------------------------
class IndexStore:
    def __init__(self, name: str):
        self.name = name
        self.rows = deque(maxlen=MAX_ROWS_PER_INDEX)
        self.lock = threading.Lock()
        self.csv_path = os.path.join(
            DATA_DIR, f"{name}_{datetime.now(IST).strftime('%Y%m%d')}.csv"
        )
        if not os.path.exists(self.csv_path):
            with open(self.csv_path, "w", newline="") as f:
                csv.writer(f).writerow(CSV_HEADER)

    def add(self, row: dict):
        with self.lock:
            self.rows.append(row)
        with open(self.csv_path, "a", newline="") as f:
            csv.writer(f).writerow([row.get(k, "") for k in CSV_HEADER])

    def snapshot_3m(self) -> list[dict]:
        with self.lock:
            return [r for r in self.rows if r["kind"] == "3m"]

    def snapshot_all(self) -> list[dict]:
        with self.lock:
            return list(self.rows)[::-1]  # newest first for UI


CSV_HEADER = [
    "time_frame", "kind", "index", "expiry",
    "spot", "spot_change", "spot_pct", "atm_strike",
    "ce_oi", "ce_oi_change",
    "pe_oi", "pe_oi_change",
    "pcr", "pcr_change",
    "ce_vol_delta", "ce_vol_total",
    "pe_vol_delta", "pe_vol_total",
    "observation", "ce_observation", "pe_observation",
]

STORES: dict[str, IndexStore] = {name: IndexStore(name) for name in INDICES}
EXPIRY_CACHE: dict[str, str] = {}     # symbol -> "29-May-2026"
EXPIRY_CACHED_AT: dict[str, float] = {}

# Latest per-strike chain snapshot per index (overwritten in place each cycle).
LATEST_CHAIN: dict[str, dict] = {}
LATEST_CHAIN_LOCK = threading.Lock()


def get_expiry(symbol: str) -> str | None:
    """Cache nearest expiry per symbol; refresh once a day."""
    now = time.time()
    cached_at = EXPIRY_CACHED_AT.get(symbol, 0)
    if symbol in EXPIRY_CACHE and (now - cached_at) < 6 * 3600:
        return EXPIRY_CACHE[symbol]
    exp = fetch_nearest_expiry(symbol)
    if exp:
        EXPIRY_CACHE[symbol] = exp
        EXPIRY_CACHED_AT[symbol] = now
        log.info(f"{symbol} nearest expiry: {exp}")
    return exp


# -----------------------------------------------------------------------------
# Polling
# -----------------------------------------------------------------------------
def is_market_open(now: datetime) -> bool:
    if now.weekday() >= 5:
        return False
    return MARKET_OPEN <= now.time() <= MARKET_CLOSE


def seconds_until(target: dtime, now: datetime) -> float:
    target_dt = now.replace(
        hour=target.hour, minute=target.minute,
        second=target.second, microsecond=0,
    )
    return (target_dt - now).total_seconds()


def build_row(symbol: str, parsed: dict, ref: dict | None,
              kind: str, ts: str) -> dict:
    expiry = parsed["expiry"]
    spot = parsed["spot"]
    ce_oi = parsed["ce_oi"]
    pe_oi = parsed["pe_oi"]
    pcr = parsed["pcr"]
    ce_vol_total = parsed["ce_vol_total"]
    pe_vol_total = parsed["pe_vol_total"]

    step = STRIKE_STEP.get(symbol, 50)
    atm_strike = int(round(spot / step) * step) if spot else ""

    # Per-side observation, derived from the SAME day-cumulative NSE fields the
    # chain pages use per strike (premium change + OI change), summed across each
    # side. This makes the aggregate label the OI-summed counterpart of the
    # chain's per-strike labels at the same timestamp — consistent by
    # construction, and independent of the snapshot history (`ref`).
    ce_observation = classify_strike(
        parsed.get("ce_price_chg", 0.0), parsed.get("ce_oi_chg_session", 0)
    )
    pe_observation = classify_strike(
        parsed.get("pe_price_chg", 0.0), parsed.get("pe_oi_chg_session", 0)
    )

    if ref is None:
        return {
            "time_frame": ts, "kind": kind, "index": symbol, "expiry": expiry,
            "spot": spot, "spot_change": "", "spot_pct": "",
            "atm_strike": atm_strike,
            "ce_oi": ce_oi, "ce_oi_change": "",
            "pe_oi": pe_oi, "pe_oi_change": "",
            "pcr": pcr, "pcr_change": "",
            "ce_vol_delta": "", "ce_vol_total": ce_vol_total,
            "pe_vol_delta": "", "pe_vol_total": pe_vol_total,
            "observation": "— (first reading)",
            "ce_observation": ce_observation, "pe_observation": pe_observation,
        }

    spot_change = spot - ref["spot"]
    spot_pct = (spot_change / ref["spot"] * 100) if ref["spot"] else 0
    ce_oi_change = ce_oi - ref["ce_oi"]
    pe_oi_change = pe_oi - ref["pe_oi"]
    pcr_change = pcr - ref["pcr"]
    ce_vol_delta = ce_vol_total - ref["ce_vol_total"]
    pe_vol_delta = pe_vol_total - ref["pe_vol_total"]

    observation = classify(spot_pct, pcr_change)

    return {
        "time_frame": ts, "kind": kind, "index": symbol, "expiry": expiry,
        "spot": spot,
        "spot_change": round(spot_change, 2),
        "spot_pct": round(spot_pct, 3),
        "atm_strike": atm_strike,
        "ce_oi": ce_oi, "ce_oi_change": ce_oi_change,
        "pe_oi": pe_oi, "pe_oi_change": pe_oi_change,
        "pcr": pcr, "pcr_change": round(pcr_change, 3),
        "ce_vol_delta": ce_vol_delta, "ce_vol_total": ce_vol_total,
        "pe_vol_delta": pe_vol_delta, "pe_vol_total": pe_vol_total,
        "observation": observation,
        "ce_observation": ce_observation,
        "pe_observation": pe_observation,
    }


def poll_all(big_window: bool):
    """Fetch all three indices; for each, append a 3m row (+ optional 15m row)."""
    now = datetime.now(IST)
    ts = now.strftime("%H:%M:%S")

    for symbol in INDICES:
        expiry = get_expiry(symbol)
        if not expiry:
            log.error(f"{symbol}: no expiry resolved, skipping")
            continue

        payload = fetch_option_chain(symbol, expiry)
        parsed = parse_chain(payload, expiry)
        if not parsed:
            log.error(f"{symbol}: parse failed, skipping")
            continue

        # Same payload → per-strike chain snapshot (no extra NSE request).
        chain = parse_chain_strikes(payload, expiry)
        if chain:
            with LATEST_CHAIN_LOCK:
                LATEST_CHAIN[symbol] = chain

        store = STORES[symbol]
        prev_3m = store.snapshot_3m()  # in insertion order (oldest first)

        # 3-min row — compare with previous 3m row
        ref_3m = prev_3m[-1] if prev_3m else None
        row_3m = build_row(symbol, parsed, ref_3m, "3m", ts)
        store.add(row_3m)
        log.info(
            f"[{symbol}] 3m {ts} spot={parsed['spot']} "
            f"PCR={parsed['pcr']} -> {row_3m['observation']}"
        )

        # 15-min row — compare with the 3m row 5 slots back (i.e., ~15 min ago)
        if big_window:
            ref_15m = prev_3m[-5] if len(prev_3m) >= 5 else (
                prev_3m[0] if prev_3m else None
            )
            row_15m = build_row(symbol, parsed, ref_15m, "15m", ts + " (15m)")
            store.add(row_15m)
            log.info(
                f"[{symbol}] 15m {ts} -> {row_15m['observation']}"
            )


def scheduler_loop():
    log.info("Scheduler thread started")
    counter = 0
    while True:
        now = datetime.now(IST)

        if not is_market_open(now):
            log.info("Market closed; sleeping 60s")
            time.sleep(60)
            continue

        if now.time() < FIRST_FETCH:
            wait = max(1, seconds_until(FIRST_FETCH, now))
            log.info(f"Waiting {wait:.0f}s until first fetch at {FIRST_FETCH}")
            time.sleep(wait)
            continue

        big_window = (counter > 0) and (counter % 6 == BIG_WINDOW_CYCLE)
        try:
            poll_all(big_window=big_window)
        except Exception as e:
            log.exception(f"poll_all failed: {e}")

        counter += 1
        time.sleep(POLL_INTERVAL_SEC)


# -----------------------------------------------------------------------------
# Flask app
# -----------------------------------------------------------------------------
app = Flask(__name__)


@app.route("/")
def index():
    return render_template("index.html", indices=INDICES)


@app.route("/api/data")
def api_data():
    out = {n: STORES[n].snapshot_all() for n in INDICES}
    return jsonify(out)


@app.route("/api/chain")
def api_chain():
    with LATEST_CHAIN_LOCK:
        out = {n: LATEST_CHAIN.get(n) for n in INDICES}
    return jsonify(out)


@app.route("/api/health")
def health():
    now = datetime.now(IST)
    return jsonify({
        "now_ist": now.isoformat(),
        "market_open": is_market_open(now),
        "rows": {n: len(STORES[n].rows) for n in INDICES},
        "expiries": EXPIRY_CACHE,
        "using_nselib": HAS_NSELIB,
    })


@app.route("/api/manual_fetch")
def manual_fetch():
    """Force one fetch — handy outside market hours."""
    poll_all(big_window=False)
    return jsonify({"ok": True})


@app.route("/download/xlsx")
def download_xlsx():
    """Excel with 3 sheets (one per index), all rows of the day."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for sym in INDICES:
        ws = wb.create_sheet(title=sym)
        ws.append(CSV_HEADER)
        # newest-first → write oldest-first for natural reading
        with STORES[sym].lock:
            rows = list(STORES[sym].rows)
        for r in rows:
            ws.append([r.get(k, "") for k in CSV_HEADER])
        # Auto-size first few columns
        for col_idx, col_name in enumerate(CSV_HEADER, start=1):
            ws.column_dimensions[
                openpyxl.utils.get_column_letter(col_idx)
            ].width = max(12, len(col_name) + 2)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    fname = f"oi_tracker_{datetime.now(IST).strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(
        buf,
        as_attachment=True,
        download_name=fname,
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )


# -----------------------------------------------------------------------------
# Boot
# -----------------------------------------------------------------------------
def demo_loop():
    """Inject fake snapshots every 5 seconds so the UI is exercisable."""
    import random
    log.info("DEMO MODE: synthetic data every 5 seconds")

    # Seed expiries so the UI shows them
    EXPIRY_CACHE.update({
        "NIFTY":     "27-May-2026",
        "BANKNIFTY": "29-May-2026",
        "FINNIFTY":  "29-May-2026",
    })

    base = {
        "NIFTY":     {"spot": 25100.0, "ce_oi": 12_000_000, "pe_oi": 13_500_000,
                      "ce_vol": 0, "pe_vol": 0, "step": 50},
        "BANKNIFTY": {"spot": 55200.0, "ce_oi": 3_500_000,  "pe_oi": 3_800_000,
                      "ce_vol": 0, "pe_vol": 0, "step": 100},
        "FINNIFTY":  {"spot": 26350.0, "ce_oi":   850_000,  "pe_oi":   900_000,
                      "ce_vol": 0, "pe_vol": 0, "step": 50},
    }

    def synth_chain(sym, spot, step):
        """Fake a per-strike chain window around ATM so chain pages render."""
        atm = round(spot / step) * step
        rows = []
        for i in range(-CHAIN_STRIKES_EACH_SIDE, CHAIN_STRIKES_EACH_SIDE + 1):
            strike = atm + i * step
            # Crude intrinsic + time value so LTP looks plausible.
            ce_intrinsic = max(0.0, spot - strike)
            pe_intrinsic = max(0.0, strike - spot)
            tv = max(2.0, 40.0 - abs(i) * 3.0)
            ce_ltp = round(ce_intrinsic + tv + random.uniform(-3, 3), 2)
            pe_ltp = round(pe_intrinsic + tv + random.uniform(-3, 3), 2)
            ce_chg = round(random.uniform(-25, 25), 2)
            pe_chg = round(random.uniform(-25, 25), 2)
            ce_chg_oi = random.randint(-50_000, 80_000)
            pe_chg_oi = random.randint(-50_000, 80_000)
            ce_leg = {
                "oi": random.randint(1_000, 300_000), "chng_oi": ce_chg_oi,
                "vol": random.randint(0, 5_000_000),
                "iv": round(random.uniform(8, 38), 2),
                "ltp": max(0.05, ce_ltp), "chng": ce_chg,
                "bid_qty": random.randint(50, 5_000),
                "bid": max(0.05, round(ce_ltp - 0.2, 2)),
                "ask": round(ce_ltp + 0.2, 2),
                "ask_qty": random.randint(50, 5_000),
            }
            pe_leg = {
                "oi": random.randint(1_000, 300_000), "chng_oi": pe_chg_oi,
                "vol": random.randint(0, 5_000_000),
                "iv": round(random.uniform(8, 38), 2),
                "ltp": max(0.05, pe_ltp), "chng": pe_chg,
                "bid_qty": random.randint(50, 5_000),
                "bid": max(0.05, round(pe_ltp - 0.2, 2)),
                "ask": round(pe_ltp + 0.2, 2),
                "ask_qty": random.randint(50, 5_000),
            }
            rows.append({
                "strike": float(strike), "ce": ce_leg, "pe": pe_leg,
                "ce_obs": classify_strike(ce_chg, ce_chg_oi),
                "pe_obs": classify_strike(pe_chg, pe_chg_oi),
            })
        return {
            "expiry": EXPIRY_CACHE[sym], "spot": round(spot, 2),
            "atm_strike": float(atm),
            "timestamp": datetime.now(IST).strftime("%H:%M:%S"),
            "rows": rows,
        }

    counter = 0
    while True:
        ts = datetime.now(IST).strftime("%H:%M:%S")
        for sym in INDICES:
            b = base[sym]
            # Random walk on spot (±0.15%)
            b["spot"] *= 1 + random.uniform(-0.0015, 0.0015)
            # Random walk on OI (±0.5%)
            b["ce_oi"] = int(b["ce_oi"] * (1 + random.uniform(-0.005, 0.005)))
            b["pe_oi"] = int(b["pe_oi"] * (1 + random.uniform(-0.005, 0.005)))
            # Volume just accumulates
            b["ce_vol"] += random.randint(5_000, 50_000)
            b["pe_vol"] += random.randint(5_000, 50_000)

            # Build the chain first, then derive the aggregate day-change sums
            # from the SAME synthesized rows — so the per-side observation on the
            # aggregate page is the sum of the chain's per-strike inputs (the two
            # pages stay consistent in demo, exactly as they do with live data).
            chain = synth_chain(sym, b["spot"], b["step"])
            with LATEST_CHAIN_LOCK:
                LATEST_CHAIN[sym] = chain

            pcr = b["pe_oi"] / b["ce_oi"]
            parsed = {
                "expiry": EXPIRY_CACHE[sym],
                "spot": round(b["spot"], 2),
                "ce_oi": b["ce_oi"],
                "ce_oi_chg_session": sum(r["ce"]["chng_oi"] for r in chain["rows"]),
                "ce_price_chg": sum(r["ce"]["chng"] for r in chain["rows"]),
                "ce_vol_total": b["ce_vol"],
                "pe_oi": b["pe_oi"],
                "pe_oi_chg_session": sum(r["pe"]["chng_oi"] for r in chain["rows"]),
                "pe_price_chg": sum(r["pe"]["chng"] for r in chain["rows"]),
                "pe_vol_total": b["pe_vol"],
                "pcr": round(pcr, 3),
                "strike_count": 30,
            }

            store = STORES[sym]
            prev_3m = store.snapshot_3m()
            ref_3m = prev_3m[-1] if prev_3m else None
            row_3m = build_row(sym, parsed, ref_3m, "3m", ts)
            store.add(row_3m)

            # Trigger 15m row every 6th cycle (just like real cadence)
            big_window = (counter > 0) and (counter % 6 == BIG_WINDOW_CYCLE)
            if big_window:
                ref_15m = prev_3m[-5] if len(prev_3m) >= 5 else (
                    prev_3m[0] if prev_3m else None
                )
                row_15m = build_row(sym, parsed, ref_15m, "15m", ts + " (15m)")
                store.add(row_15m)

        counter += 1
        time.sleep(5)  # 5-second cycles instead of 3 minutes


_worker_lock = threading.Lock()
_worker_started = False


def ensure_worker_started(demo: bool = False):
    """
    Start the background polling thread exactly once.

    Local dev (`python app.py`) calls this from main(). Under a WSGI server like
    gunicorn, main() never runs, so the production start command sets
    START_POLLER=1 and this is invoked at import time instead. The lock + flag
    make it idempotent, so it is safe to call from both paths. IMPORTANT: run
    gunicorn with a single worker (`--workers 1`) — multiple workers would each
    start their own poller and hit NSE in duplicate.
    """
    global _worker_started
    with _worker_lock:
        if _worker_started:
            return
        target = demo_loop if demo else scheduler_loop
        threading.Thread(target=target, daemon=True).start()
        _worker_started = True
        log.info(f"Background worker started (demo={demo})")


# When served by a WSGI server (gunicorn), main() is never called — start the
# poller at import time if the start command opts in via START_POLLER=1.
if os.environ.get("START_POLLER") == "1":
    log.info(f"Using nselib: {HAS_NSELIB} | START_POLLER=1 (WSGI mode)")
    ensure_worker_started(demo=False)


def main():
    import sys
    demo = "--demo" in sys.argv

    log.info(f"Using nselib: {HAS_NSELIB} | demo mode: {demo}")
    ensure_worker_started(demo=demo)

    host = os.environ.get("HOST", "127.0.0.1")
    port = int(os.environ.get("PORT", 5000))
    app.run(host=host, port=port, debug=False, use_reloader=False)


if __name__ == "__main__":
    main()