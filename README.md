# NSE Option-Chain OI Tracker

Local dashboard polling NSE option chain every 3 minutes for **NIFTY**,
**BANKNIFTY**, **FINNIFTY** at the nearest expiry. Classifies each snapshot
as **Long Buildup / Short Covering / Short Buildup / Long Unwinding** based on
**spot direction + PCR direction**.

## Quick start

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python app.py
```

Open <http://127.0.0.1:5000>. Three tabs (NIFTY / BANKNIFTY / FINNIFTY).
Backend polls every 3 minutes starting at 09:18:30 IST. Every 6th cycle adds
a "15m" row that compares against the snapshot 15 minutes ago. Cadence:
`3, 3, 3, 3, 3, +15m, 3, 3, 3, 3, 3, +15m…`

Click **Download Excel** to grab all rows of the day for all three indices in
one .xlsx (one sheet per index). Daily CSVs are also written automatically to
`./data/`.

Six tabs total: `NIFTY-chain / BANKNIFTY-chain / FINNIFTY-chain` (full NSE-style
option chain, CALLS | STRIKE | PUTS, updated in place every cycle) followed by
`NIFTY / BANKNIFTY / FINNIFTY` (the append-only aggregate time series).

## Deploy

This is a stateful, always-on app (a background thread polls NSE every 3 min and
keeps state in memory), so deploy it to a **persistent-process host** — Render,
Railway, or Fly.io. It is **not** suited to serverless (Vercel/Lambda): there is
no always-on process to poll, in-memory state would reset on every cold start,
and the `data/` CSV writes need a writable disk.

Run a **single worker** — multiple workers would each start their own poller and
hit NSE in duplicate. The background thread starts only when `START_POLLER=1`
(set in every start command below), because under gunicorn `main()` never runs.

- **Render**: New → Blueprint, point at this repo (`render.yaml` is included).
- **Railway**: New Project → Deploy from repo. It auto-detects the `Procfile`.
- **Fly.io**: `fly launch` (uses the included `Dockerfile`), then `fly deploy`.
- **Any host / locally**:
  ```bash
  START_POLLER=1 gunicorn app:app --workers 1 --threads 4 --timeout 120 \
    --bind 0.0.0.0:$PORT
  ```

**NSE IP note:** NSE often blocks datacenter IPs (AWS/GCP/Render/Fly), so a cloud
deploy may see 401/403 and empty tables even though it works locally. If that
happens, route NSE calls through a residential/Indian proxy. State is in-memory
and resets on redeploy — fine for an intraday tracker.

## Columns

| Column | Meaning |
|---|---|
| Time | Snapshot time (IST) |
| Expiry | Nearest expiry used (weekly for Nifty, monthly for BN/FN) |
| Spot | Underlying index value |
| Δ Spot, % Δ | Change vs previous snapshot |
| CE OI | Total Call OI across all strikes of the chosen expiry |
| Δ CE OI | Change vs previous snapshot |
| PE OI | Total Put OI |
| Δ PE OI | Change vs previous snapshot |
| PCR | Put-Call Ratio = PE OI / CE OI |
| Δ PCR | Change vs previous snapshot |
| CE Vol Δ | Call volume delta (this snapshot vs previous) |
| CE Vol Total | Cumulative Call volume since market open |
| PE Vol Δ, PE Vol Total | Same for Puts |
| Observation | Long Buildup / Short Covering / Short Buildup / Long Unwinding / Sideways |

## Observation logic

| Spot | PCR | Label | Why |
|---|---|---|---|
| ↑ | ↑ | Long Buildup | Price rising while put-writing dominates — bulls in control |
| ↑ | ↓ | Short Covering | Price rising as call writers exit — squeeze |
| ↓ | ↑ | Short Buildup | Price falling while call-writing dominates — bears in control |
| ↓ | ↓ | Long Unwinding | Price falling as put writers exit |
| ≈0 | ≈0 | Sideways | Small moves on both — no clear bias |

Thresholds: |Δspot| < 0.03 % **and** |ΔPCR| < 0.01 → Sideways.

## Data source

- Primary: `nselib.derivatives.get_func.get_nse_option_chain()` (handles the
  session/cookie dance against `nseindia.com/api/option-chain-v3`).
- Fallback: direct `requests.Session` against the same endpoint, with the
  legacy `option-chain-indices` endpoint used to discover the nearest expiry.

If nselib breaks after an NSE change, the fallback usually still works (or
vice versa). Either way the rest of the pipeline is unchanged.

## Endpoints

| URL | Purpose |
|---|---|
| `/` | Dashboard |
| `/api/data` | JSON of all rows per index |
| `/api/health` | Current status, expiries in use, row counts |
| `/api/manual_fetch` | Force one fetch immediately |
| `/download/xlsx` | Download today's data as Excel |

## Caveats

1. **NSE has no official public API.** The endpoints we use are reverse-
   engineered. They require browser-like session cookies (handled
   automatically) and occasionally change. Expect to update `HEADERS` or
   upgrade `nselib` once or twice a year.
2. **FINNIFTY weekly expiries were discontinued** in November 2024 by SEBI.
   Only monthly expiries exist for FINNIFTY and BANKNIFTY. NIFTY still has
   weekly + monthly. The code automatically picks whatever is nearest.
3. **FINNIFTY has low intraday volume.** PCR can be noisy minute-to-minute,
   especially on the monthly expiry between expiry weeks. This is real, not
   a bug.
4. **NSE rate-limiting.** ~3 req/sec is the rough cap. Three indices every
   3 minutes is well within that.

## Files

- `app.py` — fetcher + scheduler + Flask routes
- `templates/index.html` — UI
- `data/<INDEX>_<YYYYMMDD>.csv` — daily logs (auto-written)
