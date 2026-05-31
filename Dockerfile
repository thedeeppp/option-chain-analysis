# Host-agnostic container (Fly.io, Railway, any container host).
FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1 \
    TZ=Asia/Kolkata \
    START_POLLER=1 \
    PORT=8080

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 8080

# Single worker on purpose: one background poller + one in-memory state store.
CMD ["sh", "-c", "gunicorn app:app --workers 1 --threads 4 --timeout 120 --bind 0.0.0.0:$PORT"]
