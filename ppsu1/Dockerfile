# Backend image for Render (free tier). Build context: the ppsu1/ folder.
#
# Python is PINNED to 3.11: resvg-py (auto-icons) ships manylinux wheels only
# up to cp312, and this venv/codebase is developed on 3.11. On a Python with
# no wheel the engine silently loses icons — don't bump without checking PyPI.
FROM python:3.11-slim

# LibreOffice renders finished decks to preview PNGs (pptx -> PDF -> PyMuPDF).
# Carlito/Caladea are metric-compatible with Calibri/Cambria — the fonts this
# engine writes — so LibreOffice previews line-break the way PowerPoint would.
RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-impress \
        libreoffice-core \
        fonts-crosextra-carlito \
        fonts-crosextra-caladea \
        fonts-liberation \
        fonts-dejavu \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY backend/requirements.txt backend/requirements.txt
RUN pip install --no-cache-dir -r backend/requirements.txt

# Ships the bundled template AND the pre-computed caches (catalog JSON, design
# thumbnails, icon PNGs). Rebuilding those on a 512 MB / 0.1 vCPU instance
# would OOM (catalog build of the 50 MB template peaks at ~840 MB) or take
# many minutes (302 LibreOffice slide renders), so they must come in the image.
COPY . .

ENV PYTHONUNBUFFERED=1
WORKDIR /app/backend

# One worker, but worker-class MUST be gthread: with the default "sync" class
# --threads is silently ignored and the ONE worker fully blocks on a /build,
# so nothing else (not even /) answers until it finishes or Render's ~100s
# edge-proxy timeout kills the connection with a 502 — which is not just a
# slow response, it looks like the whole service is down. gthread lets the
# 8 threads interleave, so /progress and /result stay reachable mid-build.
# workers=1 because build progress and design-rotation state live in process
# memory, not a shared store. Long --timeout: a full build takes minutes at
# 0.1 vCPU, longer than the proxy will wait for any single response, but the
# frontend already recovers from that via /progress + /result (see app.py).
CMD gunicorn --bind 0.0.0.0:${PORT:-5001} --workers 1 --worker-class gthread --threads 8 --timeout 900 app:app
