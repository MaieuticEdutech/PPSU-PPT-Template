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

# One worker: build progress (PROGRESS dict) and design-rotation state live in
# process memory. Threads keep /progress polls answered while a /build runs.
# Long timeout: a full build takes minutes at 0.1 vCPU.
CMD gunicorn --bind 0.0.0.0:${PORT:-5001} --workers 1 --threads 8 --timeout 900 app:app
