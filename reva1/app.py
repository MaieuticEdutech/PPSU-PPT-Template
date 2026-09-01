#!/usr/bin/env python3
"""
Minimal website for the PPT template engine.
    pip install flask --break-system-packages
    python3 app.py
Then open http://localhost:5000

Flow: upload a raw PPT + a designed PPT -> the template gets auto-analyzed
(cached by file hash, so re-uploading the same template skips re-analysis)
-> raw content is matched against detected layouts -> final PPT to download.
"""
import hashlib
import json
import tempfile
import uuid
from pathlib import Path

from flask import Flask, request, send_file, render_template_string

from analyze_template import build_catalog
from engine import build as run_engine

APP_DIR = Path(__file__).parent
CACHE_DIR = APP_DIR / "catalog_cache"
CACHE_DIR.mkdir(exist_ok=True)
UPLOAD_DIR = Path(tempfile.gettempdir()) / "ppt_engine_uploads"
UPLOAD_DIR.mkdir(exist_ok=True)

app = Flask(__name__)


def template_hash(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()[:16]


def get_catalog(template_path: Path) -> Path:
    h = template_hash(template_path)
    cached = CACHE_DIR / f"{h}.json"
    if not cached.exists():
        catalog = build_catalog(template_path)
        cached.write_text(json.dumps(catalog, indent=2))
    return cached


def get_page_content():
    html_file = APP_DIR / "index.html"
    if not html_file.exists():
        return "<h2>index.html is missing. Please create it in the application directory.</h2>"
    return html_file.read_text(encoding="utf-8")


import traceback

@app.route("/")
def index():
    print("========== INDEX ROUTE CALLED ==========")
    return "Backend is alive!", 200

@app.route("/generate", methods=["POST"])
def generate():
    raw_file = request.files["raw"]
    template_file = request.files["template"]

    token = uuid.uuid4().hex
    work = UPLOAD_DIR / token
    work.mkdir()
    raw_path = work / "raw.pptx"
    template_path = work / "template.pptx"
    raw_file.save(raw_path)
    template_file.save(template_path)

    catalog_path = get_catalog(template_path)
    out_path = work / "final.pptx"

    report = run_engine(raw_path, template_path, catalog_path, out_path)
    (work / "report.json").write_text(json.dumps(report))

    return render_template_string(get_page_content(), report=report, token=token)


@app.route("/download/<token>")
def download(token):
    out_path = UPLOAD_DIR / token / "final.pptx"
    return send_file(out_path, as_attachment=True, download_name="final.pptx")


import os

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
