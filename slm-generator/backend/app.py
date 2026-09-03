#!/usr/bin/env python3
"""app.py — Phase 5 web layer for the SLM generator (FastAPI, per spec).

One deliberate deviation from the spec's "same LAN pattern as ../ppsu1"
(separate static server + a hard-coded BACKEND_HOST constant): the frontend
is served FROM this backend on ONE port, and the page calls the API with
relative paths. The ppsu1 pattern's separate origin caused this project a
whole afternoon of CORS/DHCP-IP-drift debugging during its cloud deployment
(see ../../DEPLOYMENT.md); single-origin makes that entire failure class
impossible, and LAN sharing still works (uvicorn binds 0.0.0.0 — run.bat
prints the shareable address).

Jobs run in a background thread, ONE at a time (a single 8 GB GPU serialises
generation anyway); a second submission while one runs is queued nowhere —
it's refused with 409 and the page says so. Progress is the generator's own
per-call labels, polled from an in-memory registry (single-process server).
"""
import json
import shutil
import tempfile
import threading
import uuid
from pathlib import Path

from fastapi import FastAPI, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse

import docx2pdf
import figures
from docx_builder import build as build_docx
from unit_generator import GenerationError, generate_unit

APP_DIR = Path(__file__).parent
FRONTEND = APP_DIR.parent / "frontend" / "index.html"
JOB_ROOT = Path(tempfile.gettempdir()) / "slm_generator_jobs"
JOB_ROOT.mkdir(exist_ok=True)

app = FastAPI(title="PPSU SLM Generator")

JOBS = {}                      # job_id -> dict (in-memory, single process)
_RUN_LOCK = threading.Lock()   # one generation at a time (one GPU)

# tests inject a fast fake here; production uses the real generator
app.state.generate_fn = generate_unit


def _job_dir(job_id: str) -> Path:
    if not job_id or not job_id.isalnum():
        raise HTTPException(404)
    d = JOB_ROOT / job_id
    if not d.is_dir():
        raise HTTPException(404)
    return d


def _run_job(job_id, meta, syllabus, toc_text, source_path, brand="ppsu"):
    job = JOBS[job_id]
    work = JOB_ROOT / job_id

    def progress(msg):
        job["calls_done"] += 1
        job["current"] = msg

    try:
        unit, report = app.state.generate_fn(
            meta, syllabus_topics=syllabus, toc_text=toc_text,
            source_path=source_path, progress=progress)
        job["report"] = report
        (work / "unit.json").write_text(json.dumps(unit, indent=2),
                                        encoding="utf-8")
        (work / "report.json").write_text(json.dumps(report, indent=2),
                                          encoding="utf-8")
        (work / "figures.txt").write_text(figures.figure_list_text(unit),
                                          encoding="utf-8")
        if report["validation"]["errors"]:
            job["state"] = "failed"
            job["error"] = ("The generated unit failed validation and was "
                            "not rendered: "
                            + "; ".join(report["validation"]["errors"]))
            return
        build_docx(unit, brand=brand).save(str(work / "unit.docx"))
        job["pdf"] = False
        if docx2pdf.available():
            job["current"] = "exporting PDF…"
            job["pdf"] = docx2pdf.convert(work / "unit.docx",
                                          work / "unit.pdf")
        job["state"] = "done"
    except GenerationError as e:
        job["state"] = "failed"
        job["error"] = str(e)
    except Exception as e:                              # noqa: BLE001
        job["state"] = "failed"
        job["error"] = f"unexpected error: {e.__class__.__name__}: {e}"
    finally:
        job["current"] = ""
        _RUN_LOCK.release()


@app.get("/", response_class=HTMLResponse)
def index():
    return FRONTEND.read_text(encoding="utf-8")


@app.get("/api/status")
def status():
    return {"service": "slm-generator", "pdf_export": docx2pdf.available(),
            "busy": _RUN_LOCK.locked()}


@app.get("/api/branding")
def get_branding():
    import branding
    return {"logo": branding.logo_path() is not None,
            "profile": branding.load_profile()}


@app.post("/api/branding")
async def set_branding(logo: UploadFile | None = None,
                       reference: UploadFile | None = None):
    """Institution branding, set once and applied to every future unit:
    a logo image (cover + page headers) and/or a reference SLM PDF whose
    style signals (font, sizes, heading colour, page geometry) override
    the builder defaults."""
    import branding
    out = {}
    if logo is not None and logo.filename:
        suffix = Path(logo.filename).suffix.lower()
        if suffix not in (".png", ".jpg", ".jpeg"):
            raise HTTPException(400, "logo must be .png or .jpg")
        data = await logo.read()
        if len(data) > 5_000_000:
            raise HTTPException(400, "logo larger than 5 MB")
        try:
            branding.save_logo(data, suffix)
        except ValueError as e:
            raise HTTPException(400, str(e))
        out["logo"] = "saved"
    if reference is not None and reference.filename:
        if not reference.filename.lower().endswith(".pdf"):
            raise HTTPException(400, "reference must be a .pdf")
        data = await reference.read()
        try:
            out["profile"] = branding.save_reference_pdf(data)
        except ValueError as e:
            raise HTTPException(400, str(e))
    if not out:
        raise HTTPException(400, "upload a logo and/or a reference PDF")
    return out


@app.post("/api/generate")
async def generate(programme: str = Form(""), course_code: str = Form(""),
                   course_name: str = Form(""),
                   unit_number: int = Form(1),
                   unit_title: str = Form(""),
                   syllabus_topics: str = Form(""),
                   toc_text: str = Form(""),
                   textbook_citation: str = Form(""),
                   brand: str = Form("ppsu"),
                   source: UploadFile | None = None):
    if brand not in ("ppsu", "reva"):
        raise HTTPException(400, "brand must be 'ppsu' or 'reva'")
    for name, val in (("programme", programme),
                      ("course code", course_code),
                      ("course name", course_name),
                      ("unit title", unit_title)):
        if not val.strip():
            raise HTTPException(400, f"missing {name}")

    if not _RUN_LOCK.acquire(blocking=False):
        raise HTTPException(409, "another unit is generating — one at a "
                                 "time on this machine's GPU")
    try:
        job_id = uuid.uuid4().hex
        work = JOB_ROOT / job_id
        work.mkdir()

        source_path = None
        if source is not None and source.filename:
            suffix = Path(source.filename).suffix.lower()
            if suffix not in (".pdf", ".docx", ".txt", ".md"):
                raise HTTPException(400, "source must be .pdf, .docx or "
                                          ".txt")
            source_path = work / f"source{suffix}"
            source_path.write_bytes(await source.read())

        meta = {"programme": programme.strip(),
                "course_code": course_code.strip(),
                "course_name": course_name.strip(),
                "unit_number": unit_number,
                "unit_title": unit_title.strip()}
        if textbook_citation.strip():
            meta["textbook_citation"] = textbook_citation.strip()
        syllabus = [t.strip() for t in syllabus_topics.splitlines()
                    if t.strip()] or None
        toc = toc_text.strip() or None

        JOBS[job_id] = {"state": "running", "calls_done": 0,
                        "current": "starting…", "error": None,
                        "report": None, "pdf": False, "brand": brand,
                        "meta": meta}
        threading.Thread(target=_run_job,
                         args=(job_id, meta, syllabus, toc, source_path,
                               brand),
                         daemon=True).start()
        return {"job_id": job_id}
    except HTTPException:
        _RUN_LOCK.release()
        raise
    except Exception:
        _RUN_LOCK.release()
        raise


@app.get("/api/progress/{job_id}")
def progress(job_id: str):
    job = JOBS.get(job_id)
    if job is None:
        raise HTTPException(404)
    out = {k: job[k] for k in ("state", "calls_done", "current", "error",
                               "pdf")}
    if job["state"] in ("done", "failed") and job.get("report"):
        r = job["report"]
        out["summary"] = {
            "source_mode": r.get("source_mode"),
            "failures": len(r.get("failures", [])),
            "failed_calls": [f.get("call") for f in r.get("failures", [])],
            "uk_spelling_fixes": r.get("uk_spelling_fixes"),
            "warnings": r.get("validation", {}).get("warnings", []),
            "errors": r.get("validation", {}).get("errors", []),
            "review_note": r.get("review_note"),
        }
    return out


def _serve(job_id, filename, download_name, media_type):
    path = _job_dir(job_id) / filename
    if not path.is_file():
        raise HTTPException(404, f"{filename} not available for this job")
    return FileResponse(str(path), filename=download_name,
                        media_type=media_type)


def _download_name(job_id, ext):
    """REVA's file-naming convention (CourseName_UnitNN_Title) when that
    brand generated the job; a plain name otherwise."""
    job = JOBS.get(job_id) or {}
    if job.get("brand") == "reva" and job.get("meta"):
        import brands
        return brands.reva_filename(job["meta"], ext)
    return f"slm_unit{ext}"


@app.get("/api/download/{job_id}/docx")
def download_docx(job_id: str):
    return _serve(job_id, "unit.docx", _download_name(job_id, ".docx"),
                  "application/vnd.openxmlformats-officedocument"
                  ".wordprocessingml.document")


@app.get("/api/download/{job_id}/pdf")
def download_pdf(job_id: str):
    return _serve(job_id, "unit.pdf", _download_name(job_id, ".pdf"),
                  "application/pdf")


@app.get("/api/download/{job_id}/figures")
def download_figures(job_id: str):
    return _serve(job_id, "figures.txt", "figure_placeholders.txt",
                  "text/plain")


@app.get("/api/download/{job_id}/report")
def download_report(job_id: str):
    return _serve(job_id, "report.json", "generation_report.json",
                  "application/json")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8010)
