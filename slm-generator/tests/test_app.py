#!/usr/bin/env python3
"""Offline tests for the Phase 5 web layer (FastAPI TestClient, generator
faked for speed) and the Phase 6 figure list. PDF export is exercised only
as 'gracefully absent/present' — the live Word-COM conversion is checked by
the live server smoke, not here.

Run: python tests/test_app.py
"""
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "backend"))

from fastapi.testclient import TestClient

import app as app_module
import figures

passed = failed = 0


def check(name, condition):
    global passed, failed
    if condition:
        passed += 1
        print(f"  OK   {name}")
    else:
        failed += 1
        print(f"  FAIL {name}")


UNIT = {
    "meta": {"programme": "MSc", "course_code": "T", "course_name": "C",
             "unit_number": 3, "unit_title": "U", "source_mode": "ai"},
    "introduction": ["Hello."],
    "learning_objectives": [{"verb": "Define", "rest": "r"}] * 4,
    "sections": [
        {"number": f"3.{i}", "title": f"S{i}", "intro": "",
         "subsections": [
             {"number": f"3.{i}.1", "title": "Sub",
              "blocks": [{"type": "prose", "text": "T."},
                          {"type": "figure",
                           "caption": f"Figure {i}: Fig {i}"}]}]}
        for i in (1, 2, 3)],
    "summary": ["s"], "glossary": [{"term": "t", "definition": "d"}],
    "case_study": {"title": "c", "background": ["b"], "questions": ["q"]},
    "self_assessment": {"mcq": [{"q": "q", "options": ["1", "2", "3", "4"],
                                 "answer": "a"}] * 8, "fill_blanks": []},
    "terminal": {"short": [{"q": "q", "answer": "a"}] * 5,
                 "long": [{"q": "q", "answer": "a"}] * 5},
    "references": ["r"],
}
REPORT_OK = {"source_mode": "ai", "model": "fake", "calls": 3,
             "failures": [], "uk_spelling_fixes": 1,
             "review_note": "SMEs must review.",
             "validation": {"errors": [], "warnings": ["a warning"]}}


print("=== figures.py ===")
figs = figures.figure_list(UNIT)
check("all placeholders collected with locations",
      len(figs) == 3 and figs[0]["section"] == "3.1 S1"
      and figs[2]["caption"] == "Figure 3: Fig 3")
txt = figures.figure_list_text(UNIT)
check("DTP text file lists every figure with its location",
      txt.count("Location:") == 3 and "Unit 3: U" in txt)

print("\n=== web layer (generator faked) ===")
client = TestClient(app_module.app)

FORM = {"programme": "MSc", "course_code": "T", "course_name": "C",
        "unit_number": "3", "unit_title": "U",
        "syllabus_topics": "a\nb", "toc_text": "", "textbook_citation": ""}


def fake_ok(meta, syllabus_topics=None, toc_text=None, source_path=None,
            progress=None):
    progress("outline: ok")
    progress("prose: ok")
    unit = {**UNIT, "meta": {**UNIT["meta"], **meta}}
    return unit, dict(REPORT_OK)


r = client.get("/")
check("GET / serves the frontend", r.status_code == 200
      and "SLM Generator" in r.text and "api/generate" in r.text)
r = client.get("/api/status")
check("GET /api/status", r.status_code == 200
      and r.json()["service"] == "slm-generator")

r = client.post("/api/generate", data={**FORM, "course_name": "  "})
check("missing field -> 400", r.status_code == 400)
r = client.post("/api/generate", data={**FORM, "level": "phd"})
check("invalid level -> 400", r.status_code == 400)

app_module.app.state.generate_fn = fake_ok
r = client.post("/api/generate", data=FORM)
check("job accepted", r.status_code == 200 and "job_id" in r.json())
job_id = r.json()["job_id"]

for _ in range(50):
    p = client.get(f"/api/progress/{job_id}").json()
    if p["state"] != "running":
        break
    time.sleep(0.1)
check("job completes with progress counted",
      p["state"] == "done" and p["calls_done"] == 2)
check("summary surfaced (warnings + review note)",
      p["summary"]["warnings"] == ["a warning"]
      and "SMEs" in p["summary"]["review_note"])
check("meta from the form reached the generator (source_mode ai, no file)",
      p["summary"]["source_mode"] == "ai")

r = client.get(f"/api/download/{job_id}/docx")
check("docx download", r.status_code == 200
      and r.content[:2] == b"PK" and len(r.content) > 10_000)
r = client.get(f"/api/download/{job_id}/figures")
check("figure list download", r.status_code == 200
      and r.text.count("Location:") == 3)
r = client.get(f"/api/download/{job_id}/report")
check("report download", r.status_code == 200
      and r.json()["model"] == "fake")
check("unknown job 404s",
      client.get("/api/download/deadbeef99/docx").status_code == 404)

print("\n=== busy lock (one generation at a time) ===")


def fake_slow(meta, **kw):
    time.sleep(0.8)
    return {**UNIT, "meta": {**UNIT["meta"], **meta}}, dict(REPORT_OK)


app_module.app.state.generate_fn = fake_slow
r1 = client.post("/api/generate", data=FORM)
r2 = client.post("/api/generate", data=FORM)
check("second submission while busy -> 409",
      r1.status_code == 200 and r2.status_code == 409)
jid = r1.json()["job_id"]
for _ in range(60):
    if client.get(f"/api/progress/{jid}").json()["state"] != "running":
        break
    time.sleep(0.1)
r3 = client.post("/api/generate", data=FORM)
check("lock released after the job finishes", r3.status_code == 200)
for _ in range(60):
    if client.get(f"/api/progress/{r3.json()['job_id']}").json()["state"] != "running":
        break
    time.sleep(0.1)

print("\n=== failed validation never yields a docx ===")


def fake_invalid(meta, **kw):
    bad_report = {**REPORT_OK,
                  "validation": {"errors": ["MCQs: 0 (target 8)"],
                                 "warnings": []}}
    return {**UNIT, "meta": {**UNIT["meta"], **meta}}, bad_report


app_module.app.state.generate_fn = fake_invalid
jid = client.post("/api/generate", data=FORM).json()["job_id"]
for _ in range(50):
    p = client.get(f"/api/progress/{jid}").json()
    if p["state"] != "running":
        break
    time.sleep(0.1)
check("validation errors -> job failed with the errors named",
      p["state"] == "failed" and "MCQs" in p["error"])
check("no docx served for a failed job",
      client.get(f"/api/download/{jid}/docx").status_code == 404)
check("but the report IS still downloadable for diagnosis",
      client.get(f"/api/download/{jid}/report").status_code == 200)

print(f"\n{passed} passed, {failed} failed")
sys.exit(1 if failed else 0)
