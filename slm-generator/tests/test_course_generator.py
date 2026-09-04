#!/usr/bin/env python3
"""Offline tests for the full-course engine (N source docs -> target_units
units). The critical property under test is the COVERAGE GUARANTEE: every
topic detected in the sources lands in exactly one generated unit, in
order, no matter how broken the model's plan is.

Run: python tests/test_course_generator.py
"""
import io
import json
import sys
import tempfile
import time
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "backend"))

import course_generator as cg
from unit_generator import GenerationError

passed = failed = 0


def check(name, condition):
    global passed, failed
    if condition:
        passed += 1
        print(f"  OK   {name}")
    else:
        failed += 1
        print(f"  FAIL {name}")


META = {"programme": "MSc", "course_code": "ICCS7010",
        "course_name": "Information Security", "level": "postgraduate"}


def topics(n, src="a.txt"):
    return [{"heading": f"{i+1}.1 Topic {i+1}", "text": f"Body {i+1}.",
             "source": src} for i in range(n)]


class StubCall:
    """Mimics unit_generator._Caller for plan_course."""

    def __init__(self, reply):
        self.reply = reply
        self.calls = []

    def __call__(self, label, prompt, schema, temperature=0.3,
                 num_predict=None):
        self.calls.append({"label": label, "prompt": prompt,
                           "schema": schema, "num_predict": num_predict})
        return self.reply


print("=== _repair_cuts ===")
good = [{"end_index": 2}, {"end_index": 5}, {"end_index": 9}]
check("valid monotonic plan kept as-is",
      cg._repair_cuts(good, 10, 3) == [2, 5, 9])
short = [{"end_index": 2}, {"end_index": 5}, {"end_index": 7}]
check("final cut forced to the last topic",
      cg._repair_cuts(short, 10, 3) == [2, 5, 9])
bad = [{"end_index": 5}, {"end_index": 3}, {"end_index": 9}]
cuts = cg._repair_cuts(bad, 10, 3)
check("non-monotonic plan replaced by an equal split covering all",
      cuts[-1] == 9 and cuts == sorted(set(cuts)) and len(cuts) == 3)
cuts = cg._repair_cuts([{"end_index": "x"}, {}], 7, 2)
check("garbage indices -> equal split still ends at last topic",
      cuts == sorted(set(cuts)) and cuts[-1] == 6 and len(cuts) == 2)
cuts = cg._repair_cuts([{"end_index": i} for i in range(14)], 14, 14)
check("one-topic-per-unit edge case (topics == units)",
      cuts == list(range(14)))
greedy = [{"end_index": 8}, {"end_index": 9}, {"end_index": 9}]
cuts = cg._repair_cuts(greedy, 10, 3)
check("plan leaving no topics for later units is repaired",
      len(cuts) == 3 and cuts == sorted(set(cuts)) and cuts[-1] == 9)

print("\n=== plan_course ===")
tps = topics(10)
call = StubCall({"units": [{"title": "Alpha", "end_index": 3},
                           {"title": "2.1 Beta (per the TOC)",
                            "end_index": 6},
                           {"title": "Gamma", "end_index": 9}]})
plan, repaired = cg.plan_course(META, tps, 3, call)
assigned = [c["heading"] for p in plan for c in p["chunks"]]
check("every topic assigned exactly once, in source order",
      assigned == [t["heading"] for t in tps])
check("units are contiguous and disjoint",
      [len(p["chunks"]) for p in plan] == [4, 3, 3])
check("titles cleaned of numbering/TOC noise",
      plan[1]["title"] == "Beta" and not repaired)
check("plan call carries a raised num_predict budget",
      call.calls[0]["num_predict"] == 4096)
check("numbered topic list with provenance reaches the prompt",
      "0. 1.1 Topic 1  [a.txt]" in call.calls[0]["prompt"])

call = StubCall({"units": [{"title": "A", "end_index": 9},
                           {"title": "B", "end_index": 1},
                           {"title": "C", "end_index": 4}]})
plan, repaired = cg.plan_course(META, tps, 3, call)
assigned = [c["heading"] for p in plan for c in p["chunks"]]
check("broken plan repaired -> coverage still complete",
      repaired and assigned == [t["heading"] for t in tps]
      and all(p["chunks"] for p in plan))

try:
    cg.plan_course(META, topics(3), 5, StubCall(None))
    check("too few topics raises GenerationError", False)
except GenerationError as e:
    check("too few topics raises GenerationError", "3" in str(e))

try:
    cg.plan_course(META, tps, 3, StubCall(None))
    check("plan call failing twice raises GenerationError", False)
except GenerationError:
    check("plan call failing twice raises GenerationError", True)

pre = [{"heading": "(preamble)", "text": "ignored", "source": "a.txt"}]
call = StubCall({"units": [{"title": "A", "end_index": 4},
                           {"title": "B", "end_index": 9}]})
plan, _ = cg.plan_course(META, pre + tps, 2, call)
check("preamble chunks excluded from the topic inventory",
      sum(len(p["chunks"]) for p in plan) == 10)

print("\n=== generate_course (unit generator faked) ===")
tmp = Path(tempfile.mkdtemp(prefix="slm_course_test_"))
src1 = tmp / "unitA.txt"
src1.write_text("1.1 Sets\nSets body.\n1.2 Relations\nRelations body.\n"
                "1.3 Functions\nFunctions body.\n", encoding="utf-8")
src2 = tmp / "unitB.txt"
src2.write_text("2.1 Graphs\nGraphs body.\n2.2 Trees\nTrees body.\n"
                "2.3 Logic\nLogic body.\n", encoding="utf-8")


class StubEngine:
    """Answers only the plan call (unit generation is monkeypatched)."""

    def ask(self, prompt, schema, system=None, temperature=0.3,
            num_predict=None):
        return {"units": [{"title": "Foundations", "end_index": 1},
                          {"title": "Mappings", "end_index": 2},
                          {"title": "Structures", "end_index": 4},
                          {"title": "Reasoning", "end_index": 5}]}


seen_units = []


def fake_generate_unit(meta, syllabus_topics=None, toc_text=None,
                       source_path=None, source_chunks=None,
                       source_label=None, engine=None, progress=None,
                       figures_dir=None):
    seen_units.append({"meta": meta, "chunks": source_chunks,
                       "label": source_label})
    progress("outline: ok")
    if meta["unit_number"] == 3:
        raise RuntimeError("boom")   # containment: others must survive
    unit = {"meta": meta, "sections": []}
    report = {"validation": {"errors": [], "warnings": []}}
    if meta["unit_number"] == 4:
        report["validation"]["errors"] = ["too few MCQs"]
    return unit, report


real_generate_unit = cg.generate_unit
cg.generate_unit = fake_generate_unit
msgs = []
results, creport = cg.generate_course(
    META, [src1, src2], target_units=4, engine=StubEngine(),
    progress=msgs.append)
cg.generate_unit = real_generate_unit

check("all 6 topics from both files assigned, none skipped",
      creport["coverage"] == {"topics_total": 6, "topics_assigned": 6,
                              "complete": True})
all_headings = [c["heading"] for u in seen_units for c in u["chunks"]]
check("units received disjoint contiguous slices in source order",
      all_headings == ["1.1 Sets", "1.2 Relations", "1.3 Functions",
                       "2.1 Graphs", "2.2 Trees", "2.3 Logic"])
check("per-file provenance recorded",
      [s["file"] for s in creport["sources"]] == ["unitA.txt", "unitB.txt"]
      and all(s["topics"] == 3 for s in creport["sources"]))
check("planned unit numbers and titles land in each unit's meta",
      seen_units[0]["meta"]["unit_number"] == 1
      and seen_units[0]["meta"]["unit_title"] == "Foundations"
      and seen_units[3]["meta"]["unit_number"] == 4)
check("course meta fields (level etc.) inherited by every unit",
      all(u["meta"]["level"] == "postgraduate" for u in seen_units))
check("unit 3 crash contained -> marked failed, others generated",
      creport["units"][2]["status"] == "failed"
      and "boom" in creport["units"][2]["error"]
      and creport["units"][0]["status"] == "ok"
      and creport["units"][1]["status"] == "ok")
check("validation-failed unit not treated as ok",
      creport["units"][3]["status"] == "failed"
      and "too few MCQs" in creport["units"][3]["error"])
check("unit progress labelled with its position",
      any(m.startswith("[unit 2/4]") for m in msgs))
check("plan recorded per unit with its topics",
      creport["plan"][0]["topics"] == ["1.1 Sets", "1.2 Relations"])

print("\n=== write_course_outputs ===")
GOOD_UNIT = {
    "meta": {"programme": "MSc", "course_code": "T", "course_name": "C",
             "unit_number": 1, "unit_title": "U", "source_mode": "textbook",
             "level": "postgraduate"},
    "introduction": ["Hello."],
    "learning_objectives": [{"verb": "Define", "rest": "r"}] * 4,
    "sections": [
        {"number": "1.1", "title": "S", "intro": "",
         "subsections": [
             {"number": "1.1.1", "title": "Sub",
              "blocks": [{"type": "prose", "text": "T."},
                         {"type": "figure", "caption": "Figure 1: F"}]}]}],
    "summary": ["s"], "glossary": [{"term": "t", "definition": "d"}],
    "case_study": {"title": "c", "background": ["b"], "questions": ["q"]},
    "self_assessment": {"mcq": [{"q": "q", "options": ["1", "2", "3", "4"],
                                 "answer": "a"}] * 8, "fill_blanks": []},
    "terminal": {"short": [{"q": "q", "answer": "a"}] * 5,
                 "long": [{"q": "q", "answer": "a"}] * 5},
    "references": ["r"],
}
fake_results = [
    {"unit_number": 1, "title": "U",
     "unit": GOOD_UNIT, "report": {"validation": {"errors": [],
                                                  "warnings": []}},
     "error": None},
    {"unit_number": 2, "title": "V", "unit": None,
     "report": {"validation": {"errors": ["bad"], "warnings": []}},
     "error": "validation: bad"},
]
fake_creport = {"units": [], "coverage": {"complete": True}}
out_dir = tmp / "out"
zip_path = cg.write_course_outputs(fake_results, fake_creport, out_dir,
                                   brand="ppsu", make_pdf=False,
                                   progress=lambda m: None)
names = set(zipfile.ZipFile(zip_path).namelist())
check("zip holds the good unit's docx + figures + both reports + course "
      "report",
      {"unit01.docx", "unit01.figures.txt", "unit01.report.json",
       "unit02.report.json", "course_report.json"} <= names)
check("failed unit produced no docx", "unit02.docx" not in names)
docx_bytes = zipfile.ZipFile(zip_path).read("unit01.docx")
check("unit docx is a valid OOXML zip",
      zipfile.is_zipfile(io.BytesIO(docx_bytes)))
reva_zip = cg.write_course_outputs(fake_results, fake_creport,
                                   tmp / "out_reva", brand="reva",
                                   make_pdf=False, progress=lambda m: None)
check("REVA brand uses the mandated file naming",
      any(n.startswith("C_Unit01") and n.endswith(".docx")
          for n in zipfile.ZipFile(reva_zip).namelist()))

print("\n=== web layer (/api/generate_course, generators faked) ===")
from fastapi.testclient import TestClient

import app as app_module

client = TestClient(app_module.app)
COURSE_FORM = {"programme": "MSc", "course_code": "T", "course_name": "C",
               "target_units": "4", "toc_text": "", "brand": "ppsu",
               "level": "postgraduate"}
FILES = [("sources", ("u1.txt", b"1.1 A\nBody.\n1.2 B\nBody.\n",
                      "text/plain")),
         ("sources", ("u2.txt", b"2.1 C\nBody.\n2.2 D\nBody.\n",
                      "text/plain"))]


def fake_course_fn(meta_base, source_paths, target_units=14, toc_text=None,
                   progress=None, figures_root=None):
    progress("[unit 1/4] outline: ok")
    creport = {"sources": [{"file": Path(p).name, "chars": 1, "topics": 2}
                           for p in source_paths],
               "target_units": target_units, "plan_repaired": False,
               "units": [{"unit_number": i, "title": f"U{i}",
                          "status": "ok", "error": None}
                         for i in range(1, target_units + 1)],
               "coverage": {"topics_total": 4, "topics_assigned": 4,
                            "complete": True},
               "calls": 1, "failures": []}
    return ([{"unit_number": i, "title": f"U{i}", "unit": GOOD_UNIT,
              "report": {"validation": {"errors": [], "warnings": []}},
              "error": None} for i in range(1, target_units + 1)],
            creport)


def fake_write_fn(results, creport, out_dir, brand="ppsu",
                  progress=None, **kw):
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    z = out_dir / "course_units.zip"
    with zipfile.ZipFile(z, "w") as zf:
        zf.writestr("unit01.docx", b"x")
    return z


app_module.app.state.generate_course_fn = fake_course_fn
app_module.app.state.write_course_fn = fake_write_fn

r = client.post("/api/generate_course", data=COURSE_FORM)
check("missing sources rejected 400", r.status_code == 400)
r = client.post("/api/generate_course",
                data={**COURSE_FORM, "course_name": ""}, files=FILES)
check("missing course name rejected 400", r.status_code == 400)
r = client.post("/api/generate_course",
                data={**COURSE_FORM, "target_units": "1"}, files=FILES)
check("target_units out of range rejected 400", r.status_code == 400)
r = client.post("/api/generate_course",
                data={**COURSE_FORM, "brand": "oxford"}, files=FILES)
check("unknown brand rejected 400", r.status_code == 400)
check("rejections released the run lock",
      not app_module._RUN_LOCK.locked())

r = client.post("/api/generate_course", data=COURSE_FORM, files=FILES)
check("valid course submission accepted", r.status_code == 200)
job_id = r.json()["job_id"]
for _ in range(100):
    p = client.get(f"/api/progress/{job_id}").json()
    if p["state"] != "running":
        break
    time.sleep(0.05)
check("course job completes", p["state"] == "done")
check("progress marks the job as a course", p["kind"] == "course")
s = p.get("summary", {})
check("course summary carries coverage + per-unit statuses",
      s.get("coverage", {}).get("complete") is True
      and len(s.get("units", [])) == 4
      and s["units"][0]["status"] == "ok")
r = client.get(f"/api/download/{job_id}/zip")
check("course ZIP downloads", r.status_code == 200
      and r.headers["content-type"] == "application/zip")
r = client.get(f"/api/download/{job_id}/report")
check("course report downloads via the standard report route",
      r.status_code == 200 and r.json()["coverage"]["complete"] is True)
check("run lock released after the course job",
      not app_module._RUN_LOCK.locked())

print(f"\n{passed} passed, {failed} failed")
sys.exit(1 if failed else 0)
