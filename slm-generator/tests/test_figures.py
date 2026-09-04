#!/usr/bin/env python3
"""Offline tests for auto-generated figures: the deterministic diagram
renderer (figure_render.py), docx embedding with placeholder fallback,
the DTP list's rendered/placeholder statuses, and the generate_unit
figure pass (stub engine — spec calls, containment, report counters).

Run: python tests/test_figures.py
"""
import sys
import tempfile
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "backend"))

import figure_render
import figures
import schemas
from ai_engine import OllamaEngine
from docx_builder import build
from unit_generator import generate_unit

passed = failed = 0


def check(name, condition):
    global passed, failed
    if condition:
        passed += 1
        print(f"  OK   {name}")
    else:
        failed += 1
        print(f"  FAIL {name}")


tmp = Path(tempfile.mkdtemp(prefix="slm_fig_test_"))

print("=== figure_render: every kind produces a real PNG ===")
SPECS = {
    "flow": {"kind": "flow", "items": [
        {"label": "Input", "detail": "raw data"},
        {"label": "Process"}, {"label": "Output"}]},
    "flow_long": {"kind": "flow", "items": [
        {"label": f"Step {i}"} for i in range(1, 7)]},
    "cycle": {"kind": "cycle", "items": [
        {"label": "Plan"}, {"label": "Do"}, {"label": "Check"},
        {"label": "Act"}]},
    "hierarchy": {"kind": "hierarchy", "root": "Ciphers", "items": [
        {"label": "Substitution", "detail": "replace"},
        {"label": "Transposition", "detail": "reorder"}]},
    "bar_chart": {"kind": "bar_chart", "items": [
        {"label": "E", "value": 12.7}, {"label": "T", "value": 9.1}]},
}
for name, spec in SPECS.items():
    p = figure_render.render(spec, tmp / f"{name}.png")
    data = p.read_bytes()
    check(f"{name}: valid non-trivial PNG",
          data[:8] == b"\x89PNG\r\n\x1a\n" and len(data) > 3000)

print("\n=== figure_render: unrenderable specs rejected ===")
BAD = [
    ("unknown kind", {"kind": "pie", "items": SPECS["flow"]["items"]}),
    ("too few items", {"kind": "flow", "items": [{"label": "one"}]}),
    ("too many items", {"kind": "flow",
                        "items": [{"label": f"i{i}"} for i in range(9)]}),
    ("empty label", {"kind": "flow", "items": [{"label": " "},
                                               {"label": "b"}]}),
    ("bar chart without numeric values",
     {"kind": "bar_chart", "items": [{"label": "a", "value": "big"},
                                     {"label": "b", "value": 2}]}),
    ("bar chart with boolean value",
     {"kind": "bar_chart", "items": [{"label": "a", "value": True},
                                     {"label": "b", "value": 2}]}),
]
for name, spec in BAD:
    try:
        figure_render.render(spec, tmp / "bad.png")
        check(f"{name} raises ValueError", False)
    except ValueError:
        check(f"{name} raises ValueError", True)

print("\n=== docx embedding + placeholder fallback ===")
png = tmp / "flow.png"
UNIT = {
    "meta": {"programme": "MSc", "course_code": "T", "course_name": "C",
             "unit_number": 1, "unit_title": "U", "source_mode": "ai"},
    "introduction": ["Hello."],
    "learning_objectives": [{"verb": "Define", "rest": "r"}] * 4,
    "sections": [
        {"number": "1.1", "title": "S", "intro": "",
         "subsections": [
             {"number": "1.1.1", "title": "Sub",
              "blocks": [
                  {"type": "figure", "caption": "Figure 1: Rendered",
                   "image": str(png)},
                  {"type": "figure", "caption": "Figure 2: Missing file",
                   "image": str(tmp / "nope.png")},
                  {"type": "figure", "caption": "Figure 3: No image"},
              ]}]}],
    "summary": ["s"], "glossary": [{"term": "t", "definition": "d"}],
    "case_study": {"title": "c", "background": ["b"], "questions": ["q"]},
    "self_assessment": {"mcq": [{"q": "q", "options": ["1", "2", "3", "4"],
                                 "answer": "a"}] * 8, "fill_blanks": []},
    "terminal": {"short": [{"q": "q", "answer": "a"}] * 5,
                 "long": [{"q": "q", "answer": "a"}] * 5},
    "references": ["r"],
}
docx_path = tmp / "unit.docx"
build(UNIT, use_branding=False).save(str(docx_path))
zf = zipfile.ZipFile(docx_path)
media = [n for n in zf.namelist() if n.startswith("word/media/")]
doc_xml = zf.read("word/document.xml").decode("utf-8")
check("rendered figure embedded as a media image", len(media) == 1)
check("missing-file and no-image figures keep the placeholder box",
      doc_xml.count("FIGURE PLACEHOLDER") == 2)
check("all three captions still render",
      all(f"Figure {i}:" in doc_xml for i in (1, 2, 3)))

print("\n=== DTP figure list statuses ===")
figs = figures.figure_list(UNIT)
check("list marks rendered vs placeholder",
      [f["rendered"] for f in figs] == [True, True, False])
txt = figures.figure_list_text(UNIT)
fig_lines = [ln for ln in txt.splitlines() if ln[:1].isdigit()]
check("handoff text labels each figure's status",
      [ln.split("] ")[0] + "]" for ln in fig_lines]
      == ["1. [auto-rendered]", "2. [auto-rendered]",
          "3. [PLACEHOLDER]"])

print("\n=== generate_unit figure pass (stub engine) ===")
META = {"programme": "MSc SEMESTER (I)", "course_code": "TEST101",
        "course_name": "Test Course", "unit_number": 2,
        "unit_title": "Stubbed Unit"}
CANNED = {
    id(schemas.OUTLINE): {
        "example_style": "code",
        "sections": [
            {"title": f"Major Topic {i}", "intro": f"Opener {i}.",
             "subsections": [{"title": f"Sub {i}.1"},
                             {"title": f"Sub {i}.2"},
                             {"title": f"Sub {i}.3"}]}
            for i in (1, 2, 3)
        ],
    },
    id(schemas.INTRODUCTION): {"paragraphs": ["One.", "Two.", "Three."]},
    id(schemas.LEARNING_OBJECTIVES): {"learning_objectives": [
        {"verb": "Define", "rest": "the key terms of the subject clearly"},
        {"verb": "Apply", "rest": "the core operations to given datasets"},
        {"verb": "Analyse", "rest": "relations between multiple entities"},
        {"verb": "Illustrate", "rest": "concepts using standard diagrams"}]},
    id(schemas.PROSE): {"paragraphs": ["A teaching paragraph.", "More."]},
    id(schemas.TABLE): {"caption_title": "Concepts",
                        "columns": ["Concept", "Meaning"],
                        "rows": [["A", "1"], ["B", "2"]]},
    id(schemas.CODE): {"language": "python", "code": "print(1)",
                       "explanation": "Prints one."},
    id(schemas.DID_YOU_KNOW): {"text": "An interesting aside."},
    id(schemas.SECTION_EXTRAS): {"think_and_apply": "Apply this.",
                                 "figure_caption": "How the parts connect"},
    id(schemas.SUMMARY): {"summary": [f"Point {i}" for i in range(1, 8)]},
    id(schemas.GLOSSARY): {"glossary": [
        {"term": f"Term {i:02d}", "definition": f"Definition {i}."}
        for i in range(1, 15)]},
    id(schemas.CASE_STUDY): {"title": "Case",
                             "background": ["P1.", "P2.", "P3."],
                             "questions": ["Q1?", "Q2?", "Q3?"]},
    id(schemas.MCQS): {"mcq": [
        {"q": f"Question {i}?", "options": ["w", "x", "y", "z"],
         "answer": "b"} for i in range(1, 9)]},
    id(schemas.FILL_BLANKS): {"fill_blanks": [
        {"q": f"The ______ number {i}.", "answer": f"a{i}"}
        for i in range(1, 6)]},
    id(schemas.TERMINAL_SHORT): {"short": [
        {"q": f"Short Q{i}?", "answer": f"A{i}."} for i in range(1, 6)]},
    id(schemas.TERMINAL_LONG_QS): {"long_questions": [
        f"Long Q{i}?" for i in range(1, 6)]},
    id(schemas.LONG_ANSWER): {"answer": "A model essay answer."},
    id(schemas.REFERENCES): {"references": [
        f"Author {i}. (2020). Book {i}. Publisher." for i in range(1, 7)]},
    id(schemas.RELEVANCE): {"off_topic": []},
}


class FigStubEngine(OllamaEngine):
    def __init__(self, figure_reply):
        super().__init__(model="stub", host="http://stub")
        self.figure_reply = figure_reply
        self.figure_calls = 0

    def ask(self, task_prompt, schema=None, *, system=None,
            temperature=0.3):
        if id(schema) == id(schemas.FIGURE_SPEC):
            self.figure_calls += 1
            return self.figure_reply
        return CANNED[id(schema)]


def quiet(_msg):
    pass


figdir = tmp / "unit_figs"
engine = FigStubEngine({"kind": "flow", "items": [
    {"label": "A"}, {"label": "B"}, {"label": "C"}]})
unit, report = generate_unit(META, engine=engine, progress=quiet,
                             figures_dir=figdir)
fig_blocks = [b for sec in unit["sections"] for sub in sec["subsections"]
              for b in sub["blocks"] if b["type"] == "figure"]
check("one spec call per figure placeholder",
      engine.figure_calls == len(fig_blocks) and fig_blocks)
check("every figure block got an image path that exists",
      all(b.get("image") and Path(b["image"]).is_file()
          for b in fig_blocks))
check("PNGs land in the requested figures_dir",
      all(Path(b["image"]).parent == figdir for b in fig_blocks))
check("report counts planned and rendered figures",
      report["figures"] == {"planned": len(fig_blocks),
                            "rendered": len(fig_blocks)})
check("figures never break validation",
      report["validation"]["errors"] == [])

engine = FigStubEngine({"kind": "bar_chart", "items": [
    {"label": "a"}, {"label": "b"}]})     # unrenderable: no values
unit, report = generate_unit(META, engine=engine, progress=quiet,
                             figures_dir=tmp / "unit_figs_bad")
fig_blocks = [b for sec in unit["sections"] for sub in sec["subsections"]
              for b in sub["blocks"] if b["type"] == "figure"]
check("unrenderable spec keeps the placeholder (no image key)",
      fig_blocks and not any(b.get("image") for b in fig_blocks))
check("report shows planned > rendered on failure",
      report["figures"]["planned"] == len(fig_blocks)
      and report["figures"]["rendered"] == 0)
check("unit still validates without rendered figures",
      report["validation"]["errors"] == [])

unit, report = generate_unit(META, engine=FigStubEngine({}),
                             progress=quiet)
check("no figures_dir -> no figure calls, no report entry (old behaviour)",
      "figures" not in report)

print(f"\n{passed} passed, {failed} failed")
sys.exit(1 if failed else 0)
