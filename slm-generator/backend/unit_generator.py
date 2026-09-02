#!/usr/bin/env python3
"""unit_generator.py — Phase 3 orchestrator: many small AI calls -> one
assembled, validated unit JSON -> (if clean) a rendered .docx.

Input modes (chosen automatically, recorded in meta.source_mode):
- "textbook": full source material supplied  [Phase 4 — not wired yet]
- "toc+ai":   no textbook, but the textbook's table of contents was given.
              The TOC pins the outline's structure/ordering (merged with
              syllabus topics when both exist); the AI writes the content.
- "ai":       only metadata + syllabus topics; the AI does everything.
Both non-textbook modes are AI-generated content and must be flagged for
harder SME review — a TOC pins structure, not facts.

Failure containment: every AI call is individually wrapped. A block that
fails twice (ai_engine's retry contract) is recorded in the report and
skipped — the unit still completes. Only an outline failure aborts, since
nothing can be generated without one. Numbering (1.1, 1.1.1, Table/Figure/
Problem numbers) is assigned HERE in code, never trusted from the model.
"""
import argparse
import json
import re
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

import prompts
import schemas
from ai_engine import AIEngineError, get_engine
from uk_style import apply_to_unit
from validate_unit import validate_unit


class GenerationError(Exception):
    """The unit could not be generated at all (outline failed / no model)."""


def _default_progress(msg):
    print(f"[gen] {msg}", flush=True)


# The first live toc+ai run showed the model annotating outline titles with
# TOC bookkeeping — "The Caesar Cipher (2.1.1, 2.1.2 in the textbook TOC)" —
# which would land verbatim in the document's headings. The outline prompt
# now forbids it, and this strips any that slip through anyway: parentheticals
# mentioning the TOC/syllabus, plus leading list numbering like "2.1.1 ".
_TITLE_NOISE = re.compile(r"\s*\([^)]*\b(?:TOC|table of contents|syllabus)"
                          r"[^)]*\)", re.IGNORECASE)
_TITLE_LEAD_NUM = re.compile(r"^\s*\d+(?:\.\d+)*\.?\s+")


def clean_title(title):
    return _TITLE_LEAD_NUM.sub("", _TITLE_NOISE.sub("", title)).strip()


# The first live run also showed the model embedding option letters INSIDE
# MCQ option text ("a) AES") — the builder prefixes letters itself, which
# would render as "a) a) AES". Strip any leading letter marker.
_OPTION_PREFIX = re.compile(r"^\s*[a-dA-D][\).:]\s*")


def clean_mcq_options(mcqs):
    for item in mcqs:
        item["options"] = [_OPTION_PREFIX.sub("", o).strip()
                           for o in item["options"]]
    return mcqs


class _Caller:
    """Wraps engine.ask with failure containment + report bookkeeping."""

    def __init__(self, engine, report, progress):
        self.engine = engine
        self.report = report
        self.progress = progress

    def __call__(self, label, prompt, schema, temperature=0.3):
        self.report["calls"] += 1
        t0 = time.time()
        try:
            data = self.engine.ask(prompt, schema,
                                   system=prompts.SYSTEM_STYLE,
                                   temperature=temperature)
            self.progress(f"{label}: ok ({time.time()-t0:.1f}s)")
            return data
        except AIEngineError as e:
            self.progress(f"{label}: FAILED ({e})")
            self.report["failures"].append({"call": label, "error": str(e)})
            return None


def generate_unit(meta, syllabus_topics=None, toc_text=None, engine=None,
                  progress=_default_progress):
    """Returns (unit_dict, report_dict). Raises GenerationError only when
    generation cannot even start (no outline)."""
    engine = engine or get_engine()
    meta = dict(meta)
    meta["source_mode"] = "toc+ai" if toc_text else "ai"

    report = {"source_mode": meta["source_mode"], "model": engine.model,
              "calls": 0, "failures": [], "uk_spelling_fixes": 0,
              "review_note": ("Content is AI-generated"
                              + (" (structure follows the supplied TOC)"
                                 if toc_text else "")
                              + " — SMEs must review facts, worked examples "
                                "and references before issue.")}
    call = _Caller(engine, report, progress)
    n = meta["unit_number"]

    # ---- outline (the one call that must succeed) -------------------------
    out = call("outline",
               prompts.outline(meta, syllabus_topics, toc_text),
               schemas.OUTLINE)
    if out is None:
        raise GenerationError("outline generation failed twice — cannot "
                              "build a unit without one")
    example_style = out["example_style"]
    for sec in out["sections"]:
        sec["title"] = clean_title(sec["title"])
        for sub in sec["subsections"]:
            sub["title"] = clean_title(sub["title"])
    titles = "; ".join(s["title"] for s in out["sections"])

    # ---- front matter -----------------------------------------------------
    intro = call("introduction", prompts.introduction(meta, titles),
                 schemas.INTRODUCTION)
    objectives = call("learning objectives",
                      prompts.learning_objectives(meta, titles),
                      schemas.LEARNING_OBJECTIVES)

    # ---- sections ---------------------------------------------------------
    sections = []
    fig_no = 0
    for si, sec in enumerate(out["sections"], start=1):
        sec_no = f"{n}.{si}"
        sub_titles = "; ".join(s["title"] for s in sec["subsections"])
        extras = call(f"section {sec_no} extras",
                      prompts.section_extras(meta, sec["title"], sub_titles),
                      schemas.SECTION_EXTRAS)
        problem_seq = 0
        subsections = []
        for ki, sub in enumerate(sec["subsections"]):
            sub_no = f"{sec_no}.{ki + 1}"
            label = f"{sub_no} {sub['title']}"
            blocks = []

            p = call(f"{label}: prose",
                     prompts.prose(meta, sec["title"], sub["title"]),
                     schemas.PROSE)
            if p:
                blocks += [{"type": "prose", "text": t}
                           for t in p["paragraphs"]]

            # enrichment rotation: 1st subsection a table, 2nd a worked
            # example (code or problem per the outline's example_style),
            # 3rd a did-you-know — mirrors the reference sample's rhythm
            if ki == 0:
                t = call(f"{label}: table",
                         prompts.table(meta, sec["title"], sub["title"]),
                         schemas.TABLE)
                if t:
                    blocks.append({"type": "table",
                                   "caption": f"Table {sub_no}: "
                                              f"{t['caption_title']}",
                                   "columns": t["columns"],
                                   "rows": t["rows"]})
            elif ki == 1:
                if example_style == "code":
                    c = call(f"{label}: code example",
                             prompts.code(meta, sec["title"], sub["title"]),
                             schemas.CODE)
                    if c:
                        blocks.append({"type": "code", "text": c["code"]})
                        blocks.append({"type": "prose",
                                       "text": c["explanation"]})
                else:
                    problem_seq += 1
                    pr = call(f"{label}: worked problem",
                              prompts.problem(meta, sec["title"],
                                              sub["title"]),
                              schemas.PROBLEM)
                    if pr:
                        blocks.append({"type": "problem",
                                       "label": f"Problem {sec_no}."
                                                f"{problem_seq}",
                                       "statement": pr["statement"],
                                       "solution": pr["solution"]})
            else:
                d = call(f"{label}: did-you-know",
                         prompts.did_you_know(meta, sec["title"],
                                              sub["title"]),
                         schemas.DID_YOU_KNOW)
                if d:
                    blocks.append({"type": "did_you_know", "text": d["text"]})

            # section extras land deterministically: figure on the first
            # subsection, think-and-apply closing the last
            if extras and ki == 0:
                fig_no += 1
                blocks.append({"type": "figure",
                               "caption": f"Figure {fig_no}: "
                                          f"{extras['figure_caption']}"})
            if extras and ki == len(sec["subsections"]) - 1:
                blocks.append({"type": "think_and_apply",
                               "title": "Think and Apply",
                               "text": extras["think_and_apply"]})

            subsections.append({"number": sub_no, "title": sub["title"],
                                "blocks": blocks})

        sections.append({"number": sec_no, "title": sec["title"],
                         "intro": sec.get("intro", ""),
                         "subsections": subsections})

    # ---- back matter ------------------------------------------------------
    summ = call("summary", prompts.summary(meta, titles), schemas.SUMMARY)
    glos = call("glossary", prompts.glossary(meta, titles), schemas.GLOSSARY)
    case = call("case study", prompts.case_study(meta, titles),
                schemas.CASE_STUDY)
    mcq = call("MCQs", prompts.mcqs(meta, titles), schemas.MCQS)
    blanks = call("fill-in-the-blanks", prompts.fill_blanks(meta, titles),
                  schemas.FILL_BLANKS)
    tshort = call("terminal short", prompts.terminal_short(meta, titles),
                  schemas.TERMINAL_SHORT)
    tlong = call("terminal long", prompts.terminal_long(meta, titles),
                 schemas.TERMINAL_LONG)
    refs = call("references", prompts.references(meta, titles),
                schemas.REFERENCES, temperature=0.1)

    unit = {
        "meta": meta,
        "introduction": intro["paragraphs"] if intro else [],
        "learning_objectives": (objectives or {}).get("learning_objectives",
                                                      []),
        "sections": sections,
        "summary": (summ or {}).get("summary", []),
        "glossary": (glos or {}).get("glossary", []),
        "case_study": case or {"title": "", "background": [],
                                "questions": []},
        "self_assessment": {"mcq": clean_mcq_options((mcq or {}).get("mcq",
                                                                     [])),
                            "fill_blanks": (blanks or {}).get("fill_blanks",
                                                              [])},
        "terminal": {"short": (tshort or {}).get("short", []),
                     "long": (tlong or {}).get("long", [])},
        "references": (refs or {}).get("references", []),
    }

    report["uk_spelling_fixes"] = apply_to_unit(unit)
    errors, warnings = validate_unit(unit)
    report["validation"] = {"errors": errors, "warnings": warnings}
    return unit, report


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--meta", required=True, type=Path,
                    help="JSON: programme, course_code, course_name, "
                         "unit_number, unit_title, syllabus_topics[]")
    ap.add_argument("--toc", type=Path,
                    help="optional textbook table-of-contents text file "
                         "(activates toc+ai mode)")
    ap.add_argument("--out", required=True, type=Path, help=".docx path")
    ap.add_argument("--report", type=Path)
    ap.add_argument("--json-out", type=Path,
                    help="also save the intermediate unit JSON")
    args = ap.parse_args()

    meta = json.loads(args.meta.read_text(encoding="utf-8"))
    syllabus = meta.pop("syllabus_topics", None)
    toc_text = args.toc.read_text(encoding="utf-8") if args.toc else None

    t0 = time.time()
    unit, report = generate_unit(meta, syllabus_topics=syllabus,
                                 toc_text=toc_text)
    report["seconds"] = round(time.time() - t0, 1)

    report_path = args.report or args.out.with_suffix(".report.json")
    report_path.write_text(json.dumps(report, indent=2), encoding="utf-8")
    if args.json_out:
        args.json_out.write_text(json.dumps(unit, indent=2),
                                 encoding="utf-8")

    errors = report["validation"]["errors"]
    if errors:
        # a failing unit never reaches the builder (CLAUDE.md convention)
        print(f"\nNOT RENDERED — validation errors "
              f"(report: {report_path}):")
        for e in errors:
            print(f"  - {e}")
        sys.exit(1)

    from docx_builder import build
    build(unit).save(str(args.out))
    print(f"\nwrote {args.out}  ({report['calls']} AI calls, "
          f"{len(report['failures'])} failed, "
          f"{report['uk_spelling_fixes']} UK-spelling fixes, "
          f"{report['seconds']}s)")
    for w in report["validation"]["warnings"]:
        print(f"  warning: {w}")


if __name__ == "__main__":
    main()
