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

import ingest
import prompts
import schemas
from ai_engine import AIEngineError, get_engine
from uk_style import apply_to_unit
from validate_unit import validate_unit

# how much source text one subsection call may carry (chars); ~1500-2000
# tokens, comfortable inside num_ctx=16384 with prompt + reply headroom
SUB_SOURCE_CAP = 6000


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

    def __call__(self, label, prompt, schema, temperature=0.3,
                 num_predict=None):
        self.report["calls"] += 1
        t0 = time.time()
        # num_predict forwarded only when overridden, so simple engine
        # stubs (tests) need not accept the parameter
        extra = {"num_predict": num_predict} if num_predict else {}
        try:
            data = self.engine.ask(prompt, schema,
                                   system=prompts.SYSTEM_STYLE,
                                   temperature=temperature, **extra)
            self.progress(f"{label}: ok ({time.time()-t0:.1f}s)")
            return data
        except AIEngineError as e:
            self.progress(f"{label}: FAILED ({e})")
            self.report["failures"].append({"call": label, "error": str(e)})
            return None


def _relevance_audit(unit, syllabus_topics, call):
    """Post-generation check that everything in the unit belongs to the
    stated course/unit/topics: one temperature-0.1 audit call per section
    plus one for the assessment material, each fed a digest of what was
    actually written. Returns [{scope, item, why}]. An audit call failing
    twice skips that scope (recorded via the normal failure log) — the
    audit must never block generation itself."""
    meta = unit["meta"]
    flags = []

    def run(scope, digest):
        if not digest.strip():
            return
        out = call(f"relevance audit: {scope}",
                   prompts.relevance_check(meta, syllabus_topics, scope,
                                           digest),
                   schemas.RELEVANCE, temperature=0.1)
        for f in (out or {}).get("off_topic", []):
            flags.append({"scope": scope, "item": f["item"][:120],
                          "why": f["why"][:200]})

    for sec in unit.get("sections", []):
        lines = [f'Section {sec["number"]}: {sec["title"]}']
        for sub in sec.get("subsections", []):
            lines.append(f'  Subsection {sub["number"]}: {sub["title"]}')
            for b in sub.get("blocks", []):
                t = b.get("type")
                if t == "prose":
                    lines.append(f'    prose: {b.get("text", "")[:200]}')
                elif t == "bullets":
                    lines.append('    points: '
                                 + "; ".join(b.get("items", []))[:250])
                elif t == "table":
                    lines.append(f'    table: {b.get("caption", "")}')
                elif t == "problem":
                    lines.append(f'    problem: '
                                 f'{b.get("statement", "")[:150]}')
                elif t in ("did_you_know", "think_and_apply"):
                    lines.append(f'    {t}: {b.get("text", "")[:150]}')
                elif t == "code":
                    lines.append(f'    code: {b.get("text", "")[:120]}')
        run(f'section {sec["number"]}', "\n".join(lines))

    assess = []
    if unit.get("glossary"):
        assess.append("Glossary terms: "
                      + "; ".join(e["term"] for e in unit["glossary"]))
    if unit.get("case_study", {}).get("title"):
        assess.append(f'Case study: {unit["case_study"]["title"]} — '
                      f'{(unit["case_study"].get("background") or [""])[0][:150]}')
    for m in unit.get("self_assessment", {}).get("mcq", []):
        assess.append(f'MCQ: {m["q"]}')
    for t in (unit.get("terminal", {}).get("short", [])
              + unit.get("terminal", {}).get("long", [])):
        assess.append(f'Question: {t["q"]}')
    run("assessment material", "\n".join(assess))
    return flags


def generate_unit(meta, syllabus_topics=None, toc_text=None,
                  source_path=None, source_chunks=None, source_label=None,
                  engine=None, progress=_default_progress,
                  figures_dir=None):
    """Returns (unit_dict, report_dict). Raises GenerationError only when
    generation cannot even start (no outline).

    Mode precedence (CLAUDE.md): source material (a file via source_path,
    or pre-chunked content via source_chunks — the course engine passes
    each planned unit its slice of the combined sources) -> "textbook";
    else toc_text -> "toc+ai"; else "ai"."""
    engine = engine or get_engine()
    meta = dict(meta)

    # ---- ingest the source, if any ---------------------------------------
    chunks, source_headings, condensed = [], None, None
    source_warnings = []
    source_present = source_path is not None or source_chunks is not None
    if source_present:
        meta["source_mode"] = "textbook"
        if source_chunks is not None:
            chunks = source_chunks
            text = "\n".join(c["text"] for c in chunks)
        else:
            text = ingest.extract_text(source_path)
            chunks = ingest.chunk_by_headings(text)
        real = [c for c in chunks if c["heading"] != "(preamble)"]
        if len(real) >= 3:
            source_headings = [c["heading"] for c in real]
        elif real and source_chunks is not None:
            # a course-planned unit may legitimately hold only 1-2 topics;
            # its headings still steer the outline
            source_headings = [c["heading"] for c in real]
        else:
            source_warnings.append(
                "source did not split into sections (no numbered headings "
                "detected) — the whole document grounds every call instead "
                "of per-subsection chunks")
        condensed = ingest.condense(chunks)
    else:
        meta["source_mode"] = "toc+ai" if toc_text else "ai"

    review = {"textbook": "Content is rewritten from the uploaded source — "
                          "SMEs verify fidelity to it before issue.",
              "toc+ai": "Content is AI-generated (structure follows the "
                        "supplied TOC) — SMEs must review facts, worked "
                        "examples and references before issue.",
              "ai": "Content is AI-generated — SMEs must review facts, "
                    "worked examples and references before issue."}
    meta.setdefault("level", prompts.DEFAULT_LEVEL)
    report = {"source_mode": meta["source_mode"], "level": meta["level"],
              "model": engine.model,
              "calls": 0, "failures": [], "uk_spelling_fixes": 0,
              "review_note": review[meta["source_mode"]]}
    if source_present:
        report["source"] = {"file": (str(source_path) if source_path
                                      else (source_label
                                            or "course sources")),
                            "chars": len(text), "chunks": len(chunks),
                            "unmatched_headings": []}
        report.setdefault("warnings_extra", []).extend(source_warnings)
    call = _Caller(engine, report, progress)
    n = meta["unit_number"]

    # ---- outline (the one call that must succeed) -------------------------
    out = call("outline",
               prompts.outline(meta, syllabus_topics, toc_text,
                               source_headings=source_headings),
               schemas.OUTLINE_TEXTBOOK if source_headings
               else schemas.OUTLINE)
    if out is None:
        raise GenerationError("outline generation failed twice — cannot "
                              "build a unit without one")

    def source_for(sub):
        """The grounding text for one subsection: its mapped chunks, else
        the condensed whole-source digest. None outside textbook mode."""
        if not source_present:
            return None
        wanted = sub.get("source_headings") or []
        matched, unmatched = ingest.match_chunks(wanted, chunks)
        if unmatched:
            report["source"]["unmatched_headings"].extend(unmatched)
        if matched:
            joined = "\n\n".join(f"## {c['heading']}\n{c['text']}"
                                 for c in matched)
            return joined[:SUB_SOURCE_CAP]
        return condensed
    example_style = out["example_style"]
    for sec in out["sections"]:
        sec["title"] = clean_title(sec["title"])
        for sub in sec["subsections"]:
            sub["title"] = clean_title(sub["title"])

    # ---- code-enforced syllabus coverage ----------------------------------
    # (user requirement: "I don't want any single topic/concept missed.")
    # The outline prompt demands full coverage, but a 7B can still drop a
    # topic — so it is checked HERE: any syllabus topic not represented in
    # the outline (by significant-word overlap) is appended as its own
    # subsection of the best-matching section, guaranteeing content is
    # generated for it. Mirrors the course engine's principle: coverage
    # is code, not trust.
    if syllabus_topics:
        def _tokens(s):
            # singular-ised significant words, so "topics" matches "topic"
            # (applied to BOTH sides, so consistency is what matters)
            return {w[:-1] if len(w) > 4 and w.endswith("s") else w
                    for w in re.findall(r"[a-z0-9]+", s.lower())
                    if len(w) > 3}

        sec_tokens = [
            _tokens(sec["title"] + " "
                    + " ".join(s["title"] for s in sec["subsections"]))
            for sec in out["sections"]]
        all_tokens = set().union(*sec_tokens) if sec_tokens else set()
        injected = []
        for topic in syllabus_topics:
            t = _tokens(topic)
            if not t or len(t & all_tokens) / len(t) >= 0.5:
                continue
            best = max(range(len(out["sections"])),
                       key=lambda i: len(t & sec_tokens[i]))
            out["sections"][best]["subsections"].append(
                {"title": clean_title(topic)})
            sec_tokens[best] |= t
            all_tokens |= t
            injected.append(topic)
        if injected:
            report.setdefault("warnings_extra", []).append(
                "syllabus coverage enforced in code — these topics were "
                "missing from the AI outline and were added as their own "
                "subsections: " + "; ".join(injected))
            progress(f"syllabus coverage: {len(injected)} missed "
                     f"topic(s) injected into the outline")

    titles = "; ".join(s["title"] for s in out["sections"])

    # ---- front matter -----------------------------------------------------
    intro = call("introduction", prompts.introduction(meta, titles),
                 schemas.INTRODUCTION)
    objectives = call("learning objectives",
                      prompts.learning_objectives(meta, titles),
                      schemas.LEARNING_OBJECTIVES)

    # ---- sections ---------------------------------------------------------
    if figures_dir is not None:
        figures_dir = Path(figures_dir)
        figures_dir.mkdir(parents=True, exist_ok=True)
    sections = []
    fig_no = 0
    for si, sec in enumerate(out["sections"], start=1):
        sec_no = f"{n}.{si}"
        sub_titles = "; ".join(s["title"] for s in sec["subsections"])
        extras = call(f"section {sec_no} extras",
                      prompts.section_extras(meta, sec["title"], sub_titles,
                                             source=condensed),
                      schemas.SECTION_EXTRAS)
        problem_seq = 0
        subsections = []
        for ki, sub in enumerate(sec["subsections"]):
            sub_no = f"{sec_no}.{ki + 1}"
            label = f"{sub_no} {sub['title']}"
            blocks = []
            src = source_for(sub)

            p = call(f"{label}: content",
                     prompts.prose(meta, sec["title"], sub["title"],
                                   source=src),
                     schemas.PROSE)
            if p:
                # points-first (user requirement): one lead-in paragraph,
                # then the substance as bullet points — not story prose
                blocks.append({"type": "prose", "text": p["lead_in"]})
                blocks.append({"type": "bullets", "items": p["points"]})

            # enrichment rotation: 1st subsection a table, 2nd a worked
            # example (code or problem per the outline's example_style),
            # 3rd a did-you-know — mirrors the reference sample's rhythm
            if ki == 0:
                t = call(f"{label}: table",
                         prompts.table(meta, sec["title"], sub["title"],
                                       source=src),
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
                             prompts.code(meta, sec["title"], sub["title"],
                                          source=src),
                             schemas.CODE)
                    if c:
                        blocks.append({"type": "code", "text": c["code"]})
                        blocks.append({"type": "prose",
                                       "text": c["explanation"]})
                else:
                    problem_seq += 1
                    pr = call(f"{label}: worked problem",
                              prompts.problem(meta, sec["title"],
                                              sub["title"], source=src),
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
                                              sub["title"], source=src),
                         schemas.DID_YOU_KNOW)
                if d:
                    blocks.append({"type": "did_you_know", "text": d["text"]})

            # figures: with a figures_dir, EVERY subsection gets its own
            # drawn figure (user requirement: more images per topic, with
            # a description under each) — caption+description+diagram in
            # one call, rendered deterministically, placeholder on any
            # failure. Without a figures_dir (legacy/tests), the old
            # extras-based per-section placeholder is kept.
            if figures_dir is not None:
                fig_no += 1
                fs = call(f"{label}: figure",
                          prompts.figure_spec(meta, sec["title"],
                                              sub["title"], source=src),
                          schemas.FIGURE_SPEC)
                if fs:
                    fig_block = {"type": "figure",
                                 "caption": f"Figure {fig_no}: "
                                            f"{clean_title(fs['caption'])}",
                                 "description": fs["description"]}
                    report["fig_planned"] = report.get("fig_planned", 0) + 1
                    png = Path(figures_dir) / f"figure_{fig_no:02d}.png"
                    try:
                        import figure_render
                        figure_render.render(fs, png)
                        fig_block["image"] = str(png)
                        report["fig_rendered"] = (
                            report.get("fig_rendered", 0) + 1)
                    except Exception as e:              # noqa: BLE001
                        progress(f"figure {fig_no} unrenderable ({e}) — "
                                 f"placeholder kept")
                    blocks.append(fig_block)
                else:
                    fig_no -= 1     # failed call: no figure, keep numbering
            elif extras and ki == 0:
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

    if figures_dir is not None:
        report["figures"] = {"planned": report.pop("fig_planned", 0),
                             "rendered": report.pop("fig_rendered", 0)}

    # ---- back matter (grounded on the condensed source in textbook mode) --
    summ = call("summary", prompts.summary(meta, titles, source=condensed),
                schemas.SUMMARY)
    glos = call("glossary", prompts.glossary(meta, titles, source=condensed),
                schemas.GLOSSARY)
    case = call("case study",
                prompts.case_study(meta, titles, source=condensed),
                schemas.CASE_STUDY)
    mcq = call("MCQs", prompts.mcqs(meta, titles, source=condensed),
               schemas.MCQS)
    blanks = call("fill-in-the-blanks",
                  prompts.fill_blanks(meta, titles, source=condensed),
                  schemas.FILL_BLANKS)
    tshort = call("terminal short",
                  prompts.terminal_short(meta, titles, source=condensed),
                  schemas.TERMINAL_SHORT)
    # long questions first, then ONE essay per call — a combined 5-essay
    # generation regularly overruns the engine's token cap (see schemas)
    tlong_qs = call("terminal long questions",
                    prompts.terminal_long_questions(meta, titles,
                                                    source=condensed),
                    schemas.TERMINAL_LONG_QS)
    long_items = []
    for qi, q in enumerate((tlong_qs or {}).get("long_questions", []),
                           start=1):
        ans = call(f"terminal long answer {qi}",
                   prompts.terminal_long_answer(meta, q, source=condensed),
                   schemas.LONG_ANSWER)
        if ans:
            long_items.append({"q": q, "answer": ans["answer"]})
    refs = call("references", prompts.references(meta, titles),
                schemas.REFERENCES, temperature=0.1)

    ref_list = (refs or {}).get("references", [])
    if source_present:
        # spec: in textbook mode the source textbook comes first
        src_kind = ("Source textbook" if source_path
                    else "Source material")
        src_name = (Path(source_path).name if source_path
                    else (source_label or "uploaded course sources"))
        citation = meta.get("textbook_citation") or (
            f"[{src_kind}: {src_name} — complete "
            f"citation to be added by the SME.]")
        ref_list = [citation] + [r for r in ref_list if r != citation]

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
                     "long": long_items},
        "references": ref_list,
    }

    report["uk_spelling_fixes"] = apply_to_unit(unit)
    relevance_flags = _relevance_audit(unit, syllabus_topics, call)
    report["relevance"] = {
        "flags": len(relevance_flags),
        "details": relevance_flags,
        "note": ("AI relevance audit: every section and the assessment "
                 "material were re-read against the stated course, unit "
                 "and syllabus topics. Flags are surfaced as warnings for "
                 "SME judgement — nothing is auto-deleted, because a "
                 "false positive would silently remove valid teaching "
                 "content."),
    }
    errors, warnings = validate_unit(unit)
    for f in relevance_flags:
        warnings.append(f'relevance ({f["scope"]}): "{f["item"]}" — '
                        f'{f["why"]}')
    warnings.extend(report.pop("warnings_extra", []))
    if source_present and report["source"]["unmatched_headings"]:
        warnings.append(
            "some outline subsections could not be matched to source "
            "sections and fell back to the whole-source digest: "
            + "; ".join(report["source"]["unmatched_headings"]))
    report["validation"] = {"errors": errors, "warnings": warnings}
    return unit, report


def _finish_outputs(unit, report, out_docx, json_out=None, report_path=None,
                    brand="ppsu"):
    """Shared tail for single and batch runs: write report/json/figure list,
    render only if validation passed. Returns True when rendered."""
    import figures
    report_path = report_path or out_docx.with_suffix(".report.json")
    report_path.write_text(json.dumps(report, indent=2), encoding="utf-8")
    if json_out:
        json_out.write_text(json.dumps(unit, indent=2), encoding="utf-8")
    out_docx.with_suffix(".figures.txt").write_text(
        figures.figure_list_text(unit), encoding="utf-8")

    if report["validation"]["errors"]:
        print(f"\nNOT RENDERED — validation errors (report: {report_path}):")
        for e in report["validation"]["errors"]:
            print(f"  - {e}")
        return False
    from docx_builder import build
    build(unit, brand=brand).save(str(out_docx))
    print(f"\nwrote {out_docx}  ({report['calls']} AI calls, "
          f"{len(report['failures'])} failed, "
          f"{report['uk_spelling_fixes']} UK-spelling fixes)")
    for w in report["validation"]["warnings"]:
        print(f"  warning: {w}")
    return True


def run_batch(batch_path: Path, out_dir: Path):
    """Phase 6 multi-unit batch: batch.json is a list of
    {"meta": {...with syllabus_topics}, "toc": path|null,
     "source": path|null}. Units generate sequentially (one GPU); one
    unit's failure never stops the rest."""
    entries = json.loads(batch_path.read_text(encoding="utf-8"))
    out_dir.mkdir(parents=True, exist_ok=True)
    results = []
    for i, entry in enumerate(entries, 1):
        meta = dict(entry["meta"])
        syllabus = meta.pop("syllabus_topics", None)
        toc_text = (Path(entry["toc"]).read_text(encoding="utf-8")
                    if entry.get("toc") else None)
        source = Path(entry["source"]) if entry.get("source") else None
        stem = f'unit{meta.get("unit_number", i):02d}'
        print(f"\n===== batch {i}/{len(entries)}: {stem} "
              f'{meta.get("unit_title", "")} =====')
        try:
            unit, report = generate_unit(
                meta, syllabus_topics=syllabus, toc_text=toc_text,
                source_path=source,
                figures_dir=out_dir / f"{stem}_figures")
            ok = _finish_outputs(unit, report, out_dir / f"{stem}.docx",
                                 json_out=out_dir / f"{stem}.json")
            results.append((stem, "ok" if ok else "validation failed"))
        except Exception as e:                          # noqa: BLE001
            print(f"  {stem} FAILED: {e}")
            results.append((stem, f"failed: {e}"))
    print("\n===== batch summary =====")
    for stem, status in results:
        print(f"  {stem}: {status}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--batch", type=Path,
                    help="batch JSON (list of {meta, toc?, source?}) — "
                         "generates every unit into --out-dir")
    ap.add_argument("--out-dir", type=Path, default=Path("output"))
    ap.add_argument("--meta", type=Path,
                    help="JSON: programme, course_code, course_name, "
                         "unit_number, unit_title, syllabus_topics[]")
    ap.add_argument("--toc", type=Path,
                    help="optional textbook table-of-contents text file "
                         "(activates toc+ai mode)")
    ap.add_argument("--source", type=Path,
                    help="optional textbook chapter (.pdf/.docx/.txt) — "
                         "activates textbook mode (overrides --toc for "
                         "mode selection; both still inform the outline)")
    ap.add_argument("--brand", choices=("ppsu", "reva"), default="ppsu",
                    help="target brand format (see brands.py)")
    ap.add_argument("--out", type=Path, help=".docx path (single-unit mode)")
    ap.add_argument("--report", type=Path)
    ap.add_argument("--json-out", type=Path,
                    help="also save the intermediate unit JSON")
    args = ap.parse_args()

    if args.batch:
        run_batch(args.batch, args.out_dir)
        return
    if not args.meta or not args.out:
        ap.error("--meta and --out are required (or use --batch)")

    meta = json.loads(args.meta.read_text(encoding="utf-8"))
    syllabus = meta.pop("syllabus_topics", None)
    toc_text = args.toc.read_text(encoding="utf-8") if args.toc else None

    t0 = time.time()
    unit, report = generate_unit(
        meta, syllabus_topics=syllabus, toc_text=toc_text,
        source_path=args.source,
        figures_dir=args.out.parent / f"{args.out.stem}_figures")
    report["seconds"] = round(time.time() - t0, 1)
    ok = _finish_outputs(unit, report, args.out, json_out=args.json_out,
                         report_path=args.report, brand=args.brand)
    if not ok:
        sys.exit(1)


if __name__ == "__main__":
    main()
