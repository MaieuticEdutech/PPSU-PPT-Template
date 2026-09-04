#!/usr/bin/env python3
"""course_generator.py — full-course mode: N uploaded source documents
(e.g. an existing 5-unit SLM set) re-partitioned into a TARGET number of
units (e.g. REVA's 14), generating each unit with the existing textbook-
mode pipeline grounded on its own slice of the combined sources.

THE COVERAGE GUARANTEE IS CODE, NOT TRUST. The AI only proposes where each
unit ENDS (cut points over the ordered topic list) and what to call it;
the topics themselves are assigned by contiguous slicing in code, so every
heading from every source lands in exactly one unit — nothing can be
skipped even if the model's plan is sloppy (invalid/non-monotonic cut
points are repaired to an equal split; the final cut is always forced to
the last topic). Contiguous slicing also preserves the sources' teaching
sequence (global_rules: preserve the logical teaching sequence).
"""
import json
import zipfile
from pathlib import Path

import ingest
import prompts
from ai_engine import get_engine
from unit_generator import (GenerationError, _Caller, _default_progress,
                            clean_title, generate_unit)


def _plan_schema(n):
    return {
        "type": "object",
        "properties": {
            "units": {
                "type": "array", "minItems": n, "maxItems": n,
                "items": {
                    "type": "object",
                    "properties": {
                        "title": {"type": "string", "minLength": 3},
                        "end_index": {"type": "integer", "minimum": 0},
                    },
                    "required": ["title", "end_index"],
                },
            },
        },
        "required": ["units"],
    }


def _repair_cuts(units_plan, n_topics, target_units):
    """Cut points must be strictly increasing and end at the last topic;
    anything else is repaired. Returns the list of inclusive end indices,
    len == target_units, last == n_topics - 1."""
    raw = [u.get("end_index", -1) for u in units_plan]
    cuts, prev = [], -1
    usable = True
    for i, c in enumerate(raw):
        remaining_after = target_units - 1 - i
        max_allowed = n_topics - 1 - remaining_after
        if not isinstance(c, int) or c <= prev or c > max_allowed:
            usable = False
            break
        cuts.append(c)
        prev = c
    if usable and cuts and cuts[-1] != n_topics - 1:
        cuts[-1] = n_topics - 1
        usable = cuts == sorted(set(cuts))
    if not usable or len(cuts) != target_units:
        # equal split fallback — coverage always wins over the model's plan
        base, rem = divmod(n_topics, target_units)
        cuts, acc = [], 0
        for i in range(target_units):
            acc += base + (1 if i < rem else 0)
            cuts.append(acc - 1)
    return cuts


def plan_course(meta, all_chunks, target_units, call):
    """(plan, repaired) — plan is [{title, chunks:[...]}] covering every
    chunk exactly once, in source order."""
    topics = [c for c in all_chunks if c["heading"] != "(preamble)"]
    if len(topics) < target_units:
        raise GenerationError(
            f"the sources contain only {len(topics)} detectable topics — "
            f"cannot split them into {target_units} units (need at least "
            f"one topic per unit)")
    numbered = "\n".join(f"{i}. {c['heading']}  [{c['source']}]"
                         for i, c in enumerate(topics))
    out = call("course plan",
               prompts.course_plan(meta, numbered, target_units),
               _plan_schema(target_units), num_predict=4096)
    if out is None:
        raise GenerationError("course planning failed twice — cannot "
                              "partition the sources into units")

    cuts = _repair_cuts(out["units"], len(topics), target_units)
    repaired = cuts != [u.get("end_index") for u in out["units"]]
    plan, start = [], 0
    for i, end in enumerate(cuts):
        title = clean_title(out["units"][i].get("title", "")) \
            or f"Unit {i + 1}"
        plan.append({"title": title, "chunks": topics[start:end + 1]})
        start = end + 1
    return plan, repaired


def generate_course(meta_base, source_paths, target_units=14,
                    toc_text=None, engine=None,
                    progress=_default_progress, figures_root=None):
    """Returns (results, course_report). results[i] = {unit_number, title,
    unit (dict|None), report (dict|None), error (str|None)}. One unit's
    failure never stops the rest."""
    engine = engine or get_engine()

    # ---- ingest every source, keeping order + provenance ------------------
    all_chunks, per_source = [], []
    for path in source_paths:
        path = Path(path)
        text = ingest.extract_text(path)
        chunks = ingest.chunk_by_headings(text)
        for c in chunks:
            c["source"] = path.name
        n_topics = sum(1 for c in chunks if c["heading"] != "(preamble)")
        per_source.append({"file": path.name, "chars": len(text),
                           "topics": n_topics})
        all_chunks.extend(chunks)

    course_report = {"sources": per_source, "target_units": target_units,
                     "plan_repaired": False, "units": [],
                     "calls": 0, "failures": []}
    call = _Caller(engine, course_report, progress)

    plan, repaired = plan_course(meta_base, all_chunks, target_units, call)
    course_report["plan_repaired"] = repaired
    course_report["plan"] = [
        {"unit_number": i + 1, "title": p["title"],
         "topics": [c["heading"] for c in p["chunks"]]}
        for i, p in enumerate(plan)]

    # code-enforced coverage proof, recorded for the reviewer
    assigned = [c["heading"] for p in plan for c in p["chunks"]]
    all_topics = [c["heading"] for c in all_chunks
                  if c["heading"] != "(preamble)"]
    course_report["coverage"] = {
        "topics_total": len(all_topics),
        "topics_assigned": len(assigned),
        "complete": assigned == all_topics,
    }

    results = []
    for i, p in enumerate(plan, start=1):
        meta = {**meta_base, "unit_number": i, "unit_title": p["title"]}

        def unit_progress(msg, _i=i):
            progress(f"[unit {_i}/{target_units}] {msg}")

        try:
            unit, report = generate_unit(
                meta, toc_text=toc_text, source_chunks=p["chunks"],
                source_label=f"course sources (unit {i} slice)",
                engine=engine, progress=unit_progress,
                figures_dir=(Path(figures_root) / f"unit{i:02d}"
                             if figures_root else None))
            errors = report["validation"]["errors"]
            results.append({"unit_number": i, "title": p["title"],
                            "unit": unit if not errors else None,
                            "report": report,
                            "error": ("validation: " + "; ".join(errors))
                                     if errors else None})
        except Exception as e:                          # noqa: BLE001
            results.append({"unit_number": i, "title": p["title"],
                            "unit": None, "report": None,
                            "error": f"{e.__class__.__name__}: {e}"})
        course_report["units"].append(
            {"unit_number": i, "title": p["title"],
             "status": "ok" if results[-1]["unit"] else "failed",
             "error": results[-1]["error"]})
    return results, course_report


def write_course_outputs(results, course_report, out_dir, brand="ppsu",
                         make_pdf=True, progress=_default_progress):
    """Render every successful unit (docx [+pdf] + report + figure list)
    into out_dir and zip the lot. Returns the zip path."""
    import brands
    import docx2pdf
    import figures
    from docx_builder import build

    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    files = []

    def name_for(res, ext):
        meta = res["unit"]["meta"]
        if brand == "reva":
            return brands.reva_filename(meta, ext)
        return f'unit{res["unit_number"]:02d}{ext}'

    pdf_ok = make_pdf and docx2pdf.available()
    for res in results:
        n = res["unit_number"]
        if res["report"] is not None:
            rp = out_dir / f"unit{n:02d}.report.json"
            rp.write_text(json.dumps(res["report"], indent=2),
                          encoding="utf-8")
            files.append(rp)
        if res["unit"] is None:
            continue
        progress(f"[unit {n}] rendering docx"
                 + (" + pdf" if pdf_ok else ""))
        docx_path = out_dir / name_for(res, ".docx")
        build(res["unit"], brand=brand).save(str(docx_path))
        files.append(docx_path)
        figs = out_dir / f"unit{n:02d}.figures.txt"
        figs.write_text(figures.figure_list_text(res["unit"]),
                        encoding="utf-8")
        files.append(figs)
        if pdf_ok:
            pdf_path = out_dir / name_for(res, ".pdf")
            if docx2pdf.convert(docx_path, pdf_path):
                files.append(pdf_path)

    crp = out_dir / "course_report.json"
    crp.write_text(json.dumps(course_report, indent=2), encoding="utf-8")
    files.append(crp)

    zip_path = out_dir / "course_units.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in files:
            zf.write(f, f.name)
    return zip_path
