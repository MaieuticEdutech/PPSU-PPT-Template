"""validate_unit.py — the gate between generation and rendering (CLAUDE.md:
"A unit JSON failing schema/count checks never reaches the builder; the
report says which block failed and why").

Two severities:
- errors  -> the unit must NOT be rendered (structural schema failure, or a
             count so wrong the issued format is violated)
- warnings -> render anyway, but surface to the reviewer (e.g. a block the
             model failed twice on and was skipped)
"""
import jsonschema

import schemas


def validate_unit(unit):
    """Returns (errors, warnings) — both lists of human-readable strings."""
    errors, warnings = [], []

    try:
        jsonschema.validate(unit, schemas.FULL_UNIT)
    except jsonschema.ValidationError as e:
        errors.append(f"unit schema: {e.message} (at "
                      f"{'/'.join(str(p) for p in e.absolute_path) or 'root'})")
        return errors, warnings   # structure is broken; counts are moot

    n_obj = len(unit["learning_objectives"])
    if not 4 <= n_obj <= 5:
        errors.append(f"learning objectives: {n_obj} (need 4-5)")

    sa = unit.get("self_assessment", {})
    n_mcq = len(sa.get("mcq", []))
    if n_mcq != 8:
        (errors if n_mcq == 0 else warnings).append(
            f"MCQs: {n_mcq} (target 8)")
    n_blanks = len(sa.get("fill_blanks", []))
    if n_blanks not in (0, 5):
        warnings.append(f"fill-in-the-blanks: {n_blanks} (target 5)")
    elif n_blanks == 0:
        warnings.append("fill-in-the-blanks: none generated")

    term = unit.get("terminal", {})
    for key in ("short", "long"):
        n = len(term.get(key, []))
        if n != 5:
            (errors if n == 0 else warnings).append(
                f"terminal {key} questions: {n} (target 5)")

    for section in unit["sections"]:
        for sub in section["subsections"]:
            if not sub.get("blocks"):
                errors.append(f"subsection {sub.get('number')} "
                              f"'{sub.get('title')}' has no content blocks")
            elif not any(b["type"] == "prose" for b in sub["blocks"]):
                warnings.append(f"subsection {sub.get('number')} has no "
                                f"prose (only boxed/annex blocks)")

    if len(unit.get("glossary", [])) < 10:
        warnings.append(f"glossary: {len(unit.get('glossary', []))} terms "
                        f"(sample units carry ~16)")
    if len(unit.get("references", [])) < 5:
        warnings.append(f"references: {len(unit.get('references', []))} "
                        f"(target 6-8)")
    if unit.get("references"):
        warnings.append("references are AI-suggested and must be verified "
                        "by an SME — local models can cite non-existent "
                        "editions")

    return errors, warnings
