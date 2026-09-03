"""figures.py — Phase 6: the DTP team's work list. The docx renders every
figure as a labelled placeholder box; this extracts one flat list of every
placeholder with its exact location and caption so the designers know what
to draw and where to drop it in."""


def figure_list(unit):
    """[{section, subsection, caption}] in document order."""
    out = []
    for sec in unit.get("sections", []):
        for sub in sec.get("subsections", []):
            for block in sub.get("blocks", []):
                if block.get("type") == "figure":
                    out.append({
                        "section": f'{sec["number"]} {sec["title"]}',
                        "subsection": f'{sub["number"]} {sub["title"]}',
                        "caption": block.get("caption", ""),
                    })
    return out


def figure_list_text(unit) -> str:
    """The list as a plain-text handoff file for the DTP team."""
    meta = unit.get("meta", {})
    lines = [
        f"FIGURE PLACEHOLDERS — Unit {meta.get('unit_number')}: "
        f"{meta.get('unit_title', '')}",
        f"{meta.get('course_code', '')} {meta.get('course_name', '')}",
        "Each figure below appears in the .docx as a grey placeholder box "
        "at the stated location. Replace the box with the artwork; keep "
        "the caption line beneath it.",
        "",
    ]
    figs = figure_list(unit)
    if not figs:
        lines.append("(no figure placeholders in this unit)")
    for i, f in enumerate(figs, 1):
        lines += [f"{i}. {f['caption']}",
                  f"   Location: {f['section']}  ->  {f['subsection']}",
                  ""]
    return "\n".join(lines)
