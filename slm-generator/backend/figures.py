"""figures.py — Phase 6: the DTP team's figure work list. Figures are now
auto-rendered as diagrams (figure_render.py) and embedded in the docx;
any figure whose diagram could not be generated stays a grey placeholder
box. This list gives the DTP team every figure with its location, caption
and status — auto-rendered diagrams may still be replaced with better
artwork if desired, placeholders MUST be drawn."""


def figure_list(unit):
    """[{section, subsection, caption, rendered}] in document order."""
    out = []
    for sec in unit.get("sections", []):
        for sub in sec.get("subsections", []):
            for block in sub.get("blocks", []):
                if block.get("type") == "figure":
                    out.append({
                        "section": f'{sec["number"]} {sec["title"]}',
                        "subsection": f'{sub["number"]} {sub["title"]}',
                        "caption": block.get("caption", ""),
                        "rendered": bool(block.get("image")),
                    })
    return out


def figure_list_text(unit) -> str:
    """The list as a plain-text handoff file for the DTP team."""
    meta = unit.get("meta", {})
    lines = [
        f"FIGURES — Unit {meta.get('unit_number')}: "
        f"{meta.get('unit_title', '')}",
        f"{meta.get('course_code', '')} {meta.get('course_name', '')}",
        "Figures marked [auto-rendered] are embedded in the .docx as "
        "generated diagrams (replace with better artwork if desired). "
        "Figures marked [PLACEHOLDER] appear as grey boxes and MUST be "
        "drawn; keep the caption line beneath the artwork.",
        "",
    ]
    figs = figure_list(unit)
    if not figs:
        lines.append("(no figures in this unit)")
    for i, f in enumerate(figs, 1):
        status = "auto-rendered" if f["rendered"] else "PLACEHOLDER"
        lines += [f"{i}. [{status}] {f['caption']}",
                  f"   Location: {f['section']}  ->  {f['subsection']}",
                  ""]
    return "\n".join(lines)
