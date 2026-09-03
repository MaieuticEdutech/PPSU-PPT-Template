"""brands.py — built-in brand profiles the generator can target.

"ppsu" is the decoded P P Savani format Phase 1 was built against.
"reva" is transcribed EXPLICITLY from "REVA University Online – SLM Style
Guide" (PDF supplied 2026-09-03) — every value below traces to a line of
that guide; comments cite the rule. Where the guide is silent or
ambiguous, the choice is stated.

A brand profile is the BASE theme; an uploaded reference-PDF style profile
(branding.py) and logo still layer on top of it.
"""

PPSU = {
    "key": "ppsu",
    "font": "Calibri",
    "title_pt": 24, "h1_pt": 20, "h2_pt": 15, "h3_pt": 13,
    "body_pt": 11, "caption_pt": 10,
    "heading_color": "0E2841", "body_color": None,      # default black
    # per-level heading bars: none for PPSU (plain coloured text)
    "h_fills": {},
    "h_text_colors": {},
    "heading_upper": False, "justify_body": False, "line_spacing": None,
    "did_you_know_label": "Did you know?", "fill_did_you_know": "E7E6E6",
    "think_apply_label": "Think and Apply", "fill_think_apply": "FCE4D6",
    "fill_problem": "FCE4D6", "fill_code": "F2F2F2",
    "table_header_fill": "0E2841", "table_header_text": "FFFFFF",
    "table_alt_row": None, "table_border": None,
    "fig_prefix": "Figure", "caption_center": False,
    "summary_box": None, "glossary_fill": None,
    "case_box": None, "terminal_box": None,
    "header_layout": "logo_left",            # logo left, unit label right
    "header_text": "unit",                   # "Unit NN"
    "hf_border_color": None,
    "footer_pt": 9, "footer_bold": True, "footer_color": "0E2841",
    "references_title": "References",
    "strict_filename": False,
}

REVA = {
    "key": "reva",
    # "Font Style: Plus Jakarta Sans (all elements)" — must be installed on
    # the authoring machine or Word substitutes silently.
    "font": "Plus Jakarta Sans",
    # Typography Hierarchy table:
    #   Topic Name 30 Bold Left 1.5 | Topic (Heading) 18 Bold UPPERCASE
    #   Subtopic 16 Bold UPPERCASE | Sub of Subtopic 14 Bold UPPERCASE
    #   Body 12 Regular 1.5 Justify | Figures/Tables caption 11
    "title_pt": 30, "h1_pt": 18, "h2_pt": 16, "h3_pt": 14,
    "body_pt": 12, "caption_pt": 11,
    "heading_color": "F7A35B",               # REVA Orange
    "body_color": "333333",                  # Dark Charcoal body text
    # Element-wise Color Mapping:
    #   Topic Heading: fill #F7A35B, text #FFFFFF
    #   Subtopic Heading: fill #FEF0E6, text #4A4C55
    #   Sub of Subtopic: fill #F2F2F2, text #4A4C55
    "h_fills": {1: "F7A35B", 2: "FEF0E6", 3: "F2F2F2"},
    "h_text_colors": {1: "FFFFFF", 2: "4A4C55", 3: "4A4C55"},
    "heading_upper": True,                   # Bold, Uppercase, Justify
    "justify_body": True, "line_spacing": 1.5,
    # REVA's aside boxes: Study Note (#FFD966 fill) and Activity (#A9D18E),
    # both #333333 text -> our did_you_know / think_and_apply become them
    "did_you_know_label": "Study Note", "fill_did_you_know": "FFD966",
    "think_apply_label": "Activity", "fill_think_apply": "A9D18E",
    # no worked-problem box in the guide; Light Orange Tint is a permitted
    # colour ("alternate table row shading, summary box fill") — used here
    "fill_problem": "FEF0E6",
    "fill_code": "F2F2F2",                   # Light Grey (permitted)
    # Table Header Row #F7A35B/white; Alternate Row #F2F2F2/#333333;
    # Table Border #CCCCCC
    "table_header_fill": "F7A35B", "table_header_text": "FFFFFF",
    "table_alt_row": "F2F2F2", "table_border": "CCCCCC",
    # Caption examples: "Fig. 1: Concept of Knowledge", 11pt, Center,
    # "Fig No. & Name Bold"
    "fig_prefix": "Fig.", "caption_center": True,
    # box fills: Summary #FEF0E6 (#4A4C55 text), Glossary #F2F2F2,
    # Case Study #F2F2F2, Terminal Questions #F2F2F2 (all #333333 text)
    "summary_box": "FEF0E6", "glossary_fill": "F2F2F2",
    "case_box": "F2F2F2", "terminal_box": "F2F2F2",
    # Header: "Unit Name | Unit Number + REVA logo (right)";
    # Footer: "Course Code | Course Name + Page Number (right)", 10 Bold;
    # header/footer border lines REVA Orange
    "header_layout": "logo_right",
    "header_text": "unit_name",              # "Unit Name | Unit N"
    "hf_border_color": "F7A35B",
    "footer_pt": 10, "footer_bold": True, "footer_color": "4A4C55",
    # SLM Structure: "... Summary, Suggested Books and References"
    "references_title": "Suggested Books and References",
    # File Naming: CourseName_UnitNumber_DocumentTitle, CamelCase,
    # zero-padded UnitNN, underscores only
    "strict_filename": True,
}

PROFILES = {"ppsu": PPSU, "reva": REVA}


def reva_filename(meta, ext):
    """REVA naming convention: CourseName_UnitNN_DocumentTitle.ext —
    CamelCase, no spaces, underscore separators, zero-padded unit."""
    def camel(text):
        words = "".join(c if c.isalnum() or c.isspace() else " "
                        for c in text).split()
        return "".join(w[:1].upper() + w[1:] for w in words) or "SLM"
    unit_no = meta.get("unit_number", 1)
    n = f"{unit_no:02d}" if isinstance(unit_no, int) else str(unit_no)
    return (f"{camel(meta.get('course_name', 'Course'))}_Unit{n}_"
            f"{camel(meta.get('unit_title', 'Unit'))}{ext}")
