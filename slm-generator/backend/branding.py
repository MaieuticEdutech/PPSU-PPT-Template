"""branding.py — institution branding assets, set once and applied to every
generated unit until replaced:

- logo image  -> embedded on the cover and in every page header
- reference SLM PDF -> its STYLE SIGNALS are extracted (dominant font
  family, body/heading point sizes, heading colour, page size, margins) and
  override the builder's defaults; the PDF itself is kept for side-by-side
  review. This is honest style adaptation, not pixel-cloning: the document
  STRUCTURE always follows the decoded PPSU SLM format.

Everything degrades gracefully: no assets -> builder defaults; a PDF the
extractor can't read -> defaults + the problem reported to the caller.
"""
import json
import re
from collections import Counter
from pathlib import Path

ASSETS = Path(__file__).parent.parent / "assets"
LOGO_STEMS = ("logo.png", "logo.jpg", "logo.jpeg")
PROFILE_PATH = ASSETS / "style_profile.json"
REFERENCE_PDF = ASSETS / "reference.pdf"


def logo_path():
    for name in LOGO_STEMS:
        p = ASSETS / name
        if p.is_file():
            return p
    return None


def _docx_can_parse(data: bytes) -> bool:
    from io import BytesIO
    from docx.image.image import Image as DocxImage
    try:
        DocxImage.from_blob(data)
        return True
    except Exception:
        return False


def save_logo(data: bytes, suffix: str) -> Path:
    """Store the logo, guaranteeing python-docx can embed it. A file that
    docx parses as-is is kept byte-identical (preserves PNG transparency);
    otherwise it is re-encoded to a clean PNG via PyMuPDF — a real logo
    JPEG whose header carried an XMP-only APP1 segment (no JFIF/EXIF
    marker) raised UnrecognizedImageError at render time and killed the
    whole generation. Raises ValueError when the bytes aren't a readable
    image at all."""
    if not _docx_can_parse(data):
        import fitz
        try:
            img = fitz.open(stream=data, filetype="image")
            data = img[0].get_pixmap().tobytes("png")
            img.close()
        except Exception:
            raise ValueError(
                "the logo file could not be read as an image — please "
                "export it as a standard PNG or JPG and upload again")
        if not _docx_can_parse(data):
            raise ValueError("the logo could not be converted to a "
                             "Word-embeddable image")
        suffix = ".png"
    ASSETS.mkdir(exist_ok=True)
    for name in LOGO_STEMS:          # replace whatever format was there
        (ASSETS / name).unlink(missing_ok=True)
    p = ASSETS / f"logo{suffix.lower()}"
    p.write_bytes(data)
    return p


def load_profile():
    if PROFILE_PATH.is_file():
        try:
            return json.loads(PROFILE_PATH.read_text(encoding="utf-8"))
        except Exception:
            return None
    return None


def _clean_font_name(raw: str) -> str:
    """'ABCDEE+Calibri-Bold' -> 'Calibri'."""
    name = raw.split("+")[-1]
    name = re.split(r"[-,]", name)[0]
    # split CamelCase compounds like 'TimesNewRomanPSMT' conservatively:
    # only strip well-known style suffixes
    for suffix in ("Bold", "Italic", "BoldItalic", "Regular", "PSMT", "PS"):
        if name.endswith(suffix) and len(name) > len(suffix) + 2:
            name = name[: -len(suffix)]
    return re.sub(r"(?<!^)(?=[A-Z][a-z])", " ", name).strip()


def extract_profile(pdf_path) -> dict:
    """Style signals from a reference SLM PDF via PyMuPDF. Raises ValueError
    with a readable message if the PDF has no extractable text."""
    import fitz
    doc = fitz.open(str(pdf_path))
    sizes, fonts, colors_by_size = Counter(), Counter(), {}
    page_rect = None
    text_left, text_top, text_right, text_bottom = [], [], [], []

    # skip the cover (page 0) — its display type sizes would skew the stats
    pages = list(doc)[1:12] or list(doc)[:1]
    for page in pages:
        page_rect = page.rect
        d = page.get_text("dict")
        for block in d.get("blocks", []):
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    txt = span.get("text", "").strip()
                    if not txt:
                        continue
                    size = round(span["size"])
                    sizes[size] += len(txt)
                    fonts[_clean_font_name(span.get("font", ""))] += len(txt)
                    colors_by_size.setdefault(size, Counter())[
                        span.get("color", 0)] += len(txt)
                    x0, y0, x1, y1 = span["bbox"]
                    text_left.append(x0)
                    text_top.append(y0)
                    text_right.append(x1)
                    text_bottom.append(y1)
    doc.close()
    if not sizes:
        raise ValueError("the reference PDF has no extractable text "
                         "(scanned images only?) — style cannot be read "
                         "from it")

    body_pt = sizes.most_common(1)[0][0]
    font_name = fonts.most_common(1)[0][0] or None

    # heading colour: the most common NON-black colour among text clearly
    # larger than the body (the navy section headings in the real sample)
    heading_colors = Counter()
    heading_sizes = [s for s in sizes if s >= body_pt + 3]
    for s in heading_sizes:
        for color, n in colors_by_size.get(s, {}).items():
            if color != 0:
                heading_colors[color] += n
    heading_color = (f"{heading_colors.most_common(1)[0][0]:06X}"
                     if heading_colors else None)
    h1_pt = max(heading_sizes) if heading_sizes else None

    profile = {"font_name": font_name, "body_pt": body_pt, "h1_pt": h1_pt,
               "heading_color": heading_color}
    if page_rect is not None and text_left:
        # PDF points -> EMU-friendly mm (1 pt = 0.3528 mm)
        def mm(v):
            return round(v * 0.352778, 1)
        profile.update({
            "page_width_mm": mm(page_rect.width),
            "page_height_mm": mm(page_rect.height),
            "margin_left_mm": mm(max(min(text_left), 0)),
            "margin_top_mm": mm(max(min(text_top), 0)),
            "margin_right_mm": mm(max(page_rect.width - max(text_right), 0)),
            "margin_bottom_mm": mm(max(page_rect.height - max(text_bottom),
                                       0)),
        })
    return profile


def save_reference_pdf(data: bytes) -> dict:
    """Store the PDF, extract + persist its style profile, return it."""
    ASSETS.mkdir(exist_ok=True)
    REFERENCE_PDF.write_bytes(data)
    profile = extract_profile(REFERENCE_PDF)
    PROFILE_PATH.write_text(json.dumps(profile, indent=2), encoding="utf-8")
    return profile


def clear():
    for p in [PROFILE_PATH, REFERENCE_PDF] + [ASSETS / n for n in LOGO_STEMS]:
        p.unlink(missing_ok=True)
