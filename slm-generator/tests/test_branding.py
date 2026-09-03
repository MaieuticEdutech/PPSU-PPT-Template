#!/usr/bin/env python3
"""Tests for the branding feature: reference-PDF style extraction, logo
round-trip, builder integration (logo embedded, page geometry + fonts
applied, defaults restored), and the API endpoints.

Run: python tests/test_branding.py
"""
import io
import json
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "backend"))

import tempfile

import fitz

import branding

# ISOLATION: redirect the branding store to a throwaway dir BEFORE anything
# touches it. An earlier version of this suite ran against the real
# slm-generator/assets/ and its cleanup deleted the user's actually-uploaded
# logo and reference PDF. Never again.
_TMP_ASSETS = Path(tempfile.mkdtemp(prefix="slm_test_assets_"))
branding.ASSETS = _TMP_ASSETS
branding.PROFILE_PATH = _TMP_ASSETS / "style_profile.json"
branding.REFERENCE_PDF = _TMP_ASSETS / "reference.pdf"

import styles
from docx_builder import build

passed = failed = 0


def check(name, condition):
    global passed, failed
    if condition:
        passed += 1
        print(f"  OK   {name}")
    else:
        failed += 1
        print(f"  FAIL {name}")


GOLDEN = json.loads((Path(__file__).parent / "golden_unit1.json")
                    .read_text(encoding="utf-8"))


def make_reference_pdf() -> bytes:
    """A synthetic reference SLM: cover page + 2 content pages with a known
    style — Helvetica, 11pt black body, 20pt coloured headings."""
    doc = fitz.open()
    cover = doc.new_page(width=595, height=842)          # A4 portrait
    cover.insert_text((200, 400), "COVER", fontsize=40)
    for _ in range(2):
        page = doc.new_page(width=595, height=842)
        page.insert_text((72, 80), "1.1 A Section Heading", fontsize=20,
                         color=(14 / 255, 40 / 255, 65 / 255))
        y = 120
        for i in range(12):
            page.insert_text(
                (72, y), "Body text line with plenty of characters to "
                         "dominate the statistics.", fontsize=11)
            y += 20
    out = doc.tobytes()
    doc.close()
    return out


def make_logo_png() -> bytes:
    pix = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, 120, 60))
    pix.clear_with(90)
    return pix.tobytes("png")


def docx_media(document):
    buf = io.BytesIO()
    document.save(buf)
    with zipfile.ZipFile(io.BytesIO(buf.getvalue())) as z:
        return [n for n in z.namelist() if n.startswith("word/media/")]


branding.clear()
try:
    print("=== extract_profile ===")
    ref_pdf = make_reference_pdf()
    tmp = Path(__file__).parent.parent / "output" / "_test_ref.pdf"
    tmp.parent.mkdir(exist_ok=True)
    tmp.write_bytes(ref_pdf)
    prof = branding.extract_profile(tmp)
    check("body size detected (11pt)", prof["body_pt"] == 11)
    check("heading size + colour detected",
          prof["h1_pt"] == 20 and prof["heading_color"] == "0E2841")
    check("font family cleaned", prof["font_name"] and
          "+" not in prof["font_name"])
    check("A4 page geometry detected",
          abs(prof["page_width_mm"] - 210) < 2
          and abs(prof["page_height_mm"] - 297) < 2)

    print("\n=== builder WITHOUT branding (defaults) ===")
    doc = build(GOLDEN, use_branding=False)
    check("no images embedded by default", docx_media(doc) == [])

    print("\n=== builder WITH branding ===")
    branding.save_logo(make_logo_png(), ".png")
    branding.save_reference_pdf(ref_pdf)
    doc = build(GOLDEN)
    media = docx_media(doc)
    check("logo embedded (cover + header)", len(media) >= 1)
    check("page geometry applied (A4 width)",
          abs(doc.sections[0].page_width.mm - 210) < 2)
    check("style override active during build (font from reference)",
          styles.BRAND_FONT == prof["font_name"])

    print("\n=== defaults restored for unbranded builds ===")
    build(GOLDEN, use_branding=False)
    check("apply_profile(None) resets font/colour defaults",
          styles.BRAND_FONT == "Calibri"
          and styles.NAVY == styles._DEFAULTS["NAVY"])

    print("\n=== logo normalisation (the UnrecognizedImageError incident) ===")
    # a REAL, valid JPEG whose only APP1 segment is XMP (no JFIF/EXIF
    # marker) — python-docx cannot parse this header shape and killed a
    # live generation with UnrecognizedImageError
    plain_jpeg = fitz.open(stream=make_logo_png(),
                           filetype="image")[0].get_pixmap().tobytes("jpeg")
    assert plain_jpeg[:2] == b"\xff\xd8"
    xmp_payload = b"http://ns.adobe.com/xap/1.0/\x00<x:xmpmeta/>"
    app1 = (b"\xff\xe1" + (len(xmp_payload) + 2).to_bytes(2, "big")
            + xmp_payload)
    xmp_jpeg = plain_jpeg[:2] + app1 + plain_jpeg[2:]
    check("repro: docx cannot parse an XMP-only-APP1 jpeg",
          not branding._docx_can_parse(xmp_jpeg))
    saved = branding.save_logo(xmp_jpeg, ".jpeg")
    check("save_logo re-encodes it to a docx-embeddable PNG",
          saved.name == "logo.png"
          and branding._docx_can_parse(saved.read_bytes()))
    try:
        branding.save_logo(b"this is not an image at all", ".png")
        check("garbage bytes rejected with a readable error", False)
    except ValueError as e:
        check("garbage bytes rejected with a readable error",
              "could not be read" in str(e))

    print("\n=== builder belt: a bad on-disk logo never kills a build ===")
    branding.clear()
    branding.ASSETS.mkdir(exist_ok=True)
    (branding.ASSETS / "logo.png").write_bytes(b"corrupted-not-an-image")
    doc = build(GOLDEN)                     # must not raise
    check("build survives an unreadable logo (skips it)",
          docx_media(doc) == [])
    branding.clear()

    print("\n=== API endpoints ===")
    from fastapi.testclient import TestClient
    import app as app_module
    client = TestClient(app_module.app)
    branding.clear()
    r = client.get("/api/branding")
    check("GET branding: empty state",
          r.json() == {"logo": False, "profile": None})
    r = client.post("/api/branding",
                    files={"logo": ("l.png", make_logo_png(), "image/png"),
                           "reference": ("r.pdf", ref_pdf,
                                          "application/pdf")})
    check("POST branding saves both", r.status_code == 200
          and r.json().get("logo") == "saved"
          and r.json()["profile"]["body_pt"] == 11)
    r = client.get("/api/branding")
    check("GET branding reflects saved state",
          r.json()["logo"] is True and r.json()["profile"] is not None)
    r = client.post("/api/branding",
                    files={"logo": ("l.gif", b"GIF89a", "image/gif")})
    check("bad logo type -> 400", r.status_code == 400)
    r = client.post("/api/branding", data={})
    check("empty POST -> 400", r.status_code == 400)
finally:
    branding.clear()

print(f"\n{passed} passed, {failed} failed")
sys.exit(1 if failed else 0)
