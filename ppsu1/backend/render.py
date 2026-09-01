#!/usr/bin/env python3
"""
render.py — turn a finished .pptx into one PNG per slide so the website can
show a visual preview before the user downloads.

Two rendering engines are tried, in order:

  1. LibreOffice (headless): pptx -> PDF -> one PNG per page (via PyMuPDF).
     This is the cross-platform path and the one used on the server (Railway).
  2. Microsoft PowerPoint (COM automation): Windows-only fallback so previews
     work on a local dev machine that has PowerPoint but not LibreOffice.

If neither engine is available the functions return an empty list; callers
should treat that as "preview unavailable" and still allow the download.
"""
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import traceback
from pathlib import Path

# PowerPoint automation is single-instance and stateful: two conversions
# running at once (two decks of one batch, or a redesign while a batch is
# still going) make it throw "call was rejected by callee" and return no
# images at all. Serialise every COM render behind this lock.
_PPT_LOCK = threading.Lock()
_PPT_ATTEMPTS = 3

# PyMuPDF (imported as fitz) converts the LibreOffice PDF into page images.
try:
    import fitz  # noqa: F401
    _HAVE_FITZ = True
except Exception:
    _HAVE_FITZ = False


def _find_soffice():
    """Locate the LibreOffice binary across platforms, or None if absent."""
    env = os.environ.get("SOFFICE_BIN")
    if env and Path(env).exists():
        return env
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/usr/bin/soffice",
        "/usr/bin/libreoffice",
        "/opt/libreoffice/program/soffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]
    for c in candidates:
        if Path(c).exists():
            return c
    return None


def _sorted_pngs(out_dir: Path):
    """PNG files in a directory, de-duplicated (Windows' filesystem is
    case-insensitive, so a single Slide1.PNG must not be counted twice) and
    ordered by the trailing number in their name."""
    seen = {}
    for p in out_dir.iterdir():
        if p.is_file() and p.suffix.lower() == ".png":
            seen[p.name.lower()] = p

    def num(p):
        m = re.search(r"(\d+)", p.stem)
        return int(m.group(1)) if m else 0
    return sorted(seen.values(), key=num)


def _render_with_libreoffice(pptx_path: Path, out_dir: Path):
    soffice = _find_soffice()
    if not soffice or not _HAVE_FITZ:
        return []

    # A dedicated user profile avoids "another LibreOffice instance is
    # running" profile-lock errors when several conversions run back to back.
    profile = Path(tempfile.mkdtemp(prefix="lo_profile_"))
    try:
        cmd = [
            soffice, "--headless", "--norestore", "--nologo", "--nofirststartwizard",
            "-env:UserInstallation=%s" % profile.as_uri(),
            "--convert-to", "pdf", "--outdir", str(out_dir), str(pptx_path),
        ]
        subprocess.run(cmd, check=True, timeout=180,
                       stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    except Exception:
        shutil.rmtree(profile, ignore_errors=True)
        return []
    finally:
        shutil.rmtree(profile, ignore_errors=True)

    pdf_path = out_dir / (pptx_path.stem + ".pdf")
    if not pdf_path.exists():
        return []

    images = []
    zoom = 130 / 72  # ~130 DPI: crisp thumbnails without huge files
    try:
        doc = fitz.open(str(pdf_path))
        matrix = fitz.Matrix(zoom, zoom)
        for i, page in enumerate(doc):
            pix = page.get_pixmap(matrix=matrix)
            img_path = out_dir / f"slide_{i}.png"
            pix.save(str(img_path))
            images.append(img_path)
        doc.close()
    except Exception:
        return []
    finally:
        try:
            pdf_path.unlink()
        except OSError:
            pass
    return images


def _powerpoint_export(pptx_path: Path, target, slide_index=None):
    """One PowerPoint COM export attempt. `target` is a directory (whole deck)
    or a file path (single slide when `slide_index` is given). Raises on
    failure so the caller can retry."""
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    app = pres = None
    try:
        # DispatchEx starts a PRIVATE PowerPoint process: quitting it at the
        # end can never close a presentation the user has open themselves.
        app = win32com.client.DispatchEx("PowerPoint.Application")
        pres = app.Presentations.Open(
            str(pptx_path.resolve()), ReadOnly=1, Untitled=0, WithWindow=0)
        if slide_index is None:
            pres.Export(str(Path(target).resolve()), "PNG")
        else:
            if not (0 <= slide_index < pres.Slides.Count):
                return False
            pres.Slides(slide_index + 1).Export(
                str(Path(target).resolve()), "PNG")
        return True
    finally:
        try:
            if pres is not None:
                pres.Close()
        except Exception:
            pass
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


def _powerpoint_export_retry(pptx_path: Path, target, slide_index=None):
    """Run one export, retrying transient COM failures. PowerPoint routinely
    rejects a call while it is still shutting down from the previous deck, so
    a lone failure means nothing — only give up after several tries, and say
    why on the server log instead of silently producing no previews."""
    if sys.platform != "win32":
        return False
    try:
        import pythoncom  # noqa: F401
        import win32com.client  # noqa: F401
    except Exception:
        return False

    with _PPT_LOCK:
        for attempt in range(1, _PPT_ATTEMPTS + 1):
            try:
                return _powerpoint_export(pptx_path, target, slide_index)
            except Exception as e:
                print(f"[render] PowerPoint export failed for "
                      f"{pptx_path.name} (attempt {attempt}/{_PPT_ATTEMPTS}): "
                      f"{e.__class__.__name__}: {e}", file=sys.stderr)
                if attempt == _PPT_ATTEMPTS:
                    traceback.print_exc()
                    return False
                time.sleep(2 * attempt)     # let PowerPoint settle, then retry
    return False


def _render_with_powerpoint(pptx_path: Path, out_dir: Path):
    """Windows-only fallback using PowerPoint COM automation."""
    raw = []
    for attempt in range(2):
        if not _powerpoint_export_retry(pptx_path, out_dir):
            return []
        # an export can report success yet write nothing (PowerPoint gave up
        # part-way); treat an empty directory as a failure worth one retry
        raw = _sorted_pngs(out_dir)
        if raw:
            break
        print(f"[render] {pptx_path.name}: export wrote no images, retrying",
              file=sys.stderr)
        time.sleep(2)
    if not raw:
        return []

    # normalise PowerPoint's names (Slide1.PNG ...) to slide_0.png ... in order
    images = []
    for i, p in enumerate(raw):
        target = out_dir / f"slide_{i}.png"
        if p != target:
            try:
                p.replace(target)
            except OSError:
                target = p
        images.append(target)
    return images


def render_deck_to_images(pptx_path, out_dir):
    """Render every slide of `pptx_path` to a PNG inside `out_dir`.

    Returns the list of image Paths in slide order (empty if no engine could
    render — the caller should then fall back to a download-only experience).
    """
    pptx_path = Path(pptx_path)
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    # clear any stale images from a previous (re)design of this deck
    for old in out_dir.glob("*.png"):
        try:
            old.unlink()
        except OSError:
            pass

    images = _render_with_libreoffice(pptx_path, out_dir)
    if not images:
        images = _render_with_powerpoint(pptx_path, out_dir)
    return images


def render_one_slide(pptx_path, index, out_path):
    """Render a SINGLE slide (0-based `index`) of `pptx_path` to `out_path`.

    Used after a per-slide swap so only the changed slide is re-rendered instead
    of the whole deck. PowerPoint (COM) can export one slide directly and is
    fast; LibreOffice can only convert the whole deck, so there we fall back to
    rendering everything and copying out the one page. Returns True on success.
    """
    pptx_path, out_path = Path(pptx_path), Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    # 1) PowerPoint COM — export just the one slide (fast)
    if _powerpoint_export_retry(pptx_path, out_path, slide_index=index):
        if out_path.exists():
            return True

    # 2) LibreOffice fallback — render the whole deck, copy out the one page
    tmp = out_path.parent / "_one"
    imgs = _render_with_libreoffice(pptx_path, tmp)
    if 0 <= index < len(imgs):
        try:
            out_path.write_bytes(Path(imgs[index]).read_bytes())
            return True
        except OSError:
            return False
    return False


def renderer_available():
    """True if any rendering engine is present on this machine."""
    if _find_soffice() and _HAVE_FITZ:
        return True
    if sys.platform == "win32":
        try:
            import win32com  # noqa: F401
            return True
        except Exception:
            return False
    return False
