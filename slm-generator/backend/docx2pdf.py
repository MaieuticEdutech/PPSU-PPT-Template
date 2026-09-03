"""docx2pdf.py — Phase 6 PDF export, same engine-fallback pattern as
../../ppsu1/backend/render.py: Word COM automation on Windows (this office
machine has Office installed), headless LibreOffice anywhere it exists,
otherwise "unavailable" — callers offer the .docx only, never fail a job
over a missing PDF.

Word COM is single-instance and stateful, so conversions are serialised
behind one lock (the ppsu1 PowerPoint lesson applied to Word).
"""
import shutil
import subprocess
import sys
import tempfile
import threading
from pathlib import Path

_WORD_LOCK = threading.Lock()
_WD_FORMAT_PDF = 17


def _convert_with_word(docx_path: Path, pdf_path: Path) -> bool:
    if sys.platform != "win32":
        return False
    try:
        import pythoncom
        import win32com.client
    except Exception:
        return False
    with _WORD_LOCK:
        pythoncom.CoInitialize()
        word = doc = None
        try:
            # DispatchEx: a PRIVATE Word process, so quitting it can never
            # close a document the user has open themselves
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(str(docx_path.resolve()),
                                      ReadOnly=True)
            doc.SaveAs2(str(pdf_path.resolve()), FileFormat=_WD_FORMAT_PDF)
            return True
        except Exception:
            return False
        finally:
            try:
                if doc is not None:
                    doc.Close(False)
            except Exception:
                pass
            try:
                if word is not None:
                    word.Quit()
            except Exception:
                pass
            pythoncom.CoUninitialize()


def _find_soffice():
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    for c in (r"C:\Program Files\LibreOffice\program\soffice.exe",
              "/usr/bin/soffice", "/usr/bin/libreoffice"):
        if Path(c).exists():
            return c
    return None


def _convert_with_soffice(docx_path: Path, pdf_path: Path) -> bool:
    soffice = _find_soffice()
    if not soffice:
        return False
    profile = Path(tempfile.mkdtemp(prefix="lo_profile_"))
    try:
        subprocess.run(
            [soffice, "--headless", "--norestore", "--nologo",
             "-env:UserInstallation=%s" % profile.as_uri(),
             "--convert-to", "pdf", "--outdir",
             str(pdf_path.parent), str(docx_path)],
            check=True, timeout=180,
            stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    except Exception:
        return False
    finally:
        shutil.rmtree(profile, ignore_errors=True)
    produced = pdf_path.parent / (docx_path.stem + ".pdf")
    if produced != pdf_path and produced.exists():
        produced.replace(pdf_path)
    return pdf_path.exists()


def convert(docx_path, pdf_path) -> bool:
    """docx -> pdf. True on success; False means 'no engine available or
    conversion failed' — offer the .docx alone in that case."""
    docx_path, pdf_path = Path(docx_path), Path(pdf_path)
    if _convert_with_word(docx_path, pdf_path) and pdf_path.exists():
        return True
    return _convert_with_soffice(docx_path, pdf_path)


def available() -> bool:
    if _find_soffice():
        return True
    if sys.platform == "win32":
        try:
            import win32com  # noqa: F401
            return True
        except Exception:
            return False
    return False
