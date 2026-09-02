"""ingest.py — Phase 4 source ingestion: PDF/DOCX/TXT -> plain text ->
heading-based chunks that the generator can feed, one at a time, to
source-grounded AI calls (each call sees ONLY its subsection's chunk, per
CLAUDE.md's generation order).

The chunker keys on NUMBERED headings ("2.1 Classical Encryption", "2.1.1
The Caesar Cipher", "Chapter 2: Foundations") because that's how the
prescribed textbooks and the issued SLMs are structured. A source with no
detectable numbered headings still works — it becomes one chunk and the
report warns that splitting failed, rather than silently degrading.
"""
import re
from pathlib import Path

# a heading line: "Chapter N ..." or dotted-number + title; short, and not
# ending like a sentence (guards against numbered list items in prose)
_HEADING_RE = re.compile(
    r"^\s*(?:chapter\s+\d+[:.]?\s+\S.*|\d+(?:\.\d+)*\.?\s+[A-Za-z(\"'‘“].*)$",
    re.IGNORECASE)
_MAX_HEADING_LEN = 90


def extract_text(path) -> str:
    """Plain text from a .pdf / .docx / .txt source file."""
    path = Path(path)
    suffix = path.suffix.lower()
    if suffix == ".pdf":
        import fitz
        with fitz.open(str(path)) as doc:
            return "\n".join(page.get_text() for page in doc)
    if suffix == ".docx":
        from docx import Document
        doc = Document(str(path))
        return "\n".join(p.text for p in doc.paragraphs)
    if suffix in (".txt", ".md"):
        return path.read_text(encoding="utf-8", errors="replace")
    raise ValueError(f"unsupported source type: {suffix} "
                     f"(use .pdf, .docx or .txt)")


def _is_heading(line: str) -> bool:
    line = line.strip()
    return (0 < len(line) <= _MAX_HEADING_LEN
            and not line.endswith(".")
            and bool(_HEADING_RE.match(line)))


def chunk_by_headings(text: str):
    """[{"heading", "text"}] in document order. Content before the first
    heading lands in a "(preamble)" chunk."""
    chunks = []
    heading, buf = "(preamble)", []

    def flush():
        body = "\n".join(buf).strip()
        if body or heading != "(preamble)":
            chunks.append({"heading": heading, "text": body})

    for line in text.splitlines():
        if _is_heading(line):
            flush()
            heading, buf = line.strip(), []
        else:
            buf.append(line)
    flush()
    return chunks


def _norm(s: str) -> str:
    """Normalise a heading/title for matching: drop numbering, case,
    punctuation."""
    s = re.sub(r"^\s*(?:chapter\s+\d+[:.]?|\d+(?:\.\d+)*\.?)\s*", "",
               s.strip(), flags=re.IGNORECASE)
    return re.sub(r"[^a-z0-9 ]", "", s.lower()).strip()


def match_chunks(wanted_headings, chunks):
    """Map the outline's chosen source headings to actual chunks: exact
    normalised match first, then substring containment either way.
    Returns (matched_chunks_in_order, unmatched_headings)."""
    matched, unmatched = [], []
    for want in wanted_headings:
        w = _norm(want)
        hit = next((c for c in chunks if _norm(c["heading"]) == w), None)
        if hit is None and w:
            hit = next((c for c in chunks
                        if w in _norm(c["heading"])
                        or (_norm(c["heading"])
                            and _norm(c["heading"]) in w)), None)
        if hit is not None:
            if hit not in matched:
                matched.append(hit)
        else:
            unmatched.append(want)
    return matched, unmatched


def condense(chunks, per_chunk=700, cap=12000) -> str:
    """A whole-source digest for back-matter calls (summary/glossary/MCQs
    must span the unit, but a full chapter may not fit num_ctx): every
    chunk's heading + its opening `per_chunk` characters, capped overall."""
    parts = []
    total = 0
    for c in chunks:
        piece = f"## {c['heading']}\n{c['text'][:per_chunk]}"
        if total + len(piece) > cap:
            break
        parts.append(piece)
        total += len(piece)
    return "\n\n".join(parts)
