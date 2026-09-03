#!/usr/bin/env python3
"""Offline tests for Phase 4 (textbook mode): ingestion/chunking/matching,
and generate_unit with a source file — grounding text reaching the prompts,
per-subsection chunk mapping, fallbacks, provenance. Model stubbed.

Run: python tests/test_textbook_mode.py
"""
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "backend"))

import ingest
import schemas
from ai_engine import OllamaEngine
from unit_generator import generate_unit

passed = failed = 0


def check(name, condition):
    global passed, failed
    if condition:
        passed += 1
        print(f"  OK   {name}")
    else:
        failed += 1
        print(f"  FAIL {name}")


CHAPTER = """Chapter 2: Cryptographic Foundations
This chapter introduces the building blocks of cryptography.

2.1 The Caesar Cipher
The Caesar cipher shifts each letter by a fixed amount. With a shift of 3,
SECURITY becomes VHFXULWB. It offers 25 usable keys.
1. It is trivially breakable by brute force.

2.2 Transposition Techniques
A transposition cipher rearranges letters without replacing them. The rail
fence cipher writes plaintext diagonally across rows.

2.3 Frequency Analysis
Every natural language has a characteristic letter frequency. In English, E
is the most common letter, which betrays substitution ciphers.
"""


print("=== ingest: chunking ===")
chunks = ingest.chunk_by_headings(CHAPTER)
check("chapter heading + 3 numbered sections detected (4 chunks)",
      len(chunks) == 4 and chunks[0]["heading"].startswith("Chapter 2"))
check("numbered list item NOT mistaken for a heading ('1. It is...' ends "
      "with a full stop)",
      all("trivially breakable" not in c["heading"] for c in chunks)
      and "trivially breakable" in chunks[1]["text"])
check("chunk text captured under its heading",
      "VHFXULWB" in chunks[1]["text"]
      and chunks[1]["heading"] == "2.1 The Caesar Cipher")

print("\n=== ingest: extract_text round trips ===")
tmp = Path(tempfile.mkdtemp())
txt_path = tmp / "src.txt"
txt_path.write_text(CHAPTER, encoding="utf-8")
check(".txt extraction", "Caesar cipher shifts" in
      ingest.extract_text(txt_path))

from docx import Document as _Doc
docx_path = tmp / "src.docx"
d = _Doc()
for line in CHAPTER.splitlines():
    d.add_paragraph(line)
d.save(str(docx_path))
docx_text = ingest.extract_text(docx_path)
check(".docx extraction preserves headings for the chunker",
      len(ingest.chunk_by_headings(docx_text)) == 4)

import fitz
pdf_path = tmp / "src.pdf"
pdf = fitz.open()
page = pdf.new_page()
page.insert_text((72, 72), "2.1 The Caesar Cipher")
page.insert_text((72, 100), "A shifted alphabet.")
pdf.save(str(pdf_path))
pdf.close()
check(".pdf extraction", "Caesar Cipher" in ingest.extract_text(pdf_path))

print("\n=== ingest: matching + condense ===")
m, um = ingest.match_chunks(["2.1 The Caesar Cipher"], chunks)
check("exact heading match", len(m) == 1 and "VHFXULWB" in m[0]["text"])
m, um = ingest.match_chunks(["The Caesar cipher"], chunks)
check("numbering/case-insensitive match", len(m) == 1 and um == [])
m, um = ingest.match_chunks(["Quantum Entanglement"], chunks)
check("unmatched heading reported", m == [] and um == ["Quantum Entanglement"])
digest = ingest.condense(chunks, per_chunk=50, cap=10_000)
check("condense keeps every heading, truncates bodies",
      digest.count("## ") == 4 and "VHFXULWB"[:3] not in digest.split("## ")[2][60:])

print("\n=== generate_unit: textbook mode (stubbed) ===")

CANNED_TB = {
    id(schemas.OUTLINE_TEXTBOOK): {
        "example_style": "problem",
        "sections": [
            {"title": f"Sec {i}", "intro": "Opener.",
             "subsections": [
                 {"title": "Caesar", "source_headings":
                     ["2.1 The Caesar Cipher"]},
                 {"title": "Ghost topic", "source_headings":
                     ["Quantum Entanglement"]}]}
            for i in (1, 2, 3)],
    },
}
CANNED_TB.update({
    id(schemas.INTRODUCTION): {"paragraphs": ["P1.", "P2.", "P3."]},
    id(schemas.LEARNING_OBJECTIVES): {"learning_objectives": [
        {"verb": "Define", "rest": "r1"}, {"verb": "Apply", "rest": "r2"},
        {"verb": "Analyse", "rest": "r3"}, {"verb": "Solve", "rest": "r4"}]},
    id(schemas.PROSE): {"paragraphs": ["Teaching text.", "More text."]},
    id(schemas.TABLE): {"caption_title": "T", "columns": ["A", "B"],
                        "rows": [["1", "2"], ["3", "4"], ["5", "6"]]},
    id(schemas.PROBLEM): {"statement": "S", "solution": "Sol ✓"},
    id(schemas.SECTION_EXTRAS): {"think_and_apply": "Try it.",
                                  "figure_caption": "A figure"},
    id(schemas.SUMMARY): {"summary": [f"S{i}" for i in range(6)]},
    id(schemas.GLOSSARY): {"glossary": [
        {"term": f"T{i:02d}", "definition": "D."} for i in range(12)]},
    id(schemas.CASE_STUDY): {"title": "Case", "background": ["B1.", "B2.",
                                                            "B3."],
                              "questions": ["Q1?", "Q2?", "Q3?"]},
    id(schemas.MCQS): {"mcq": [{"q": f"Q{i}?",
                                "options": ["1", "2", "3", "4"],
                                "answer": "a"} for i in range(8)]},
    id(schemas.FILL_BLANKS): {"fill_blanks": [
        {"q": f"The ______ {i}.", "answer": "x"} for i in range(5)]},
    id(schemas.TERMINAL_SHORT): {"short": [
        {"q": f"SQ{i}?", "answer": "A."} for i in range(5)]},
    id(schemas.TERMINAL_LONG_QS): {"long_questions": [
        f"LQ{i}?" for i in range(5)]},
    id(schemas.LONG_ANSWER): {"answer": "A model essay."},
    id(schemas.REFERENCES): {"references": [
        f"Ref {i}." for i in range(6)]},
})


class StubEngine(OllamaEngine):
    def __init__(self):
        super().__init__(model="stub", host="http://stub")
        self.prompts_by_schema = {}

    def ask(self, task_prompt, schema=None, *, system=None, temperature=0.3):
        self.prompts_by_schema.setdefault(id(schema), []).append(task_prompt)
        return CANNED_TB[id(schema)]


def quiet(_msg):
    pass


engine = StubEngine()
unit, report = generate_unit(
    {"programme": "MSc", "course_code": "T1", "course_name": "C",
     "unit_number": 2, "unit_title": "U"},
    source_path=txt_path, engine=engine, progress=quiet)

check("source_mode is 'textbook'", unit["meta"]["source_mode"] == "textbook")
check("outline used the TEXTBOOK schema and got the source headings list",
      id(schemas.OUTLINE_TEXTBOOK) in engine.prompts_by_schema
      and "- 2.1 The Caesar Cipher"
      in engine.prompts_by_schema[id(schemas.OUTLINE_TEXTBOOK)][0])
prose_prompts = engine.prompts_by_schema[id(schemas.PROSE)]
check("matched subsection's prose call carries ITS chunk + grounding rules",
      any("VHFXULWB" in p and "STRICT SOURCE RULES" in p
          for p in prose_prompts))
# the whole-source digest is the only prompt carrying EVERY section heading
check("unmatched subsection fell back to the condensed digest",
      any("STRICT SOURCE RULES" in p
          and "## 2.2 Transposition Techniques" in p
          and "## 2.3 Frequency Analysis" in p
          for p in prose_prompts))
check("back matter grounded on the condensed digest",
      "## 2.3 Frequency Analysis"
      in engine.prompts_by_schema[id(schemas.MCQS)][0])
check("source textbook placed first in references (SME-citation stub)",
      unit["references"][0].startswith("[Source textbook: src.txt"))
check("provenance in report (file/chars/chunks + unmatched headings)",
      report["source"]["chunks"] == 4
      and "Quantum Entanglement" in report["source"]["unmatched_headings"])
check("unmatched-headings warning surfaced",
      any("Quantum Entanglement" in w
          for w in report["validation"]["warnings"]))
check("validation has no errors", report["validation"]["errors"] == [])

print("\n=== generate_unit: source with no detectable headings ===")
flat_path = tmp / "flat.txt"
flat_path.write_text("Just prose with no numbered structure at all. " * 40,
                     encoding="utf-8")
CANNED_TB[id(schemas.OUTLINE)] = {
    "example_style": "problem",
    "sections": [{"title": f"S{i}", "intro": "o",
                  "subsections": [{"title": "a"}, {"title": "b"}]}
                 for i in (1, 2, 3)]}
engine = StubEngine()
unit, report = generate_unit(
    {"programme": "MSc", "course_code": "T1", "course_name": "C",
     "unit_number": 2, "unit_title": "U"},
    source_path=flat_path, engine=engine, progress=quiet)
check("falls back to the plain OUTLINE schema",
      id(schemas.OUTLINE) in engine.prompts_by_schema)
check("whole-document grounding still reaches content calls",
      any("Just prose with no numbered structure"
          in p for p in engine.prompts_by_schema[id(schemas.PROSE)]))
check("split-failure warning surfaced",
      any("did not split" in w for w in report["validation"]["warnings"]))

print(f"\n{passed} passed, {failed} failed")
sys.exit(1 if failed else 0)
