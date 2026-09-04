#!/usr/bin/env python3
"""Offline tests for the Phase 3 pipeline — generate_unit with the model
stubbed (dispatch by schema identity), through validation, UK-style pass,
and a real render through the Phase 1 docx builder.

Run: python tests/test_unit_generator.py
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "backend"))

import schemas
from ai_engine import OllamaEngine
from docx_builder import build
from uk_style import apply_uk_spelling, apply_to_unit
from unit_generator import GenerationError, generate_unit
from validate_unit import validate_unit

passed = failed = 0


def check(name, condition):
    global passed, failed
    if condition:
        passed += 1
        print(f"  OK   {name}")
    else:
        failed += 1
        print(f"  FAIL {name}")


META = {"programme": "MSc SEMESTER (I)", "course_code": "TEST101",
        "course_name": "Test Course", "unit_number": 2,
        "unit_title": "Stubbed Unit"}

CANNED = {
    id(schemas.OUTLINE): {
        "example_style": "code",
        "sections": [
            {"title": f"Major Topic {i}", "intro": f"Opener {i}.",
             "subsections": [{"title": f"Sub {i}.1"},
                              {"title": f"Sub {i}.2"},
                              {"title": f"Sub {i}.3"}]}
            for i in (1, 2, 3)
        ],
    },
    id(schemas.INTRODUCTION): {"paragraphs": [
        "Welcome to the unit; we are analyzing collections of things.",
        "Second paragraph.", "Third paragraph."]},
    id(schemas.LEARNING_OBJECTIVES): {"learning_objectives": [
        {"verb": "Define", "rest": "the key terms of the subject clearly"},
        {"verb": "Apply", "rest": "the core operations to given datasets"},
        {"verb": "Analyse", "rest": "relationships between multiple entities"},
        {"verb": "Illustrate", "rest": "concepts using standard diagrams"}]},
    id(schemas.PROSE): {"paragraphs": ["A teaching paragraph about color "
                                       "coding.", "Another paragraph."]},
    id(schemas.TABLE): {"caption_title": "Concepts and Applications",
                        "columns": ["Concept", "Meaning"],
                        "rows": [["A", "1"], ["B", "2"], ["C", "3"]]},
    id(schemas.CODE): {"language": "python",
                       "code": "colors = {'analyze'}\nprint(colors)",
                       "explanation": "Shows a set literal."},
    id(schemas.DID_YOU_KNOW): {"text": "A genuinely interesting aside."},
    id(schemas.SECTION_EXTRAS): {"think_and_apply": "Apply this to a shop.",
                                  "figure_caption": "How the parts connect"},
    id(schemas.SUMMARY): {"summary": [f"Point {i}" for i in range(1, 8)]},
    id(schemas.GLOSSARY): {"glossary": [
        {"term": f"Term {i:02d}", "definition": f"Definition {i}."}
        for i in range(1, 15)]},
    id(schemas.CASE_STUDY): {"title": "Case: Riverside Analytics",
                              "background": ["Para one.", "Para two.",
                                             "Para three."],
                              "questions": ["Q1?", "Q2?", "Q3?"]},
    id(schemas.MCQS): {"mcq": [
        {"q": f"Question {i}?", "options": ["w", "x", "y", "z"],
         "answer": "b"} for i in range(1, 9)]},
    id(schemas.FILL_BLANKS): {"fill_blanks": [
        {"q": f"The ______ number {i}.", "answer": f"answer{i}"}
        for i in range(1, 6)]},
    id(schemas.TERMINAL_SHORT): {"short": [
        {"q": f"Short Q{i}?", "answer": f"Short A{i}."} for i in range(1, 6)]},
    id(schemas.TERMINAL_LONG_QS): {"long_questions": [
        f"Long Q{i}?" for i in range(1, 6)]},
    id(schemas.LONG_ANSWER): {"answer": "A thorough model essay answer."},
    id(schemas.REFERENCES): {"references": [
        f"Author {i}. (2020). Book {i}. Publisher." for i in range(1, 7)]},
}


class StubEngine(OllamaEngine):
    def __init__(self, fail_schemas=()):
        super().__init__(model="stub", host="http://stub")
        self.fail_schemas = {id(s) for s in fail_schemas}
        self.prompts_by_schema = {}

    def ask(self, task_prompt, schema=None, *, system=None, temperature=0.3):
        self.prompts_by_schema.setdefault(id(schema), []).append(task_prompt)
        if id(schema) in self.fail_schemas:
            from ai_engine import AIEngineError
            raise AIEngineError("stubbed failure")
        return CANNED[id(schema)]


def quiet(_msg):
    pass


print("=== clean_title (TOC noise stripped from outline titles) ===")
from unit_generator import clean_title
check("TOC parenthetical stripped",
      clean_title("The Caesar Cipher (2.1.1, 2.1.2 in the textbook TOC)")
      == "The Caesar Cipher")
check("'table of contents' variant stripped",
      clean_title("Stream Ciphers (see table of contents 2.2)")
      == "Stream Ciphers")
check("leading list numbering stripped",
      clean_title("2.1.1 The Caesar Cipher") == "The Caesar Cipher")
check("ordinary parentheticals kept",
      clean_title("Public-Key Cryptography (RSA)")
      == "Public-Key Cryptography (RSA)")

from unit_generator import clean_mcq_options
cleaned = clean_mcq_options([{"q": "?", "options":
                              ["a) AES", "b. RSA", "C: DES", "Vigenère"],
                              "answer": "d"}])
check("embedded option-letter prefixes stripped ('a) AES' -> 'AES')",
      cleaned[0]["options"] == ["AES", "RSA", "DES", "Vigenère"])

print("\n=== uk_style ===")
fixed, n = apply_uk_spelling("We analyze color and behavior while modeling.")
check("US spellings fixed",
      fixed == "We analyse colour and behaviour while modelling." and n == 4)
fixed, n = apply_uk_spelling("Analyzing the Center's organization")
check("case preserved on capitalised words",
      fixed.startswith("Analysing the Centre's organis"))
fixed, n = apply_uk_spelling("The programme licenses practice sessions.")
check("ambiguous words untouched", n == 0)

print("\n=== generate_unit: ai mode, all calls succeed ===")
engine = StubEngine()
unit, report = generate_unit(META, syllabus_topics=["topic a", "topic b"],
                             engine=engine, progress=quiet)
check("source_mode is 'ai' without a TOC",
      unit["meta"]["source_mode"] == "ai" and report["source_mode"] == "ai")
check("no failures reported", report["failures"] == [])
errors, warnings = report["validation"]["errors"], report["validation"]["warnings"]
check("validation passes with no errors", errors == [])
check("section numbering assigned in code (2.1..2.3, 2.1.1..)",
      [s["number"] for s in unit["sections"]] == ["2.1", "2.2", "2.3"]
      and unit["sections"][0]["subsections"][0]["number"] == "2.1.1")
first_sub = unit["sections"][0]["subsections"][0]
check("first subsection: prose + numbered table + numbered figure",
      any(b["type"] == "table" and b["caption"].startswith("Table 2.1.1: ")
          for b in first_sub["blocks"])
      and any(b["type"] == "figure" and b["caption"].startswith("Figure 1: ")
              for b in first_sub["blocks"]))
second_sub = unit["sections"][0]["subsections"][1]
check("second subsection: code example + explanation prose "
      "(example_style=code)",
      any(b["type"] == "code" for b in second_sub["blocks"])
      and not any(b["type"] == "problem" for b in second_sub["blocks"]))
last_sub = unit["sections"][0]["subsections"][2]
check("last subsection: did_you_know + think_and_apply",
      any(b["type"] == "did_you_know" for b in last_sub["blocks"])
      and last_sub["blocks"][-1]["type"] == "think_and_apply")
check("figures numbered globally across sections (Figure 3 in section 3)",
      any(b["type"] == "figure" and b["caption"].startswith("Figure 3: ")
          for b in unit["sections"][2]["subsections"][0]["blocks"]))
check("UK pass ran and fixed the stubbed US spellings",
      report["uk_spelling_fixes"] >= 3
      and "analysing" in unit["introduction"][0])
code_blocks = [b for s in unit["sections"] for ss in s["subsections"]
               for b in ss["blocks"] if b["type"] == "code"]
check("code blocks NOT UK-respelled ('analyze' survives in code)",
      all("analyze" in b["text"] for b in code_blocks))
check("references warning always surfaced for SME verification",
      any("SME" in w for w in warnings))
check("terminal long assembled from question + per-answer calls (5 items)",
      len(unit["terminal"]["long"]) == 5
      and unit["terminal"]["long"][0]["q"] == "Long Q1?"
      and "essay" in unit["terminal"]["long"][0]["answer"])

print("\n=== academic level steering ===")
engine = StubEngine()
unit, report = generate_unit({**META, "level": "undergraduate"},
                             engine=engine, progress=quiet)
ug_outline = engine.prompts_by_schema[id(schemas.OUTLINE)][0]
ug_objectives = engine.prompts_by_schema[id(schemas.LEARNING_OBJECTIVES)][0]
ug_prose = engine.prompts_by_schema[id(schemas.PROSE)][0]
check("undergraduate guidance rides on EVERY call's context",
      all("UNDERGRADUATE" in p and "first principles" in p
          for p in (ug_outline, ug_prose)))
check("undergraduate objectives steered to foundational Bloom's verbs",
      "foundational verbs" in ug_objectives)
check("level recorded in unit meta and report",
      unit["meta"]["level"] == "undergraduate"
      and report["level"] == "undergraduate")

engine = StubEngine()
unit, report = generate_unit(META, engine=engine, progress=quiet)
pg_prose = engine.prompts_by_schema[id(schemas.PROSE)][0]
pg_objectives = engine.prompts_by_schema[id(schemas.LEARNING_OBJECTIVES)][0]
check("default level is postgraduate (preserves prior behaviour)",
      unit["meta"]["level"] == "postgraduate"
      and "POSTGRADUATE" in pg_prose and "rigour" in pg_prose)
check("postgraduate objectives steered to higher-order verbs",
      "higher-order verbs" in pg_objectives)

print("\n=== generate_unit: toc+ai mode ===")
engine = StubEngine()
unit, report = generate_unit(META, syllabus_topics=["topic a"],
                             toc_text="1. Alpha\n2. Beta\n3. Gamma",
                             engine=engine, progress=quiet)
check("source_mode is 'toc+ai' when a TOC is supplied",
      unit["meta"]["source_mode"] == "toc+ai")
outline_prompt = engine.prompts_by_schema[id(schemas.OUTLINE)][0]
check("outline prompt carries the TOC verbatim",
      "1. Alpha" in outline_prompt and "authoritative outline skeleton"
      in outline_prompt)
check("outline prompt also carries syllabus topics (both used together)",
      "topic a" in outline_prompt)
check("review note flags TOC-guided AI generation for SME review",
      "structure follows the supplied TOC" in report["review_note"])

print("\n=== failure containment ===")
engine = StubEngine(fail_schemas=[schemas.CODE, schemas.GLOSSARY])
unit, report = generate_unit(META, engine=engine, progress=quiet)
check("failed calls recorded, generation continues",
      len(report["failures"]) == 4)   # code x3 sections + glossary
check("unit completes without the failed pieces",
      unit["glossary"] == [] and not any(
          b["type"] == "code" for s in unit["sections"]
          for ss in s["subsections"] for b in ss["blocks"]))
check("missing glossary downgraded to warning, not error",
      report["validation"]["errors"] == []
      and any("glossary" in w for w in report["validation"]["warnings"]))

engine = StubEngine(fail_schemas=[schemas.OUTLINE])
try:
    generate_unit(META, engine=engine, progress=quiet)
    check("outline failure raises GenerationError", False)
except GenerationError:
    check("outline failure raises GenerationError", True)

print("\n=== validator gate ===")
engine = StubEngine(fail_schemas=[schemas.MCQS, schemas.TERMINAL_SHORT])
unit, report = generate_unit(META, engine=engine, progress=quiet)
check("zero MCQs / zero short questions are ERRORS (unit must not render)",
      any("MCQ" in e for e in report["validation"]["errors"])
      and any("short" in e for e in report["validation"]["errors"]))

print("\n=== end-to-end render through the Phase 1 builder ===")
engine = StubEngine()
unit, report = generate_unit(META, engine=engine, progress=quiet)
out = Path(__file__).parent.parent / "output" / "stub_unit.docx"
out.parent.mkdir(exist_ok=True)
build(unit).save(str(out))
check("stub-generated unit renders to a real docx",
      out.stat().st_size > 20_000)

print(f"\n{passed} passed, {failed} failed")
sys.exit(1 if failed else 0)
