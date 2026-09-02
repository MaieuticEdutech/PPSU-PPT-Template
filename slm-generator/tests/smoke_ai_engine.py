#!/usr/bin/env python3
"""LIVE smoke test for ai_engine.py — needs a running Ollama server with the
configured model pulled (unlike test_ai_engine.py, which stubs the model).

Asks for real content in the shapes Phase 1's builder renders, validates it
against real schemas, and prints what came back so a human can judge quality,
not just structure.

Run: python tests/smoke_ai_engine.py
"""
import json
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "backend"))

from ai_engine import AIEngineError, get_engine

engine = get_engine()
server_up, model_present = engine.available()
if not server_up:
    print("BLOCKED: Ollama server is not running (start it, or install from "
          "ollama.com). Not a code failure.")
    sys.exit(2)
if not model_present:
    print(f"BLOCKED: model {engine.model!r} is not pulled. "
          f"Run: ollama pull {engine.model}")
    sys.exit(2)

print(f"Ollama up, model {engine.model!r} present. num_ctx={engine.num_ctx}")

MCQ_SCHEMA = {
    "type": "object",
    "properties": {
        "mcq": {
            "type": "array", "minItems": 3, "maxItems": 3,
            "items": {
                "type": "object",
                "properties": {
                    "q": {"type": "string", "minLength": 10},
                    "options": {"type": "array",
                                "items": {"type": "string"},
                                "minItems": 4, "maxItems": 4},
                    "answer": {"type": "string", "enum": ["a", "b", "c", "d"]},
                },
                "required": ["q", "options", "answer"],
            },
        },
    },
    "required": ["mcq"],
}

OBJECTIVES_SCHEMA = {
    "type": "object",
    "properties": {
        "learning_objectives": {
            "type": "array", "minItems": 4, "maxItems": 5,
            "items": {
                "type": "object",
                "properties": {
                    "verb": {"type": "string"},
                    "rest": {"type": "string", "minLength": 5},
                },
                "required": ["verb", "rest"],
            },
        },
    },
    "required": ["learning_objectives"],
}

SYSTEM = ("You write UK-English academic content for P P Savani University "
          "self-learning materials. Output only JSON.")

t0 = time.time()
print("\n--> 3 MCQs on set theory (temperature 0.3)")
data = engine.ask(
    "Write exactly 3 multiple-choice questions testing basic set theory "
    "(sets, subsets, union, intersection) for an MSc Data Science unit. "
    "Each question has exactly 4 options and one correct answer given as "
    "the letter a, b, c or d.",
    schema=MCQ_SCHEMA, system=SYSTEM)
print(f"    valid in {time.time()-t0:.1f}s:")
for i, m in enumerate(data["mcq"], 1):
    print(f"    {i}. {m['q']}  [answer: {m['answer']}]")

t0 = time.time()
print("\n--> 4-5 learning objectives with Bloom's verbs (temperature 0.3)")
data = engine.ask(
    "Write 4 to 5 learning objectives for a unit titled 'Foundations of "
    "Discrete Mathematics' covering sets, subsets, power sets, set "
    "operations, Venn diagrams and relations. Each objective is split into "
    "'verb' (a single Bloom's-taxonomy verb such as Define, Explain, Apply, "
    "Analyse — never Understand, Know or Learn) and 'rest' (8-12 words, one "
    "measurable outcome).",
    schema=OBJECTIVES_SCHEMA, system=SYSTEM)
print(f"    valid in {time.time()-t0:.1f}s:")
banned = {"understand", "know", "learn"}
for o in data["learning_objectives"]:
    flag = "  [!] banned verb" if o["verb"].lower() in banned else ""
    print(f"    - {o['verb']} {o['rest']}{flag}")

t0 = time.time()
print("\n--> extraction-style call (temperature 0.1)")
data = engine.ask(
    "From this passage, extract the glossary terms it defines, as "
    '{"glossary": [{"term": ..., "definition": ...}]}: '
    "\"A set is a well-defined collection of distinct objects. The union "
    "of A and B is the set of elements in A, in B, or in both. A Venn "
    "diagram represents sets as overlapping circles.\"",
    schema={
        "type": "object",
        "properties": {"glossary": {
            "type": "array", "minItems": 2,
            "items": {"type": "object",
                      "properties": {"term": {"type": "string"},
                                     "definition": {"type": "string"}},
                      "required": ["term", "definition"]}}},
        "required": ["glossary"],
    },
    system=SYSTEM, temperature=0.1)
print(f"    valid in {time.time()-t0:.1f}s:")
for g in data["glossary"]:
    print(f"    - {g['term']}: {g['definition'][:70]}")

print("\nsmoke test PASSED — engine produces valid, schema-conforming "
      "content end to end")
