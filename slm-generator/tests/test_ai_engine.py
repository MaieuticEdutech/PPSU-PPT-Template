#!/usr/bin/env python3
"""Offline tests for ai_engine.py — fence-stripping and the parse/validate/
retry contract, with the model stubbed out. No Ollama server needed (the
LIVE end-to-end check is tests/smoke_ai_engine.py, which does need one).

Run: python tests/test_ai_engine.py
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "backend"))

from ai_engine import AIEngineError, OllamaEngine, strip_json_fences

passed = failed = 0


def check(name, condition):
    global passed, failed
    if condition:
        passed += 1
        print(f"  OK   {name}")
    else:
        failed += 1
        print(f"  FAIL {name}")


RAISE = object()   # sentinel: this scripted "reply" is a transport failure


class StubEngine(OllamaEngine):
    """Feeds a scripted sequence of raw model replies to ask()."""
    def __init__(self, replies):
        super().__init__(model="stub", host="http://stub")
        self.replies = list(replies)
        self.prompts = []

    def _chat(self, prompt, schema, system, temperature, num_predict=None):
        self.prompts.append(prompt)
        reply = self.replies.pop(0)
        if reply is RAISE:
            raise AIEngineError("stubbed timeout")
        return reply


MCQ_SCHEMA = {
    "type": "object",
    "properties": {
        "mcq": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "q": {"type": "string"},
                    "options": {"type": "array", "items": {"type": "string"},
                                "minItems": 4, "maxItems": 4},
                    "answer": {"type": "string", "enum": ["a", "b", "c", "d"]},
                },
                "required": ["q", "options", "answer"],
            },
        },
    },
    "required": ["mcq"],
}

GOOD = '{"mcq": [{"q": "2+2?", "options": ["1","2","3","4"], "answer": "d"}]}'


print("=== strip_json_fences ===")
check("plain JSON untouched", strip_json_fences(GOOD) == GOOD)
check("```json fences stripped",
      strip_json_fences(f"```json\n{GOOD}\n```") == GOOD)
check("bare ``` fences stripped",
      strip_json_fences(f"```\n{GOOD}\n```") == GOOD)
check("leading prose cut to outermost object",
      strip_json_fences(f"Here is your JSON:\n{GOOD}\nHope this helps!") == GOOD)
check("array payloads survive",
      strip_json_fences('the answer: [1, 2, 3] ok') == "[1, 2, 3]")

print("\n=== ask(): parse/validate/retry contract ===")

e = StubEngine([GOOD])
data = e.ask("write mcqs", schema=MCQ_SCHEMA)
check("valid first reply accepted, no retry",
      data["mcq"][0]["answer"] == "d" and len(e.prompts) == 1)

e = StubEngine([f"```json\n{GOOD}\n```"])
data = e.ask("write mcqs", schema=MCQ_SCHEMA)
check("fenced first reply repaired without burning the retry",
      data["mcq"][0]["q"] == "2+2?" and len(e.prompts) == 1)

e = StubEngine(["this is not json at all", GOOD])
data = e.ask("write mcqs", schema=MCQ_SCHEMA)
check("unparseable first reply -> ONE retry succeeds", len(e.prompts) == 2)
check("retry prompt carries the corrective reminder",
      "ONLY a single valid JSON" in e.prompts[1])

bad_shape = '{"mcq": [{"q": "2+2?", "options": ["1","2"], "answer": "z"}]}'
e = StubEngine([bad_shape, GOOD])
data = e.ask("write mcqs", schema=MCQ_SCHEMA)
check("schema-invalid first reply -> ONE retry succeeds",
      len(e.prompts) == 2 and data["mcq"][0]["answer"] == "d")
check("retry reminder names the validation problem",
      "did not match the required structure" in e.prompts[1])

e = StubEngine(["garbage one", "garbage two"])
try:
    e.ask("write mcqs", schema=MCQ_SCHEMA)
    check("second failure raises AIEngineError", False)
except AIEngineError:
    check("second failure raises AIEngineError", True)
check("exactly two attempts were made, never a third", len(e.prompts) == 2)

e = StubEngine(['{"anything": 1}'])
data = e.ask("free-form", schema=None)
check("schema=None still parses JSON without validating",
      data == {"anything": 1})

e = StubEngine([RAISE, GOOD])
data = e.ask("write mcqs", schema=MCQ_SCHEMA)
check("transport failure (timeout) -> ONE retry succeeds",
      len(e.prompts) == 2 and data["mcq"][0]["answer"] == "d")

e = StubEngine([RAISE, RAISE])
try:
    e.ask("write mcqs", schema=MCQ_SCHEMA)
    check("two transport failures raise AIEngineError", False)
except AIEngineError:
    check("two transport failures raise AIEngineError", True)

print(f"\n{passed} passed, {failed} failed")
sys.exit(1 if failed else 0)
