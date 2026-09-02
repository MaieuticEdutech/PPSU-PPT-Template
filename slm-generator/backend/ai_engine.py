#!/usr/bin/env python3
"""ai_engine.py — the ONE interface every AI call in this tool goes through
(CLAUDE.md: "No Ollama imports outside services/ai_engine.py"; routes and the
docx builder never talk to a model directly, so a future Claude/Grok backend
is a drop-in).

Usage:
    from ai_engine import get_engine
    engine = get_engine()
    data = engine.ask("Write 3 MCQs about set theory.", schema=MCQ_SCHEMA)

Contract (matches CLAUDE.md's "Ollama specifics"):
- Always requests structured output: the JSON schema is passed to Ollama's
  `format` field (supported since Ollama 0.5) so generation is constrained
  server-side; with no schema, `format: "json"` still forces JSON mode.
- Still defensive client-side: ``` fences are stripped, the parsed object is
  validated against the schema, and on a parse OR validation failure ONE
  retry is made with a corrective reminder appended. A second failure raises
  AIEngineError — the caller fails that block, not the whole unit.
- num_ctx is set explicitly (default 16384): Ollama's out-of-the-box 2k/4k
  context silently truncates textbook chunks — the #1 silent quality killer.
- temperature 0.3 for content generation, 0.1 for extraction/classification
  (pass temperature=0.1 explicitly for those calls).

Model default here is `qwen2.5:7b-instruct`, NOT the spec's 14b: this
machine's RTX 4060 has 8 GB VRAM, and the 14b Q4 weights (~9 GB) plus a
16k-context KV cache would force heavy CPU offload. The 7b Q4 (~4.7 GB)
fits fully in VRAM with the KV cache alongside. Override with OLLAMA_MODEL
when running on a bigger card.
"""
import json
import os
import re
import urllib.error
import urllib.request

import jsonschema

DEFAULT_MODEL = "qwen2.5:7b-instruct"
DEFAULT_HOST = "http://localhost:11434"
DEFAULT_NUM_CTX = 16384
DEFAULT_TEMPERATURE = 0.3
REQUEST_TIMEOUT = 600     # local 7b models are fast, but long generations
                          # (a full glossary) can take minutes on first load

_FENCE_RE = re.compile(r"```(?:json)?\s*(.*?)\s*```", re.S)

RETRY_REMINDER = (
    "\n\nIMPORTANT: your previous reply was not valid for this task "
    "({why}). Output ONLY a single valid JSON object matching the "
    "requested structure — no markdown fences, no commentary, no text "
    "before or after the JSON."
)


class AIEngineError(Exception):
    """The model could not produce a valid response after the one allowed
    retry, or the Ollama server is unreachable. Callers should fail the
    current block and carry on — never abort the whole unit for this."""


def strip_json_fences(text: str) -> str:
    """Return the JSON payload from a model reply that may have wrapped it
    in ``` fences or surrounded it with prose. Pure function — unit-tested
    without any model."""
    text = text.strip()
    m = _FENCE_RE.search(text)
    if m:
        return m.group(1).strip()
    # no fences: if there's leading/trailing chatter, cut to the outermost
    # JSON object/array
    for open_ch, close_ch in (("{", "}"), ("[", "]")):
        start = text.find(open_ch)
        end = text.rfind(close_ch)
        if start != -1 and end > start:
            return text[start:end + 1]
    return text


class OllamaEngine:
    def __init__(self, model=None, host=None, num_ctx=None):
        self.model = model or os.environ.get("OLLAMA_MODEL", DEFAULT_MODEL)
        self.host = (host or os.environ.get("OLLAMA_HOST", DEFAULT_HOST)).rstrip("/")
        self.num_ctx = int(num_ctx or os.environ.get("OLLAMA_NUM_CTX",
                                                     DEFAULT_NUM_CTX))

    # -- transport (separated so tests can stub it without a server) --------

    def _post(self, path: str, payload: dict) -> dict:
        req = urllib.request.Request(
            f"{self.host}{path}",
            data=json.dumps(payload).encode("utf-8"),
            headers={"Content-Type": "application/json"})
        try:
            with urllib.request.urlopen(req, timeout=REQUEST_TIMEOUT) as r:
                return json.loads(r.read())
        except urllib.error.URLError as e:
            raise AIEngineError(
                f"Ollama server unreachable at {self.host} ({e}). Is "
                f"`ollama serve` running and the model pulled?") from e

    def _get(self, path: str) -> dict:
        try:
            with urllib.request.urlopen(f"{self.host}{path}", timeout=10) as r:
                return json.loads(r.read())
        except urllib.error.URLError as e:
            raise AIEngineError(
                f"Ollama server unreachable at {self.host} ({e})") from e

    # -- public API ---------------------------------------------------------

    def available(self):
        """(server_up, model_present) — lets callers give a precise error
        ('model not pulled' vs 'Ollama not running') instead of guessing."""
        try:
            tags = self._get("/api/tags")
        except AIEngineError:
            return False, False
        names = {m.get("name", "") for m in tags.get("models", [])}
        # tags come back as e.g. "qwen2.5:7b-instruct"; a bare model name
        # (no tag) matches its ":latest" form
        want = self.model if ":" in self.model else self.model + ":latest"
        return True, want in names

    def ask(self, task_prompt: str, schema: dict | None = None, *,
            system: str | None = None,
            temperature: float = DEFAULT_TEMPERATURE) -> dict:
        """One narrow task -> one parsed, schema-validated dict.

        Raises AIEngineError after the single allowed retry fails; the
        caller fails that block, not the whole unit."""
        prompt = task_prompt
        last_why = None
        for attempt in (1, 2):
            if attempt == 2:
                prompt = task_prompt + RETRY_REMINDER.format(why=last_why)
            content = self._chat(prompt, schema, system, temperature)
            try:
                data = json.loads(strip_json_fences(content))
            except json.JSONDecodeError as e:
                last_why = f"it was not parseable JSON: {e}"
                continue
            if schema is not None:
                try:
                    jsonschema.validate(data, schema)
                except jsonschema.ValidationError as e:
                    last_why = (f"it did not match the required structure: "
                                f"{e.message}")
                    continue
            return data
        raise AIEngineError(
            f"model {self.model!r} failed twice to produce a valid "
            f"response ({last_why})")

    # -- internals ----------------------------------------------------------

    def _chat(self, prompt, schema, system, temperature) -> str:
        messages = []
        if system:
            messages.append({"role": "system", "content": system})
        messages.append({"role": "user", "content": prompt})
        payload = {
            "model": self.model,
            "messages": messages,
            "stream": False,
            # a real JSON schema constrains generation server-side
            # (Ollama >= 0.5 structured outputs); "json" alone still forces
            # JSON mode on older servers
            "format": schema if schema is not None else "json",
            "options": {
                "temperature": temperature,
                "num_ctx": self.num_ctx,
            },
        }
        reply = self._post("/api/chat", payload)
        return reply.get("message", {}).get("content", "")


_default_engine = None


def get_engine() -> OllamaEngine:
    global _default_engine
    if _default_engine is None:
        _default_engine = OllamaEngine()
    return _default_engine
