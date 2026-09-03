# CLAUDE.md — PPSU SLM Generator

Read this at the start of every session. Update the "Current state" section as phases complete.

---

## What we're building

A tool for **Maieutic Edutech** that generates complete, PPSU-branded **Self-Learning Material (SLM)** units — the ~45-page academic booklets P P Savani University issues per course unit (cover page, TOC, Introduction, Learning Objectives, numbered teaching sections with worked problems, Summary, Glossary, Case Study, Self-Assessment, Terminal Questions, Answers, References).

**The output is a document (.docx), NOT slides.** This is the opposite direction from the existing PPT designer (`../ppsu1`), which consumes finished content. Eventually: SLM generator → SLM doc → (existing pipeline) raw PPT → designed PPT. Do not conflate the two tools; this one only produces the SLM document.

### Input modes (all must work)

1. **Textbook mode** — user uploads source material (PDF/DOCX textbook chapter(s)) plus unit metadata (programme, course code/name, unit number/title, syllabus topics). The AI **restructures and rewrites only what the source contains** — it must never invent facts beyond the source.
2. **TOC mode (`toc+ai`)** — no textbook content, but the textbook's **table of contents** is supplied (pasted text). The TOC is the authoritative outline skeleton: it pins the unit's structure, topic ordering and coverage; the AI generates the teaching content for each entry from its own knowledge. When syllabus topics are ALSO given, both feed the outline (the TOC fixes ordering/coverage, the syllabus fills gaps). A TOC pins structure, not facts — this is still AI-generated content and is flagged for the same harder SME review as mode 3.
3. **No-textbook mode (`ai`)** — user provides only the metadata + syllabus topics. The AI generates outline and content entirely from its own knowledge. Mark generated-from-AI units clearly in the tool's report (SMEs must review harder).

The mode is chosen automatically: textbook present → textbook mode; else TOC present → `toc+ai`; else `ai`. `meta.source_mode` records which.

---

## Tech stack

- Python 3.11+ (backend), FastAPI (mirrors the org's existing tools)
- **Ollama** for all AI calls — local, free, no API keys. This machine is the "strong PC" chosen to run it.
- `python-docx` for document assembly; PyMuPDF (`fitz`) for PDF text extraction; `python-docx`/`docx` for DOCX extraction
- Frontend: single static `index.html` like `../ppsu1/frontend` (upload form → progress → download). Keep the same LAN pattern (`BACKEND_HOST` constant, `run.bat`).

### Ollama specifics (important)

- Model: default `qwen2.5:14b-instruct`; if the GPU has ≥24 GB VRAM prefer `qwen2.5:32b-instruct` or `llama3.3:70b` (quantised). Make the model an env var `OLLAMA_MODEL`; host `OLLAMA_HOST` (default `http://localhost:11434`).
- ALWAYS request JSON with Ollama's structured output (`format: json`, or a JSON schema via `format={...}` on newer versions). Still defensively strip ``` fences and retry once on `json.JSONDecodeError` with an "output ONLY valid JSON" reminder appended. A second failure fails that block, not the whole unit.
- Set `num_ctx` explicitly (16384 minimum; 32768 if RAM allows) — Ollama's default 2k/4k context silently truncates textbook chunks and the model "forgets" the source. This is the #1 silent quality killer.
- `temperature 0.3` for content generation, `0.1` for extraction/classification.
- One unit = many small AI calls (one per block type per section), NOT one giant call. Small calls are reliable on local models; a 45-page single-shot generation is not.
- Local models are weaker than cloud APIs: keep every prompt narrow, give one example of the expected JSON in each prompt, and validate every response against the schema below before accepting it.

### AI layer must be pluggable

Wrap all AI calls behind one interface (`services/ai_engine.py`, e.g. `ask(task_prompt, schema) -> dict`) with the Ollama client as the only current implementation. Never call Ollama directly from routes or the document builder. (Same rule as the LMS `PaymentGateway` convention — a future Claude/Grok backend must be a drop-in.)

### Reuse from REVA-AI-PPT-Creator (github.com/MaieuticEdutech/REVA-AI-PPT-Creator)

- `backend/app/prompts/global_rules.txt` — the house style rulebook (UK English, Bloom's verbs, "exactly four learning objectives", JSON-only). Copy it in and extend; it IS the SLM voice.
- `slm_reader.py` (PDF/DOCX text extraction) and `ollama_service.py` are usable starting points.
- ⚠️ That repo has a committed `backend/.env` with an `XAI_API_KEY` — treat as leaked; do not reuse; never commit `.env` here (gitignore it from the first commit).

---

## The SLM document specification

Decoded from a real issued SLM (MSc Data Science, "Discrete Mathematics Unit 1", ICSH7010). Put a sample PDF in `samples/` and re-check against it whenever formatting is in doubt.

### Document skeleton, in order

1. **Cover page** — PPSU logo, "SELF-LEARNING MATERIAL" banner, programme line ("MSc Data Science: Semester (I)"), course name. Red/navy/orange PPSU chrome.
2. **Unit heading** — "Unit NN: <Title>" + **Table of Contents** with page numbers (Word TOC field, updateable).
3. **Introduction** — ~3 paragraphs: why the subject matters, what the unit covers, section-by-section roadmap.
4. **Learning Objectives** — "By the end of this unit, you will be able to:" + **exactly 4–5 bullets**, each starting with a bolded Bloom's verb (Define/Construct/Enumerate/Apply/Analyse…), 8–12 words, one measurable outcome. Never "Understand/Know/Learn".
5. **Teaching sections** numbered `N.1`, `N.2`, `N.3` (typically three majors per unit), subsections `N.1.1`, `N.1.2`, sub-subs `N.1.1.1` where needed. Each section mixes these repeating **block types**:
   - `prose` — explanatory paragraphs, UK English, formal but readable
   - `table` — captioned "Table N.1.1: <title>", concept→meaning→application columns
   - `did_you_know` — grey box, historical/contextual aside, one paragraph
   - `problem` + `solution` — orange boxes: "Problem N.M" statement, then fully worked "Solution" with verification ticks (✓); several per section, increasing difficulty
   - `key_takeaway` — bold-led paragraph closing each subsection: "**Key Takeaway:** …"
   - `think_and_apply` — orange box, open-ended applied prompt (no solution given)
   - `figure` — captioned "Figure N: <title>"; generator inserts a placeholder box with the caption for the DTP team unless a chart can be drawn programmatically
6. **N.4 Summary** — one bullet per major concept, mirroring section order.
7. **N.5 Glossary** — two-column table, alphabetised Term | Definition (≈20–25 terms).
8. **N.6 Case Study** — one titled applied scenario with given data, then "Questions" each followed by its worked answer.
9. **N.7 Self-Assessment Questions** — "A. Multiple Choice Questions": 8 MCQs, options a)–d). "B. Fill in the Blanks": 5 items.
10. **N.8 Terminal Questions** — "Short Questions": 5 computational items. "Long Questions": 5 essay/derivation items.
11. **N.9 Answers** — `N.9.1` self-assessment answers (letter + one-line justification each), `N.9.2` terminal answers (fully worked, incl. long-question model essays).
12. **N.10 Suggested Books and References** — APA-style bullet list, real books relevant to the subject (in textbook mode, put the source textbook first).

### House style (enforce in prompts AND validate in code)

- UK English (analyse, summarise, colour).
- Textbook mode: content derived from source only; never invent facts; preserve the source's teaching sequence.
- Every problem's solution ends with verification where possible ("Verify: 23+12+16+9=60 ✓").
- Headers/footers: course code + unit number on every page, page numbers.
- Data-science flavouring: where the subject allows, tables/examples connect concepts to applications (the sample maps set theory to SQL).

### The intermediate JSON (AI output ↔ docx builder input)

The AI half fills this; the builder renders it. Keep them decoupled — the builder must render a hand-written JSON perfectly with no AI involved (that's how it's tested).

**This is the schema as actually implemented** (`backend/docx_builder.py`),
corrected against the real reference sample — see "Current state" for what
changed from the original pre-sample sketch and why:

```json
{
  "meta": {"programme": "", "course_code": "", "course_name": "", "unit_number": 1, "unit_title": "", "source_mode": "textbook|ai"},
  "introduction": ["para", "para", "para"],
  "learning_objectives": [{"verb": "Define", "rest": "…"}],
  "sections": [
    {"number": "1.1", "title": "", "intro": "",
     "subsections": [
       {"number": "1.1.1", "title": "",
        "blocks": [
          {"type": "prose", "text": ""},
          {"type": "table", "caption": "", "columns": [], "rows": [[]]},
          {"type": "did_you_know", "text": ""},
          {"type": "code", "text": ""},
          {"type": "problem", "label": "Problem 1.1", "statement": "", "solution": ""},
          {"type": "key_takeaway", "text": ""},
          {"type": "think_and_apply", "title": "", "text": ""},
          {"type": "figure", "caption": "", "placeholder": true}
        ]}]}
  ],
  "summary": [""],
  "glossary": [{"term": "", "definition": ""}],
  "case_study": {"title": "", "background": ["para", "para"], "questions": ["question text, no paired answer"]},
  "self_assessment": {"mcq": [{"q": "", "options": ["","","",""], "answer": "b", "why": "optional"}], "fill_blanks": [{"q": "", "answer": ""}]},
  "terminal": {"short": [{"q": "", "answer": ""}], "long": [{"q": "", "answer": ""}]},
  "references": [""]
}
```

Notes on fields that moved since the original sketch:
- `background` is a list of paragraphs, not one string.
- `case_study.questions` is flat strings, not `{"q","a"}` — this unit's
  case study has no provided answers; a future unit that DOES pair
  answers to case-study questions would need a schema addition, not
  reuse of the old shape (the builder doesn't support `{"q","a"}` there).
- `self_assessment.fill_blanks` may be `[]` — renders nothing, not an
  empty heading.
- There is no top-level `answers` object — see "Current state".

### Generation order (per unit)

1. Ingest textbook (if any) → split into topic chunks (reuse the numbered-heading splitter idea from `content_processor.py`).
2. AI call: unit outline (3 major sections + subsection titles) — from syllabus + source TOC.
3. Per subsection, sequential AI calls: prose blocks → tables → problems (+solutions) → did-you-know → key takeaway → think-and-apply. Feed each call ONLY that subsection's source chunk (textbook mode) or the outline context (AI mode).
4. AI calls for back matter: summary, glossary, case study, self-assessment, terminal Qs, answers, references.
5. Validate the assembled JSON (schema + counts: 4–5 objectives, 8 MCQs, 5 blanks, 5 short, 5 long).
6. Render `.docx`. Optionally export PDF via Word COM (Windows) or LibreOffice — reuse the pattern in `../ppsu1/backend/render.py`.

---

## Conventions (non-negotiable)

- **AI behind the interface.** No Ollama imports outside `services/ai_engine.py` + the ollama client module.
- **Builder is deterministic.** `docx_builder.py` takes JSON in, document out. No AI calls, no randomness. All PPSU styling (colours, box shading, fonts) lives in one `styles.py`.
- **Validate before render.** A unit JSON failing schema/count checks never reaches the builder; the report says which block failed and why.
- **Tests ship with features.** Golden-JSON → docx test (open the docx with python-docx and assert structure); extraction tests on a sample PDF; JSON-repair tests for fenced/malformed model output. Run tests before claiming a phase done.
- **Never commit secrets or `.env`.** Also gitignore `uploads/`, `output/`, `__pycache__/`.
- **Source fidelity in textbook mode.** Prompts must carry "use only the supplied content; never invent facts" (from global_rules.txt), and generated units must record `source_mode` so reviewers know the provenance.
- **Quality over speed.** Flag tradeoffs; don't silently take shortcuts.

## Working agreements for Claude Code

- Before adding a major dependency or changing the JSON schema: propose it, wait for confirmation.
- After each phase: summarise what was built, what's tested, what's left, in plain text.
- When blocked by something outside your control (model not pulled, GPU too small, no sample PDF in `samples/`), say so explicitly — don't guess around it.
- Keep this file's "Current state" section updated as phases complete.

---

## Build phases

1. **Docx builder + styles** — render a hand-written golden JSON into a correctly formatted SLM (cover, TOC, all block types, back matter). No AI. *This proves the hard half first.*
2. **Ollama engine** — `ai_engine.py` with JSON-mode calls, schema validation, retry/repair; smoke-test against the local model.
3. **AI mode (no textbook)** — outline + per-block generation + back matter for a topic; assemble + validate + render end to end.
4. **Textbook mode** — PDF/DOCX ingestion, chunking, source-grounded generation; provenance in the report.
5. **Web UI + run.bat** — upload form, progress polling, download; same LAN conventions as `../ppsu1`.
6. **Polish** — PDF export, figure placeholders list for the DTP team, multi-unit batch.

## Current state

**Branding & reference-format feature added, 2026-09-03** (user request).
Set-once institution branding in `assets/` (gitignored — per-installation
uploads, not source), applied to every generated unit until replaced:

- **Logo upload** (.png/.jpg): embedded on the cover and in every page
  header (logo left, unit label pushed right via a tab stop — the real
  issued sample's running-header layout).
- **Reference SLM PDF upload**: `branding.extract_profile` (PyMuPDF) reads
  its STYLE SIGNALS — dominant font family (subset-prefix/style-suffix
  cleaned), body/heading point sizes, heading colour (most common
  non-black colour among larger-than-body text), page size and margins
  (clamped 8-40 mm; the cover page is skipped so display sizes don't skew
  stats) — and overrides the builder defaults. HONEST scope, stated to the
  user: style adaptation, not pixel-cloning; the document structure always
  follows the decoded PPSU format. The PDF is kept for side-by-side review.

Plumbing worth knowing: `docx_builder` was converted from frozen
`from styles import CONSTANTS` to live `st.<NAME>` attribute access so
`styles.apply_profile()` overrides are visible; apply_profile ALWAYS
resets to `_DEFAULTS` first so unbranded builds are never polluted by a
previous branded one; `set_run`/`add_paragraph` defaults resolve at call
time (def-time defaults would freeze the originals).
`build(data, use_branding=False)` bypasses assets (tests assert the
default look with it). Endpoints: GET/POST `/api/branding`; frontend has
a "Branding & reference format" card showing current state.
`tests/test_branding.py`: 14 checks (synthetic reference PDF with known
font/sizes/colours round-trips through extraction; logo embeds into
word/media; geometry applied; defaults restored; API validation) — 126
offline checks across six suites.

**Phases 5+6 (web UI + run.bat; PDF export, DTP figure list, batch) built,
2026-09-02.** `backend/app.py`: FastAPI on ONE port (8010) serving the
frontend AND the API — a deliberate deviation from ppsu1's two-server
BACKEND_HOST pattern, because that pattern's separate origin caused a full
afternoon of CORS/IP-drift debugging during the PPT designer's deployment
(../DEPLOYMENT.md); single-origin + relative fetch paths makes the whole
failure class impossible while LAN sharing still works. Jobs run in a
background thread ONE at a time (`_RUN_LOCK` — a single 8 GB GPU
serialises generation anyway; a second submit gets 409 and the page says
so). Endpoints: `/` (frontend), `/api/status` (incl. pdf_export
availability), `/api/generate` (multipart: meta fields + syllabus/TOC
textareas + optional source file; FastAPI gotcha fixed — form fields need
`Form()` annotations or they silently become query params),
`/api/progress/{job}` (the generator's own per-call labels + a result
summary), `/api/download/{job}/{docx|pdf|figures|report}`. A unit failing
validation is never rendered: the job fails naming the errors, the report
stays downloadable for diagnosis. `frontend/index.html`: single static
page, PPSU-styled, mode auto-detection explained inline, live progress,
download buttons (PDF button only when the export engine exists), review
note + warnings surfaced. `run.bat`: one server; creates the venv +
installs on first run; checks Ollama and pulls the model if missing;
prints/refreshes the LAN link + "Open SLM Generator.url".

Phase 6 pieces: `backend/docx2pdf.py` (Word COM first — office machine has
Office — LibreOffice fallback, serialised behind a lock, gracefully
"unavailable" otherwise: the docx alone is offered); `backend/figures.py`
(flat DTP handoff list: every figure placeholder with section/subsection
location + caption, saved per job and via CLI as *.figures.txt);
multi-unit batch: `unit_generator.py --batch samples/batch_example.json
--out-dir DIR` — sequential (one GPU), one unit's failure never stops the
rest, per-unit docx/json/report/figures + a summary table.

`tests/test_app.py`: 18 checks (TestClient, generator faked — routes,
job lifecycle, busy-409 + lock release, validation-failure path, all
downloads). 111 offline checks across five suites. NOTE: test_app's
busy-lock section is timing-sensitive and can flake under heavy machine
load; rerun it solo before trusting a failure.

**Phase 4 LIVE textbook-mode run passed, 2026-09-02**: a full unit
generated from samples/source_chapter_example.txt in 265s — 28 calls, 0
failures, 11 UK fixes, validation clean. Grounding fidelity verified: the
source's specific facts (SECURITY→VHFXULWB, Al-Kindi/9th-century frequency
analysis, the rail-fence MEETMEATNOON example, E=12.7%, the 26! keyspace)
all survived into the output, and prose reads as a faithful rewrite of the
source rather than free-styled model knowledge. Two lessons from this run:

- **Runaway-generation fix (ai_engine)**: the FIRST attempt hung — one
  prose call pegged the GPU at 98% for 8+ minutes. Under schema-constrained
  decoding a small model can loop forever inside a JSON string. Every call
  now carries `num_predict` (default 2048, env OLLAMA_NUM_PREDICT): a
  runaway is cut off → parse failure → the normal single retry with fresh
  sampling. Transport failures (timeouts) now share that same one-retry
  budget. Request timeout dropped to 300s. Regression-tested (engine suite
  16 checks).
- **Syllabus broader than the source** (deliberate in the sample fixtures):
  the outline blended syllabus topics the chapter doesn't cover (RSA,
  hashing), producing subsections the matcher correctly could NOT map to
  source sections — they fell back to the whole-source digest and the
  report warned, naming them. Working as designed, but reviewers should
  treat digest-fallback subsections as effectively ai-mode content. A
  future refinement: in textbook mode, tell the outline call to prefer
  source coverage and drop syllabus topics absent from the source.

**Phase 4 (textbook mode) built, 2026-09-02.** `backend/ingest.py`:
PDF (PyMuPDF) / DOCX / TXT extraction; a numbered-heading chunker
("Chapter N", "2.1 Title", "2.1.1 Title" — short lines not ending in a
full stop, so numbered list items in prose don't split the text); heading
matching (exact-normalised then substring, numbering/case-insensitive);
and `condense()` — a per-chunk-truncated whole-source digest for calls
that must span the unit (back matter) when a full chapter would blow
num_ctx. Generator wiring: mode precedence textbook > toc+ai > ai; in
textbook mode the outline call gets the detected source headings and must
return, per subsection, which verbatim headings it teaches
(OUTLINE_TEXTBOOK schema) — each subsection's calls then carry ONLY its
own mapped chunk(s) (capped 6000 chars) with the STRICT SOURCE RULES
grounding block (from global_rules.txt: use only the source, never invent
facts, preserve teaching sequence); unmatched subsections fall back to the
condensed digest with a warning; back matter is grounded on the digest;
the source textbook is placed FIRST in references (meta.textbook_citation
if given, else a filename stub marked for the SME to complete). Report
carries provenance (file/chars/chunks/unmatched headings). A source with
no detectable numbered headings still works: whole document grounds every
call, plain OUTLINE schema, split-failure warning. CLI: `--source
chapter.pdf|.docx|.txt`. `tests/test_textbook_mode.py`: 22 offline checks
(ingestion round-trips incl. a generated PDF/DOCX, chunker guards,
matching, digest, full stubbed textbook-mode run, no-headings fallback),
all passing — 91 offline checks across the four suites now.

**Phase 3 LIVE end-to-end run passed, 2026-09-02**: a full unit
("Fundamentals of Cryptography", toc+ai mode, samples/meta_example.json +
samples/toc_example.txt) generated in 300s — 30 AI calls, 0 failures, 9
automatic UK-spelling fixes, validation clean, real docx rendered. The TOC
demonstrably steered the outline (Caesar → Feistel → stream → modes → RSA
→ Diffie-Hellman, straight from the supplied TOC), and the model correctly
chose problem-style worked examples for a crypto unit (a correct
step-by-step Caesar worked problem with mod-26 arithmetic). Two 7b quirks
found in that run and fixed deterministically in code (plus prompt
tightening): TOC bookkeeping leaking into outline titles ("... (2.1.1 in
the textbook TOC)") — `clean_title()`; and option letters embedded inside
MCQ option text ("a) AES", which would double-prefix in the docx) —
`clean_mcq_options()`. Both regression-tested; the committed
output/-excluded sample docx predates the title fix.

**Phase 3 (AI-mode generation, incl. the new TOC modes) built, 2026-09-02.**
`backend/unit_generator.py` orchestrates the full pipeline: outline (the one
call that MUST succeed — anything else failing twice is recorded in the
report and skipped, the unit still completes) → front matter → per-
subsection content → back matter → UK-spelling pass → validation gate →
render via the Phase 1 builder. Numbering (N.1/N.1.1, Table/Figure/Problem
numbers) is assigned in code, never trusted from the model. Enrichment
blocks rotate deterministically per subsection (1st: table, 2nd: worked
example — `code` or `problem` per the outline call's `example_style`
decision for the subject, 3rd: did-you-know; figure on the first, think-
and-apply closing the last), mirroring the reference sample's rhythm
without fragile oneOf schemas. Supporting modules: `schemas.py` (per-call
+ full-unit schemas), `prompts.py` (house voice adapted from
`prompts/global_rules.txt`, copied verbatim from REVA-AI-PPT-Creator —
only that file; the temp clone with its leaked .env was deleted),
`uk_style.py` (conservative US→UK respelling, code blocks exempt,
ambiguous pairs like program/programme untouched), `validate_unit.py`
(errors block rendering; warnings surface to the reviewer — references
always warned as needing SME verification since local models can cite
non-existent editions). CLI: `unit_generator.py --meta m.json [--toc
toc.txt] --out unit.docx [--json-out unit.json] [--report r.json]`;
sample inputs in `samples/meta_example.json` + `samples/toc_example.txt`.
`tests/test_unit_generator.py`: 24 offline checks (stub engine dispatched
by schema identity), all passing — mode detection, TOC-in-outline-prompt,
numbering, rotation placement, UK pass (incl. code exemption), failure
containment, validator gate, end-to-end render.

**Phase 2 (Ollama engine) done and live-tested, 2026-09-02.** Ollama 0.33.2
installed (winget, per-user, server auto-runs from the tray);
`qwen2.5:7b-instruct` pulled. Model default is deliberately **7b, not the
spec's 14b**: the 14b Q4 weights (~9 GB) plus a 16k KV cache don't fit the
RTX 4060's 8 GB VRAM without heavy CPU offload; the 7b fits fully and
generates in 3-10s per call. Override via `OLLAMA_MODEL` on a bigger card.

`backend/ai_engine.py`: the single `ask(task_prompt, schema) -> dict`
interface (nothing else in the codebase may import/talk to Ollama). Real
JSON schemas are passed to Ollama's `format` field so generation is
constrained server-side (structured outputs); client-side the reply is
fence-stripped, parsed, and `jsonschema`-validated, with exactly ONE
corrective retry on a parse OR validation failure — a second failure
raises `AIEngineError` and fails that block, not the unit. `num_ctx`
defaults to 16384 explicitly. `available()` distinguishes "server down"
from "model not pulled" so callers can say which.

Tests: `tests/test_ai_engine.py` (14 offline checks, model stubbed —
fence-stripping variants, the full parse/validate/retry contract, the
two-attempts-never-three guarantee) and `tests/smoke_ai_engine.py` (LIVE,
needs server+model; generates MCQs / Bloom's-verb objectives / a
temperature-0.1 extraction against real schemas and prints the content
for human quality judgement). Both passing on this machine.

Known quality item for Phase 3 prompts: the model produced "Analyze" (US
spelling) despite a UK-English system prompt — the house-style rulebook
(`global_rules.txt` from REVA-AI-PPT-Creator) must be baked into the real
prompts and UK spelling ideally validated in code, not just requested.

**Phase 1 (docx builder + styles) done and tested, 2026-09-02.** Correction
to this file's earlier claim: the office PC turned out to have an RTX 4060
(8 GB VRAM) + 64 GB RAM — capable of running Ollama for Phase 2+ too, not
just the AI-free Phase 1. Ollama itself is still not installed.

The real reference sample arrived (P P Savani University, ICCS7010 "Information
Security and Applications" Unit 1, "Foundations of Discrete Mathematics" —
note the course/unit title mismatch in the source file itself, not a
transcription error here) and it disagreed with this spec's pre-sample
guesses in several concrete ways — the code now matches the REAL sample,
not the paragraph above:

- **New block type: `code`.** Not in the original block list at all. This
  unit is data-science-flavoured and has a Python/SQL snippet in nearly
  every subsection (1.1.3, 1.2.1, 1.2.2, 1.2.3, 1.2.4, 1.3.2 ×2, 1.3.3) —
  far more common here than `problem`/`solution` or `key_takeaway`, which
  this unit uses **zero** times. Block-type mix is topic-dependent; the
  builder renders whichever subset of the (now 8) known types a unit's
  JSON contains, not a fixed set per unit.
- **`case_study.questions`** is a flat list of strings, not `{"q","a"}`
  pairs — the real case study's questions have no provided answers at all.
- **Self-assessment**: this unit has 15 MCQs (not the ~8 guessed) and NO
  "Fill in the Blanks" subsection — `fill_blanks` can be `[]`, and the
  builder skips that heading entirely when it is.
- **No separate "answers" object.** Each MCQ/terminal item carries its own
  `answer` (+ optional `why`, unused by this unit — the real answer key is
  just "letter) option text", no separate justification prose); the
  builder renders the bare question under N.7/N.8 and the answer view
  under N.9 from the SAME object, so nothing is written twice in the JSON.
- Glossary: 16 terms here, not "≈20-25" — cosmetic, no schema impact.
- PPSU brand colours are no longer guessed from the PDF render: pulled
  directly from `../ppsu1/template`'s slide master shape fills (the
  template's `theme1.xml` colour scheme is just unmodified Office default,
  not useful) — navy `#0E2841`, red `#D8181F`, orange `#F47820`,
  gold `#FBB217`. See `backend/styles.py`.

Built: `backend/styles.py`, `backend/docx_builder.py` (cover, running
header/footer, unit heading + a REAL updateable Word TOC field via raw
`w:fldChar`/`instrText` XML — not a static hand-typed page list, all 8
block types, back matter). `tests/golden_unit1.json` is a representative
(not exhaustive — the real unit is 33 pages) transcription of the real
sample; `tests/block_types_synthetic.json` is a small synthetic fixture
proving `problem`/`solution` and `key_takeaway` render correctly even
though the real sample never exercises them. `tests/test_docx_builder.py`:
26 checks, all passing — builds both fixtures, reopens the output with
python-docx, and asserts real structural facts (heading text, table
counts, verbatim block content, glossary alphabetisation, TOC field
presence, header/footer content, valid-zip/OOXML integrity). Own venv at
`backend/venv/` (python-docx only so far — kept separate from ppsu1's venv
since FastAPI/Ollama-client deps will diverge from it in Phase 2+).

**Not done / next**: the real PDF itself still needs to be dropped into
`samples/` (I only had its content pasted into a conversation, not the
file — `samples/README.md` explains what's needed and why). No PPSU logo
image asset yet — the cover renders a labelled placeholder box instead
of the real photography/logo, which needs actual asset files from
whoever has PPSU's brand kit. Phase 2 (Ollama engine) has not started —
Ollama is not installed on this machine yet.

Related repos: `MaieuticEdutech/Template_Designer` — **name unverified,
worth confirming this actually exists on GitHub as a repo distinct from
`MaieuticEdutech/PPSU-PPT-Template`.** The PPT-designer work done in this
same working tree this session was pushed directly to
`MaieuticEdutech/PPSU-PPT-Template` (git remote `origin`) — there was no
separate `Template_Designer` remote involved, so either that name is stale
or it refers to something not yet reconciled with this checkout. The PPT
designer itself lives in `../ppsu1` (deployed to Render+Vercel, see
`../DEPLOYMENT.md`) and `../reva1` (local-only fork, not deployed).
`MaieuticEdutech/REVA-AI-PPT-Creator` (prompt rulebook + extraction code
to reuse in Phase 2 — has a leaked `XAI_API_KEY` committed, see
"Conventions" above, do not reuse carelessly).
