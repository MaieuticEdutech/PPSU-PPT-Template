# CLAUDE.md — PPSU SLM Generator

Read this at the start of every session. Update the "Current state" section as phases complete.

---

## What we're building

A tool for **Maieutic Edutech** that generates complete, PPSU-branded **Self-Learning Material (SLM)** units — the ~45-page academic booklets P P Savani University issues per course unit (cover page, TOC, Introduction, Learning Objectives, numbered teaching sections with worked problems, Summary, Glossary, Case Study, Self-Assessment, Terminal Questions, Answers, References).

**The output is a document (.docx), NOT slides.** This is the opposite direction from the existing PPT designer (`../ppsu1`), which consumes finished content. Eventually: SLM generator → SLM doc → (existing pipeline) raw PPT → designed PPT. Do not conflate the two tools; this one only produces the SLM document.

### Input modes (both must work)

1. **Textbook mode** — user uploads source material (PDF/DOCX textbook chapter(s)) plus unit metadata (programme, course code/name, unit number/title, syllabus topics). The AI **restructures and rewrites only what the source contains** — it must never invent facts beyond the source.
2. **No-textbook mode** — user provides only the metadata + syllabus topics. The AI generates the teaching content itself from its own knowledge. Mark generated-from-AI units clearly in the tool's report (SMEs must review harder).

The mode is chosen automatically: textbook present → textbook mode, else AI mode.

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

```json
{
  "meta": {"programme": "", "semester": "", "course_code": "", "course_name": "", "unit_number": 1, "unit_title": "", "source_mode": "textbook|ai"},
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
          {"type": "problem", "label": "Problem 1.1", "statement": "", "solution": ""},
          {"type": "key_takeaway", "text": ""},
          {"type": "think_and_apply", "title": "", "text": ""},
          {"type": "figure", "caption": "", "placeholder": true}
        ]}]}
  ],
  "summary": [""],
  "glossary": [{"term": "", "definition": ""}],
  "case_study": {"title": "", "background": "", "questions": [{"q": "", "a": ""}]},
  "self_assessment": {"mcq": [{"q": "", "options": ["","","",""], "answer": "b", "why": ""}], "fill_blanks": [{"q": "", "answer": ""}]},
  "terminal": {"short": [{"q": "", "answer": ""}], "long": [{"q": "", "answer": ""}]},
  "references": [""]
}
```

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

**Nothing built yet.** This spec was written 2026-09-02 on the office PC (which has no Ollama and no GPU); development happens on the strong PC. The sample SLM PDF ("MSc_Sem 1_Discrete Mathematics_Unit 1.pdf") still needs to be copied into `samples/` — ask the user for it if missing.

Related repos: `MaieuticEdutech/Template_Designer` (this repo — the PPT designer lives in `../ppsu1`, `../reva1`), `MaieuticEdutech/PPSU-PPT-Template` (hosted deployment fork), `MaieuticEdutech/REVA-AI-PPT-Creator` (prompt rulebook + extraction code to reuse).
