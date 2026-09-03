"""JSON schemas for every AI call the unit generator makes, plus the schema
the fully-assembled unit is validated against before it may reach the docx
builder (CLAUDE.md: "A unit JSON failing schema/count checks never reaches
the builder").

Each per-call schema is passed to ai_engine.ask(), which forwards it to
Ollama's `format` field (server-side constrained generation) AND validates
the reply client-side. Tests stub the engine and dispatch canned replies by
schema identity, so keep these as module-level constants.
"""

_STR = {"type": "string", "minLength": 1}


def _arr(items, lo, hi):
    return {"type": "array", "items": items, "minItems": lo, "maxItems": hi}


# --- outline ---------------------------------------------------------------
# Numbering is NOT requested from the model: unit_generator assigns N.1,
# N.1.1 etc. in code, so a model miscount can never corrupt the numbering.
OUTLINE = {
    "type": "object",
    "properties": {
        "sections": _arr({
            "type": "object",
            "properties": {
                "title": _STR,
                "intro": _STR,           # 1-2 sentence section opener
                "subsections": _arr({"type": "object",
                                     "properties": {"title": _STR},
                                     "required": ["title"]}, 2, 3),
            },
            "required": ["title", "intro", "subsections"],
        }, 3, 3),
        # worked-example style for the whole unit: computational subjects
        # get `code` blocks, mathematical ones get problem/solution boxes
        "example_style": {"type": "string", "enum": ["code", "problem"]},
    },
    "required": ["sections", "example_style"],
}

# Textbook mode: same outline, but every subsection must ALSO name which
# source headings (verbatim strings from the supplied list) it teaches, so
# the generator can feed each subsection call ONLY its own source chunk.
OUTLINE_TEXTBOOK = {
    "type": "object",
    "properties": {
        "sections": _arr({
            "type": "object",
            "properties": {
                "title": _STR,
                "intro": _STR,
                "subsections": _arr({
                    "type": "object",
                    "properties": {
                        "title": _STR,
                        "source_headings": _arr({"type": "string"}, 1, 4),
                    },
                    "required": ["title", "source_headings"]}, 2, 3),
            },
            "required": ["title", "intro", "subsections"],
        }, 3, 3),
        "example_style": {"type": "string", "enum": ["code", "problem"]},
    },
    "required": ["sections", "example_style"],
}

# --- front matter ----------------------------------------------------------
INTRODUCTION = {
    "type": "object",
    "properties": {"paragraphs": _arr(_STR, 3, 4)},
    "required": ["paragraphs"],
}

LEARNING_OBJECTIVES = {
    "type": "object",
    "properties": {
        "learning_objectives": _arr({
            "type": "object",
            "properties": {"verb": _STR, "rest": _STR},
            "required": ["verb", "rest"],
        }, 4, 5),
    },
    "required": ["learning_objectives"],
}

# --- per-subsection content ------------------------------------------------
PROSE = {
    "type": "object",
    "properties": {"paragraphs": _arr(_STR, 2, 3)},
    "required": ["paragraphs"],
}

TABLE = {
    "type": "object",
    "properties": {
        "caption_title": _STR,       # caption text after "Table N.M.K: "
        "columns": _arr(_STR, 2, 4),
        "rows": _arr(_arr({"type": "string"}, 2, 4), 3, 6),
    },
    "required": ["caption_title", "columns", "rows"],
}

CODE = {
    "type": "object",
    "properties": {
        "language": {"type": "string", "enum": ["python", "sql"]},
        "code": _STR,
        "explanation": _STR,         # 1 short paragraph after the snippet
    },
    "required": ["language", "code", "explanation"],
}

PROBLEM = {
    "type": "object",
    "properties": {
        "statement": _STR,
        "solution": _STR,            # fully worked, ends with verification
    },
    "required": ["statement", "solution"],
}

DID_YOU_KNOW = {
    "type": "object",
    "properties": {"text": _STR},
    "required": ["text"],
}

SECTION_EXTRAS = {
    "type": "object",
    "properties": {
        "think_and_apply": _STR,     # open applied prompt, no solution
        "figure_caption": _STR,      # caption text after "Figure N: "
    },
    "required": ["think_and_apply", "figure_caption"],
}

# --- back matter -----------------------------------------------------------
SUMMARY = {
    "type": "object",
    "properties": {"summary": _arr(_STR, 6, 10)},
    "required": ["summary"],
}

GLOSSARY = {
    "type": "object",
    "properties": {
        "glossary": _arr({
            "type": "object",
            "properties": {"term": _STR, "definition": _STR},
            "required": ["term", "definition"],
        }, 12, 20),
    },
    "required": ["glossary"],
}

CASE_STUDY = {
    "type": "object",
    "properties": {
        "title": _STR,
        "background": _arr(_STR, 3, 5),
        "questions": _arr(_STR, 3, 3),
    },
    "required": ["title", "background", "questions"],
}

MCQS = {
    "type": "object",
    "properties": {
        "mcq": _arr({
            "type": "object",
            "properties": {
                "q": _STR,
                "options": _arr(_STR, 4, 4),
                "answer": {"type": "string", "enum": ["a", "b", "c", "d"]},
            },
            "required": ["q", "options", "answer"],
        }, 8, 8),
    },
    "required": ["mcq"],
}

FILL_BLANKS = {
    "type": "object",
    "properties": {
        "fill_blanks": _arr({
            "type": "object",
            "properties": {"q": _STR, "answer": _STR},
            "required": ["q", "answer"],
        }, 5, 5),
    },
    "required": ["fill_blanks"],
}

TERMINAL_SHORT = {
    "type": "object",
    "properties": {
        "short": _arr({"type": "object",
                       "properties": {"q": _STR, "answer": _STR},
                       "required": ["q", "answer"]}, 5, 5),
    },
    "required": ["short"],
}

# Long terminal questions are generated in SIX calls, not one: five
# multi-paragraph model essays in a single generation regularly exceeds the
# engine's num_predict cap (observed live 2026-09-03: the combined call was
# truncated mid-JSON twice -> 0 long questions -> validation blocked the
# unit). Questions first, then one essay per call — each well under the cap.
TERMINAL_LONG_QS = {
    "type": "object",
    "properties": {"long_questions": _arr(_STR, 5, 5)},
    "required": ["long_questions"],
}

LONG_ANSWER = {
    "type": "object",
    "properties": {"answer": {"type": "string", "minLength": 200}},
    "required": ["answer"],
}

REFERENCES = {
    "type": "object",
    "properties": {"references": _arr(_STR, 6, 8)},
    "required": ["references"],
}

# --- the assembled unit (validated before rendering) ------------------------
_BLOCK = {
    "type": "object",
    "properties": {"type": {"type": "string",
                            "enum": ["prose", "table", "did_you_know",
                                     "code", "problem", "key_takeaway",
                                     "think_and_apply", "figure"]}},
    "required": ["type"],
}

FULL_UNIT = {
    "type": "object",
    "properties": {
        "meta": {
            "type": "object",
            "properties": {
                "programme": _STR, "course_code": _STR, "course_name": _STR,
                "unit_number": {"type": "integer"}, "unit_title": _STR,
                "source_mode": {"type": "string",
                                "enum": ["textbook", "toc+ai", "ai"]},
            },
            "required": ["programme", "course_code", "course_name",
                          "unit_number", "unit_title", "source_mode"],
        },
        "introduction": _arr(_STR, 1, 6),
        "learning_objectives": _arr({"type": "object"}, 4, 5),
        "sections": _arr({
            "type": "object",
            "properties": {
                "number": _STR, "title": _STR,
                "subsections": {"type": "array",
                                "items": {"type": "object",
                                          "properties": {"blocks": {
                                              "type": "array",
                                              "items": _BLOCK}},
                                          "required": ["blocks"]}},
            },
            "required": ["number", "title", "subsections"],
        }, 3, 3),
        "summary": {"type": "array"},
        "glossary": {"type": "array"},
        "case_study": {"type": "object"},
        "self_assessment": {"type": "object"},
        "terminal": {"type": "object"},
        "references": {"type": "array"},
    },
    "required": ["meta", "introduction", "learning_objectives", "sections",
                  "summary", "glossary", "case_study", "self_assessment",
                  "terminal", "references"],
}
