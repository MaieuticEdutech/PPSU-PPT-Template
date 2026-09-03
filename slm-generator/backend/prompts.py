"""Prompt builders for every AI call in the unit generator.

The house voice is adapted from prompts/global_rules.txt (copied verbatim
from REVA-AI-PPT-Creator per CLAUDE.md — only that one file was taken; the
rest of that repo, including its leaked .env, was not). Slide-specific rules
(bullet limits, topic continuation) don't apply to a 45-page document; the
general rules, Bloom's-verb discipline, and definition/formula preservation
carry over directly.
"""

# Adapted from global_rules.txt GENERAL RULES + QUALITY, SLM-flavoured.
SYSTEM_STYLE = (
    "You write academic self-learning material for P P Savani University "
    "(PPSU). Rules:\n"
    "1. UK English ONLY (analyse, summarise, colour, programme). Never US "
    "spellings.\n"
    "2. Formal but readable teaching prose for postgraduate learners.\n"
    "3. Never change academic meaning; never modify mathematical formulae "
    "or programming syntax.\n"
    "4. Every new idea gets a short realistic example from business or "
    "data-processing contexts where natural.\n"
    "5. Accuracy over brevity; instructional quality over aesthetics; "
    "academic integrity over summarisation.\n"
    "6. Output ONLY valid JSON matching the requested structure — no "
    "markdown, no commentary."
)

# From global_rules.txt LEARNING OBJECTIVES — the verb discipline verbatim.
BLOOMS_VERBS = ("Define, Identify, List, Recall, State, Explain, Describe, "
                "Discuss, Illustrate, Summarise, Apply, Use, Demonstrate, "
                "Calculate, Solve, Analyse, Compare, Differentiate, "
                "Classify, Examine, Evaluate, Assess, Justify, Critique, "
                "Recommend, Design, Develop, Construct, Create, Formulate")
BANNED_VERBS = "Understand, Know, Learn, Be familiar with, Appreciate"

# Textbook-mode grounding — from global_rules.txt GENERAL RULES 2/3/5.
GROUNDING_RULES = (
    "STRICT SOURCE RULES: use ONLY the supplied source material below. "
    "Never invent facts beyond it. Never change its academic meaning. "
    "Preserve its logical teaching sequence. You may rephrase and "
    "restructure, but every fact must come from the source.")


def _with_source(prompt, source):
    if not source:
        return prompt
    return (f"{prompt}\n\n{GROUNDING_RULES}\n"
            f"--- SOURCE MATERIAL ---\n{source}\n--- END SOURCE ---")


def _unit_context(meta, syllabus_topics=None, toc_text=None):
    lines = [f"Programme: {meta['programme']}",
             f"Course: {meta['course_code']} {meta['course_name']}",
             f"Unit {meta['unit_number']}: {meta['unit_title']}"]
    if syllabus_topics:
        lines.append("Syllabus topics for this unit: "
                     + "; ".join(syllabus_topics))
    if toc_text:
        lines.append(
            "The prescribed textbook's table of contents for this unit's "
            "material is below. Follow ITS structure and topic ordering — "
            "it is the authoritative outline skeleton:\n" + toc_text.strip())
    return "\n".join(lines)


def outline(meta, syllabus_topics=None, toc_text=None,
            source_headings=None):
    if source_headings:
        headings_block = (
            "\nThe uploaded source material contains these sections:\n"
            + "\n".join(f"- {h}" for h in source_headings)
            + "\nBuild the outline FROM these source sections (preserving "
            "their teaching sequence), and for every subsection return "
            "'source_headings': the exact heading strings (copied verbatim "
            "from the list above) whose content that subsection teaches.")
    else:
        headings_block = ""
    return (
        f"{_unit_context(meta, syllabus_topics, toc_text)}{headings_block}\n\n"
        "Design this unit's teaching outline: exactly 3 major sections, "
        "each with 2-3 subsections. Section titles mirror the syllabus "
        "topics"
        + (" and the textbook TOC above (merge both: the TOC fixes the "
           "ordering and coverage, the syllabus fills any gaps)"
           if toc_text else "")
        + ". Titles must be clean academic headings ONLY — never include "
        "numbering, cross-references, or notes like '(2.1.1 in the "
        "textbook TOC)' inside a title. "
        "Each section gets a 1-2 sentence 'intro' opener. Also decide "
        "'example_style' for the unit's worked examples: 'code' if the "
        "subject is computational (programming, data science, databases), "
        "'problem' if it is mathematical/derivational.")


def introduction(meta, outline_titles):
    return (
        f"{_unit_context(meta)}\n"
        f"The unit's sections are: {outline_titles}.\n\n"
        "Write the unit's Introduction: 3-4 paragraphs — why the subject "
        "matters, what the unit covers, and a section-by-section roadmap. "
        "Welcome the learner to the unit in the first paragraph.")


def learning_objectives(meta, outline_titles):
    return (
        f"{_unit_context(meta)}\n"
        f"The unit's sections are: {outline_titles}.\n\n"
        "Write 4-5 learning objectives. Each splits into 'verb' (ONE "
        f"measurable Bloom's verb from: {BLOOMS_VERBS}) and 'rest' (8-12 "
        "words, exactly one measurable outcome). NEVER use these verbs: "
        f"{BANNED_VERBS}.")


def prose(meta, section_title, subsection_title, source=None):
    return _with_source((
        f"{_unit_context(meta)}\n"
        f"Section: {section_title}\nSubsection: {subsection_title}\n\n"
        "Write this subsection's teaching prose: 2-3 substantial "
        "paragraphs. Define every new concept precisely, then ground it "
        "with one short realistic example."), source)


def table(meta, section_title, subsection_title, source=None):
    return _with_source((
        f"{_unit_context(meta)}\n"
        f"Section: {section_title}\nSubsection: {subsection_title}\n\n"
        "Create one comparison/reference table for this subsection "
        "(concept -> meaning -> application style, 2-4 columns, 3-6 rows). "
        "'caption_title' is the caption text only (no 'Table N:' prefix — "
        "numbering is added automatically)."), source)


def code(meta, section_title, subsection_title, source=None):
    return _with_source((
        f"{_unit_context(meta)}\n"
        f"Section: {section_title}\nSubsection: {subsection_title}\n\n"
        "Write one short, correct, runnable code example (python or sql) "
        "demonstrating this subsection's concept, with inline # comments "
        "showing expected output, plus a one-paragraph 'explanation' of "
        "what it shows. Keep it under 15 lines."), source)


def problem(meta, section_title, subsection_title, source=None):
    return _with_source((
        f"{_unit_context(meta)}\n"
        f"Section: {section_title}\nSubsection: {subsection_title}\n\n"
        "Write one worked problem for this subsection: a precise "
        "'statement' with concrete numbers, and a fully worked step-by-step "
        "'solution' that ends with a verification line where possible "
        "(e.g. 'Verify: 23+12+16+9=60 ✓')."), source)


def did_you_know(meta, section_title, subsection_title, source=None):
    return _with_source((
        f"{_unit_context(meta)}\n"
        f"Section: {section_title}\nSubsection: {subsection_title}\n\n"
        "Write one 'Did you know?' aside for this subsection: a single "
        "paragraph of genuine historical or contextual interest. Only "
        "well-established facts — nothing invented."), source)


def section_extras(meta, section_title, subsection_titles, source=None):
    return _with_source((
        f"{_unit_context(meta)}\n"
        f"Section: {section_title} (subsections: {subsection_titles})\n\n"
        "Write (1) 'think_and_apply': one open-ended applied scenario "
        "prompt for this section — realistic business/data context, asks "
        "the learner to apply the section's ideas, gives NO solution; and "
        "(2) 'figure_caption': a caption for one illustrative figure the "
        "design team will draw (caption text only, no 'Figure N:' "
        "prefix)."), source)


def summary(meta, outline_titles, source=None):
    return _with_source((
        f"{_unit_context(meta)}\n"
        f"The unit's sections are: {outline_titles}.\n\n"
        "Write the unit summary: 6-10 bullets, one per major concept, in "
        "section order. Key takeaways only — no new information."), source)


def glossary(meta, outline_titles, source=None):
    return _with_source((
        f"{_unit_context(meta)}\n"
        f"The unit's sections are: {outline_titles}.\n\n"
        "Write the unit glossary: 12-20 term/definition pairs covering the "
        "unit's key vocabulary. One-sentence definitions preserving full "
        "academic meaning. (They are alphabetised automatically — any "
        "order is fine.)"), source)


def case_study(meta, outline_titles, source=None):
    return _with_source((
        f"{_unit_context(meta)}\n"
        f"The unit's sections are: {outline_titles}.\n\n"
        "Write one applied case study: a 'title' naming a fictional "
        "organisation and its problem, 3-5 'background' paragraphs telling "
        "how a named practitioner applied this unit's concepts (with "
        "concrete given data), and exactly 3 open 'questions' for the "
        "learner (no answers)."), source)


def mcqs(meta, outline_titles, source=None):
    return _with_source((
        f"{_unit_context(meta)}\n"
        f"The unit's sections are: {outline_titles}.\n\n"
        "Write exactly 8 multiple-choice questions spanning the whole "
        "unit, easier first. 4 options each; 'answer' is the correct "
        "letter a-d. Wrong options must be plausible, not silly."), source)


def fill_blanks(meta, outline_titles, source=None):
    return _with_source((
        f"{_unit_context(meta)}\n"
        f"The unit's sections are: {outline_titles}.\n\n"
        "Write exactly 5 fill-in-the-blank items: each 'q' is a sentence "
        "with one blank written as ______, and 'answer' is the missing "
        "word or phrase."), source)


def terminal_short(meta, outline_titles, source=None):
    return _with_source((
        f"{_unit_context(meta)}\n"
        f"The unit's sections are: {outline_titles}.\n\n"
        "Write exactly 5 short terminal questions (each answerable in 3-5 "
        "sentences, computational or definitional), each with its model "
        "'answer'."), source)


def terminal_long_questions(meta, outline_titles, source=None):
    return _with_source((
        f"{_unit_context(meta)}\n"
        f"The unit's sections are: {outline_titles}.\n\n"
        "Write exactly 5 long terminal QUESTIONS for this unit (essay/"
        "derivation depth, spanning the whole unit). Questions only — no "
        "answers."), source)


def terminal_long_answer(meta, question, source=None):
    return _with_source((
        f"{_unit_context(meta)}\n\n"
        f"Long terminal question: {question}\n\n"
        "Write the model answer: a thorough, well-structured essay of 2-4 "
        "paragraphs a strong postgraduate student would submit."), source)


def references(meta, outline_titles, source=None):
    return _with_source((
        f"{_unit_context(meta)}\n"
        f"The unit's sections are: {outline_titles}.\n\n"
        "List 6-8 real, well-known textbooks/resources for this subject in "
        "APA style (author, year, title, edition, publisher). Only books "
        "you are certain actually exist — standard, widely-cited texts."), source)
