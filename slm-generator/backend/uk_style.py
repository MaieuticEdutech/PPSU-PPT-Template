"""uk_style.py — conservative US->UK spelling normalisation for generated
content (house style: UK English, "enforce in prompts AND validate in code";
Phase 2's live smoke test showed the model writing "Analyze" despite a
UK-English system prompt, so prompting alone demonstrably isn't enough).

Deliberately conservative: only unambiguous mappings are applied (analyze->
analyse, color->colour...), matched on word boundaries, case-preserving for
the leading letter. Context-dependent pairs (program/programme, license/
licence, practice/practise) are NOT touched — a wrong "fix" there changes
meaning ("computer program" is correct UK English). `code` blocks are never
modified at all: identifiers like colorsys or df.normalize() must survive
verbatim (the FORMULA RULE in prompts/global_rules.txt).
"""
import re

# unambiguous stems: every inflection is generated from these
_STEMS = {
    "analyz": "analys",        # analyze/analyzed/analyzing/analyzer
    "organiz": "organis",
    "summariz": "summaris",
    "categoriz": "categoris",
    "generaliz": "generalis",
    "normaliz": "normalis",
    "optimiz": "optimis",
    "recogniz": "recognis",
    "visualiz": "visualis",
    "emphasiz": "emphasis",
    "prioritiz": "prioritis",
    "standardiz": "standardis",
    "utiliz": "utilis",
    "color": "colour",
    "behavior": "behaviour",
    "flavor": "flavour",
    "labor": "labour",
    "neighbor": "neighbour",
    "center": "centre",
    "modeling": "modelling",
    "labeling": "labelling",
    "labeled": "labelled",
    "catalog ": "catalogue ",   # trailing space: don't touch cataloging etc.
    "defense": "defence",
    "fulfill": "fulfil",
}

_PATTERNS = [(re.compile(r"\b" + us, re.IGNORECASE), uk)
             for us, uk in _STEMS.items()]


def _fix_word(match, uk):
    src = match.group(0)
    return uk.capitalize() if src[0].isupper() else uk


def apply_uk_spelling(text):
    """(fixed_text, n_fixes) for one string."""
    n = 0
    for pat, uk in _PATTERNS:
        text, k = pat.subn(lambda m, uk=uk: _fix_word(m, uk), text)
        n += k
    return text, n


def apply_to_unit(unit):
    """Walk every text field of an assembled unit dict IN PLACE, skipping
    code blocks entirely. Returns the total number of fixes made (reported
    so reviewers can see how much the model drifted)."""
    total = 0

    def fix(value):
        nonlocal total
        fixed, n = apply_uk_spelling(value)
        total += n
        return fixed

    def walk(node, in_code=False):
        if isinstance(node, dict):
            is_code = node.get("type") == "code"
            return {k: (v if (is_code and k == "text") else walk(v, in_code or is_code))
                    for k, v in node.items()}
        if isinstance(node, list):
            return [walk(v, in_code) for v in node]
        if isinstance(node, str) and not in_code:
            return fix(node)
        return node

    fixed_unit = walk(unit)
    unit.clear()
    unit.update(fixed_unit)
    return total
