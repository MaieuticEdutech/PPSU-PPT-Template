#!/usr/bin/env python3
"""Tests for the brand profiles: a REVA build must follow the REVA SLM
Style Guide explicitly, and a PPSU build must be unchanged by the feature.

Run: python tests/test_brands.py
"""
import io
import json
import re
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "backend"))

import brands
from docx import Document
from docx_builder import build

passed = failed = 0


def check(name, condition):
    global passed, failed
    if condition:
        passed += 1
        print(f"  OK   {name}")
    else:
        failed += 1
        print(f"  FAIL {name}")


GOLDEN = json.loads((Path(__file__).parent / "golden_unit1.json")
                    .read_text(encoding="utf-8"))


def doc_xml(document):
    buf = io.BytesIO()
    document.save(buf)
    with zipfile.ZipFile(io.BytesIO(buf.getvalue())) as z:
        return (z.read("word/document.xml").decode("utf-8"),
                z.read("word/header2.xml").decode("utf-8", "ignore")
                if "word/header2.xml" in z.namelist() else
                z.read([n for n in z.namelist()
                        if n.startswith("word/header")][0]).decode("utf-8"))


def all_text(document):
    from docx.oxml.ns import qn
    parts = []
    for t in document.element.body.iter(qn('w:t')):
        if t.text:
            parts.append(t.text)
    return "\n".join(parts)


print("=== REVA filename convention ===")
meta = {"course_name": "Computer Networks", "unit_number": 1,
        "unit_title": "Fundamentals of Computer Networks"}
check("CamelCase_UnitNN_Title",
      brands.reva_filename(meta, ".docx")
      == "ComputerNetworks_Unit01_FundamentalsOfComputerNetworks.docx")
check("special characters stripped",
      brands.reva_filename({"course_name": "Maths & Stats!", "unit_number": 12,
                             "unit_title": "Sets, Maps"}, ".pdf")
      == "MathsStats_Unit12_SetsMaps.pdf")

print("\n=== REVA build follows the style guide ===")
reva_doc = build(GOLDEN, brand="reva", use_branding=False)
xml, header_xml = doc_xml(reva_doc)
text = all_text(reva_doc)

check("Plus Jakarta Sans set on runs", 'w:ascii="Plus Jakarta Sans"' in xml)
check("REVA Orange heading bars (#F7A35B paragraph shading)",
      re.search(r'w:shd[^>]*w:fill="F7A35B"', xml) is not None)
check("subtopic bars Light Orange Tint (#FEF0E6)",
      re.search(r'w:shd[^>]*w:fill="FEF0E6"', xml) is not None)
check("headings uppercase (Topic Heading rule)",
      "1.1 INTRODUCTION; IMPORTANCE IN DATA SCIENCE" in text)
check("aside boxes renamed: Study Note replaces Did you know?",
      "Study Note" in text and "Did you know?" not in text)
check("Think and Apply becomes Activity",
      "Activity" in text and "Think and Apply" not in text)
check("Study Note fill #FFD966 present",
      re.search(r'w:shd[^>]*w:fill="FFD966"', xml) is not None)
check("Activity fill #A9D18E present",
      re.search(r'w:shd[^>]*w:fill="A9D18E"', xml) is not None)
check("table headers REVA Orange, borders #CCCCCC",
      re.search(r'w:shd[^>]*w:fill="F7A35B"', xml) is not None
      and 'w:color="CCCCCC"' in xml)
check("figure captions use 'Fig. N:' not 'Figure N:'",
      "Fig. 1: Discrete Mathematics" in text
      and "Figure 1: Discrete Mathematics" not in text)
check("references section renamed per REVA structure",
      "SUGGESTED BOOKS AND REFERENCES" in text.upper())
check("body text Dark Charcoal (#333333)", 'w:val="333333"' in xml)
check("REVA header: 'Unit Name | Unit Number' + orange rule",
      "Foundations of Discrete Mathematics | Unit 01" in header_xml
      and 'w:color="F7A35B"' in header_xml)

print("\n=== PPSU build unaffected by the feature ===")
ppsu_doc = build(GOLDEN, brand="ppsu", use_branding=False)
p_xml, _ = doc_xml(ppsu_doc)
p_text = all_text(ppsu_doc)
check("PPSU keeps Calibri", 'w:ascii="Calibri"' in p_xml)
check("PPSU keeps 'Did you know?' and 'Figure N:'",
      "Did you know?" in p_text and "Figure 1: Discrete Mathematics" in p_text)
check("PPSU has no orange heading bars",
      re.search(r'w:shd[^>]*w:fill="F7A35B"', p_xml) is None)
check("PPSU headings not uppercased",
      "1.1 Introduction; Importance in Data Science" in p_text)

print(f"\n{passed} passed, {failed} failed")
sys.exit(1 if failed else 0)
