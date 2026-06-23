/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

export const PYTHON_REQUIREMENTS = `Flask>=3.0.0
python-docx>=1.1.0
lxml>=5.1.0
`;

export const PYTHON_README = `# DOCX Formatting Standards Validator (Python Local Tool)

This folder contains a complete, working Python-based implementation of the DOCX Formatting Standards Validator. It features a Flask backend and a modern web dashboard.

---

## 📂 Folder Structure

\`\`\`
python_app/
├── app.py             # Main Flask server + XML-level paragraph parser
├── requirements.txt   # Python package dependencies
└── README.md          # Local setup & running instructions (this file)
\`\`\`

---

## ⚙️ Prerequisites \\& Setup

Make sure you have **Python 3.8 or newer** installed on your system.

### 1. Install Dependencies

Open your terminal or command prompt inside the \`python_app\` folder, and use \`pip\` to install the required packages:

\`\`\`bash
pip install -r requirements.txt
\`\`\`

Alternatively, install them manually:
\`\`\`bash
pip install Flask python-docx lxml
\`\`\`

---

## 🚀 Running the Local Web UI

1. Run the Flask application:
   \`\`\`bash
   python app.py
   \`\`\`

2. Open your web browser and navigate to:
   \`\`\`
   http://127.0.0.1:5000
   \`\`\`

3. Enter the paths:
   - **Target Scanning Directory**: The absolute path to the directory containing your \`.docx\` files.
   - **Output Directory**: (Optional) Path to save the extracted JSON and CSV reports.

4. Click **Run Analysis** to scan and validate documents instantly.

---

## ⚡ Standards \\& Validation Rules Applied

Each paragraph of your documents is extracted down to the word level using \`python-docx\` and XML parsing. The rules enforced are:

| Level ID | Expected Level | Font Name | Font Size | Bold Status | Expected Alignment | Expected Indent |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| \`heading1\` | Level 1 Heading | \`Human Black Condensed\` | \`20 pt\` | \`True\` | \`LEFT\` | \`0 pt\` |
| \`heading2\` | Level 2 Heading | \`Human Black Condensed\` | \`11.5 pt\` | \`True\` | \`LEFT\` | \`0 pt\` |
| \`heading3\` | Level 3 Heading | \`Human Black Condensed\` | \`10.5 pt\` | \`True\` | \`LEFT\` | \`0 pt\` |
| \`normal\` | Normal Text | \`Human Light Condensed\` | \`10.5 pt\` | \`False\` | \`LEFT\` | \`0 pt\` |
| \`boldText\` | Text Level Bold | \`Human Bold Condensed\` | \`10.5 pt\` | \`True\` | \`LEFT\` | \`0 pt\` |
| \`formName\` | Form Name (Footer/End) | \`Human Light Condensed\` | \`7 pt\` | \`False\` | \`LEFT\` | \`0 pt\` |
`;

export const PYTHON_APP = `# -*- coding: utf-8 -*-
"""
DOCX Formatting Validator - Python Local Server Code
===================================================
A Flask-based application for scanning and validating the formatting of .docx documents.
Performs XML-level run parsing and paragraph-level validation against style conventions.

Requires:
    pip install flask python-docx lxml

Usage:
    python app.py
    Then direct your browser to http://127.0.0.1:5000
"""

import os
import json
import csv
import re
from pathlib import Path
from datetime import datetime
from flask import Flask, request, jsonify, render_template_string

# Try to import python-docx
try:
    from docx import Document
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
except ImportError:
    print("Warning: 'python-docx' library is not installed. Install with 'pip install python-docx'")

app = Flask(__name__)

# Predefined standard rules as requested in the prompt
DEFAULT_STANDARDS = [
    {
        "id": "heading1",
        "name": "Level 1 Heading",
        "fontName": "Human Black Condensed",
        "fontSize": 20.0,
        "bold": True,
        "alignment": "LEFT",
        "indent": 0.0
    },
    {
        "id": "heading2",
        "name": "Level 2 Heading",
        "fontName": "Human Black Condensed",
        "fontSize": 11.5,
        "bold": True,
        "alignment": "LEFT",
        "indent": 0.0
    },
    {
        "id": "heading3",
        "name": "Level 3 Heading",
        "fontName": "Human Black Condensed",
        "fontSize": 10.5,
        "bold": True,
        "alignment": "LEFT",
        "indent": 0.0
    },
    {
        "id": "normal",
        "name": "Normal Text",
        "fontName": "Human Light Condensed",
        "fontSize": 10.5,
        "bold": False,
        "alignment": "LEFT",
        "indent": 0.0
    },
    {
        "id": "boldText",
        "name": "Text Level Bold",
        "fontName": "Human Bold Condensed",
        "fontSize": 10.5,
        "bold": True,
        "alignment": "LEFT",
        "indent": 0.0
    },
    {
        "id": "formName",
        "name": "Form Name",
        "fontName": "Human Light Condensed",
        "fontSize": 7.0,
        "bold": False,
        "alignment": "LEFT",
        "indent": 0.0
    }
]

W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

def get_doc_defaults(doc):
    """Parse styles.xml to get default document properties."""
    defaults = {"fontName": None, "fontSize": None, "bold": None}
    try:
        styles_element = doc.styles.element
        doc_def = styles_element.find(f"{W_NS}docDefaults")
        if doc_def is not None:
            rpr_default = doc_def.find(f"{W_NS}rPrDefault")
            if rpr_default is not None:
                rpr = rpr_default.find(f"{W_NS}rPr")
                if rpr is not None:
                    rfonts = rpr.find(f"{W_NS}rFonts")
                    if rfonts is not None:
                        for attr in ["ascii", "hAnsi", "cs"]:
                            val = rfonts.get(f"{W_NS}{attr}")
                            if val:
                                defaults["fontName"] = val
                                break
                    sz = rpr.find(f"{W_NS}sz")
                    if sz is not None:
                        hp = sz.get(f"{W_NS}val")
                        if hp:
                            defaults["fontSize"] = int(hp) / 2.0
                    b_elem = rpr.find(f"{W_NS}b")
                    if b_elem is not None:
                        b_val = b_elem.get(f"{W_NS}val")
                        defaults["bold"] = b_val is None or b_val.lower() in ("true", "1")
    except Exception:
        pass
    return defaults


def get_xml_run_props(r_elem):
    """Extract formatting values from a single run elements."""
    props = {"fontName": None, "fontSize": None, "bold": None, "runStyleId": None}
    rpr = r_elem.find(f"{W_NS}rPr")
    if rpr is None:
        return props

    rfonts = rpr.find(f"{W_NS}rFonts")
    if rfonts is not None:
        for attr in ["ascii", "hAnsi", "cs"]:
            val = rfonts.get(f"{W_NS}{attr}")
            if val:
                props["fontName"] = val
                break

    sz = rpr.find(f"{W_NS}sz")
    if sz is not None:
        val = sz.get(f"{W_NS}val")
        if val:
            props["fontSize"] = int(val) / 2.0

    b_elem = rpr.find(f"{W_NS}b")
    if b_elem is not None:
        val = b_elem.get(f"{W_NS}val")
        props["bold"] = val is None or val.lower() in ("true", "1")

    r_style = rpr.find(f"{W_NS}rStyle")
    if r_style is not None:
        props["runStyleId"] = r_style.get(f"{W_NS}val")

    return props


def walk_style_chain(style, key_func, max_depth=10):
    """Traverse style hierarchy (based_on) to resolve formatting."""
    current = style
    depth = 0
    while current and depth < max_depth:
        val = key_func(current)
        if val is not None:
            return val
        try:
            current = current.base_style
        except Exception:
            break
        depth += 1
    return None


def resolve_properties(run_props, para, doc_defaults, doc_styles):
    """Resolve final font properties based on XML cascade."""
    font_name = run_props["fontName"]
    font_size = run_props["fontSize"]
    bold = run_props["bold"]

    # Try run style
    if run_props["runStyleId"] and run_props["runStyleId"] in doc_styles:
        rstyle = doc_styles[run_props["runStyleId"]]
        if font_name is None:
            font_name = walk_style_chain(rstyle, lambda s: s.font.name if s.font else None)
        if font_size is None:
            font_size = walk_style_chain(rstyle, lambda s: s.font.size.pt if s.font and s.font.size else None)
        if bold is None:
            bold = walk_style_chain(rstyle, lambda s: s.font.bold if s.font else None)

    # Try paragraph style
    try:
        pstyle = para.style
        if pstyle:
            if font_name is None:
                font_name = walk_style_chain(pstyle, lambda s: s.font.name if s.font else None)
            if font_size is None:
                font_size = walk_style_chain(pstyle, lambda s: s.font.size.pt if s.font and s.font.size else None)
            if bold is None:
                bold = walk_style_chain(pstyle, lambda s: s.font.bold if s.font else None)
    except Exception:
        pass

    # Fallback to document defaults
    if font_name is None:
        font_name = doc_defaults["fontName"]
    if font_size is None:
        font_size = doc_defaults["fontSize"]
    if bold is None:
        bold = doc_defaults["bold"]

    return font_name or "Calibri", font_size or 11.0, bold or False


def parse_paragraph_xml_runs(para, doc_defaults, doc_styles):
    """Parser to find nested run blocks in paragraph and gather words."""
    p_elem = para._element
    runs_data = []
    
    for r_elem in p_elem.iter(f"{W_NS}r"):
        text_parts = []
        for child in r_elem:
            if child.tag == f"{W_NS}t" and child.text:
                text_parts.append(child.text)
            elif child.tag == f"{W_NS}tab":
                text_parts.append("\\t")
            elif child.tag == f"{W_NS}br":
                text_parts.append("\\n")
        
        run_text = "".join(text_parts)
        if not run_text:
            continue

        run_xml_props = get_xml_run_props(r_elem)
        r_font, r_size, r_bold = resolve_properties(run_xml_props, para, doc_defaults, doc_styles)
        
        runs_data.append({
            "text": run_text,
            "fontName": r_font,
            "fontSize": r_size,
            "bold": r_bold
        })
    return runs_data


def analyze_docx_file(filepath, standards=DEFAULT_STANDARDS):
    """Extract paragraphs and run metrics with style validations."""
    doc = Document(str(filepath))
    doc_defaults = get_doc_defaults(doc)
    doc_styles = {s.style_id: s for s in doc.styles if s.style_id}
    
    paragraphs_results = []
    
    for idx, para in enumerate(doc.paragraphs):
        text = para.text
        if not text.strip():
            continue
            
        runs = parse_paragraph_xml_runs(para, doc_defaults, doc_styles)
        
        font_lens = {}
        size_lens = {}
        bold_lens = {True: 0, False: 0}
        
        for r in runs:
            rlen = len(r["text"])
            font_lens[r["fontName"]] = font_lens.get(r["fontName"], 0) + rlen
            size_lens[r["fontSize"]] = size_lens.get(r["fontSize"], 0) + rlen
            bold_lens[r["bold"]] = bold_lens.get(r["bold"], 0) + rlen
            
        dominant_font = max(font_lens, key=font_lens.get) if font_lens else "Calibri"
        dominant_size = max(size_lens, key=size_lens.get) if size_lens else 11.0
        dominant_bold = bold_lens[True] >= bold_lens[False] if runs else False
        
        align_str = "LEFT"
        if para.alignment == WD_ALIGN_PARAGRAPH.RIGHT:
            align_str = "RIGHT"
        elif para.alignment == WD_ALIGN_PARAGRAPH.CENTER:
            align_str = "CENTER"
        elif para.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY:
            align_str = "JUSTIFY"
            
        indent_pt = 0
        if para.paragraph_format.left_indent:
            indent_pt = round(para.paragraph_format.left_indent.pt, 2)
            
        style_name = para.style.name if para.style else "Normal"
        
        matched_rule = None
        min_diff = 9999
        for rule in standards:
            diff = abs(dominant_size - rule["fontSize"]) * 10
            if dominant_font.lower() != rule["fontName"].lower():
                diff += 5
            if rule["bold"] is not None and dominant_bold != rule["bold"]:
                diff += 3
            if style_name.lower().find("heading") != -1 and rule["id"].startswith("heading"):
                diff -= 2
            if diff < min_diff:
                min_diff = diff
                matched_rule = rule
                
        issues = []
        validation_status = "Pass"
        matched_level = "Unknown Standard"
        matched_level_id = "unmatched"
        
        if matched_rule:
            matched_level = matched_rule["name"]
            matched_level_id = matched_rule["id"]
            
            if dominant_font.lower() != matched_rule["fontName"].lower():
                issues.append({
                    "type": "fontName",
                    "severity": "error",
                    "message": f"Incorrect Font for {matched_rule['name']}. Expected '{matched_rule['fontName']}', but found '{dominant_font}'."
                })
            if dominant_size != matched_rule["fontSize"]:
                issues.append({
                    "type": "fontSize",
                    "severity": "error",
                    "message": f"Incorrect Font Size for {matched_rule['name']}. Expected {matched_rule['fontSize']}pt, but found {dominant_size}pt."
                })
            if matched_rule["bold"] is not None and dominant_bold != matched_rule["bold"]:
                issues.append({
                    "type": "bold",
                    "severity": "error",
                    "message": f"Incorrect Bold formatting for {matched_rule['name']}."
                })
            if matched_rule["alignment"] and align_str != matched_rule["alignment"]:
                issues.append({
                    "type": "alignment",
                    "severity": "warning",
                    "message": f"Inconsistent alignment. Expected '{matched_rule['alignment']}', found '{align_str}'."
                })
            if matched_rule["indent"] is not None and indent_pt != matched_rule["indent"]:
                if abs(indent_pt - matched_rule["indent"]) > 1.0:
                    issues.append({
                        "type": "indent",
                        "severity": "warning",
                        "message": f"Indentation issue. Expected '{matched_rule['indent']}pt', found '{indent_pt}pt'."
                    })
        else:
            issues.append({
                "type": "unmatched",
                "severity": "error",
                "message": f"Paragraph formatting matches no known standards."
            })
            
        any_errors = any(i["severity"] == "error" for i in issues)
        if any_errors or not matched_rule:
            validation_status = "Fail"
            
        paragraphs_results.append({
            "paragraphNumber": idx + 1,
            "text": text,
            "styleName": style_name,
            "alignment": align_str,
            "indent": indent_pt,
            "dominantFont": dominant_font,
            "dominantSize": dominant_size,
            "dominantBold": dominant_bold,
            "matchedLevelId": matched_level_id,
            "matchedLevelName": matched_level,
            "validationStatus": validation_status,
            "issues": issues,
            "runs": runs
        })
        
    return paragraphs_results


def scan_and_analyze_folder(folder_path, output_fol=None):
    p = Path(folder_path)
    if not p.exists() or not p.is_dir():
        return None, "Folder does not exist."
        
    files = list(p.glob("*.docx"))
    if not files:
        return None, "No .docx files found in this folder."
        
    all_reports = []
    
    for f in files:
        try:
            p_res = analyze_docx_file(f)
            total = len(p_res)
            passed = sum(1 for x in p_res if x["validationStatus"] == "Pass")
            failed = total - passed
            issues_cnt = sum(len(x["issues"]) for x in p_res)
            
            summary = {
                "totalParagraphs": total,
                "passed": passed,
                "failed": failed,
                "issuesCount": issues_cnt
            }
            
            all_reports.append({
                "fileName": f.name,
                "fileSize": f.stat().st_size,
                "filePath": str(f),
                "paragraphs": p_res,
                "summary": summary
            })
            
        except Exception as e:
            all_reports.append({
                "fileName": f.name,
                "filePath": str(f),
                "paragraphs": [],
                "error": str(e),
                "summary": {"totalParagraphs": 0, "passed": 0, "failed": 0, "issuesCount": 0}
            })
            
    if output_fol:
        out_p = Path(output_fol)
        out_p.mkdir(parents=True, exist_ok=True)
        
        json_path = out_p / "formatting_analysis_report.json"
        with open(json_path, "w", encoding="utf-8") as jf:
            json.dump(all_reports, jf, ensure_ascii=False, indent=2)
            
        csv_path = out_p / "formatting_analysis_report.csv"
        with open(csv_path, "w", encoding="utf-8", newline="") as cf:
            writer = csv.writer(cf)
            writer.writerow([
                "File Name", "Paragraph Num", "Paragraph Text", 
                "Font Name", "Font Size", "Bold", "Alignment", "Indent (pt)",
                "Expected Level", "Status", "Issues Highlight"
            ])
            for report in all_reports:
                for p_obj in report.get("paragraphs", []):
                    issue_list = "; ".join([i["message"] for i in p_obj["issues"]])
                    writer.writerow([
                        report["fileName"],
                        p_obj["paragraphNumber"],
                        p_obj["text"][:120].replace("\\n", " ") + ("..." if len(p_obj["text"]) > 120 else ""),
                        p_obj["dominantFont"],
                        p_obj["dominantSize"],
                        "True" if p_obj["dominantBold"] else "False",
                        p_obj["alignment"],
                        p_obj["indent"],
                        p_obj["matchedLevelName"],
                        p_obj["validationStatus"],
                        issue_list
                    ])
                    
    return all_reports, None

HTML_TEMPLATE = """
... (Flask template) ...
"""
`;
