"""
diagnose.py
Prints the raw structure of a Word document so we can see exactly
what python-docx is reading — styles, bold runs, tables, indent levels.

Usage:
    python diagnose.py input.docx
"""

import sys
from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph


def get_indent_level(para):
    pPr = para._p.find(qn("w:pPr"))
    if pPr is None:
        return 0
    numPr = pPr.find(qn("w:numPr"))
    if numPr is None:
        return 0
    ilvl = numPr.find(qn("w:ilvl"))
    return int(ilvl.get(qn("w:val"), 0)) if ilvl is not None else 0


def is_list_paragraph(para):
    if "list" in para.style.name.lower() or "bullet" in para.style.name.lower():
        return True
    pPr = para._p.find(qn("w:pPr"))
    return pPr is not None and pPr.find(qn("w:numPr")) is not None


def main():
    if len(sys.argv) < 2:
        print("Usage: python diagnose.py input.docx")
        sys.exit(1)

    doc = Document(sys.argv[1])
    print("=" * 70)
    print("DOCUMENT STRUCTURE DIAGNOSIS")
    print("=" * 70)

    element_count = 0
    for child in doc.element.body:
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag

        # ── Table ────────────────────────────────────────────────────────────
        if tag == "tbl":
            element_count += 1
            print(f"\n[ELEMENT {element_count}] TABLE")
            table = Table(child, doc)
            for r_idx, row in enumerate(table.rows):
                row_texts = [cell.text.strip() for cell in row.cells]
                print(f"  Row {r_idx}: {row_texts}")

        # ── Paragraph ────────────────────────────────────────────────────────
        elif tag == "p":
            para = Paragraph(child, doc)
            text = para.text.strip()
            if not text:
                continue

            element_count += 1
            is_list = is_list_paragraph(para)
            indent = get_indent_level(para)
            style = para.style.name

            # Bold run breakdown
            run_info = []
            for run in para.runs:
                if run.text.strip():
                    run_info.append(f"{'[BOLD]' if run.bold else '[norm]'} {repr(run.text[:40])}")

            print(f"\n[ELEMENT {element_count}] PARAGRAPH")
            print(f"  Style : {style}")
            print(f"  List  : {is_list}  |  Indent level: {indent}")
            print(f"  Text  : {repr(text[:80])}")
            if run_info:
                print(f"  Runs  :")
                for r in run_info:
                    print(f"    {r}")

    print("\n" + "=" * 70)
    print(f"Total elements processed: {element_count}")
    print("=" * 70)


if __name__ == "__main__":
    main()
