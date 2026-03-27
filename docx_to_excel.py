"""
docx_to_excel.py

Parses a Word document structured as:
  - A table where one column contains section labels: Progress, Plans, Problems
  - Division headers (plain, non-bold paragraphs above bullet groups)
  - Bullet points where the Initiative is bold and the update follows in the same paragraph
  - Nested sub-bullets are appended as notes to the parent bullet

Output: Formatted Excel with columns:
  Section | Division | Initiative | Progress Update

Requirements (install once):
    pip install python-docx openpyxl

Usage:
    python docx_to_excel.py input.docx output.xlsx
"""

import sys
from itertools import groupby
from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

SECTION_NAMES = {"progress", "plans", "problems"}


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


def extract_bold_and_rest(para):
    bold_parts, rest_parts, switched = [], [], False
    for run in para.runs:
        if not switched and run.bold and run.text.strip():
            bold_parts.append(run.text)
        else:
            switched = True
            rest_parts.append(run.text)
    bold = "".join(bold_parts).strip().rstrip(":")
    rest = "".join(rest_parts).strip().lstrip(": ")
    return bold, rest


def is_division_header(para):
    text = para.text.strip()
    if not text or len(text) > 120:
        return False
    if is_list_paragraph(para):
        return False
    if text.lower() in SECTION_NAMES:
        return False
    style = para.style.name.lower()
    if any(x in style for x in ["heading", "title"]):
        return True
    bold_len = sum(len(r.text) for r in para.runs if r.bold and r.text.strip())
    return bold_len / max(len(text), 1) < 0.5


def parse_document(path):
    doc = Document(path)
    rows = []
    current_section = "Unknown"
    current_division = "Unknown"
    last_top_level = None

    for child in doc.element.body:
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag

        if tag == "tbl":
            table = Table(child, doc)
            for row in table.rows:
                for cell in row.cells:
                    if cell.text.strip().lower() in SECTION_NAMES:
                        current_section = cell.text.strip().capitalize()

        elif tag == "p":
            para = Paragraph(child, doc)
            text = para.text.strip()
            if not text:
                continue

            if text.lower() in SECTION_NAMES:
                current_section = text.capitalize()
                continue

            if is_division_header(para):
                current_division = text
                last_top_level = None
                continue

            if is_list_paragraph(para):
                indent = get_indent_level(para)
                bold, rest = extract_bold_and_rest(para)

                if indent == 0:
                    entry = {
                        "section": current_section,
                        "division": current_division,
                        "initiative": bold if bold else text,
                        "update": rest if bold else "",
                        "notes": [],
                    }
                    rows.append(entry)
                    last_top_level = entry
                else:
                    if last_top_level is not None:
                        note = (bold + (": " + rest if rest else "")) if bold else text
                        last_top_level["notes"].append(note)

    result = []
    for e in rows:
        update = e["update"]
        if e["notes"]:
            notes_str = " | ".join(e["notes"])
            update = (update + "\n[Notes: " + notes_str + "]") if update else "[Notes: " + notes_str + "]"
        result.append({
            "section": e["section"],
            "division": e["division"],
            "initiative": e["initiative"],
            "update": update,
        })
    return result


def thin_border():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)


def apply_header_style(cell, bg_hex):
    cell.font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    cell.fill = PatternFill("solid", start_color=bg_hex)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = thin_border()


def apply_section_style(cell, bg_hex):
    cell.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    cell.fill = PatternFill("solid", start_color=bg_hex)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = thin_border()


def write_excel(data, out_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Tracker"

    section_colors = {
        "Progress": "1F4E79",
        "Plans":    "375623",
        "Problems": "833C00",
    }

    even_fill = PatternFill("solid", start_color="EBF3FB")
    odd_fill  = PatternFill("solid", start_color="FFFFFF")

    for col, (h, c) in enumerate(
        zip(["Section", "Division", "Initiative", "Progress Update"],
            ["1F4E79", "1F4E79", "2E75B6", "2E75B6"]), start=1
    ):
        apply_header_style(ws.cell(row=1, column=col, value=h), c)
    ws.row_dimensions[1].height = 28

    section_order = ["Progress", "Plans", "Problems"]
    data_sorted = sorted(data, key=lambda x: section_order.index(x["section"])
                         if x["section"] in section_order else 99)

    current_row = 2
    for section_name, sec_group in groupby(data_sorted, key=lambda x: x["section"]):
        entries = list(sec_group)
        sec_color = section_colors.get(section_name, "404040")
        section_start = current_row
        div_idx = 0

        for division_name, div_group in groupby(entries, key=lambda x: x["division"]):
            div_entries = list(div_group)
            div_start = current_row

            for entry in div_entries:
                fill = even_fill if div_idx % 2 == 0 else odd_fill

                ws.cell(row=current_row, column=1, value=section_name)
                apply_section_style(ws.cell(row=current_row, column=1), sec_color)

                div_cell = ws.cell(row=current_row, column=2, value=division_name)
                div_cell.font = Font(name="Arial", bold=True, size=10)
                div_cell.fill = PatternFill("solid", start_color="D9E1F2")
                div_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                div_cell.border = thin_border()

                init_cell = ws.cell(row=current_row, column=3, value=entry["initiative"])
                init_cell.font = Font(name="Arial", bold=True, size=10)
                init_cell.fill = fill
                init_cell.alignment = Alignment(vertical="top", wrap_text=True)
                init_cell.border = thin_border()

                upd_cell = ws.cell(row=current_row, column=4, value=entry["update"])
                upd_cell.font = Font(name="Arial", size=10)
                upd_cell.fill = fill
                upd_cell.alignment = Alignment(vertical="top", wrap_text=True)
                upd_cell.border = thin_border()

                ws.row_dimensions[current_row].height = 50
                current_row += 1

            if current_row - 1 >= div_start:
                ws.merge_cells(start_row=div_start, start_column=2,
                               end_row=current_row - 1, end_column=2)
                dc = ws.cell(row=div_start, column=2)
                dc.font = Font(name="Arial", bold=True, size=10)
                dc.fill = PatternFill("solid", start_color="D9E1F2")
                dc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                dc.border = thin_border()

            div_idx += 1

        if current_row - 1 >= section_start:
            ws.merge_cells(start_row=section_start, start_column=1,
                           end_row=current_row - 1, end_column=1)
            apply_section_style(ws.cell(row=section_start, column=1), sec_color)

    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 32
    ws.column_dimensions["D"].width = 65
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = "A1:D1"

    wb.save(out_path)
    print(f"✓ Saved: {out_path}")


def main():
    if len(sys.argv) < 3:
        print("Usage: python docx_to_excel.py input.docx output.xlsx")
        sys.exit(1)

    data = parse_document(sys.argv[1])
    if not data:
        print("No entries found. Verify your document structure.")
        sys.exit(1)

    counts = {}
    for e in data:
        counts[e["section"]] = counts.get(e["section"], 0) + 1
    for sec, n in counts.items():
        print(f"  • {sec}: {n} entries")

    write_excel(data, sys.argv[2])


if __name__ == "__main__":
    main()
