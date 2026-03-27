"""
docx_to_excel.py

Handles Word documents where the entire content is inside a table:
  - Column 0: Section label (Progress / Plans / Problems) — may span multiple rows
  - Column 1+: Division header followed by bullet paragraphs within the cell

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
    """Short, non-list paragraph that isn't a section name."""
    text = para.text.strip()
    if not text or len(text) > 120:
        return False
    if text.lower() in SECTION_NAMES:
        return False
    if is_list_paragraph(para):
        return False
    return True


def parse_cell_paragraphs(cell_paragraphs, section, results):
    """
    Parse paragraphs from a content cell.
    First non-empty non-section paragraph is treated as the division header.
    Subsequent list paragraphs are initiatives/updates.
    Nested bullets become notes on the parent.
    """
    current_division = "Unknown"
    last_top_level = None
    found_division = False

    for para in cell_paragraphs:
        text = para.text.strip()
        if not text:
            continue

        if text.lower() in SECTION_NAMES:
            continue

        # First real paragraph in cell = division header
        if not found_division and not is_list_paragraph(para):
            current_division = text
            found_division = True
            continue

        # Some docs put division header as a non-list paragraph mid-cell
        if is_division_header(para) and not is_list_paragraph(para):
            current_division = text
            last_top_level = None
            continue

        if is_list_paragraph(para):
            indent = get_indent_level(para)
            bold, rest = extract_bold_and_rest(para)

            if indent == 0:
                entry = {
                    "section": section,
                    "division": current_division,
                    "initiative": bold if bold else text,
                    "update": rest if bold else "",
                    "notes": [],
                }
                results.append(entry)
                last_top_level = entry
            else:
                if last_top_level is not None:
                    note = (bold + (": " + rest if rest else "")) if bold else text
                    last_top_level["notes"].append(note)
        else:
            # Non-list paragraph after division header = treat as plain text entry
            if found_division:
                bold, rest = extract_bold_and_rest(para)
                entry = {
                    "section": section,
                    "division": current_division,
                    "initiative": bold if bold else text,
                    "update": rest if bold else "",
                    "notes": [],
                }
                results.append(entry)
                last_top_level = entry


def parse_document(path):
    doc = Document(path)
    raw_results = []
    current_section = "Unknown"

    for table in doc.tables:
        for row in table.rows:
            cells = row.cells

            # Check if any cell in this row is a section label
            for cell in cells:
                cell_text = cell.text.strip()
                if cell_text.lower() in SECTION_NAMES:
                    current_section = cell_text.capitalize()
                    break

            # Parse content cells (skip cells that are just section labels)
            for cell in cells:
                cell_text = cell.text.strip()
                if cell_text.lower() in SECTION_NAMES:
                    continue
                if not cell_text:
                    continue
                parse_cell_paragraphs(cell.paragraphs, current_section, raw_results)

    # Flatten notes
    result = []
    for e in raw_results:
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


# ── Excel output ──────────────────────────────────────────────────────────────

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
    data_sorted = sorted(
        data,
        key=lambda x: section_order.index(x["section"]) if x["section"] in section_order else 99
    )

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

                dc = ws.cell(row=current_row, column=2, value=division_name)
                dc.font = Font(name="Arial", bold=True, size=10)
                dc.fill = PatternFill("solid", start_color="D9E1F2")
                dc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                dc.border = thin_border()

                ic = ws.cell(row=current_row, column=3, value=entry["initiative"])
                ic.font = Font(name="Arial", bold=True, size=10)
                ic.fill = fill
                ic.alignment = Alignment(vertical="top", wrap_text=True)
                ic.border = thin_border()

                uc = ws.cell(row=current_row, column=4, value=entry["update"])
                uc.font = Font(name="Arial", size=10)
                uc.fill = fill
                uc.alignment = Alignment(vertical="top", wrap_text=True)
                uc.border = thin_border()

                ws.row_dimensions[current_row].height = 50
                current_row += 1

            # Merge division cells
            if current_row - 1 >= div_start:
                ws.merge_cells(start_row=div_start, start_column=2,
                               end_row=current_row - 1, end_column=2)
                dc = ws.cell(row=div_start, column=2)
                dc.font = Font(name="Arial", bold=True, size=10)
                dc.fill = PatternFill("solid", start_color="D9E1F2")
                dc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                dc.border = thin_border()

            div_idx += 1

        # Merge section cells
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
        print("No entries found. Run diagnose.py and share the output.")
        sys.exit(1)

    counts = {}
    for e in data:
        counts[e["section"]] = counts.get(e["section"], 0) + 1
    for sec, n in counts.items():
        print(f"  • {sec}: {n} entries")

    write_excel(data, sys.argv[2])


if __name__ == "__main__":
    main()
