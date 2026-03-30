"""
docx_to_excel.py

Parses a Word document (table-based structure) into a flat Excel table
suitable for Power BI and Power Automate.

Every row contains: Section | Division | Initiative | Progress Update
No merged cells. One record per row.

Requirements (install once):
    pip install python-docx openpyxl

Usage:
    python docx_to_excel.py input.docx output.xlsx
"""

import sys
from docx import Document
from docx.oxml.ns import qn
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SECTION_NAMES = {"progress", "plans", "problems"}


# ── Parsing helpers ───────────────────────────────────────────────────────────

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
    if text.lower() in SECTION_NAMES:
        return False
    if is_list_paragraph(para):
        return False
    return True


def parse_cell_paragraphs(cell_paragraphs, section, results):
    current_division = "Unknown"
    last_top_level = None
    found_division = False

    for para in cell_paragraphs:
        text = para.text.strip()
        if not text:
            continue
        if text.lower() in SECTION_NAMES:
            continue

        if not found_division and not is_list_paragraph(para):
            current_division = text
            found_division = True
            continue

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
            for cell in cells:
                if cell.text.strip().lower() in SECTION_NAMES:
                    current_section = cell.text.strip().capitalize()
                    break
            for cell in cells:
                cell_text = cell.text.strip()
                if cell_text.lower() in SECTION_NAMES or not cell_text:
                    continue
                parse_cell_paragraphs(cell.paragraphs, current_section, raw_results)

    result = []
    for e in raw_results:
        update = e["update"]
        if e["notes"]:
            notes_str = " | ".join(e["notes"])
            update = (update + " [Notes: " + notes_str + "]") if update else "[Notes: " + notes_str + "]"
        result.append({
            "section":    e["section"],
            "division":   e["division"],
            "initiative": e["initiative"],
            "update":     update,
        })
    return result


# ── Excel output ──────────────────────────────────────────────────────────────

def header_border():
    s = Side(style="medium", color="1F4E79")
    return Border(left=s, right=s, top=s, bottom=s)


def cell_border():
    s = Side(style="thin", color="BFBFBF")
    return Border(left=s, right=s, top=s, bottom=s)


def write_excel(data, out_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Tracker"

    # ── Header row ────────────────────────────────────────────────────────────
    headers = ["Section", "Division", "Initiative", "Progress Update"]
    header_fill = PatternFill("solid", start_color="1F4E79")

    for col, h in enumerate(headers, start=1):
        c = ws.cell(row=1, column=col, value=h)
        c.font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
        c.fill = header_fill
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = header_border()
    ws.row_dimensions[1].height = 24

    # ── Section color map (text only — no background fill on data rows) ───────
    section_colors = {
        "Progress": "1F4E79",
        "Plans":    "375623",
        "Problems": "833C00",
    }

    # Alternating row fills (subtle, won't interfere with PBI)
    even_fill = PatternFill("solid", start_color="EEF3F9")
    odd_fill  = PatternFill("solid", start_color="FFFFFF")

    # ── Data rows — one record per row, every cell populated ─────────────────
    section_order = ["Progress", "Plans", "Problems"]
    data_sorted = sorted(
        data,
        key=lambda x: section_order.index(x["section"]) if x["section"] in section_order else 99
    )

    for row_idx, entry in enumerate(data_sorted, start=2):
        fill = even_fill if row_idx % 2 == 0 else odd_fill
        sec_color = section_colors.get(entry["section"], "404040")

        values = [entry["section"], entry["division"], entry["initiative"], entry["update"]]
        for col, val in enumerate(values, start=1):
            c = ws.cell(row=row_idx, column=col, value=val)
            c.fill = fill
            c.border = cell_border()
            c.alignment = Alignment(vertical="top", wrap_text=True)

            if col == 1:
                # Section cell: colored bold text
                c.font = Font(name="Arial", bold=True, size=10, color=sec_color)
                c.alignment = Alignment(horizontal="center", vertical="top")
            elif col == 2:
                # Division: bold
                c.font = Font(name="Arial", bold=True, size=10)
            elif col == 3:
                # Initiative: bold
                c.font = Font(name="Arial", bold=True, size=10)
            else:
                # Progress Update: normal
                c.font = Font(name="Arial", size=10)

        ws.row_dimensions[row_idx].height = 40

    # ── Column widths ─────────────────────────────────────────────────────────
    ws.column_dimensions["A"].width = 14   # Section
    ws.column_dimensions["B"].width = 26   # Division
    ws.column_dimensions["C"].width = 32   # Initiative
    ws.column_dimensions["D"].width = 70   # Progress Update

    # ── Freeze header + auto-filter ───────────────────────────────────────────
    ws.freeze_panes = "A2"
    last_col = get_column_letter(len(headers))
    ws.auto_filter.ref = f"A1:{last_col}1"

    wb.save(out_path)
    print(f"✓ Saved: {out_path}  ({len(data_sorted)} rows)")


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
