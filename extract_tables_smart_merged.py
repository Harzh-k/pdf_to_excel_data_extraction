#!/usr/bin/env python3
"""
Universal IRDAI PDF → Excel Extractor
Each table gets its own sheet in the workbook.

Usage:
  python extract_tables_smart_merged.py <pdf_path> [output_path]
"""

import re
import sys
import os
from pathlib import Path

try:
    import pdfplumber
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError as e:
    print(f"Error: {e}\nInstall: pip install pdfplumber openpyxl")
    sys.exit(1)


# ============================================================
# STYLE CONSTANTS
# ============================================================

_thin  = Side(style="thin")
_thick = Side(style="medium")

BORDER_NONE         = Border()
BORDER_BOTTOM_THIN  = Border(bottom=_thin)
BORDER_BOTTOM_THICK = Border(bottom=_thick)
BORDER_H_THIN       = Border(top=_thin, bottom=_thin)

FILL_PAGE_HDR = PatternFill("solid", fgColor="1F3864")  # dark navy
FILL_FORM_HDR = PatternFill("solid", fgColor="C00000")  # IRDAI red
FILL_COL_HDR  = PatternFill("solid", fgColor="2E75B6")  # blue
FILL_SUB_HDR  = PatternFill("solid", fgColor="D6E4F0")  # light blue
FILL_TOTAL    = PatternFill("solid", fgColor="E2EFDA")  # light green
FILL_NONE     = PatternFill("none")

FONT_PAGE_HDR = Font(bold=True, color="FFFFFF", size=10)
FONT_WHITE    = Font(bold=True, color="FFFFFF", size=9)
FONT_BOLD     = Font(bold=True, size=9)
FONT_NORMAL   = Font(size=9)

ALIGN_TOP_WRAP = Alignment(vertical="top", wrap_text=True)
ALIGN_CENTER   = Alignment(vertical="center", indent=1)

_COL_HDR_KEYS = {
    "LIFE", "PENSION", "HEALTH", "ANNUITY", "PARTICULARS",
    "LINKED BUSINESS", "PARTICIPATING", "NON-PARTICIPATING",
    "VAR.INS", "VAR. INS"
}
_TOTAL_KEYS = {
    "TOTAL (A)", "TOTAL (B)", "TOTAL (C)", "TOTAL (D)",
    "SUB TOTAL", "TOTAL", "SURPLUS", "DEFICIT",
    "AMOUNT AVAILABLE", "APPROPRIATION"
}
_FORM_HDR_KEYS = {
    "FORM L-", "REVENUE ACCOUNT", "PROFIT AND LOSS",
    "BALANCE SHEET", "POLICYHOLDERS", "NAME OF THE INSURER",
    "REGISTRATION", "SCHEDULE"
}


# ============================================================
# ROW CLASSIFIER
# ============================================================

def _classify_row(row, row_idx):
    text = " ".join(str(c) for c in row if c).upper().strip()
    if row_idx < 7 and any(k in text for k in _FORM_HDR_KEYS):
        return "form_header"
    if any(k in text for k in _COL_HDR_KEYS):
        return "col_header"
    if any(k in text for k in _TOTAL_KEYS):
        return "total"
    return "normal"


# ============================================================
# SHEET NAME BUILDER
# ============================================================

def _make_sheet_name(used_names, form_name, page_number, table_index):
    """
    Build a unique, Excel-safe sheet name.
    Excel limit: 31 chars, no special chars [ ] : * ? / \
    """
    # Extract short form code e.g. 'L-1-A-RA' from 'FORM L-1-A-RA'
    if form_name:
        short = re.sub(r"^FORM\s+", "", form_name).strip()
    else:
        short = "Sheet"

    # Remove Excel-illegal chars
    short = re.sub(r"[\[\]:*?/\\]", "-", short)

    # Build candidate name
    candidate = f"{short}_P{page_number}"
    if table_index > 1:
        candidate += f"_T{table_index}"

    # Truncate to 31 chars
    candidate = candidate[:31]

    # Ensure uniqueness by appending suffix if needed
    base = candidate
    suffix = 2
    while candidate in used_names:
        candidate = f"{base[:28]}_{suffix}"
        suffix += 1

    used_names.add(candidate)
    return candidate


# ============================================================
# SHEET WRITER
# ============================================================

def _write_table_to_sheet(ws, table_data, page_number, form_name):
    """Write a single table to a worksheet with formatting."""

    cur_row = 1

    # Sheet header banner
    hdr_text = f"Page {page_number}"
    if form_name:
        hdr_text += f"   |   {form_name}"

    cell = ws.cell(row=cur_row, column=1, value=hdr_text)
    cell.font      = FONT_PAGE_HDR
    cell.fill      = FILL_PAGE_HDR
    cell.border    = Border(bottom=_thick)
    cell.alignment = ALIGN_CENTER
    cur_row += 1

    # Table rows
    for row_idx, row in enumerate(table_data):
        row_type = _classify_row(row, row_idx)

        for ci, val in enumerate(row, start=1):
            cell = ws.cell(row=cur_row, column=ci, value=val if val else None)
            cell.alignment = ALIGN_TOP_WRAP

            if row_type == "form_header":
                if row_idx == 0:
                    cell.fill   = FILL_FORM_HDR
                    cell.font   = FONT_WHITE
                    cell.border = BORDER_BOTTOM_THIN
                else:
                    cell.fill   = FILL_SUB_HDR
                    cell.font   = FONT_BOLD
                    cell.border = BORDER_BOTTOM_THIN

            elif row_type == "col_header":
                cell.fill   = FILL_COL_HDR
                cell.font   = FONT_WHITE
                cell.border = BORDER_BOTTOM_THICK

            elif row_type == "total":
                cell.fill   = FILL_TOTAL
                cell.font   = FONT_BOLD
                cell.border = BORDER_H_THIN

            else:
                cell.font   = FONT_NORMAL
                cell.fill   = FILL_NONE
                cell.border = BORDER_NONE

        cur_row += 1

    # Auto column widths
    col_widths = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value:
                first_line = str(cell.value).split("\n")[0]
                col_widths[cell.column] = max(
                    col_widths.get(cell.column, 0), len(first_line)
                )
    for col, width in col_widths.items():
        ws.column_dimensions[get_column_letter(col)].width = min(width + 2, 40)

    # Freeze top 2 rows (banner + first data row)
    ws.freeze_panes = "A3"


# ============================================================
# EXTRACTION ENGINE
# ============================================================

def extract_forms_from_pdf(pdf_path):
    from src.extractor import TABLE_SETTINGS, header_needs_rebuild, rebuild_using_header_spans

    document_pages = []
    current_form_name = None

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""

            m = re.search(r"FORM\s+L-[\dA-Za-z\-]+", text)
            if m:
                current_form_name = m.group().strip()

            page_block = {
                "page_number": page.page_number,
                "form_name": current_form_name,
                "content": []
            }

            found_tables = page.find_tables(table_settings=TABLE_SETTINGS)

            if found_tables:
                for idx, table in enumerate(found_tables):
                    table_data = table.extract()

                    if header_needs_rebuild(table_data):
                        rebuilt = rebuild_using_header_spans(page, table.bbox)
                        if rebuilt:
                            table_data = rebuilt

                    page_block["content"].append({
                        "type": "table",
                        "index_on_page": idx + 1,
                        "data": table_data
                    })

            if not page_block["content"]:
                page_block["content"].append({
                    "type": "text",
                    "data": text.split("\n")
                })

            document_pages.append(page_block)

    print(f"Processed {len(document_pages)} pages.")
    return document_pages


# ============================================================
# EXCEL CREATION — ONE SHEET PER TABLE
# ============================================================

def create_excel_from_forms(document_pages, output_path):
    print("Creating Excel workbook (one sheet per table)...")

    wb = Workbook()
    wb.remove(wb.active)

    used_sheet_names = set()
    total_sheets = 0

    for pb in document_pages:
        page_number = pb["page_number"]
        form_name   = pb["form_name"]
        contents    = pb["content"]

        # Count tables on this page for naming
        table_count = 0

        for item in contents:

            if item["type"] == "table":
                table_count += 1
                sheet_name = _make_sheet_name(
                    used_sheet_names, form_name, page_number, table_count
                )
                ws = wb.create_sheet(title=sheet_name)
                _write_table_to_sheet(ws, item["data"], page_number, form_name)
                total_sheets += 1
                print(f"  Sheet '{sheet_name}': {len(item['data'])} rows")

            elif item["type"] == "text":
                # Text-only pages (no tables found) → one sheet with raw text
                lines = [l for l in item["data"] if l.strip()]
                if not lines:
                    continue
                sheet_name = _make_sheet_name(
                    used_sheet_names, form_name, page_number, 1
                )
                ws = wb.create_sheet(title=sheet_name)

                # Simple banner
                cell = ws.cell(row=1, column=1, value=f"Page {page_number} — text only")
                cell.font   = FONT_PAGE_HDR
                cell.fill   = FILL_PAGE_HDR
                cell.border = Border(bottom=_thick)

                for i, line in enumerate(lines, start=2):
                    ws.cell(row=i, column=1, value=line).font = FONT_NORMAL

                ws.column_dimensions["A"].width = 80
                ws.freeze_panes = "A2"
                total_sheets += 1

    # Save
    try:
        wb.save(output_path)
        size_mb = os.path.getsize(output_path) / (1024 * 1024)
        print(f"\n✓ Excel saved: {output_path}")
        print(f"  Size: {size_mb:.2f} MB  |  Sheets: {total_sheets}  |  Pages: {len(document_pages)}")
        return output_path
    except PermissionError:
        print(f"Error: Close the file first: {output_path}")
        sys.exit(1)
    except Exception as e:
        print(f"Error saving: {e}")
        sys.exit(1)


# ============================================================
# MAIN
# ============================================================

def main():
    if len(sys.argv) < 2:
        print("Usage: python extract_tables_smart_merged.py <pdf_path> [output_path]")
        sys.exit(1)

    pdf_path = sys.argv[1]
    if not os.path.exists(pdf_path):
        print(f"File not found: {pdf_path}")
        sys.exit(1)

    if len(sys.argv) >= 3:
        output_path = sys.argv[2]
    else:
        output_path = Path(pdf_path).stem + "_IRDA_Forms.xlsx"

    if not output_path.lower().endswith(".xlsx"):
        output_path += ".xlsx"

    print("=" * 60)
    print("Universal IRDAI PDF → Excel Extractor")
    print("=" * 60)

    forms = extract_forms_from_pdf(pdf_path)
    if not forms:
        print("No content extracted.")
        sys.exit(0)

    create_excel_from_forms(forms, output_path)

    print("=" * 60)
    print("Done!")
    print("=" * 60)


if __name__ == "__main__":
    main()