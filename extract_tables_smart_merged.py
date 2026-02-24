#!/usr/bin/env python3
"""
extract_tables_smart_merged.py — Universal IRDAI PDF → Excel Extractor
=======================================================================
Converts IRDAI public disclosure PDFs into clean, multi-sheet Excel files.

Usage:
    python extract_tables_smart_merged.py <pdf_path> [output_path]

Requirements:
    pip install pdfplumber openpyxl
"""

import re
import sys
import os
import logging
from pathlib import Path

# ── Dependency check ──────────────────────────────────────────────────────────
try:
    import pdfplumber
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError as e:
    print(f"Missing dependency: {e}")
    print("Install with:  pip install pdfplumber openpyxl")
    sys.exit(1)

from src.extractor import TABLE_SETTINGS, process_table

# ── Logging ───────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-7s  %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)


# ══════════════════════════════════════════════════════════════════════════════
# STYLE DEFINITIONS  (minimal: dark-blue headers + rule lines only)
# ══════════════════════════════════════════════════════════════════════════════

_DARK_BLUE = "1F4E79"
_MID_BLUE  = "2E75B6"
_GREY_FONT = "595959"

# Sides
_S_THIN  = Side(style="thin",   color="BFBFBF")
_S_MED   = Side(style="medium", color=_DARK_BLUE)
_S_WHITE = Side(style="thin",   color="FFFFFF")

# Borders
B_NONE       = Border()
B_HDR        = Border(top=_S_MED, bottom=_S_MED)   # thick rule above+below header
B_TOTAL      = Border(top=_S_THIN, bottom=_S_THIN)  # thin rule for totals
B_SHEET_HDR  = Border(bottom=_S_MED)

# Fills
F_DARK_BLUE = PatternFill("solid", fgColor=_DARK_BLUE)
F_MID_BLUE  = PatternFill("solid", fgColor=_MID_BLUE)
F_NONE      = PatternFill("none")

# Fonts
FT_SHEET_HDR = Font(name="Calibri", bold=True,  size=10, color="FFFFFF")
FT_COL_HDR   = Font(name="Calibri", bold=True,  size=9,  color="FFFFFF")
FT_BOLD      = Font(name="Calibri", bold=True,  size=9,  color="000000")
FT_NORMAL    = Font(name="Calibri", bold=False, size=9,  color="000000")
FT_MUTED     = Font(name="Calibri", bold=False, size=8,  color=_GREY_FONT)

# Alignments
A_WRAP    = Alignment(vertical="top",    wrap_text=True)
A_CENTER  = Alignment(vertical="center", horizontal="center", wrap_text=True)
A_MID_L   = Alignment(vertical="center", indent=1)
A_NUM     = Alignment(vertical="top",    horizontal="right")


# ══════════════════════════════════════════════════════════════════════════════
# ROW CLASSIFICATION  (drives per-row styling)
# ══════════════════════════════════════════════════════════════════════════════

# IRDAI column-header keywords
_COL_HDR_KW = {
    "LIFE", "PENSION", "HEALTH", "ANNUITY",
    "LINKED BUSINESS", "PARTICIPATING", "NON-PARTICIPATING",
    "VAR.INS", "VAR. INS", "PARTICULARS",
}

# Total / subtotal markers
_TOTAL_KW = {
    "TOTAL (A)", "TOTAL (B)", "TOTAL (C)", "TOTAL (D)",
    "SUB TOTAL", "SUB-TOTAL", "SURPLUS", "DEFICIT",
    "AMOUNT AVAILABLE", "TOTAL",
}

# Form-level title markers
_FORM_HDR_KW = {
    "FORM L-", "REVENUE ACCOUNT", "PROFIT AND LOSS",
    "BALANCE SHEET", "POLICYHOLDERS", "NAME OF THE INSURER",
}


def classify_row(row, row_idx, is_index_page=False):
    """
    Return row type string used to select styling.

    Types: 'col_header' | 'total' | 'form_meta' | 'section' | 'normal'
    """
    if is_index_page:
        return "col_header" if row_idx == 0 else "normal"

    text = " ".join(str(c) for c in row if c).upper().strip()

    if any(k in text for k in _COL_HDR_KW):
        return "col_header"

    # Total rows: must match at word-boundary level to avoid false positives
    for k in _TOTAL_KW:
        if re.search(r"\b" + re.escape(k) + r"\b", text):
            return "total"

    if row_idx < 6 and any(k in text for k in _FORM_HDR_KW):
        return "form_meta"

    # ALL-CAPS section headings (e.g. "APPROPRIATIONS", "SOURCES OF FUNDS")
    words = [w for w in text.split() if len(w) > 2]
    if words and all(w.isupper() for w in words) and len(words) >= 2:
        return "section"

    return "normal"


def _is_numeric_cell(value):
    """True if the cell looks like a financial number."""
    if not value:
        return False
    v = str(value).strip().replace(",", "").replace("(", "").replace(")", "").replace("-", "")
    try:
        float(v)
        return True
    except ValueError:
        return False


def _apply_row_style(ws, excel_row, row, row_type, num_cols):
    """Write one data row to the worksheet with appropriate styling."""
    for ci, val in enumerate(row, start=1):
        cell       = ws.cell(row=excel_row, column=ci, value=val if val else None)
        is_numeric = ci > 1 and _is_numeric_cell(val)

        if row_type == "col_header":
            cell.font      = FT_COL_HDR
            cell.fill      = F_DARK_BLUE
            cell.border    = B_HDR
            cell.alignment = A_CENTER

        elif row_type == "total":
            cell.font      = FT_BOLD
            cell.fill      = F_NONE
            cell.border    = B_TOTAL
            cell.alignment = A_NUM if is_numeric else A_WRAP

        elif row_type == "form_meta":
            cell.font      = FT_MUTED
            cell.fill      = F_NONE
            cell.border    = B_NONE
            cell.alignment = A_WRAP

        elif row_type == "section":
            cell.font      = FT_BOLD
            cell.fill      = F_NONE
            cell.border    = B_NONE
            cell.alignment = A_WRAP

        else:  # normal
            cell.font      = FT_NORMAL
            cell.fill      = F_NONE
            cell.border    = B_NONE
            cell.alignment = A_NUM if is_numeric else A_WRAP


# ══════════════════════════════════════════════════════════════════════════════
# EXTRACTION ENGINE
# ══════════════════════════════════════════════════════════════════════════════

def extract_document(pdf_path):
    """
    Extract all pages from a PDF.

    Returns a list of page_block dicts:
        {
            page_number : int,
            form_name   : str | None,
            is_index    : bool,
            content     : [{'type': 'table'|'text', 'data': ...}]
        }
    """
    document = []
    current_form = None

    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)
        logger.info(f"Opened '{Path(pdf_path).name}' — {total} pages")

        for page in pdf.pages:
            pg_num = page.page_number
            text   = page.extract_text() or ""

            # Track current form name
            m = re.search(r"FORM\s+L-[\dA-Za-z\-]+", text)
            if m:
                current_form = m.group().strip()

            page_block = {
                "page_number": pg_num,
                "form_name":   current_form,
                "is_index":    False,
                "content":     [],
            }

            tables = page.find_tables(table_settings=TABLE_SETTINGS)

            if tables:
                for t_idx, table in enumerate(tables):
                    raw = table.extract()

                    # Detect index page
                    is_idx = (
                        len(raw) == 2
                        and raw[1]
                        and any("\n" in str(c) for c in raw[1] if c)
                    )
                    if is_idx:
                        page_block["is_index"] = True

                    finalised = process_table(page, table)

                    page_block["content"].append({
                        "type": "table",
                        "data": finalised,
                    })
                    logger.info(
                        f"  Page {pg_num} Table {t_idx+1}: "
                        f"{len(finalised)}r x {len(finalised[0]) if finalised else 0}c"
                    )
            else:
                # No tables — store plain text lines
                lines = [l for l in text.split("\n") if l.strip()]
                if lines:
                    page_block["content"].append({"type": "text", "data": lines})
                logger.info(f"  Page {pg_num}: no tables, {len(lines)} text lines")

            document.append(page_block)

    return document


# ══════════════════════════════════════════════════════════════════════════════
# EXCEL WRITER
# ══════════════════════════════════════════════════════════════════════════════

def _make_sheet_name(pb, used):
    """
    Build a unique Excel sheet name ≤ 31 chars.
    Format: P{n}_{FORM} e.g. 'P3_L-1-A-RA'
    """
    base = f"P{pb['page_number']}"
    if pb["form_name"]:
        suffix = pb["form_name"].replace("FORM ", "").strip()
        base   = f"P{pb['page_number']}_{suffix}"

    # Strip illegal Excel sheet-name characters
    base = re.sub(r"[\\/*?:\[\]]", "", base)[:31]

    # Deduplicate
    if base not in used:
        used.add(base)
        return base

    for n in range(2, 100):
        candidate = f"{base[:28]}_{n}"
        if candidate not in used:
            used.add(candidate)
            return candidate

    return base  # fallback (shouldn't happen)


def _set_column_widths(ws, min_w=8, max_w=40):
    """Auto-fit column widths based on content, clamped to [min_w, max_w]."""
    widths = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value:
                length = len(str(cell.value).split("\n")[0])
                widths[cell.column] = max(widths.get(cell.column, min_w), length)
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = min(w + 2, max_w)


def write_excel(document, output_path):
    """Write extracted document to a multi-sheet Excel file."""
    wb   = Workbook()
    wb.remove(wb.active)
    used = set()

    for pb in document:
        sheet_name = _make_sheet_name(pb, used)
        ws         = wb.create_sheet(title=sheet_name)
        is_index   = pb["is_index"]
        cur_row    = 1

        # ── Sheet header row ─────────────────────────────────────────────────
        title = f"Page {pb['page_number']}"
        if pb["form_name"]:
            title += f"   |   {pb['form_name']}"

        header_cell            = ws.cell(row=1, column=1, value=title)
        header_cell.font       = FT_SHEET_HDR
        header_cell.fill       = F_DARK_BLUE
        header_cell.border     = B_SHEET_HDR
        header_cell.alignment  = A_MID_L
        cur_row = 2

        # ── Content ──────────────────────────────────────────────────────────
        for item in pb["content"]:

            if item["type"] == "table":
                table_data = item["data"]
                if not table_data:
                    continue

                num_cols = max(len(r) for r in table_data)

                for ri, row in enumerate(table_data):
                    # Pad short rows
                    row = list(row) + [""] * (num_cols - len(row))
                    rt  = classify_row(row, ri, is_index)
                    _apply_row_style(ws, cur_row, row, rt, num_cols)
                    cur_row += 1

                cur_row += 1  # blank spacer after table

            elif item["type"] == "text":
                for line in item["data"]:
                    cell           = ws.cell(row=cur_row, column=1, value=line)
                    cell.font      = FT_NORMAL
                    cell.alignment = A_WRAP
                    cur_row += 1
                cur_row += 1

        # ── Formatting ───────────────────────────────────────────────────────
        _set_column_widths(ws)
        ws.freeze_panes   = "A2"
        ws.sheet_view.showGridLines = True

    # ── Save ─────────────────────────────────────────────────────────────────
    try:
        wb.save(output_path)
        size_kb = os.path.getsize(output_path) // 1024
        logger.info(f"Saved '{output_path}'  ({size_kb} KB,  {len(document)} sheets)")
    except PermissionError:
        logger.error(f"Cannot write — is the file open? {output_path}")
        sys.exit(1)
    except Exception as exc:
        logger.error(f"Save failed: {exc}")
        sys.exit(1)


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════

def main():
    if len(sys.argv) < 2:
        print(__doc__)
        print("Usage: python extract_tables_smart_merged.py <pdf_path> [output_path]")
        sys.exit(1)

    pdf_path = sys.argv[1]
    if not os.path.exists(pdf_path):
        logger.error(f"File not found: {pdf_path}")
        sys.exit(1)

    output_path = (
        sys.argv[2] if len(sys.argv) >= 3
        else str(Path(pdf_path).with_suffix("")) + "_IRDAI.xlsx"
    )
    if not output_path.lower().endswith(".xlsx"):
        output_path += ".xlsx"

    logger.info("=" * 60)
    logger.info("IRDAI PDF → Excel Extractor")
    logger.info("=" * 60)

    document = extract_document(pdf_path)

    if not document:
        logger.warning("No content extracted — check the PDF.")
        sys.exit(0)

    write_excel(document, output_path)

    logger.info("=" * 60)
    logger.info("Done.")
    logger.info("=" * 60)


if __name__ == "__main__":
    main()