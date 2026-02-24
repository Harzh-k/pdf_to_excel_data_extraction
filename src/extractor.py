"""
extractor.py — IRDAI PDF Table Extractor (Production Grade)

Handles all IRDAI form types:
  - Multi-column revenue forms (FORM L-1-A-RA) with 20 sub-columns
  - Standard forms (P&L, Balance Sheet) with 4-6 columns
  - Index/content pages

Key design decisions (from diagnostic analysis of HDFC pdf2.pdf):
  - Always rebuild from word positions — more reliable than pdfplumber's line-based extraction
  - merge_gap=3px: merges VAR.(x1=331.97)+INS(x0=332.94) gap=1px, keeps all others (min 5.5px)
  - Column zones: Particulars(x<230), Schedule(230≤x<data_start-5), Data(x≥data_start-5)
  - Number split fix: '5 ,36,897' → '5,36,897' via regex post-processing
  - Index pages (2-row tables with multi-line cells) are expanded into individual rows
"""

import re
import logging

logger = logging.getLogger(__name__)

# ── Table detection settings ──────────────────────────────────────────────────

TABLE_SETTINGS = {
    "vertical_strategy":       "lines",
    "horizontal_strategy":     "lines",
    "snap_tolerance":          3,
    "join_tolerance":          3,
    "edge_min_length":         50,
    "intersection_y_tolerance": 10,
    "text_x_tolerance":        1,
}

# ── Constants ─────────────────────────────────────────────────────────────────

# x-coordinate below which text belongs to Particulars column
_PARTICULARS_X_MAX = 230

# Gap (px) between first data column x0 and Schedule column right edge
_SCHEDULE_GAP = 5

# Max px gap to merge two header words into one column label (e.g. VAR. + INS)
_HEADER_MERGE_GAP = 3

# ── Internal helpers ──────────────────────────────────────────────────────────

def _group_into_rows(words, y_tol=3):
    """Cluster word dicts into rows by y-proximity."""
    if not words:
        return []
    words = sorted(words, key=lambda w: w["top"])
    rows, cur = [], [words[0]]
    for w in words[1:]:
        if abs(w["top"] - cur[-1]["top"]) <= y_tol:
            cur.append(w)
        else:
            rows.append(sorted(cur, key=lambda w: w["x0"]))
            cur = [w]
    rows.append(sorted(cur, key=lambda w: w["x0"]))
    return rows


def _find_column_header_row(rows):
    """
    Return the row that best defines the column structure.

    Heuristic: the true column-header row has the most words
    where every inter-word gap is >= 2px (rules out sentence text
    and title rows with touching characters).

    Searches the first 12 rows to cover IRDAI's multi-row header blocks.
    """
    best, best_n = None, 0
    for row in rows[:12]:
        if len(row) < 3:
            continue
        xp   = [w["x0"] for w in row]
        gaps = [xp[i + 1] - xp[i] for i in range(len(xp) - 1)]
        if not gaps or min(gaps) < 2:
            continue
        if len(row) > best_n:
            best_n, best = len(row), row
    return best


def _merge_header_words(header_words):
    """
    Merge adjacent header words within _HEADER_MERGE_GAP px.

    Verified on HDFC data:
      VAR. x1=331.97, INS x0=332.94 → gap=0.97px → merged to 'VAR. INS'
      All other column gaps ≥ 5.5px  → kept separate
    """
    hw = sorted(header_words, key=lambda w: w["x0"])
    groups = [{"x0": hw[0]["x0"], "x1": hw[0]["x1"], "text": hw[0]["text"]}]
    for w in hw[1:]:
        if w["x0"] - groups[-1]["x1"] <= _HEADER_MERGE_GAP:
            groups[-1]["x1"]   = w["x1"]
            groups[-1]["text"] += " " + w["text"]
        else:
            groups.append({"x0": w["x0"], "x1": w["x1"], "text": w["text"]})
    return groups


def _fix_number_splits(row):
    """
    Fix numbers fragmented by pdfplumber's char-level spacing analysis.
    Examples:  '5 ,36,897' → '5,36,897'   |   '( 12,195)' → '(12,195)'
    """
    fixed = []
    for cell in row:
        if cell:
            v = re.sub(r"(\d)\s+([,\d])", r"\1\2", str(cell))   # '5 ,36,897'
            v = re.sub(r"\(\s+",          r"(",     v)           # '( 12,195)'
            fixed.append(v)
        else:
            fixed.append(cell)
    return fixed


def _is_index_table(table_data):
    """
    True if this is an index/content page — identified by a 2-row table
    where the second row has cells containing multiple newlines.
    """
    if len(table_data) != 2:
        return False
    row = table_data[1]
    return any("\n" in str(c) for c in row if c)


def _expand_index_table(table_data):
    """
    Convert a 2-row pdfplumber index table (with multi-line cells) into
    individual rows: one row per form entry.

    Input:  [header_row, giant_multiline_row]
    Output: [header_row, row1, row2, ...]
    """
    header   = table_data[0]
    data_row = table_data[1]
    cols     = [str(c).split("\n") if c else [] for c in data_row]
    max_rows = max((len(c) for c in cols), default=0)
    cols     = [c + [""] * (max_rows - len(c)) for c in cols]
    result   = [header]
    for i in range(max_rows):
        result.append([c[i] if i < len(c) else "" for c in cols])
    return result


# ── Public API ────────────────────────────────────────────────────────────────

def should_rebuild(table_data):
    """
    Returns True when the raw pdfplumber extraction needs to be replaced
    with our word-coordinate-based rebuild.

    Triggers on:
      1. Index pages (2-row multi-line tables) — needs expansion
      2. IRDAI column-header detection (LIFE/PENSION/HEALTH in same cell)
      3. Any table where a cell has split numbers (digit-space-comma pattern)
    """
    if not table_data:
        return False

    # Index tables always need expansion
    if _is_index_table(table_data):
        return True

    # IRDAI multi-column keyword detection
    IRDAI_COL_KW = ["LIFE", "PENSION", "HEALTH", "ANNUITY", "VAR.INS", "VAR. INS"]
    for row in table_data[:8]:
        for cell in row:
            if not cell:
                continue
            text = str(cell)
            if "\n" in text:
                continue
            if sum(1 for k in IRDAI_COL_KW if k in text.upper()) >= 2:
                return True

    # Split number detection (catches P&L, Balance Sheet etc.)
    for row in table_data[:5]:
        for cell in row:
            if cell and re.search(r"\d\s+[,\d]", str(cell)):
                return True

    return False


def rebuild_table(page, bbox):
    """
    Rebuild a table from raw word coordinates instead of pdfplumber's
    line-based cell extraction.

    Returns list-of-rows (each row is a list of strings), or None on failure.

    Column assignment:
      col 0  — Particulars  (x0 < 230)
      col 1  — Schedule     (230 ≤ x0 < first_data_col - 5)
      col 2+ — Data columns (boundaries = midpoints between header word groups)
    """
    x0, top, x1, bottom = bbox

    words = page.extract_words(x_tolerance=2, y_tolerance=2)
    words = [w for w in words
             if x0 <= w["x0"] <= x1 and top <= w["top"] <= bottom]
    if not words:
        return None

    rows   = _group_into_rows(words)
    header = _find_column_header_row(rows)
    if not header or len(header) < 2:
        logger.warning("rebuild_table: could not find usable header row")
        return None

    groups = _merge_header_words(header)
    if not groups:
        return None

    data_x0          = groups[0]["x0"]
    schedule_x_max   = data_x0 - _SCHEDULE_GAP
    boundaries       = [
        (groups[i]["x1"] + groups[i + 1]["x0"]) / 2
        for i in range(len(groups) - 1)
    ]
    num_cols = 2 + len(groups)  # Particulars + Schedule + data cols

    result = []
    for rw in rows:
        row = [""] * num_cols
        for w in rw:
            xc = (w["x0"] + w["x1"]) / 2
            if w["x0"] < _PARTICULARS_X_MAX:
                ci = 0
            elif w["x0"] < schedule_x_max:
                ci = 1
            else:
                ci = 2
                for i, b in enumerate(boundaries):
                    if xc > b:
                        ci = i + 3
                    else:
                        break
            ci = min(ci, num_cols - 1)
            row[ci] = (row[ci] + " " + w["text"]).strip()
        if any(c.strip() for c in row):
            result.append(_fix_number_splits(row))

    return result or None


def process_table(page, table):
    """
    High-level entry point called per table.
    Returns finalised list-of-rows ready for Excel output.
    """
    raw = table.extract()

    # Index page: expand multi-line cells into individual rows
    if _is_index_table(raw):
        return _expand_index_table(raw)

    # All other tables: try word-coordinate rebuild
    if should_rebuild(raw):
        rebuilt = rebuild_table(page, table.bbox)
        if rebuilt:
            return rebuilt
        logger.warning(f"rebuild_table failed for bbox={table.bbox}, using raw extraction")

    return raw