"""
extractor.py — IRDAI PDF Table Extractor
Handles multi-column IRDAI forms with space-separated sub-columns.

Key insight from diagnostic data (pdf2.pdf page 3):
- pdfplumber detects only 9 columns (line-based) but table has 20
- Header row has 18 words; VAR. and INS are 1px apart (same column)
- merge_gap=3 correctly merges only VAR.+INS (gap=1px)
  while keeping all other columns separate (min gap 5.5px)
- Particulars (x<230), Schedule (230≤x<257), data cols (x≥257)
- Split numbers like '5 ,36,897' are post-fixed to '5,36,897'
"""

import re
import logging
logger = logging.getLogger(__name__)

# ============================================================
# TABLE DETECTION SETTINGS
# ============================================================

TABLE_SETTINGS = {
    "vertical_strategy": "lines",
    "horizontal_strategy": "lines",
    "snap_tolerance": 3,
    "join_tolerance": 3,
    "edge_min_length": 50,
    "intersection_y_tolerance": 10,
    "text_x_tolerance": 1,
}

# ============================================================
# INTERNAL HELPERS
# ============================================================

def _group_rows(words, y_tol=3):
    """Group word dicts into rows by y-position proximity."""
    if not words:
        return []
    words = sorted(words, key=lambda w: w["top"])
    rows, current = [], [words[0]]
    for w in words[1:]:
        if abs(w["top"] - current[-1]["top"]) <= y_tol:
            current.append(w)
        else:
            rows.append(sorted(current, key=lambda w: w["x0"]))
            current = [w]
    rows.append(sorted(current, key=lambda w: w["x0"]))
    return rows


def _find_header_row(rows):
    """
    Find the row that defines actual data columns.
    Picks the row with most words where all inter-word gaps >= 2px.
    Selects LIFE PENSION HEALTH... row over title rows.
    Checks first 12 rows to handle IRDAI's multi-row header blocks.
    """
    best, best_n = None, 0
    for row in rows[:12]:
        if len(row) < 3:
            continue
        xp = [w["x0"] for w in row]
        gaps = [xp[j + 1] - xp[j] for j in range(len(xp) - 1)]
        if not gaps or min(gaps) < 2:
            continue
        if len(row) > best_n:
            best_n, best = len(row), row
    return best


def _merge_header_words(header_words, merge_gap=3):
    """
    Merge adjacent header words within merge_gap px.
    merge_gap=3 handles VAR.(x1=331.97) + INS(x0=332.94) gap=1px
    while keeping all other columns (min gap 5.5px) separate.
    """
    hw = sorted(header_words, key=lambda w: w["x0"])
    groups = [{"x0": hw[0]["x0"], "x1": hw[0]["x1"], "text": hw[0]["text"]}]
    for w in hw[1:]:
        if w["x0"] - groups[-1]["x1"] <= merge_gap:
            groups[-1]["x1"] = w["x1"]
            groups[-1]["text"] += " " + w["text"]
        else:
            groups.append({"x0": w["x0"], "x1": w["x1"], "text": w["text"]})
    return groups


def _fix_number_splits(row):
    """
    Fix numbers split by pdfplumber's char-level spacing.
    e.g. '5 ,36,897' → '5,36,897' | '2 9,135' → '29,135'
    """
    result = []
    for cell in row:
        if cell:
            cleaned = re.sub(r"(\d)\s+([,\d])", r"\1\2", str(cell))
            result.append(cleaned)
        else:
            result.append(cell)
    return result


# ============================================================
# PUBLIC API
# ============================================================

def header_needs_rebuild(table_data):
    """
    Returns True if pdfplumber merged columns that should be separate.
    Detects IRDAI column keywords (LIFE, PENSION, HEALTH etc.) in same cell.
    - Skips multi-line cells (index/content pages) to avoid false positives.
    - Checks first 8 rows because IRDAI tables have 5-6 header rows before data.
    """
    IRDAI_COL_KEYWORDS = ["LIFE", "PENSION", "HEALTH", "ANNUITY", "VAR.INS", "VAR. INS"]
    if not table_data:
        return False
    for row in table_data[:8]:
        for cell in row:
            if not cell:
                continue
            text = str(cell)
            if "\n" in text:
                continue  # skip multi-line cells (index/content pages)
            found = [k for k in IRDAI_COL_KEYWORDS if k in text.upper()]
            if len(found) >= 2:
                return True
    return False


def rebuild_using_header_spans(page, bbox):
    """
    Rebuild table using true column header row as boundary anchors.

    Column layout (verified on HDFC pdf2.pdf page 3):
    - Col 0 (Particulars): x0 < 230
    - Col 1 (Schedule):    230 <= x0 < (first_data_col - 5)
    - Col 2+ (data):       assigned by midpoint boundaries from header row

    Algorithm:
    1. Extract words within table bbox
    2. Group into rows by y-position
    3. Find true header row (most words with min gap >= 2px)
    4. Merge words within 3px (VAR. + INS → one column)
    5. Build split boundaries as midpoints between column groups
    6. Assign all words to columns using boundaries
    7. Post-fix split numbers

    Returns list of rows (list of strings), or None if rebuild fails.
    """
    x0, top, x1, bottom = bbox
    words = page.extract_words(x_tolerance=2, y_tolerance=2)
    words = [w for w in words if x0 <= w["x0"] <= x1 and top <= w["top"] <= bottom]
    if not words:
        return None

    rows = _group_rows(words)
    header = _find_header_row(rows)
    if not header or len(header) < 3:
        return None

    groups = _merge_header_words(header)
    if len(groups) < 3:
        return None

    # Column boundaries
    data_col_start = groups[0]["x0"]
    schedule_boundary = data_col_start - 5   # typically ~257
    particulars_boundary = 230

    boundaries = [
        (groups[i]["x1"] + groups[i + 1]["x0"]) / 2
        for i in range(len(groups) - 1)
    ]

    # col0=Particulars, col1=Schedule, col2..N=data columns
    num_cols = 2 + len(groups)

    result = []
    for rw in rows:
        row = [""] * num_cols
        for w in rw:
            xc = (w["x0"] + w["x1"]) / 2
            if w["x0"] < particulars_boundary:
                ci = 0
            elif w["x0"] < schedule_boundary:
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
    return result


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


def _is_index_table(table_data):
    if len(table_data) != 2:
        return False
    row = table_data[1]
    return any("\n" in str(c) for c in row if c)


def _expand_index_table(table_data):
    header   = table_data[0]
    data_row = table_data[1]
    cols     = [str(c).split("\n") if c else [] for c in data_row]
    max_rows = max((len(c) for c in cols), default=0)
    cols     = [c + [""] * (max_rows - len(c)) for c in cols]
    result   = [header]
    for i in range(max_rows):
        result.append([c[i] if i < len(c) else "" for c in cols])
    return result


def should_rebuild(table_data):
    if not table_data:
        return False
    if _is_index_table(table_data):
        return True
    IRDAI_COL_KW = ["LIFE", "PENSION", "HEALTH", "ANNUITY", "VAR.INS", "VAR. INS"]
    for row in table_data[:8]:
        for cell in row:
            if not cell: continue
            text = str(cell)
            if "\n" in text: continue
            if sum(1 for k in IRDAI_COL_KW if k in text.upper()) >= 2:
                return True
    import re
    for row in table_data[:5]:
        for cell in row:
            if cell and re.search(r"\d\s+[,\d]", str(cell)):
                return True
    return False


def rebuild_table(page, bbox):
    return rebuild_using_header_spans(page, bbox)