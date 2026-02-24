"""
extractor_gap.py — Gap-based column extractor for h_only and mixed PDFs
========================================================================
Handles PDFs where tables have horizontal lines but NO vertical lines
(Aditya Birla, Tata AIA, Max Life IRDAI disclosures etc.)

Algorithm:
  1. Detect data column clusters from numeric token x-positions
  2. For normal rows: x < particulars_boundary → col0, else by cluster boundaries
  3. For header/date rows: only "Particulars" word → col0, all else by boundaries
  4. Post-process: fix smashed date tokens ("202530th" → "2025")
  5. Suppress footer rows ("Version:1 Date of upload:")
"""

import re
import logging

logger = logging.getLogger(__name__)

_MIN_COL_GAP        = 20   # min px gap between two different column clusters
_MIN_CLUSTER_TOKENS = 3    # min number tokens for a cluster to be a real column
_NUM_RE    = re.compile(r'^[\d,\(\)\-\.]+$')
_DATE_RE   = re.compile(r'^\d{4}$|^30th$|^31st$|^September|^March|^December|^June')
_FOOTER_RE = re.compile(r'Version\s*:\s*\d|Date of upload', re.IGNORECASE)
# Smashed year pattern: "2025 30th" or "202530th" or "2024September"
_SMASH_RE  = re.compile(r'(\d{4})\s*(30th|31st|[A-Z][a-z]+)')


# ── Helpers ────────────────────────────────────────────────────────────────────

def _group_rows(words, y_tol=4):
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


def _detect_data_columns(words, page_width=None):
    """
    Find data column x-ranges from numeric token clustering.
    Uses only tokens in the right 60% of page to avoid row-number tokens.
    Returns list of (x_min, x_max) sorted left-to-right, or None.
    """
    left_cutoff = (page_width * 0.40) if page_width else 200

    num_xs = sorted([
        w["x0"] for w in words
        if _NUM_RE.match(w["text"]) and w["x0"] > left_cutoff
    ])

    if len(num_xs) < _MIN_CLUSTER_TOKENS * 2:
        return None

    clusters = [[num_xs[0]]]
    for x in num_xs[1:]:
        if x - clusters[-1][-1] > _MIN_COL_GAP:
            clusters.append([x])
        else:
            clusters[-1].append(x)

    clusters = [c for c in clusters if len(c) >= _MIN_CLUSTER_TOKENS]
    if not clusters:
        return None

    return [(min(c), max(c)) for c in clusters]


def _build_boundaries(col_ranges):
    """Midpoint between end of one cluster and start of next."""
    return [
        (col_ranges[i][1] + col_ranges[i + 1][0]) / 2
        for i in range(len(col_ranges) - 1)
    ]


def _fix_number_splits(row):
    """
    Fix numbers fragmented by pdfplumber character spacing.
      '1 13,330'  → '113,330'
      '5 ,36,897' → '5,36,897'
      '( 12,195)' → '(12,195)'
    Also fix smashed date tokens:
      '202530th'        → '2025'
      '2025 30th'       → '2025'
      'September, 2025 30th' → 'September, 2025'
    """
    fixed = []
    for cell in row:
        if cell:
            v = str(cell)
            # Fix split numbers
            v = re.sub(r"(\d)\s+([,\d])", r"\1\2", v)
            v = re.sub(r"\(\s+", "(", v)
            # Fix smashed dates: "202530th" or "2025 30th" → "2025"
            v = _SMASH_RE.sub(r"\1", v).strip()
            fixed.append(v)
        else:
            fixed.append(cell)
    return fixed


def _is_footer_row(row):
    return bool(_FOOTER_RE.search(" ".join(str(c) for c in row if c)))


def _row_is_text_only(rw, left_cutoff):
    """True if all words are in the left (Particulars) area."""
    return all(w["x0"] < left_cutoff for w in rw)


def _has_date_tokens(rw):
    """True if row contains date-like words (header/period row)."""
    return any(_DATE_RE.match(w["text"]) for w in rw)


def _has_data_numbers(rw, boundary):
    """True if row has numeric tokens to the right of boundary."""
    return any(
        _NUM_RE.match(w["text"]) and w["x0"] > boundary
        for w in rw
    )


# ── Column assignment ──────────────────────────────────────────────────────────

def _assign_by_boundary(xc, boundaries, num_cols):
    """Assign to column by midpoint boundaries. Col1=first data col."""
    ci = 1
    for i, b in enumerate(boundaries):
        if xc > b:
            ci = i + 2
        else:
            break
    return min(ci, num_cols - 1)


# ── Public API ─────────────────────────────────────────────────────────────────

def extract_gap_based(page, bbox=None):
    """
    Extract table from a page using gap-based column detection.

    Returns list of rows (list of strings), or None on failure.
    """
    if bbox:
        x0b, top, x1b, bottom = bbox
        words = page.extract_words(x_tolerance=2, y_tolerance=4)
        words = [w for w in words
                 if x0b <= w["x0"] <= x1b and top <= w["top"] <= bottom]
    else:
        words = page.extract_words(x_tolerance=2, y_tolerance=4)

    if not words:
        return None

    page_width  = page.width
    col_ranges  = _detect_data_columns(words, page_width)
    if not col_ranges:
        logger.warning(f"gap extractor: no data columns on page {page.page_number}")
        return None

    # particulars_boundary: left of first data cluster
    particulars_boundary = col_ranges[0][0] - 15
    boundaries           = _build_boundaries(col_ranges)
    num_cols             = 1 + len(col_ranges)  # col0 = Particulars, rest = data

    logger.info(
        f"Page {page.page_number} [gap]: {len(col_ranges)} cols "
        f"@ x={[f'{r[0]:.0f}-{r[1]:.0f}' for r in col_ranges]} "
        f"| part_boundary={particulars_boundary:.0f}"
    )

    rows_raw = _group_rows(words)
    result   = []

    for rw in rows_raw:
        # Skip empty
        if not rw:
            continue

        # Pure text rows (company name, headings): everything to col0
        if _row_is_text_only(rw, particulars_boundary):
            text = " ".join(w["text"] for w in rw).strip()
            if text and not _FOOTER_RE.search(text):
                result.append([text] + [""] * (num_cols - 1))
            continue

        # Detect header/date row: has date tokens but no actual data numbers
        is_header = (
            _has_date_tokens(rw)
            and not _has_data_numbers(rw, particulars_boundary)
        )

        row = [""] * num_cols

        for w in rw:
            xc = (w["x0"] + w["x1"]) / 2

            if is_header:
                # Header rows: only the literal word "Particulars" → col0
                # Everything else assigned by cluster boundaries
                if w["text"] == "Particulars":
                    ci = 0
                else:
                    ci = _assign_by_boundary(xc, boundaries, num_cols)
            else:
                # Normal data rows
                if w["x0"] < particulars_boundary:
                    ci = 0
                else:
                    ci = _assign_by_boundary(xc, boundaries, num_cols)

            row[ci] = (row[ci] + " " + w["text"]).strip()

        if not any(c.strip() for c in row):
            continue
        if _is_footer_row(row):
            continue

        result.append(_fix_number_splits(row))

    return result or None