"""
extractor_gap.py — Gap-based + Header-driven column extractor
=============================================================
Two extraction strategies, selected automatically per page:

  Engine A — gap_based      : clusters numeric-token x-positions into columns.
                              For PDFs with horizontal lines only (Aditya Birla etc.)

  Engine B — header_cols    : reads LIFE/PENSION/HEALTH/ANNUITY column-header words
                              to map every word to the exact column it belongs to.
                              Solves the Schedule-merged-into-Particulars bug that
                              affects Tata AIA, ICICI, Bajaj, LIC, Shriram, Canara
                              and any insurer whose PDF lacks vertical grid lines on
                              the segmental (L-1-A-RA) pages.

Selection rule in should_use_header_cols():
  Use header_cols when EITHER:
    (a) header engine found MORE columns than pdfplumber  -> pp under-detected
    (b) pdfplumber col-0 contains a numeric value        -> pp merged data into Particulars
  Otherwise pdfplumber is trusted (e.g. HDFC which has full vertical grid lines).
"""

import re
import logging

logger = logging.getLogger(__name__)

# ── Constants ──────────────────────────────────────────────────────────────────
_MIN_COL_GAP        = 20
_MIN_CLUSTER_TOKENS = 3
_NUM_RE    = re.compile(r'^[\d,\(\)\-\.]+$')
_DATE_RE   = re.compile(r'^\d{4}$|^30th$|^31st$|^September|^March|^December|^June')
_FOOTER_RE = re.compile(r'Version\s*:\s*\d|Date of upload', re.IGNORECASE)
_SMASH_RE  = re.compile(r'(\d{4})\s*(30th|31st|[A-Z][a-z]+)')

# Words that appear as column sub-headers in segmental IRDAI forms
_HDR_COL_WORDS = {
    'LIFE', 'PENSION', 'HEALTH', 'HEALTH#',
    'ANNUITY', 'TOTAL', 'VAR.INS', 'VAR.', 'INS', 'INS.',
    'VARIABLE',
}
_SCHED_COL_RE  = re.compile(r'^L-\d+$')   # exact: matches L-4 but NOT L-1-A-RA
_MIN_HDR_SCORE = 4    # minimum matching header-words to activate header_cols
_MIN_HDR_COLS  = 6    # minimum data-columns for this engine to make sense

_FOOTER_HDR_RE = re.compile(
    r'Version\s*[:\.]?\s*\d|Date\s+of\s+upload|^\s*\d+\s*$', re.IGNORECASE
)

# A cell with an embedded number signals pdfplumber collapsed data into Particulars
_EMBEDDED_NUM_RE = re.compile(r'\d{3,}[,\d]*')


# ══════════════════════════════════════════════════════════════════════════════
# Shared helpers
# ══════════════════════════════════════════════════════════════════════════════

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


def _fix_number_splits(row):
    """
    Fix numbers fragmented by PDF character spacing and smashed date tokens.
      '1 13,330'  -> '113,330'   '5 ,36,897' -> '5,36,897'
      '( 12,195)' -> '(12,195)'  '202530th'  -> '2025'
    """
    fixed = []
    for cell in row:
        if cell:
            v = str(cell)
            v = re.sub(r"(\d)\s+([,\d])", r"\1\2", v)
            v = re.sub(r"\(\s+", "(", v)
            v = _SMASH_RE.sub(r"\1", v).strip()
            fixed.append(v)
        else:
            fixed.append(cell)
    return fixed


def _is_footer_row(row):
    return bool(_FOOTER_RE.search(" ".join(str(c) for c in row if c)))


def _row_is_text_only(rw, left_cutoff):
    return all(w["x0"] < left_cutoff for w in rw)


def _has_date_tokens(rw):
    return any(_DATE_RE.match(w["text"]) for w in rw)


def _has_data_numbers(rw, boundary):
    return any(
        _NUM_RE.match(w["text"]) and w["x0"] > boundary
        for w in rw
    )


# ══════════════════════════════════════════════════════════════════════════════
# Engine A — gap-based
# ══════════════════════════════════════════════════════════════════════════════

def _detect_data_columns(words, page_width=None):
    """
    Find data column x-ranges from numeric token clustering.
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
    return [(min(c), max(c)) for c in clusters] if clusters else None


def _build_boundaries(col_ranges):
    return [
        (col_ranges[i][1] + col_ranges[i + 1][0]) / 2
        for i in range(len(col_ranges) - 1)
    ]


def _assign_by_boundary(xc, boundaries, num_cols):
    ci = 1
    for i, b in enumerate(boundaries):
        if xc > b:
            ci = i + 2
        else:
            break
    return min(ci, num_cols - 1)


# Split "L-14-Investments - Assets …" → ("L-14", "Investments - Assets …")
_FORM_SPLIT_RE = re.compile(
    r'^(L-[\dA-Za-z]+(?:\s*&\s*L[\dA-Za-z]+)?'
    r'(?:\s*\([^)]+\)\s*&\s*\([^)]+\))?)'
    r'\s*-\s*(.+)$'
)


def detect_index_page(words):
    """
    Return True if this page is a TOC / index-of-forms page.

    Handles two IRDAI formats:
    Format A (Axis/Shriram/most insurers):
        Header row has 'Description' AND 'Pages' as separate words
        Data rows: Sl.No | Form No | Description | Pages

    Format B (ICICI Prudential and some others):
        Header row has 'Sr.' / 'No.' / 'Form' — no separate Description/Pages columns
        Form No and Description are fused: "L-1-Revenue Account"
        Data rows: Sr.No | L-N-Description  (no Pages column)
    """
    rows = _group_rows(words)
    for row in rows[:15]:
        texts = {w['text'] for w in row}

        # Format A: standard 4-col index with Description + Pages headers
        if 'Description' in texts and 'Pages' in texts:
            return True

        # Format B: ICICI-style — header has Sr./No./Form but no Description/Pages
        # and data rows have fused "L-N-…" cells
        if {'Sr.', 'Form'}.issubset(texts):
            return True

    return False


def extract_index_page(page):
    """
    Engine INDEX: extracts the IRDAI 'List of Website Disclosures' table
    that appears as a ToC at the start of every insurer's PDF.

    The table has 4 columns:
      Sl. No. | Form No. | Description | Pages

    pdfplumber and gap-based both fail here because the three left columns
    (Sl.No / Form No / Description) have no visible column separators, so they
    get merged into one cell.  We split them by fixed x-boundaries derived from
    the actual word positions in the header row.

    Returns list of rows (list of str), or None if detection fails.
    """
    words = page.extract_words(x_tolerance=2, y_tolerance=4)
    if not words:
        return None

    if not detect_index_page(words):
        return None

    rows = _group_rows(words)

    # ── Detect which format ───────────────────────────────────────────────────
    is_format_b = False   # ICICI-style: fused "L-N-Description", no Pages column
    header_ri   = None
    pages_x     = None

    for ri, row in enumerate(rows[:15]):
        texts = {w['text'] for w in row}
        if 'Description' in texts and 'Pages' in texts:
            header_ri = ri
            for w in row:
                if w['text'] == 'Pages':
                    pages_x = w['x0'] - 5
                    break
            break
        if {'Sr.', 'Form'}.issubset(texts) and 'Description' not in texts:
            header_ri  = ri
            is_format_b = True
            break

    if header_ri is None:
        return None

    # ── FORMAT B — ICICI: split fused "L-N-Description" cells ────────────────
    if is_format_b:
        result = []
        # Header row → output 3 clean column headers
        result.append(['Sr. No.', 'Form No.', 'Description'])

        # First pass: collect raw entries preserving the weird layout
        # Pattern: some rows have only the form entry (no Sr.No), and
        # the Sr.No appears on the NEXT row as an orphan.
        # e.g. row16="L-14-Investments..." row17="14"
        # We need to look ahead: if the NEXT row has only a number and no form ref,
        # that number is the Sr.No for the CURRENT entry.
        i = header_ri + 1
        while i < len(rows):
            row = rows[i]
            if not row:
                i += 1
                continue

            row_sorted = sorted(row, key=lambda w: w['x0'])
            sr_text   = ''
            rest_text = ''
            for w in row_sorted:
                if w['x0'] < 135:
                    sr_text += (' ' if sr_text else '') + w['text']
                else:
                    rest_text += (' ' if rest_text else '') + w['text']

            flat = (sr_text + ' ' + rest_text).strip()
            if not flat or _FOOTER_HDR_RE.search(flat):
                i += 1
                continue

            # Case 1: normal row — has both Sr.No and form entry
            if sr_text and rest_text:
                m = _FORM_SPLIT_RE.match(rest_text.strip())
                if m:
                    result.append([sr_text.strip(), m.group(1).strip(), m.group(2).strip()])
                else:
                    result.append([sr_text.strip(), rest_text.strip(), ''])
                i += 1
                continue

            # Case 2: only a form entry, no Sr.No — check if NEXT row is orphan number
            if rest_text and not sr_text:
                # Look ahead for an orphan number row
                orphan_sr = ''
                if i + 1 < len(rows):
                    next_row = rows[i + 1]
                    next_sorted = sorted(next_row, key=lambda w: w['x0'])
                    next_sr   = ' '.join(w['text'] for w in next_sorted if w['x0'] < 135)
                    next_rest = ' '.join(w['text'] for w in next_sorted if w['x0'] >= 135)
                    # If next row has only a number and no form-ref text, it's orphan Sr.No
                    if next_sr and not next_rest and next_sr.strip().isdigit():
                        orphan_sr = next_sr.strip()
                        i += 1  # consume the orphan row

                m = _FORM_SPLIT_RE.match(rest_text.strip())
                if m:
                    result.append([orphan_sr, m.group(1).strip(), m.group(2).strip()])
                else:
                    result.append([orphan_sr, rest_text.strip(), ''])
                i += 1
                continue

            # Case 3: only an orphan Sr.No (no rest_text) — attach to previous entry
            if sr_text and not rest_text:
                if result and not result[-1][0]:
                    result[-1][0] = sr_text.strip()
                i += 1
                continue

            i += 1

        return result or None

    # ── FORMAT A — Standard 4-col index ──────────────────────────────────────
    if pages_x is None:
        return None

    # Derive left-column boundaries from the header row.
    # Note: 'Description' header is CENTER-aligned in its column (x0~275),
    # but description TEXT is LEFT-aligned starting at ~x0=110.
    # Use the right edge of Form No cluster + 5px, not midpoint to Description.
    hdr_row = rows[header_ri]
    form_x  = None
    desc_x  = None

    hdr_words_sorted = sorted(hdr_row, key=lambda w: w['x0'])
    seen_x = []
    for w in hdr_words_sorted:
        if not seen_x or (w['x0'] - seen_x[-1]) > 20:
            seen_x.append(w['x0'])

    if len(seen_x) >= 2:
        form_x = (seen_x[0] + seen_x[1]) / 2

    if form_x and pages_x:
        form_col_words = [
            w for w in hdr_row
            if w['x0'] >= (form_x or 50) and w['x0'] < pages_x / 2
            and w['text'] not in ('Description', 'Pages')
        ]
        if form_col_words:
            desc_x = max(w['x1'] for w in form_col_words) + 5
        else:
            desc_x = (seen_x[1] if len(seen_x) > 1 else 63) + 40

    if form_x is None: form_x = 50
    if desc_x is None: desc_x = 105

    # ── Build Format A output rows ────────────────────────────────────────────
    result = []
    for ri, row in enumerate(rows):
        if ri < header_ri:
            continue          # skip title/header rows above the table

        cells = ['', '', '', '']
        for w in row:
            x = w['x0']
            if x < form_x:
                cells[0] += (' ' if cells[0] else '') + w['text']
            elif x < desc_x:
                cells[1] += (' ' if cells[1] else '') + w['text']
            elif x < pages_x:
                cells[2] += (' ' if cells[2] else '') + w['text']
            else:
                cells[3] += (' ' if cells[3] else '') + w['text']

        # Skip footer-only rows and completely blank rows
        flat = ' '.join(c for c in cells if c)
        if not flat or _FOOTER_HDR_RE.search(flat):
            continue
        result.append([c.strip() for c in cells])

    return result or None


def extract_gap_based(page, bbox=None):
    """
    Engine A: gap-based column detection.
    Returns list of rows (list of str), or None on failure.
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

    particulars_boundary = col_ranges[0][0] - 15
    boundaries           = _build_boundaries(col_ranges)
    num_cols             = 1 + len(col_ranges)

    logger.info(
        f"Page {page.page_number} [gap]: {len(col_ranges)} cols "
        f"@ x={[f'{r[0]:.0f}-{r[1]:.0f}' for r in col_ranges]} "
        f"| part_boundary={particulars_boundary:.0f}"
    )

    rows_raw = _group_rows(words)
    result   = []

    for rw in rows_raw:
        if not rw:
            continue

        if _row_is_text_only(rw, particulars_boundary):
            text = " ".join(w["text"] for w in rw).strip()
            if text and not _FOOTER_RE.search(text):
                result.append([text] + [""] * (num_cols - 1))
            continue

        is_header = (
            _has_date_tokens(rw)
            and not _has_data_numbers(rw, particulars_boundary)
        )

        row = [""] * num_cols
        for w in rw:
            xc = (w["x0"] + w["x1"]) / 2
            if is_header:
                ci = 0 if w["text"] == "Particulars" else _assign_by_boundary(xc, boundaries, num_cols)
            else:
                ci = 0 if w["x0"] < particulars_boundary else _assign_by_boundary(xc, boundaries, num_cols)
            row[ci] = (row[ci] + " " + w["text"]).strip()

        if not any(c.strip() for c in row):
            continue
        if _is_footer_row(row):
            continue

        result.append(_fix_number_splits(row))

    return result or None


# ══════════════════════════════════════════════════════════════════════════════
# Engine B — header_cols
# ══════════════════════════════════════════════════════════════════════════════

def detect_col_centers(words, page_width):
    """
    Locate the row with the most LIFE/PENSION/HEALTH/ANNUITY/TOTAL keyword hits.
    Builds a sorted list of column-center x-coordinates.

    Three improvements over the original:
    1. sched_x uses ^L-[0-9]+$ (exact) so L-1-A-RA form title is NOT picked up
    2. 'Total' header (which lives in the row ABOVE the Life/Pension/... row)
       is captured by scanning ±2 rows around the best header row
    3. Adjacent header words that form one column (e.g. "Variable Insurance")
       are merged into a single center spanning both words, giving a more
       accurate column center that matches right-aligned data values

    Returns (col_centers, sched_x, part_boundary)
      or    (None, None, None) if not enough column headers found.
    """
    rows = _group_rows(words)

    # ── Find best header row (most LIFE/PENSION/HEALTH/... hits) ─────────────
    best_row_idx, best_score = None, 0
    for ri, row in enumerate(rows):
        score = sum(
            1 for w in row
            if w['text'].upper().rstrip('#') in _HDR_COL_WORDS
        )
        if score > best_score:
            best_score, best_row_idx = score, ri

    if best_score < _MIN_HDR_SCORE:
        return None, None, None

    # ── Collect column centers from best header row ───────────────────────────
    # Merge adjacent header words that belong to one column (e.g. "Variable Insurance")
    # by extending the span of a matching word to include immediately adjacent words.
    col_centers = []
    seen = set()
    hdr_row = rows[best_row_idx]
    used_indices = set()

    for wi, w in enumerate(hdr_row):
        if wi in used_indices:
            continue
        norm = w['text'].upper().rstrip('#')
        if norm not in _HDR_COL_WORDS:
            continue

        # Extend span to consume immediately adjacent words (gap < 5px)
        x0_span = w['x0']
        x1_span = w['x1']
        for wj in range(wi + 1, len(hdr_row)):
            nxt = hdr_row[wj]
            if nxt['x0'] - x1_span < 5:   # adjacent
                x1_span = nxt['x1']
                used_indices.add(wj)
            else:
                break

        c = round((x0_span + x1_span) / 2, 0)
        if not any(abs(c - s) < 10 for s in seen):
            col_centers.append((x0_span + x1_span) / 2)
            seen.add(c)

    # ── Also scan adjacent rows (±2) for 'Total' which often lives one row above ──
    for ri in range(max(0, best_row_idx - 2), min(len(rows), best_row_idx + 3)):
        if ri == best_row_idx:
            continue
        for w in rows[ri]:
            if w['text'].upper().rstrip('#') == 'TOTAL':
                c = round((w['x0'] + w['x1']) / 2, 0)
                if not any(abs(c - s) < 10 for s in seen):
                    col_centers.append((w['x0'] + w['x1']) / 2)
                    seen.add(c)

    # ── Scan ALL rows for "GRAND … TOTAL" ────────────────────────────────────
    for ri in range(len(rows)):
        for w in rows[ri]:
            if w['text'] in ('GRAND', 'Grand'):
                nxt = [ww for ww in rows[ri] if ww['x0'] >= w['x0'] + 5]
                if nxt and nxt[0]['text'] in ('TOTAL', 'Total'):
                    c = round((nxt[0]['x0'] + nxt[0]['x1']) / 2, 0)
                    if not any(abs(c - s) < 10 for s in seen):
                        col_centers.append((nxt[0]['x0'] + nxt[0]['x1']) / 2)
                        seen.add(c)

    col_centers = sorted(col_centers)
    if len(col_centers) < _MIN_HDR_COLS:
        return None, None, None

    # ── Find Schedule column x (L-4, L-5, L-6 …) — exact match only ─────────
    # Use only words from rows BELOW row 6 to skip form-title lines like L-1-A-RA
    sched_xs = [
        w['x0'] for w in words
        if _SCHED_COL_RE.match(w['text']) and w['top'] > 60
    ]
    sched_x = min(sched_xs) if sched_xs else None

    # ── Particulars boundary = first col_center minus half-gap ───────────────
    gap           = (col_centers[1] - col_centers[0]) / 2 if len(col_centers) > 1 else 20
    part_boundary = col_centers[0] - gap

    return col_centers, sched_x, part_boundary


def extract_header_cols(page):
    """
    Engine B: word-level column assignment using detected column-header centers.

    Layout produced:
      col 0        = Particulars      (x < part_boundary)
      col 1        = Schedule ref     (x ≈ sched_x)  — only when L-4/L-5 refs exist
      col 2 … N   = data columns      (nearest center)

    Returns list of rows (list of str), or None if detection fails.
    """
    words = page.extract_words(x_tolerance=2, y_tolerance=4)
    if not words:
        return None

    col_centers, sched_x, part_bnd = detect_col_centers(words, page.width)
    if not col_centers:
        return None

    has_sched  = sched_x is not None
    total_cols = len(col_centers) + (2 if has_sched else 1)

    # Build midpoint boundaries between consecutive column centers for zone assignment.
    # Zone-based assignment is more robust than nearest-center for right-aligned data:
    # when a number's center falls slightly past the midpoint due to right-alignment,
    # the zone boundary still gives the correct column.
    _boundaries = [
        (col_centers[i] + col_centers[i+1]) / 2
        for i in range(len(col_centers) - 1)
    ]

    def _assign(w):
        xc = (w['x0'] + w['x1']) / 2
        # Schedule check FIRST (sched col may overlap part_bnd zone)
        if has_sched and abs(xc - sched_x) < 20:
            return 1
        if xc < part_bnd:
            return 0
        # Zone-based: find which interval xc falls into
        ci = 0
        for bi, bnd in enumerate(_boundaries):
            if xc > bnd:
                ci = bi + 1
            else:
                break
        return ci + (2 if has_sched else 1)

    rows_raw = _group_rows(words)
    result   = []
    for word_row in rows_raw:
        if not word_row:
            continue
        row = [''] * total_cols
        for w in word_row:
            ci = _assign(w)
            if 0 <= ci < total_cols:
                sep      = ' ' if row[ci] else ''
                row[ci] += sep + w['text']
        flat = ' '.join(row).strip()
        if not flat or _FOOTER_HDR_RE.search(flat):
            continue
        result.append(_fix_number_splits(row))

    return result or None


# ══════════════════════════════════════════════════════════════════════════════
# Engine selector  (called from extract_tables_smart_merged.py)
# ══════════════════════════════════════════════════════════════════════════════

def _pp_col0_has_embedded_number(pp_data):
    """
    Return True if any data row's col-0 cell contains a bare number ≥ 3 digits.
    This signals pdfplumber collapsed a data column into the Particulars column.
    We skip the first 5 rows (title/header rows).
    """
    if not pp_data:
        return False
    for row in pp_data[5:20]:
        if row and row[0] and _EMBEDDED_NUM_RE.search(str(row[0])):
            return True
    return False


# Matches a Schedule ref that ran into an adjacent number e.g. "L-48,568" or "L-51,234"
_MERGED_SCHED_RE = re.compile(r'^L-\d+[,\d\(\)]+')


def _pp_has_merged_schedule(pp_data):
    """
    Return True if pdfplumber merged a Schedule column reference (L-4, L-5 …)
    with the adjacent data value into a single cell, e.g. "L-48,568".

    This is the Shriram / Star-Union symptom: because there are no vertical grid
    lines between the Schedule column and the first data column, pdfplumber treats
    them as one cell.  header_cols handles this correctly.
    Scans the first 30 data rows, all columns.
    """
    if not pp_data:
        return False
    for row in pp_data[3:30]:
        for cell in row:
            if cell and _MERGED_SCHED_RE.match(str(cell).strip()):
                return True
    return False


def should_use_header_cols(page, pp_data):
    """
    Decide whether to use header_cols instead of pdfplumber for this page.

    Returns (use_hdr: bool, hdr_data: list|None)

    Rule — prefer header_cols when ANY of:
      (a) header engine found MORE columns than pdfplumber  [pp under-detected cols]
      (b) pdfplumber col-0 contains an embedded number     [pp merged data into Particulars]
      (c) pdfplumber has a merged Schedule+value cell       [e.g. "L-48,568" — Shriram/Star-Union]
    In all other cases, trust pdfplumber (e.g. HDFC with full vertical grid lines).
    """
    words = page.extract_words(x_tolerance=2, y_tolerance=4)
    col_centers, sched_x, part_bnd = detect_col_centers(words, page.width)

    if not col_centers or len(col_centers) < _MIN_HDR_COLS:
        return False, None

    hdr_data = extract_header_cols(page)
    if not hdr_data or len(hdr_data) < 5:
        return False, None

    hdr_cols = len(hdr_data[0])
    pp_cols  = len(pp_data[0]) if pp_data else 0

    reason_a = hdr_cols > pp_cols
    reason_b = _pp_col0_has_embedded_number(pp_data)
    reason_c = _pp_has_merged_schedule(pp_data)   # "L-48,568" → Shriram/Star-Union

    if reason_a or reason_b or reason_c:
        why = []
        if reason_a: why.append(f"hdr_cols({hdr_cols})>pp_cols({pp_cols})")
        if reason_b: why.append("pp col-0 has number")
        if reason_c: why.append("pp merged schedule+value")
        logger.info(
            f"  Page {page.page_number} [header_cols chosen]: "
            f"{len(hdr_data)}r x {hdr_cols}c  reason={', '.join(why)}"
        )
        return True, hdr_data

    return False, None