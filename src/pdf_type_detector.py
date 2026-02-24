"""
pdf_type_detector.py — Fast PDF structure classifier
Determines which extraction engine to use per page.

Types:
  'lines'    → has vertical lines  → pdfplumber (fast, current engine)
  'h_only'   → horizontal only     → gap-based column detect (medium)
  'no_lines' → no lines at all     → docling (ML, slower)
  'scanned'  → image-based page    → docling with OCR
"""

import logging
logger = logging.getLogger(__name__)


def classify_page(page) -> str:
    """
    Classify a single pdfplumber page object.
    Returns: 'lines' | 'h_only' | 'no_lines' | 'scanned'
    Cost: <5ms per page (just reads PDF metadata, no rendering)
    """
    # Check if page is image-based (scanned)
    images = page.images or []
    chars  = page.chars  or []

    # Scanned: has large image(s) covering most of page, very few chars
    if images and len(chars) < 20:
        page_area = page.width * page.height
        img_area  = sum(
            (i.get("width", 0) or 0) * (i.get("height", 0) or 0)
            for i in images
        )
        if img_area > 0.5 * page_area:
            return "scanned"

    # Check for vector lines
    lines   = page.lines or []
    v_lines = [l for l in lines if abs(l.get("x0",0) - l.get("x1",0)) < 1
               and (l.get("y1",0) - l.get("y0",0)) > 20]
    h_lines = [l for l in lines if abs(l.get("y0",0) - l.get("y1",0)) < 1
               and (l.get("x1",0) - l.get("x0",0)) > 20]

    # Also check rects — some PDFs use thin rectangles as lines
    rects   = page.rects or []
    v_rects = [r for r in rects if (r.get("width",  999) < 2)
               and  (r.get("height", 0) > 20)]
    h_rects = [r for r in rects if (r.get("height", 999) < 2)
               and  (r.get("width",  0) > 20)]

    total_v = len(v_lines) + len(v_rects)
    total_h = len(h_lines) + len(h_rects)

    if total_v >= 3:
        return "lines"      # HDFC, LIC — full grid
    if total_h >= 3:
        return "h_only"     # Aditya Birla — horizontal rules only
    return "no_lines"       # Pure whitespace separation


def classify_pdf(pdf_path: str) -> dict:
    """
    Classify all pages in a PDF. Returns summary dict.
    Used to decide which engine to spin up.

    Returns:
        {
            'dominant': 'lines' | 'h_only' | 'no_lines' | 'scanned' | 'mixed',
            'per_page': {1: 'lines', 2: 'h_only', ...},
            'counts':   {'lines': 40, 'h_only': 50, 'no_lines': 2, 'scanned': 0}
        }
    """
    import pdfplumber

    result = {"dominant": None, "per_page": {}, "counts": {
        "lines": 0, "h_only": 0, "no_lines": 0, "scanned": 0
    }}

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            t = classify_page(page)
            result["per_page"][page.page_number] = t
            result["counts"][t] += 1

    counts = result["counts"]
    total  = sum(counts.values())
    dominant_type = max(counts, key=counts.get)

    # "mixed" if no single type is >80% of pages
    if counts[dominant_type] / total < 0.8:
        result["dominant"] = "mixed"
    else:
        result["dominant"] = dominant_type

    logger.info(
        f"PDF type: {result['dominant']} "
        f"(lines={counts['lines']}, h_only={counts['h_only']}, "
        f"no_lines={counts['no_lines']}, scanned={counts['scanned']})"
    )
    return result