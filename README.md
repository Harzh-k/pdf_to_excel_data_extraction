# IRDAI PDF Extractor

> Convert IRDAI public disclosure PDFs into clean, structured Excel files — automatically, with no manual cleanup.

---

## What It Does

Insurance companies regulated by IRDAI file quarterly disclosure PDFs containing dense multi-segment financial tables: Revenue Accounts (L-1-A-RA), P&L, Balance Sheets, and supporting schedules. These PDFs are notoriously difficult to parse because:

- Tables span 18–22 columns with no vertical grid lines
- Column headers (`LIFE`, `PENSION`, `HEALTH`…) are repeated across segments
- The `Schedule` reference column (`L-4`, `L-5`…) silently merges into `Particulars`
- Some PDFs use character-level font encoding that scrambles header text

This tool extracts every table from every page into a formatted multi-sheet Excel workbook — one sheet per page, headers styled in navy, totals bolded, numbers right-aligned.

---

## Supported Insurers & Forms

| Category | Details |
|---|---|
| **Insurer types** | Life, Health, General (any IRDAI filer) |
| **Forms** | All `FORM L-*` types: L-1-A-RA, L-2-A-PL, L-3-A-BS, L-4 through L-7, L-25, L-26, and all others |
| **Tested insurers** | HDFC Life, Tata AIA, Aditya Birla, Bajaj Allianz, ICICI Prudential, LIC, Canara HSBC, Shriram Life, Star Union, Bandhan Life |
| **Page orientations** | Portrait (612px) and Landscape (792px+) |

---

## Architecture

The extractor uses a **3-engine hybrid** that classifies each page and routes it to the best parser:

```
PDF page
   │
   ▼
pdf_type_detector.classify_page()
   │
   ├── "lines"    → pdfplumber (vertical grid lines present)
   │                  └── smart override: should_use_header_cols()
   │                        ├── header_cols engine  (when pp merges columns)
   │                        └── pdfplumber          (when pp is correct)
   │
   ├── "h_only"   → header_cols engine  (LIFE/PENSION/HEALTH detection)
   │                  └── fallback: gap-based clustering
   │
   ├── "no_lines" → docling (ML-based, optional)
   │                  └── fallback: gap-based
   │
   └── "scanned"  → docling with OCR
```

### Engine Details

**Engine 1 — pdfplumber** (`src/extractor.py`)
Uses PDF line geometry to detect table cells. Fast and accurate for PDFs with full vertical grids (e.g. HDFC Life).

**Engine 2 — header_cols** (`src/extractor_gap.py → extract_header_cols`)
The key innovation for IRDAI disclosures. Instead of relying on grid lines, it:
1. Finds the row with the most `LIFE / PENSION / HEALTH / ANNUITY / TOTAL` keyword hits
2. Records the x-center of each keyword as a column anchor
3. Scans all rows for `GRAND … TOTAL` to capture the Grand Total column
4. Assigns every word on every data row to its nearest column anchor
5. Detects `L-4`, `L-5`… schedule reference words and places them in a dedicated `Schedule` column

This correctly separates `Particulars | Schedule | LIFE | PENSION | … | GRAND TOTAL` for all 19-20 column segmental forms.

**Smart selector — `should_use_header_cols()`**
Runs pdfplumber first, then overrides with header_cols when either:
- `header_cols` detected more columns than pdfplumber (pdfplumber under-detected)
- pdfplumber's col-0 contains an embedded number (pdfplumber merged data into Particulars)

This ensures HDFC Life (full vertical grid → pdfplumber 22c ✓) and Tata AIA (no verticals → header_cols 20c ✓) both get the right engine automatically.

**Engine 3 — gap-based** (`src/extractor_gap.py → extract_gap_based`)
Clusters numeric token x-positions to detect columns. Used as fallback for simple multi-column pages without LIFE/PENSION headers (e.g. Aditya Birla schedule pages).

**Engine 4 — docling** (`src/extractor_docling.py`)
ML-based table detection for scanned or zero-line PDFs. Optional — install separately.

---

## File Structure

```
.
├── app.py                        # Streamlit UI
├── extract_tables_smart_merged.py # CLI entry point & orchestrator
├── src/
│   ├── extractor.py              # pdfplumber engine + table post-processing
│   ├── extractor_gap.py          # gap-based + header_cols engines + smart selector
│   ├── extractor_docling.py      # docling/OCR engine (optional)
│   └── pdf_type_detector.py      # per-page classifier (lines/h_only/no_lines/scanned)
└── README.md
```

---

## Installation

```bash
pip install pdfplumber openpyxl streamlit
```

Optional (for scanned PDFs):
```bash
pip install docling torch torchvision
```

---

## Usage

### Web UI
```bash
streamlit run app.py
```
Open `http://localhost:8501`, upload a PDF, click **Extract to Excel**, download.

### Command Line
```bash
python extract_tables_smart_merged.py path/to/disclosure.pdf [output.xlsx]

# Force docling for all pages:
python extract_tables_smart_merged.py disclosure.pdf output.xlsx --force-docling
```


## Excel Output Format

Each page → one sheet, named `P{n}_{FORM_CODE}` (e.g. `P3_L-1-A-RA`).

| Row type | Style |
|---|---|
| Sheet title | Navy fill, white bold text |
| Column headers (`LIFE`, `PENSION`…) | Navy fill, white bold, centered |
| Total rows (`TOTAL (A)`, `SURPLUS`…) | Bold, thin top/bottom border |
| Section headings | Bold, no fill |
| Form metadata | Small grey italic |
| Data rows | Normal, numbers right-aligned |

Empty rows and footer rows (`Version:`, `Date of upload:`) are suppressed automatically.

---

## Known Limitations

| Issue | Status |
|---|---|
| Scanned / image-only PDFs | Requires `docling` install |
| Star Union value alignment on landscape pages | In progress |
| HDFC fine-tune (minor) | Tracked |

---

## How the Schedule Column Works

Every segmental form has a hidden column between `Particulars` and the first data column containing schedule references like `L-4`, `L-5`, `L-6`, `L-7`. Most PDF parsers miss it because it has no vertical separator from `Particulars`.

The header_cols engine detects it by scanning for words matching `^L-\d` and recording their minimum x-position as `sched_x`. During word assignment, the Schedule check runs *before* the Particulars boundary check, so `L-4` at x=156px (which is less than `part_boundary=167px`) still correctly lands in col-1 rather than col-0.

---

*Built for processing IRDAI quarterly public disclosures. All FORM L-* types supported.*