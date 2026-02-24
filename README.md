# IRDAI PDF → Excel Extractor

Automatically extracts tables from IRDAI insurance company public disclosure PDFs and outputs a clean, formatted Excel file — one sheet per table.

Built to handle the complex multi-column structure of IRDAI forms (e.g. FORM L-1-A-RA with 20 sub-columns across LINKED / PARTICIPATING / NON-PARTICIPATING business segments).

---

## The Problem This Solves

IRDAI public disclosure PDFs have tables with space-separated sub-columns that standard PDF extractors (pdfplumber, Adobe, etc.) collapse into a single merged cell. For example, `FORM L-1-A-RA` has 20 data columns but pdfplumber detects only 9.

This tool detects and correctly rebuilds those tables using coordinate-based column detection.

---

## Project Structure

```
irdai-pdf-extractor/
│
├── extract_tables_smart_merged.py   # Main script — run this
├── src/
│   └── extractor.py                 # Core extraction & rebuild engine
├── requirements.txt
└── README.md
```

---

## Requirements

- Python 3.8+
- pdfplumber
- openpyxl

Install dependencies:

```bash
pip install -r requirements.txt
```

---

## Usage

```bash
python extract_tables_smart_merged.py <pdf_path> [output_path]
```

**Examples:**

```bash
# Output file auto-named from PDF
python extract_tables_smart_merged.py HDFC_Life_Q3_2025.pdf

# Custom output path
python extract_tables_smart_merged.py HDFC_Life_Q3_2025.pdf output/HDFC_extracted.xlsx
```

---

## Output

- One Excel file with **one sheet per table**
- Sheet names follow the pattern: `L-1-A-RA_P3` (form code + page number)
- If a page has multiple tables: `L-1-A-RA_P3_T2`

### Formatting

| Row Type | Style |
|---|---|
| Page / Form header | Dark navy / IRDAI red background |
| Column headers (LIFE, PENSION...) | Blue background, white bold text |
| Total / Sub Total rows | Light green background, thin border |
| Normal data rows | No borders, clean white |

---

## How It Works

### The Core Problem
Standard PDF extractors use line-detection to find columns. IRDAI tables have **horizontal lines only** — sub-columns are separated by whitespace, not vertical lines. This causes pdfplumber to merge all sub-columns into one cell.

### The Fix (in `src/extractor.py`)

1. **`header_needs_rebuild()`** — detects if pdfplumber has incorrectly merged columns by looking for IRDAI column keywords (LIFE, PENSION, HEALTH, ANNUITY, VAR.INS) in the same cell

2. **`rebuild_using_header_spans()`** — rebuilds the table from scratch:
   - Extracts all words from the table bbox
   - Finds the true column header row (row with most spaced-out words)
   - Merges adjacent header words within 3px (handles `VAR.` + `INS` → `VAR. INS`)
   - Computes column boundaries as midpoints between column groups
   - Assigns every word to the correct column using x-center position

3. **Column layout** (verified on HDFC pdf2.pdf):
   - `x < 230` → Particulars column
   - `230 ≤ x < 257` → Schedule column  
   - `x ≥ 257` → Data columns (LIFE, PENSION, HEALTH, VAR.INS, TOTAL × 3 segments)

4. **Split number fix** — pdfplumber occasionally splits numbers like `5,36,897` into `5` + `,36,897` due to char-level spacing. Post-processing regex merges these back.

---

## Supported Forms

Tested on HDFC Life Insurance IRDAI Public Disclosures (December 2025):

| Form | Description | Columns |
|---|---|---|
| L-1-A-RA | Revenue Account | 20 |
| L-2-A-PL | Profit & Loss Account | 4 |
| L-3-A-BS | Balance Sheet | 3 |
| L-4 to L-45 | Various schedules | varies |

---

## Accuracy

| Scenario | Accuracy |
|---|---|
| FORM L-1-A-RA (Revenue Account) | ~95% |
| Simple 4–6 column forms | ~90% |
| Overall on full HDFC 92-page PDF | ~85% |

Accuracy across other companies' PDFs may vary slightly depending on column x-positions. The detection is dynamic (not hardcoded), so it adapts to most standard IRDAI layouts.

---

## Limitations

- Scanned PDFs (image-based) are not supported — PDF must have selectable text
- Very complex forms like L-27 (ULIP Fund, 13 pages) may need manual review
- `particulars_boundary=230` and `schedule_boundary=257` are calibrated for standard IRDAI layouts — extreme outliers may need adjustment in `extractor.py`

---

## Contributing

Pull requests welcome. If you find a form type that extracts incorrectly:

1. Note the form name and page number
2. Run the diagnostic script to get word coordinates
3. Open an issue with the output

---

## License

MIT License — free to use, modify, and distribute.