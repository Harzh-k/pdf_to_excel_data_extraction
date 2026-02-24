# Codebase Structure

**Analysis Date:** 2026-02-12

## Directory Layout

```
[project-root]/
├── .gemini/            # GSD Tools and Agents
├── .planning/          # GSD Planning Artifacts
├── Output_excel/       # Default output directory for generated Excel files
├── frontend/           # React frontend application
├── pdf_to_excel_tables.py  # Core extraction script
├── batch_extract_tables.py # Batch processing script
├── extract_tables_simple.py # Simple wrapper script
├── README_PDF_TO_EXCEL.md  # Documentation
├── QUICKSTART.md       # Quick start guide
└── USAGE_GUIDE.md      # Detailed usage guide
```

## Directory Purposes

**Output_excel:**
- Purpose: Stores generated Excel files
- Contains: `.xlsx` files
- Generated: Yes

**.gemini:**
- Purpose: Contains "Get Shit Done" (GSD) agent definitions and tools
- Contains: `agents/`, `commands/`, `get-shit-done/`

**.planning:**
- Purpose: GSD workspace for Project, Roadmap, and Codebase Map
- Contains: `PROJECT.md`, `ROADMAP.md`, `codebase/`

## Key File Locations

**Entry Points:**
- `pdf_to_excel_tables.py`: Main script for single-file extraction
- `batch_extract_tables.py`: Script for batch folder processing

**Core Logic:**
- `pdf_to_excel_tables.py`: Contains `extract_tables_from_pdf` and `create_excel_from_tables` functions

## Naming Conventions

**Files:**
- Python scripts: `snake_case.py` (e.g., `pdf_to_excel_tables.py`)
- Documentation: `UPPERCASE.md` (e.g., `USAGE_GUIDE.md`)

## Where to Add New Code

**New Feature (Extraction logic):**
- Primary code: `pdf_to_excel_tables.py`

**New Feature (Batch logic):**
- Primary code: `batch_extract_tables.py`

---

*Structure analysis: 2026-02-12*
