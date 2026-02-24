# Coding Conventions

**Analysis Date:** 2026-02-12

## Naming Patterns

**Files:**
- `snake_case.py` for Python scripts

**Functions:**
- `snake_case` for function names (e.g., `extract_tables_from_pdf`)

**Variables:**
- `snake_case` for variables
- `UPPER_CASE` for constants (e.g., `HEADER_BG_COLOR`)

## Code Style

**Formatting:**
- Standard Python (PEP 8 style generally followed)

**Linting:**
- Not strictly enforced in repo, but code follows standard conventions.

## Import Organization

**Order:**
1. Standard library (`sys`, `os`)
2. Third-party libraries (`pdfplumber`, `openpyxl`)
3. Local modules (in `batch_extract_tables.py`)

## Error Handling

**Patterns:**
- `try/except` blocks around external file operations and library imports
- Print error message to stdout and exit with `sys.exit(1)`

## Logging

**Framework:** `print` statements (console logging)

**Patterns:**
- Progress info print with specific formatting (e.g., `print(f"[{idx}/{total}] Processing: ...")`)
- Error messages prefixed with `Error:` or icons like `❌`

## Comments

**Docstrings:**
- Module-level docstrings describing purpose and usage
- Function-level docstrings in standard format (Args/Returns)

---

*Convention analysis: 2026-02-12*
