# Context Phase 1: Foundation & Cleanup

**Goal:** Refactor the existing monolithic `pdf_to_excel_tables.py` into a modular, maintainable structure without changing external behavior.

## Current State
- `pdf_to_excel_tables.py` contains all logic (extraction, excel writing, style definitions).
- `batch_extract_tables.py` imports from the main script.
- Hardcoded styles and settings.

## Desired State
- **Modular Structure:**
  - `src/extractor.py`: Core `pdfplumber` logic.
  - `src/writer.py`: `openpyxl` logic.
  - `src/styles.py`: Excel style definitions.
  - `src/utils.py`: Helpers.
- **Type Hinting:** All functions should have type hints.
- **Tests:** Basic unit tests for the new modules (using `unittest` or `pytest`).

## Constraints
- Must backward compatible with existing CLI usage.
- Keep `pdf_to_excel_tables.py` as an entry point file (importing from `src`).

---

*Context definition: 2026-02-12*
