# Context Phase 2: Core Extraction Improvement

**Goal:** Improve `pdfplumber` extraction parameters to prevent multiple distinct text rows from merging into a single table row.

## Problem Description
The user reported that "many information data is in one row" instead of being split into separate rows. This typically occurs when:
1.  **Lattice Mode Fail:** The PDF has visual table borders, but `pdfplumber` (defaulting to lattice) misses some horizontal lines, causing adjacent rows to merge.
2.  **Stream Mode Fail:** The PDF relies on whitespace, but the text is packed too closely vertically, so `pdfplumber` (in stream mode) guesses they belong to the same row.

## Solution Strategy
We will modify `src/extractor.py` to use a **Hybrid Strategy**:
1.  **Explicit Table Settings:** instead of bare `page.extract_tables()`, we will pass a customizable configuration.
2.  **Strategy Iteration:** We will try to extract with `vertical_strategy='lines'` and `horizontal_strategy='lines'` (strict lattice) first.
3.  **Fallback:** If that fails or yields poor results, we might try `horizontal_strategy='text'` to force row splitting based on text alignment.
4.  **Configuration:** We will expose these settings or use a robust default that favors splitting over merging.

## Proposed Changes
- Update `extract_tables_from_pdf` to accept `table_settings`.
- Implement a default setting dictionary that uses:
  ```python
  {
    "vertical_strategy": "lines", 
    "horizontal_strategy": "text", # Use text layout to define rows if lines are missing
    "intersection_y_tolerance": 10,
    "snap_tolerance": 3
  }
  ```
- Or perform a two-pass check.

---

*Context definition: Phase 2*
