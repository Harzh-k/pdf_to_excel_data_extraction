# Architecture

**Analysis Date:** 2026-02-12

## Pattern Overview

**Overall:** Script-based synchronous processing.

**Key Characteristics:**
- CLI-driven execution
- Sequential processing (read PDF -> extract tables -> write Excel)
- No persistent state (database)

## Layers

**Core Logic:**
- Purpose: Orchestrates PDF reading and data transformation
- Location: `pdf_to_excel_tables.py`
- Contains: `extract_tables_from_pdf`, `create_excel_from_tables`

**Batch Processing:**
- Purpose: Iterates over directories to apply core logic
- Location: `batch_extract_tables.py`
- Depends on: `pdf_to_excel_tables.py`

## Data Flow

**Extraction Flow:**
1.  User invokes script with PDF path
2.  `pdfplumber` opens PDF and extracts tables as lists of lists
3.  Tables are stored in memory with page metadata
4.  `openpyxl` creates workbook
5.  Data is written to sheet-per-table
6.  Styles (fonts, colors, borders) are applied
7.  Workbook is saved to disk

## Key Abstractions

**Table Data Structure:**
- Purpose: Intermediate representation of extracted data
- Format: `{'table': [[row1], [row2]], 'page': 1, 'index_on_page': 1}`

## Entry Points

**Single File Extraction:**
- Location: `pdf_to_excel_tables.py`
- Triggers: CLI execution
- Responsibilities: Extract tables from one PDF to one Excel file

**Batch Extraction:**
- Location: `batch_extract_tables.py`
- Triggers: CLI execution
- Responsibilities: Scan directory, process all PDFs found

## Error Handling

**Strategy:** Try/Except blocks with user-friendly error messages and `sys.exit(1)`.

**Patterns:**
- Check for library imports (`try-import-except`)
- Check file existence before processing
- Catch exceptions during processing and print to stdout

---

*Architecture analysis: 2026-02-12*
