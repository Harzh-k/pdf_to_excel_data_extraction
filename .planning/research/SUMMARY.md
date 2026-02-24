# Research Summary: PDF Table Extractor

**Domain:** Data Extraction Tool
**Researched:** 2026-02-12
**Overall confidence:** MEDIUM

## Executive Summary

The current implementation uses `pdfplumber`, which is a solid choice for text-based PDFs. However, libraries like `camelot-py` (specifically with 'lattice' mode) often perform better for tables with grid lines, while `pdfplumber` excels at text layout analysis.

## Key Findings

**Stack:** `pdfplumber` + `openpyxl` is effective but could be enhanced with `camelot` or `tabula-py` for specific table types.
**Architecture:** Current script is monolithic. Refactoring into a modular pipeline (Reader -> Extractor -> Formatter -> Writer) would improve maintainability.
**Critical pitfall:** Relying solely on one extraction method. Different PDFs require different strategies (stream vs. lattice).

## Implications for Roadmap

Based on research, suggested phase structure:

1.  **Modularization** - Break down the monolithic script.
    - Addresses: Maintainability and extensibility.
    - Avoids: Spaghetti code as features grow.

2.  **Hybrid Extraction Strategy** - Integrate `camelot` or advanced `pdfplumber` config.
    - Addresses: Handling of complex tables (merged cells, no borders).

3.  **Smart Formatting** - Heuristics for header detection and data typing.

## Gaps to Address

-   OCR capabilities for scanned PDFs (currently out of scope but worth noting).
