# Project Roadmap

**Analysis Date:** 2026-02-12

## Phase 1: Foundation & Cleanup
- Goal: ensuring code quality and basic structure.
- Tasks:
  - Normalize project structure.
  - Add type hinting and linting.
  - Standardize error handling.

## Phase 2: Core Extraction Improvement
- Goal: Improve accuracy of table extraction.
- Tasks:
  - **Fix: Prevent multiple data rows merging into single row (Switching to text-based row detection).**
  - **Fix: Prevent multiple data rows merging into single row (Switching to text-based row detection).**
  - ~~Fix: Prevent data columns merging (Reverted due to over-splitting).~~
  - Investigate `pdfplumber` settings for better detection.
  - Handle edge cases (merged cells, multi-line headers).

## Phase 3: Advanced Features
- Goal: Add smart processing.
- Tasks:
  - Detect and format data types (dates, currencies).
  - Add configuration file support.

## Phase 4: User Experience
- Goal: Make it easier to use.
- Tasks:
  - Improve CLI arguments and help messages.
  - Create a simple GUI (optional).

---

*Roadmap definition: 2026-02-12*
