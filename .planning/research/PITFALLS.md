# Domain Pitfalls

**Domain:** PDF Table Extraction
**Researched:** 2026-02-12

## Critical Pitfalls

### Pitfall 1: One-Size-Fits-All Extraction
**What goes wrong:** Using a single strategy (e.g., `pdfplumber` default) for all PDFs.
**Why it happens:** Different PDFs construct tables differently (lines vs. whitespace).
**Consequences:** Poor extraction on some files, merged columns, missed rows.
**Prevention:** Implement strategy pattern: try Lattice (lines) first, fall back to Stream (whitespace).

### Pitfall 2: Memory Usage
**What goes wrong:** Loading entire large PDFs into memory.
**Why it happens:** `pdfplumber.open()` loads the whole file structure.
**Consequences:** Script crashes on large reports.
**Prevention:** Process page-by-page and release resources.

## Moderate Pitfalls

### Pitfall 1: Formatting Loss
**What goes wrong:** Excel output loses meaning (e.g., headers indistinguishable from data).
**Prevention:** Heuristics to detect bold text or larger font size as headers.

---

*Pitfall analysis: 2026-02-12*
