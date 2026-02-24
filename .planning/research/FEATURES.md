# Feature Landscape

**Domain:** PDF Table Extraction
**Researched:** 2026-02-12

## Table Stakes

Features users expect. Missing = product feels incomplete.

| Feature | Why Expected | Complexity | Notes |
|---------|--------------|------------|-------|
| Grid Table Extraction | Core function | Med | Standard `pdfplumber` |
| Batch Processing | Efficiency for multiple files | Low | Already implemented |
| Basic Excel Formatting | Usability | Low | Headers, alternating rows |

## Differentiators

Features that set product apart.

| Feature | Value Proposition | Complexity | Notes |
|---------|-------------------|------------|-------|
| Hybrid Extraction | Handles both stream (text) and lattice (lines) tables | High | Requires strategy pattern |
| Smart Type Detection | Formats dates/numbers correctly in Excel | Med | Heuristics on cell content |
| Configuration | Users can tune settings per project | Med | `config.json` support |

## Anti-Features

Features to explicitly NOT build.

| Anti-Feature | Why Avoid | What to Do Instead |
|--------------|-----------|-------------------|
| Full OCR | High complexity, large deps (Tesseract) | Use text-based PDF lib only |
| GUI (Qt/Tkinter) | Maintenance burden | Keep CLI simple |

## Feature Dependencies

```
Hybrid Extraction -> Configuration (users act as tie-breaker)
Smart Type Detection -> Hybrid Extraction (needs clean data)
```

## MVP Recommendation

Prioritize:
1.  Modular Refactor (Foundation)
2.  Hybrid Extraction (Accuracy)
3.  Smart Type Detection (Usability)

Defer: GUI, OCR

---

*Feature analysis: 2026-02-12*
