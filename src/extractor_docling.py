"""
extractor_docling.py — Docling-based extractor for no-line / scanned PDFs
==========================================================================
Used as fallback when pdfplumber and gap-based extraction both fail.
Handles: scanned PDFs, image-based pages, complex no-line tables.

Performance:
  CPU:  ~1-3 pages/min  → use for small files (<20 pages)
  GPU:  ~10-20 pages/min → use for large files

Chunking strategy for large files:
  - Split PDF into chunks of CHUNK_SIZE pages
  - Process chunks in parallel using ProcessPoolExecutor
  - Merge results in page order

Requirements:
  pip install docling
  (optional) pip install torch torchvision  # for GPU acceleration
"""

import os
import logging
import tempfile
from pathlib import Path
from concurrent.futures import ProcessPoolExecutor, as_completed

logger = logging.getLogger(__name__)

# ── Config ────────────────────────────────────────────────────────────────────

CHUNK_SIZE  = 20    # pages per chunk for parallel processing
MAX_WORKERS = None  # None = auto (os.cpu_count())


# ── Device detection ──────────────────────────────────────────────────────────

def get_device() -> str:
    """
    Auto-detect best available device for docling.
    Returns: 'cuda' | 'mps' | 'cpu'
    """
    try:
        import torch
        if torch.cuda.is_available():
            gpu = torch.cuda.get_device_name(0)
            vram_gb = torch.cuda.get_device_properties(0).total_memory / 1e9
            logger.info(f"GPU detected: {gpu} ({vram_gb:.1f} GB VRAM)")
            return "cuda"
        if hasattr(torch.backends, "mps") and torch.backends.mps.is_available():
            logger.info("Apple MPS detected")
            return "mps"
    except ImportError:
        pass
    logger.info("No GPU detected — using CPU (slower for large files)")
    return "cpu"


# ── Docling table extraction ───────────────────────────────────────────────────

def _docling_available() -> bool:
    try:
        import docling  # noqa
        return True
    except ImportError:
        return False


def _extract_chunk_worker(args):
    """
    Worker function for multiprocessing.
    Extracts tables from a single PDF chunk (temp file with subset of pages).

    Args:
        args: (chunk_pdf_path, start_page, device)

    Returns:
        list of page_block dicts
    """
    chunk_path, start_page, device = args

    try:
        from docling.document_converter import DocumentConverter
        from docling.datamodel.pipeline_options import PdfPipelineOptions
        from docling.datamodel.base_models import InputFormat

        pipeline_options = PdfPipelineOptions()
        pipeline_options.do_ocr            = True
        pipeline_options.do_table_structure = True
        pipeline_options.table_structure_options.do_cell_matching = True

        # GPU / CPU setting
        if device == "cuda":
            pipeline_options.accelerator_options = {"device": "cuda", "dtype": "float16"}
        elif device == "mps":
            pipeline_options.accelerator_options = {"device": "mps"}

        converter = DocumentConverter()
        result    = converter.convert(chunk_path)
        doc       = result.document

        page_blocks = []
        for pg_idx, page in enumerate(doc.pages):
            pg_num = start_page + pg_idx
            tables = []

            # Extract tables from this page
            for table in doc.tables:
                # Check if this table belongs to this page
                if hasattr(table, 'prov') and table.prov:
                    if table.prov[0].page_no != (pg_idx + 1):
                        continue

                # Convert docling table to list-of-rows
                rows = []
                if hasattr(table, 'data') and table.data:
                    grid = table.data.grid
                    for row in grid:
                        rows.append([
                            cell.text.strip() if cell and cell.text else ""
                            for cell in row
                        ])
                if rows:
                    tables.append({"type": "table", "data": rows})

            page_blocks.append({
                "page_number": pg_num,
                "form_name":   None,  # will be detected by caller
                "is_index":    False,
                "content":     tables,
                "engine":      "docling",
            })

        return page_blocks

    except Exception as e:
        logger.error(f"Docling chunk failed (start_page={start_page}): {e}")
        return []


def _split_pdf_to_chunks(pdf_path: str, chunk_size: int) -> list:
    """
    Split a PDF into chunk temp files.
    Returns list of (temp_path, start_page_number) tuples.
    """
    try:
        from pypdf import PdfReader, PdfWriter
    except ImportError:
        try:
            from PyPDF2 import PdfReader, PdfWriter
        except ImportError:
            raise ImportError("Install pypdf: pip install pypdf")

    reader = PdfReader(pdf_path)
    total  = len(reader.pages)
    chunks = []

    for start in range(0, total, chunk_size):
        end    = min(start + chunk_size, total)
        writer = PdfWriter()
        for i in range(start, end):
            writer.add_page(reader.pages[i])

        tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
        with open(tmp.name, "wb") as f:
            writer.write(f)
        chunks.append((tmp.name, start + 1))  # 1-indexed page number

    logger.info(f"Split {total} pages into {len(chunks)} chunks of {chunk_size}")
    return chunks


def extract_with_docling(pdf_path: str, page_numbers: list = None) -> list:
    """
    Extract tables from a PDF using docling.
    Handles large files via chunking + parallel processing.

    Args:
        pdf_path:     path to PDF file
        page_numbers: optional list of specific pages to process (1-indexed)
                      if None, processes all pages

    Returns:
        list of page_block dicts (same format as pdfplumber path)
    """
    if not _docling_available():
        raise ImportError(
            "docling not installed.\n"
            "Install with: pip install docling\n"
            "For GPU support: pip install docling torch torchvision"
        )

    device = get_device()
    chunks = _split_pdf_to_chunks(pdf_path, CHUNK_SIZE)

    # Filter chunks if specific pages requested
    if page_numbers:
        pg_set = set(page_numbers)
        chunks = [
            (p, s) for p, s in chunks
            if any(s <= pg <= s + CHUNK_SIZE - 1 for pg in pg_set)
        ]

    workers  = min(MAX_WORKERS or os.cpu_count() or 1, len(chunks))
    args     = [(path, start, device) for path, start in chunks]
    all_blocks = []

    logger.info(
        f"Docling extraction: {len(chunks)} chunks, "
        f"{workers} workers, device={device}"
    )

    try:
        if workers == 1 or device in ("cuda", "mps"):
            # GPU: single process (GPU can't be shared across processes)
            # Also single-process for small jobs
            for a in args:
                all_blocks.extend(_extract_chunk_worker(a))
        else:
            # CPU: parallel chunks
            with ProcessPoolExecutor(max_workers=workers) as pool:
                futures = {pool.submit(_extract_chunk_worker, a): a for a in args}
                for fut in as_completed(futures):
                    all_blocks.extend(fut.result())
    finally:
        # Clean up temp files
        for path, _ in chunks:
            try:
                os.unlink(path)
            except OSError:
                pass

    # Sort by page number and filter to requested pages
    all_blocks.sort(key=lambda b: b["page_number"])
    if page_numbers:
        pg_set     = set(page_numbers)
        all_blocks = [b for b in all_blocks if b["page_number"] in pg_set]

    logger.info(f"Docling: extracted {len(all_blocks)} pages")
    return all_blocks