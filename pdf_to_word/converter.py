"""
Core conversion logic.

Uses the ``pdf2docx`` library which relies on PyMuPDF (fitz) under the hood.
It faithfully reproduces:
  - Text with fonts, sizes, colours, bold/italic/underline
  - Images and figures (embedded in the DOCX)
  - Tables (with merged cells, borders, shading)
  - Hyperlinks
  - Page layout (margins, columns, headers)

For scenarios where pdf2docx cannot handle a page (e.g. scanned / image-only
PDFs), we fall back to rendering the page as a high-resolution image and
inserting it into the DOCX so no content is lost.
"""

from __future__ import annotations

import os
import sys
import tempfile
from pathlib import Path
from typing import Optional, Sequence

import fitz  # PyMuPDF
from docx import Document
from docx.shared import Inches


def _is_image_only_page(page: fitz.Page) -> bool:
    """Heuristic: a page with no extractable text is image-only."""
    return len(page.get_text("text").strip()) == 0


def _fallback_page_as_image(
    page: fitz.Page, doc_out: Document, tmp_dir: str, dpi: int = 300
) -> None:
    """Render *page* at *dpi* and insert the resulting image into *doc_out*."""
    zoom = dpi / 72
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    img_path = os.path.join(tmp_dir, f"page_{page.number}.png")
    pix.save(img_path)

    # Calculate width to fit within standard letter margins (6.5 in usable)
    img_width_in = pix.width / dpi
    width = Inches(min(img_width_in, 6.5))
    doc_out.add_picture(img_path, width=width)
    doc_out.add_page_break()


def convert_pdf_to_docx(
    pdf_path: str | Path,
    docx_path: Optional[str | Path] = None,
    *,
    pages: Optional[Sequence[int]] = None,
    dpi: int = 300,
    verbose: bool = False,
) -> Path:
    """Convert a PDF file to DOCX, preserving layout, images, tables & links.

    Parameters
    ----------
    pdf_path:
        Path to the source PDF file.
    docx_path:
        Destination path for the DOCX.  Defaults to the same name with a
        ``.docx`` extension.
    pages:
        Optional 0-based page indices to convert.  ``None`` converts all.
    dpi:
        Resolution used when falling back to image-based conversion for
        scanned / image-only pages.
    verbose:
        Print progress to stderr.

    Returns
    -------
    Path
        The path to the generated DOCX file.
    """
    pdf_path = Path(pdf_path).resolve()
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    if docx_path is None:
        docx_path = pdf_path.with_suffix(".docx")
    else:
        docx_path = Path(docx_path).resolve()

    # Ensure output directory exists
    docx_path.parent.mkdir(parents=True, exist_ok=True)

    # ------------------------------------------------------------------
    # Step 1 – Try pdf2docx for high-fidelity conversion
    # ------------------------------------------------------------------
    try:
        from pdf2docx import Converter

        if verbose:
            print(f"Converting {pdf_path.name} → {docx_path.name} …", file=sys.stderr)

        cv = Converter(str(pdf_path))

        if pages is not None:
            # pdf2docx uses 0-based page indices
            cv.convert(str(docx_path), pages=list(pages))
        else:
            cv.convert(str(docx_path))

        cv.close()

        if verbose:
            print("Conversion complete.", file=sys.stderr)

        return docx_path

    except Exception as exc:
        if verbose:
            print(
                f"pdf2docx failed ({exc}); falling back to image-based conversion.",
                file=sys.stderr,
            )

    # ------------------------------------------------------------------
    # Step 2 – Fallback: render each page as a high-res image
    # ------------------------------------------------------------------
    pdf_doc = fitz.open(str(pdf_path))
    word_doc = Document()

    page_indices = list(pages) if pages is not None else range(len(pdf_doc))

    with tempfile.TemporaryDirectory() as tmp_dir:
        for idx in page_indices:
            page = pdf_doc[idx]
            if verbose:
                print(
                    f"  Rendering page {idx + 1}/{len(pdf_doc)} as image …",
                    file=sys.stderr,
                )
            _fallback_page_as_image(page, word_doc, tmp_dir, dpi=dpi)

        word_doc.save(str(docx_path))

    pdf_doc.close()

    if verbose:
        print("Fallback conversion complete.", file=sys.stderr)

    return docx_path
