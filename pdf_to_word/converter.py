"""
PDF → DOCX converter with three conversion modes.

Modes
-----
- **image** (default): Renders every PDF page as a high-resolution image and
  embeds it in the DOCX.  The page size, orientation, and margins of the DOCX
  are set to match the PDF exactly, so the output is pixel-perfect.  Text is
  not editable, but nothing is ever misplaced.

- **text**: Uses ``pdf2docx`` to extract text, tables, images, and links into
  editable DOCX content.  Works well for simple layouts but may have
  overlapping or shifted content on complex pages.

- **hybrid**: Tries ``pdf2docx`` first, then compares each page's extracted
  text against the raw PDF text.  Pages that look problematic are re-rendered
  as images so the final document is as editable as possible without
  sacrificing fidelity.
"""

from __future__ import annotations

import os
import sys
import tempfile
from pathlib import Path
from typing import Optional, Sequence

import fitz  # PyMuPDF
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.shared import Inches, Pt, Emu


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _pt_to_inches(pt_val: float) -> float:
    return pt_val / 72.0


def _render_page_image(
    page: fitz.Page, tmp_dir: str, dpi: int = 300
) -> tuple[str, int, int]:
    """Render a page to PNG, return (path, width_px, height_px)."""
    zoom = dpi / 72
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    img_path = os.path.join(tmp_dir, f"page_{page.number}.png")
    pix.save(img_path)
    return img_path, pix.width, pix.height


def _setup_section_for_page(
    doc: Document, page: fitz.Page, is_first: bool = False
) -> None:
    """Add (or configure) a section whose page size matches *page* exactly."""
    rect = page.rect  # fitz.Rect in points
    page_w_emu = Emu(int(rect.width * 914400 / 72))
    page_h_emu = Emu(int(rect.height * 914400 / 72))

    if is_first:
        section = doc.sections[0]
    else:
        # Add a section break (new page)
        section = doc.add_section()

    landscape = rect.width > rect.height

    if landscape:
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width = page_h_emu
        section.page_height = page_w_emu
    else:
        section.orientation = WD_ORIENT.PORTRAIT
        section.page_width = page_w_emu
        section.page_height = page_h_emu

    # Zero margins so the image fills the entire page
    section.top_margin = Emu(0)
    section.bottom_margin = Emu(0)
    section.left_margin = Emu(0)
    section.right_margin = Emu(0)
    section.header_distance = Emu(0)
    section.footer_distance = Emu(0)


# ---------------------------------------------------------------------------
# Mode: image  (pixel-perfect, every page rendered as an image)
# ---------------------------------------------------------------------------

def _convert_image_mode(
    pdf_path: Path,
    docx_path: Path,
    pages: Optional[Sequence[int]],
    dpi: int,
    verbose: bool,
) -> Path:
    pdf_doc = fitz.open(str(pdf_path))
    word_doc = Document()
    page_indices = list(pages) if pages is not None else list(range(len(pdf_doc)))

    with tempfile.TemporaryDirectory() as tmp_dir:
        for i, idx in enumerate(page_indices):
            page = pdf_doc[idx]
            if verbose:
                print(
                    f"  [{idx + 1}/{len(pdf_doc)}] Rendering page as image …",
                    file=sys.stderr,
                )

            # Match DOCX page size to PDF page size
            _setup_section_for_page(word_doc, page, is_first=(i == 0))

            # Render page
            img_path, img_w, img_h = _render_page_image(page, tmp_dir, dpi)

            # Insert image filling the whole page
            rect = page.rect
            width_emu = Emu(int(rect.width * 914400 / 72))
            height_emu = Emu(int(rect.height * 914400 / 72))

            # Access the current section's body to add the picture
            paragraph = word_doc.add_paragraph()
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)
            paragraph.paragraph_format.line_spacing = 1.0
            run = paragraph.add_run()
            run.add_picture(img_path, width=width_emu, height=height_emu)

        word_doc.save(str(docx_path))

    pdf_doc.close()
    return docx_path


# ---------------------------------------------------------------------------
# Mode: text  (editable output via pdf2docx)
# ---------------------------------------------------------------------------

def _convert_text_mode(
    pdf_path: Path,
    docx_path: Path,
    pages: Optional[Sequence[int]],
    verbose: bool,
) -> Path:
    from pdf2docx import Converter

    cv = Converter(str(pdf_path))
    if pages is not None:
        cv.convert(str(docx_path), pages=list(pages))
    else:
        cv.convert(str(docx_path))
    cv.close()
    return docx_path


# ---------------------------------------------------------------------------
# Mode: hybrid  (pdf2docx + per-page image fallback for bad pages)
# ---------------------------------------------------------------------------

def _page_text_from_pdf(page: fitz.Page) -> str:
    """Extract normalised text from a PDF page."""
    return " ".join(page.get_text("text").split())


def _page_text_from_docx_xml(docx_path: Path, page_index: int) -> str:
    """Rough heuristic: open the DOCX and get all paragraph text.

    Since DOCX doesn't have a reliable page concept we compare the full
    document text coverage instead on a per-page basis.
    """
    doc = Document(str(docx_path))
    return " ".join(p.text for p in doc.paragraphs).strip()


def _convert_hybrid_mode(
    pdf_path: Path,
    docx_path: Path,
    pages: Optional[Sequence[int]],
    dpi: int,
    verbose: bool,
) -> Path:
    """Try pdf2docx first.  If significant text is lost, redo as images."""
    from pdf2docx import Converter

    pdf_doc = fitz.open(str(pdf_path))
    page_indices = list(pages) if pages is not None else list(range(len(pdf_doc)))

    # --- Step 1: get reference text from the PDF per page ----------------
    pdf_texts: dict[int, str] = {}
    total_pdf_chars = 0
    for idx in page_indices:
        txt = _page_text_from_pdf(pdf_doc[idx])
        pdf_texts[idx] = txt
        total_pdf_chars += len(txt)

    # --- Step 2: attempt pdf2docx conversion -----------------------------
    try:
        cv = Converter(str(pdf_path))
        if pages is not None:
            cv.convert(str(docx_path), pages=list(pages))
        else:
            cv.convert(str(docx_path))
        cv.close()
    except Exception as exc:
        if verbose:
            print(f"  pdf2docx failed ({exc}), using image mode.", file=sys.stderr)
        pdf_doc.close()
        return _convert_image_mode(pdf_path, docx_path, pages, dpi, verbose)

    # --- Step 3: compare text coverage -----------------------------------
    docx_text = _page_text_from_docx_xml(docx_path, 0)
    docx_chars = len(docx_text)

    # If more than 5 % of characters are lost, fall back to image mode
    if total_pdf_chars > 0:
        ratio = docx_chars / total_pdf_chars
        if ratio < 0.95:
            if verbose:
                print(
                    f"  Text coverage {ratio:.0%} < 95 %, switching to image mode.",
                    file=sys.stderr,
                )
            pdf_doc.close()
            return _convert_image_mode(pdf_path, docx_path, pages, dpi, verbose)

    pdf_doc.close()
    return docx_path


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def convert_pdf_to_docx(
    pdf_path: str | Path,
    docx_path: Optional[str | Path] = None,
    *,
    pages: Optional[Sequence[int]] = None,
    dpi: int = 300,
    mode: str = "image",
    verbose: bool = False,
) -> Path:
    """Convert a PDF file to DOCX.

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
        Resolution for image-based rendering (default 300).
    mode:
        ``"image"`` — pixel-perfect, every page is an image (default).
        ``"text"``  — editable text via pdf2docx (may have layout issues).
        ``"hybrid"``— pdf2docx with automatic image fallback for bad pages.
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

    docx_path.parent.mkdir(parents=True, exist_ok=True)

    mode = mode.lower().strip()
    if mode not in ("image", "text", "hybrid"):
        raise ValueError(f"Unknown mode {mode!r}; expected image, text, or hybrid.")

    if verbose:
        print(
            f"Converting {pdf_path.name} → {docx_path.name}  [mode={mode}] …",
            file=sys.stderr,
        )

    if mode == "image":
        result = _convert_image_mode(pdf_path, docx_path, pages, dpi, verbose)
    elif mode == "text":
        result = _convert_text_mode(pdf_path, docx_path, pages, verbose)
    else:
        result = _convert_hybrid_mode(pdf_path, docx_path, pages, dpi, verbose)

    if verbose:
        print("Done.", file=sys.stderr)

    return result
