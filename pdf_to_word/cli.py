"""Command-line interface for the PDF → DOCX converter."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from pdf_to_word.converter import convert_pdf_to_docx


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(
        prog="pdf-to-word",
        description="Convert a PDF to a Word (.docx) document preserving "
        "images, tables, links, and layout.",
    )
    parser.add_argument("pdf", help="Path to the input PDF file.")
    parser.add_argument(
        "-o",
        "--output",
        default=None,
        help="Output DOCX path. Defaults to <input>.docx.",
    )
    parser.add_argument(
        "-m",
        "--mode",
        choices=["image", "text", "hybrid"],
        default="image",
        help=(
            "Conversion mode (default: image). "
            "'image'  = pixel-perfect, each page rendered as a high-res image. "
            "'text'   = editable text via pdf2docx (may have layout issues). "
            "'hybrid' = pdf2docx with automatic image fallback for bad pages."
        ),
    )
    parser.add_argument(
        "-p",
        "--pages",
        default=None,
        help="Comma-separated 0-based page numbers to convert (e.g. 0,1,3). "
        "Defaults to all pages.",
    )
    parser.add_argument(
        "--dpi",
        type=int,
        default=300,
        help="DPI for image-based rendering (default: 300).",
    )
    parser.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        help="Print progress information.",
    )

    args = parser.parse_args(argv)

    pdf_path = Path(args.pdf)
    if not pdf_path.exists():
        print(f"Error: file not found: {pdf_path}", file=sys.stderr)
        sys.exit(1)

    pages = None
    if args.pages:
        try:
            pages = [int(p.strip()) for p in args.pages.split(",")]
        except ValueError:
            print("Error: --pages must be comma-separated integers.", file=sys.stderr)
            sys.exit(1)

    out = convert_pdf_to_docx(
        pdf_path,
        args.output,
        pages=pages,
        dpi=args.dpi,
        mode=args.mode,
        verbose=args.verbose,
    )

    print(f"Saved: {out}")


if __name__ == "__main__":
    main()
