"""Command-line interface for the PDF → DOCX converter."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from pdf_to_word.converter import convert_pdf_to_docx


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(
        prog="pdf-to-word",
        description="Convert a PDF to a Word (.docx) document "
        "preserving layout, images, and text.",
    )
    parser.add_argument("pdf", help="Path to the input PDF file.")
    parser.add_argument(
        "-o",
        "--output",
        default=None,
        help="Output DOCX path. Defaults to <input>.docx.",
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
        help="DPI for rendering (default: 300). "
        "Higher = sharper output but larger file.",
    )
    parser.add_argument(
        "-m",
        "--mode",
        choices=["exact", "editable"],
        default="editable",
        help="Conversion mode: 'editable' (default) extracts text, images, "
        "and shapes as fully editable elements. 'exact' renders pages as "
        "images for perfect visual match but limited editability.",
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
