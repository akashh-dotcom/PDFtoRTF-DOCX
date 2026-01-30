# PDFtoRTF-DOCX

Convert PDF files to Word (.docx) documents while preserving:

- **Text** — fonts, sizes, colours, bold/italic/underline
- **Images & figures** — embedded at original resolution
- **Tables** — with merged cells, borders, and shading
- **Hyperlinks** — clickable links carried over
- **Page layout** — margins, columns, spacing

If a page is image-only (e.g. scanned PDF), the converter automatically falls back to rendering the page as a high-resolution image so no content is lost.

## Quick start

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Convert a PDF
python convert.py input.pdf                  # → input.docx
python convert.py input.pdf -o output.docx   # custom output path
```

## Installation (editable / dev)

```bash
pip install -e .
# Now you can use the CLI anywhere:
pdf-to-word input.pdf
```

## CLI options

```
usage: pdf-to-word [-h] [-o OUTPUT] [-p PAGES] [--dpi DPI] [-v] pdf

positional arguments:
  pdf                   Path to the input PDF file.

options:
  -h, --help            show this help message and exit
  -o, --output OUTPUT   Output DOCX path. Defaults to <input>.docx.
  -p, --pages PAGES     Comma-separated 0-based page numbers (e.g. 0,1,3).
  --dpi DPI             DPI for fallback image rendering (default: 300).
  -v, --verbose         Print progress information.
```

## Use as a library

```python
from pdf_to_word import convert_pdf_to_docx

output = convert_pdf_to_docx("report.pdf", "report.docx", verbose=True)
print(f"Saved to {output}")
```

## Requirements

- Python >= 3.9
- pdf2docx
- python-docx
- PyMuPDF
- Pillow
