# PDFtoRTF-DOCX

Convert PDF files to **editable** Word (.docx) documents while preserving:

- **Text** — fonts, sizes, colours, bold/italic/underline (fully editable)
- **Images & figures** — embedded at high resolution (movable/resizable)
- **Tables** — with merged cells, borders, shading (editable cells)
- **Hyperlinks** — clickable links carried over
- **Page layout** — margins, columns, spacing matched to the PDF

## Quick start

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Convert a PDF → editable DOCX
python convert.py "input.pdf"

# 3. Custom output path
python convert.py "input.pdf" -o "output.docx"

# 4. Verbose progress
python convert.py "input.pdf" -v
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
  --dpi DPI             DPI for images (default: 300). Higher = sharper.
  -v, --verbose         Print progress information.
```

## Use as a library

```python
from pdf_to_word import convert_pdf_to_docx

convert_pdf_to_docx("report.pdf", verbose=True)
```

## How it works

1. **pdf2docx** extracts editable text, tables, images, and links with tuned
   parameters that reduce overlapping and line-shift artefacts.
2. **Post-processing** corrects page dimensions, orientation, and paragraph
   spacing so the DOCX layout matches the PDF as closely as possible.

## Requirements

- Python >= 3.9
- pdf2docx, python-docx, PyMuPDF, Pillow
