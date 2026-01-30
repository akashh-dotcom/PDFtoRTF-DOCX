# PDFtoRTF-DOCX

Convert PDF files to Word (.docx) documents with pixel-perfect layout.

## Conversion modes

| Mode | Layout fidelity | Editable text? | Best for |
|------|----------------|----------------|----------|
| `image` (default) | Pixel-perfect | No | Any PDF — guaranteed exact match |
| `text` | Approximate | Yes | Simple, text-heavy PDFs |
| `hybrid` | Mixed | Partially | When you want editable text where possible |

## Quick start

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Convert (default = pixel-perfect image mode)
python convert.py "input.pdf"

# 3. Editable text mode (may have layout differences)
python convert.py "input.pdf" -m text

# 4. Hybrid mode (auto-falls back to image for bad pages)
python convert.py "input.pdf" -m hybrid
```

## CLI options

```
usage: pdf-to-word [-h] [-o OUTPUT] [-m {image,text,hybrid}] [-p PAGES]
                   [--dpi DPI] [-v] pdf

positional arguments:
  pdf                   Path to the input PDF file.

options:
  -h, --help            show this help message and exit
  -o, --output OUTPUT   Output DOCX path. Defaults to <input>.docx.
  -m, --mode {image,text,hybrid}
                        Conversion mode (default: image).
  -p, --pages PAGES     Comma-separated 0-based page numbers (e.g. 0,1,3).
  --dpi DPI             DPI for image rendering (default: 300).
  -v, --verbose         Print progress information.
```

## Use as a library

```python
from pdf_to_word import convert_pdf_to_docx

# Pixel-perfect (default)
convert_pdf_to_docx("report.pdf")

# Editable text
convert_pdf_to_docx("report.pdf", mode="text")

# Custom DPI for higher quality
convert_pdf_to_docx("report.pdf", dpi=400, verbose=True)
```

## Requirements

- Python >= 3.9
- pdf2docx, python-docx, PyMuPDF, Pillow
