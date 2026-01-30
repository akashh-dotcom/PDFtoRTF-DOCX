from setuptools import setup, find_packages

setup(
    name="pdf-to-word",
    version="1.0.0",
    description="Convert PDF to Word (.docx) preserving images, tables, links and layout",
    packages=find_packages(),
    python_requires=">=3.9",
    install_requires=[
        "pdf2docx>=0.5.8",
        "python-docx>=1.1.0",
        "PyMuPDF>=1.24.0",
        "Pillow>=10.0.0",
    ],
    entry_points={
        "console_scripts": [
            "pdf-to-word=pdf_to_word.cli:main",
        ],
    },
)
