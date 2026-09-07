# pdf-image-ocr-to-markdown

Convert PDFs, scanned PDFs, images, and DOCX files to Markdown locally using an automatic hybrid flow:

- normal text PDFs use a lighter standard extractor first
- scanned or image-only PDFs use Docling OCR
- images use Docling OCR
- DOCX uses a lightweight python-docx path

This project is designed to handle both ordinary text PDFs and harder OCR-heavy inputs without forcing every PDF through the heavier Docling path.

---

## Quick Start For Windows

1. Put your files in this folder.
2. Double-click `CONVERT_TO_MARKDOWN.bat`.
3. Wait while it creates a local `.venv` (if needed), installs Python libraries into that venv, and checks or prepares the local Docling OCR models.
4. Markdown files appear in `md_output\`.

The batch file keeps dependencies in `.venv` so your global Python stays untouched.

## Prerequisites

```bash
pip install -r requirements.txt
```

For contributors running tests:

```bash
pip install -r requirements-dev.txt
pytest
```

```bash
python prepare_models.py
```

See repository history / local restore payload for the full README body if this commit is incomplete.
