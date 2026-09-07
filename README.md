# pdf-image-ocr-to-markdown

Convert PDFs, scanned PDFs, images, and DOCX files to Markdown locally using an automatic hybrid flow:

- normal text PDFs use a lighter standard extractor first
- scanned or image-only PDFs use Docling OCR
- images use Docling OCR
- DOCX uses a lightweight python-docx path

This project is designed to handle both ordinary text PDFs and harder OCR-heavy inputs without forcing every PDF through the heavier Docling path.

---

## What Is The Big Download

This project uses [Docling](https://docling-project.github.io/docling/), an open-source document parsing system, to process scanned PDFs and images locally on your PC.

That means setup can be much heavier than a basic script because it may need to download:

- the Docling package itself
- OCR and layout AI models used for scanned documents
- supporting runtime libraries needed to run those models locally

Important clarification:

- In this project, Docling is **not** being used as a chat-style LLM.
- It is being used as a **local document AI/OCR pipeline** that reads page images, detects layout, and extracts text.
- Nothing needs to be uploaded to a cloud service for the default workflow in this repo.
- A simple way to think about it: this project runs small document-AI models on your PC, using model weights downloaded during setup for Docling layout analysis and OCR.

Why that matters:

- better for scanned PDFs and image-based documents
- heavier setup than a basic converter
- more RAM usage than the lighter `pdf-docx-to-markdown` project

If you want to learn more about the underlying project, see:

- [Docling documentation](https://docling-project.github.io/docling/)
- [Docling GitHub repository](https://github.com/docling-project/docling)

---

## Quick Start For Windows

If you just want to use it:

1. Put your `.pdf`, `.png`, `.jpg`, `.jpeg`, `.tiff`, `.bmp`, `.webp`, or `.docx` files in this folder.
2. Double-click `CONVERT_TO_MARKDOWN.bat`.
3. Wait while it creates a local `.venv` (if needed), installs Python libraries into that venv, and checks or prepares the local Docling OCR models.
4. If Windows asks for permission, allow it.
5. Your Markdown files will appear in the `md_output\` folder.

Setup can still take a while because Docling and its OCR/layout models are checked and downloaded if needed for scanned-PDF and image cases. The batch file keeps dependencies in `.venv` so your global Python stays untouched.

---

## Prerequisites

Most Windows users do not need to install anything manually. `CONVERT_TO_MARKDOWN.bat` tries to set up Python, the Python libraries, and the Docling OCR models for you.

Manual setup is mainly for users who want to run the script from the command line.

### 1. Install Python 3.10 or newer

Download from [python.org](https://www.python.org/downloads/).

> Warning: During install, check **"Add Python to PATH"**

### 2. Install required libraries

```bash
pip install -r requirements.txt
```

For contributors running tests:

```bash
pip install -r requirements-dev.txt
pytest
```

### 3. Prepare the local Docling models

```bash
python prepare_models.py
```
