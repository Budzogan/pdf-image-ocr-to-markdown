"""Embedded-text PDF detection tests (pymupdf only; no Docling OCR)."""

from __future__ import annotations

import fitz

from scan_to_markdown_docling import _pdf_has_embedded_text


def test_pdf_has_embedded_text_true(tmp_path):
    path = tmp_path / "with_text.pdf"
    doc = fitz.open()
    page = doc.new_page()
    page.insert_text(
        (72, 72),
        "This is a sample document with enough words to count as usable text content for detection.",
    )
    doc.save(str(path))
    doc.close()
    assert _pdf_has_embedded_text(path) is True


def test_pdf_has_embedded_text_false_for_blank(tmp_path):
    path = tmp_path / "blank.pdf"
    doc = fitz.open()
    doc.new_page()
    doc.save(str(path))
    doc.close()
    assert _pdf_has_embedded_text(path) is False
