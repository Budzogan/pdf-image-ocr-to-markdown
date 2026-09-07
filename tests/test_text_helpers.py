"""Markdown / text helper unit tests."""

from __future__ import annotations

from scan_to_markdown_docling import (
    _escape_markdown_cell,
    _has_usable_text,
    _normalize_pdf_text_to_markdown,
)


def test_normalize_pdf_text_to_markdown_collapses_whitespace():
    text = "Hello   world\n\n  Next line  "
    result = _normalize_pdf_text_to_markdown(text)
    assert "Hello world" in result
    assert "Next line" in result


def test_normalize_pdf_text_to_markdown_empty():
    assert _normalize_pdf_text_to_markdown("") == ""
    assert _normalize_pdf_text_to_markdown(None) == ""


def test_normalize_pdf_text_to_markdown_trims_trailing_blank_lines():
    result = _normalize_pdf_text_to_markdown("One\n\n\n")
    assert result == "One"


def test_has_usable_text():
    assert _has_usable_text("x" * 25) is True
    assert _has_usable_text("short") is False
    assert _has_usable_text("") is False
    assert _has_usable_text(None) is False
    assert _has_usable_text("   \n\t  ") is False


def test_escape_markdown_cell():
    assert _escape_markdown_cell(None) == ""
    assert _escape_markdown_cell("a|b") == "a\\|b"
    assert _escape_markdown_cell("a\nb") == "a<br>b"
    assert _escape_markdown_cell("a\\b") == "a\\\\b"
    assert _escape_markdown_cell("  spaced  ") == "spaced"
