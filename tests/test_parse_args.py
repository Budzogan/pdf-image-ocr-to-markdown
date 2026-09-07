"""CLI argument parsing tests."""

from __future__ import annotations

from pathlib import Path

import pytest

from scan_to_markdown_docling import OUTPUT_DIR, _parse_args


def test_parse_args_help():
    args = _parse_args(["--help"])
    assert args["help"] is True
    args_h = _parse_args(["-h"])
    assert args_h["help"] is True


def test_parse_args_pdf_mode_variants():
    assert _parse_args(["--pdf-mode", "text"])["pdf_mode"] == "text"
    assert _parse_args(["--pdf-mode=ocr"])["pdf_mode"] == "ocr"
    assert _parse_args(["--pdf-text"])["pdf_mode"] == "text"
    assert _parse_args(["--pdf-ocr"])["pdf_mode"] == "ocr"
    assert _parse_args([])["pdf_mode"] == "auto"


def test_parse_args_unsafe_and_force_full_ocr():
    args = _parse_args(["--unsafe", "--force-full-ocr"])
    assert args["unsafe"] is True
    assert args["force_full_ocr"] is True


def test_parse_args_positionals():
    args = _parse_args(["report.pdf", "custom_out"])
    assert args["input_path"] == "report.pdf"
    assert args["output_dir"] == Path("custom_out")


def test_parse_args_defaults_and_flags_with_positionals():
    args = _parse_args(["--pdf-mode", "auto", "in.pdf"])
    assert args["input_path"] == "in.pdf"
    assert args["output_dir"] == OUTPUT_DIR
    assert args["unsafe"] is False
    assert args["force_full_ocr"] is False
    assert args["system_report"] is False


def test_parse_args_invalid_pdf_mode():
    with pytest.raises(ValueError, match="PDF mode"):
        _parse_args(["--pdf-mode", "nope"])


def test_parse_args_pdf_mode_missing_value():
    with pytest.raises(ValueError, match="--pdf-mode requires"):
        _parse_args(["--pdf-mode"])
