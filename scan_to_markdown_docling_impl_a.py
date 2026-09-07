import logging
import os
import re
import threading
import sys
import time
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path

import fitz
import pdfplumber
import psutil
import pypdfium2
from docling.datamodel.accelerator_options import AcceleratorDevice, AcceleratorOptions
from docling.datamodel.base_models import ConversionStatus, InputFormat
from docling.datamodel.pipeline_options import PdfPipelineOptions, RapidOcrOptions
from docling.document_converter import DocumentConverter, ImageFormatOption, PdfFormatOption
from docling_core.types.doc import ImageRefMode
from docx import Document
from docx.oxml.ns import qn

SUPPORTED_EXTENSIONS = {
    ".bmp",
    ".docx",
    ".jpeg",
    ".jpg",
    ".pdf",
    ".png",
    ".tif",
    ".tiff",
    ".webp",
}
IMAGE_EXTENSIONS = {".bmp", ".jpeg", ".jpg", ".png", ".tif", ".tiff", ".webp"}
SCRIPT_DIR = Path(__file__).parent
OUTPUT_DIR = SCRIPT_DIR / "md_output"
MODEL_DIR = SCRIPT_DIR / "docling_models"
DISPLAY_SCRIPT_NAME = "convert_to_markdown.py"

MIN_TOTAL_RAM_GB = 8
MIN_AVAILABLE_RAM_GB = 4
LOW_RAM_WARNING_GB = 8
SAFE_MAX_PDF_MB = 150
SAFE_MAX_PDF_PAGES = 150
SAFE_MAX_IMAGE_MB = 25
PDF_TEXT_SAMPLE_PAGES = 5
PDF_TEXT_MIN_CHARS = 25
PDF_TEXT_MIN_WORDS = 4
PDF_TEXT_MIN_TOTAL_CHARS = 50
LOW_CPU_WARNING_THREADS = 4
_DOCLING_CONVERTER_CACHE = {}
DOCLING_LAYOUT_DIR = MODEL_DIR / "docling-project--docling-layout-heron"
DOCLING_TABLE_DIR = MODEL_DIR / "docling-project--docling-models"
DOCLING_TABLEFORMER_DIR = MODEL_DIR / "docling-project--TableFormerV2"
DOCLING_RAPIDOCR_DIR = MODEL_DIR / "RapidOcr"


def _has_cached_files(path):
    return path.exists() and any(candidate.is_file() for candidate in path.rglob("*"))


def _print_docling_component_status():
    components = (
        ("Layout", DOCLING_LAYOUT_DIR),
        ("Table", DOCLING_TABLE_DIR),
        ("TableFormer", DOCLING_TABLEFORMER_DIR),
        ("RapidOCR", DOCLING_RAPIDOCR_DIR),
    )
    parts = []
    missing = []

    for label, path in components:
        if _has_cached_files(path):
            parts.append(f"{label}=ready")
        else:
            parts.append(f"{label}=missing")
            missing.append(label)

    print("  Components: " + ", ".join(parts))
    if missing:
        print("  Note      : Missing components will be downloaded during setup.")


class _DoclingProgressMonitor:
    _PAGE_PROGRESS_RE = re.compile(r"Finished converting pages (\d+)/(\d+)")

    def __init__(self, total_pages=None):
        self.total_pages = total_pages
        self.pages_done = 0
        self.stage = "starting"
        self._lock = threading.Lock()
        self._stop_event = threading.Event()
        self._thread = None
        self._start_time = None
        self._line_length = 0
        self._logger_states = []
        self._handler = self._build_handler()

    def _build_handler(self):
        monitor = self

        class ProgressHandler(logging.Handler):
            def emit(self, record):
                message = record.getMessage()
                match = monitor._PAGE_PROGRESS_RE.search(message)
                if not match:
                    return

                pages_done = int(match.group(1))
                total_pages = int(match.group(2))
                with monitor._lock:
                    monitor.pages_done = pages_done
                    monitor.total_pages = total_pages
                    monitor.stage = "processing"

        handler = ProgressHandler()
        handler.setLevel(logging.DEBUG)
        return handler

    def start(self):
        self._start_time = time.monotonic()
        for logger_name in ("docling.pipeline.base_pipeline", "RapidOCR", "onnxruntime"):
            logger = logging.getLogger(logger_name)
            handler_levels = [handler.level for handler in logger.handlers]
            self._logger_states.append(
                (logger, logger.level, logger.propagate, logger.disabled, handler_levels)
            )
            if logger_name == "docling.pipeline.base_pipeline":
                logger.setLevel(logging.DEBUG)
                logger.addHandler(self._handler)
            else:
                logger.setLevel(logging.WARNING)
                logger.disabled = True
                for handler in logger.handlers:
                    handler.setLevel(logging.WARNING)

        self._thread = threading.Thread(target=self._render_loop, daemon=True)
        self._thread.start()

    def stop(self, success=True):
        with self._lock:
            if success:
                self.stage = "done"
                if self.total_pages and self.pages_done < self.total_pages:
                    self.pages_done = self.total_pages
            else:
                self.stage = "failed"

        self._stop_event.set()
        if self._thread is not None:
            self._thread.join(timeout=2)

        self._print_line(final=True)
        if self._line_length:
            sys.stdout.write("\n")
            sys.stdout.flush()

        for logger, level, propagate, disabled, handler_levels in self._logger_states:
            if logger is logging.getLogger("docling.pipeline.base_pipeline"):
                logger.removeHandler(self._handler)
            logger.setLevel(level)
            logger.propagate = propagate
            logger.disabled = disabled
            for handler, handler_level in zip(logger.handlers, handler_levels):
                handler.setLevel(handler_level)

    def _render_loop(self):
        while not self._stop_event.wait(0.5):
            self._print_line(final=False)

    def _print_line(self, final=False):
        with self._lock:
            elapsed = 0 if self._start_time is None else int(time.monotonic() - self._start_time)
            mins, secs = divmod(elapsed, 60)
            timer = f"{mins:02d}:{secs:02d}"

            if self.total_pages:
                percent = 100 if final and self.stage == "done" else min(
                    99, int((self.pages_done / self.total_pages) * 100)
                )
                label = self.stage.capitalize()
                line = (
                    f"  Progress: {percent:3d}% ({self.pages_done}/{self.total_pages} pages) "
                    f"| {label} | Elapsed {timer}"
                )
            else:
                label = self.stage.capitalize()
                suffix = " | Complete" if final and self.stage == "done" else ""
                line = f"  Progress: working | {label} | Elapsed {timer}{suffix}"

        padded = line.ljust(self._line_length)
        self._line_length = max(self._line_length, len(line))
        sys.stdout.write("\r" + padded)
        sys.stdout.flush()


@contextmanager
def _docling_progress(total_pages=None):
    monitor = _DoclingProgressMonitor(total_pages=total_pages)
    monitor.start()
    try:
        yield
    except Exception:
        monitor.stop(success=False)
        raise
    else:
        monitor.stop(success=True)


def convert_document_to_markdown(
    input_path,
    output_dir=None,
    unsafe=False,
    force_full_ocr=False,
    pdf_mode="auto",
):
    source = Path(input_path).expanduser().resolve()
    if not source.exists():
        raise FileNotFoundError(f"File not found: {source}")

    suffix = source.suffix.lower()
    if suffix not in SUPPORTED_EXTENSIONS:
        raise ValueError(f"Unsupported file type: {suffix}")

    target_dir = Path(output_dir).expanduser().resolve() if output_dir else source.parent
    target_dir.mkdir(parents=True, exist_ok=True)

    file_size_mb = source.stat().st_size / (1024 * 1024)
    print(f"\n{'=' * 60}")
    print(f"  File    : {source.name}")
    print(f"  Size    : {file_size_mb:.2f} MB")
    print(f"  Started : {datetime.now().strftime('%H:%M:%S')}")
    print("  Stop    : Press Ctrl+C to cancel")
    print(f"{'=' * 60}")

    t_start = time.time()

    if suffix == ".docx":
        print("  Processing DOCX with python-docx...")
        markdown_content = _convert_with_python_docx(source, target_dir)
    elif suffix == ".pdf":
        markdown_content = _convert_pdf_with_best_path(
            source,
            target_dir,
            unsafe=unsafe,
            force_full_ocr=force_full_ocr,
            pdf_mode=pdf_mode,
        )
    else:
        runtime = _get_runtime_guardrails(unsafe=unsafe)
        _preflight_source(source, unsafe=unsafe)
        _print_runtime_status(runtime)
        markdown_content = _convert_with_docling(
            source,
            target_dir,
            threads=runtime["threads"],
            unsafe=unsafe,
            force_full_ocr=force_full_ocr,
        )

    elapsed = time.time() - t_start

    if markdown_content is None:
        print(f"  ERROR: Failed to convert {source.name}.")
        return None

    md_path = target_dir / f"{source.stem}.md"
    md_path.write_text(markdown_content, encoding="utf-8")

    out_size_kb = md_path.stat().st_size / 1024
    mins, secs = divmod(int(elapsed), 60)
    time_str = f"{mins}m {secs}s" if mins else f"{secs:.1f}s"

    print(f"\n{'=' * 60}")
    print("  DONE!")
    print(f"  Output  : {md_path.name}")
    print(f"  Size    : {out_size_kb:.1f} KB")
    print(f"  Time    : {time_str}")
    print(f"  Finished: {datetime.now().strftime('%H:%M:%S')}")
    print(f"{'=' * 60}\n")
    return str(md_path)


def _convert_pdf_with_best_path(source, target_dir, unsafe=False, force_full_ocr=False, pdf_mode="auto"):
    if pdf_mode not in {"auto", "text", "ocr"}:
        raise ValueError(f"Unsupported PDF mode: {pdf_mode}")

    detected_text_pdf = False
    if pdf_mode == "auto":
        detected_text_pdf = _pdf_has_embedded_text(source)
        route_label = "embedded text detected" if detected_text_pdf else "scan-like PDF detected"
        print(f"  PDF mode : auto ({route_label})")
    elif pdf_mode == "text":
        print("  PDF mode : forced text extraction")
    else:
        print("  PDF mode : forced OCR")

    if pdf_mode == "text" or (pdf_mode == "auto" and detected_text_pdf):
        print("  Processing PDF with standard text extraction...")
        markdown_content = _convert_pdf_with_text_extractors(source)
        if markdown_content:
            return markdown_content
        print("  Standard text extraction produced no usable output.")
        if pdf_mode == "text":
            return None
        print("  Falling back to Docling OCR...")

    runtime = _get_runtime_guardrails(unsafe=unsafe)
    _preflight_source(source, unsafe=unsafe)
    _print_runtime_status(runtime)
    return _convert_with_docling(
        source,
        target_dir,
        threads=runtime["threads"],
        unsafe=unsafe,
        force_full_ocr=force_full_ocr,
    )


def _get_runtime_guardrails(unsafe=False):
    vm = psutil.virtual_memory()
    total_ram_gb = vm.total / (1024**3)
    available_ram_gb = vm.available / (1024**3)
    cpu_count = os.cpu_count() or 2
    threads = max(1, min(4, cpu_count // 2 or 1))
    cpu_warning = None

    if cpu_count <= LOW_CPU_WARNING_THREADS:
        cpu_warning = (
            f"Only {cpu_count} CPU thread(s) detected. OCR should still work, "
            "but it may run noticeably slower on this system."
        )

    if total_ram_gb < MIN_TOTAL_RAM_GB and not unsafe:
        raise RuntimeError(
            f"This system has {total_ram_gb:.1f} GB RAM. "
            f"At least {MIN_TOTAL_RAM_GB} GB is recommended for the Docling OCR mode. "
            "Use --unsafe only if you accept the risk."
        )

    if available_ram_gb < MIN_AVAILABLE_RAM_GB and not unsafe:
        raise RuntimeError(
            f"Only {available_ram_gb:.1f} GB RAM is free right now. "
            f"This tool requires at least {MIN_AVAILABLE_RAM_GB} GB free in safe mode."
        )

    if available_ram_gb < LOW_RAM_WARNING_GB:
        threads = max(1, min(2, threads))

    return {
        "threads": threads,
        "cpu_count": cpu_count,
        "cpu_warning": cpu_warning,
        "total_ram_gb": total_ram_gb,
        "available_ram_gb": available_ram_gb,
    }


def _preflight_source(source, unsafe=False):
    size_mb = source.stat().st_size / (1024 * 1024)
    suffix = source.suffix.lower()

    if suffix in IMAGE_EXTENSIONS:
        if size_mb > SAFE_MAX_IMAGE_MB and not unsafe:
            raise RuntimeError(
                f"{source.name} is {size_mb:.1f} MB. "
                f"Safe mode blocks image files over {SAFE_MAX_IMAGE_MB} MB. "
                "Use --unsafe to override."
            )
        return

    if suffix != ".pdf":
        return

    page_count = _get_pdf_page_count(source)
    if size_mb > SAFE_MAX_PDF_MB and not unsafe:
        raise RuntimeError(
            f"{source.name} is {size_mb:.1f} MB. "
            f"Safe mode blocks PDFs over {SAFE_MAX_PDF_MB} MB. "
            "Use --unsafe to override."
        )

    if page_count > SAFE_MAX_PDF_PAGES and not unsafe:
        raise RuntimeError(
            f"{source.name} has {page_count} pages. "
            f"Safe mode blocks PDFs over {SAFE_MAX_PDF_PAGES} pages. "
            "Use --unsafe to override."
        )


def _get_pdf_page_count(source):
    with pypdfium2.PdfDocument(str(source)) as pdf:
        return len(pdf)


def _pdf_has_embedded_text(source):
    try:
        with fitz.open(str(source)) as pdf:
            sample_pages = min(len(pdf), PDF_TEXT_SAMPLE_PAGES)
            max_chars_on_page = 0
            total_chars = 0
            total_words = 0

            for page_index in range(sample_pages):
                text = pdf.load_page(page_index).get_text("text")
                cleaned = re.sub(r"\s+", " ", text).strip()
                if not cleaned:
                    continue

                char_count = len(cleaned)
                word_count = len(cleaned.split())
                max_chars_on_page = max(max_chars_on_page, char_count)
                total_chars += char_count
                total_words += word_count

            return (
                (max_chars_on_page >= PDF_TEXT_MIN_CHARS and total_words >= PDF_TEXT_MIN_WORDS)
                or total_chars >= PDF_TEXT_MIN_TOTAL_CHARS
            )
    except Exception as exc:
        print(f"  PDF text detection failed: {exc}")
        return False


def _convert_pdf_with_text_extractors(source):
    extractors = (
        ("PyMuPDF", _convert_pdf_with_pymupdf),
        ("pdfplumber", _convert_pdf_with_pdfplumber),
    )

    for name, extractor in extractors:
        try:
            markdown_content = extractor(source)
        except Exception as exc:
            print(f"  {name} failed: {exc}")
            continue

        if _has_usable_text(markdown_content):
            print(f"  Extractor: {name}")
            return markdown_content

    return None


def _convert_pdf_with_pymupdf(source):
    md_lines = []
    with fitz.open(str(source)) as pdf:
        for page_index, page in enumerate(pdf, start=1):
            text = page.get_text("text")
            page_markdown = _normalize_pdf_text_to_markdown(text)
            if page_markdown:
                if md_lines:
                    md_lines.append("")
                md_lines.append(f"<!-- Page {page_index} -->")
                md_lines.append("")
                md_lines.append(page_markdown)
    return "\n".join(md_lines).strip()


def _convert_pdf_with_pdfplumber(source):
    md_lines = []
    with pdfplumber.open(str(source)) as pdf:
        for page_index, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            page_markdown = _normalize_pdf_text_to_markdown(text)
            if page_markdown:
                if md_lines:
                    md_lines.append("")
                md_lines.append(f"<!-- Page {page_index} -->")
                md_lines.append("")
                md_lines.append(page_markdown)
    return "\n".join(md_lines).strip()


def _normalize_pdf_text_to_markdown(text):
    if not text:
        return ""

    normalized_lines = []
    for raw_line in text.splitlines():
        line = re.sub(r"\s+", " ", raw_line).strip()
        if not line:
            if normalized_lines and normalized_lines[-1] != "":
                normalized_lines.append("")
            continue
        normalized_lines.append(line)

    while normalized_lines and normalized_lines[-1] == "":
        normalized_lines.pop()

    return "\n".join(normalized_lines).strip()


def _has_usable_text(text):
    if not text:
        return False
    cleaned = re.sub(r"\s+", " ", text).strip()
    return len(cleaned) >= PDF_TEXT_MIN_CHARS
