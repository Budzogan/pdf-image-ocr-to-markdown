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


def _build_docling_converter(threads, force_full_ocr=False):
    cache_key = (threads, force_full_ocr)
    cached = _DOCLING_CONVERTER_CACHE.get(cache_key)
    if cached is not None:
        return cached

    MODEL_DIR.mkdir(parents=True, exist_ok=True)
    _print_docling_component_status()

    pipeline_options = PdfPipelineOptions()
    pipeline_options.artifacts_path = MODEL_DIR
    pipeline_options.do_ocr = True
    pipeline_options.do_code_enrichment = False
    pipeline_options.do_formula_enrichment = False
    pipeline_options.do_picture_description = False
    pipeline_options.do_picture_classification = False
    pipeline_options.do_chart_extraction = False
    pipeline_options.generate_page_images = False
    pipeline_options.generate_picture_images = False
    pipeline_options.generate_table_images = False
    pipeline_options.ocr_options = RapidOcrOptions(force_full_page_ocr=force_full_ocr)
    pipeline_options.accelerator_options = AcceleratorOptions(
        num_threads=threads,
        device=AcceleratorDevice.CPU,
    )

    converter = DocumentConverter(
        allowed_formats=[InputFormat.DOCX, InputFormat.IMAGE, InputFormat.PDF],
        format_options={
            InputFormat.PDF: PdfFormatOption(pipeline_options=pipeline_options),
            InputFormat.IMAGE: ImageFormatOption(pipeline_options=pipeline_options),
        },
    )
    _DOCLING_CONVERTER_CACHE[cache_key] = converter
    return converter


def _convert_with_docling(source, target_dir, threads, unsafe=False, force_full_ocr=False):
    try:
        converter = _build_docling_converter(threads=threads, force_full_ocr=force_full_ocr)
        if source.suffix.lower() == ".pdf" and not unsafe:
            max_file_size = int(SAFE_MAX_PDF_MB * 1024 * 1024)
            max_pages = SAFE_MAX_PDF_PAGES
        else:
            max_file_size = sys.maxsize
            max_pages = sys.maxsize
        total_pages = _get_pdf_page_count(source) if source.suffix.lower() == ".pdf" else None

        with _docling_progress(total_pages=total_pages):
            result = converter.convert(
                str(source),
                raises_on_error=False,
                max_num_pages=max_pages,
                max_file_size=max_file_size,
            )

        if result.status not in {ConversionStatus.SUCCESS, ConversionStatus.PARTIAL_SUCCESS}:
            if result.errors:
                for error in result.errors[:3]:
                    print(f"  Docling error: {error}")
            return None

        images_dir = target_dir / f"{source.stem}_images"
        images_dir.mkdir(parents=True, exist_ok=True)

        markdown_content = result.document.export_to_markdown(image_mode=ImageRefMode.REFERENCED)
        result.document.save_as_markdown(
            target_dir / f"{source.stem}.md",
            artifacts_dir=images_dir,
            image_mode=ImageRefMode.REFERENCED,
        )

        if not any(images_dir.iterdir()):
            try:
                images_dir.rmdir()
            except OSError:
                pass

        return markdown_content
    except Exception as exc:
        print(f"Docling conversion failed: {exc}")
        return None


def _convert_with_python_docx(source, target_dir):
    """Convert DOCX using python-docx to keep DOCX handling lightweight."""
    try:
        from docx.table import Table
        from docx.text.paragraph import Paragraph

        doc = Document(str(source))
        md_lines = []

        heading_map = {
            "Title": "# ",
            "Subtitle": "## ",
            "Heading 1": "# ",
            "Heading 2": "## ",
            "Heading 3": "### ",
            "Heading 4": "#### ",
            "Heading 5": "##### ",
        }

        images_dir = target_dir / f"{source.stem}_images"
        image_map = {}
        try:
            for rel in doc.part.rels.values():
                if "image" not in rel.reltype:
                    continue
                img_data = rel.target_part.blob
                ext = rel.target_part.content_type.split("/")[-1]
                img_filename = f"img_{len(image_map) + 1}.{ext}"
                images_dir.mkdir(parents=True, exist_ok=True)
                (images_dir / img_filename).write_bytes(img_data)
                image_map[rel.rId] = img_filename
        except Exception:
            pass

        for element in doc.element.body:
            tag = element.tag.split("}")[-1]

            if tag == "p":
                para = Paragraph(element, doc)
                content = _extract_docx_paragraph_content(para, source.stem, image_map).strip()

                if not content:
                    md_lines.append("")
                    continue

                style_name = para.style.name if para.style else ""
                prefix = heading_map.get(style_name, "")
                if prefix:
                    md_lines.append(f"\n{prefix}{content}\n")
                    continue

                num_pr = element.find(".//" + qn("w:numPr"))
                if num_pr is not None:
                    ilvl = num_pr.find(qn("w:ilvl"))
                    level = int(ilvl.get(qn("w:val"), 0)) if ilvl is not None else 0
                    md_lines.append("  " * level + f"- {content}")
                else:
                    md_lines.append(content)

            elif tag == "tbl":
                table = Table(element, doc)
                rows = table.rows
                if not rows:
                    continue

                header = [_escape_markdown_cell(cell.text) for cell in rows[0].cells]
                md_lines.append("")
                md_lines.append("| " + " | ".join(header) + " |")
                md_lines.append("| " + " | ".join(["---"] * len(header)) + " |")

                for row in rows[1:]:
                    cells = [_escape_markdown_cell(cell.text) for cell in row.cells]
                    md_lines.append("| " + " | ".join(cells) + " |")

                md_lines.append("")

        if image_map:
            print(f"  Extracted {len(image_map)} image(s) to {images_dir}")
        elif images_dir.exists():
            try:
                images_dir.rmdir()
            except OSError:
                pass

        return "\n".join(md_lines)
    except Exception as exc:
        print(f"python-docx conversion failed: {exc}")
        return None


def _extract_docx_paragraph_content(para, source_stem, image_map):
    parts = []

    for run in para.runs:
        if run.text:
            parts.append(run.text)

        for blip in run._element.findall(".//{http://schemas.openxmlformats.org/drawingml/2006/main}blip"):
            rel_id = blip.get(qn("r:embed"))
            img_filename = image_map.get(rel_id)
            if not img_filename:
                continue
            rel_path = f"{source_stem}_images/{img_filename}"
            if parts and not parts[-1].endswith((" ", "\n")):
                parts.append(" ")
            parts.append(f"![{img_filename}]({rel_path})")
            parts.append(" ")

    return "".join(parts)


def _escape_markdown_cell(cell):
    if cell is None:
        return ""

    text = str(cell).replace("\r\n", "\n").replace("\r", "\n").strip()
    text = text.replace("\\", "\\\\")
    text = text.replace("\n", "<br>")
    text = text.replace("|", "\\|")
    return text


def _print_system_report():
    runtime = _get_runtime_guardrails(unsafe=True)
    print(f"CPU threads   : {runtime['cpu_count']}")
    print(f"Total RAM     : {runtime['total_ram_gb']:.1f} GB")
    print(f"Available RAM : {runtime['available_ram_gb']:.1f} GB")
    print(f"Safe threads  : {runtime['threads']}")
    print(f"Model cache   : {MODEL_DIR}")
    print("PDF mode      : auto (embedded text => standard extractor, scan => Docling OCR)")
    print("Safe mode     : CPU-only Docling OCR")
    if runtime["cpu_warning"]:
        print(f"CPU note      : {runtime['cpu_warning']}")
    print("Stop          : Press Ctrl+C to cancel the current run")


def _print_help():
    print("Usage:")
    print(f"  python {DISPLAY_SCRIPT_NAME}")
    print(f"  python {DISPLAY_SCRIPT_NAME} <input_path>")
    print(f"  python {DISPLAY_SCRIPT_NAME} <input_path> <output_dir>")
    print(f"  python {DISPLAY_SCRIPT_NAME} [options] <input_path> [output_dir]")
    print("")
    print("Options:")
    print("  -h, --help          Show this help message")
    print("  --system-report     Show RAM, safe thread count, and model cache path")
    print("  --unsafe            Bypass safe-mode RAM and file-size guardrails")
    print("  --force-full-ocr    Force full-page OCR for image-like documents")
    print("  --pdf-text          Force the standard text-PDF path")
    print("  --pdf-ocr           Force the OCR path")
    print("  --pdf-mode MODE     Set PDF mode: auto, text, or ocr")
    print("")
    print("Tip:")
    print("  Press Ctrl+C during a run if you want to cancel.")
    print("")
    print("Examples:")
    print(f"  python {DISPLAY_SCRIPT_NAME}")
    print(rf"  python {DISPLAY_SCRIPT_NAME} report.pdf")
    print(rf"  python {DISPLAY_SCRIPT_NAME} report.pdf md_output")
    print(rf"  python {DISPLAY_SCRIPT_NAME} --pdf-mode ocr scanned.pdf")
    print(rf"  python {DISPLAY_SCRIPT_NAME} --system-report")


def _print_runtime_status(runtime):
    print(
        "  Safe mode: CPU-only, "
        f"{runtime['threads']} threads, "
        f"{runtime['available_ram_gb']:.1f} GB RAM free"
    )
    if runtime["cpu_warning"]:
        print(f"  CPU note : {runtime['cpu_warning']}")
    print("  Stop     : Press Ctrl+C to cancel if it is too slow")


def _parse_args(argv):
    args = {
        "input_path": None,
        "output_dir": OUTPUT_DIR,
        "pdf_mode": "auto",
        "unsafe": False,
        "force_full_ocr": False,
        "system_report": False,
        "help": False,
    }

    positionals = []
    i = 0
    while i < len(argv):
        token = argv[i]
        if token in {"-h", "--help"}:
            args["help"] = True
        elif token == "--unsafe":
            args["unsafe"] = True
        elif token == "--force-full-ocr":
            args["force_full_ocr"] = True
        elif token == "--pdf-text":
            args["pdf_mode"] = "text"
        elif token == "--pdf-ocr":
            args["pdf_mode"] = "ocr"
        elif token.startswith("--pdf-mode="):
            args["pdf_mode"] = token.split("=", 1)[1].strip().lower()
        elif token == "--pdf-mode":
            if i + 1 >= len(argv):
                raise ValueError("--pdf-mode requires one of: auto, text, ocr")
            args["pdf_mode"] = argv[i + 1].strip().lower()
            i += 1
        elif token == "--system-report":
            args["system_report"] = True
        else:
            positionals.append(token)
        i += 1

    if args["pdf_mode"] not in {"auto", "text", "ocr"}:
        raise ValueError("PDF mode must be one of: auto, text, ocr")

    if positionals:
        args["input_path"] = positionals[0]
    if len(positionals) > 1:
        args["output_dir"] = Path(positionals[1])

    return args


def main(argv=None):
    argv = sys.argv[1:] if argv is None else argv
    args = _parse_args(argv)

    if args["help"]:
        _print_help()
        return 0

    if args["system_report"]:
        _print_system_report()
        return 0

    batch_start = time.time()

    if args["input_path"]:
        result = convert_document_to_markdown(
            args["input_path"],
            args["output_dir"],
            unsafe=args["unsafe"],
            force_full_ocr=args["force_full_ocr"],
            pdf_mode=args["pdf_mode"],
        )
        return 0 if result else 1

    files = [file for file in SCRIPT_DIR.iterdir() if file.suffix.lower() in SUPPORTED_EXTENSIONS]
    if not files:
        print(f"No supported files found in {SCRIPT_DIR}")
        return 0

    print(f"\nFound {len(files)} file(s) to convert.")
    print("Press Ctrl+C at any time to stop the run.")
    ok, failed = 0, 0

    for i, file in enumerate(files):
        result = convert_document_to_markdown(
            file,
            OUTPUT_DIR,
            unsafe=args["unsafe"],
            force_full_ocr=args["force_full_ocr"],
            pdf_mode=args["pdf_mode"],
        )
        if result:
            ok += 1
        else:
            failed += 1

        if i < len(files) - 1:
            remaining = len(files) - i - 1
            print(f"  Next: {files[i + 1].name}  ({remaining} file(s) remaining)")

    total_elapsed = time.time() - batch_start
    mins, secs = divmod(int(total_elapsed), 60)
    time_str = f"{mins}m {secs}s" if mins else f"{secs:.1f}s"
    print(f"\n{'#' * 60}")
    print(f"  ALL DONE - {ok} converted, {failed} failed")
    print(f"  Total time: {time_str}")
    print(f"{'#' * 60}\n")
    return 1 if failed else 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except FileNotFoundError as exc:
        print(f"ERROR: {exc}")
        sys.exit(1)
    except KeyboardInterrupt:
        print("\nStopped by user (Ctrl+C).")
        sys.exit(130)
    except Exception as exc:
        print(f"ERROR: Unexpected failure: {exc}")
        sys.exit(1)
