"""Second half of converter implementation."""
from scan_to_markdown_docling_impl_a import *  # noqa: F403

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

        md_path = target_dir / f"{source.stem}.md"
        result.document.save_as_markdown(
            md_path,
            artifacts_dir=images_dir,
            image_mode=ImageRefMode.REFERENCED,
        )

        if not any(images_dir.iterdir()):
            try:
                images_dir.rmdir()
            except OSError:
                pass

        return md_path.read_text(encoding="utf-8")
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
