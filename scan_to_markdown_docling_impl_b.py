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
