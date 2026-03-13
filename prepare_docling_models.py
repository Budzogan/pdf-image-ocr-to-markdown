import logging
import sys
from pathlib import Path

from docling.datamodel.pipeline_options import LayoutOptions
from docling.models.stages.layout.layout_model import LayoutModel
from docling.models.stages.ocr.rapid_ocr_model import RapidOcrModel
from docling.models.stages.table_structure.table_structure_model import TableStructureModel


TABLEFORMER_DIRNAME = "docling-project--TableFormerV2"


def _has_cached_files(path):
    return path.exists() and any(candidate.is_file() for candidate in path.rglob("*"))


def _quiet_third_party_logs():
    for logger_name in ("RapidOCR", "onnxruntime", "docling"):
        logging.getLogger(logger_name).setLevel(logging.WARNING)


def _ensure_component(label, local_dir, downloader, extra_paths=None):
    paths_to_check = [local_dir, *(extra_paths or [])]
    ready = all(_has_cached_files(path) for path in paths_to_check)

    if ready:
        print(f"[ready] {label}")
        return

    print(f"[setup] {label} missing. Downloading now...")
    downloader()
    if all(_has_cached_files(path) for path in paths_to_check):
        print(f"[ok]    {label} ready")
    else:
        raise RuntimeError(f"{label} is still missing after download.")


def main():
    model_dir = Path(sys.argv[1]).expanduser().resolve() if len(sys.argv) > 1 else Path.cwd() / "docling_models"
    model_dir.mkdir(parents=True, exist_ok=True)
    _quiet_third_party_logs()

    print(f"Preparing Docling model cache in {model_dir}")
    print("Checking components: layout, tableformer, rapidocr")

    layout_dir = model_dir / LayoutOptions().model_spec.model_repo_folder
    table_dir = model_dir / TableStructureModel._model_repo_folder
    tableformer_dir = model_dir / TABLEFORMER_DIRNAME
    rapidocr_dir = model_dir / RapidOcrModel._model_repo_folder

    _ensure_component(
        "Layout model",
        layout_dir,
        lambda: LayoutModel.download_models(
            local_dir=layout_dir,
            force=False,
            progress=False,
        ),
    )
    _ensure_component(
        "Table structure model",
        table_dir,
        lambda: TableStructureModel.download_models(
            local_dir=table_dir,
            force=False,
            progress=False,
        ),
        extra_paths=[tableformer_dir],
    )
    _ensure_component(
        "RapidOCR model",
        rapidocr_dir,
        lambda: RapidOcrModel.download_models(
            backend="onnxruntime",
            local_dir=rapidocr_dir,
            force=False,
            progress=False,
        ),
    )

    print("Docling models are ready.")


if __name__ == "__main__":
    main()
