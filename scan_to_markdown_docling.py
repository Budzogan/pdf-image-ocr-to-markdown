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
