"""
ocr.py
======
OCR fallback for scanned (image-only) electoral roll pages.

Uses pdf2image (Poppler) to rasterize a page to an image, then
pytesseract (Tesseract OCR, English + Telugu language packs) to pull
text back out. This is only invoked for pages where PyMuPDF/pdfplumber
found effectively no extractable text — see
extractor.PDFExtractor._is_scanned_page().

Both Tesseract and Poppler are external system binaries and must be
installed separately (see requirements.txt header comment). If they
are missing, OCR methods raise a clear RuntimeError rather than a
cryptic ImportError deep in a worker thread.
"""

from __future__ import annotations

from typing import Optional

import config
from utils import logger

try:
    import pytesseract
    from pdf2image import convert_from_path
    from PIL import Image
    _OCR_AVAILABLE = True
except ImportError as exc:  # pragma: no cover - environment dependent
    _OCR_AVAILABLE = False
    _OCR_IMPORT_ERROR = exc


class OCRError(RuntimeError):
    """Raised when OCR cannot be performed (missing binaries, bad page, etc.)."""


def ocr_available() -> bool:
    return _OCR_AVAILABLE


def ocr_page_from_pdf(pdf_path: str, page_number: int, dpi: int = config.OCR_DPI) -> str:
    """Render a single 1-indexed PDF page to an image and OCR it.

    Raises OCRError if pdf2image/pytesseract/Tesseract/Poppler are not
    available, so callers can log a clear, actionable failure instead
    of crashing the worker thread.
    """
    if not _OCR_AVAILABLE:
        raise OCRError(
            "OCR dependencies not available "
            f"(pytesseract/pdf2image/Pillow). Original import error: {_OCR_IMPORT_ERROR}"
        )

    try:
        images = convert_from_path(
            pdf_path, dpi=dpi, first_page=page_number, last_page=page_number
        )
    except Exception as exc:
        raise OCRError(f"Failed to rasterize page {page_number} via Poppler: {exc}") from exc

    if not images:
        raise OCRError(f"No image produced for page {page_number}.")

    return ocr_image(images[0])


def ocr_image(image: "Image.Image") -> str:
    if not _OCR_AVAILABLE:
        raise OCRError(
            "OCR dependencies not available "
            f"(pytesseract/pdf2image/Pillow). Original import error: {_OCR_IMPORT_ERROR}"
        )
    try:
        text = pytesseract.image_to_string(image, lang=config.TESSERACT_LANGS)
    except pytesseract.TesseractNotFoundError as exc:
        raise OCRError(
            "Tesseract OCR engine executable not found on PATH. "
            "Install it from https://github.com/UB-Mannheim/tesseract/wiki "
            "and ensure tesseract.exe is on PATH."
        ) from exc
    except Exception as exc:
        raise OCRError(f"Tesseract OCR failed: {exc}") from exc
    return text


def safe_ocr_page(pdf_path: str, page_number: int) -> str:
    """Never-raise wrapper used by the extractor pipeline: on any OCR
    failure, logs it and returns an empty string so extraction can
    continue with the remaining pages instead of aborting the run."""
    try:
        return ocr_page_from_pdf(pdf_path, page_number)
    except OCRError as exc:
        logger.error("OCR failed on page %d: %s", page_number, exc)
        return ""
