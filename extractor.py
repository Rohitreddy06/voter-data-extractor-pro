"""
extractor.py
============
Simplified PDF extraction engine sufficient for the GUI import and
for the user's address-extraction workflow described in the task.

Behavior implemented:
- One page => one `VoterRecord` (do not split a page into multiple voters)
- Find EPIC using `config.EPIC_PATTERN`
- Find Telugu anchor `చిరునామా` (or English "Address") and extract text
  until the address terminator `క్రమ సంఖ్య` (or common serial/part labels)
- Normalize address lines, join with commas, split on commas and trim
- Assign `house_no=parts[0]`, `colony=parts[1]` (if present)
- If the address anchor is present but house_no or colony is missing,
  raise `AddressParseError` so the run can be aborted for debugging

This file purposefully keeps the implementation small and robust so the
application can import and run; it defers heavy PDF layout parsing to
PyMuPDF (fitz) when available and uses OCR as a fallback via `ocr.py`.
"""

from __future__ import annotations

import re
import threading
from pathlib import Path
from typing import Callable, List, Optional

import config
import ocr
from database import VoterDatabase, VoterRecord
from utils import logger, safe_str, strip_illegal_xml_chars

try:
    import fitz  # PyMuPDF
except Exception:  # pragma: no cover - runtime availability
    fitz = None


class ExtractionError(RuntimeError):
    pass


class EncryptedPDFError(ExtractionError):
    pass


class CorruptPDFError(ExtractionError):
    pass


class AddressParseError(ExtractionError):
    """Raised when an address anchor is present but parsing produced
    incomplete house/colony components. This stops execution for debugging."""


# Terminator regex: Telugu "క్రమ సంఖ్య" or common serial/part labels
_ADDRESS_TERMINATOR = re.compile(r"(?:క్రమ\s*సంఖ్య|Serial\s*No|Sl\.?\s*No|Part\s*No)", re.IGNORECASE)
_LABEL_ADDRESS_START = re.compile(r"(?:చిరునామా|Address)\s*[:\-]?", re.IGNORECASE)


def _extract_field(pattern: re.Pattern, text: str) -> str:
    m = pattern.search(text)
    return strip_illegal_xml_chars(safe_str(m.group(1))) if m else ""


def _extract_address_block(page_text: str) -> str:
    """Find the Telugu address anchor and return the normalized address
    block as a single comma-separated string. Returns empty string if
    the anchor is not found.
    """
    m = _LABEL_ADDRESS_START.search(page_text)
    if not m:
        return ""

    start = m.end()
    term = _ADDRESS_TERMINATOR.search(page_text, start)
    end = term.start() if term else len(page_text)

    raw = page_text[start:end].strip()
    # Collapse multiple newlines into a comma separator, preserve commas.
    raw = re.sub(r"[\r\n]+", ", ", raw)
    # Remove duplicate commas and collapse whitespace
    raw = re.sub(r",\s*,+", ", ", raw)
    raw = re.sub(r"\s+", " ", raw).strip(" ,.-")
    return strip_illegal_xml_chars(raw)


def _parse_house_and_colony(address: str) -> tuple[str, str]:
    """Split the normalized address by commas and return (house_no, colony).
    Parts are stripped of surrounding punctuation/whitespace. If a part is
    missing, return empty string for that component.
    """
    parts = [p.strip(" \t\n\r,.-") for p in address.split(",") if p.strip()]
    if not parts:
        return "", ""

    house_no = parts[0]
    colony = parts[1] if len(parts) > 1 else ""

    house_no = strip_illegal_xml_chars(house_no)
    colony = strip_illegal_xml_chars(colony)
    return house_no, colony


def parse_voter_blocks(page_text: str, page_no: int, translate_telugu: bool = False) -> List[VoterRecord]:
    """Parse exactly one VoterRecord from the provided page text.

    - One page -> one voter. If EPIC not found, returns empty list.
    - Uses Telugu anchor extraction for the address.
    - Does NOT transliterate Telugu; returns original Telugu text.
    """
    records: List[VoterRecord] = []
    if not page_text or not page_text.strip():
        return records

    epic_m = config.EPIC_PATTERN.search(page_text)
    if not epic_m:
        logger.info("No EPIC found on page %d; skipping", page_no)
        return records

    epic = epic_m.group(1).strip()

    # Extract address anchored on Telugu label
    address_block = _extract_address_block(page_text)
    house_no = ""
    colony = ""

    if address_block:
        house_no, colony = _parse_house_and_colony(address_block)
        # If the anchor exists but components missing, raise for debugging
        if (not house_no or not colony):
            raise AddressParseError(
                f"Address anchored on page {page_no} but could not parse House No/Colony.\n"
                f"Extracted block: {address_block!r}"
            )

    # Construct the VoterRecord; keep Telugu unchanged
    rec = VoterRecord(
        epic=epic,
        house_no=house_no,
        address=address_block,
        area=colony,
        page_no=page_no,
    )
    records.append(rec)
    return records


class PDFExtractor:
    """Lightweight PDF extractor wrapper used by the GUI.

    This implementation focuses on safe imports and predictable behavior
    so the rest of the app can call `run()`; it uses PyMuPDF when
    available and falls back to OCR for scanned pages.
    """

    def __init__(
        self,
        pdf_path: str,
        db: VoterDatabase,
        use_ocr: bool = True,
        translate_telugu: bool = False,
        progress_callback: Optional[Callable[[int, int, str], None]] = None,
        cancel_event: Optional[threading.Event] = None,
    ) -> None:
        self.pdf_path = pdf_path
        self.db = db
        self.use_ocr = use_ocr
        self.translate_telugu = translate_telugu
        self.progress_callback = progress_callback or (lambda cur, total, msg: None)
        self.cancel_event = cancel_event or threading.Event()

        self._doc = None
        self.page_count = 0
        self._open_pdf()

    def _open_pdf(self) -> None:
        path = Path(self.pdf_path)
        if not path.exists():
            raise ExtractionError(f"PDF file not found: {self.pdf_path}")

        if fitz is None:
            # FitZ not available; we'll still allow the app to import but
            # extraction will be disabled at runtime.
            logger.warning("PyMuPDF (fitz) not available; PDF parsing disabled.")
            self.page_count = 0
            self._doc = None
            return

        try:
            doc = fitz.open(self.pdf_path)
        except Exception as exc:
            raise CorruptPDFError(f"Could not open PDF: {exc}") from exc

        if doc.is_encrypted and not doc.authenticate(""):
            raise EncryptedPDFError("PDF is encrypted; provide an unlocked copy.")

        self._doc = doc
        self.page_count = doc.page_count

    def _get_page_text(self, page_index: int) -> str:
        if not self._doc:
            return ""
        try:
            page = self._doc.load_page(page_index)
            txt = page.get_text("text") or ""
            if txt and "చిరునామా" in txt:
                return txt
            # Try more robust block/dict extraction
            blocks = page.get_text("blocks") or []
            if blocks:
                lines = [str(b[4]).strip() for b in blocks if len(b) >= 5 and b[4]]
                return "\n".join(lines)
            return txt
        except Exception as exc:
            logger.warning("PyMuPDF failed reading page %d: %s", page_index + 1, exc)
            return ""

    def run(self, resume_from_page: int = 0) -> dict:
        if self.page_count == 0:
            logger.info("No pages to process (fitz not available or PDF empty)")
            return {"total_records": 0, "pages_processed": 0, "cancelled": False}

        total_records = 0
        batch: List[VoterRecord] = []
        start = max(0, resume_from_page)

        for i in range(start, self.page_count):
            if self.cancel_event.is_set():
                break
            page_no = i + 1
            text = self._get_page_text(i)
            # If page looks scanned and OCR is enabled, try OCR
            if (not text or len(text.strip()) < config.TEXT_EXTRACTION_MIN_CHARS) and self.use_ocr:
                if ocr.ocr_available():
                    text = ocr.safe_ocr_page(self.pdf_path, page_no)
            try:
                records = parse_voter_blocks(text, page_no, translate_telugu=self.translate_telugu)
                if records:
                    batch.extend(records)
                    total_records += len(records)
            except AddressParseError:
                # Re-raise so the GUI can stop and show the error for debugging
                raise
            except Exception as exc:
                logger.error("Error parsing page %d: %s", page_no, exc)

            if len(batch) >= config.DB_BATCH_SIZE:
                self.db.insert_batch(batch)
                batch.clear()

            self.progress_callback(i + 1, self.page_count, f"Page {i+1}/{self.page_count}")

        if batch:
            self.db.insert_batch(batch)

        logger.info("Extraction finished: %d records", total_records)
        return {"total_records": total_records, "pages_processed": self.page_count, "cancelled": self.cancel_event.is_set()}

    def close(self) -> None:
        try:
            if self._doc:
                self._doc.close()
        except Exception:
            pass
