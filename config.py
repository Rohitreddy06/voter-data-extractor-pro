"""
config.py
=========
Central configuration for Voter Data Extractor Pro.

Holds file paths, regex patterns used for parsing Telangana / Andhra
Pradesh electoral roll PDFs, GUI constants, and tunable performance
knobs. Keeping these in one module means the rest of the codebase
never hard-codes a magic string or number.

NOTE ON REGEX PATTERNS
-----------------------
Electoral roll layouts differ slightly between publication years and
between TS/AP CEO office templates. The patterns below cover the most
common layout (EPIC number in the format AAA1234567, followed by
Name / Relative Name / House No / Age / Gender fields). If your PDFs
use a different template, adjust EPIC_PATTERN and the field labels in
extractor.py's LABEL_* constants — everything else in the pipeline is
layout-agnostic.
"""

from __future__ import annotations

import os
import re
from pathlib import Path

# --------------------------------------------------------------------------
# Directories
# --------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
LOGS_DIR = BASE_DIR / "logs"
OUTPUT_DIR = BASE_DIR / "output"
DATABASE_DIR = BASE_DIR / "database"

for _dir in (LOGS_DIR, OUTPUT_DIR, DATABASE_DIR):
    _dir.mkdir(parents=True, exist_ok=True)

LOG_FILE = LOGS_DIR / "processing.log"
DEFAULT_DB_PATH = DATABASE_DIR / "voters.db"
RESUME_STATE_FILE = DATABASE_DIR / "resume_state.json"
LAST_FOLDER_FILE = DATABASE_DIR / "last_folder.json"

# --------------------------------------------------------------------------
# EPIC number pattern
# --------------------------------------------------------------------------
# Standard Indian EPIC format: 3 letters followed by 7 digits, e.g. ABC1234567.
# Some OCR engines insert a stray space or hyphen between the letters and
# digits, so we tolerate that here. The pattern is also case-insensitive
# so lowercase OCR output is still matched.
EPIC_PATTERN = re.compile(
    r"\b([A-Z]{3}[-\s]*[0-9]{7}|[A-Z]{2,3}/?\d{2,3}/?\d{3,6})\b",
    re.IGNORECASE,
)

# --------------------------------------------------------------------------
# OCR / language settings
# --------------------------------------------------------------------------
TESSERACT_LANGS = "eng+tel"          # English + Telugu
OCR_DPI = 300                        # render resolution for scanned pages
TEXT_EXTRACTION_MIN_CHARS = 40       # below this, a page is considered "scanned"

# --------------------------------------------------------------------------
# Performance / concurrency
# --------------------------------------------------------------------------
MAX_WORKER_THREADS = min(8, (os.cpu_count() or 4))
DB_BATCH_SIZE = 500                  # rows per executemany() batch
PROGRESS_UPDATE_EVERY_N_PAGES = 1

# --------------------------------------------------------------------------
# Excel columns (must match the template's header row, case-insensitive)
# --------------------------------------------------------------------------
COL_SNO = "S.No"
COL_EPIC = "Epic Number"
COL_NAME = "Elector Name"
COL_RELATIVE = "Relative Name"
COL_AGE = "Age"
COL_GENDER = "Gender"
COL_HOUSE_NO = "House No"
COL_COLONY = "Colony"
COL_ADDRESS = "Address"
COL_PHONE = "Phone Number"

REQUIRED_EXCEL_COLUMNS = [COL_EPIC, COL_HOUSE_NO, COL_COLONY]

NOT_FOUND_TEXT = "NOT FOUND"

# --------------------------------------------------------------------------
# GUI
# --------------------------------------------------------------------------
APP_TITLE = "Voter Data Extractor Pro"
APP_GEOMETRY = "1000x720"
THEME_MODE = "dark"
THEME_COLOR = "blue"

# --------------------------------------------------------------------------
# Address labels to strip out while isolating House No / Colony
# --------------------------------------------------------------------------
ADDRESS_NOISE_LABELS = [
    "polling station", "assembly constituency", "part no", "part number",
    "serial no", "photo", "booth", "page", "electoral roll",
    "constituency", "ward no", "gram panchayat", "mandal",
]
