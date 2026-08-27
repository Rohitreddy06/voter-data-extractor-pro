"""
utils.py
========
Shared helpers used across the project: logging setup, JSON-based
resume/state persistence, "last folder" memory, and small formatting
utilities. Nothing in this module depends on the GUI or on any other
project module, so it can be imported anywhere without circular
import issues.
"""

from __future__ import annotations

import json
import logging
import re
import time
from pathlib import Path
from typing import Any, Optional

import config


def setup_logger(name: str = "voter_extractor") -> logging.Logger:
    """Create (or fetch) the shared application logger.

    Writes to logs/processing.log and mirrors to console. Safe to call
    repeatedly (e.g. once per module) since Python's logging module
    de-duplicates handlers per logger name only if we guard for it.
    """
    logger = logging.getLogger(name)
    if logger.handlers:
        return logger

    logger.setLevel(logging.DEBUG)

    file_handler = logging.FileHandler(config.LOG_FILE, encoding="utf-8")
    file_handler.setLevel(logging.DEBUG)

    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)

    fmt = logging.Formatter(
        "%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    file_handler.setFormatter(fmt)
    console_handler.setFormatter(fmt)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    return logger


logger = setup_logger()


class Stopwatch:
    """Small helper to track elapsed processing time for the GUI."""

    def __init__(self) -> None:
        self._start: Optional[float] = None
        self._elapsed_at_pause: float = 0.0

    def start(self) -> None:
        self._start = time.monotonic()
        self._elapsed_at_pause = 0.0

    def elapsed(self) -> float:
        if self._start is None:
            return self._elapsed_at_pause
        return self._elapsed_at_pause + (time.monotonic() - self._start)

    def elapsed_str(self) -> str:
        total = int(self.elapsed())
        h, rem = divmod(total, 3600)
        m, s = divmod(rem, 60)
        if h:
            return f"{h:02d}:{m:02d}:{s:02d}"
        return f"{m:02d}:{s:02d}"

    def stop(self) -> None:
        self._elapsed_at_pause = self.elapsed()
        self._start = None


def save_json(path: Path, data: dict) -> None:
    try:
        path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    except OSError as exc:
        logger.error("Failed to write JSON state file %s: %s", path, exc)


def load_json(path: Path) -> Optional[dict]:
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        logger.error("Failed to read JSON state file %s: %s", path, exc)
        return None


def save_resume_state(pdf_path: str, last_page_completed: int, db_path: str) -> None:
    """Persist how far page-by-page extraction has progressed so a future
    run can skip already-processed pages after an interruption/crash."""
    save_json(
        config.RESUME_STATE_FILE,
        {
            "pdf_path": pdf_path,
            "last_page_completed": last_page_completed,
            "db_path": db_path,
            "timestamp": time.time(),
        },
    )


def load_resume_state(pdf_path: str) -> Optional[dict]:
    state = load_json(config.RESUME_STATE_FILE)
    if state and state.get("pdf_path") == pdf_path:
        return state
    return None


def clear_resume_state() -> None:
    try:
        if config.RESUME_STATE_FILE.exists():
            config.RESUME_STATE_FILE.unlink()
    except OSError as exc:
        logger.error("Failed to clear resume state: %s", exc)


def remember_last_folder(folder: str) -> None:
    save_json(config.LAST_FOLDER_FILE, {"last_folder": folder})


def get_last_folder() -> str:
    state = load_json(config.LAST_FOLDER_FILE)
    if state and "last_folder" in state:
        return state["last_folder"]
    return str(Path.home())


def _is_xml_char(codepoint: int) -> bool:
    return (
        codepoint == 0x9
        or codepoint == 0xA
        or codepoint == 0xD
        or 0x20 <= codepoint <= 0xD7FF
        or 0xE000 <= codepoint <= 0xFFFD
        or 0x10000 <= codepoint <= 0x10FFFF
    )


def strip_illegal_xml_chars(text: str) -> str:
    if not text:
        return ""
    return "".join(ch for ch in text if _is_xml_char(ord(ch)))


def safe_str(value: Any) -> str:
    """Normalize any extracted value into a clean, whitespace-collapsed
    string, never raising even if given None or a non-str type."""
    if value is None:
        return ""
    text = strip_illegal_xml_chars(str(value))
    return " ".join(text.split())


def normalize_epic(value: Any) -> str:
    """Normalize EPIC values by stripping whitespace/hyphens and uppercasing."""
    text = safe_str(value)
    return re.sub(r"[\s\-]+", "", text).upper()


def format_bytes(num_bytes: int) -> str:
    step = 1024.0
    for unit in ("B", "KB", "MB", "GB"):
        if num_bytes < step:
            return f"{num_bytes:.1f} {unit}"
        num_bytes /= step
    return f"{num_bytes:.1f} TB"
