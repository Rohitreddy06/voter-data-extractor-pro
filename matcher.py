"""
matcher.py
==========
Matches EPIC numbers from the user's Excel template against the
SQLite voter database built by extractor.py, and prepares the values
to write back (House No / Colony), plus statistics (matched,
not-found, duplicate EPICs).

This module does not touch openpyxl directly — it hands a plain list
of MatchResult objects to excel_writer.py, keeping "decide what to
write" separate from "how to write an .xlsx".
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Callable, List, Optional

import config
from database import VoterDatabase
from utils import logger, normalize_epic, safe_str


@dataclass
class MatchResult:
    row_index: int          # 1-indexed Excel data row (excluding header)
    epic: str
    found: bool
    duplicate: bool
    house_no: str = ""
    colony: str = ""
    address: str = ""


@dataclass
class MatchStats:
    total: int = 0
    matched: int = 0
    not_found: int = 0
    duplicates: int = 0
    blank_epic: int = 0


class Matcher:
    def __init__(self, db: VoterDatabase) -> None:
        self.db = db

    def match_epics(
        self,
        epic_rows: List[tuple[int, str]],
        progress_callback: Optional[Callable[[int, int], None]] = None,
    ) -> tuple[List[MatchResult], MatchStats]:
        """`epic_rows` is a list of (excel_row_index, epic_value) pairs
        pulled from the Excel sheet by excel_writer.py. Returns per-row
        match results plus aggregate stats."""
        results: List[MatchResult] = []
        stats = MatchStats(total=len(epic_rows))
        progress_callback = progress_callback or (lambda cur, total: None)

        for i, (row_idx, raw_epic) in enumerate(epic_rows, start=1):
            epic = normalize_epic(raw_epic)

            if not epic:
                stats.blank_epic += 1
                results.append(MatchResult(row_idx, epic, found=False, duplicate=False))
                progress_callback(i, stats.total)
                continue

            db_rows = self.db.find_by_epic(epic)

            if not db_rows:
                stats.not_found += 1
                results.append(MatchResult(row_idx, epic, found=False, duplicate=False))
                logger.info("EPIC not found: %s (row %d)", epic, row_idx)
            else:
                is_dup = len(db_rows) > 1
                if is_dup:
                    stats.duplicates += 1
                    logger.warning(
                        "Duplicate EPIC %s found %d times in electoral roll "
                        "(row %d in Excel). Using first match.",
                        epic, len(db_rows), row_idx,
                    )
                best = db_rows[0]
                stats.matched += 1
                results.append(
                    MatchResult(
                        row_idx, epic, found=True, duplicate=is_dup,
                        house_no=safe_str(best["HouseNo"]),
                        colony=safe_str(best["Area"]),
                        address=safe_str(best["Address"]),
                    )
                )

            progress_callback(i, stats.total)

        logger.info(
            "Matching complete: %d total, %d matched, %d not found, "
            "%d duplicates, %d blank EPICs",
            stats.total, stats.matched, stats.not_found, stats.duplicates, stats.blank_epic,
        )
        return results, stats
