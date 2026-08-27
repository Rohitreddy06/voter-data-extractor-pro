"""
database.py
===========
SQLite persistence layer.

All extracted voter records are written to a single SQLite database so
that:
  * Excel matching never has to re-scan the PDF.
  * Ad-hoc searches (by EPIC / Name / House No) are fast even on
    500,000+ row rolls, thanks to an index on EPIC.
  * The extracted data can be exported/reused across sessions.

The schema deliberately keeps every field as TEXT (EPIC numbers and
serial numbers can have leading characters/zeros that must not be
mangled by numeric coercion).
"""

from __future__ import annotations

import sqlite3
from contextlib import contextmanager
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Iterable, Iterator, List, Optional, Sequence

import config
from utils import logger


@dataclass
class VoterRecord:
    """One row extracted from the electoral roll PDF."""

    epic: str
    name: str = ""
    relative: str = ""
    house_no: str = ""
    address: str = ""
    area: str = ""
    part_no: str = ""
    serial_no: str = ""
    page_no: int = 0
    age: str = ""
    gender: str = ""

    def as_tuple(self) -> tuple:
        return (
            self.epic, self.name, self.relative, self.house_no,
            self.address, self.area, self.part_no, self.serial_no,
            self.page_no, self.age, self.gender,
        )


SCHEMA = """
CREATE TABLE IF NOT EXISTS Voters (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    EPIC        TEXT NOT NULL,
    Name        TEXT,
    Relative    TEXT,
    HouseNo     TEXT,
    Address     TEXT,
    Area        TEXT,
    PartNo      TEXT,
    SerialNo    TEXT,
    PageNo      INTEGER,
    Age         TEXT,
    Gender      TEXT
);
"""

INDEX_SQL = "CREATE INDEX IF NOT EXISTS idx_voters_epic ON Voters (EPIC);"
INDEX_NAME_SQL = "CREATE INDEX IF NOT EXISTS idx_voters_name ON Voters (Name);"
INDEX_HOUSE_SQL = "CREATE INDEX IF NOT EXISTS idx_voters_house ON Voters (HouseNo);"


class VoterDatabase:
    """Thin, thread-safe-enough wrapper around the SQLite connection.

    A new connection is opened per call to `connection()` (SQLite
    connections are cheap and this avoids cross-thread sharing issues
    with the GUI's worker threads). Bulk inserts use a single
    connection + executemany for speed.
    """

    def __init__(self, db_path: Optional[Path] = None) -> None:
        self.db_path = Path(db_path) if db_path else config.DEFAULT_DB_PATH
        self._init_schema()

    @contextmanager
    def connection(self) -> Iterator[sqlite3.Connection]:
        conn = sqlite3.connect(str(self.db_path), timeout=30)
        conn.execute("PRAGMA journal_mode=WAL;")
        conn.execute("PRAGMA synchronous=NORMAL;")
        try:
            yield conn
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    def _init_schema(self) -> None:
        with self.connection() as conn:
            conn.execute(SCHEMA)
            conn.execute(INDEX_SQL)
            conn.execute(INDEX_NAME_SQL)
            conn.execute(INDEX_HOUSE_SQL)
        logger.info("Database ready at %s", self.db_path)

    def clear_all(self) -> None:
        with self.connection() as conn:
            conn.execute("DELETE FROM Voters;")
        logger.info("Cleared all rows from Voters table.")

    def insert_batch(self, records: Sequence[VoterRecord]) -> int:
        """Insert a batch of records. Returns number of rows inserted."""
        if not records:
            return 0
        rows = [r.as_tuple() for r in records]
        with self.connection() as conn:
            conn.executemany(
                """INSERT INTO Voters
                   (EPIC, Name, Relative, HouseNo, Address, Area,
                    PartNo, SerialNo, PageNo, Age, Gender)
                   VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
                rows,
            )
        return len(rows)

    def count(self) -> int:
        with self.connection() as conn:
            cur = conn.execute("SELECT COUNT(*) FROM Voters;")
            return cur.fetchone()[0]

    def find_by_epic(self, epic: str) -> List[sqlite3.Row]:
        """Return all rows matching an EPIC number (there may be more
        than one if the same EPIC appears twice in the roll — this is
        how duplicates are detected upstream in matcher.py)."""
        with self.connection() as conn:
            conn.row_factory = sqlite3.Row
            cur = conn.execute(
                "SELECT * FROM Voters WHERE EPIC = ? COLLATE NOCASE;", (epic.strip(),)
            )
            return cur.fetchall()

    def search_by_name(self, name_fragment: str, limit: int = 200) -> List[sqlite3.Row]:
        with self.connection() as conn:
            conn.row_factory = sqlite3.Row
            cur = conn.execute(
                "SELECT * FROM Voters WHERE Name LIKE ? COLLATE NOCASE LIMIT ?;",
                (f"%{name_fragment.strip()}%", limit),
            )
            return cur.fetchall()

    def search_by_house_no(self, house_no: str, limit: int = 200) -> List[sqlite3.Row]:
        with self.connection() as conn:
            conn.row_factory = sqlite3.Row
            cur = conn.execute(
                "SELECT * FROM Voters WHERE HouseNo LIKE ? COLLATE NOCASE LIMIT ?;",
                (f"%{house_no.strip()}%", limit),
            )
            return cur.fetchall()

    def duplicate_epics(self) -> List[sqlite3.Row]:
        with self.connection() as conn:
            conn.row_factory = sqlite3.Row
            cur = conn.execute(
                """SELECT EPIC, COUNT(*) as cnt FROM Voters
                   GROUP BY EPIC HAVING cnt > 1;"""
            )
            return cur.fetchall()

    def export_csv(self, csv_path: Path) -> None:
        import csv

        with self.connection() as conn:
            conn.row_factory = sqlite3.Row
            cur = conn.execute("SELECT * FROM Voters;")
            rows = cur.fetchall()

        if not rows:
            logger.warning("No rows to export to CSV.")
            return

        with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow(rows[0].keys())
            for row in rows:
                writer.writerow(tuple(row))
        logger.info("Exported %d rows to %s", len(rows), csv_path)

    def export_sqlite_copy(self, dest_path: Path) -> None:
        """Export/copy the live database file to another location
        (e.g. for archival or sharing) using SQLite's backup API so it
        works even while WAL journaling is active."""
        with self.connection() as src_conn:
            dest_conn = sqlite3.connect(str(dest_path))
            with dest_conn:
                src_conn.backup(dest_conn)
            dest_conn.close()
        logger.info("Exported database copy to %s", dest_path)
