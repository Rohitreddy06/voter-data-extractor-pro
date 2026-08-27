# Voter Data Extractor Pro

A Windows desktop application that extracts voter records from
Telangana / Andhra Pradesh Electoral Roll PDFs and fills an Excel
template's **House No** and **Colony** columns by matching **EPIC
numbers**.

## Features

- Auto-detects text-based vs. scanned PDF pages; OCRs scanned pages
  (English + Telugu) with Tesseract.
- Parses each page into individual voter records (EPIC, Name,
  Relative Name, House No, Address/Area, Part No, Serial No, Age,
  Gender) and stores them in an indexed SQLite database, so the PDF
  is only ever read once — later searches hit the database.
- Matches every EPIC number in your Excel template against the
  database and fills **House No** / **Colony**, writing `NOT FOUND`
  where no match exists and flagging duplicate EPICs.
- Optional Telugu → English **transliteration** (not translation) of
  addresses, e.g. `కొంపల్లి` → `Kompally`.
- Excel output preserves the original template's formatting, colors,
  and formulas — only the target cells are updated.
- Dark-themed CustomTkinter GUI: file pickers, drag & drop, live
  progress bar, elapsed time, records-found / matched / remaining
  counters, and a live log.
- Multithreaded so the GUI never freezes during long PDF runs, with
  resume-after-interruption, CSV/DB export, and search-by-EPIC /
  Name / House No.

## Installation

1. Install Python 3.12+.
2. Install system dependencies (NOT via pip):
   - **Tesseract OCR** — https://github.com/UB-Mannheim/tesseract/wiki
     Install the English **and Telugu** language packs, add
     `tesseract.exe` to your PATH.
   - **Poppler** (for rendering scanned PDF pages) —
     https://github.com/oschwartz10612/poppler-windows — add its
     `bin` folder to PATH.
3. Install Python dependencies:
   ```
   pip install -r requirements.txt
   ```

## Running

```
python main.py
```

1. Click **Select PDF** and choose the electoral roll PDF.
2. Click **Select Excel** and choose your template (must contain at
   least `Epic Number`, `House No`, and `Colony` columns).
3. Choose an **Output Folder**.
4. Toggle **Translate Telugu → English** and **Use OCR** as needed.
5. Click **Start**. Progress, elapsed time, and live logs are shown
   as the PDF is processed. `Completed.xlsx` is written to the
   output folder when finished.

If a run is interrupted (crash, cancel, power loss), starting again
on the same PDF will offer to **resume** from the last completed
page instead of reprocessing from scratch.

## Screenshots

### Main interface

![Main interface](screenshots/main-interface.png)

### PDF processing

![PDF processing](screenshots/pdf-processing.png)

### Excel matching

![Excel matching](screenshots/excel-matching.png)

### Search results

![Search results](screenshots/search-results.png)

## Data privacy

Input PDFs, Excel templates, generated workbooks, CSV exports, SQLite
databases, resume state, and processing logs may contain personal voter
data. Keep these files local and do not commit or publish them. The
repository `.gitignore` excludes generated files in `database/`, `logs/`,
and `output/`, as well as common voter-data formats. Review `git status`
before pushing changes.

## Adjusting to your PDF's exact layout

Electoral roll PDF layouts vary slightly by year and by which CEO
office template was used. The field-extraction regular expressions
live in `extractor.py` (`_LABEL_NAME`, `_LABEL_RELATIVE`,
`_LABEL_HOUSE`, etc.) and the EPIC number pattern lives in
`config.py` (`EPIC_PATTERN`). If your PDFs use different field
labels or an EPIC format outside `AAA1234567`, adjust those patterns
— the rest of the pipeline (database, matching, Excel writing, GUI)
is layout-agnostic and needs no changes.

## Project layout

```
VoterDataExtractor/
├── main.py           # entry point
├── gui.py             # CustomTkinter GUI + threading
├── extractor.py        # PDF parsing (text + OCR) into VoterRecords
├── ocr.py               # Tesseract OCR fallback for scanned pages
├── database.py           # SQLite schema, batch insert, search, export
├── matcher.py              # EPIC matching logic against the database
├── excel_writer.py          # openpyxl read/write, preserves formatting
├── translator.py              # Telugu -> English transliteration
├── utils.py                    # logging, resume state, stopwatch
├── config.py                    # paths, regexes, constants
├── requirements.txt
├── logs/               # local processing.log (ignored)
├── output/             # local Completed.xlsx / exports (ignored)
└── database/           # local voters.db + resume state (ignored)
```

## Notes / limitations

- This tool only works with **unencrypted** PDFs. Password-protected
  files will raise a clear error asking for an unlocked copy.
- OCR accuracy on scanned rolls depends heavily on scan quality; review
  the local `logs/processing.log` file for pages that produced no records.
- Always verify a sample of matched rows in `Completed.xlsx` against
  the source PDF before relying on the output for official use.
