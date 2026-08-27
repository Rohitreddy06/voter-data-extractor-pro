import os
from pathlib import Path
from database import VoterDatabase
from extractor import PDFExtractor
from excel_writer import run_matching_pipeline
import openpyxl

pdf_path = r'C:\Users\chenn\Downloads\Enumeration_Form_S29_45_59_S29_45_59_20260614094549_TEL - Copy.pdf'
if not os.path.exists(pdf_path):
    raise SystemExit('PDF missing')

# Prepare minimal Excel template
excel_path = Path('template_test.xlsx')
if not excel_path.exists():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(['Epic Number','House No','Colony'])
    ws.append(['NVT8141996','',''])
    wb.save(excel_path)
    print('Created', excel_path)

# Clear DB and run extraction
db = VoterDatabase()
try:
    db.clear_all()
except Exception as e:
    print('DB clear failed:', e)

extractor = PDFExtractor(pdf_path, db, use_ocr=False, translate_telugu=False)
try:
    summary = extractor.run()
    print('Extraction summary:', summary)
except Exception as e:
    print('Extraction halted with error:', type(e).__name__, e)
    raise

# Run matching pipeline
out_path, stats = run_matching_pipeline(str(excel_path), db, 'output')
print('Saved matched workbook to', out_path)
print('Match stats:', stats)
