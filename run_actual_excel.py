from pathlib import Path
import os
from database import VoterDatabase
from extractor import PDFExtractor
from excel_writer import run_matching_pipeline

pdf_path = r'C:\Users\chenn\Downloads\Enumeration_Form_S29_45_59_S29_45_59_20260614094549_TEL - Copy.pdf'
excel_path = r'C:\Users\chenn\Downloads\59_EF_Filled_Forms_Complete_1-8-2026.xlsx'
output_folder = r'C:\Users\chenn\Downloads\VoterDataExtractor\output'

for path in [pdf_path, excel_path]:
    if not Path(path).exists():
        raise SystemExit(f'Missing required file: {path}')

Path(output_folder).mkdir(parents=True, exist_ok=True)

db = VoterDatabase()
print('Clearing existing DB rows...')
db.clear_all()

print('Extracting PDF...')
extractor = PDFExtractor(pdf_path, db, use_ocr=False, translate_telugu=False)
summary = extractor.run()
print('Extraction summary:', summary)

print('Matching Excel...')
out_path, stats = run_matching_pipeline(excel_path, db, output_folder)
print('Saved completed workbook to', out_path)
print('Match stats:', stats)
