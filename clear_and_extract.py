from database import VoterDatabase
from extractor import PDFExtractor
import os

pdf_path = r'C:\Users\chenn\Downloads\Enumeration_Form_S29_45_59_S29_45_59_20260614094549_TEL - Copy.pdf'
if not os.path.exists(pdf_path):
    raise SystemExit('PDF missing')

print('PDF:', pdf_path)
db = VoterDatabase()
print('DB before clear count:', db.count())
db.clear_all()
print('DB after clear count:', db.count())
extractor = PDFExtractor(pdf_path, db, use_ocr=False, translate_telugu=False)
try:
    summary = extractor.run()
    print('Extraction summary:', summary)
except Exception as e:
    print('Extraction halted with error:', type(e).__name__, e)
    raise

print('DB after extraction count:', db.count())
rows = db.find_by_epic('NVT8141996')
print('Rows for EPIC:', len(rows))
for r in rows:
    print(dict(r))
