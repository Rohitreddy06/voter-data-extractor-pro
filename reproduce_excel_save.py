from pathlib import Path
from database import VoterDatabase
from excel_writer import ExcelWriter
from matcher import Matcher

excel_path = Path(r'C:\Users\chenn\Downloads\59_EF_Filled_Forms_Complete_1-8-2026.xlsx')
output_folder = Path(r'C:\Users\chenn\Downloads\VoterDataExtractor')

db = VoterDatabase()
writer = ExcelWriter(str(excel_path))
writer.load()
epic_rows = writer.read_epic_rows()
matcher = Matcher(db)
results, stats = matcher.match_epics(epic_rows)
print('matched', stats.matched, 'duplicates', stats.duplicates, 'not_found', stats.not_found)
try:
    writer.apply_results(results)
    writer.save(str(output_folder), filename='reproduce_save.xlsx')
    print('save succeeded')
except Exception as exc:
    print('save failed', type(exc).__name__, exc)
    raise
