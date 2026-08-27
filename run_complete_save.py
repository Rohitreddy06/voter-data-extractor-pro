from pathlib import Path
from database import VoterDatabase
from excel_writer import run_matching_pipeline, ExcelWriter

excel_path = Path(r'C:\Users\chenn\Downloads\59_EF_Filled_Forms_Complete_1-8-2026.xlsx')
output_folder = Path(r'C:\Users\chenn\Downloads\VoterDataExtractor')

print('Running matching pipeline...')
out_path, stats = run_matching_pipeline(str(excel_path), VoterDatabase(), str(output_folder))
print('Pipeline saved to', out_path)

# Load saved workbook and print first 20 H.No/Colony values for verification
writer = ExcelWriter(out_path)
writer.load()
rows = writer.read_epic_rows()
house_col = writer._col('H.No/Flat.No') or writer._col('H.No') or writer._col('house_no')
colony_col = writer._col('Colony') or writer._col('Area') or writer._col('colony')
print('\nFirst 20 filled H.No/Colony from saved workbook:')
for row_idx, _ in rows[:20]:
    h = writer.worksheet.cell(row=row_idx, column=house_col).value if house_col else None
    c = writer.worksheet.cell(row=row_idx, column=colony_col).value if colony_col else None
    print() 
    print('Row', row_idx)
    print('H.No/Flat.No ->', h)
    print('Colony      ->', c)

print('\nStats:', stats)
