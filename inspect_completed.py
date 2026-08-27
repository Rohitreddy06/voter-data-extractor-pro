from pathlib import Path
from excel_writer import ExcelWriter
import config

path = Path('Completed.xlsx')
if not path.exists():
    print('Completed.xlsx not found')
    raise SystemExit(1)

writer = ExcelWriter(str(path))
writer.load()
rows = writer.read_epic_rows()
house_col = writer._col(config.COL_HOUSE_NO)
colony_col = writer._col(config.COL_COLONY)
print('Found house_col=', house_col, 'colony_col=', colony_col)

# Print header row for debugging
hdr = writer.header_row
print('\nHeader row (col index : value):')
for cell in writer.worksheet[hdr]:
    print(cell.column, ' : ', repr(cell.value))
count = 0
for row_idx, _ in rows:
    if count >= 20:
        break
    h = writer.worksheet.cell(row=row_idx, column=house_col).value if house_col else None
    c = writer.worksheet.cell(row=row_idx, column=colony_col).value if colony_col else None
    if h or c:
        print(f'Row {row_idx} -> H.No={repr(h)} | Colony={repr(c)}')
        count += 1

if count==0:
    print('No filled H.No/Colony cells found in first 20 rows')
