from pathlib import Path
import openpyxl
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

excel_path = Path(r'C:\Users\chenn\Downloads\59_EF_Filled_Forms_Complete_1-8-2026.xlsx')
wb = openpyxl.load_workbook(excel_path, data_only=False)
ws = wb.active
errors = []
for row in ws.iter_rows(values_only=True):
    for cell in row:
        if isinstance(cell, str) and ILLEGAL_CHARACTERS_RE.search(cell):
            errors.append((cell, [hex(ord(c)) for c in cell if ILLEGAL_CHARACTERS_RE.match(c)]))
            if len(errors) >= 20:
                break
    if len(errors) >= 20:
        break
print('illegal count', len(errors))
for cell, codes in errors[:20]:
    print(repr(cell), codes)
