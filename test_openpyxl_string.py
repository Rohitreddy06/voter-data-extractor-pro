from openpyxl.cell.cell import Cell
from openpyxl import Workbook

wb = Workbook()
ws = wb.active
s = 'న్ సీ ల్ నార్త్'
try:
    ws.cell(row=1, column=1).value = s
    print('OK')
except Exception as e:
    print(type(e).__name__, e)
    import traceback
    traceback.print_exc()
