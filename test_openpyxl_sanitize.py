import openpyxl
from utils import safe_str

text = 'ఎన్\u200d సీ ఎల్\u200d నార్త్ \x00\x00వేన్యు'
print('original repr:', repr(text))
print('safe repr:', repr(safe_str(text)))
print('safe ords:', [hex(ord(c)) for c in safe_str(text)])
wb = openpyxl.Workbook()
ws = wb.active
ws['A1'] = safe_str(text)
wb.save('test_sanitized.xlsx')
print('saved test_sanitized.xlsx')
