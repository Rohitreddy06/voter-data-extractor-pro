import sqlite3
import config
from utils import _is_xml_char

conn = sqlite3.connect(config.DEFAULT_DB_PATH)
cur = conn.cursor()
for row in cur.execute('SELECT rowid, EPIC, HouseNo, Area, Address FROM voters'):
    rowid, epic, house, area, address = row
    for col, val in [('HouseNo', house), ('Area', area), ('Address', address)]:
        if not val:
            continue
        for i, ch in enumerate(val):
            if not _is_xml_char(ord(ch)):
                print('BAD', rowid, epic, col, i, hex(ord(ch)), repr(ch), repr(val))
                break
conn.close()
