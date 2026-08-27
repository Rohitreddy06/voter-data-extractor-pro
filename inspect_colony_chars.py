from database import VoterDatabase

if __name__ == '__main__':
    db = VoterDatabase()
    rows = db.find_by_epic('NVT9474750')
    for r in rows:
        s = r['Area']
        print('RAW:', s)
        print('REPR:', repr(s))
        print('LOWS:', [hex(ord(c)) for c in s if ord(c) < 0x20])
        print('ONES:', [hex(ord(c)) for c in s if ord(c) > 0xFFFF])
        print('LEN:', len(s))
        print('CODEPOINTS:', [(c, hex(ord(c))) for c in s])
        print('---')
