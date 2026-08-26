"""Regenerate data.js from the So-matome N3 daily kanji workbook.

Usage:  python build_data.py [path/to/Nihongo_so-matome_N3_Kanji_daily.xlsx]

Reads the "All Cards" sheet (Week, Day, Day Card, Note ID, Tags, Front, Back)
and writes one CATEGORIES entry per week/day. The Back column holds
"<reading> - <meaning>"; the first " - " is the separator, so meanings may
themselves contain a dash.

Bump SEED_VERSION in "Kanji no Michi.html" after running this, otherwise
browsers with an existing local database keep the previous word list.
"""
import zipfile, re, json, sys, io, os
import xml.etree.ElementTree as ET
from collections import OrderedDict

log = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
PATH = sys.argv[1] if len(sys.argv) > 1 else r'E:/Nihongo_so-matome_N3_Kanji_daily.xlsx'
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data.js')

z = zipfile.ZipFile(PATH)
shared = [''.join(t.text or '' for t in si.iter(NS + 't'))
          for si in ET.fromstring(z.read('xl/sharedStrings.xml')).findall(NS + 'si')]

wb = ET.fromstring(z.read('xl/workbook.xml'))
rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
RNS = '{http://schemas.openxmlformats.org/package/2006/relationships}'
relmap = {r.get('Id'): r.get('Target') for r in rels.findall(RNS + 'Relationship')}
RID = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
sheetmap = {}
for sh in wb.find(NS + 'sheets').findall(NS + 'sheet'):
    t = relmap[sh.get(RID)].lstrip('/')
    sheetmap[sh.get('name')] = t if t.startswith('xl/') else 'xl/' + t


def colnum(ref):
    n = 0
    for ch in re.match(r'([A-Z]+)', ref).group(1):
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def read_sheet(name):
    root = ET.fromstring(z.read(sheetmap[name]))
    rows = []
    for row in root.iter(NS + 'row'):
        cells, auto = {}, 0
        for c in row.findall(NS + 'c'):
            t, v, isel = c.get('t'), c.find(NS + 'v'), c.find(NS + 'is')
            if t == 's' and v is not None:
                val = shared[int(v.text)]
            elif t == 'inlineStr' and isel is not None:
                val = ''.join(x.text or '' for x in isel.iter(NS + 't'))
            else:
                val = v.text if v is not None else ''
            r = c.get('r')
            idx = colnum(r) if r else auto
            auto = idx + 1
            cells[idx] = (val or '').strip()
        if cells:
            rows.append([cells.get(i, '') for i in range(max(cells) + 1)])
    return rows


rows = read_sheet('All Cards')
header, body = rows[0], rows[1:]
assert header[:7] == ['Week', 'Day', 'Day Card', 'Note ID', 'Tags', 'Front', 'Back'], header

# Readings that repeat the kanji instead of giving kana. Keyed by (front, raw reading).
READING_FIXES = {
    ('便利（な）', '便利な'): 'べんりな',
}

notes = []
days = OrderedDict()          # (week, day) -> list of {k,h,m}
skipped = []

for i, r in enumerate(body):
    week, day, front, back = int(r[0]), int(r[1]), r[5], r[6]

    reading, _, meaning = back.partition(' - ')
    reading, meaning = reading.strip(), meaning.strip()

    # The workbook occasionally puts the " - " one word late, leaving trailing
    # English on the reading ("とりかえる to - exchange"). Move it to the meaning.
    m = re.match(r'^(.*?)[ \u3000]+([A-Za-z][A-Za-z \'\-]*)$', reading)
    if m:
        reading = m.group(1).strip()
        meaning = (m.group(2).strip() + ' ' + meaning).strip()
        notes.append(f'W{week}D{day} {front}: re-split reading/meaning -> {reading!r} / {meaning!r}')

    fixed = READING_FIXES.get((front, reading))
    if fixed:
        notes.append(f'W{week}D{day} {front}: reading {reading!r} -> {fixed!r} (source repeated the kanji)')
        reading = fixed

    if not meaning:
        skipped.append((week, day, front, reading, back))
        continue

    days.setdefault((week, day), []).append({'k': front, 'h': reading, 'm': meaning})

for week, day, front, reading, back in skipped:
    print(f'SKIPPED W{week}D{day} {front} | {back!r} (no meaning in source)', file=log)
for n in notes:
    print('FIXED  ' + n, file=log)

NUMERALS = ['一', '二', '三', '四', '五', '六', '七']

lines = [
    '/* data.js — central data for Kanji no Michi app',
    ' *',
    ' * Source: Nihongo So-matome N3 Kanji — daily study plan',
    ' * (Nihongo_so-matome_N3_Kanji_daily.xlsx, "All Cards" sheet).',
    ' *',
    ' * The course is split into 6 weeks of 7 days. Each day is one study set;',
    ' * its `id` doubles as the cat_id used by the local SQL database, so ids',
    ' * must stay stable across regenerations.',
    ' */',
    '',
    'const WEEK_NUMERALS = ["一","二","三","四","五","六"];',
    '',
    'const CATEGORIES = [',
]

total = 0
for (week, day), words in days.items():
    total += len(words)
    lines.append(
        f'  {{ id:"w{week}d{day}", week:{week}, day:{day}, route:"w{week}/d{day}", '
        f'jp:"{NUMERALS[day - 1]}日目", en:"Day {day}", glyph:"日",'
    )
    lines.append('    words:[')
    for w in words:
        k = json.dumps(w['k'], ensure_ascii=False)
        h = json.dumps(w['h'], ensure_ascii=False)
        mm = json.dumps(w['m'], ensure_ascii=False)
        lines.append(f'      {{k:{k}, h:{h}, m:{mm}}},')
    lines.append('    ]},')

lines += [
    '];',
    '',
    '/* Weeks are derived from CATEGORIES so they stay correct after the app',
    '   reloads word lists out of the local SQL database. */',
    'const WEEKS = [1,2,3,4,5,6].map(n => ({',
    '  id:`w${n}`,',
    '  n,',
    '  jp:`第${WEEK_NUMERALS[n-1]}週`,',
    '  en:`Week ${n}`,',
    '  glyph:"週",',
    '  get days(){ return CATEGORIES.filter(c => c.week === n); },',
    '  get wordCount(){ return this.days.reduce((sum, d) => sum + d.words.length, 0); },',
    '}));',
    '',
    'const byId = id => CATEGORIES.find(c => c.id === id);',
    'const weekById = id => WEEKS.find(w => w.id === id);',
    '',
]

with open(OUT, 'w', encoding='utf-8', newline='\n') as f:
    f.write('\n'.join(lines))

print(f'\nwrote {OUT}', file=log)
print(f'day sets: {len(days)}   cards: {total}   dropped: {len(skipped)}', file=log)
for w in range(1, 7):
    counts = [len(days[(w, d)]) for d in range(1, 8)]
    print(f'  week {w}: {counts}  total {sum(counts)}', file=log)
log.flush()
