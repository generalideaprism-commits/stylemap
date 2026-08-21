# -*- coding: utf-8 -*-
"""웹에 게시된 구글 스프레드시트에서 도식화 이미지를 뽑아 gs_images.json 에 합친다.

  게시 HTML(pubhtml/sheet) 은 셀 위 이미지를 <img> 로 그대로 내보내므로
  'MAIN NO' 행의 품번 ↔ 같은 열 '도식화' 행의 이미지를 짝지어 받는다.

사용법: python fetch_pubhtml_images.py [gid ...]      (인자 없으면 GIDS 기본값)
"""
import base64, io, json, os, sys, urllib.request
from html.parser import HTMLParser
from PIL import Image

PUB = ('https://docs.google.com/spreadsheets/d/e/2PACX-1vQmTAV1eWaOmbPcslMrrpy0doVkN'
       '_H7FsgMHL29J6bfig1LqdIaipbpxW3Y2r2JPdxi3sKjv5aIRq0e')
GIDS = ['1620257578']          # 게시된 시트의 gid (추가로 게시하면 여기에 덧붙이면 됨)
HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, 'gs_images.json')
MAX_PX, JPEG_Q = 560, 72


class Grid(HTMLParser):
    """게시 HTML 표 -> 행/열 격자 (colspan 반영, 셀마다 text/img)"""
    def __init__(self):
        super().__init__()
        self.rows, self._row, self._cell = [], None, None

    def handle_starttag(self, tag, attrs):
        d = dict(attrs)
        if tag == 'tr':
            self._row = []
        elif tag == 'td' and self._row is not None:
            self._cell = {'text': '', 'img': None, 'span': int(d.get('colspan', 1) or 1)}
        elif tag == 'img' and self._cell is not None:
            self._cell['img'] = d.get('src')

    def handle_endtag(self, tag):
        if tag == 'td' and self._cell is not None:
            self._row.append(self._cell)
            for _ in range(self._cell['span'] - 1):
                self._row.append({'text': '', 'img': None, 'span': 1})
            self._cell = None
        elif tag == 'tr' and self._row is not None:
            self.rows.append(self._row)
            self._row = None

    def handle_data(self, data):
        if self._cell is not None:
            self._cell['text'] += data.strip()


def sheet_grid(gid):
    url = '%s/pubhtml/sheet?headers=false&gid=%s' % (PUB, gid)
    html = urllib.request.urlopen(url, timeout=120).read().decode('utf-8', 'ignore')
    g = Grid()
    g.feed(html)
    return g.rows


mapping = {}
for gid in (sys.argv[1:] or GIDS):
    rows = sheet_grid(gid)
    hit = 0
    for i, row in enumerate(rows):
        texts = [c['text'] for c in row]
        if 'MAIN NO' not in texts:
            continue
        # 'MAIN NO' 행 아래에서 '도식화' 행 찾기 (보통 2칸 아래)
        draw = next((k for k in range(i + 1, min(i + 6, len(rows)))
                     if '도식화' in [c['text'] for c in rows[k]]), None)
        if draw is None:
            continue
        for j, c in enumerate(row):
            code = c['text'].strip().upper()
            if len(code) < 6 or code == 'MAIN NO':
                continue
            cell = rows[draw][j] if j < len(rows[draw]) else None
            if cell and cell['img']:
                mapping.setdefault(code, cell['img'])
                hit += 1
    print('gid %s : 행 %d / 품번-이미지 %d' % (gid, len(rows), hit))

print('총 매칭:', len(mapping))

store = json.load(open(OUT, encoding='utf-8')) if os.path.exists(OUT) else {}
before, added = len(store), 0
for code, url in sorted(mapping.items()):
    if code in store:
        continue
    try:
        # =w183-h168 같은 크기 접미사를 떼면 원본 크기로 받는다
        full = url.split('=w')[0].split('=s')[0]
        raw = urllib.request.urlopen(full, timeout=60).read()
        im = Image.open(io.BytesIO(raw))
        if im.mode in ('RGBA', 'LA', 'P'):
            bg = Image.new('RGB', im.size, (255, 255, 255))
            im = im.convert('RGBA')
            bg.paste(im, mask=im.split()[-1])
            im = bg
        else:
            im = im.convert('RGB')
        im.thumbnail((MAX_PX, MAX_PX), Image.LANCZOS)
        buf = io.BytesIO()
        im.save(buf, 'JPEG', quality=JPEG_Q, optimize=True)
        store[code] = 'data:image/jpeg;base64,' + base64.b64encode(buf.getvalue()).decode()
        added += 1
    except Exception as e:
        print('  실패', code, e)

json.dump(store, open(OUT, 'w', encoding='utf-8'), ensure_ascii=False)
print('gs_images.json:', before, '->', len(store), '(추가', added, ')')
