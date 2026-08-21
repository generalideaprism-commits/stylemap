# -*- coding: utf-8 -*-
"""'26 SEASON' 구글 시트(_IMGMAP + MAIN 시트)에서 도식화 이미지를 받아 gs_images.json 에 합친다.

  - _IMGMAP 시트: code -> 구글드라이브 fileId
  - 각 MAIN 시트: 'MAIN NO' 행의 품번 ↔ 2칸 위 'SAMPLE NO' 행의 샘플번호
  - 품번 자체 또는 '<샘플번호>_도식' 으로 _IMGMAP 을 찾아 이미지를 내려받는다.

사용법: python fetch_gsheet_imgmap.py
"""
import base64, csv, io, json, os, urllib.parse, urllib.request
from PIL import Image

SHEET_ID = '1KFzJy6uOGlPm3QWuV39sa3Zu63bq-VMvLxgkcIA69w4'
SHEETS = ['_IMGMAP', '26FW 우먼 MAIN', '26FW 유니맨 MAIN', '26FW 우먼 촬영MAP',
          '26FW 유니 촬영MAP', '26SS 우먼 MAIN', '26SS 우먼 MAIN 신규품번',
          '26SS 유니 MAIN 신규품번']
HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, 'gs_images.json')
MAX_PX, JPEG_Q = 560, 72


def csv_sheet(name):
    url = ('https://docs.google.com/spreadsheets/d/%s/gviz/tq?tqx=out:csv&sheet=%s'
           % (SHEET_ID, urllib.parse.quote(name)))
    try:
        raw = urllib.request.urlopen(url, timeout=60).read().decode('utf-8')
    except Exception as e:
        print('  시트 읽기 실패', name, e)
        return []
    return list(csv.reader(io.StringIO(raw)))


imgmap = {}
for r in csv_sheet('_IMGMAP')[1:]:
    if len(r) > 1 and r[0].strip():
        imgmap[r[0].strip().upper()] = r[1].strip()
print('_IMGMAP:', len(imgmap), '건')

# 품번 -> fileId
resolved = {}
for name in SHEETS[1:]:
    rows = csv_sheet(name)
    for i, r in enumerate(rows):
        if len(r) < 2 or r[1].strip() != 'MAIN NO':
            continue
        samp = rows[i - 2] if i >= 2 and len(rows[i - 2]) > 1 and rows[i - 2][1].strip() == 'SAMPLE NO' else []
        for j, c in enumerate(r):
            code = c.strip().upper()
            if len(code) < 6 or code == 'MAIN NO':
                continue
            s = samp[j].strip().upper() if len(samp) > j else ''
            fid = imgmap.get(code) or imgmap.get(code + '_도식') or \
                  (imgmap.get(s + '_도식') or imgmap.get(s) if s else None)
            if fid:
                resolved.setdefault(code, fid)
print('품번 매칭:', len(resolved), '건')

store = json.load(open(OUT, encoding='utf-8')) if os.path.exists(OUT) else {}
before = len(store)
added = 0
for code, fid in sorted(resolved.items()):
    if code in store:
        continue
    try:
        url = 'https://drive.google.com/uc?export=download&id=' + fid
        raw = urllib.request.urlopen(url, timeout=60).read()
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
