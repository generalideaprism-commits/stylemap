# -*- coding: utf-8 -*-
"""구글 스프레드시트(스타일맵)에서 도식화 이미지를 뽑아 gs_images.json 으로 저장.

  - 'MAIN NO' 가 적힌 행의 스타일 넘버를 읽고, 같은 열 2칸 아래에 붙어 있는 이미지를 연결한다.
  - 이미지는 축소/압축해서 base64 data URI 로 저장 -> 카드 HTML 안에 정적으로 박힌다.

사용법:  python extract_gsheet_images.py [이미 받아둔 xlsx 경로]
"""
import base64, io, json, os, re, sys, urllib.request, zipfile
import xml.etree.ElementTree as ET
from PIL import Image

SHEET_ID = '1dnrfBBj1gO3Exhfh0NX6FuYJf0JOKW1gBxyRisfsdiY'
EXPORT = 'https://docs.google.com/spreadsheets/d/%s/export?format=xlsx' % SHEET_ID
HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, 'gs_images.json')

MAX_PX = 560          # 긴 변 기준 축소 크기
JPEG_Q = 72

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
XDR = '{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}'
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
CODE_RE = re.compile(r'^[A-Z]{2,3}[A-Z0-9]{5,}$')

src = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, '_gsheet_export.xlsx')
if not os.path.exists(src):
    print('다운로드 중...', EXPORT)
    urllib.request.urlretrieve(EXPORT, src)
print('원본:', src, round(os.path.getsize(src) / 1e6, 1), 'MB')

z = zipfile.ZipFile(src)

shared = []
for _, el in ET.iterparse(z.open('xl/sharedStrings.xml')):
    if el.tag == NS + 'si':
        shared.append(''.join(t.text or '' for t in el.iter(NS + 't')))
        el.clear()


def cells(path):
    d = {}
    for _, c in ET.iterparse(z.open(path)):
        if c.tag == NS + 'c':
            v = c.find(NS + 'v')
            if v is not None and v.text is not None:
                d[c.get('r')] = shared[int(v.text)] if c.get('t') == 's' else v.text
            c.clear()
    return d


def rels(path):
    try:
        x = z.read(path).decode('utf8')
    except KeyError:
        return {}
    return {m.group(1): m.group(2)
            for m in re.finditer(r'Id="([^"]+)"[^>]*?Target="([^"]+)"', x)}


def col2i(s):
    n = 0
    for ch in s:
        n = n * 26 + ord(ch) - 64
    return n - 1


mapping = {}   # 스타일 -> zip 내부 이미지 경로
for i in range(1, 60):
    sp = 'xl/worksheets/sheet%d.xml' % i
    if sp not in z.namelist():
        continue
    d = cells(sp)
    head = [k for k, v in d.items() if str(v).strip() == 'MAIN NO']
    if not head:
        continue
    dr = [t for t in rels('xl/worksheets/_rels/sheet%d.xml.rels' % i).values() if 'drawing' in t]
    if not dr:
        continue
    dpath = 'xl/' + dr[0].replace('../', '')
    drel = rels(re.sub(r'drawings/', 'drawings/_rels/', dpath) + '.rels')
    anchors = {}
    for a in list(ET.fromstring(z.read(dpath))):
        frm = a.find(XDR + 'from')
        blip = a.find('.//' + A + 'blip')
        if frm is None or blip is None:
            continue
        key = (int(frm.find(XDR + 'row').text), int(frm.find(XDR + 'col').text))
        anchors[key] = 'xl/' + drel[blip.get(R + 'embed')].replace('../', '')
    hit = 0
    for h in head:
        row = int(re.sub(r'\D', '', h))
        for k, v in d.items():
            if re.sub(r'\D', '', k) != str(row):
                continue
            code = str(v).strip().upper()
            if not CODE_RE.match(code):
                continue
            # 스타일 넘버 기준 2칸 아래(0-based 행 = row+1) 의 이미지
            img = anchors.get((row + 1, col2i(re.sub(r'\d', '', k))))
            if img and code not in mapping:
                mapping[code] = img
                hit += 1
    print('  sheet%-3d MAIN NO행 %2d / 이미지 %3d / 매칭 %3d' % (i, len(head), len(anchors), hit))

print('스타일-이미지 매칭:', len(mapping))

out, total = {}, 0
for code, path in sorted(mapping.items()):
    try:
        im = Image.open(io.BytesIO(z.read(path)))
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
        out[code] = 'data:image/jpeg;base64,' + base64.b64encode(buf.getvalue()).decode()
        total += len(out[code])
    except Exception as e:
        print('  이미지 변환 실패', code, e)

json.dump(out, open(OUT, 'w', encoding='utf-8'), ensure_ascii=False)
print('저장:', OUT, len(out), '건 /', round(total / 1e6, 1), 'MB(base64)')
