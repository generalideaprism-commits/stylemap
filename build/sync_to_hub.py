# -*- coding: utf-8 -*-
"""stylemap 산출물을 production-hub 저장소의 share/ 로 복사하고 카탈로그 날짜를 맞춘다.

  하는 일
    1. share/*.html 과 share/img/*.jpg 를 허브 share/ 로 복사 (안 쓰는 이미지는 삭제)
    2. 허브 share/index.html 의 cards26fw · sum26fw 항목 `updated` 를 이번 데이터 날짜로 갱신
    3. --note "..." 를 주면 두 항목의 history 맨 위에 한 줄 추가
    4. --commit 을 주면 허브 저장소에 커밋·푸시까지

  배포는 자동으로 하지 않는다. 허브 규칙대로 마지막에 직접:  npx vercel --prod

사용법
    python build/sync_to_hub.py
    python build/sync_to_hub.py --note "260825 스타일맵 반영" --commit
    python build/sync_to_hub.py --hub "C:/경로/production-hub"
"""
import argparse, os, re, shutil, subprocess, sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SRC_SHARE = os.path.join(ROOT, 'share')
KEYS = ['cards26fw', 'sum26fw']          # 허브 share/index.html 의 FILES 키

ap = argparse.ArgumentParser()
ap.add_argument('--hub', default=os.environ.get('HUB_REPO') or
                os.path.join(os.path.dirname(ROOT), 'production-hub'),
                help='production-hub 저장소 경로 (기본: stylemap 옆 폴더)')
ap.add_argument('--note', help='변경 이력에 추가할 한 줄')
ap.add_argument('--commit', action='store_true', help='허브 저장소에 커밋·푸시까지')
a = ap.parse_args()

HUB_SHARE = os.path.join(a.hub, 'share')
INDEX = os.path.join(HUB_SHARE, 'index.html')
if not os.path.exists(INDEX):
    raise SystemExit('허브 저장소를 찾지 못했습니다: %s\n--hub 로 경로를 지정하세요.' % INDEX)

# 1) 파일 복사 --------------------------------------------------------------
htmls = [f for f in os.listdir(SRC_SHARE) if f.endswith('.html')]
for f in htmls:
    shutil.copy2(os.path.join(SRC_SHARE, f), os.path.join(HUB_SHARE, f))
print('HTML 복사:', ', '.join(htmls))

src_img, dst_img = os.path.join(SRC_SHARE, 'img'), os.path.join(HUB_SHARE, 'img')
os.makedirs(dst_img, exist_ok=True)
names = set(os.listdir(src_img)) if os.path.isdir(src_img) else set()
copied = 0
for n in names:
    s, d = os.path.join(src_img, n), os.path.join(dst_img, n)
    if not os.path.exists(d) or open(s, 'rb').read() != open(d, 'rb').read():
        shutil.copy2(s, d)
        copied += 1
removed = 0
for n in os.listdir(dst_img):
    if n not in names:
        os.remove(os.path.join(dst_img, n))
        removed += 1
print('이미지 %d장 (변경 %d, 삭제 %d)' % (len(names), copied, removed))

# 2) 데이터 날짜(= 카드 화면의 업데이트 날짜) 읽기 ---------------------------
card = open(os.path.join(SRC_SHARE, '26fw-sales-cards.html'), encoding='utf-8').read()
m = re.search(r'업데이트 날짜 <b>(\d{4}-\d{2}-\d{2})</b>', card)
if not m:
    raise SystemExit('판매카드에서 업데이트 날짜를 찾지 못했습니다.')
date = m.group(1)
print('데이터 날짜:', date)

# 3) 허브 카탈로그 갱신 ------------------------------------------------------
idx = open(INDEX, encoding='utf-8').read()
for key in KEYS:
    i = idx.find("key:'%s'" % key)
    if i < 0:
        print('  ! FILES 에 %s 항목이 없습니다 — 건너뜀' % key)
        continue
    end = idx.index('\n  }', i)                       # 이 항목의 끝
    block = idx[i:end]
    new = re.sub(r"updated:'[^']*'", "updated:'%s'" % date, block)
    if a.note:
        new = re.sub(r"history:\[\s*", "history:[\n      {d:'%s', t:'%s'},\n      "
                     % (date, a.note.replace("'", "’")), new, count=1)
    idx = idx[:i] + new + idx[end:]
open(INDEX, 'w', encoding='utf-8').write(idx)
print('카탈로그 갱신: %s (updated=%s%s)' % (', '.join(KEYS), date, ', history 추가' if a.note else ''))

# 4) 커밋 --------------------------------------------------------------------
if a.commit:
    msg = '공유파일 26FW 판매카드·집계표 갱신 (%s)' % date + (('\n\n- ' + a.note) if a.note else '')
    subprocess.run(['git', 'add', 'share'], cwd=a.hub, check=True)
    r = subprocess.run(['git', 'commit', '-m', msg], cwd=a.hub)
    if r.returncode == 0:
        subprocess.run(['git', 'push', 'origin', 'HEAD'], cwd=a.hub, check=True)
        print('허브 저장소에 커밋·푸시 완료')
    else:
        print('변경 사항이 없어 커밋하지 않았습니다')

print('\n다음 단계 — 허브는 수동 배포입니다. 저장소 루트에서:\n  cd "%s" && npx vercel --prod' % a.hub)
