# -*- coding: utf-8 -*-
"""26FW 판매자료 카드 지난 주차 백필 — 허브(production-hub) git 이력의 배포본에서 주차별 JSON 을 만든다

  허브의 share/26fw-sales-cards.html 커밋을 전부 훑어 마감 일요일(CARDEND, 없으면 업데이트 날짜 직전 일요일)별로
  **가장 나중 배포본**을 share/26fw-weeks/MMDD.json 으로 저장한다 (그 주에 배포했던 판 그대로 — 지표 정의는 그때 빌드 기준).
  이번 빌드 주차(현재 share/26fw-sales-cards.html 의 CARDEND)와 같거나 늦은 주는 건드리지 않는다 (빌드가 직접 저장).

  사용:  python build/fw_weeks_backfill.py            # 허브 = ../production-hub
         python build/fw_weeks_backfill.py --hub 경로 --force   # 이미 있는 주차도 다시 쓴다
"""
import argparse
import os
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import fwweeks

try: sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception: pass

ap = argparse.ArgumentParser()
ap.add_argument('--hub', default=os.path.join(os.path.dirname(fwweeks.ROOT), 'production-hub'))
ap.add_argument('--force', action='store_true')
a = ap.parse_args()

REL = 'share/26fw-sales-cards.html'
live_end = ''
cur = os.path.join(fwweeks.ROOT, 'share', '26fw-sales-cards.html')
if os.path.exists(cur):
    live_end = fwweeks.parse_card_html(open(cur, encoding='utf-8').read())['end']
print('이번 빌드 주차(건너뜀):', live_end or '-')

log = subprocess.run(['git', 'log', '--format=%H %ad %s', '--date=short', '--', REL], cwd=a.hub, capture_output=True).stdout.decode('utf-8', 'replace')
done = {}
for line in log.splitlines():
    sha, d, msg = line.split(' ', 2)
    html = subprocess.run(['git', 'show', '%s:%s' % (sha, REL)], cwd=a.hub, capture_output=True).stdout.decode('utf-8', 'replace')
    p = fwweeks.parse_card_html(html)
    if not p['data'] or not p['end']:
        print('  %s %s: 데이터/마감일 없음 — 건너뜀' % (sha[:8], d)); continue
    if live_end and p['end'] >= live_end:
        continue
    if p['end'] in done:
        continue                                           # 같은 주는 가장 나중 커밋(먼저 나온 것)만
    k = fwweeks.week_key(p['end'])
    path = os.path.join(fwweeks.WEEKS_DIR, k + '.json')
    if os.path.exists(path) and not a.force:
        done[p['end']] = 'skip'; print('  %s 있음 — 건너뜀 (--force 로 다시)' % k); continue
    fwweeks.save_week(p['end'], p['data'], p['trend'], p['chat'], p['updated'], '허브 %s (%s)' % (sha[:8], d), 'deployed',
                      '그 주에 배포했던 판 그대로 — 지표 정의는 그때 빌드 기준')
    done[p['end']] = sha[:8]
    print('  %s ← %s %s 카드 %d장 (업데이트 %s · MARK %s)' % (k, sha[:8], d, len(p['data']), p['updated'] or '-', p['chat'] or '-'))

weeks = fwweeks.rebuild_index()
print('index.json: 주차 %d개 — %s' % (len(weeks), ' '.join(w['k'] for w in weeks)))
