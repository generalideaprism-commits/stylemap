# -*- coding: utf-8 -*-
"""26FW 판매자료 카드 — 주차별 보관 (2026-11-16 사용자 "26SS 판매카드처럼 '주차' 카테고리")

  share/26fw-weeks/MMDD.json   한 주(일요일 마감) 치 카드 데이터 — 카드 페이지의 DATA·TREND·CARDEND·CHAT·업데이트 날짜 그대로
  share/26fw-weeks/index.json  주차 목록 — 카드 페이지가 열릴 때 읽어 '주차' 줄을 그리고, 지난 주를 누르면 그 주 JSON 을 받아 갈아 끼운다

  · 매주 FW 빌드(build_salecards.py)가 자기 주차를 저장하고 index 를 다시 쓴다.
  · 지난 주차 백필은 fw_weeks_backfill.py — 허브(production-hub) git 이력의 배포본에서 뽑는다(그 주에 배포했던 판 그대로).
  · 같은 주차를 다시 빌드하면 덮어쓴다(가장 나중 판이 남는다).
"""
import datetime
import json
import os
import re

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE) if os.path.basename(HERE) == 'build' else HERE
WEEKS_DIR = os.path.join(ROOT, 'share', '26fw-weeks')
_DEC = json.JSONDecoder()


def _grab_json(html, name):
    m = re.search(r'const %s\s*=\s*(?:/\*__[A-Z]+__\*/)?' % name, html)
    if not m:
        return None
    try:
        v, _ = _DEC.raw_decode(html[m.end():])
        return v
    except ValueError:
        return None


def _grab_str(html, name):
    m = re.search(r'const %s\s*=\s*"([^"]*)"' % name, html)
    return m.group(1) if m else ''


def parse_card_html(html):
    """카드 HTML(어느 판이든) -> {data, trend, end, chat, updated}. end 는 CARDEND, 없으면 업데이트 날짜(월) 전날(일)."""
    data = _grab_json(html, 'DATA') or []
    trend = _grab_json(html, 'TREND') or {}
    end = _grab_str(html, 'CARDEND')
    if end.startswith('__'):
        end = ''
    chat = _grab_str(html, 'CHAT')
    if chat.startswith('__'):
        chat = ''
    m = re.search(r'업데이트 날짜 <b>(\d{4}-\d{2}-\d{2})</b>', html)
    updated = m.group(1) if m else ''
    if not end and updated:
        d = datetime.date.fromisoformat(updated)
        end = (d - datetime.timedelta(days=(d.weekday() + 1) % 7 or 7)).isoformat()   # 업데이트 날짜 직전 일요일
    return {'data': data, 'trend': trend, 'end': end, 'chat': chat, 'updated': updated}


def week_key(end):
    return end[5:7] + end[8:10]


def save_week(end, data, trend, chat, updated, src, how, note=''):
    """한 주 저장 -> 파일 경로. how: 'live'(이번 빌드) · 'deployed'(허브 배포본) · 'rebuild'(엑셀로 다시 만든 판)"""
    os.makedirs(WEEKS_DIR, exist_ok=True)
    doc = {'k': week_key(end), 'end': end, 'updated': updated or '', 'chat': chat or '', 'src': src, 'how': how, 'note': note,
           'n': len(data), 'hasLots': any(s.get('lots') for s in data), 'hasWk': any(s.get('wk') for s in data),
           'hasCh': any(s.get('ch') for s in data), 'built': datetime.datetime.now().strftime('%Y-%m-%d %H:%M'),
           'trend': trend or {}, 'data': data}
    path = os.path.join(WEEKS_DIR, doc['k'] + '.json')
    tmp = path + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(doc, f, ensure_ascii=False, separators=(',', ':'))
    os.replace(tmp, path)
    return path


def rebuild_index():
    """폴더의 MMDD.json 을 훑어 index.json 을 다시 쓴다 (데이터는 빼고 머리만)."""
    weeks = []
    if os.path.isdir(WEEKS_DIR):
        for f in sorted(os.listdir(WEEKS_DIR)):
            if not re.match(r'^\d{4}\.json$', f):
                continue
            try:
                d = json.load(open(os.path.join(WEEKS_DIR, f), encoding='utf-8'))
            except Exception as e:
                print('[26fw-weeks] %s 읽기 실패: %s' % (f, e))
                continue
            weeks.append({k: d.get(k) for k in ('k', 'end', 'updated', 'chat', 'src', 'how', 'note', 'n', 'hasLots', 'hasWk', 'hasCh', 'built')})
    weeks.sort(key=lambda w: w['end'])
    os.makedirs(WEEKS_DIR, exist_ok=True)
    with open(os.path.join(WEEKS_DIR, 'index.json'), 'w', encoding='utf-8') as f:
        json.dump({'season': '26FW', 'at': datetime.datetime.now().strftime('%Y-%m-%d %H:%M'), 'weeks': weeks}, f, ensure_ascii=False)
    return weeks
