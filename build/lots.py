# -*- coding: utf-8 -*-
"""판매자료 카드 — 차수(초도 · 리오더 N차)별 입고·판매 만들기.

원천
  · 총량(정본)  : 스타일맵 CO 시트 — 최초기획 S · 리오더기획 T · 입고 V · 리오더입고 X (컬러 단위)
  · 차수 쪼개기 : '입고스케줄' 구글시트의 '리오더' 탭(차순 = 리오더 차수) — 지시·투입일·입고 파샬
                  '메인' 탭의 차순(1~4차)은 리오더 차수가 아니라 초도 행의 구분이라 차수로는 쓰지 않는다
                  (메인 탭은 초도의 입고 상태 — 파샬 여부·입고일 범위·입고 예정 — 에만 쓴다)
  · 판매        : 차수별 기록이 없다 → 컬러별 선입선출(초도 → 1차 → 2차 …)로 나눈 '추정'. 합계는 항상 카드 누계와 같다.
  · 러닝        : 구품번 시점재고(이월재고)의 판매는 구품번 코드 판매 그대로(추정 아님)

읽는 순서: 허브 스냅샷 CSV(INTAKE_DIR 로 바꿀 수 있음, 24시간 이내) → 구글시트 직접(curl) → 지난 빌드가 남긴
build/intake_snapshot.json 과 오래된 허브 CSV 중 더 새것 → 없으면 차수 정보 없이 기존 팝업으로 돈다.
스냅샷은 '새로 읽은 값'일 때만 덮어쓴다(오래된 값·반쪽 응답으로 좋은 스냅샷을 망가뜨리지 않게).
"""
import csv
import datetime
import io
import json
import os
import re
import subprocess
import tempfile

SHEET_ID = '1Ju2e1KihpRVh1tJuMRuktcqElJA-UbSfal69Q3ar9NM'
STYLE_RE = re.compile(r'^[A-Z]{1,4}\d')
MAX_AGE_H = 24                       # 허브 스냅샷이 이보다 오래되면 구글시트를 직접 받아 본다
DONE = ('종결', '입고종결', '입고완료')


def parse_date(v):
    """'26. 10. 20' · '2026.09.10' · '2026-09-10' · '26.9.1' → 'YYYY-MM-DD' (못 읽거나 월·일이 이상하면 '')"""
    m = re.match(r'^\s*(\d{2,4})\s*[.\-/]\s*(\d{1,2})\s*[.\-/]\s*(\d{1,2})', str(v or ''))
    if not m:
        return ''
    y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
    y += 2000 if y < 100 else 0
    try:
        datetime.date(y, mo, d)                              # 9/31·2/30 같은 날짜는 버린다
    except ValueError:
        return ''
    if y < 2000:
        return ''
    return '%04d-%02d-%02d' % (y, mo, d)


def _n(v):
    try:
        s = str(v if v is not None else '').replace(',', '').replace('₩', '').replace('%', '').strip()
        return int(round(float(s))) if s not in ('', '-') else 0
    except ValueError:
        return 0


def md(iso):
    """'2026-09-28' → '9/28'"""
    return '%d/%d' % (int(iso[5:7]), int(iso[8:10])) if iso and len(iso) >= 10 else ''


def _range(dates):
    dates = sorted(d for d in dates if d)
    if not dates:
        return ''
    lo, hi = md(dates[0]), md(dates[-1])
    return lo if lo == hi else '%s~%s' % (lo, hi)


def parse_tab(path, tab, today):
    """입고스케줄 CSV(gviz headers=0) 한 탭 → 행 목록. 열 위치는 두 탭이 같다(0부터):
       1 차순 · 3 구분 · 4 스타일 · 6 컬러코드 · 7 컬러명 · 17 지시합 · 18 투입일 · 19/20 납기 · 22 입고종결 ·
       31 입고합 · 41+9b 파샬 b 날짜 · 49+9b 파샬 b 합계(b=0~4)
       입고 = 입고합. 종결인데 입고합 0이면 전량 입고로 본다. 오늘(스타일맵 기준일) 이후 날짜 파샬은 아직 안 들어온 것이라 뺀다."""
    out = []
    last = {'style': '', 'seq': 1}
    with io.open(path, encoding='utf-8-sig', newline='') as f:
        for cells in csv.reader(f):
            def S(i):
                return str(cells[i] if i < len(cells) else '').strip()
            style = S(4).upper()
            color = S(6).upper()
            if not STYLE_RE.match(style):
                # 병합 셀로 스타일이 비어 온 행 — 컬러와 지시가 있으면 앞 행 값으로 채운다
                if not style and color and _n(S(17)) and last['style']:
                    style = last['style']
                else:
                    continue
            if style != last['style']:
                last['seq'] = 1                                       # 차순 앞 행 채움은 같은 스타일 안에서만
            last['style'] = style
            seq = None
            if tab == 're':
                m = re.match(r'^(\d+)', S(1))
                seq = int(m.group(1)) if m else last['seq']           # 차순 칸이 비면(병합) 같은 스타일의 앞 행 차수
                last['seq'] = seq
            tot, rl = _n(S(17)), _n(S(31))
            status = S(22)
            assumed = status in DONE and rl == 0 and tot > 0
            recv = tot if assumed else rl
            parts = []
            for b in range(5):
                d, t = parse_date(S(41 + 9 * b)), _n(S(49 + 9 * b))
                if t or d:
                    parts.append([d, t])
            fut = sum(t for d, t in parts if d and d >= today)          # 아직 안 들어온 파샬
            recv0 = recv                                                 # 미래 파샬을 빼기 전 값(스냅샷을 다른 날 쓸 때 다시 계산)
            recv = max(0, recv - fut)
            out.append({'tab': tab, 'seq': seq, 'style': style, 'color': color, 'cname': S(7), 'ord': tot, 'recv': recv, 'recv0': recv0,
                        'inv': parse_date(S(18)), 'due': parse_date(S(20)) or parse_date(S(19)), 'status': status,
                        'partial': status == '파샬', 'parts': parts, 'pend': fut,
                        'rd': [d for d, t in parts if d and d < today and t > 0]})
    return out


def _curl(tab_name, dst):
    from urllib.parse import quote
    url = 'https://docs.google.com/spreadsheets/d/%s/gviz/tq?tqx=out:csv&headers=0&sheet=%s' % (SHEET_ID, quote(tab_name))
    try:
        r = subprocess.run(['curl', '-sfL', '-m', '120', '-o', dst, url], capture_output=True)
    except Exception:
        return False
    if r.returncode != 0 or not os.path.exists(dst) or os.path.getsize(dst) < 5000:
        return False
    with io.open(dst, encoding='utf-8-sig', errors='replace') as f:                # 오류 HTML 이 아니라 시트 CSV 인지
        head = f.read(4000)
    return '스타일번호' in head and '<html' not in head.lower()


def _valid_rows(rows):
    need = ('tab', 'style', 'color', 'ord', 'recv', 'parts')
    return isinstance(rows, list) and bool(rows) and all(isinstance(r, dict) and all(k in r for k in need) for r in rows)


def _rebase(rows, today):
    """스냅샷 행을 오늘 기준으로 다시 계산 — 저장 뒤 지나간 파샬은 '입고'로, 아직 안 온 것은 '예정'으로."""
    out = []
    for r in rows:
        r = dict(r)
        fut = sum(t for d, t in r['parts'] if d and d >= today)
        r['recv'] = max(0, r.get('recv0', r['recv'] + r.get('pend', 0)) - fut)
        r['pend'] = fut
        r['rd'] = [d for d, t in r['parts'] if d and d < today and t > 0]
        out.append(r)
    return out


def _write_snapshot(path, doc):
    tmp = path + '.tmp'
    with io.open(tmp, 'w', encoding='utf-8') as f:
        json.dump(doc, f, ensure_ascii=False)
    os.replace(tmp, path)                                                         # 쓰다 끊겨도 기존 스냅샷이 남는다


def load_intake(build_dir, hub_data_dir, today, log=print):
    """(행 목록, 출처 문구) — 못 읽으면 ([], '')."""
    snap = os.path.join(build_dir, 'intake_snapshot.json')
    main_p, re_p = (os.path.join(hub_data_dir, n) for n in ('intake_main.csv', 'intake_reorder.csv'))
    now = datetime.datetime.now()
    fresh = None                                                                  # (경로 둘, 출처 문구) — 새로 읽을 수 있는 값
    stale_hub = None
    if os.path.exists(main_p) and os.path.exists(re_p):
        mt = min(os.path.getmtime(main_p), os.path.getmtime(re_p))
        label = datetime.datetime.fromtimestamp(os.path.getmtime(re_p)).strftime('%m/%d %H:%M')
        if (now.timestamp() - mt) / 3600 <= MAX_AGE_H:
            fresh = ((main_p, re_p), '허브 스냅샷 %s' % label)
        else:
            stale_hub = ((main_p, re_p), '허브 스냅샷(오래됨) %s' % label, mt)
    if fresh is None:
        tmpd = tempfile.gettempdir()
        tm, tr = os.path.join(tmpd, 'cards_intake_main.csv'), os.path.join(tmpd, 'cards_intake_reorder.csv')
        if _curl('메인', tm) and _curl('리오더', tr):
            fresh = ((tm, tr), '구글시트 직접 %s' % now.strftime('%m/%d %H:%M'))
    if fresh:
        try:
            rows = parse_tab(fresh[0][0], 'main', today) + parse_tab(fresh[0][1], 're', today)
            if any(r['tab'] == 'main' for r in rows) and any(r['tab'] == 're' for r in rows):
                try:                                              # 스냅샷 저장이 실패해도 방금 읽은 값은 그대로 쓴다
                    _old = None
                    if os.path.exists(snap):
                        try:
                            _old = json.load(io.open(snap, encoding='utf-8'))
                        except Exception:
                            _old = None
                    doc = {'at': fresh[1], 'today': today, 'rows': rows}
                    if not (isinstance(_old, dict) and str(_old.get('today', '')) > today):   # 더 나중 날짜 스냅샷을 옛 스타일맵 재빌드가 덮지 않게
                        if not isinstance(_old, dict) or any(_old.get(k) != doc[k] for k in doc):   # 내용이 같으면 쓰지 않는다(깃 잡음 방지)
                            _write_snapshot(snap, doc)
                except Exception as e:
                    log('[경고] 입고스케줄 스냅샷 저장 실패(읽은 값은 사용) — %s' % e)
                return rows, fresh[1]
            log('[경고] 입고스케줄 한 탭이 비어 있어(메인 %d행·리오더 %d행) 스냅샷을 덮지 않습니다' % (
                sum(1 for r in rows if r['tab'] == 'main'), sum(1 for r in rows if r['tab'] == 're')))
        except Exception as e:
            log('[경고] 입고스케줄 읽기 실패 — %s' % e)
    snap_doc = None
    if os.path.exists(snap):
        try:
            snap_doc = json.load(io.open(snap, encoding='utf-8'))
        except Exception as e:
            log('[경고] 입고스케줄 스냅샷을 못 읽음 — %s' % e)
    # 오래된 허브 CSV 와 스냅샷 중 더 새것을 쓴다 (스냅샷은 덮어쓰지 않는다)
    if stale_hub and (not snap_doc or stale_hub[2] > os.path.getmtime(snap)):
        try:
            rows = parse_tab(stale_hub[0][0], 'main', today) + parse_tab(stale_hub[0][1], 're', today)
            if any(r['tab'] == 'main' for r in rows) and any(r['tab'] == 're' for r in rows):
                log('[경고] 입고스케줄을 새로 못 받아 오래된 허브 CSV 를 씁니다 — %s' % stale_hub[1])
                return rows, stale_hub[1]
        except Exception as e:
            log('[경고] 오래된 허브 CSV 읽기 실패 — %s' % e)
    if isinstance(snap_doc, dict) and _valid_rows(snap_doc.get('rows')):
        rows = snap_doc['rows']
        if snap_doc.get('today') != today:
            rows = _rebase(rows, today)
        log('[경고] 입고스케줄을 새로 못 읽어 지난 빌드 스냅샷 사용 (%s)' % snap_doc.get('at'))
        return rows, '지난 빌드 스냅샷 · %s' % snap_doc.get('at')
    return [], ''


def _alloc(total, weights):
    """total 을 weights 비율로 정수 배분(최대잔여법) — 합이 total 과 정확히 같다."""
    s = sum(weights)
    if total <= 0 or s <= 0:
        return [0] * len(weights)
    raw = [total * w / s for w in weights]
    out = [int(x) for x in raw]
    for i in sorted(range(len(raw)), key=lambda i: raw[i] - out[i], reverse=True)[:total - sum(out)]:
        out[i] += 1
    return out


def make_lots(key_style, colors, intake_rows, today, base=None, first_in_date=''):
    """한 카드의 차수 목록.
    colors : [{'color','cname','p0'(최초기획),'pr'(리오더기획),'ir'(리오더입고),'qtyIn','total'}]  — 신품번 컬러 단위, 스타일맵 값
    base   : 러닝 이월재고 [{'color','cname','recv'(시점재고),'sold'(구품번 7/27주~ 판매)}]  또는 None/빈 목록
    반환 {'lots':[...], 'meta':{...}} — 판매(sold)는 컬러별 선입선출 추정, 합은 카드 누계와 같다."""
    rows_re = [r for r in intake_rows if r['tab'] == 're' and r['style'] == key_style]
    rows_main = [r for r in intake_rows if r['tab'] == 'main' and r['style'] == key_style]
    seqs = sorted({r['seq'] for r in rows_re})
    no_sheet = False
    if not seqs and sum((c.get('ir') or 0) + (c.get('pr') or 0) for c in colors) > 0:
        seqs, no_sheet = [1], True                          # 스타일맵엔 리오더가 있는데 입고스케줄에 행이 없다 → 차수 미상 한 덩어리
    by_code = {c['color'].upper(): c for c in colors}
    by_name = {(c.get('cname') or '').strip().upper(): c for c in colors if c.get('cname')}
    extra = {}                                              # 카드에 없는 컬러에 걸린 리오더 행

    def card_color(r):
        c = by_code.get(r['color']) or by_name.get(r['cname'].strip().upper())
        if c:
            return c['color']
        return extra.setdefault(r['color'] or r['cname'], {'color': r['color'] or r['cname'], 'cname': r['cname']})['color']

    # 컬러 × 차수 시트 행
    cell = {}
    for r in rows_re:
        cell.setdefault((card_color(r), r['seq']), []).append(r)

    lots = []
    if base and (sum(b['recv'] for b in base) or sum(b['sold'] for b in base)):      # 합이 0 인 이월재고(음수 재고가 서로 상쇄)는 빈 차수를 만들지 않는다
        lots.append({'n': -1, 'kind': 'base', 'label': '이월재고', 'ord': sum(b['recv'] for b in base), 'recv': sum(b['recv'] for b in base),
                     'sold': sum(b['sold'] for b in base), 'pend': 0, 'state': '종결', 'rd': '7/31 시점', 'inv': '', 'due': '',
                     'colors': [{'c': b['color'], 'cn': b.get('cname', ''), 'ord': b['recv'], 'recv': b['recv'], 'sold': b['sold'], 'old': bool(b.get('old'))}
                                for b in base]})
    # 초도: 입고일 범위·상태·예정은 입고스케줄 메인 탭(파샬이면 입고중), 없으면 스타일맵 첫 입고일
    m_rd = [d for r in rows_main for d in r['rd']]
    m_pd = [d for r in rows_main for d, t in r['parts'] if d and t and d >= today]
    first = {'n': 0, 'kind': 'first', 'label': '초도', 'ord': 0, 'recv': 0, 'sold': 0, 'pend': sum(r['pend'] for r in rows_main), 'state': '종결',
             'rd': _range(m_rd) or (md(first_in_date) if re.match(r'^\d{4}-', first_in_date or '') else (first_in_date or '')),
             'pendDate': _range(m_pd), 'inv': md(min([r['inv'] for r in rows_main if r['inv']], default='')),
             'due': md(max([r['due'] for r in rows_main if r['due']], default='')), 'colors': []}
    re_lots = {n: {'n': n, 'kind': 're', 'label': ('리오더(차수 미상)' if no_sheet else '%d차 리오더' % n), 'ord': 0, 'recv': 0, 'sold': 0, 'pend': 0, 'warn': False,
                   'colors': [], '_inv': [], '_due': [], '_rd': [], '_pd': []} for n in seqs}
    over_total = 0
    warn_cols = []
    for c in colors + list(extra.values()):
        code = c['color']
        ir = c.get('ir') or 0
        qin = c.get('qtyIn') or 0
        first_recv = max(0, qin - ir)
        first_ord = c.get('p0') if c.get('p0') is not None else 0
        # 리오더 차수별 시트 입고 → 스타일맵 리오더입고(정본) 합에 맞춘다
        sheet = [(n, sum(r['recv'] for r in cell.get((code, n), []))) for n in seqs]
        s_sum = sum(v for _, v in sheet)
        if s_sum == ir:
            alloc = [v for _, v in sheet]
            cwarn = False
        else:
            cwarn = (s_sum > 0 or ir > 0) and bool(seqs) and not no_sheet
            if ir > 0 and s_sum > 0:
                alloc = _alloc(ir, [v for _, v in sheet])
            elif ir > 0 and seqs:                           # 시트엔 입고 기록이 없는데 스타일맵엔 리오더입고가 있다 → 첫 차수에 둔다
                alloc = [ir if i == 0 else 0 for i in range(len(seqs))]
            else:
                alloc = [0] * len(seqs)
        if cwarn:
            warn_cols.append(code)
        # 선입선출 판매 추정 — 입고 순서 = 초도, 1차, 2차 …
        remain = c.get('total') or 0
        sold_first = min(remain, first_recv)
        remain -= sold_first
        sold_re = []
        for v in alloc:
            s_ = min(remain, v)
            sold_re.append(s_)
            remain -= s_
        if remain > 0:                                      # 입고보다 판매가 많다(스타일맵 입고·판매 시차) → 마지막 입고 차수에 얹고 따로 센다
            over_total += remain
            tgt = max([i for i, v in enumerate(alloc) if v > 0], default=-1)
            if tgt >= 0:
                sold_re[tgt] += remain
            else:
                sold_first += remain
        first['ord'] += first_ord
        first['recv'] += first_recv
        first['sold'] += sold_first
        first['colors'].append({'c': code, 'cn': c.get('cname', ''), 'ord': first_ord, 'recv': first_recv, 'sold': sold_first})
        for i, n in enumerate(seqs):
            rr = cell.get((code, n), [])
            if no_sheet:                                    # 시트 행이 없으면 스타일맵 값(리오더기획·리오더입고)으로
                ordq, pend = (c.get('pr') or 0), 0
            else:
                ordq, pend = sum(r['ord'] for r in rr), sum(r['pend'] for r in rr)
            L = re_lots[n]
            L['ord'] += ordq
            L['recv'] += alloc[i]
            L['sold'] += sold_re[i]
            L['pend'] += pend
            if rr or alloc[i] > 0 or (no_sheet and ordq):   # 시트 행이 없어도 스타일맵이 입고로 잡은 컬러는 컬러 행에 넣는다(헤더와 표 합계가 같게)
                L['colors'].append({'c': code, 'cn': c.get('cname', ''), 'ord': ordq, 'recv': alloc[i], 'sold': sold_re[i], 'pend': pend})
                L['_inv'] += [r['inv'] for r in rr if r['inv']]
                L['_due'] += [r['due'] for r in rr if r['due']]
                L['_rd'] += [d for r in rr for d in r['rd']]
                L['_pd'] += [d for r in rr for d, t in r['parts'] if d and t and d >= today]
    out = ([lots[0]] if lots else []) + [first]
    for n in seqs:
        L = re_lots[n]
        inv, due, rd, pd = L.pop('_inv'), L.pop('_due'), L.pop('_rd'), L.pop('_pd')
        L['inv'] = md(min(inv)) if inv else ''
        L['due'] = md(max(due)) if due else ''
        L['rd'] = _range(rd)
        L['pendDate'] = _range(pd)
        L['state'] = '예정' if L['recv'] == 0 else ('종결' if L['pend'] == 0 and L['recv'] >= L['ord'] * 0.98 else '입고중')
        L['warn'] = any(w for w in warn_cols if any(cc['c'] == w for cc in L['colors']))
        out.append(L)
    # 초도 상태: 입고가 없으면 예정, 메인 탭이 파샬(일부 입고)이거나 입고 예정이 남았으면 입고중
    if first['recv'] == 0:
        first['state'] = '예정'
    elif any(r['partial'] for r in rows_main) or first['pend'] > 0:
        first['state'] = '입고중'
    meta = {'over': over_total, 'warnColors': warn_cols, 'extra': [e['color'] for e in extra.values()], 'noSheet': no_sheet}
    return {'lots': out, 'meta': meta}
