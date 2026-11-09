# -*- coding: utf-8 -*-
"""허브 '26FW 시즌 판매 현황'(share/26fw-season-status.html)에 이번 주 종합(2) 계산값을 한 주차로 넣는다.

  그 화면은 원래 NAS 의 주별 스타일맵 엑셀 '종합' 시트 셀 값을 그대로 뜬 것(KWAN 2026-11-06).
  11/09 주부터 스타일맵이 구글 시트 + 종합(2) 로 바뀌어 엑셀 '종합' 이 더 안 생기므로,
  빌드가 구글 시트에서 계산한 종합(2) 값을 같은 모양(최신 주차의 셀 배치·서식·병합)으로 끼워 넣는다.

  동작
    1. 허브의 현재 페이지를 읽어 DATA(주차 목록) 를 파싱
    2. 가장 최근 주차를 '틀' 로 삼아 — 셀 자리·스타일·병합·숫자 서식은 그대로, 값만 종합(2) 계산값으로 바꾼 새 주차를 만든다
       (서식은 틀 칸의 '표시 글자 ↔ 원값' 에서 추정: 천단위 콤마·소수 자리·백만원 단위·% · 0 은 '-')
    3. 같은 key(MMDD) 주차가 있으면 바꾸고 없으면 뒤에 붙인 뒤, stylemap/share/ 에 쓴다 → sync_to_hub 가 허브로 복사

  빌드(build_salecards.py, 구글 시트 모드)가 끝에서 부른다. 따로 돌리려면:
    python build/season_status.py            (구글 시트를 다시 읽는다)
"""
import datetime, hashlib, json, os, re, sys
from decimal import Decimal, ROUND_HALF_UP

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
HUB = os.environ.get('HUB_REPO') or os.path.join(os.path.dirname(ROOT), 'production-hub')
PAGE = '26fw-season-status.html'
HUB_PAGE = os.path.join(HUB, 'share', PAGE)
OUT_PAGE = os.path.join(ROOT, 'share', PAGE)
SHEET_LABEL = "구글 시트 '대쉬보드PIP업데이트파일' → 종합(2) 계산값"


def col_idx(letters):
    n = 0
    for ch in letters: n = n * 26 + ord(ch) - 64
    return n - 1


# ------------------------------------------------------------------ 숫자 서식 추정/적용
def _rhu(v, d):
    q = Decimal(1).scaleb(-d)
    return Decimal(repr(float(v))).quantize(q, rounding=ROUND_HALF_UP)


def _fmt_num(v, d, scale):
    x = _rhu(v / scale, d)
    s = '{:,.%df}' % d
    return s.format(x)


def infer(text, raw):
    """(표시 글자, 원값) → 서식 {'k': 'pct'|'num', 'd': 소수자리, 's': 배율} 또는 None."""
    if raw is None or not isinstance(raw, (int, float)) or not isinstance(text, str): return None
    t = text.strip()
    if t in ('-', '', '0', '0%'):                       # 0 표시 — 서식은 못 알아냄
        return None
    m = re.fullmatch(r'(-?)([\d,]*\.?\d+)%', t)
    if m:
        d = len(m.group(2).split('.')[1]) if '.' in m.group(2) else 0
        return {'k': 'pct', 'd': d}
    m = re.fullmatch(r'(-?)([\d,]*\.?\d+)', t)
    if not m or raw == 0: return None
    n = float(m.group(2).replace(',', '')) * (-1 if m.group(1) else 1)
    d = len(m.group(2).split('.')[1]) if '.' in m.group(2) else 0
    if n == 0: return None
    ratio = raw / n
    scale = min((1, 1000, 1e6, 1e8), key=lambda s: abs(ratio - s))
    if abs(ratio / scale - 1) > 0.35: return None            # 반올림으로 설명 안 되는 차이 → 포기
    return {'k': 'num', 'd': d, 's': scale}


def apply(fmt, v, zero='-'):
    if v is None: v = 0
    if isinstance(v, str): return v
    if v == 0: return zero
    if fmt is None: return '{:,.0f}'.format(v) if float(v).is_integer() else '{:,.2f}'.format(v)
    if fmt['k'] == 'pct': return ('{:,.%df}' % fmt['d']).format(_rhu(v * 100, fmt['d'])) + '%'
    return _fmt_num(v, fmt['d'], fmt['s'])


# ------------------------------------------------------------------ 본체
def load_page(path=HUB_PAGE):
    html = open(path, encoding='utf-8').read()
    m = re.search(r'const DATA = (\{.*?\});\n', html, re.S)
    if not m: raise SystemExit('%s 에서 DATA 를 찾지 못했습니다' % path)
    return html, m, json.loads(m.group(1))


def grid_get(rows, r, c):
    """종합(2) 격자(0-based 행 리스트)에서 엑셀 (행번호, 열index) 값."""
    if r - 1 < len(rows):
        row = rows[r - 1]
        if c < len(row): return row[c]
    return None


class Formats:
    """틀 주차의 칸들에서 열별 서식(0 아닌 칸)과 0 표시 글자('-'/'0'/'0%')를 모아 둔다. 자기 칸 → 같은 열의 가까운 행 순으로 찾는다."""
    def __init__(self, tpl):
        self.fmt, self.zero = {}, {}
        for rn, cells in tpl['rows']:
            for c in cells:
                if len(c) > 3 and isinstance(c[3], (int, float)):
                    if c[3] == 0: self.zero.setdefault(c[0], []).append((rn, c[1]))
                    else:
                        f = infer(c[1], c[3])
                        if f: self.fmt.setdefault(c[0], []).append((rn, f))

    @staticmethod
    def _near(cand, rn, default):
        return min(cand, key=lambda x: abs(x[0] - rn))[1] if cand else default

    def of(self, c, rn):
        """틀 칸 c(행 rn) 의 (서식, 0 표시)."""
        raw = c[3] if len(c) > 3 else None
        f = infer(c[1], raw) if raw else None
        z = c[1] if raw == 0 else None
        return (f or self._near(self.fmt.get(c[0], []), rn, None),
                z if z is not None else self._near(self.zero.get(c[0], []), rn, '-'))


def build_week(tpl, jong, key, fname):
    """틀 주차(tpl)의 배치로 종합(2) 계산값(jong) 주차를 만든다. → (week, 보고 dict)"""
    cols = tpl['cols']
    letters = [c[0] for c in cols]
    F = Formats(tpl)
    rows_out, rep = [], {'label_diff': [], 'num_changed': 0, 'cells': 0, 'text_from_grid': 0, 'new_nonzero_missing': []}
    for rn, cells in tpl['rows']:
        out = []
        seen = set()
        for c in cells:
            ci = c[0]; seen.add(ci)
            addr_c = col_idx(letters[ci])
            v = grid_get(jong, rn, addr_c)
            c2 = list(c)
            has_raw = len(c) > 3 and isinstance(c[3], (int, float))
            if has_raw or (isinstance(v, (int, float)) and not isinstance(v, bool)):
                if isinstance(v, str):                        # 숫자 칸에 글자(#DIV/0! 등)
                    c2[1] = v; c2[3] = None if len(c2) > 3 else None
                else:
                    fmt, zero = F.of(c, rn)
                    vv = 0 if v is None else float(v)
                    if has_raw and abs(vv - float(c[3])) > 1e-9: rep['num_changed'] += 1
                    c2[1] = apply(fmt, vv, zero)
                    if len(c2) > 3: c2[3] = vv if not vv.is_integer() else int(vv)
                    else: c2.append(vv if not vv.is_integer() else int(vv))
                rep['cells'] += 1
            else:                                              # 글자 칸(머리글·이름)
                if isinstance(v, str) and v.strip() and v.strip() != str(c[1]).strip():
                    rep['label_diff'].append((letters[ci] + str(rn), c[1], v))
                    c2[1] = v; rep['text_from_grid'] += 1
            out.append(c2)
        # 틀에 없던 자리인데 계산값이 0 이 아닌 칸 — 보고만 한다(배치를 모르므로 넣지 않음)
        if rn - 1 < len(jong):
            for ci, L in enumerate(letters):
                if ci in seen: continue
                v = grid_get(jong, rn, col_idx(L))
                if isinstance(v, (int, float)) and not isinstance(v, bool) and v != 0:
                    rep['new_nonzero_missing'].append(L + str(rn))
        rows_out.append([rn, out])
    week = {k: tpl[k] for k in tpl if k not in ('diff',)}
    week.update({'key': key, 'file': fname, 'base': None, 'prev': None, 'rows': rows_out,
                 'warn': [], 'sig': hashlib.sha1(json.dumps(rows_out, ensure_ascii=False).encode()).hexdigest()[:10]})
    return week, rep


def selfcheck(tpl):
    """서식 추정이 틀 자체를 되살리는지: 틀의 원값을 다시 표시해 글자가 같은지 센다."""
    F = Formats(tpl)
    ok = bad = 0; ex = []
    for rn, cells in tpl['rows']:
        for c in cells:
            if len(c) > 3 and isinstance(c[3], (int, float)):
                f, z = F.of(c, rn)
                if c[3] == 0:                                  # 0 칸은 자기 글자 대신 이웃에서 빌린 0 표시로 검증(엄격)
                    z = F._near([x for x in F.zero.get(c[0], []) if x[0] != rn], rn, '-')
                t = apply(f, c[3], z)
                if t == c[1]: ok += 1
                else:
                    bad += 1
                    if len(ex) < 12: ex.append((tpl['cols'][c[0]][0] + str(rn), c[1], t, c[3]))
    return ok, bad, ex


def update(jong, ref, src_name, hub_page=HUB_PAGE, out_page=OUT_PAGE, verbose=True):
    """jong: 종합(2) 계산 격자, ref: 기준 일요일(date), src_name: 원본 표시 이름."""
    html, m, D = load_page(hub_page)
    weeks = D['weeks']
    if not weeks: raise SystemExit('페이지에 주차가 없습니다')
    file_day = ref + datetime.timedelta(days=1)                 # 파일 날짜 = 기준 일요일 + 1 (월요일)
    key = file_day.strftime('%m%d')
    # 틀 = 우리가 만들지 않은(엑셀에서 뜬) 가장 최근 주차. 없으면 최근 주차.
    tpl = next((w for w in reversed(weeks) if not w.get('gen_by')), weeks[-1])
    ok, bad, ex = selfcheck(tpl)
    week, rep = build_week(tpl, jong, key, src_name)
    week['gen_by'] = 'stylemap-build'
    # 틀이 우리 주차가 아니면, 그 주차의 diff(최신 주 대비 머리글 차이) 표시는 그대로 둔다
    idx = next((i for i, w in enumerate(weeks) if w['key'] == key), None)
    if idx is None:
        weeks.append(week); weeks.sort(key=lambda w: w['key'])   # 같은 해 안에서만 MMDD 정렬
        # 연도 바뀜(12월→1월) 대비: 정렬은 기존 순서를 깨지 않게 — 새 주차가 끝이 아니면 끝으로 보낸다
        if weeks[-1] is not week:
            weeks.remove(week); weeks.append(week)
    else:
        weeks[idx] = week
    D['gen'] = datetime.datetime.now().strftime('%Y-%m-%d %H:%M')
    if SHEET_LABEL not in D.get('src', ''):
        D['src'] = (D.get('src') or '') + ' · 11/09 주부터 ' + SHEET_LABEL
    new_json = json.dumps(D, ensure_ascii=False, separators=(',', ':'))
    html2 = html[:m.start(1)] + new_json + html[m.end(1):]
    os.makedirs(os.path.dirname(out_page), exist_ok=True)
    open(out_page, 'w', encoding='utf-8').write(html2)
    if verbose:
        print('시즌 판매 현황: 주차 %s 를 %s (틀: %s 주차 %s) → %s' % (key, '바꿈' if idx is not None else '추가', tpl['key'], tpl['file'], os.path.relpath(out_page, ROOT)))
        print('  서식 되살리기 검증: 일치 %d · 불일치 %d%s' % (ok, bad, (' 예: %s' % ex[:5]) if bad else ''))
        print('  숫자 칸 %d (틀과 값 다른 칸 %d) · 머리글/이름 틀과 다른 칸 %d%s' % (
            rep['cells'], rep['num_changed'], len(rep['label_diff']), (' 예: %s' % rep['label_diff'][:6]) if rep['label_diff'] else ''))
        if rep['new_nonzero_missing']:
            print('  [경고] 틀에 자리가 없는데 값이 있는 칸 %d: %s' % (len(rep['new_nonzero_missing']), rep['new_nonzero_missing'][:12]))
    return week, rep


if __name__ == '__main__':
    try: sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception: pass
    sys.path.insert(0, HERE)
    import gsource
    GS = gsource.load()
    update(GS['jonghap'], GS['ref'], '%s_26FW 스타일맵(구글시트)' % (GS['ref'] + datetime.timedelta(days=1)).strftime('%y%m%d'))
