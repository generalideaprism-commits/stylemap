# -*- coding: utf-8 -*-
"""구글 스프레드시트('대쉬보드PIP업데이트파일')를 스타일맵 엑셀처럼 읽는다 — 빌드의 원천을 엑셀에서 구글 시트로 바꾸는 어댑터.

  원자료 탭(ERP 직접 조회 → erp-http 가 매주 씀)
    26FW생판재(ST)                 A 신상 · B 러닝여부 · C 신규품번 · D IMC(테마) · E~ 원본
    26FW주간판매(CO)               A 신상 · B 러닝여부 · C 스타일 수 · D~ 원본
    26FW전넌비교(~동시점,생판재)    A 러닝여부 · B~ 원본
  사용자가 올려 둔 탭
    시점재고(~26.07.31,컬러생판재) · 시점재고(~25.07.31, 컬러생판재)   — 엑셀 시트와 같은 배치
    시즌별 러닝아이템              — G '26FW'(신품번) · H '26FW - 과거품번' · M·N 품번→테마 · B~F 시즌별 러닝 목록

  돌려주는 것: 엑셀 스타일맵(구조조정)과 같은 시트 이름·같은 열 배치의 격자. 그래서 빌드와 종합 계산(jonghap)이 그대로 돈다.
    · 생판재 탭은 IMC 열(D)이 엑셀에 없는 열이라 격자에서는 빼고, 테마는 따로 {품번: 테마} 로 넘긴다.
    · 종합(2)는 원자료 격자로 계산해 넣는다 (엑셀 수식과 같은 규칙).
  읽기 전용이다 — 이 모듈은 시트에 아무것도 쓰지 않는다.
"""
import datetime, os, re, sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
ERP_HTTP = os.environ.get('ERP_HTTP_DIR') or os.path.join(os.path.dirname(ROOT), 'erp-http')

# 구글 탭 → 엑셀 시트 이름. gid 로 찾는다(탭 이름을 바꿔도 따라간다).
TABS = {
    'st':    (1334512621, '생판재(ST)'),
    'co':    (1180665516, '주간판매(CO)'),
    'back':  (1278219249, '전년비교BACK(~동시점,생판재)'),
    'stk26': (1413884224, '시점재고(~26.07.31, 컬러생판재)'),
    'stk25': (1422508689, '시점재고(~25.07.31, 컬러생판재)'),
    'run':   (1537621745, '26FW 러닝스타일 리스트'),
}


class GridSheet:
    """openpyxl 읽기 전용 시트처럼 쓰는 격자."""
    def __init__(self, title, rows):
        self.title, self.rows = title, rows

    def iter_rows(self, min_row=1, max_row=None, min_col=1, max_col=None, values_only=True):
        end = len(self.rows) if max_row is None else min(max_row, len(self.rows))
        for i in range(max(min_row, 1) - 1, end):
            r = self.rows[i]
            yield tuple(r[min_col - 1:max_col] if (max_col or min_col > 1) else r)

    def calculate_dimension(self):
        return 'A1:%s%d' % (_col(max((len(r) for r in self.rows), default=1) - 1), len(self.rows))


class GridWB:
    def __init__(self, sheets):
        self._s = {s.title: s for s in sheets}
        self.sheetnames = [s.title for s in sheets]

    def __getitem__(self, name): return self._s[name]
    def __contains__(self, name): return name in self._s


def _col(i):
    s = ''; i += 1
    while i: i, r = divmod(i - 1, 26); s = chr(65 + r) + s
    return s


def _t(v): return '' if v is None else str(v).strip()


def _norm(rows):
    """구글 값 → 엑셀 읽기와 같은 모양: 빈 글자는 None, 행 길이 맞춤, 정수 모양 실수는 그대로 둔다."""
    w = max((len(r) for r in rows), default=0)
    return [[(None if v == '' else v) for v in r] + [None] * (w - len(r)) for r in rows]


def _client():
    sys.path.insert(0, os.path.join(ERP_HTTP, 'lib')); sys.path.insert(0, ERP_HTTP)
    import gsheet
    cfg = gsheet.load_config()
    return gsheet, gsheet.Sheets(cfg)


def load(verbose=True, season='FW'):
    """구글 시트 → {'wb': GridWB, 'themes': {품번: 테마}, 'ref': 기준 일요일(date), 'title': 파일 제목, 'info': [로그 줄]}
    season='SS' 이면 러닝 리스트를 '시즌별 러닝아이템' 26SS 열(구품번 31, 짝 없음)로 만들고 컬러생판재 두 시점 탭(현재·12/31)도 읽는다."""
    import jonghap
    gsheet, S = _client()
    title, tabs = S.meta()
    by_gid = {v['sheetId']: k for k, v in tabs.items()}
    info = []

    def read(key):
        gid, xname = TABS[key]
        name = by_gid.get(gid)
        if name is None: raise SystemExit("구글 시트 '%s' 에 %s 탭(gid %s)이 없습니다" % (title, xname, gid))
        q = "'" + name.replace("'", "''") + "'"
        v = S._req('GET', '/values/' + gsheet._enc(q) + '?valueRenderOption=UNFORMATTED_VALUE').get('values', [])
        return name, _norm(v)

    def read_name(part):
        name = next((k for k in tabs if part.replace(' ', '') in k.replace(' ', '')), None)
        if name is None: return None, None
        q = "'" + name.replace("'", "''") + "'"
        v = S._req('GET', '/values/' + gsheet._enc(q) + '?valueRenderOption=UNFORMATTED_VALUE').get('values', [])
        return name, _norm(v)

    # --- 생판재: IMC 열(D)을 테마로 떼어 내고 엑셀 배치(A~C 플래그 + 원본)로 ---
    nm, st = read('st')
    if _t(st[2][3]) != 'IMC' or _t(st[1][4 + 12]) != '스타일':
        raise SystemExit("%s 탭의 배치가 예상과 다릅니다 (3행 D열 'IMC', 2행 Q열 '스타일' 이어야 함): D3=%r Q2=%r" % (nm, st[2][3], st[1][16]))
    si = 16
    themes = {}
    for r in st[3:]:
        code, th = _t(r[si]), _t(r[3])
        if code and th and th not in ('0', '#N/A'): themes[code] = th
    st_x = [r[:3] + r[4:] for r in st]
    info.append('%s: %d스타일 · 신상 %d · 러닝여부 %d · 테마(IMC) %d건' % (nm, sum(1 for r in st[3:] if _t(r[si])), sum(1 for r in st[3:] if r[0] == 1),
                                                              sum(1 for r in st[3:] if r[1] == 1), len(themes)))
    # --- 주간판매 ---
    nm, co = read('co')
    h2 = [_t(v) for v in co[1]]
    if _t(co[2][0]) != '신상' or '스타일' not in h2: raise SystemExit('%s 탭의 배치가 예상과 다릅니다 (3행 A열 신상 · 2행에 스타일)' % nm)
    w1 = next((v for v in h2 if re.match(r'^W1\(', v)), '')
    m = re.match(r'^W1\((\d{2})/(\d{2}) -(\d{2})\)', w1)
    if not m: raise SystemExit('%s 탭에서 W1 주간 머리글을 찾지 못했습니다: %r' % (nm, w1))
    info.append('%s: %d행 · 신상 %d · 주간 %s' % (nm, len(co) - 3, sum(1 for r in co[3:] if r[0] == 1), w1))
    # --- 전년 동시점 · 시점재고 ---
    nm_b, back = read('back'); nm_6, stk26 = read('stk26'); nm_5, stk25 = read('stk25')
    # 전년 탭 A열 '러닝여부': 엑셀은 글자 '러닝', 구글 탭은 1 로 적혀 있다(2026-11-06). 종합 수식은 '러닝' 글자와 비교하므로 맞춰 준다.
    n_fix = 0
    for r in back[3:]:
        if r and r[0] in (1, 1.0, '1'): r[0] = '러닝'; n_fix += 1
    if n_fix: info.append("%s: A열 러닝여부 1 → '러닝' 으로 읽음 (%d행)" % (nm_b, n_fix))
    # 깨진 칸 검사: 엑셀에서 열 너비가 좁은 채로 복사하면 큰 숫자가 '#######' 글자로 붙는다(2026-11-06 시점재고 두 탭에서 실제 발생).
    # 그대로 계산하면 그 칸이 합계에서 빠져 금액이 크게 틀어지므로, 깨진 칸이 있으면 멈춘다.
    # 임시 우회(점검용): 환경변수 STYLEMAP_STOCK_XLSX 에 엑셀 스타일맵 경로를 주면 시점재고 두 시트만 그 파일에서 읽는다.
    broken = {}
    for nm, g in ((nm_b, back), (nm_6, stk26), (nm_5, stk25), ('생판재', st), ('주간판매', co)):
        n = sum(1 for r in g[3:] for v in r if isinstance(v, str) and len(v) >= 3 and set(v) == {'#'})
        if n: broken[nm] = n
    alt = os.environ.get('STYLEMAP_STOCK_XLSX')
    if broken and alt and all(k in (nm_6, nm_5) for k in broken):
        import openpyxl
        wx = openpyxl.load_workbook(alt, read_only=True, data_only=True)
        stk26 = [list(r) for r in wx[TABS['stk26'][1]].iter_rows(values_only=True)]
        stk25 = [list(r) for r in wx[TABS['stk25'][1]].iter_rows(values_only=True)]
        info.append('[임시 우회] 시점재고 두 시트를 엑셀에서 읽음: %s (구글 탭에 깨진 칸 %s)' % (os.path.basename(alt), broken))
    elif broken:
        raise SystemExit("구글 시트에 '####' 로 깨진 칸이 있습니다 — 엑셀에서 열 너비가 좁은 채 복사된 값입니다: %s\n"
                         "  해당 탭을 다시 붙여 넣어 주세요(엑셀에서 열 너비를 넓히거나 '값만 붙여넣기' 전에 표시 형식을 '일반'으로)." % broken)
    info.append('%s %d행 · %s %d행 · %s %d행' % (nm_b, len(back) - 3, nm_6, len(stk26) - 3, nm_5, len(stk25) - 3))
    # --- 러닝 목록: 엑셀 '26FW 러닝스타일 리스트' 모양으로 다시 짠다 ---
    nm_r, run = read('run')
    hdr = [_t(v) for v in run[1]]
    try: i_new, i_old = hdr.index('26FW'), hdr.index('26FW - 과거품번')
    except ValueError: raise SystemExit("%s 탭 2행에서 '26FW' / '26FW - 과거품번' 머리글을 찾지 못했습니다: %s" % (nm_r, hdr))
    i_tc, i_tt = (hdr.index('품번'), hdr.index('테마')) if '품번' in hdr and '테마' in hdr else (None, None)
    pairs = [(_t(r[i_old]), _t(r[i_new])) for r in run[2:] if _t(r[i_old]) and _t(r[i_new])]
    if season == 'SS':                                   # 26SS 러닝 = 26SS 열의 구품번 그대로(신품번 없음) → 자기 자신과 짝
        if '26SS' not in hdr: raise SystemExit("%s 탭 2행에 '26SS' 머리글이 없습니다" % nm_r)
        j = hdr.index('26SS')
        pairs = [(_t(r[j]).upper(), _t(r[j]).upper()) for r in run[2:] if len(r) > j and _t(r[j])]
        # 2026-11-09: 26SS 러닝 = 러닝 구품번 전체(annual_gs 와 같은 규칙) — 26SS 열 뒤에 다른 시즌 열의 구품번도 붙인다.
        # 2026 신상(러닝 아님)과 생판재에 없는 품번은 뺀다. 이래야 26SS 카드·26SS 탭·연간의 러닝 품번 수가 같다.
        import annual_gs as _ag
        _psi = _ag.Psi(st_x, '생판재(ST)')
        _new26 = set(c for c in _psi.rows if _t(_psi.g(c, 'year')) == '2026' and '러닝' not in _t(_psi.g(c, 'season')))
        _seen = set(a for a, _ in pairs)
        for lab in _ag.RUN_SEASON_COLS:
            if lab not in hdr: continue
            jj = hdr.index(lab)
            for r in run[2:]:
                c = _t(r[jj]).upper() if len(r) > jj else ''
                if c and c not in _seen: _seen.add(c); pairs.append((c, c))
        _n0 = len(pairs)
        pairs = [(a, b) for a, b in pairs if a not in _new26 and a in _psi.rows]
        info.append('26SS 러닝 목록: 구품번 %d개 (러닝아이템 탭 시즌 열 합집합 %d − 2026 신상·생판재에 없는 품번 %d)' % (len(pairs), _n0, _n0 - len(pairs)))
    y25 = []
    for lab in ('25SS', '25FW'):
        if lab in hdr:
            for r in run[2:]:
                c = _t(r[hdr.index(lab)])
                if c and c not in y25: y25.append(c)
    tmap = [(_t(r[i_tc]), _t(r[i_tt])) for r in run[2:] if i_tc is not None and _t(r[i_tc])]
    rows = [['기존품번', '신규러닝품번', None, None, '품번', '테마', None, None, '25년 러닝', None]]
    body = pairs + ([(n, n) for _, n in pairs] if season == 'FW' else [])      # 엑셀과 같게: 구품번→신품번 뒤에 신품번→신품번
    for i in range(max(len(body), len(tmap), len(y25))):
        a, b = body[i] if i < len(body) else (None, None)
        e, f = tmap[i] if i < len(tmap) else (None, None)
        rows.append([a, b, None, None, e, f, None, None, y25[i] if i < len(y25) else None, '러닝' if i < len(y25) else None])
    info.append('%s: 러닝 구→신 %d쌍 · 품번→테마 %d건 · 25년 러닝 %d개' % (nm_r, len(pairs), len(tmap), len(y25)))
    # --- 기준일 ---
    year = datetime.date.today().year
    ref = datetime.date(year, int(m.group(1)), int(m.group(2))) + datetime.timedelta(days=6)
    if ref > datetime.date.today() + datetime.timedelta(days=14): ref = ref.replace(year=year - 1)      # 1월에 12월 주가 오는 경우
    # --- 종합(2) 계산 ---
    grids = {TABS['st'][1]: st_x, TABS['co'][1]: co, TABS['back'][1]: back, TABS['stk26'][1]: stk26, TABS['stk25'][1]: stk25}
    tpl = jonghap.load_template()
    jong = jonghap.compute(grids, tpl)
    info.append('종합 계산: %s 의 수식 %d개 (틀: %s)' % (tpl['sheet'], sum(1 for v in tpl['cells'].values() if isinstance(v, str) and v.startswith('=')), tpl.get('source', '?')))
    sheets = [GridSheet(TABS['run'][1], rows), GridSheet(tpl['sheet'], jong)] + [GridSheet(n, g) for n, g in grids.items()]
    if verbose:
        print("원천: 구글 시트 '%s' · 기준 일요일 %s" % (title, ref))
        for l in info: print('  ' + l)
    out = {'wb': GridWB(sheets), 'themes': themes, 'ref': ref, 'title': title, 'info': info, 'grids': grids, 'jonghap': jong}
    if season == 'SS':
        # 컬러생판재 두 시점 → {(품번, 칼라): {'stock': 전체 재고, 'prod': 누적 생산, 'sales': 누적 판매(기간)}}
        def color_map(part):
            nm_c, g = read_name(part)
            if not g: return nm_c, None
            h1 = [_t(v) for v in g[0]]; h2 = [_t(v) for v in g[1]]
            grp, cur_g = [], ''
            for a in h1:
                if a: cur_g = a.split('|')[0]
                grp.append(cur_g)
            def col(gname, sub):
                return next(i for i in range(len(h2)) if grp[i] == gname and h2[i] == sub)
            si, ci = h2.index('스타일'), h2.index('칼라')
            i_stk, i_prod, i_sales = col('재고', '전체'), col('생산', '수량'), col('판매(기간)', '수량')
            m = {}
            for r in g[2:]:
                code = _t(r[si]).upper() if len(r) > si else ''
                if not re.match(r'^[A-Z0-9]{8,}$', code): continue
                k = (code, _t(r[ci]))
                a = m.setdefault(k, {'stock': 0.0, 'prod': 0.0, 'sales': 0.0})
                for key, i in (('stock', i_stk), ('prod', i_prod), ('sales', i_sales)):
                    v = r[i] if i < len(r) else None
                    a[key] += float(v) if isinstance(v, (int, float)) else 0.0
            return nm_c, m
        nm1, out['color_now'] = color_map('컬러생판재(현재')
        nm2, out['color_dec'] = color_map('시점재고(~25.12.31,컬러')
        info.append('컬러생판재: 현재 %s %s행 · 12/31 %s %s행' % (nm1, len(out['color_now'] or {}), nm2, len(out['color_dec'] or {})))
        if verbose: print('  ' + info[-1])
    return out


if __name__ == '__main__':
    try: sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception: pass
    sys.path.insert(0, HERE)
    d = load()
    print('시트:', d['wb'].sheetnames)
