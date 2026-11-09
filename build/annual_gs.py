# -*- coding: utf-8 -*-
"""26년 연간 집계표 — 구글 시트(ERP 자동 조회) 한 원천으로 26SS 신상·26FW 신상·러닝을 '품번은 한 자리에만' 규칙으로 나눠 센다 (2026-11-09 사용자 '2단계').

왜
  26SS 판매자료 카드(NAS 26SS 스타일맵, 1주 늦음)와 26FW 카드를 더하면 ① 러닝 구품번 10개(SS 러닝 ∩ FW 구품번) ② 26SS 신상 10개(FW 러닝 구품번으로 재등장)
  ③ 시즌이 봄→가을로 바뀐 2품번(GF1MSH004·GF1LSS506) 이 두 번 세어지고, 기준일도 한 주 어긋났다.

귀속 규칙(품번 하나 = 묶음 하나)
  · 26SS 신상 = ERP 년도 2026 · 시즌 봄/여름 (사은품 등 '기타' 품목 제외)        — 시즌 전체(1/1~기준일) 값 = 생판재(ST) 그대로
  · 26FW 신상 = ERP 년도 2026 · 시즌 가을/겨울 (= 26FW 집계표 그대로, 여기서는 안 만든다)
  · 러닝 신품번 = 시즌 가을(러닝)/겨울(러닝) (GF7·GF8)                             — 26FW 집계표 러닝 행 그대로
  · 러닝 구품번 = '시즌별 러닝아이템' 탭의 모든 시즌 열(24SS~26SS, 26FW-과거품번) 합집합 − 2026 신상 품번
                 → **1/1~ 기준 한 번만**: 입고(투입) = 12/31 시점재고 + 2026 생산, 판매 = 1/1~ 판매(기간), 재고 = 현재고
                 (탭 '26년생판재(1.1~,ST)' = 기간 1/1~기준일 조회, '시점재고(~25.12.31,ST)' = 작년 12/31 시점 조회)
    · 그중 26SS 목록 31개 = 'SS 러닝'(26SS 기간 화면의 러닝), 나머지 = '구품번 추가'(연간에서만 더함)
  · 25SS 동시점(전년) = '26FW전넌비교(~동시점,생판재)' 탭(작년 같은 날 시점, 년도 2024·2025)의 2025 봄/여름 행

산출: ss_data.load() 와 같은 모양의 SS(block1·newCat·seasonCat·genderItem·genderCat·runList·run·end·prevEnd) + ANN(귀속 결과·러닝 상세·검증).
"""
import collections
import datetime
import re

ITEMS = ('OUTER', 'TOP', 'BOTTOM', 'ACC')
GRP_OF = {'OUTER': 'OUTER', 'KNIT': 'TOP', 'INNER (W/O KNIT)': 'TOP', 'BOTTOM': 'BOTTOM', 'ACC': 'ACC'}
SS_SEASONS = ('봄', '여름')
RUN_SEASON_COLS = ('24SS', '24FW', '25SS', '25FW', '26SS', '26FW - 과거품번')
GIFT_EXTRA_25SS = ('WBE1L95512',)
SS_EXCLUDE = ('QBF1U98001',)
FW_GIFT = ('GF3UFG101', 'GF4LFG512', 'GF4LFG513', 'GF4UFG101', 'GF4UFG409', 'GF4UFG410')      # 26FW 카드·집계표가 빼는 사은품(build_salecards 와 같은 목록)
FW_SEASONS = ('가을', '겨울')
KEYS = ('styles', 'sku', 'plan', 'inM', 'inSku', 'qtyIn', 'planCost', 'cost', 'tag', 'sales', 'salesTag', 'salesAmt', 'salesCost', 'wkQty', 'wkAmt', 'stock')


def t(v): return '' if v is None else str(v).strip()


def num(v): return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else 0.0


def derive(v):
    """ss_data.derive 와 같은 규칙(원가 VAT 포함 표시값·마크업·비율)."""
    cost, pc, sc, tag, tg = v.get('cost'), v.get('planCost'), v.get('salesCost'), v.get('tag'), v.get('salesTag')
    v['costV'] = round(cost * 1.1) if cost else None
    v['planCostV'] = round(pc * 1.1) if pc else None
    v['salesCostV'] = round(sc * 1.1) if sc else None
    v['mu'] = round(tag / cost / 1.1, 2) if tag and cost else None
    v['muPlan'] = round(tag / pc / 1.1, 2) if tag and pc else None
    v['muReal'] = round(v['salesAmt'] / sc / 1.1, 2) if v.get('salesAmt') and sc else None
    v['sellRate'] = (v.get('sales') or 0) / v['qtyIn'] if v.get('qtyIn') else None
    v['disc'] = 1 - (v.get('salesAmt') or 0) / tg if tg else None
    v['inRate'] = (v.get('qtyIn') or 0) / v['plan'] if v.get('plan') else None
    return v


class Psi:
    """생판재(스타일 단위) 격자 — 머리글 2줄을 (묶음, 항목) 이름으로 읽어 열을 찾는다. 앞에 플래그 열이 몇 개 붙어 있어도 된다."""
    COLS = {'year': ('년도',), 'season': ('시즌',), 'sex': ('성별',), 'it': ('품목',), 'bok': ('복종',), 'item': ('아이템',), 'style': ('스타일',), 'name': ('스타일명',),
            'unitCost': ('원가(-)',), 'tagPrice': ('UP-TAG가',),
            'planAll': ('기획', '수량'), 'planAllCost': ('기획', '원가금액'),                      # 작지(최초+리오더)
            'plan': ('최초기획', '수량'), 'planTag': ('최초기획', 'UP-TAG가금액'), 'planCost': ('최초기획', '원가금액'),
            'prodAll': ('생산', '수량'), 'prodAllTag': ('생산', 'UP-TAG가금액'), 'prodAllCost': ('생산', '원가금액'),
            'qtyIn': ('기획생산', '수량'), 'tag': ('기획생산', 'UP-TAG가금액'), 'cost': ('기획생산', '원가금액'),
            'etc': ('기타', '수량'),
            'wkQty': ('판매(최근)', '수량'), 'wkAmt': ('판매(최근)', '판매금액'),
            'sales': ('판매(기간)', '수량'), 'salesTag': ('판매(기간)', 'UP-TAG가금액'), 'salesAmt': ('판매(기간)', '판매금액'), 'salesCost': ('판매(기간)', '원가금액'),
            'stock': ('재고', '전체'), 'stockTag': ('재고', '금액'), 'stockCost': ('재고', '원가금액')}

    def __init__(self, grid, label):
        self.label = label
        h1 = [t(v) for v in grid[0]]; h2 = [t(v) for v in grid[1]]
        n = max(len(h1), len(h2)); h1 += [''] * (n - len(h1)); h2 += [''] * (n - len(h2))
        grp, g = [], ''
        for a in h1:
            if a: g = a.split('|')[0] if '|' in a else a
            grp.append(g)
        self.idx = {}
        for key, spec in self.COLS.items():
            if len(spec) == 1:
                j = next((i for i in range(n) if h2[i] == spec[0] or (not h2[i] and h1[i] == spec[0])), None)
            else:
                j = next((i for i in range(n) if grp[i] == spec[0] and h2[i] == spec[1]), None)
            if j is None: raise SystemExit("%s: 머리글 %s 을 못 찾음 (h1=%s…)" % (label, '/'.join(spec), h1[:20]))
            self.idx[key] = j
        si = self.idx['style']
        self.rows = {}
        for r in grid[2:]:
            code = t(r[si]).upper() if len(r) > si else ''
            if not code or code == '스타일' or not re.match(r'^[A-Z0-9]{8,}$', code): continue
            if code in self.rows: raise SystemExit('%s: 품번 %s 가 두 번 있음' % (label, code))
            self.rows[code] = r

    def g(self, code, key):
        r = self.rows.get(code)
        if r is None: return None
        j = self.idx[key]
        return r[j] if j < len(r) else None

    def n(self, code, key): return num(self.g(code, key))


def read_tab_by_name(S, gsheet, name_part, norm):
    title, tabs = S.meta()
    name = next((k for k in tabs if name_part.replace(' ', '') in k.replace(' ', '')), None)
    if not name: return None, None
    q = "'" + name.replace("'", "''") + "'"
    v = S._req('GET', '/values/' + gsheet._enc(q) + '?valueRenderOption=UNFORMATTED_VALUE').get('values', [])
    return name, norm(v)


BASIS = 'all'      # 'all' = 기획 작지·생산 전체(리오더 포함, 26SS 종합 시트와 같음) / 'plan0' = 최초기획·기획생산(리오더 제외, 26FW 카드·집계표 규칙)


def metrics(psi, codes, co_sku=None, basis=None):
    """품번 묶음의 시즌 지표 합(ss_data 의 cur 모양). basis 에 따라 기획·입고를 리오더 포함/제외로 센다."""
    basis = basis or BASIS
    kp, kpc, ki, kt, kc = ('planAll', 'planAllCost', 'prodAll', 'prodAllTag', 'prodAllCost') if basis == 'all' else ('plan', 'planCost', 'qtyIn', 'tag', 'cost')
    a = {k: 0.0 for k in KEYS}
    for c in codes:
        if c not in psi.rows: continue
        a['plan'] += psi.n(c, kp); a['planCost'] += psi.n(c, kpc); a['qtyIn'] += psi.n(c, ki); a['tag'] += psi.n(c, kt); a['cost'] += psi.n(c, kc)
        for k in ('wkQty', 'wkAmt', 'sales', 'salesTag', 'salesAmt', 'salesCost', 'stock'):
            a[k] += psi.n(c, k)
        if psi.n(c, kp) > 0: a['styles'] += 1
        if psi.n(c, ki) > 0: a['inM'] += 1
        if co_sku:
            a['sku'] += co_sku.get(c, (0, 0))[0]; a['inSku'] += co_sku.get(c, (0, 0))[1]
    return derive(a)


def run_metrics(ytd, dec, cur, codes):
    """러닝 구품번: 1/1~ 기준. 입고(투입) = 12/31 재고 + 2026 생산(= 현재 누적 생산 − 12/31 누적 생산) · 판매 = 1/1~ (= 현재 누계 − 12/31 누계) · 재고 = 현재고. 기획은 없다.
    ERP 의 '기간' 조건은 판매(기간)에만 걸리고 생산·재고는 누적값 그대로라(2026-11-09 확인) 두 시점의 차로 올해 값을 만든다. 1/1~ 탭의 판매(기간)은 대조용."""
    a = {k: 0.0 for k in KEYS}; a['plan'] = None; a['planCost'] = None
    resid, ytd_diff = [], []
    for c in codes:
        if c not in cur.rows: continue
        d = lambda k: cur.n(c, k) - dec.n(c, k)
        d_stk, d_tag, d_cost = dec.n(c, 'stock'), dec.n(c, 'stockTag'), dec.n(c, 'stockCost')
        a['qtyIn'] += d_stk + d('prodAll'); a['tag'] += d_tag + d('prodAllTag'); a['cost'] += d_cost + d('prodAllCost')
        for k in ('sales', 'salesTag', 'salesAmt', 'salesCost'): a[k] += d(k)
        for k in ('wkQty', 'wkAmt'): a[k] += cur.n(c, k)
        a['stock'] += cur.n(c, 'stock')
        if d_stk + d('prodAll') > 0: a['inM'] += 1
        a['styles'] += 1
        resid.append((round(d_stk + d('prodAll') - d('sales') - d('etc') - cur.n(c, 'stock')), c))
        if ytd is not None and c in ytd.rows: ytd_diff.append((round(ytd.n(c, 'sales') - d('sales')), c))
    return derive(a), resid, ytd_diff


def build(GS, verbose=True):
    import gsource
    g = GS['grids']
    cur = Psi(g[gsource.TABS['st'][1]], '생판재(ST)')
    py = Psi(g[gsource.TABS['back'][1]], '전년비교')
    gsheet, S = gsource._client()
    nm_y, ytd_g = read_tab_by_name(S, gsheet, '26년생판재(1.1~', gsource._norm)
    nm_d, dec_g = read_tab_by_name(S, gsheet, '시점재고(~25.12.31', gsource._norm)
    if not ytd_g or not dec_g: raise SystemExit("연간 집계용 탭이 아직 없습니다: 26년생판재(1.1~,ST) / 시점재고(~25.12.31,ST) — erp_client.py --screen psi_ytd/stk_dec --gsheet")
    ytd = Psi(ytd_g, nm_y); dec = Psi(dec_g, nm_d)
    ref = GS['ref']
    # 머리글에서 조회 기간 확인(최근판매 머리글 등은 안 쓰고, 행 수와 기준일만 로그)
    # --- 러닝 목록(시즌별 러닝아이템 원탭) ---
    nm_r, run_raw = read_tab_by_name(S, gsheet, '러닝아이템', gsource._norm)
    hdr = [t(v) for v in run_raw[1]]
    season_codes = {}
    for lab in RUN_SEASON_COLS:
        if lab in hdr:
            j = hdr.index(lab)
            season_codes[lab] = [t(r[j]).upper() for r in run_raw[2:] if len(r) > j and t(r[j])]
    i_new = hdr.index('26FW') if '26FW' in hdr else None
    run_new = [t(r[i_new]).upper() for r in run_raw[2:] if i_new is not None and len(r) > i_new and t(r[i_new])]
    # --- 품번 귀속 ---
    yr = lambda c: t(cur.g(c, 'year')); se = lambda c: t(cur.g(c, 'season')); it = lambda c: GRP_OF.get(t(cur.g(c, 'it')))
    all_codes = list(cur.rows)
    ss_new = [c for c in all_codes if yr(c) == '2026' and se(c) in SS_SEASONS and it(c) and c not in SS_EXCLUDE]
    ss_gift = [c for c in all_codes if yr(c) == '2026' and se(c) in SS_SEASONS and (not it(c) or c in SS_EXCLUDE)]
    fw_new = [c for c in all_codes if yr(c) == '2026' and se(c) in ('가을', '겨울')]
    run_new_erp = [c for c in all_codes if yr(c) == '2026' and '러닝' in se(c)]
    new2026 = set(c for c in all_codes if yr(c) == '2026' and '러닝' not in se(c))
    old_union = []
    for lab in RUN_SEASON_COLS:
        for c in season_codes.get(lab, []):
            if c not in old_union: old_union.append(c)
    run_old = [c for c in old_union if c not in new2026]
    old_as_new = [c for c in old_union if c in new2026]                  # 러닝 목록에 있지만 2026 신상 → 신상으로만 센다
    missing = [c for c in run_old if c not in cur.rows]
    run_old = [c for c in run_old if c in cur.rows]
    ss_run = [c for c in season_codes.get('26SS', []) if c in run_old]
    old_add = [c for c in run_old if c not in ss_run]
    # 유일성 검사
    buckets = {'26SS 신상': ss_new, '26FW 신상': fw_new, '러닝 신품번': run_new_erp, '러닝 구품번': run_old}
    seen = collections.Counter(c for v in buckets.values() for c in v)
    dup = sorted(c for c, n in seen.items() if n > 1)
    # --- SKU(컬러) 수: 주간판매(CO) ---
    co = g[gsource.TABS['co'][1]]
    h2 = [t(v) for v in co[1]]
    ci_s, ci_p, ci_in = h2.index('스타일'), h2.index('작지' if BASIS == 'all' else '최초기획'), h2.index('입고' if BASIS == 'all' else '기획입고')
    co_sku = {}
    for r in co[2:]:
        c = t(r[ci_s]).upper() if len(r) > ci_s else ''
        if not re.match(r'^[A-Z0-9]{8,}$', c): continue
        a = co_sku.setdefault(c, [0, 0])
        if num(r[ci_p]) > 0: a[0] += 1
        if num(r[ci_in]) > 0: a[1] += 1
    # --- 25SS 동시점(전년) 품번 ---
    py_year = lambda c: t(py.g(c, 'year')); py_se = lambda c: t(py.g(c, 'season')); py_it = lambda c: GRP_OF.get(t(py.g(c, 'it')))
    py_codes = [c for c in py.rows if py_year(c) == '2025' and py_se(c) in SS_SEASONS and py_it(c)]
    gifts25 = [c for c in py_codes if ('U98' in c and not py.n(c, 'salesAmt')) or c in GIFT_EXTRA_25SS]
    py_codes = [c for c in py_codes if c not in gifts25]

    def rows_for(codes, psi, key_fn):
        out = collections.defaultdict(list)
        for c in codes: out[key_fn(c)].append(c)
        return out

    def new_block(codes, seasons, py_codes_, with_run):
        """신상 묶음(codes, 시즌 목록) → block1/newCat/seasonCat/genderItem/genderCat (ss_data 모양). with_run 이면 러닝(26SS 31) 행도 붙인다."""
        b1 = []
        by_se_it = rows_for(codes, cur, lambda c: (se(c), it(c)))
        py_by = rows_for(py_codes_, py, lambda c: (py_se(c), py_it(c)))
        ttl_codes, pyt_codes = [], []
        for s_ in seasons:
            s_codes, p_codes = [], []
            for item in ITEMS:
                cs, ps = by_se_it.get((s_, item), []), py_by.get((s_, item), [])
                s_codes += cs; p_codes += ps
                b1.append({'season': s_, 'item': item, 'mode': 'new', 'kind': '항목', 'cur': metrics(cur, cs, co_sku), 'py': metrics(py, ps)})
            b1.append({'season': s_, 'item': 'TTL', 'mode': 'new', 'kind': '%s TTL' % s_, 'cur': metrics(cur, s_codes, co_sku), 'py': metrics(py, p_codes)})
            ttl_codes += s_codes; pyt_codes += p_codes
        ttl_new_ = metrics(cur, ttl_codes, co_sku)
        b1.append({'season': '', 'item': 'TTL', 'mode': 'new', 'kind': '신상품 TTL', 'cur': ttl_new_, 'py': metrics(py, pyt_codes)})
        m_run_ = None
        if with_run:
            run_by = rows_for(ss_run, cur, it)
            for item in ITEMS:
                m, rs, yd = run_metrics(ytd, dec, cur, run_by.get(item, [])); resid_all.extend(rs); ytd_all.extend(yd)
                b1.append({'season': '', 'item': item, 'mode': 'run', 'kind': '항목', 'cur': m, 'py': None})
            m_run_, _, _ = run_metrics(ytd, dec, cur, ss_run)
            b1.append({'season': '', 'item': 'TTL', 'mode': 'run', 'kind': '러닝 TTL', 'cur': m_run_, 'py': None})
            tot = {k: (ttl_new_.get(k) or 0) + ((m_run_.get(k) or 0) if k not in ('plan', 'planCost') else 0) for k in KEYS}
            b1.append({'season': '', 'item': 'TTL', 'mode': 'run', 'kind': '전체 TTL', 'cur': derive(tot), 'py': None})
        cat = lambda c: t(cur.g(c, 'item')); py_cat = lambda c: t(py.g(c, 'item'))
        new_cat_, season_cat_ = [], []
        for s_ in (None,) + tuple(seasons):
            cur_by = rows_for([c for c in codes if s_ is None or se(c) == s_], cur, lambda c: (it(c), cat(c)))
            py_by2 = rows_for([c for c in py_codes_ if s_ is None or py_se(c) == s_], py, lambda c: (py_it(c), py_cat(c)))
            for key in sorted(set(cur_by) | set(py_by2), key=lambda k: (ITEMS.index(k[0]), k[1])):
                rec = {'item': 'ACCESSORY' if key[0] == 'ACC' else key[0], 'cat': key[1], 'code': '', 'cur': metrics(cur, cur_by.get(key, []), co_sku), 'py': metrics(py, py_by2.get(key, []))}
                if not any((rec['cur'].get(k) or 0) for k in ('plan', 'qtyIn', 'sales')) and not any((rec['py'].get(k) or 0) for k in ('qtyIn', 'sales')): continue
                if s_: rec['season'] = s_; season_cat_.append(rec)
                else: new_cat_.append(rec)
        sex = lambda c: '여성' if t(cur.g(c, 'sex')) == '여성' else '남성'
        gi_, gc_ = [], []
        for gnd in ('여성', '남성'):
            gc = [c for c in codes if sex(c) == gnd]
            for item in ITEMS:
                gi_.append({'gender': gnd, 'item': item, 'cur': metrics(cur, [c for c in gc if it(c) == item], co_sku), 'py': None})
            gi_.append({'gender': gnd, 'item': 'TOTAL', 'cur': metrics(cur, gc, co_sku), 'py': None})
            for key, cs in sorted(rows_for(gc, cur, lambda c: (it(c), cat(c))).items(), key=lambda k: (ITEMS.index(k[0][0]), k[0][1])):
                gc_.append({'gender': gnd, 'item': 'ACCESSORY' if key[0] == 'ACC' else key[0], 'cat': key[1], 'code': '', 'cur': metrics(cur, cs, co_sku), 'py': None})
        return b1, new_cat_, season_cat_, gi_, gc_, ttl_new_, m_run_

    resid_all, ytd_all = [], []
    block1, new_cat, season_cat, gender_item, gender_cat, ttl_new, m_run = new_block(ss_new, SS_SEASONS, py_codes, True)
    # 26FW 신상(같은 기준) — 연간 탭에서 26FW 집계표 대신 쓴다(리오더 포함/제외 기준을 SS 와 맞추려고)
    fw_codes = [c for c in fw_new if it(c) and t(cur.g(c, 'item')) != '사은품' and c not in FW_GIFT]
    fw_gift = [c for c in fw_new if c not in fw_codes]
    py_fw = [c for c in py.rows if py_year(c) == '2025' and py_se(c) in FW_SEASONS and py_it(c) and not ('U98' in c and not py.n(c, 'salesAmt'))]
    fb1, fnc, fsc, fgi, fgc, fw_ttl, _ = new_block(fw_codes, FW_SEASONS, py_fw, False)
    fw_block = {'block1': fb1, 'newCat': fnc, 'seasonCat': fsc, 'genderItem': fgi, 'genderCat': fgc, 'gift': [(c, t(cur.g(c, 'name'))) for c in fw_gift], 'n': len(fw_codes)}
    # --- 연간 러닝 상세(ANN) ---
    def detail(c, src):
        d = lambda k: cur.n(c, k) - dec.n(c, k)
        d_stk, p_in = dec.n(c, 'stock'), d('prodAll')
        return {'code': c, 'name': t(cur.g(c, 'name')), 'item': it(c) or '', 'src': src, 'seasons': [lab for lab in RUN_SEASON_COLS if c in season_codes.get(lab, [])],
                'decStock': round(d_stk), 'newIn': round(p_in), 'qtyIn': round(d_stk + p_in), 'sales': round(d('sales')), 'stock': round(cur.n(c, 'stock')),
                'cost': round(dec.n(c, 'stockCost') + d('prodAllCost')), 'salesAmt': round(d('salesAmt')), 'wkQty': round(cur.n(c, 'wkQty')),
                'resid': round(d_stk + p_in - d('sales') - d('etc') - cur.n(c, 'stock')), 'ytdSalesDiff': round(ytd.n(c, 'sales') - d('sales')) if c in ytd.rows else None,
                'lifeSales': round(cur.n(c, 'sales')), 'lifeIn': round(cur.n(c, 'prodAll'))}
    run_items = {}
    for item in ITEMS + ('TOTAL',):
        f = (lambda c: True) if item == 'TOTAL' else (lambda c, item=item: it(c) == item)
        a, _, _ = run_metrics(ytd, dec, cur, [c for c in ss_run if f(c)]); b, rb, yb = run_metrics(ytd, dec, cur, [c for c in old_add if f(c)])
        if item != 'TOTAL': resid_all += rb; ytd_all += yb
        n_codes = [c for c in run_new_erp if f(c)]
        run_items[item] = {'ssRun': a, 'oldAdd': b, 'fwNew': metrics(cur, n_codes, co_sku), 'nSs': len([c for c in ss_run if f(c)]), 'nAdd': len([c for c in old_add if f(c)]), 'nNew': len(n_codes)}
    resid_max = max((abs(r[0]) for r in resid_all), default=0)
    ytd_max = max((abs(r[0]) for r in ytd_all), default=0)
    # 옛 방식(카드 단순 합산)이라면 두 번 세었을 품번 — 설명용
    legacy = {'ssRunAndFwOld': sorted(set(season_codes.get('26SS', [])) & set(season_codes.get('26FW - 과거품번', []))),
              'ssNewAsFwOld': sorted(set(season_codes.get('26FW - 과거품번', [])) & set(ss_new)),
              'seasonMoved': sorted(c for c in old_as_new if c in fw_new) + sorted(c for c in fw_new if c[:3] in ('GF1', 'GF2'))}
    legacy['seasonMoved'] = sorted(set(legacy['seasonMoved']))
    ann = {'ref': ref.isoformat(), 'ytdTab': nm_y, 'decTab': nm_d, 'gen': datetime.datetime.now().strftime('%Y-%m-%d %H:%M'),
           'buckets': {k: len(v) for k, v in buckets.items()}, 'bucketCodes': {'러닝 구품번': run_old, '러닝 신품번': run_new_erp},
           'ssGift': [(c, t(cur.g(c, 'name'))) for c in ss_gift], 'gifts25': gifts25, 'oldAsNew': old_as_new, 'missing': missing, 'dup': dup,
           'nSsRun': len(ss_run), 'nOldAdd': len(old_add), 'nRunNew': len(run_new_erp), 'runNewList': run_new,
           'items': run_items, 'oldRows': [detail(c, '26SS 러닝' if c in ss_run else '구품번 추가') for c in run_old],
           'basis': BASIS, 'fw': fw_block,
           'residMax': resid_max, 'residTop': sorted(resid_all, key=lambda x: -abs(x[0]))[:5], 'ytdSalesDiffMax': ytd_max, 'legacy': legacy}
    ss = {'file': '구글 시트 ERP 자동 조회 (%s 기준)' % ref, 'end': ref.isoformat(), 'prevEnd': (ref.replace(year=ref.year - 1)).isoformat(), 'source': 'gsheet',
          'gifts25': ['%s %s' % (c, t(py.g(c, 'name'))) for c in gifts25],
          'gifts25Sum': {'n': len(gifts25), 'qtyIn': round(sum(py.n(c, 'qtyIn') for c in gifts25)), 'costV': round(sum(py.n(c, 'cost') for c in gifts25) * 1.1)},
          'gifts26': [{'code': c, 'name': t(cur.g(c, 'name')), 'season': se(c), 'qtyIn': round(cur.n(c, 'qtyIn')), 'costV': round(cur.n(c, 'cost') * 1.1), 'rows': 0} for c in ss_gift],
          'block1': block1, 'newCat': new_cat, 'seasonCat': season_cat, 'genderItem': gender_item, 'genderCat': gender_cat,
          'runList': ss_run, 'run': {'ssList': len(ss_run), 'fwList': len(season_codes.get('26FW - 과거품번', [])), 'overlap': 0, 'overlapStyles': [], 'overlapQty': 0, 'overlapQtyLo': 0, 'since': ''}, 'check': []}
    if verbose:
        b = ann['buckets']
        print('연간(구글 시트 단일 원천, 기준 %s): 26SS 신상 %d · 26FW 신상 %d · 러닝 신품번 %d · 러닝 구품번 %d(SS 러닝 %d + 추가 %d) · 겹침 %s · 러닝목록에 있지만 2026 신상인 품번 %d · ST 에 없는 구품번 %s'
              % (ref, b['26SS 신상'], b['26FW 신상'], b['러닝 신품번'], b['러닝 구품번'], len(ss_run), len(old_add), dup or '없음', len(old_as_new), missing or '없음'))
        print('  1/1~ 탭 %s %d품번 · 12/31 탭 %s %d품번 · 재고 항등식 잔차 최대 %d장(%s) · 1/1~ 탭 판매 vs (현재−12/31) 최대 차 %d장 · 25SS 사은품 제외 %d · 26SS 기타 제외 %d'
              % (nm_y, len(ytd.rows), nm_d, len(dec.rows), resid_max, ann['residTop'][:3], ytd_max, len(gifts25), len(ss_gift)))
        tn = ttl_new; tr = m_run
        print('  26FW 신상(같은 기준, 사은품 %d 제외): 입고 %s · 누계 %s · 주간 %s · 원가(VAT+) %s' % (len(fw_gift), format(round(fw_ttl['qtyIn']), ','), format(round(fw_ttl['sales']), ','), format(round(fw_ttl['wkQty']), ','), format(fw_ttl['costV'] or 0, ',')))
        print('  26SS 신상: 입고 %s · 누계 %s · 주간 %s · 원가(VAT+) %s | 러닝(SS 31, 1/1~): 투입 %s · 판매 %s · 재고 %s | 구품번 추가 %d: 투입 %s · 판매 %s'
              % (format(round(tn['qtyIn']), ','), format(round(tn['sales']), ','), format(round(tn['wkQty']), ','), format(tn['costV'] or 0, ','),
                 format(round(tr['qtyIn']), ','), format(round(tr['sales']), ','), format(round(tr['stock']), ','), len(old_add),
                 format(round(run_items['TOTAL']['oldAdd']['qtyIn']), ','), format(round(run_items['TOTAL']['oldAdd']['sales']), ',')))
    return ss, ann
