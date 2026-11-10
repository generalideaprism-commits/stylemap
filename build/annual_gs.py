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
GRP_OF = {'OUTER': 'OUTER', 'KNIT': 'TOP', 'INNER (W/O KNIT)': 'TOP', 'BOTTOM': 'BOTTOM', 'ACC': 'ACC', 'SHOES': 'ACC', 'SWIM': 'ACC'}     # SHOES·SWIM 은 26FW 집계표(BOKMAP 기본값)처럼 ACC 로 (2026-11-10: 러닝 WBE3L97504 로퍼가 복종 표에서 빠지던 것)
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


class ColorPsi:
    """칼라 단위 격자(시점재고 컬러생판재) → 품번별 합계. Psi 와 같은 머리글 규칙, 숫자 열만 더한다. Psi 처럼 .rows / .n() 로 읽는다."""
    def __init__(self, grid, label):
        self.label = label
        h1 = [t(v) for v in grid[0]]; h2 = [t(v) for v in grid[1]]
        n = max(len(h1), len(h2)); h1 += [''] * (n - len(h1)); h2 += [''] * (n - len(h2))
        grp, g = [], ''
        for a in h1:
            if a: g = a
            grp.append(g)
        self.idx = {}
        for key, spec in Psi.COLS.items():
            if len(spec) == 1:
                j = next((i for i in range(n) if h2[i] == spec[0] or (not h2[i] and h1[i] == spec[0])), None)
            else:
                j = next((i for i in range(n) if grp[i] == spec[0] and h2[i] == spec[1]), None)
            if j is not None: self.idx[key] = j
        si = self.idx['style']
        rows = collections.defaultdict(lambda: collections.defaultdict(float))
        for r in grid[2:]:
            code = t(r[si]).upper() if len(r) > si else ''
            if not re.match(r'^[A-Z0-9]{8,}$', code): continue
            for key, j in self.idx.items():
                if key in ('year', 'season', 'sex', 'it', 'bok', 'item', 'style', 'name'): continue
                if j < len(r): rows[code][key] += num(r[j])
        self.rows = {k: dict(v) for k, v in rows.items()}

    def n(self, code, key): return self.rows[code].get(key, 0.0) if code in self.rows else 0.0


def read_tab_by_name(S, gsheet, name_part, norm):
    title, tabs = S.meta()
    name = next((k for k in tabs if name_part.replace(' ', '') in k.replace(' ', '')), None)
    if not name: return None, None
    q = "'" + name.replace("'", "''") + "'"
    v = S._req('GET', '/values/' + gsheet._enc(q) + '?valueRenderOption=UNFORMATTED_VALUE').get('values', [])
    return name, norm(v)


BASIS = 'all'      # 'all' = 기획 작지·생산 전체(리오더 포함, 26SS 종합 시트와 같음) / 'plan0' = 최초기획·기획생산(리오더 제외, 26FW 카드·집계표 규칙)


AMT_ADJ = {}      # {품번: (누계 실판매금액에서 뺄 값, 주간 판매금액에서 뺄 값)} — 26FW 집계표의 실판매금액 보정(AMT_FIX)과 같은 값을 빌드가 넣어 준다
CUT731 = None     # (현재 Psi, 7/31 ColorPsi, {품번}) — 이 품번들은 metrics 에서 생산·판매·재고를 7/31 시점 값으로 센다(8/1~ 는 26FW 러닝 이월로 가므로). 26SS 신상이면서 FW 과거품번인 GF1 10개


def metrics(psi, codes, co_sku=None, basis=None, co_in=None):
    """품번 묶음의 시즌 지표 합(ss_data 의 cur 모양). basis 에 따라 기획·입고를 리오더 포함/제외로 센다.
    co_in({품번: 주간판매(CO) 입고 합})을 주면 입고 수량만 그 값으로 센다 — 26FW 집계표가 CO 입고를 쓰므로 연간의 26FW 몫도 같은 화면 값으로 맞춘다
    (ERP 두 화면의 입고 수량이 다른 품번이 있다: 2026-11-09 GF7LKC507·GF4LJP505·GF3UKP008). 금액은 두 쪽 모두 생판재(ST) 값."""
    basis = basis or BASIS
    kp, kpc, ki, kt, kc = ('planAll', 'planAllCost', 'prodAll', 'prodAllTag', 'prodAllCost') if basis == 'all' else ('plan', 'planCost', 'qtyIn', 'tag', 'cost')
    a = {k: 0.0 for k in KEYS}
    for c in codes:
        if c not in psi.rows: continue
        cut = bool(CUT731) and psi is CUT731[0] and c in CUT731[2]
        src = CUT731[1] if cut else psi                              # 7/31 로 자르는 품번은 생산·판매·재고를 7/31 시점 값으로
        q_in = co_in[c] if (co_in is not None and c in co_in and not cut) else src.n(c, ki)
        a['plan'] += psi.n(c, kp); a['planCost'] += psi.n(c, kpc); a['qtyIn'] += q_in; a['tag'] += src.n(c, kt); a['cost'] += src.n(c, kc)
        for k in ('sales', 'salesTag', 'salesAmt', 'salesCost', 'stock'):
            a[k] += src.n(c, k)
        if not cut:
            for k in ('wkQty', 'wkAmt'): a[k] += psi.n(c, k)
        if c in AMT_ADJ and not cut:
            a['salesAmt'] -= AMT_ADJ[c][0]; a['wkAmt'] -= AMT_ADJ[c][1]
        if psi.n(c, kp) > 0: a['styles'] += 1
        if q_in > 0: a['inM'] += 1
        if co_sku:
            a['sku'] += co_sku.get(c, (0, 0))[0]; a['inSku'] += co_sku.get(c, (0, 0))[1]
    return derive(a)


def run_metrics(ytd, dec, cur, codes, start=None, end=None):
    """러닝 구품번. 기본은 1/1~ 기준(start=12/31 시점, end=현재). start/end 로 구간을 바꾼다(2026-11-10):
      · 26SS 몫 = start 12/31, end 7/31 시점  → 투입 = 12/31 재고 + 1/1~7/31 생산 · 판매 = 7/31 누계 − 12/31 누계 · 재고 = 7/31 재고
      · 26FW 몫 = start 7/31 시점, end 현재    → 투입 = 8/1~ 생산만(7/31 이월 재고는 SS 투입에 이미 있음 · carryQty/carryCost 로 따로) · 판매 = 현재 − 7/31 · 재고 = 현재고
    1/1~ 기준. 입고(투입) = 12/31 재고 + 2026 생산(= 현재 누적 생산 − 12/31 누적 생산) · 판매 = 1/1~ (= 현재 누계 − 12/31 누계) · 재고 = 현재고. 기획은 없다.
    ERP 의 '기간' 조건은 판매(기간)에만 걸리고 생산·재고는 누적값 그대로라(2026-11-09 확인) 두 시점의 차로 올해 값을 만든다. 1/1~ 탭의 판매(기간)은 대조용."""
    a = {k: 0.0 for k in KEYS}; a['plan'] = None; a['planCost'] = None
    base = start if start is not None else dec; top = end if end is not None else cur
    opening = base is dec                       # 12/31 에서 시작하면 기초재고를 투입에 넣는다
    carry_q = carry_c = 0.0
    resid, ytd_diff = [], []
    for c in codes:
        if c not in cur.rows: continue
        d = lambda k: top.n(c, k) - base.n(c, k)
        d_stk, d_tag, d_cost = (base.n(c, 'stock'), base.n(c, 'stockTag'), base.n(c, 'stockCost')) if opening else (0.0, 0.0, 0.0)
        a['qtyIn'] += d_stk + d('prodAll'); a['tag'] += d_tag + d('prodAllTag'); a['cost'] += d_cost + d('prodAllCost')
        for k in ('sales', 'salesTag', 'salesAmt', 'salesCost'): a[k] += d(k)
        if top is cur:
            for k in ('wkQty', 'wkAmt'): a[k] += cur.n(c, k)
        a['stock'] += top.n(c, 'stock')
        if not opening: carry_q += base.n(c, 'stock'); carry_c += base.n(c, 'stockCost')
        if d_stk + d('prodAll') > 0: a['inM'] += 1
        a['styles'] += 1
        if opening and top is cur:
            resid.append((round(d_stk + d('prodAll') - d('sales') - d('etc') - cur.n(c, 'stock')), c))
            if ytd is not None and c in ytd.rows: ytd_diff.append((round(ytd.n(c, 'sales') - d('sales')), c))
    out = derive(a); out['carryQty'] = carry_q; out['carryCost'] = carry_c
    return out, resid, ytd_diff


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
    nm_7 = gsource.TABS['stk26'][1]
    s731 = ColorPsi(g[nm_7], nm_7)                                   # 7/31 시점(칼라 행 → 품번 합) — 구품번을 26SS 몫(1/1~7/31) / 26FW 몫(8/1~) 으로 나누는 경계
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
    # (가) 2026-11-10 사용자: 26FW 러닝 이월은 러닝아이템 탭 '26FW - 과거품번' 열의 37개(대시보드 26FW 집계표의 '시점재고 합산'과 같은 목록)만.
    #   이 37개는 1/1~7/31 → 26SS 러닝, 8/1~ → 26FW 러닝(신품번과 짝). 그중 ERP 2026 봄 신상인 GF1 10개도 구품번(이월)으로 세고 26SS 신상에서 뺀다(두 번 세지 않으려고).
    #   목록 밖의 과거 시즌 러닝(32개)은 나누지 않고 1/1~ 전체를 26SS 러닝에 둔다 → 26SS + 26FW = 연간 그대로.
    fw_old_list = [c for c in season_codes.get('26FW - 과거품번', []) if c in cur.rows]
    run_old = [c for c in old_union if c not in new2026]
    old_as_new = [c for c in old_union if c in new2026]                                   # 러닝 목록에 있지만 2026 신상 → 신상(GF1 10개는 26SS 신상, 단 FW 과거품번이라 8/1~ 는 26FW 러닝 이월)
    gf1_cut = [c for c in fw_old_list if c in new2026 and c in ss_new]                    # 26SS 신상이면서 FW 과거품번: 26SS 신상은 7/31 까지 몫만
    global CUT731
    CUT731 = (cur, s731, set(gf1_cut))
    missing = [c for c in run_old if c not in cur.rows]
    run_old = [c for c in run_old if c in cur.rows]
    ss_run = [c for c in season_codes.get('26SS', []) if c in run_old]
    old_add = [c for c in run_old if c not in ss_run]
    split_codes = [c for c in run_old if c in fw_old_list]                                # 7/31 로 나누는 구품번(27 · FW 과거품번 중 2026 신상이 아닌 것)
    whole_codes = [c for c in run_old if c not in fw_old_list]                            # 1/1~ 전체가 26SS 러닝(32)
    fw_carry_codes = split_codes + gf1_cut                                                 # 26FW 러닝 이월 = FW 과거품번 37개 전부(대시보드 시점재고 합산과 같은 목록)

    def old_part(codes, part):
        """구품번 묶음의 26SS 몫('ss') / 26FW 몫('fw'). ss = 나누는 품번은 1/1~7/31 + 안 나누는 품번은 1/1~ 전체(GF1 은 제외 — 26SS 신상) ·
        fw = FW 과거품번(GF1 포함)의 8/1~ 만."""
        sp = [c for c in codes if c in fw_old_list and c not in gf1_cut]; wh = [c for c in codes if c not in fw_old_list and c not in gf1_cut]
        if part == 'fw':
            m, _, _ = run_metrics(ytd, dec, cur, [c for c in codes if c in fw_old_list], start=s731); return m
        a, _, _ = run_metrics(ytd, dec, cur, sp, end=s731); b, _, _ = run_metrics(ytd, dec, cur, wh)
        raw = {k: ((a.get(k) or 0) + (b.get(k) or 0)) if k not in ('plan', 'planCost') else None for k in KEYS}
        m = derive(raw); m['carryQty'] = 0.0; m['carryCost'] = 0.0
        return m
    # 유일성 검사
    buckets = {'26SS 신상': ss_new, '26FW 신상': fw_new, '러닝 신품번': run_new_erp, '러닝 구품번': run_old}
    seen = collections.Counter(c for v in buckets.values() for c in v)
    dup = sorted(c for c, n in seen.items() if n > 1)
    # --- SKU(컬러) 수: 주간판매(CO) ---
    co = g[gsource.TABS['co'][1]]
    h2 = [t(v) for v in co[1]]
    ci_s, ci_p, ci_in = h2.index('스타일'), h2.index('작지' if BASIS == 'all' else '최초기획'), h2.index('입고' if BASIS == 'all' else '기획입고')
    ci_all = h2.index('입고'); ci_pi = h2.index('기획입고')
    co_sku, co_in, co_in0 = {}, {}, {}                                 # co_in = CO 입고(리오더 포함) · co_in0 = CO 기획입고(리오더 제외, 26FW 집계표 본 숫자)
    for r in co[2:]:
        c = t(r[ci_s]).upper() if len(r) > ci_s else ''
        if not re.match(r'^[A-Z0-9]{8,}$', c): continue
        co_in[c] = co_in.get(c, 0.0) + num(r[ci_all]); co_in0[c] = co_in0.get(c, 0.0) + num(r[ci_pi])
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

    def new_block(codes, seasons, py_codes_, with_run, co_in=None):
        """신상 묶음(codes, 시즌 목록) → block1/newCat/seasonCat/genderItem/genderCat (ss_data 모양). with_run 이면 러닝(구품번 전체, 1/1~) 행도 붙인다.
        co_in 을 주면(26FW) 입고 수량을 주간판매(CO) 입고로 센다."""
        _metrics = globals()['metrics']
        def metrics(psi, cs, sku=None):
            return _metrics(psi, cs, sku, co_in=co_in if psi is cur else None)
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
            run_by = rows_for(ss_tab_run, cur, it)
            for item in ITEMS:
                m = old_part(run_by.get(item, []), 'ss')
                b1.append({'season': '', 'item': item, 'mode': 'run', 'kind': '항목', 'cur': m, 'py': None})
            m_run_ = old_part(ss_tab_run, 'ss')
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

    # 26SS 탭의 러닝 = 러닝 구품번 전체(26SS 러닝 목록 + 그 밖의 구품번)의 1/1~7/31 몫 (2026-11-10 사용자 정리안).
    # 8/1~ 몫은 26FW 러닝(신품번과 짝 연결)으로 간다 → 26SS 러닝 + 26FW 구품번 8/1~ = 구품번 1/1~ 전체, 연간에 겹침 없음.
    ss_tab_run = ss_run                      # 26SS 탭·연간의 26SS 러닝 = 26SS 러닝 목록 31개만(대시보드 26SS 종합과 같음)
    extra_run = old_add                      # 과년차 러닝 상품: 러닝아이템 탭 과거 시즌 열에만 있는 구품번 28개 — 26SS/26FW 어느 시즌에도 안 넣고 연간에 별도 칸(2026-11-10 2번안)
    all_old = ss_run + old_add
    miss731 = [c for c in fw_carry_codes if c not in s731.rows]           # 7/31 탭에 없는 구품번(나누는 37개 중): 7/31 = 12/31 로 보아 전부 FW 몫이 된다
    resid_all, ytd_all = [], []
    block1, new_cat, season_cat, gender_item, gender_cat, ttl_new, m_run = new_block(ss_new, SS_SEASONS, py_codes, True)
    # 26FW 신상(같은 기준) — 연간 탭에서 26FW 집계표 대신 쓴다(리오더 포함/제외 기준을 SS 와 맞추려고)
    fw_codes = [c for c in fw_new if it(c) and t(cur.g(c, 'item')) != '사은품' and c not in FW_GIFT]
    fw_gift = [c for c in fw_new if c not in fw_codes]
    py_fw = [c for c in py.rows if py_year(c) == '2025' and py_se(c) in FW_SEASONS and py_it(c) and not ('U98' in c and not py.n(c, 'salesAmt'))]
    fb1, fnc, fsc, fgi, fgc, fw_ttl, _ = new_block(fw_codes, FW_SEASONS, py_fw, False, co_in=co_in)
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
        oc = [c for c in ss_tab_run if f(c)]; fc = [c for c in fw_carry_codes if f(c)]; xc = [c for c in extra_run if f(c)]
        a = old_part(oc, 'ss')                                                  # 26SS 몫(31개): FW 과거품번은 1/1~7/31, 그 밖은 1/1~ 전체
        b = old_part(fc, 'fw')                                                  # 26FW 몫: FW 과거품번 37개(GF1 포함)의 8/1~ (투입 = 8/1~ 생산 · 이월은 carryQty/carryCost)
        x = old_part(xc, 'ss')                                                  # 과년차 러닝(28개): FW 과거품번 17개는 1/1~7/31, 그 밖 11개는 1/1~ 전체
        full, rb, yb = run_metrics(ytd, dec, cur, oc + xc)                      # 검산용 1/1~ 전체(구품번 59개)
        if item != 'TOTAL': resid_all += rb; ytd_all += yb
        n_codes = [c for c in run_new_erp if f(c)]
        empty = derive({k: 0.0 for k in KEYS})
        run_items[item] = {'ssRun': a, 'fwOld': b, 'extraRun': x, 'oldAdd': empty, 'full': full, 'fwNew': metrics(cur, n_codes, co_sku, co_in=co_in), 'nExtra': len(xc),
                           'nSs': len(oc), 'nAdd': 0, 'nNew': len(n_codes), 'nSsList': len([c for c in ss_run if f(c)]), 'nOldAdd': len([c for c in old_add if f(c)]),
                           'nSplit': len([c for c in oc if c in fw_old_list]), 'nWhole': len([c for c in oc if c not in fw_old_list]), 'nGf1': len([c for c in gf1_cut if f(c)])}
    resid_max = max((abs(r[0]) for r in resid_all), default=0)
    ytd_max = max((abs(r[0]) for r in ytd_all), default=0)
    # 옛 방식(카드 단순 합산)이라면 두 번 세었을 품번 — 설명용
    legacy = {'ssRunAndFwOld': sorted(set(season_codes.get('26SS', [])) & set(season_codes.get('26FW - 과거품번', []))),
              'ssNewAsFwOld': sorted(set(season_codes.get('26FW - 과거품번', [])) & set(ss_new)),
              'seasonMoved': sorted(c for c in old_as_new if c in fw_new) + sorted(c for c in fw_new if c[:3] in ('GF1', 'GF2'))}
    legacy['seasonMoved'] = sorted(set(legacy['seasonMoved']))
    # 신품번 ↔ 구품번 짝(러닝아이템 탭 26FW - 과거품번 → 26FW) — 26FW 러닝을 한 줄에 합산해 보이기 위한 상세
    i_old = hdr.index('26FW - 과거품번') if '26FW - 과거품번' in hdr else None
    new_of = collections.defaultdict(list)
    if i_new is not None and i_old is not None:
        for r in run_raw[2:]:
            if len(r) > max(i_new, i_old) and t(r[i_old]) and t(r[i_new]):
                o, n_ = t(r[i_old]).upper(), t(r[i_new]).upper()
                if (o in all_old or o in gf1_cut) and n_ in run_new_erp and n_ not in new_of[o]: new_of[o].append(n_)
    def new_info(n_):
        return {'code': n_, 'name': t(cur.g(n_, 'name')), 'item': it(n_) or '', 'sales': round(cur.n(n_, 'sales')), 'qtyIn': round(co_in.get(n_, cur.n(n_, 'prodAll'))),
                'cost': round(cur.n(n_, 'prodAllCost')), 'stock': round(cur.n(n_, 'stock')), 'wkQty': round(cur.n(n_, 'wkQty'))}
    def pair_row(c):
        a = old_part([c], 'ss'); b = old_part([c], 'fw')
        return {'code': c, 'name': t(cur.g(c, 'name')), 'item': it(c) or '', 'has731': c in s731.rows, 'inSsList': c in ss_run, 'split': c in fw_old_list, 'wasNew': c in new2026,
                'ssSales': round(a['sales']), 'ssIn': round(a['qtyIn']), 'ssCost': round(a['cost']), 'stock731': round(a['stock']),
                'fwSales': round(b['sales']), 'fwIn': round(b['qtyIn']), 'fwCost': round(b['cost']), 'stock': round(cur.n(c, 'stock')), 'wkQty': round(cur.n(c, 'wkQty')),
                'news': [new_info(n_) for n_ in new_of.get(c, [])]}
    pair_rows = [pair_row(c) for c in all_old + gf1_cut]
    paired_new = set(x for v in new_of.values() for x in v)
    unpaired_new = [new_info(n_) for n_ in run_new_erp if n_ not in paired_new]
    # 성별 러닝(간소화 페이지용): 여성 / 남성(유니맨) × ssRun(1/1~7/31) · fwOld(8/1~) · fwNew
    sexg = lambda c: '여성' if t(cur.g(c, 'sex')) == '여성' else '남성'
    # 26FW 리오더 제외 값(대시보드 26FW 집계표 본 숫자와 같은 기준: 최초기획·기획입고·기획생산 금액) — 정리안 페이지가 '본 숫자 + 리오더 행' 으로 나눠 보이는 데 쓴다(2026-11-10)
    fwp0 = lambda cs: metrics(cur, cs, co_sku, basis='plan0', co_in=co_in0)
    fw_plan0 = {'TOTAL': {'new': fwp0(fw_codes), 'runNew': fwp0(run_new_erp)}}
    for item in ITEMS:
        fw_plan0[item] = {'new': fwp0([c for c in fw_codes if it(c) == item]), 'runNew': fwp0([c for c in run_new_erp if it(c) == item])}
    sexg0 = lambda c: '여성' if t(cur.g(c, 'sex')) == '여성' else '남성'
    for gnd in ('여성', '남성'):
        fw_plan0[gnd] = {'new': fwp0([c for c in fw_codes if sexg0(c) == gnd]), 'runNew': fwp0([c for c in run_new_erp if sexg0(c) == gnd])}
    run_gender = {}
    for gnd in ('여성', '남성'):
        oc = [c for c in ss_tab_run if sexg(c) == gnd]; nc = [c for c in run_new_erp if sexg(c) == gnd]; xc = [c for c in extra_run if sexg(c) == gnd]
        a = old_part(oc, 'ss'); b = old_part([c for c in fw_carry_codes if sexg(c) == gnd], 'fw'); x = old_part(xc, 'ss')
        run_gender[gnd] = {'ssRun': a, 'fwOld': b, 'extraRun': x, 'fwNew': metrics(cur, nc, co_sku, co_in=co_in), 'nSs': len(oc), 'nNew': len(nc), 'nExtra': len(xc)}
    T_ = run_items['TOTAL']
    _fw_old_only = old_part(all_old, 'fw')                                               # 검산: 구품번 59개(GF1 제외) 1/1~ 전체 = 26SS 러닝 + 과년차 + FW 8/1~ 몫
    split_chk = {k: round((T_['full'].get(k) or 0) - ((T_['ssRun'].get(k) or 0) + (T_['extraRun'].get(k) or 0) + (_fw_old_only.get(k) or 0))) for k in ('sales', 'qtyIn', 'cost')}
    ann = {'split': {'tab731': nm_7, 'miss731': miss731, 'check': split_chk, 'nPairs': sum(1 for c in ss_tab_run if new_of.get(c)), 'nUnpairedNew': len(unpaired_new),
                     'nSplit': len(split_codes), 'nWhole': len(whole_codes), 'fwOldList': fw_carry_codes, 'gf1Cut': sorted(gf1_cut), 'nFwCarry': len(fw_carry_codes),
                     'nExtra': len(extra_run), 'extraSplit': len([c for c in extra_run if c in fw_old_list]), 'extraWhole': len([c for c in extra_run if c not in fw_old_list]),
                     'ssSplit': len([c for c in ss_run if c in fw_old_list]), 'ssWhole': len([c for c in ss_run if c not in fw_old_list]), 'extraList': extra_run},
           'pairRows': pair_rows, 'unpairedNew': unpaired_new,
           'runGender': run_gender, 'fwPlan0': fw_plan0,
           'ref': ref.isoformat(), 'ytdTab': nm_y, 'decTab': nm_d, 'gen': datetime.datetime.now().strftime('%Y-%m-%d %H:%M'),
           'buckets': {k: len(v) for k, v in buckets.items()}, 'bucketCodes': {'러닝 구품번': run_old, '러닝 신품번': run_new_erp},
           'ssGift': [(c, t(cur.g(c, 'name'))) for c in ss_gift], 'gifts25': gifts25, 'oldAsNew': old_as_new, 'missing': missing, 'dup': dup,
           'nSsRun': len(ss_tab_run), 'nOldAdd': 0, 'nSsList': len(ss_run), 'nOldOutside': len(old_add), 'nFwOld': len(fw_carry_codes), 'nRunNew': len(run_new_erp), 'runNewList': run_new,
           'items': run_items, 'oldRows': [detail(c, '26SS 러닝' if c in ss_run else '구품번 추가') for c in run_old],
           'basis': BASIS, 'fw': fw_block,
           'residMax': resid_max, 'residTop': sorted(resid_all, key=lambda x: -abs(x[0]))[:5], 'ytdSalesDiffMax': ytd_max, 'legacy': legacy}
    ss = {'file': '구글 시트 ERP 자동 조회 (%s 기준)' % ref, 'end': ref.isoformat(), 'prevEnd': (ref.replace(year=ref.year - 1)).isoformat(), 'source': 'gsheet',
          'gifts25': ['%s %s' % (c, t(py.g(c, 'name'))) for c in gifts25],
          'gifts25Sum': {'n': len(gifts25), 'qtyIn': round(sum(py.n(c, 'qtyIn') for c in gifts25)), 'costV': round(sum(py.n(c, 'cost') for c in gifts25) * 1.1)},
          'gifts26': [{'code': c, 'name': t(cur.g(c, 'name')), 'season': se(c), 'qtyIn': round(cur.n(c, 'qtyIn')), 'costV': round(cur.n(c, 'cost') * 1.1), 'rows': 0} for c in ss_gift],
          'block1': block1, 'newCat': new_cat, 'seasonCat': season_cat, 'genderItem': gender_item, 'genderCat': gender_cat,
          'runList': ss_tab_run, 'run': {'ssList': len(ss_tab_run), 'fwList': len(season_codes.get('26FW - 과거품번', [])), 'overlap': 0, 'overlapStyles': [], 'overlapQty': 0, 'overlapQtyLo': 0, 'since': ''}, 'check': []}
    if verbose:
        b = ann['buckets']
        print('연간(구글 시트 단일 원천, 기준 %s): 26SS 신상 %d · 26FW 신상 %d · 러닝 신품번 %d · 러닝 구품번 %d(SS 러닝 %d + 추가 %d) · 겹침 %s · 러닝목록에 있지만 2026 신상인 품번 %d · ST 에 없는 구품번 %s'
              % (ref, b['26SS 신상'], b['26FW 신상'], b['러닝 신품번'], b['러닝 구품번'], len(ss_run), len(old_add), dup or '없음', len(old_as_new), missing or '없음'))
        print('  1/1~ 탭 %s %d품번 · 12/31 탭 %s %d품번 · 재고 항등식 잔차 최대 %d장(%s) · 1/1~ 탭 판매 vs (현재−12/31) 최대 차 %d장 · 25SS 사은품 제외 %d · 26SS 기타 제외 %d'
              % (nm_y, len(ytd.rows), nm_d, len(dec.rows), resid_max, ann['residTop'][:3], ytd_max, len(gifts25), len(ss_gift)))
        tn = ttl_new; tr = m_run
        print('  26FW 신상(같은 기준, 사은품 %d 제외): 입고 %s · 누계 %s · 주간 %s · 원가(VAT+) %s' % (len(fw_gift), format(round(fw_ttl['qtyIn']), ','), format(round(fw_ttl['sales']), ','), format(round(fw_ttl['wkQty']), ','), format(fw_ttl['costV'] or 0, ',')))
        print('  26SS 신상: 입고 %s · 누계 %s · 주간 %s · 원가(VAT+) %s | 러닝(26SS 목록 %d · FW 과거품번 %d 의 1/1~7/31 + 그 밖 %d 의 1/1~): 투입 %s · 판매 %s · 재고 %s | FW 과거품번 37개(GF1 10 포함) 8/1~(26FW 몫): 투입 %s · 판매 %s | 과년차 러닝 %d: 투입 %s · 판매 %s | 7/31 탭에 없는 구품번 %s · 나누기 검산 %s'
              % (format(round(tn['qtyIn']), ','), format(round(tn['sales']), ','), format(round(tn['wkQty']), ','), format(tn['costV'] or 0, ','),
                 len(ss_tab_run), len([c for c in ss_run if c in fw_old_list]), len([c for c in ss_run if c not in fw_old_list]), format(round(tr['qtyIn']), ','), format(round(tr['sales']), ','), format(round(tr['stock']), ','),
                 format(round(run_items['TOTAL']['fwOld']['qtyIn']), ','), format(round(run_items['TOTAL']['fwOld']['sales']), ','),
                 len(extra_run), format(round(run_items['TOTAL']['extraRun']['qtyIn']), ','), format(round(run_items['TOTAL']['extraRun']['sales']), ','), miss731 or '없음', split_chk))
    return ss, ann


def fw_split_reorder(GS, SUM, verbose=True):
    """26FW 집계표(SUM)의 금액을 수량과 같은 기준(리오더 제외)으로 맞춘다 — 2026-11-09.
    종합 시트 수식은 수량(기획량·입고량)은 최초기획·기획입고(리오더 제외)인데, TAG 금액은 생판재 '생산' 금액(리오더 포함),
    러닝의 생산원가도 '생산' 금액(리오더 포함)이라 한 줄에 두 기준이 섞여 있었다(장당 원가·마크업이 어긋나고, 리오더 행을 더하면 두 번 셈).
    여기서 SUM 의 각 행에서 그 묶음 품번의 리오더 금액(= 생판재 생산 − 기획생산)을 뺀다: TAG 는 모든 행, 원가는 러닝 행(+전체 TTL 의 러닝 몫).
    리오더 행의 원가·TAG 도 같은 생판재 차액으로 바꿔 '본 행 + 리오더 행 = 연간의 26FW 몫' 이 금액에서도 성립하게 한다.
    전년(py)·지난주(was) 값은 건드리지 않는다(지난주 스냅샷은 옛 기준이라 이번 주 한 번은 TAG·러닝 원가의 전주 대비가 리오더만큼 작게 나온다).
    돌려주는 것: 점검용 dict."""
    import gsource
    cur = Psi(GS['grids'][gsource.TABS['st'][1]], '생판재(ST)')
    codes = [c for c in cur.rows if t(cur.g(c, 'year')) == '2026']
    se = lambda c: t(cur.g(c, 'season')); it = lambda c: GRP_OF.get(t(cur.g(c, 'it'))); cat = lambda c: t(cur.g(c, 'item'))
    sex = lambda c: '여성' if t(cur.g(c, 'sex')) == '여성' else '남성'
    new = [c for c in codes if se(c) in FW_SEASONS]; run = [c for c in codes if '러닝' in se(c)]
    d_tag = {c: cur.n(c, 'prodAllTag') - cur.n(c, 'tag') for c in new + run}
    d_cost = {c: cur.n(c, 'prodAllCost') - cur.n(c, 'cost') for c in new + run}
    itm = lambda x: 'ACC' if x in ('ACC', 'ACCESSORY') else x
    whole = lambda x: x in ('TTL', 'TOTAL', '', None)
    n_fix = [0]

    def fix(row, cs, cost_cs):
        r = row.get('cur') if isinstance(row.get('cur'), dict) else row
        dt, dc = sum(d_tag[c] for c in cs), sum(d_cost[c] for c in cost_cs)
        if not dt and not dc: return
        n_fix[0] += 1
        if r.get('tag') is not None: r['tag'] = r['tag'] - dt
        if r.get('cost') is not None and dc: r['cost'] = r['cost'] - dc
        if r.get('cost'):
            r['costV'] = round(r['cost'] * 1.1); r['mu'] = (r.get('tag') or 0) / r['cost'] / 1.1
        if r.get('planCost'): r['muPlan'] = (r.get('tag') or 0) / r['planCost'] / 1.1

    for d in SUM.get('block1', []):
        if d.get('kind') == '전체 TTL': fix(d, new + run, run); continue
        base = new if d.get('mode') == 'new' else run
        cs = [c for c in base if (not d.get('season') or se(c) == d['season']) and (whole(d.get('item')) or it(c) == itm(d['item']))]
        fix(d, cs, cs if d.get('mode') != 'new' else [])
    for nm, base, is_run in (('newItem', new, False), ('runItem', run, True)):
        for d in SUM.get(nm, []):
            cs = [c for c in base if whole(d.get('item')) or it(c) == itm(d['item'])]; fix(d, cs, cs if is_run else [])
    for nm, base, is_run in (('newCat', new, False), ('runCat', run, True)):
        for d in SUM.get(nm, []):
            cs = [c for c in base if it(c) == itm(d['item']) and cat(c) == d['cat']]; fix(d, cs, cs if is_run else [])
    for d in SUM.get('seasonCat', []):
        fix(d, [c for c in new if se(c) == d['season'] and it(c) == itm(d['item']) and cat(c) == d['cat']], [])
    for d in SUM.get('genderItem', []):
        g = '여성' if '여' in t(d.get('gender')) else '남성'
        fix(d, [c for c in new if sex(c) == g and (whole(d.get('item')) or it(c) == itm(d['item']))], [])
    for d in SUM.get('genderCat', []):
        g = '여성' if '여' in t(d.get('gender')) else '남성'
        fix(d, [c for c in new if sex(c) == g and it(c) == itm(d['item']) and cat(c) == d['cat']], [])
    # 리오더 행: 금액을 생판재 차액으로(수량은 CO 리오더기획·리오더입고 그대로)
    grp = {'가을': [c for c in new if se(c) == '가을'], '겨울': [c for c in new if se(c) == '겨울'], '러닝': run, 'TTL': new + run}
    chg = []
    for r in SUM.get('reorder', []):
        cs = grp.get(r.get('name'))
        if cs is None: continue
        nt, nc = sum(d_tag[c] for c in cs), sum(d_cost[c] for c in cs)
        if round(nt - (r.get('tag') or 0)) or round(nc - (r.get('cost') or 0)):
            chg.append('%s TAG %s→%s · 원가 %s→%s' % (r['name'], format(round(r.get('tag') or 0), ','), format(round(nt), ','), format(round(r.get('cost') or 0), ','), format(round(nc), ',')))
        r['tag'], r['cost'] = nt, nc
        r['costV'] = round(nc * 1.1) if nc else None
        r['mu'] = round(nt / nc / 1.1, 2) if nt and nc else None
    st_re = sorted(c for c in new + run if round(d_tag[c]) or round(d_cost[c]))
    co_re = sorted(set(c for r in SUM.get('reorder', []) if r.get('name') != 'TTL' for c in (r.get('codes') or [])))
    out = {'rows': n_fix[0], 'stOnly': sorted(set(st_re) - set(co_re)), 'coOnly': sorted(set(co_re) - set(st_re)), 'changed': chg,
           'tag': round(sum(d_tag.values())), 'cost': round(sum(d_cost.values())), 'runCost': round(sum(d_cost[c] for c in run))}
    if verbose:
        print('26FW 금액 기준 통일(리오더 제외): %d행 조정 · 리오더 TAG %s · 원가 %s(러닝 %s) · 리오더 행 금액 변경: %s · 생판재에만 리오더 금액 %s · CO 에만 리오더 %s'
              % (out['rows'], format(out['tag'], ','), format(out['cost'], ','), format(out['runCost'], ','), chg or '없음', out['stOnly'] or '없음', out['coOnly'] or '없음'))
    return out
