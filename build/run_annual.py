# -*- coding: utf-8 -*-
"""연간(SS+FW) 러닝 — 26SS 러닝과 26FW 러닝(시점재고 합산)이 같은 구품번을 두 번 세는 것을 찾아 빼고, 두 파일이 서로 맞는지(오차 50장 이내) 확인한다.

왜 겹치나
  · 26SS 러닝: 12월말 기초재고 + 신규 물량의 1/1~ 누계(구품번 31개, 26SS 스타일맵 '런닝 BACK').
  · 26FW 러닝 '시점재고 합산': 26FW 신품번 + 구품번 36개의 7/31 시점재고(이월)와 8/1~ 판매(26FW 스타일맵 '종합' CE~CR 열).
  · FW 구품번 36개 = ① SS 러닝 품번과 겹치는 10개(W 코드) ② 26SS 신상품 품번인 10개(GF1 코드 — 26SS 신상품 집계에 1/1~ 판매가 이미 들어 있다) ③ 어느 쪽에도 없는 FW 전용 16개.
  · ①②는 SS 쪽에 이미 들어 있는 입고·판매를 FW 이월이 다시 세므로 더하면 이중 집계 → 연간 러닝 = SS 러닝 전체 + FW 신품번 + (FW 이월 − ①② 의 이월) = FW 전용 이월만 더한다.
  · ②는 신상품 합계(26SS)와 겹치는 것이라 러닝을 신상품에 더할 때 함께 뺀다.

확인(오차 TOL=50장)
  ① 두 파일의 주간판매(52주)가 같은지 — 겹치는 품번의 주별 차이 최대값
  ② FW 이월 판매(8/1~, 누계 − 7/31 시점 누계)가 SS 주간판매로 센 8/3 주~ 이후 합과 7/27 주~ 이후 합 사이에 드는지(주 단위 경계라 범위로 본다)
  ③ 두 파일의 현재고가 같은지
  ④ FW 이월 계산(시점재고 시트·러닝 ST 시트)이 종합 시트의 CJ~CR 값과 같은지 — 같아야 겹치는 품번의 몫을 정확히 뺄 수 있다
"""
import collections
import re

import openpyxl

TOL = 50
SEASON_ORDER = ('24SS', '24FW', '25SS', '25FW', '26SS', '26FW')         # 러닝 원본의 시즌 열(시간순)
GRP = {'OUTER': 'OUTER', 'KNIT': 'TOP', 'INNER (W/O KNIT)': 'TOP', 'BOTTOM': 'BOTTOM', 'ACC': 'ACC'}
ITEMS = ('OUTER', 'TOP', 'BOTTOM', 'ACC')
BASIS_KEYS = ('qtyIn', 'cost', 'tag', 'sales', 'salesTag', 'salesAmt', 'salesCost', 'wkQty', 'wkAmt')        # SS 기준으로 세는 구품번 추가분의 항목
CARRY = ('qtyIn', 'sales', 'salesTag', 'salesAmt', 'wkQty', 'wkAmt', 'cost', 'salesCost', 'tag')      # cost = 7/31 시점재고의 재고 원가금액 · salesCost = 8/1~ 판매의 원가(둘 다 VAT 미포함 원천값)


def num(v):
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else 0.0


def txt(v):
    return '' if v is None else str(v).strip()


def weekly(ws, si, stock_col, maxr=6000):
    """주간판매(CO) 시트 -> 스타일별 {w: 52주 합(W52 → W1 순), stock: 재고합}"""
    rows = list(ws.iter_rows(min_row=1, max_row=maxr, values_only=True))
    hdr = rows[1]
    wcols = [j for j, h in enumerate(hdr) if h and re.match(r'^W\d+\(', str(h))]
    out = collections.defaultdict(lambda: {'w': [0.0] * len(wcols), 'stock': 0.0})
    for r in rows[3:]:
        if len(r) > si and r[si]:
            a = out[txt(r[si])]
            for k, j in enumerate(wcols):
                a['w'][k] += num(r[j]) if len(r) > j else 0.0
            a['stock'] += num(r[stock_col]) if len(r) > stock_col else 0.0
    return [str(hdr[j]) for j in wcols], out


def build(fw_wb, ss_path, ss_list, old36, fw_stk_total=None, week_start_idx=40, nas_all=None, nas_seasons=None):                # noqa: C901
    """fw_wb: 열린 26FW 스타일맵, ss_path: 26SS 스타일맵 경로, ss_list: 26SS 러닝 품번, old36: 26FW 러닝 구품번 목록.
    week_start_idx: 7/27 주의 위치(W52=0 … W1=51 중 W12=40). fw_stk_total: 종합 시트 CJ~CR 러닝 합계(대조용)."""
    wb = openpyxl.load_workbook(ss_path, read_only=True, data_only=True)
    overlap = sorted(set(ss_list) & set(old36))
    # --- SS: 런닝 BACK(스타일 단위) ---
    ssb = {}
    for r in wb['런닝 BACK'].iter_rows(min_row=3, max_row=60, max_col=40, values_only=True):
        if len(r) > 10 and r[3] and txt(r[3]) in ss_list:
            ssb[txt(r[3])] = dict(item=txt(r[2]), base=num(r[7]), neworder=num(r[8]), qtyIn=num(r[9]), sales=num(r[10]), cost=num(r[12]), salesAmt=num(r[15]), salesCost=num(r[16]), stock=num(r[0]))
    _, ssw = weekly(wb['런닝주간판매(CO)'], 21, 41)

    def life(ws):                                                          # 스타일별 누계 판매(판매기간 수량, 58열)
        out = collections.defaultdict(float)
        for r in ws.iter_rows(min_row=4, max_row=6000, max_col=70, values_only=True):
            if len(r) > 58 and r[12] and (txt(r[12]) in ss_list or txt(r[12]) in old36):
                out[txt(r[12])] += num(r[58])
        return out
    life_now, life_12 = life(wb['런닝 주마감(ST-생,판,재)']), life(wb['12월기초재고(런닝)'])
    ss_rec = {c: life_now[c] - life_12[c] for c in ss_list}                # 판매기록 기준 12월말 이후 판매

    def lifefull(ws, codes):                                               # 스타일별 누계(생산·판매·재고의 수량·택가·원가) — 12월기초재고(런닝) · 런닝 주마감 시트 공통 배치
        out = collections.defaultdict(lambda: collections.defaultdict(float))
        for r in ws.iter_rows(min_row=4, max_row=6000, max_col=70, values_only=True):
            if len(r) > 66 and r[12] and txt(r[12]) in codes:
                a = out[txt(r[12])]
                a['name'] = txt(r[13])
                a['grp'] = txt(r[5])                                           # 품목(KNIT·INNER·OUTER·BOTTOM·ACC·SHOES)
                for k, j in dict(prod=41, prodTag=42, prodCost=43, sales=58, salesTag=59, salesAmt=60, salesCost=61, stock=64, stockTag=65, stockCost=66, wkQty=54, wkAmt=56).items():
                    a[k] += num(r[j])
        return out
    chn = collections.defaultdict(float)                                    # 채널표('전주(유통1.1~)' 1블록: 온라인·오프라인·글로벌·기타 수량) = 날짜 기준 1/1~ 판매 — 12월말 시트와 무관한 두 번째 출처
    try:
        for r in wb['전주(유통1.1~)'].iter_rows(min_row=3, max_row=3000, max_col=22, values_only=True):
            if len(r) > 18 and r[2]:
                chn[txt(r[2])] += sum(num(r[j]) for j in (9, 12, 15, 18))
    except KeyError:
        chn = collections.defaultdict(float)
    # 러닝 원본(NAS 런닝품번리스트 = production-hub/inventory/running-codes.json 'all') + 25년 러닝 목록(FW 러닝스타일 리스트 I열) 중
    # SS 러닝·FW 구품번 어느 쪽에도 없는 품번 — 연간 러닝에 들어 있지 않다는 사실과 크기(SS 기준 12월말 재고 + 2026 입고, 1/1~ 판매)를 알린다
    i25 = []
    for r in fw_wb['26FW 러닝스타일 리스트'].iter_rows(min_row=2, max_row=120, max_col=12, values_only=True):
        if r and len(r) > 8 and r[8] and txt(r[8]) not in i25:
            i25.append(txt(r[8]))
    cand = list(dict.fromkeys(list(nas_all or []) + i25))
    unlisted = [c for c in cand if c not in ss_list and c not in old36]
    lf_now_u, lf_12_u = lifefull(wb['런닝 주마감(ST-생,판,재)'], set(unlisted)), lifefull(wb['12월기초재고(런닝)'], set(unlisted))
    unlisted_rows = [{'code': c, 'inNas': c in (nas_all or []), 'in25': c in i25} for c in unlisted if c not in lf_now_u]      # SS 스타일맵에 자료가 없어 넣지 못한 품번(있으면 알린다)
    extra_codes = [c for c in unlisted if c in lf_now_u]                    # 자료가 있는 것 = 러닝 원본에만 있는 과년 러닝 → 구품번 추가(SS 기준)
    # --- FW: 주간판매-러닝(CO), 러닝 ST(누계), 시점재고(7/31) ---
    wl, fww = weekly(fw_wb['주간판매-러닝(CO)'], 13, 33)
    st, name = collections.defaultdict(lambda: collections.defaultdict(float)), {}
    for r in fw_wb['생판재-러닝(ST)'].iter_rows(min_row=4, max_row=4997, max_col=70, values_only=True):
        if len(r) > 14 and r[14]:
            c = txt(r[14])
            name[c] = txt(r[15])
            a = st[c]
            a['grp'] = GRP.get(txt(r[7]), '?')
            for k, j in dict(sales=60, salesTag=61, salesAmt=62, wkQty=56, wkAmt=58, stock=66, salesCost=63).items():
                a[k] += num(r[j])
    sn = collections.defaultdict(lambda: collections.defaultdict(float))
    for r in fw_wb['시점재고(~26.07.31, 컬러생판재)'].iter_rows(min_row=4, max_row=5800, max_col=75, values_only=True):
        if len(r) > 70 and r[16]:
            a = sn[txt(r[16])]
            for k, j in dict(sales=64, salesTag=65, salesAmt=66, stock=70, stockTag=71, stockCost=72, salesCost=67).items():
                a[k] += num(r[j])

    def carry(c):                                                # 종합 시트 CJ~CR 과 같은 정의: 시점재고 수량 · 8/1~ 판매(누계 − 7/31 누계) · 금액 · 최근 1주
        return {'qtyIn': sn[c]['stock'], 'sales': st[c]['sales'] - sn[c]['sales'], 'salesTag': st[c]['salesTag'] - sn[c]['salesTag'],
                'salesAmt': st[c]['salesAmt'] - sn[c]['salesAmt'], 'wkQty': st[c]['wkQty'], 'wkAmt': st[c]['wkAmt'],
                'cost': sn[c]['stockCost'], 'salesCost': st[c]['salesCost'] - sn[c]['salesCost'], 'tag': sn[c]['stockTag']}      # tag = 7/31 재고의 택가 금액(시점재고 시트 71열)

    # --- SS 신상품(26SS) 스타일: 생,판,재 + 주간판매 — FW 구품번 중 26SS 신상품 품번(GF1 코드)과 겹치는 것을 찾는다 ---
    ssn = {}
    for r in wb['생,판,재'].iter_rows(min_row=4, max_row=1200, max_col=86, values_only=True):
        if len(r) > 77 and r[23] and str(r[13]) == '2026' and txt(r[14]) in ('봄', '여름'):
            a = ssn.setdefault(txt(r[23]), dict(item=GRP.get(txt(r[16]), '?'), name=txt(r[24]), qtyIn=0.0, sales=0.0, stock=0.0))
            a['qtyIn'] += num(r[54])
            a['sales'] += num(r[71])
            a['stock'] += num(r[77])
    _, ssw_new = weekly(wb['주간판매'], 10, 30)
    overlap_new = sorted(set(old36) & set(ssn) - set(overlap))

    items = {g: {'carryAll': {k: 0.0 for k in CARRY}, 'carryOverlap': {k: 0.0 for k in CARRY}, 'carryOverlapNew': {k: 0.0 for k in CARRY}, 'onlyBasis': {k: 0.0 for k in BASIS_KEYS}} for g in ITEMS}
    for c in old36:
        g = st[c]['grp'] if c in st else '?'
        if g not in items:
            continue
        cv = carry(c)
        for k in CARRY:
            items[g]['carryAll'][k] += cv[k]
            if c in overlap:
                items[g]['carryOverlap'][k] += cv[k]
            elif c in overlap_new:
                items[g]['carryOverlapNew'][k] += cv[k]
    rows = []
    for c in overlap:
        cv, b = carry(c), ssb.get(c, {})
        sw, fw = ssw.get(c, {'w': [0.0] * 52, 'stock': 0.0}), fww.get(c, {'w': [0.0] * 52, 'stock': 0.0})
        i0 = week_start_idx
        s27, s03 = sum(sw['w'][i0:]), sum(sw['w'][i0 + 1:])
        f27, f03 = sum(fw['w'][i0:]), sum(fw['w'][i0 + 1:])
        wd = max(abs(x - y) for x, y in zip(sw['w'], fw['w']))
        inb = s03 - TOL <= cv['sales'] <= s27 + TOL
        rows.append({'code': c, 'name': name.get(c, ''), 'item': b.get('item') or st[c]['grp'], 'ssIn': b.get('qtyIn'), 'ssSales': b.get('sales'), 'ssStock': sw['stock'], 'fwStock': fw['stock'],
                     'fwCarryStock': cv['qtyIn'], 'fwAugSales': cv['sales'], 'ssWk0727': s27, 'ssWk0803': s03, 'fwWk0727': f27, 'fwWk0803': f03, 'weeklyMaxDiff': wd,
                     'boundaryOk': inb, 'stockDiff': sw['stock'] - fw['stock'], 'weeklyDiff0727': s27 - f27, 'ssSalesRec': ss_rec.get(c), 'ssSalesDiff': (b.get('sales') or 0) - ss_rec.get(c, 0.0)})
    only_codes = [c for c in old36 if c not in overlap and c not in overlap_new]
    lf_now_o, lf_12_o = lifefull(wb['런닝 주마감(ST-생,판,재)'], set(only_codes)), lifefull(wb['12월기초재고(런닝)'], set(only_codes))
    wk_i = next((j for j, lab in enumerate(wl) if '(12/29' in lab), 10)           # 'W42(12/29 -04)'
    rows_only = []                                                        # FW 전용 이월(SS 러닝·SS 신상품 어느 쪽에도 없는 구품번)
    for c in old36:
        if c in overlap or c in overlap_new:
            continue
        cv = carry(c)
        jj = (sn[c]['sales'] - life_12[c]) if c in life_12 else None                  # 1/1~7/31 판매 = 7/31 시점 누계 − 12월말 누계(SS 스타일맵 12월기초재고 시트에 있는 품번만)
        b12, bn = lf_12_o.get(c), lf_now_o.get(c)
        basis = None
        if b12 is not None and bn is not None:                                         # SS 기준(구품번 = 26SS): 12월말 재고 + 2026 입고, 1/1~ 판매·원가
            basis = {'qtyIn': b12['stock'] + bn['prod'] - b12['prod'], 'newIn': bn['prod'] - b12['prod'], 'base': b12['stock'],
                     'cost': b12['stockCost'] + bn['prodCost'] - b12['prodCost'], 'tag': b12['stockTag'] + bn['prodTag'] - b12['prodTag'],
                     'sales': bn['sales'] - b12['sales'], 'salesTag': bn['salesTag'] - b12['salesTag'], 'salesAmt': bn['salesAmt'] - b12['salesAmt'], 'salesCost': bn['salesCost'] - b12['salesCost'],
                     'wkQty': st[c]['wkQty'], 'wkAmt': st[c]['wkAmt'], 'stockNow': bn['stock']}
            basis['resid'] = basis['qtyIn'] - basis['sales'] - bn['stock']             # 12월말 재고 + 입고 − 판매 − 현재고 (0 이면 서로 맞음)
        fwk = fww.get(c, {'w': [0.0] * len(wl)})['w']
        wk_hi, wk_lo = sum(fwk[wk_i:]), sum(fwk[wk_i + 1:])                            # 12/29 주부터 / 1/5 주부터 최근 주까지 FW 주간판매 합 — 12월말 기준일이 주 중간이라 두 값의 범위로 본다
        rows_only.append({'src': 'FW 구품번', 'nasSeasons': [k for k in SEASON_ORDER if c in (nas_seasons or {}).get(k, [])], 'chSales': (chn.get(c) if c in chn else None), 'wkLo': wk_lo, 'wkHi': wk_hi, 'wkOk': (basis is not None and wk_lo - TOL <= basis['sales'] <= wk_hi + TOL), 'basis': basis, 'code': c, 'name': name.get(c, ''), 'item': st[c]['grp'] if c in st else '?', 'fwCarryStock': cv['qtyIn'], 'fwAugSales': cv['sales'], 'fwStock': st[c]['stock'] if c in st else 0.0, 'janJulSales': jj})
    GRP_SS = {'KNIT': 'TOP', 'INNER (W/O KNIT)': 'TOP', 'OUTER': 'OUTER', 'BOTTOM': 'BOTTOM', 'ACC': 'ACC', 'SHOES': 'ACC'}        # SS 스타일맵 품목 → 4복종(SS 러닝 31개와 같은 분류)
    rows_extra = []                                                       # 러닝 원본(NAS)·25년 러닝 목록에만 있는 과년 러닝 — SS 러닝·FW 구품번 어디에도 없어 구품번 추가로 넣는다(SS 기준)
    for c in extra_codes:
        bn, b12 = lf_now_u[c], lf_12_u.get(c)
        z = b12 if b12 is not None else collections.defaultdict(float)
        basis = {'qtyIn': z['stock'] + bn['prod'] - z['prod'], 'newIn': bn['prod'] - z['prod'], 'base': z['stock'],
                 'cost': z['stockCost'] + bn['prodCost'] - z['prodCost'], 'tag': z['stockTag'] + bn['prodTag'] - z['prodTag'],
                 'sales': bn['sales'] - z['sales'], 'salesTag': bn['salesTag'] - z['salesTag'], 'salesAmt': bn['salesAmt'] - z['salesAmt'], 'salesCost': bn['salesCost'] - z['salesCost'],
                 'wkQty': bn['wkQty'], 'wkAmt': bn['wkAmt'], 'stockNow': bn['stock']}
        basis['resid'] = basis['qtyIn'] - basis['sales'] - bn['stock']
        g_ = GRP_SS.get(bn['grp'], '?')
        fwk = fww.get(c, {'w': [0.0] * len(wl)})['w']
        wk_hi, wk_lo = sum(fwk[wk_i:]), sum(fwk[wk_i + 1:])
        rows_extra.append({'src': '과년 러닝', 'chSales': (chn.get(c) if c in chn else None), 'wkLo': wk_lo, 'wkHi': wk_hi, 'wkOk': wk_lo - TOL <= basis['sales'] <= wk_hi + TOL, 'basis': basis, 'code': c, 'name': bn['name'], 'item': g_,
                           'fwCarryStock': None, 'fwAugSales': None, 'fwStock': st[c]['stock'] if c in st else None, 'janJulSales': None, 'noBase': b12 is None,
                           'stockDiff': (bn['stock'] - st[c]['stock']) if c in st else None, 'grpFw': st[c]['grp'] if c in st else None, 'inNas': c in (nas_all or []), 'in25': c in i25,
                           'nasSeasons': [k for k in SEASON_ORDER if c in (nas_seasons or {}).get(k, [])]})
    rows_all = []                                                         # SS 러닝 전 품번의 두 파일 주간판매·재고 대조(겹침 여부와 무관)
    for c in ss_list:
        sw, fw = ssw.get(c, {'w': [0.0] * 52, 'stock': 0.0}), fww.get(c, {'w': [0.0] * 52, 'stock': 0.0})
        rows_all.append({'code': c, 'overlap': c in overlap, 'ssSalesBack': ssb.get(c, {}).get('sales'), 'ssSalesRec': ss_rec.get(c), 'ssWk0727': sum(sw['w'][week_start_idx:]), 'fwWk0727': sum(fw['w'][week_start_idx:]),
                         'weeklyMaxDiff': max(abs(x - y) for x, y in zip(sw['w'], fw['w'])), 'ssStock': sw['stock'], 'fwStock': fw['stock']})
    for r_ in rows_only + rows_extra:                                     # 구품번 추가(FW 구품번 16 + 러닝 원본 13)의 SS 기준 합계를 복종별로
        g_ = r_['item']
        if r_['basis'] and g_ in items:
            for k in BASIS_KEYS:
                items[g_]['onlyBasis'][k] += r_['basis'][k]
    rows_new = []
    for c in overlap_new:
        cv, b = carry(c), ssn[c]
        sw = ssw_new.get(c, {'w': [0.0] * 52, 'stock': 0.0})
        fw = fww.get(c)
        i0 = week_start_idx
        s27, s03 = sum(sw['w'][i0:]), sum(sw['w'][i0 + 1:])
        wd = max(abs(x - y) for x, y in zip(sw['w'], fw['w'])) if fw else None
        rows_new.append({'code': c, 'name': b['name'], 'item': b['item'], 'ssIn': b['qtyIn'], 'ssSales': b['sales'], 'ssStock': b['stock'], 'fwStock': st[c]['stock'], 'fwSalesNow': st[c]['sales'], 'salesDiff': b['sales'] - st[c]['sales'],
                         'fwCarryStock': cv['qtyIn'], 'fwAugSales': cv['sales'], 'ssWk0727': s27, 'ssWk0803': s03, 'weeklyMaxDiff': wd,
                         'boundaryOk': s03 - TOL <= cv['sales'] <= s27 + TOL, 'stockDiff': b['stock'] - st[c]['stock']})
    all_only_codes = [r_['code'] for r_ in rows_only + rows_extra]
    checks = {
        'tol': TOL,
        'overlapN': len(overlap), 'overlapNewN': len(overlap_new), 'ssN': len(ss_list), 'fwOldN': len(old36),
        'weeklyMaxDiff': max((r['weeklyMaxDiff'] for r in rows), default=0.0),
        'weeklyMaxDiffAll': max((max(abs(x - y) for x, y in zip(ssw[c]['w'], fww[c]['w'])) for c in ss_list if c in ssw and c in fww), default=0.0),
        'boundaryFail': [r['code'] for r in rows if not r['boundaryOk']],
        'stockMaxDiff': max((abs(r['stockDiff']) for r in rows), default=0.0),
        'stockMaxDiffAll': max((abs(ssw[c]['stock'] - fww[c]['stock']) for c in ss_list if c in ssw and c in fww), default=0.0),
    }
    tot = lambda key, sel: sum(items[g][sel][key] for g in ITEMS)
    checks['carryAll'] = {k: tot(k, 'carryAll') for k in CARRY}
    checks['carryOverlap'] = {k: tot(k, 'carryOverlap') for k in CARRY}
    checks['carryOverlapNew'] = {k: tot(k, 'carryOverlapNew') for k in CARRY}
    checks['onlyBasis'] = {k: sum(items[g]['onlyBasis'][k] for g in ITEMS) for k in BASIS_KEYS}       # 구품번 추가(FW 구품번 16 + 러닝 원본 13)를 SS 기준(1/1~)으로 센 합계
    sumb = lambda rs: {k: sum((r_['basis'] or {}).get(k, 0.0) for r_ in rs) for k in BASIS_KEYS}
    checks['onlyFwBasis'], checks['onlyExtraBasis'] = sumb(rows_only), sumb(rows_extra)
    checks['onlyN'], checks['onlyFwN'], checks['onlyExtraN'] = len(rows_only) + len(rows_extra), len(rows_only), len(rows_extra)
    checks['onlyResidMax'] = max((abs(r_['basis']['resid']) for r_ in rows_only + rows_extra if r_['basis']), default=0.0)
    checks['onlyBasisKnown'] = sum(1 for r_ in rows_only + rows_extra if r_['basis'])
    checks['onlyItemUnknown'] = [r_['code'] for r_ in rows_only + rows_extra if r_['item'] not in ITEMS]            # 복종을 못 정한 품번(있으면 합계에서 빠진다)
    checks['extraGrpMismatch'] = [r_['code'] for r_ in rows_extra if r_['grpFw'] and r_['grpFw'] != r_['item']]      # SS 품목과 FW 시트 품목이 다른 품번
    checks['extraStockDiffMax'] = max((abs(r_['stockDiff']) for r_ in rows_extra if r_['stockDiff'] is not None), default=0.0)      # SS 현재고 − FW 스타일맵 현재고(같은 ERP 추출이라 0 이어야 한다)
    checks['extraNoBase'] = [r_['code'] for r_ in rows_extra if r_['noBase']]
    checks['nasEmptySeasons'] = [k for k in SEASON_ORDER if nas_seasons and not nas_seasons.get(k)]       # 러닝 원본에서 비어 있는 시즌 열(예: 26FW)
    # 채널표(날짜 기준 1/1~)와 대조: 구품번 추가 29개의 ‘주마감 − 12월말 시트’ 판매 vs 채널표 — 12월말 시트에 빠진 소급 입력분이 있으면 앞쪽이 커진다
    have = [r_ for r_ in rows_only + rows_extra if r_['basis'] and r_.get('chSales') is not None]
    checks['onlyChN'] = len(have)
    checks['onlyChMissing'] = [r_['code'] for r_ in rows_only + rows_extra if r_['basis'] and r_.get('chSales') is None]
    checks['onlyChBasisSum'] = sum(r_['basis']['sales'] for r_ in have)
    checks['onlyChSum'] = sum(r_['chSales'] for r_ in have)
    checks['onlyChOut'] = [{'code': r_['code'], 'src': r_['src'], 'sales': r_['basis']['sales'], 'ch': r_['chSales']} for r_ in have if abs(r_['basis']['sales'] - r_['chSales']) > TOL]
    checks['onlyChBelow'] = [r_['code'] for r_ in have if r_['basis']['sales'] < r_['chSales'] - TOL]       # 채널표보다 작은 품번(있으면 ‘소급 입력’ 설명과 반대)
    checks['onlyChCostDiff'] = sum((r_['basis']['sales'] - r_['chSales']) * (r_['basis']['cost'] / r_['basis']['qtyIn'] if r_['basis']['qtyIn'] else 0.0) for r_ in have)      # 채널표 대비 판매 초과분의 원가(VAT 미포함, 품번별 평균 원가 × 초과 장수)
    ss_ch = [c for c in ss_list if c in chn]
    checks['ssChRecSum'] = sum(ss_rec[c] for c in ss_ch)                                    # SS 러닝 31 판매기록 기준(주마감 − 12월말 시트) 합
    dif = [ssb[c]['sales'] - ss_rec[c] for c in ss_list if c in ssb]                       # 스타일맵 정의 − 판매기록 기준(품번별)
    checks['ssRecDownN'], checks['ssRecDownSum'] = sum(1 for d_ in dif if d_ < -0.5), sum(d_ for d_ in dif if d_ < -0.5)
    checks['ssRecUpN'], checks['ssRecUpSum'] = sum(1 for d_ in dif if d_ > 0.5), sum(d_ for d_ in dif if d_ > 0.5)
    # 러닝이 아닌 이월 품번(SS 러닝 31·구품번 추가 29 어디에도 없는 주마감 품번)의 1/1~ 판매 — 연간 합계에 들어 있지 않은 규모를 알린다
    def sales_by(ws_):
        o_ = collections.defaultdict(float)
        for r in ws_.iter_rows(min_row=4, max_row=6000, max_col=70, values_only=True):
            if len(r) > 58 and r[12]:
                o_[txt(r[12])] += num(r[58])
        return o_
    s_now, s_dec = sales_by(wb['런닝 주마감(ST-생,판,재)']), sales_by(wb['12월기초재고(런닝)'])
    run_set = set(ss_list) | set(all_only_codes)
    checks['nonRunN'] = sum(1 for c in s_now if c not in run_set)
    checks['nonRunSales'] = sum(s_now[c] - s_dec.get(c, 0.0) for c in s_now if c not in run_set)
    checks['ssChN'] = len(ss_ch)
    checks['ssChBackSum'] = sum(ssb.get(c, {}).get('sales', 0.0) for c in ss_ch)                 # SS 러닝 31(스타일맵 정의) 판매 합
    checks['ssChSum'] = sum(chn[c] for c in ss_ch)
    # SS 러닝 31 원가 재평가: 스타일맵 정의(올해재고 × 12월말 단가) 원가 − 판매원가 vs 주마감 재고원가
    lf_ss = lifefull(wb['런닝 주마감(ST-생,판,재)'], set(ss_list))
    checks['ssStockCostGap'] = sum(ssb[c]['cost'] - ssb[c]['salesCost'] for c in ss_list if c in ssb) - sum(lf_ss[c]['stockCost'] for c in ss_list if c in lf_ss)      # (스타일맵 정의 원가 − 판매원가) − 주마감 재고원가, VAT 미포함
    checks['ssStockCostTop'] = max(((abs((ssb[c]['cost'] - ssb[c]['salesCost']) - lf_ss[c]['stockCost']), c) for c in ss_list if c in ssb and c in lf_ss), default=(0, ''))[1]      # 재평가 차이가 가장 큰 SS 러닝 품번

    fw_new = {txt(r[1]) for r in fw_wb['26FW 러닝스타일 리스트'].iter_rows(min_row=2, max_row=300, max_col=2, values_only=True) if r and len(r) > 1 and r[1]}      # FW 신품번 목록(B열)
    checks['extraInFwNew'] = sorted(set(extra_codes) & fw_new)
    new_blocks = set(ssn) | {txt(r[14]) for r in fw_wb['생판재-26FW(ST)'].iter_rows(min_row=4, max_row=4997, max_col=16, values_only=True) if len(r) > 14 and r[14]}      # SS 신상품('생,판,재') + FW 신상품(생판재-26FW)
    all_only = [r_['code'] for r_ in rows_only + rows_extra]
    checks['onlyInNew'] = sorted(set(all_only) & new_blocks)                          # 구품번 추가가 신상품 합계에 이미 들어 있으면 두 번 세는 것
    checks['onlyInSsRun'] = sorted(set(all_only) & set(ss_list))                       # 구품번 추가가 SS 러닝 31개에도 있으면 두 번 세는 것
    checks['onlyInFwNew'] = sorted(set(all_only) & fw_new)                             # 구품번 추가가 FW 신품번에 있으면 두 번 세는 것
    checks['extraSeasons'] = [k for k in SEASON_ORDER if any(k in r_['nasSeasons'] for r_ in rows_extra)]                     # 러닝 원본에서 이 품번들이 올라 있는 시즌 목록
    checks['extraIn26'] = [r_['code'] for r_ in rows_extra if any(k in ('26SS', '26FW') for k in r_['nasSeasons'])]      # 26SS·26FW 러닝 목록에도 있으면 알린다(없어야 ‘과년 러닝’)                  # 러닝 원본 추가분이 FW 신품번에도 있으면 두 번 세는 것(없어야 정상)
    nas_set = set(nas_all or [])
    checks['nasCover'] = {'n': len(nas_set), 'ss': len(nas_set & set(ss_list)), 'fwOld': len(nas_set & set(old36) - set(ss_list)), 'extra': len(nas_set & set(extra_codes)),
                          'none': sorted(nas_set - set(ss_list) - set(old36) - set(extra_codes))}      # 러닝 원본 품번이 연간 러닝에 모두 반영됐는지
    checks['onlyWeeklyOut'] = [{'code': r_['code'], 'src': r_['src'], 'sales': r_['basis']['sales'], 'lo': r_['wkLo'], 'hi': r_['wkHi']} for r_ in rows_only + rows_extra if r_['basis'] and not r_['wkOk']]     # 주간판매 범위(±50장) 밖
    checks['onlyWeeklyN'] = sum(1 for r_ in rows_only + rows_extra if r_['basis'])
    checks['onlyWeeklyFwN'], checks['onlyWeeklyExtraN'] = sum(1 for r_ in rows_only if r_['basis']), sum(1 for r_ in rows_extra if r_['basis'])
    checks['unlisted'] = unlisted_rows                                   # 자료가 없어 못 넣은 품번
    checks['boundaryFailNew'] = [r['code'] for r in rows_new if not r['boundaryOk']]
    checks['stockMaxDiffNew'] = max((abs(r['stockDiff']) for r in rows_new), default=0.0)
    checks['weeklyMaxDiffNew'] = max((r['weeklyMaxDiff'] for r in rows_new if r['weeklyMaxDiff'] is not None), default=0.0)
    checks['ssRunInNew'] = sorted(set(ss_list) & set(ssn))                    # SS 러닝 품번이 26SS 신상품에도 있으면 안 된다(없어야 정상)
    checks['overlapNewSsSales'] = sum(r['ssSales'] or 0 for r in rows_new)
    checks['overlapNewSalesDiffMax'] = max((abs(r['salesDiff']) for r in rows_new), default=0.0)       # SS 신상품 판매(생,판,재) − FW 러닝 ST 누계 판매
    checks['overlapNewSsIn'] = sum(r['ssIn'] or 0 for r in rows_new)
    if fw_stk_total:
        checks['carryVsSheet'] = {k: [checks['carryAll'][k], fw_stk_total.get(k)] for k in CARRY if fw_stk_total.get(k) is not None}
        checks['carryMatchesSheet'] = all(abs((fw_stk_total.get(k) or 0) - checks['carryAll'][k]) <= max(1.0, abs(fw_stk_total.get(k) or 0) * 1e-4) for k in ('qtyIn', 'sales', 'salesAmt', 'wkQty'))
    checks['fwOnlyJanJulSales'] = sum(r['janJulSales'] or 0 for r in rows_only)                 # FW 전용 구품번의 1/1~7/31 판매(연간 러닝에 들어 있지 않은 몫)
    checks['fwOnlyJanJulKnown'] = sum(1 for r in rows_only if r['janJulSales'] is not None)
    checks['overlapSsSales'] = sum(r['ssSales'] or 0 for r in rows)
    checks['overlapSsSalesRec'] = sum(r['ssSalesRec'] or 0 for r in rows)                  # 판매기록 기준(주마감 누계 − 12월기초재고 누계)
    checks['ssSalesBackTotal'] = sum((ssb.get(c) or {}).get('sales') or 0 for c in ss_list)  # 스타일맵 종합 정의(런닝 BACK 판매량 = 올해재고 − 현재고)
    checks['ssSalesRecTotal'] = sum(ss_rec.values())
    checks['ssSalesDiffMax'] = max((abs(r['ssSalesDiff']) for r in rows), default=0.0)
    checks['ssSalesDiffOver'] = [r['code'] for r in rows if abs(r['ssSalesDiff']) > TOL]
    checks['overlapSsIn'] = sum(r['ssIn'] or 0 for r in rows)
    checks['weekLabels'] = [wl[week_start_idx], wl[week_start_idx + 1]] if len(wl) > week_start_idx + 1 else []
    return {'overlap': rows, 'overlapNew': rows_new, 'fwOnly': rows_only + rows_extra, 'ssAll': rows_all, 'items': items, 'checks': checks}
