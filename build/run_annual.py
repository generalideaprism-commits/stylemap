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
GRP = {'OUTER': 'OUTER', 'KNIT': 'TOP', 'INNER (W/O KNIT)': 'TOP', 'BOTTOM': 'BOTTOM', 'ACC': 'ACC'}
ITEMS = ('OUTER', 'TOP', 'BOTTOM', 'ACC')
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


def build(fw_wb, ss_path, ss_list, old36, fw_stk_total=None, week_start_idx=40):                # noqa: C901
    """fw_wb: 열린 26FW 스타일맵, ss_path: 26SS 스타일맵 경로, ss_list: 26SS 러닝 품번, old36: 26FW 러닝 구품번 목록.
    week_start_idx: 7/27 주의 위치(W52=0 … W1=51 중 W12=40). fw_stk_total: 종합 시트 CJ~CR 러닝 합계(대조용)."""
    wb = openpyxl.load_workbook(ss_path, read_only=True, data_only=True)
    overlap = sorted(set(ss_list) & set(old36))
    # --- SS: 런닝 BACK(스타일 단위) ---
    ssb = {}
    for r in wb['런닝 BACK'].iter_rows(min_row=3, max_row=60, max_col=40, values_only=True):
        if len(r) > 10 and r[3] and txt(r[3]) in ss_list:
            ssb[txt(r[3])] = dict(item=txt(r[2]), base=num(r[7]), neworder=num(r[8]), qtyIn=num(r[9]), sales=num(r[10]), cost=num(r[12]), salesAmt=num(r[15]), stock=num(r[0]))
    _, ssw = weekly(wb['런닝주간판매(CO)'], 21, 41)

    def life(ws):                                                          # 스타일별 누계 판매(판매기간 수량, 58열)
        out = collections.defaultdict(float)
        for r in ws.iter_rows(min_row=4, max_row=6000, max_col=70, values_only=True):
            if len(r) > 58 and r[12] and (txt(r[12]) in ss_list or txt(r[12]) in old36):
                out[txt(r[12])] += num(r[58])
        return out
    life_now, life_12 = life(wb['런닝 주마감(ST-생,판,재)']), life(wb['12월기초재고(런닝)'])
    ss_rec = {c: life_now[c] - life_12[c] for c in ss_list}                # 판매기록 기준 12월말 이후 판매
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

    items = {g: {'carryAll': {k: 0.0 for k in CARRY}, 'carryOverlap': {k: 0.0 for k in CARRY}, 'carryOverlapNew': {k: 0.0 for k in CARRY}} for g in ITEMS}
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
    rows_only = []                                                        # FW 전용 이월(SS 러닝·SS 신상품 어느 쪽에도 없는 구품번)
    for c in old36:
        if c in overlap or c in overlap_new:
            continue
        cv = carry(c)
        jj = (sn[c]['sales'] - life_12[c]) if c in life_12 else None                  # 1/1~7/31 판매 = 7/31 시점 누계 − 12월말 누계(SS 스타일맵 12월기초재고 시트에 있는 품번만)
        rows_only.append({'code': c, 'name': name.get(c, ''), 'item': st[c]['grp'] if c in st else '?', 'fwCarryStock': cv['qtyIn'], 'fwAugSales': cv['sales'], 'fwStock': st[c]['stock'] if c in st else 0.0, 'janJulSales': jj})
    rows_all = []                                                         # SS 러닝 전 품번의 두 파일 주간판매·재고 대조(겹침 여부와 무관)
    for c in ss_list:
        sw, fw = ssw.get(c, {'w': [0.0] * 52, 'stock': 0.0}), fww.get(c, {'w': [0.0] * 52, 'stock': 0.0})
        rows_all.append({'code': c, 'overlap': c in overlap, 'ssSalesBack': ssb.get(c, {}).get('sales'), 'ssSalesRec': ss_rec.get(c), 'ssWk0727': sum(sw['w'][week_start_idx:]), 'fwWk0727': sum(fw['w'][week_start_idx:]),
                         'weeklyMaxDiff': max(abs(x - y) for x, y in zip(sw['w'], fw['w'])), 'ssStock': sw['stock'], 'fwStock': fw['stock']})
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
    return {'overlap': rows, 'overlapNew': rows_new, 'fwOnly': rows_only, 'ssAll': rows_all, 'items': items, 'checks': checks}
