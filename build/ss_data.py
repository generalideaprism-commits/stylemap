# -*- coding: utf-8 -*-
"""26SS 스타일맵(NAS '4.26SS/04. 스타일맵 (주별 판매)') -> 시즌 집계표 제안안의 '연간(SS+FW)' 탭용 SS 데이터

읽는 곳
  · 'G.I 종합' 시트      — 시즌(봄·여름)×복종 현황(기획·입고·판매·원가) + 러닝 + 전체 합계. 26FW 스타일맵 '종합' 시트와 같은 배치(G~AA 열).
  · '전년비교BACK(1.1~현재)' — 25SS 동시점(2025-10-19까지) 스타일별 원본. 종합 시트의 '전년비교 25년 동요일'(AF~AK)은 입고·판매·금액만 있어서
                              기획·원가·판매원가·주간판매까지 필요한 전년 값은 이 시트를 (년도=2025, 시즌, 복종)으로 합쳐 만든다
                              (합친 입고량·판매량·실판매금액이 종합 시트 전년 열과 일치하는지 매번 확인한다 — 파싱이 맞는지 보는 것이지 기준일이 같다는 증명은 아니다).
                              사은품(품번에 U98 + 판매금액 0원)은 26FW 빌더와 같은 규칙으로 25SS 합에서 뺀다(26SS 종합 시트는 사은품이 4복종 밖이라 이미 빠져 있다).
                              종합 시트 전년 열(AF~AK)에는 사은품이 들어 있어서, 대조는 사은품을 빼기 전 합으로 한다.
  · '런닝' 시트          — SS 러닝 품번 31개(겹침 확인용)
결과는 build/ss_snapshot.json 에도 저장한다(NAS 가 안 열리는 PC 에서는 이 파일로 빌드).

원가 규칙(MD 확정): 원천 원가는 VAT 미포함 → 화면 표시용 costV·planCostV·salesCostV 는 ×1.1(VAT 포함). 원천 값(cost 등)은 그대로 둔다.
"""
import json
import os
import re
import sys

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
SS_DIR = os.environ.get('SS_DIR') or '//DiskStation/운영기획팀/@기획/4.26SS/04. 스타일맵 (주별 판매)'
SNAPSHOT = os.path.join(HERE, 'ss_snapshot.json')


def C(letter):
    n = 0
    for ch in letter:
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def num(v):
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def txt(v):
    return '' if v is None else str(v).strip()


# 종합 시트 지표 열 — 26FW 종합 시트와 같은 열(G~AA). 전년비교(25년 동요일)는 AF~AK.
MET = dict(styles='G', sku='H', plan='I', inM='J', inSku='K', qtyIn='L', inRate='M',
           planCost='N', cost='O', tag='Q', mu='S', sales='T', sellRate='U',
           salesTag='V', salesAmt='W', salesCost='X', disc='Y', wkQty='Z', wkAmt='AA')
PREV = dict(qtyIn='AF', sales='AG', sellRate='AH', salesTag='AI', salesAmt='AJ', disc='AK')
# 전년비교BACK 시트(0-based 열): 기획 34~36(수량·TAG금액·원가금액) · 생산(입고) 43~45 · 판매(최근=주간) 56~59 · 판매(기간=누계) 60~63
BK = dict(plan=34, planCost=36, qtyIn=43, tag=44, cost=45, wkQty=56, wkAmt=58, sales=60, salesTag=61, salesAmt=62, salesCost=63)
ITEMS = ('OUTER', 'TOP', 'BOTTOM', 'ACC')
GIFT_EXTRA_25SS = ('WBE1L95512',)         # 25SS 삭스 세트(판매금액 0원) — MD 지시 2026-10-22 로 사은품과 같이 뺀다
SS_EXCLUDE_CODES = ('QBF1U98001',)        # 26SS 현재 값에서 뺄 사은품성 품번: GNRL 2026 다이어리(U98 + 판매금액 0원) — MD 지시 2026-10-22
GRP_OF = {'OUTER': 'OUTER', 'KNIT': 'TOP', 'INNER (W/O KNIT)': 'TOP', 'BOTTOM': 'BOTTOM', 'ACC': 'ACC'}     # 생,판,재 '품목' -> 4복종 (기타=사은품 등은 뺀다)
CAT_KEYS = ('styles', 'sku', 'plan', 'inM', 'inSku', 'qtyIn', 'planCost', 'cost', 'tag', 'sales', 'salesTag', 'salesAmt', 'salesCost', 'wkQty', 'wkAmt')
SUB_KEYS = ('styles', 'sku', 'plan', 'inM', 'inSku', 'qtyIn', 'planCost', 'cost', 'tag', 'sales', 'salesTag', 'salesAmt', 'salesCost', 'wkQty', 'wkAmt')


def derive(v):
    """원가 VAT 포함 표시값 · 마크업 · 비율 — 26FW 빌더 vals() 와 같은 규칙"""
    cost, pc, sc, tag, tg = v.get('cost'), v.get('planCost'), v.get('salesCost'), v.get('tag'), v.get('salesTag')
    v['costV'] = round(cost * 1.1) if cost else None
    v['planCostV'] = round(pc * 1.1) if pc else None
    v['salesCostV'] = round(sc * 1.1) if sc else None
    v['mu'] = round(tag / cost / 1.1, 2) if tag and cost else None
    v['muReal'] = round(v['salesAmt'] / sc / 1.1, 2) if v.get('salesAmt') and sc else None
    v['sellRate'] = (v.get('sales') or 0) / v['qtyIn'] if v.get('qtyIn') else v.get('sellRate')
    v['disc'] = 1 - (v.get('salesAmt') or 0) / tg if tg else v.get('disc')
    if v.get('plan'):
        v['inRate'] = (v.get('qtyIn') or 0) / v['plan']
    return v


def find_latest(folder=SS_DIR):
    try:
        cands = []
        for f in os.listdir(folder):
            m = re.match(r'^26SS 스타일맵 (\d{4})\.xlsx$', f)
            if m and not f.startswith('~$'):
                mmdd = m.group(1)
                cands.append((1 if mmdd[:2] in ('01', '02') else 0, mmdd, f))        # 연초(1~2월) 파일이 12월 파일보다 새것
        if not cands:
            return None
        return os.path.join(folder, sorted(cands)[-1][2])
    except OSError:
        return None


def gender_rows(wb):
    """26SS 성별(여성 / 유니맨=남성+공용) 값 — 스타일별 시트 '생,판,재'(칼라 단위 행)를 복종·카테고리로 합친다. 사은품·제외 품번은 뺀다.
    전년(25SS) 성별 비교는 만들지 않는다 — 전년비교BACK 의 성별 열은 분류 기준이 올해와 같은지 확인되지 않는다(유니맨 입고 25SS 1.8만 vs 26SS 9.7만으로 크게 다름)."""
    acc = {}
    sty = {}
    for r in wb['생,판,재'].iter_rows(min_row=4, max_row=1200, max_col=86, values_only=True):
        if not (len(r) > 74 and r[23] and str(r[13]) == '2026' and txt(r[14]) in ('봄', '여름')) or txt(r[23]) in SS_EXCLUDE_CODES:
            continue
        grp = GRP_OF.get(txt(r[16]))
        if not grp:
            continue
        g = '여성' if txt(r[15]) == '여성' else '남성'
        cat = txt(r[18])
        v = {'plan': num(r[45]) or 0.0, 'planCost': num(r[47]) or 0.0, 'qtyIn': num(r[54]) or 0.0, 'tag': num(r[55]) or 0.0, 'cost': num(r[56]) or 0.0,
             'wkQty': num(r[67]) or 0.0, 'wkAmt': num(r[69]) or 0.0, 'sales': num(r[71]) or 0.0, 'salesTag': num(r[72]) or 0.0, 'salesAmt': num(r[73]) or 0.0, 'salesCost': num(r[74]) or 0.0}
        v['sku'] = 1.0 if v['plan'] > 0 else 0.0
        v['inSku'] = 1.0 if v['qtyIn'] > 0 else 0.0
        for key in ((g, grp), (g, grp, cat)):
            a = acc.setdefault(key, {k: 0.0 for k in CAT_KEYS})
            for k, x in v.items():
                a[k] += x
            st = sty.setdefault(key, {'p': set(), 'i': set()})
            if v['plan'] > 0:
                st['p'].add(txt(r[23]))
            if v['qtyIn'] > 0:
                st['i'].add(txt(r[23]))
    for key, a in acc.items():
        a['styles'] = float(len(sty[key]['p']))
        a['inM'] = float(len(sty[key]['i']))
    items, cats_ = [], []
    for g in ('여성', '남성'):
        tot = {k: 0.0 for k in CAT_KEYS}
        for grp in ITEMS:
            a = acc.get((g, grp)) or {k: 0.0 for k in CAT_KEYS}
            for k in CAT_KEYS:
                tot[k] += a[k]
            items.append({'gender': g, 'item': grp, 'cur': derive(dict(a)), 'py': None})
        items.append({'gender': g, 'item': 'TOTAL', 'cur': derive(tot), 'py': None})
    for (g, grp, cat) in sorted(k for k in acc if len(k) == 3):
        cats_.append({'gender': g, 'item': 'ACCESSORY' if grp == 'ACC' else grp, 'cat': cat, 'code': '', 'cur': derive(dict(acc[(g, grp, cat)])), 'py': None})
    return items, cats_


def exclude_current(wb, block1, codes, new_cat=None, season_cat=None):
    """26SS 종합 시트 합계에서 사은품성 품번의 값을 뺀다. 값은 스타일별 시트 '생,판,재'(칼라 단위 행)에서 합쳐 읽고,
    해당 복종 행 · 시즌 TTL · 신상품 TTL · 전체 TTL 에서 차감한다(카테고리 행은 화면에서 쓰지 않아 건드리지 않는다)."""
    per = {}
    for r in wb['생,판,재'].iter_rows(min_row=4, max_row=1200, max_col=86, values_only=True):
        if len(r) > 74 and txt(r[23]) in codes and str(r[13]) == '2026':
            a = per.setdefault(txt(r[23]), dict(season=txt(r[14]), item=txt(r[17]), cat=txt(r[18]), name=txt(r[24]), styles=1, sku=0, inM=0, inSku=0, plan=0.0, planCost=0.0,
                                                qtyIn=0.0, tag=0.0, cost=0.0, sales=0.0, salesTag=0.0, salesAmt=0.0, salesCost=0.0, wkQty=0.0, wkAmt=0.0))
            for k, j in (('plan', 45), ('planCost', 47), ('qtyIn', 54), ('tag', 55), ('cost', 56), ('wkQty', 67), ('wkAmt', 69), ('sales', 71), ('salesTag', 72), ('salesAmt', 73), ('salesCost', 74)):
                a[k] += num(r[j]) or 0.0
            a['sku'] += 1 if (num(r[45]) or 0) > 0 else 0
            a['inSku'] += 1 if (num(r[54]) or 0) > 0 else 0
    done = []
    for code, v in per.items():
        v['inM'] = 1 if v['qtyIn'] > 0 else 0
        hit = 0
        for rec in block1:
            se, it = v['season'], v['item']
            if (rec['mode'] == 'new' and ((rec['season'] == se and rec['item'] == it and rec.get('kind') == '항목') or rec.get('kind') in ('%s TTL' % se, '신상품 TTL'))) or rec.get('kind') == '전체 TTL':
                c = rec['cur']
                for k in SUB_KEYS:
                    if c.get(k) is not None and v.get(k):
                        c[k] -= v[k]
                derive(c)
                hit += 1
        for lst, by_season in ((new_cat or [], False), (season_cat or [], True)):             # 카테고리 행(시즌·복종·카테고리 화면용)에서도 뺀다
            for rec in lst:
                if rec['cat'] == v['cat'] and rec['item'][:3] == v['item'][:3] and (not by_season or rec.get('season') == v['season']):
                    for k in SUB_KEYS:
                        if rec['cur'].get(k) is not None and v.get(k):
                            rec['cur'][k] -= v[k]
                    derive(rec['cur'])
        done.append({'code': code, 'name': v['name'], 'season': v['season'], 'qtyIn': round(v['qtyIn']), 'costV': round(v['cost'] * 1.1), 'rows': hit})
    return done


def run_overlap(wb, ss_list, fw_codes, end_iso, since='2026-07-27'):
    """SS 러닝 품번과 26FW 러닝 목록(구품번)이 겹치는 품번, 그 품번들의 since(26FW 러닝 집계 시작 주)~ 판매량(SS 주간판매 시트 합).
    SS 러닝 누계(1/1~)에는 이 판매가 이미 들어 있고, FW 러닝 '시점재고 합산'(구품번 8/1~)도 같은 판매를 세므로 연간 합계에서 러닝을 더하면 이중 집계가 된다."""
    import datetime
    over = sorted(set(ss_list) & set(c for c in fw_codes if c))
    out = {'ssList': len(ss_list), 'fwList': len(set(c for c in fw_codes if c)), 'overlap': len(over), 'overlapStyles': over, 'overlapQty': None, 'overlapQtyLo': None, 'since': since}
    if not over:
        return out
    try:
        ws = wb['런닝주간판매(CO)']
        rows = list(ws.iter_rows(min_row=1, max_row=6000, values_only=True))
        hdr = rows[1]
        si = hdr.index('스타일')
        wcols = [j for j, h in enumerate(hdr) if h and re.match(r'^W\d+\(', str(h))]
        end = datetime.date.fromisoformat(end_iso)
        n = len(wcols)
        starts = {j: end - datetime.timedelta(days=6 + 7 * (n - 1 - i)) for i, j in enumerate(wcols)}      # 마지막 W 열 = 기준일이 든 주(월~일)
        lo = datetime.date.fromisoformat(since)
        tot = tot_lo = 0.0
        sset = set(over)
        lo2 = lo + datetime.timedelta(days=7)                                  # 7/27 주를 빼면(8/3 주부터) 하한
        for r in rows[3:]:
            if len(r) > si and r[si] in sset:
                tot += sum((num(r[j]) or 0.0) for j in wcols if starts[j] >= lo)
                tot_lo += sum((num(r[j]) or 0.0) for j in wcols if starts[j] >= lo2)
        out['overlapQty'] = round(tot)
        out['overlapQtyLo'] = round(tot_lo)
    except Exception as e:
        print('[경고] 러닝 겹침 판매량을 계산하지 못했습니다:', type(e).__name__, e)
    return out


def parse(path, fw_codes=None):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb['G.I 종합']
    rows = list(ws.iter_rows(min_row=1, max_row=170, values_only=True))

    def cell(r, col):
        row = rows[r - 1]
        i = C(col)
        return row[i] if len(row) > i else None

    def vals(r, cols):
        v = {k: num(cell(r, c)) for k, c in cols.items()}
        return derive(v)

    end = cell(2, 'D')
    prev_end = cell(2, 'F')
    iso = lambda d: d.strftime('%Y-%m-%d') if hasattr(d, 'strftime') else txt(d)[:10]

    # --- 전년비교BACK -> (시즌, 복종)별 합 ---
    bs = wb['전년비교BACK(1.1~현재)']
    back, back_all, gifts, gsum = {}, {}, [], {'n': 0, 'qtyIn': 0.0, 'costV': 0.0}
    back_cat = {}                                    # 카테고리별 (시즌, 복종, 아이템) — 사은품 제외
    for r in bs.iter_rows(min_row=4, max_row=1200, max_col=71, values_only=True):
        if not r[14] or str(r[4]) != '2025' or txt(r[5]) not in ('봄', '여름') or txt(r[0]) not in ITEMS:
            continue
        gift = ('U98' in txt(r[14]) and not (num(r[BK['salesAmt']]) or 0)) or txt(r[14]) in GIFT_EXTRA_25SS      # 사은품 — 판매금액이 0원인 U98 품번 + 지정 품번
        if gift and (num(r[BK['qtyIn']]) or 0) > 0:
            gifts.append('%s %s' % (txt(r[14]), txt(r[15])))
            gsum['n'] += 1
            gsum['qtyIn'] += num(r[BK['qtyIn']]) or 0.0
            gsum['costV'] += (num(r[BK['cost']]) or 0.0) * 1.1
        for tgt in ((back_all,) if gift else (back, back_all)):
            a = tgt.setdefault((txt(r[5]), txt(r[0])), {k: 0.0 for k in BK})
            for k, j in BK.items():
                a[k] += num(r[j]) or 0.0
        if not gift:
            a = back_cat.setdefault((txt(r[5]), txt(r[0]), txt(r[9])), {k: 0.0 for k in BK})
            for k, j in BK.items():
                a[k] += num(r[j]) or 0.0

    def py_of(season, item, src=None):
        src = back if src is None else src
        if item == 'TTL':
            keys = [(s, it) for s in ([season] if season else ('봄', '여름')) for it in ITEMS]
        else:
            keys = [(season, item)]
        t = {k: 0.0 for k in BK}
        hit = False
        for k in keys:
            if k in src:
                hit = True
                for kk in t:
                    t[kk] += src[k][kk]
        if not hit:
            return None
        # 전년비교BACK 의 '기획'·'판매(최근)' 은 종합 시트의 기획량·주간판매와 정의가 달라(입고율 68% vs 94% 로 어긋남) 쓰지 않는다 — 전년 값은 입고·판매·금액·원가만
        for k in ('plan', 'planCost', 'wkQty', 'wkAmt'):
            t.pop(k, None)
        return derive(t)

    def py_sum(src, keys):
        """src 의 keys 합(전년비교BACK 기준) -> py 값. 하나도 없으면 0 으로 채운 값(= 전년에 없던 구분)"""
        t = {k: 0.0 for k in BK}
        for k in keys:
            if k in src:
                for kk in t:
                    t[kk] += src[k][kk]
        for k in ('plan', 'planCost', 'wkQty', 'wkAmt'):
            t.pop(k, None)
        return derive(t)

    block1, season, mode, check = [], '', 'new', []
    for r in range(7, 24):
        c, d, e = txt(cell(r, 'C')), txt(cell(r, 'D')), txt(cell(r, 'E'))
        if d:
            season = d
        if c == '런닝':
            mode = 'run'
        rec = {'season': '' if mode == 'run' else season, 'item': e, 'mode': mode, 'cur': vals(r, MET)}
        sprev = {k: num(cell(r, col)) for k, col in PREV.items()}        # 종합 시트의 전년 열(입고·판매·금액·할인율만)
        if c == 'TTL':
            rec.update(item='TTL', season='', kind='신상품 TTL' if mode == 'new' else '전체 TTL')
            mode = 'run' if mode == 'new' else mode
        elif e == 'TTL':
            rec['kind'] = ('%s TTL' % season) if mode == 'new' else '러닝 TTL'
        elif e:
            rec['kind'] = '항목'
        else:
            continue
        if rec['mode'] == 'new' and rec['kind'] != '전체 TTL':
            rec['py'] = py_of(rec['season'], rec['item'])
            chk = py_of(rec['season'], rec['item'], back_all)             # 종합 시트 전년 열(사은품 포함)과 대조 — 사은품을 빼기 전 합으로
            if chk:
                for k in ('qtyIn', 'sales', 'salesAmt'):
                    a, b = chk.get(k), sprev.get(k)
                    if b and a is not None and abs(a - b) > max(1.0, abs(b) * 0.001):
                        check.append('%s %s %s: BACK %s ≠ 종합 %s' % (rec['season'] or '신상', rec['item'], k, round(a), round(b)))
        elif rec['mode'] == 'run':
            rec['py'] = derive({k: v for k, v in sprev.items() if v is not None}) if any(v for v in sprev.values()) else None
        block1.append(rec)

    # --- 복종·카테고리 현재값(전년 없음) — 신상품 전체(30~64행), 봄·여름 시즌별(71~144행) ---
    def cats(r0, r1, season=None):
        out, gubun = [], ''
        for r in range(r0, r1 + 1):
            d, e, f = txt(cell(r, 'D')), txt(cell(r, 'E')), txt(cell(r, 'F'))
            if d.endswith('TOTAL') or 'ACC제외' in d:
                continue
            if d:
                gubun = d
            if not f or 'S.TTL' in e:
                continue
            v = vals(r, MET)
            grp = 'ACC' if gubun.startswith('ACC') else gubun
            if any((v.get(k) or 0) for k in ('plan', 'qtyIn', 'sales')):
                rec = {'item': gubun, 'cat': f, 'code': e, 'cur': v}
                if season:
                    rec['season'] = season
                rec['py'] = py_sum(back_cat, [(season, grp, f)] if season else [(s_, grp, f) for s_ in ('봄', '여름')])
                out.append(rec)
        # 올해 행이 없는(또는 0 인) 카테고리라도 전년에 값이 있었으면 전년 값만 있는 행으로 둔다 — 복종 합계와 맞추려고
        have = {(x['item'][:3] if x['item'].startswith('ACC') else x['item'], x['cat']) for x in out}
        for (se, grp, cat) in sorted(back_cat):
            if (season and se != season) or (grp, cat) in have:
                continue
            ks = [(season, grp, cat)] if season else [(s_, grp, cat) for s_ in ('봄', '여름')]
            py = py_sum(back_cat, ks)
            if not any((py.get(k) or 0) for k in ('qtyIn', 'sales')):
                continue
            rec = {'item': 'ACCESSORY' if grp == 'ACC' else grp, 'cat': cat, 'code': '', 'cur': derive({k: 0.0 for k in MET}), 'py': py}
            if season:
                rec['season'] = season
            out.append(rec)
            have.add((grp, cat))
        return out

    new_cat = cats(30, 64)
    season_cat = cats(71, 104, '봄') + cats(110, 143, '여름')
    gender_item, gender_cat = gender_rows(wb)
    gifts26 = exclude_current(wb, block1, SS_EXCLUDE_CODES, new_cat, season_cat)
    # 성별 합 = 신상품 TTL 인지 확인(사은품성 품번 제외 후)
    ttl = next((x for x in block1 if x.get('kind') == '신상품 TTL'), None)
    if ttl:
        for g_ in ('qtyIn', 'cost', 'sales'):
            gs_ = sum((x['cur'].get(g_) or 0) for x in gender_item if x['item'] == 'TOTAL')
            ts_ = ttl['cur'].get(g_) or 0
            if ts_ and abs(gs_ - ts_) > max(1.0, abs(ts_) * 0.001):
                check.append('성별 합 %s: %s ≠ 신상품 TTL %s' % (g_, round(gs_), round(ts_)))

    # --- 러닝 품번(겹침 확인용) ---
    run_list = []
    for r in wb['런닝'].iter_rows(min_row=1, max_row=60, values_only=True):
        if len(r) > 2 and r[2] and re.match(r'^[A-Z]{2,3}[A-Z0-9]+$', str(r[2]).strip()):
            run_list.append(str(r[2]).strip())

    run = run_overlap(wb, run_list, fw_codes or [], iso(end))
    return {'file': os.path.basename(path), 'end': iso(end), 'prevEnd': iso(prev_end), 'gifts25': gifts, 'gifts25Sum': {k: round(v) for k, v in gsum.items()}, 'gifts26': gifts26,
            'block1': block1, 'newCat': new_cat, 'seasonCat': season_cat, 'genderItem': gender_item, 'genderCat': gender_cat, 'runList': run_list, 'run': run, 'check': check}


def load(path=None, quiet=False, fw_codes=None):
    """SS 데이터를 돌려준다. NAS 최신 파일 -> 스냅샷 저장, 안 열리면 저장된 스냅샷, 둘 다 없으면 빈 dict."""
    p = path or find_latest()
    data = None
    if p and os.path.exists(p):
        try:
            data = parse(p, fw_codes)
            json.dump(data, open(SNAPSHOT, 'w', encoding='utf-8'), ensure_ascii=False)
            if not quiet:
                print('26SS: %s (기준일 %s · 전년 %s) · 종합 %d행 · 카테고리 %d+%d행 · 러닝 품번 %d (FW 목록과 겹침 %s품번 · 겹침 판매 %s장)' % (
                    data['file'], data['end'], data['prevEnd'], len(data['block1']), len(data['newCat']), len(data['seasonCat']), len(data['runList']),
                    (data.get('run') or {}).get('overlap'), (data.get('run') or {}).get('overlapQty')))
                if data.get('gifts26'):
                    print('  26SS 현재 값에서 사은품성 품번 제외:', ', '.join('%s %s(입고 %s장)' % (g['code'], g['name'], g['qtyIn']) for g in data['gifts26']))
                if data.get('gifts25'):
                    print('  25SS 사은품 %d품번을 전년 합에서 제외: %s' % (len(data['gifts25']), ', '.join(data['gifts25'])))
                if data['check']:
                    print('  ⚠ 전년 값 대조 불일치:', data['check'])
        except Exception as e:                                           # 파일이 열려 있거나 형식이 바뀐 경우 — 스냅샷으로 간다
            print('[경고] 26SS 파일을 읽지 못했습니다(%s: %s) — 저장된 스냅샷을 씁니다' % (type(e).__name__, e))
    if data is None and os.path.exists(SNAPSHOT):
        data = json.load(open(SNAPSHOT, encoding='utf-8'))
        data['fromSnapshot'] = True
        if not quiet:
            print('26SS: 스냅샷 사용 — %s (기준일 %s)' % (data.get('file'), data.get('end')))
    if data is None:
        print('[경고] 26SS 자료가 없어 연간 탭에서 SS 는 비어 있습니다')
        return {}
    return data


if __name__ == '__main__':
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass
    d = load(sys.argv[1] if len(sys.argv) > 1 else None)
    if d:
        for r in d['block1']:
            c, py = r['cur'], r.get('py') or {}
            print('%-6s %-6s %-9s 입고 %9s 판매 %9s 원가(VAT+) %14s | 전년 입고 %9s 판매 %9s 원가 %14s' % (
                r['season'], r['item'], r.get('kind'), c.get('qtyIn') and round(c['qtyIn']), c.get('sales') and round(c['sales']), c.get('costV'),
                py.get('qtyIn') and round(py['qtyIn']), py.get('sales') and round(py['sales']), py.get('costV')))
