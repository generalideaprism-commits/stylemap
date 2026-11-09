# -*- coding: utf-8 -*-
"""허브 '26SS 시즌 판매 현황'(share/26ss-season-status.html)에 구글 시트 ERP 자동조회로 계산한 26SS 종합값을 한 주차로 넣는다 (2026-11-09 사용자).

  그 화면은 KWAN 이 NAS 26SS 스타일맵의 'G.I 종합' 시트 셀 값을 주별로 뜬 것. 26SS 판매자료 카드·연간 집계표가 구글 시트로 통일되면서
  이 화면도 같은 값(annual_gs 의 SS 블록)으로 채운다. 셀 배치·서식은 FW 와 같이 '엑셀에서 뜬 최신 주차'를 틀로 쓰고 값만 바꾼다.

  G.I 종합 배치(ss_data.py 와 같음): 7~23행 시즌×복종(+런닝), 30~64행 신상품 카테고리, 71~105 봄, 110~144 여름.
  열: G M · H SKU · I 기획량 · J 입고 M · K 입고 SKU · L 입고량 · M 입고율 · N 기획원가 · O 생산원가 · P 원가비중(O÷신상품 TTL O) · Q TAG · R 택가비중
      · S M/U(=Q÷O÷1.1) · T 누계판매량 · U 판매율 · V 판매TAG · W 실판매금액 · X 판매원가 · Y 할인율 · Z 주간판매량 · AA 주간판매금액
      · AB~AD 숨김(전주 대비 — 시트 수식이 금액을 집어 쓰는 칸이라 비운다) · AE 전년 라벨 · AF~AK 전년(입고·판매·판매율·TAG·실판매·할인율, 신상 행만)
"""
import datetime, hashlib, json, os, re

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
HUB = os.environ.get('HUB_REPO') or os.path.join(os.path.dirname(ROOT), 'production-hub')
PAGE = '26ss-season-status.html'
HUB_PAGE = os.path.join(HUB, 'share', PAGE)
OUT_PAGE = os.path.join(ROOT, 'share', PAGE)
SHEET_LABEL = "구글 시트 '대쉬보드PIP업데이트파일' → ERP 자동조회로 계산한 26SS 종합값"

import season_status as SS_
COLMAP = {'G': 'styles', 'H': 'sku', 'I': 'plan', 'J': 'inM', 'K': 'inSku', 'L': 'qtyIn', 'M': 'inRate', 'N': 'planCost', 'O': 'cost', 'P': 'costShare',
          'Q': 'tag', 'R': 'tagShare', 'S': 'mu', 'T': 'sales', 'U': 'sellRate', 'V': 'salesTag', 'W': 'salesAmt', 'X': 'salesCost', 'Y': 'disc', 'Z': 'wkQty', 'AA': 'wkAmt',
          'AF': 'py:qtyIn', 'AG': 'py:sales', 'AH': 'py:sellRate', 'AI': 'py:salesTag', 'AJ': 'py:salesAmt', 'AK': 'py:disc'}
BLANK_COLS = ('AB', 'AC', 'AD')
SUMK = ('styles', 'sku', 'plan', 'inM', 'inSku', 'qtyIn', 'planCost', 'cost', 'tag', 'sales', 'salesTag', 'salesAmt', 'salesCost', 'wkQty', 'wkAmt')


def t(v): return '' if v is None else str(v).strip()


def zero(): return {k: 0.0 for k in SUMK}


def add(a, b):
    for k in SUMK: a[k] += (b.get(k) or 0)
    return a


def fin(v, ref=None):
    """합계/비율 마무리 — ss_data.derive 와 같은 식. ref = 비중 분모 행."""
    v = dict(v)
    v['inRate'] = v['qtyIn'] / v['plan'] if v.get('plan') else None
    v['sellRate'] = v['sales'] / v['qtyIn'] if v.get('qtyIn') else None
    v['disc'] = 1 - v['salesAmt'] / v['salesTag'] if v.get('salesTag') else None
    v['mu'] = v['tag'] / v['cost'] / 1.1 if v.get('cost') else None
    v['costShare'] = v['cost'] / ref['cost'] if ref and ref.get('cost') else None
    v['tagShare'] = v['tag'] / ref['tag'] if ref and ref.get('tag') else None
    return v


def py_fin(p):
    if not p: return None
    q = dict(p); q['sellRate'] = q['sales'] / q['qtyIn'] if q.get('qtyIn') else None; q['disc'] = 1 - q['salesAmt'] / q['salesTag'] if q.get('salesTag') else None
    return q


def values_by_row(tpl, ss):
    """틀 주차의 행 라벨을 읽어 행번호 → {'cur': 값, 'py': 전년 값, 'blankPlan': 러닝 행 여부} 를 만든다."""
    cols = [c[0] for c in tpl['cols']]
    rows = {rn: {cols[c[0]]: c[1] for c in cells if (len(c) <= 3 or c[3] is None)} for rn, cells in tpl['rows']}
    b1 = {(b['season'], b['item'], b['kind']): b for b in ss['block1']}
    out = {}
    # 1) 7~23행
    season, mode = '', 'new'
    for rn in range(7, 24):
        lab = rows.get(rn, {}); c, d, e = lab.get('C', ''), lab.get('D', ''), lab.get('E', '')
        if d: season = d
        if c in ('런닝', '러닝'): mode = 'run'
        if c == 'TTL':
            key = ('', 'TTL', '신상품 TTL' if mode == 'new' else '전체 TTL'); mode = 'run' if mode == 'new' else mode
        elif e == 'TTL':
            key = (season if mode == 'new' else '', 'TTL', ('%s TTL' % season) if mode == 'new' else '러닝 TTL')
        elif e:
            key = (season if mode == 'new' else '', e, '항목')
        else:
            continue
        b = b1.get(key)
        if not b: continue
        out[rn] = {'cur': b['cur'], 'py': b.get('py'), 'run': key[2] in ('러닝 TTL',) or (mode == 'run' and key[2] == '항목'), 'all': key[2] == '전체 TTL', 'key': key}
    ttl_new = b1[('', 'TTL', '신상품 TTL')]['cur']
    for rn, o in out.items():
        o['cur'] = fin(o['cur'], None if o['run'] or o['all'] else ttl_new)
        if o['run'] or o['all']: o['py'] = None
    # 2) 카테고리 블록 (30~, 71~, 110~): D 복종(이어짐) · F 카테고리명 · 소계 행
    def cat_block(r0, r1, recs, ref_key):
        by = {(('ACC' if r['item'].startswith('ACC') else r['item']), r['cat']): r for r in recs}
        grp, acc_grp, acc_all, acc_knit, acc_wo, knit_done = '', zero(), zero(), zero(), zero(), False
        seq = []
        for rn in range(r0, r1 + 1):
            lab = rows.get(rn, {}); d, e, f = lab.get('D', ''), lab.get('E', ''), lab.get('F', '')
            if d and not d.endswith('TOTAL') and 'ACC제외' not in d and d.strip() != 'TOTAL':
                if d != grp: grp, acc_grp, acc_knit, acc_wo, knit_done = d, zero(), zero(), zero(), False
            g = 'ACC' if grp.startswith('ACC') else grp
            if d.strip() == 'TOTAL' and not f and not e: seq.append((rn, 'ALL')); continue
            if d.endswith('TOTAL'): seq.append((rn, 'GRP', dict(acc_grp))); continue
            if 'ACC제외' in d: seq.append((rn, 'EXACC')); continue
            if e == 'KNIT S.TTL': seq.append((rn, 'SUB', dict(acc_knit))); knit_done = True; continue
            if e == 'W/O KNIT S.TTL': seq.append((rn, 'SUB', dict(acc_wo))); continue
            if f:
                r = by.get((g, f))
                v = dict(r['cur']) if r else zero()
                seq.append((rn, 'CAT', v))
                add(acc_grp, v); add(acc_all, v)
                if g == 'TOP': add(acc_wo if knit_done else acc_knit, v)
            else:                             # 라벨이 없거나 카테고리명 없는 숨김 행 → 0 (틀에 숫자 칸이 있으면 0 으로 채운다)
                seq.append((rn, 'CAT', zero()))
        total = dict(acc_all)
        acc_sum = {}
        for rn, kind, *v in seq:
            if kind == 'GRP' and 'ACC' in (rows.get(rn, {}).get('D', '')): acc_sum = v[0]
        for rn, kind, *v in seq:
            if kind == 'ALL': cur = total
            elif kind == 'EXACC': cur = {k: total[k] - (acc_sum.get(k) or 0) for k in SUMK}
            else: cur = v[0]
            out[rn] = {'cur': fin(cur, total), 'py': None, 'run': False, 'all': False, 'key': (ref_key, rn)}
    cat_block(30, 64, ss['newCat'], 'newCat')
    cat_block(71, 105, [r for r in ss['seasonCat'] if r.get('season') == '봄'], '봄')
    cat_block(110, 144, [r for r in ss['seasonCat'] if r.get('season') == '여름'], '여름')
    return out


def build_week(tpl, ss, key, fname, ref):
    cols = [c[0] for c in tpl['cols']]
    F = SS_.Formats(tpl)
    vals = values_by_row(tpl, ss)
    rows_out, rep = [], {'cells': 0, 'changed': 0, 'unmapped': [], 'blank': 0}
    for rn, cells in tpl['rows']:
        out = []
        v = vals.get(rn)
        for c in cells:
            L = cols[c[0]]; c2 = list(c)
            has_raw = len(c) > 3 and isinstance(c[3], (int, float))
            if rn == 2 and L == 'D': c2[1] = ref.strftime('%m월 %d일').lstrip('0')
            elif rn == 2 and L == 'F': c2[1] = ref.replace(year=ref.year - 1).strftime('%m월 %d일').lstrip('0')
            elif has_raw or (v and L in COLMAP):
                if L in BLANK_COLS or rn == 24:
                    c2[1] = ''
                    if len(c2) > 3: c2[3] = None
                    rep['blank'] += 1
                elif v and L in COLMAP:
                    k = COLMAP[L]
                    if k.startswith('py:'):
                        x = (v['py'] or {}).get(k[3:]) if v['py'] else None
                        if v['py'] and k[3:] in ('sellRate', 'disc'): x = (py_fin(v['py']) or {}).get(k[3:])
                    else:
                        x = v['cur'].get(k)
                    if x is None or (v['run'] and k in ('styles', 'sku', 'plan', 'inRate', 'planCost', 'costShare', 'tagShare')):
                        c2[1] = ''
                        if len(c2) > 3: c2[3] = None
                    else:
                        fmt, z = F.of(c, rn)
                        xx = float(x)
                        if has_raw and abs(xx - float(c[3])) > 1e-9: rep['changed'] += 1
                        c2[1] = SS_.apply(fmt, xx, z)
                        if len(c2) > 3: c2[3] = xx if not xx.is_integer() else int(xx)
                        else: c2.append(xx if not xx.is_integer() else int(xx))
                    rep['cells'] += 1
                elif has_raw:
                    rep['unmapped'].append(L + str(rn))
            out.append(c2)
        rows_out.append([rn, out])
    week = {k: tpl[k] for k in tpl if k not in ('diff',)}
    week.update({'key': key, 'file': fname, 'rows': rows_out, 'warn': [], 'gen_by': 'stylemap-build',
                 'base': ref.isoformat(), 'prev': ref.replace(year=ref.year - 1).isoformat(),      # 전년 = 작년 같은 날(ERP 동시점 조회) — KWAN 판은 동요일이었다
                 'sig': hashlib.sha1(json.dumps(rows_out, ensure_ascii=False).encode()).hexdigest()[:10]})
    return week, rep


def update(ss, ref, src_name, hub_page=HUB_PAGE, out_page=OUT_PAGE, verbose=True):
    html, m, D = SS_.load_page(hub_page)
    weeks = D['weeks']
    key = (ref + datetime.timedelta(days=1)).strftime('%m%d') if False else ref.strftime('%m%d')      # 26SS 페이지의 key 는 기준 일요일 날짜(1108) — FW(파일 날짜) 와 다르다
    tpl = next((w for w in reversed(weeks) if not w.get('gen_by')), weeks[-1])
    week, rep = build_week(tpl, ss, key, src_name, ref)
    idx = next((i for i, w in enumerate(weeks) if w['key'] == key), None)
    if idx is None: weeks.append(week)
    else: weeks[idx] = week
    D['gen'] = datetime.datetime.now().strftime('%Y-%m-%d %H:%M')
    if SHEET_LABEL not in D.get('src', ''): D['src'] = (D.get('src') or '') + ' · 11/09 주부터 ' + SHEET_LABEL
    html2 = html[:m.start(1)] + json.dumps(D, ensure_ascii=False, separators=(',', ':')) + html[m.end(1):]
    open(out_page, 'w', encoding='utf-8').write(html2)
    if verbose:
        print('26SS 시즌 판매 현황: 주차 %s 를 %s (틀: %s %s) → %s' % (key, '바꿈' if idx is not None else '추가', tpl['key'], tpl['file'], os.path.relpath(out_page, ROOT)))
        print('  숫자 칸 %d (틀과 다른 값 %d) · 비운 칸(숨김 열·메모) %d · 매핑 못 한 칸 %d%s' % (rep['cells'], rep['changed'], rep['blank'], len(rep['unmapped']), (' 예: %s' % rep['unmapped'][:8]) if rep['unmapped'] else ''))
    return week, rep


if __name__ == '__main__':
    import sys
    sys.stdout.reconfigure(encoding='utf-8', errors='replace'); sys.path.insert(0, HERE)
    import gsource, annual_gs
    GS = gsource.load(verbose=False); ss, ann = annual_gs.build(GS, verbose=False)
    update(ss, GS['ref'], '구글 시트 ERP 자동조회 (%s 기준)' % GS['ref'])
