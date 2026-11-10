# -*- coding: utf-8 -*-
"""26년 연간 집계표 (정리안, 2026-11-10) · share/26-annual-simple.html

  네 칸: 26SS 신상 · 26SS 러닝(구품번 1/1~7/31) · 26FW 신상 · 26FW 러닝(신품번 + 구품번 8/1~, 짝 연결).
  캡슐 두 줄: 시즌 = 연간 / 26SS / 26FW · 화면 = 한눈에 보기 / 시즌·복종 / 성별 / 러닝 / 온·오프 / 전체 표.
  자료는 빌드가 만든 SS(26SS 탭 모양)·ANN(연간)·SUM_FW(26FW 집계표)·CHN(온·오프 채널) 그대로 JSON 으로 끼운다 — 계산은 여기서 하지 않는다.
  원가는 VAT 포함(×1.1) 표시. 온·오프는 허브 MARK 채널 자료가 26FW 뿐이라 26FW 부분만 나온다.
  기존 집계표(26fw-season-summary-v2)는 이 페이지가 자리 잡으면 없앨 예정(사용자, 2026-11-10).
"""
import datetime, io, json, os

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(os.path.dirname(HERE), 'share', '26-annual-simple.html')
ITEMS = ('OUTER', 'TOP', 'BOTTOM', 'ACC')
KEYS = ('cost', 'qtyIn', 'sales', 'salesAmt', 'salesCost', 'styles', 'stock', 'carryQty', 'carryCost', 'wkQty', 'plan', 'planCost', 'tag', 'salesTag', 'wkAmt', 'sku', 'inM')
TBL_KEYS = ('styles', 'sku', 'plan', 'qtyIn', 'inM', 'cost', 'tag', 'sales', 'salesTag', 'salesAmt', 'salesCost', 'wkQty', 'wkAmt', 'planCost', 'stock')


def cell(o):
    o = o or {}
    return {k: (o.get(k) or 0) for k in KEYS}


def add(a, b):
    return {k: (a.get(k) or 0) + (b.get(k) or 0) for k in KEYS}


def slim(rows, keys=('season', 'item', 'mode', 'kind', 'cat', 'code', 'gender', 'name')):
    """표 6종용: 행의 구분 키 + cur/py 의 숫자만 남긴다(ss_data 모양(cur/py 중첩)과 SUM 모양(평면) 모두)."""
    out = []
    for r in rows or []:
        d = {k: r.get(k) for k in keys if r.get(k) not in (None, '')}
        cur = r.get('cur') if isinstance(r.get('cur'), dict) else r
        d['cur'] = {k: cur.get(k) for k in TBL_KEYS if cur.get(k) is not None}
        py = r.get('py') if isinstance(r.get('py'), dict) else None
        if py: d['py'] = {k: py.get(k) for k in TBL_KEYS if py.get(k) is not None}
        out.append(d)
    return out


def build(GS, ANN, SS, SUM=None, CHN=None, verbose=True):
    ref = GS['ref']
    b1_ss, b1_fw = SS['block1'], ANN['fw']['block1']
    def new_ttl(b1, key='cur'): return cell(next(d for d in b1 if d.get('kind') == '신상품 TTL').get(key))
    def new_item(b1, item, key='cur'):
        out = {k: 0 for k in KEYS}
        for d in b1:
            if d.get('mode') == 'new' and d.get('kind') == '항목' and d.get('item') == item: out = add(out, cell(d.get(key)))
        return out
    def new_gender(gi, gnd): return cell(next((d for d in gi if d.get('gender') == gnd and d.get('item') == 'TOTAL'), {}).get('cur'))
    T = ANN['items']
    P0 = ANN.get('fwPlan0') or {}
    def pack(nss, run_ss, nfw, run_fw, pyss=None, pyfw=None, p0key='TOTAL'):
        p0 = P0.get(p0key) or {}
        d = {'ss_new': nss, 'ss_run': cell(run_ss['ssRun']), 'fw_new': nfw, 'fw_run_new': cell(run_fw['fwNew']), 'fw_run_old': cell(run_fw['fwOld']),
             'extra_run': cell(run_ss.get('extraRun')), 'fw_new_p0': cell(p0.get('new')), 'fw_run_new_p0': cell(p0.get('runNew'))}
        if pyss is not None: d['py_ss_new'] = pyss
        if pyfw is not None: d['py_fw_new'] = pyfw
        return d
    SUM = SUM or {}
    D = {'ref': ref.isoformat(), 'refS': ref.strftime('%Y.%m.%d'), 'gen': datetime.datetime.now().strftime('%Y-%m-%d %H:%M'),
         'ssEnd': SS.get('end'), 'ssPrevEnd': SS.get('prevEnd'),
         'nSs': ANN['nSsRun'], 'nSsList': ANN.get('nSsList', 0), 'nOut': ANN.get('nOldOutside', 0), 'nNew': ANN['nRunNew'],
         'nFwOld': ANN.get('nFwOld', 0), 'nWhole': (ANN.get('split') or {}).get('nWhole', 0), 'nSplit': (ANN.get('split') or {}).get('nSplit', 0), 'gf1Cut': (ANN.get('split') or {}).get('gf1Cut', []),
         'nExtra': (ANN.get('split') or {}).get('nExtra', 0), 'extraSplit': (ANN.get('split') or {}).get('extraSplit', 0), 'extraWhole': (ANN.get('split') or {}).get('extraWhole', 0),
         'ssSplit': (ANN.get('split') or {}).get('ssSplit', 0), 'ssWhole': (ANN.get('split') or {}).get('ssWhole', 0), 'extraList': (ANN.get('split') or {}).get('extraList', []),
         'nSsNew': int(new_ttl(b1_ss)['styles']), 'nFwNew': int(new_ttl(b1_fw)['styles']),
         'total': pack(new_ttl(b1_ss), T['TOTAL'], new_ttl(b1_fw), T['TOTAL'], new_ttl(b1_ss, 'py'), new_ttl(b1_fw, 'py')),
         'items': {it: pack(new_item(b1_ss, it), T[it], new_item(b1_fw, it), T[it], new_item(b1_ss, it, 'py'), new_item(b1_fw, it, 'py'), p0key=it) for it in ITEMS},
         'gender': {g: pack(new_gender(SS['genderItem'], g), ANN['runGender'][g], new_gender(ANN['fw']['genderItem'], g), ANN['runGender'][g], p0key=g) for g in ('여성', '남성')},
         'pairs': ANN.get('pairRows', []), 'unpaired': ANN.get('unpairedNew', []), 'split': ANN.get('split', {}), 'residMax': ANN.get('residMax', 0),
         'tables': {'ss': {k: slim(SS.get(k)) for k in ('block1', 'newCat', 'seasonCat', 'genderItem', 'genderCat')},
                    'fw': {k: slim(ANN['fw'].get(k)) for k in ('block1', 'newCat', 'seasonCat', 'genderItem', 'genderCat')},
                    'fwRun': {'runItem': slim(SUM.get('runItem')), 'runCat': slim(SUM.get('runCat'))}},
         'chn': CHN or {}}
    page = TEMPLATE.replace('/*__D__*/{}', json.dumps(D, ensure_ascii=False))
    io.open(OUT, 'w', encoding='utf-8').write(page)
    t = D['total']
    if verbose:
        f = lambda v: format(int(round(v)), ',')
        print('연간 집계표(정리안): 원가(VAT+) 26SS %s+%s · 26FW %s+%s+%s | 판매 %s+%s+%s+%s = %s | 온·오프 %d행 → %s' % (
            f(t['ss_new']['cost'] * 1.1), f(t['ss_run']['cost'] * 1.1), f(t['fw_new']['cost'] * 1.1), f(t['fw_run_new']['cost'] * 1.1), f(t['fw_run_old']['cost'] * 1.1),
            f(t['ss_new']['sales']), f(t['ss_run']['sales']), f(t['fw_new']['sales']), f(t['fw_run_new']['sales'] + t['fw_run_old']['sales']),
            f(sum(t[k]['sales'] for k in ('ss_new', 'ss_run', 'fw_new', 'fw_run_new', 'fw_run_old', 'extra_run'))), len((CHN or {}).get('rows') or []), os.path.basename(OUT)))
    return D


TEMPLATE = r"""<!doctype html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>26년 연간 집계표</title>
<style>
:root{--bg:#f6f7fb;--card:#fff;--ink:#1a1d23;--sub:#475467;--muted:#6b7280;--line:#e5e7eb;--blue:#2563eb;--ss:#2563eb;--fw:#c2410c;--up:#dc2626;--down:#2563eb}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font-family:"Pretendard","Malgun Gothic","Apple SD Gothic Neo",system-ui,sans-serif;font-size:14px;line-height:1.5}
.wrap{max-width:1180px;margin:0 auto;padding:18px 16px 48px}
.hd1{display:flex;flex-wrap:wrap;align-items:baseline;gap:8px 14px}h1{font-size:20px;margin:0}.hd1 .rng{font-size:12.5px;color:var(--muted)}
.bar{display:flex;flex-wrap:wrap;align-items:center;gap:8px 18px;margin:12px 0 6px}
.flabel{flex:none;font-size:11px;color:var(--muted);font-weight:700;margin-right:2px}
.seg{display:inline-flex;align-items:center;gap:4px;flex-wrap:wrap}.seg .opt{border:1px solid var(--line);background:#fff;color:var(--sub);padding:6px 13px;font-size:12.5px;font-weight:700;cursor:pointer;border-radius:999px;transition:.12s;white-space:nowrap}
.seg .opt:hover{border-color:#c9d2e0;color:var(--ink)}.seg .opt.on{background:var(--blue);border-color:var(--blue);color:#fff}.seg.period .opt{font-size:13px;padding:7px 17px}
.sec{margin-top:18px}h2{font-size:16px;margin:0 0 8px}h2 small{color:var(--muted);font-weight:500;font-size:12.5px;margin-left:6px}h3{font-size:14px;margin:0 0 6px}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:14px 16px;box-shadow:0 1px 2px rgba(0,0,0,.03)}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:10px;margin:12px 0 4px}.kpi{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:12px 14px}
.kpi .l{font-size:12px;color:var(--muted)}.kpi .v{font-size:24px;font-weight:800;letter-spacing:-.02em}.kpi .s{font-size:12px;color:var(--muted)}.kpi.ss .v{color:var(--ss)}.kpi.fw .v{color:var(--fw)}
table{width:100%;border-collapse:collapse}th,td{padding:8px 10px;border-bottom:1px solid var(--line);text-align:left;vertical-align:top}thead th{font-size:12px;color:var(--muted);font-weight:600;background:#fafafa;white-space:nowrap}
td.n,th.n{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}td.p{color:var(--muted);font-size:12px;white-space:nowrap}
tr.l1 th{font-weight:600}th small,td small{display:block;color:var(--muted);font-weight:400;font-size:11.5px}tr.l2 th,tr.l2 td{color:#4b5563;font-size:13px;background:#fcfcfd}tr.extra th,tr.extra td{background:#fdf6ec}
tr.sub th,tr.sub td{background:#f3f4f6;font-weight:600}tr.tot th,tr.tot td{background:#111827;color:#fff;font-weight:700;font-size:15px;border-bottom:0}tr.tot td.p,tr.tot small{color:#d1d5db}
tr.exp th{cursor:pointer}tr.exp th .tg{display:inline-block;width:14px;color:var(--muted);font-size:11px}tr.det td{background:#f8fafc;color:#374151;font-size:12.5px;padding:6px 10px 10px 34px}
.det .dg{display:flex;flex-wrap:wrap;gap:6px 22px}.det .dg b{font-weight:600}.det .dg span{display:inline-block}
.up{color:var(--up);font-weight:700}.down{color:var(--down);font-weight:700}.dim{color:#9ca3af}.note{color:var(--muted);font-size:12.5px;margin-top:8px}.warn{background:#fff7ed;border-color:#fed7aa}
.chips{display:flex;flex-wrap:wrap;gap:6px;margin:0 0 10px;align-items:center}.chip{border:1px solid var(--line);background:#fff;border-radius:999px;padding:4px 11px;font-size:12px;font-weight:600;cursor:pointer;color:var(--sub)}.chip.on{background:#111827;color:#fff;border-color:#111827}
.tw{overflow-x:auto}.rule{background:#fff;border:1px solid var(--line);border-radius:12px;padding:12px 16px;font-size:13px}.rule ul{margin:6px 0 0 18px;padding:0}.rule li{margin:3px 0}
.ok{color:#047857;font-weight:600}.bad{color:#b91c1c;font-weight:600}
details.grp{border:1px solid var(--line);border-radius:10px;margin-top:10px;background:#fff}details.grp>summary{cursor:pointer;padding:10px 14px;font-weight:700;list-style:none;display:flex;flex-wrap:wrap;gap:6px 16px;align-items:baseline}
details.grp>summary::-webkit-details-marker{display:none}details.grp>summary::before{content:'▸';color:var(--muted);margin-right:4px}details[open].grp>summary::before{content:'▾'}details.grp>summary small{display:inline;font-weight:500;color:var(--muted)}details.grp .tw{padding:0 10px 10px}
.legend{display:flex;gap:14px;flex-wrap:wrap;font-size:11px;color:var(--sub);margin:0 0 8px}.legend i{display:inline-block;width:12px;height:8px;border-radius:2px;margin-right:5px}
.st-row{display:grid;grid-template-columns:120px 1fr;gap:4px 12px;align-items:center;margin:6px 0}.st-name{font-size:12.5px;font-weight:600}.st-bar{display:flex;height:18px;border-radius:5px;overflow:hidden;background:#eef0f4}.st-bar span{display:flex;align-items:center;justify-content:center;color:#fff;font-size:10.5px;font-weight:700}.st-cap{grid-column:2;font-size:11.5px;color:var(--muted)}
.ch-on{background:#0e7490}.ch-off{background:#6d4aa5}.ch-etc{background:#6b7280}
.headline{background:#eef4ff;border:1px solid #d6e2ff;border-radius:12px;padding:12px 16px;font-size:14px}.headline .story{margin-top:4px;color:var(--sub);font-size:13px}
@media(max-width:640px){th,td{padding:6px 6px}.kpi .v{font-size:20px}.seg.period .opt{padding:6px 12px}tr.det td{padding-left:12px}}
</style></head><body><div class="wrap">
<div class="hd1"><h1>26년 연간 집계표</h1><span class="rng" id="rng"></span></div>
<div class="bar">
  <div class="seg period" id="segSeason" role="group" aria-label="시즌"><span class="flabel">시즌</span>
    <button class="opt" data-season="all">연간 (SS+FW)</button><button class="opt" data-season="ss">26SS</button><button class="opt" data-season="fw">26FW</button></div>
  <div class="seg" id="segView" role="group" aria-label="화면"><span class="flabel">화면</span>
    <button class="opt" data-view="overview">한눈에 보기</button><button class="opt" data-view="item">시즌 · 복종</button><button class="opt" data-view="gender">성별</button><button class="opt" data-view="run">러닝</button><button class="opt" data-view="onoff">온 · 오프</button><button class="opt" data-view="tables">전체 표</button></div>
</div>
<div id="view"></div>
</div>
<script>
const D = /*__D__*/{};
const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const fmtN = v => (v == null || isNaN(v)) ? '-' : Math.round(v).toLocaleString('ko-KR');
const fmtW = v => { if(v == null || isNaN(v)) return '-'; const a = Math.abs(v), s = v < 0 ? '−' : ''; if(a >= 1e8) return s + (a / 1e8).toFixed(1) + '억'; if(a >= 1e4) return s + Math.round(a / 1e4).toLocaleString('ko-KR') + '만'; return s + Math.round(a).toLocaleString('ko-KR'); };
const fmtP = v => (v == null || isNaN(v) || !isFinite(v)) ? '-' : (v * 100).toFixed(1) + '%';
const Z = () => ({cost:0, qtyIn:0, sales:0, salesAmt:0, salesCost:0, styles:0, stock:0, carryQty:0, carryCost:0, wkQty:0, plan:0, planCost:0, tag:0, salesTag:0, wkAmt:0, sku:0, inM:0});
const add = (...xs) => xs.reduce((a, b) => { Object.keys(a).forEach(k => { a[k] += (b && b[k]) || 0; }); return a; }, Z());
const sub = (a, b) => { const o = Z(); Object.keys(o).forEach(k => { o[k] = ((a && a[k]) || 0) - ((b && b[k]) || 0); }); ['sales', 'salesAmt', 'salesCost', 'salesTag', 'wkQty', 'wkAmt', 'stock', 'carryQty', 'carryCost'].forEach(k => { o[k] = 0; }); return o; };   // 리오더분 = 리오더 포함 − 리오더 제외 (판매·재고는 차수로 못 나눠 0)
const reoOf = (P, k) => sub(P[k], P[k + '_p0']);
const mainOf = (P, k) => { const all = P[k], p0 = P[k + '_p0'] || all; const o = add(p0); ['sales', 'salesAmt', 'salesCost', 'salesTag', 'wkQty', 'wkAmt', 'stock'].forEach(kk => { o[kk] = all[kk]; }); return o; };   // 본 숫자: 기획·입고·금액은 리오더 제외, 판매·재고는 전체
const MET = {costV: {l:'원가 총액', f:o => fmtW(o.cost * 1.1), v:o => o.cost}, qtyIn: {l:'입고(투입)량', f:o => fmtN(o.qtyIn), v:o => o.qtyIn}, sales: {l:'누계판매량', f:o => fmtN(o.sales), v:o => o.sales}, salesAmt: {l:'실판매금액', f:o => fmtW(o.salesAmt), v:o => o.salesAmt}};
const SNAME = {ss:'26SS', fw:'26FW', all:'연간 (SS+FW)'}, SLABEL = {ss:'26SS (신상 + 러닝)', fw:'26FW (신상 + 러닝)', all:'연간 (26SS + 26FW · 신상 + 러닝)'};
const PYLAB = {ss:'25SS 동시점', fw:'25FW 동시점', all:'25SS+25FW 동시점'};
let SEASON = 'all', VIEW = 'overview', AMET = 'costV', OSCOPE = '전체';
const IN_FRAME = (() => { try { return window.parent !== window; } catch(e){ return true; } })();
const OPEN = new Set();                                             // 펼쳐 둔 행(전년 비교)·그룹의 키 — 다시 그려도 유지
const dPct = (cur, py) => (py == null || !py) ? null : (cur - py) / py;
const dHtml = d => d == null ? '<span class="dim">-</span>' : `<span class="${d >= 0 ? 'up' : 'down'}">${d >= 0 ? '▲' : '▼'} ${Math.abs(d * 100).toFixed(1)}%</span>`;

/* 시즌별 행 — 전년(py)은 신상 행에만 있다(러닝은 비교 자료 없음) */
function rowsFor(P, season){
  const sn = P.ss_new, sr = P.ss_run, fn = P.fw_new, frn = P.fw_run_new, fro = P.fw_run_old, fr = add(frn, fro), xr = P.extra_run || Z();
  const ssT = add(sn, sr), fwT = add(fn, fr), tot = add(ssT, fwT, xr);
  const R = [];
  if(season !== 'fw'){
    R.push({k:'ss_new', name:'26SS 신상', o:sn, cls:'l1', sub:`봄·여름 ${D.nSsNew}품번 · 리오더 포함${(D.gf1Cut || []).length ? ` · GF1 ${D.gf1Cut.length}개는 7/31 까지(8/1~ 는 26FW 러닝 이월)` : ''}`, period:`~ ${D.refS} 누계`, py:P.py_ss_new, pyl:'25SS 동시점'});
    R.push({k:'ss_run', name:'26SS 러닝', o:sr, cls:'l1', sub:`26SS 러닝 목록 ${D.nSs}개(대시보드 26SS 종합과 같음) · FW 과거품번 ${D.ssSplit}개는 1/1~7/31, 그 밖 ${D.ssWhole}개는 1/1~`, period:'FW 과거품번 2026.01.01 ~ 07.31 · 그 밖 ~ 기준일', nopy:'러닝은 전년 비교 자료가 없습니다'});
    R.push({k:'ss_sub', name:'26SS 소계', o:ssT, cls:'sub', sub:'신상 + 러닝', py:P.py_ss_new, pyl:'25SS 동시점 (신상만)', pyNote:'전년은 신상만'});
  }
  if(season === 'ss') R.push({k:'extra', name:'과년차 러닝 상품', o:xr, cls:'l1 extra', sub:`구품번 ${D.nExtra}개 · 26SS 러닝 목록 밖(24SS~25FW 러닝) · 26SS 소계에 넣지 않음, 연간에만 포함`, period:`FW 과거품번 ${D.extraSplit}개 1/1~7/31 · 그 밖 ${D.extraWhole}개 ~ 기준일`, nopy:'전년 비교 자료가 없습니다'});
  if(season !== 'ss'){
    const fnMain = mainOf(P, 'fw_new'), frnMain = mainOf(P, 'fw_run_new'), reoN = reoOf(P, 'fw_new'), reoR = reoOf(P, 'fw_run_new'), frMain = add(frnMain, fro);
    R.push({k:'fw_new', name:'26FW 신상', o:fnMain, cls:'l1', sub:`가을·겨울 ${D.nFwNew}품번 · 기획·입고·원가는 리오더 제외(대시보드 26FW 집계표와 같음)`, period:`~ ${D.refS} 누계`, py:P.py_fw_new, pyl:'25FW 동시점'});
    R.push({k:'fw_new_reo', name:'　└ 리오더', o:reoN, cls:'l2', reo:true, noexp:true});
    R.push({k:'fw_run', name:'26FW 러닝', o:frMain, cls:'l1', sub:`신품번 ${D.nNew}개(리오더 제외) + FW 과거품번 ${D.nFwOld}개 8/1~ (짝 연결 합산)`, period:'', carry:true, nopy:'러닝은 전년 비교 자료가 없습니다', den: frnMain.qtyIn + fro.carryQty});
    R.push({k:'fw_run_new', name:'　└ 신품번', o:frnMain, cls:'l2', sub:'', period:`~ ${D.refS} 누계`, noexp:true});
    R.push({k:'fw_run_reo', name:'　└ 신품번 리오더', o:reoR, cls:'l2', reo:true, noexp:true});
    R.push({k:'fw_run_old', name:'　└ FW 과거품번 8/1~ 생산', o:fro, cls:'l2', sub:'대시보드에는 없는 항목', period:`2026.08.01 ~ ${D.refS} · 매주 갱신`, carry:true, noRest:true, noexp:true});
    R.push({k:'fw_sub', name:'26FW 소계', o:fwT, cls:'sub', sub:'신상 + 러닝 + 리오더', carry:true, py:P.py_fw_new, pyl:'25FW 동시점 (신상만)', pyNote:'전년은 신상만', den: fnMain.qtyIn + frnMain.qtyIn + fro.carryQty});
  }
  if(season === 'all') R.push({k:'extra', name:'과년차 러닝 상품', o:xr, cls:'l1 extra', sub:`구품번 ${D.nExtra}개 · 26SS 러닝 목록 밖(24SS~25FW 러닝) · 어느 시즌에도 안 넣고 연간에만 더함`, period:`FW 과거품번 ${D.extraSplit}개 1/1~7/31 · 그 밖 ${D.extraWhole}개 ~ 기준일`, nopy:'전년 비교 자료가 없습니다'});
  if(season === 'all') R.push({k:'tot', name:'연간 합계', o:tot, cls:'tot', sub:'신상품 + 러닝 + 과년차 러닝', den: tot.qtyIn - reoOf(P, 'fw_new').qtyIn - reoOf(P, 'fw_run_new').qtyIn + fro.carryQty, carry:true, py:(P.py_ss_new && P.py_fw_new) ? add(P.py_ss_new, P.py_fw_new) : null, pyl:'25SS+25FW 동시점 (신상만)', pyNote:'전년은 신상만'});
  return {R, tot: season === 'ss' ? ssT : season === 'fw' ? fwT : tot, ssT, fwT, xr};
}
/* 전년 비교 상세 행(접었다 펴기) — 원가·입고·판매·실판매금액·장당 원가 */
function detRow(r, span, which){
  if(r.noexp) return '';
  if(!r.py) return `<tr class="det" data-det="${esc(r.k)}"${OPEN.has(which + r.k) ? '' : ' hidden'}><td colspan="${span}"><span class="dim">${esc(r.nopy || '전년 자료 없음')}</span></td></tr>`;
  const o = r.o, p = r.py, items = [['원가 총액', fmtW(o.cost * 1.1), fmtW(p.cost * 1.1), dPct(o.cost, p.cost)], ['입고량', fmtN(o.qtyIn), fmtN(p.qtyIn), dPct(o.qtyIn, p.qtyIn)], ['누계판매량', fmtN(o.sales), fmtN(p.sales), dPct(o.sales, p.sales)],
    ['실판매금액', fmtW(o.salesAmt), fmtW(p.salesAmt), dPct(o.salesAmt, p.salesAmt)], ['장당 원가', o.qtyIn ? fmtN(o.cost * 1.1 / o.qtyIn) + '원' : '-', p.qtyIn ? fmtN(p.cost * 1.1 / p.qtyIn) + '원' : '-', (o.qtyIn && p.qtyIn) ? dPct(o.cost / o.qtyIn, p.cost / p.qtyIn) : null],
    ['판매율', fmtP(o.qtyIn ? o.sales / o.qtyIn : null), fmtP(p.qtyIn ? p.sales / p.qtyIn : null), null]];
  return `<tr class="det" data-det="${esc(r.k)}"${OPEN.has(which + r.k) ? '' : ' hidden'}><td colspan="${span}"><div class="dg"><b>${esc(r.pyl || '전년 동시점')}</b>${items.map(([l, a, b, d]) => `<span>${l} <b>${a}</b> <span class="dim">vs ${b}</span> ${d == null ? '' : dHtml(d)}</span>`).join('')}${r.pyNote ? `<span class="dim">· ${esc(r.pyNote)}</span>` : ''}</div></td></tr>`;
}
const expTh = (r, which) => r.noexp ? `<th>${esc(r.name)}${r.sub ? `<small>${esc(r.sub)}</small>` : ''}</th>` : `<th data-exp="${esc(which + r.k)}" title="전년 동시점 비교 펼치기/접기"><span class="tg">${OPEN.has(which + r.k) ? '▾' : '▸'}</span>${esc(r.name)}${r.sub ? `<small>${esc(r.sub)}</small>` : ''}</th>`;

function kpis(){
  const {R, tot, ssT, fwT, xr} = rowsFor(D.total, SEASON);
  const k = (cls, l, v, s) => `<div class="kpi ${cls}"><div class="l">${l}</div><div class="v">${v}</div><div class="s">${s}</div></div>`;
  const xC = xr.cost ? ` + 과년차 ${fmtW(xr.cost * 1.1)}` : '', xS = xr.sales ? ` + 과년차 ${fmtN(xr.sales)}` : '', xA = xr.salesAmt ? ` + 과년차 ${fmtW(xr.salesAmt)}` : '';
  const by = {}; R.forEach(r => { by[r.k] = r; });
  const nw = SEASON === 'fw' ? by.fw_new : by.ss_new, rn = SEASON === 'fw' ? by.fw_run : by.ss_run, reoT = SEASON === 'fw' ? add(reoOf(D.total, 'fw_new'), reoOf(D.total, 'fw_run_new')) : null;
  const parts = SEASON === 'all' ? `26SS ${fmtW(ssT.cost * 1.1)} + 26FW ${fmtW(fwT.cost * 1.1)}${xC}` : `신상 ${fmtW(nw.o.cost * 1.1)} + 러닝 ${fmtW(rn.o.cost * 1.1)}${reoT ? ` + 리오더 ${fmtW(reoT.cost * 1.1)}` : ''}${SEASON === 'ss' ? ' (과년차 러닝 제외)' : ''}`;
  const partsS = SEASON === 'all' ? `26SS ${fmtN(ssT.sales)} + 26FW ${fmtN(fwT.sales)}${xS}` : `신상 ${fmtN(nw.o.sales)} + 러닝 ${fmtN(rn.o.sales)}`;
  const partsA = SEASON === 'all' ? `26SS ${fmtW(ssT.salesAmt)} + 26FW ${fmtW(fwT.salesAmt)}${xA}` : `신상 ${fmtW(nw.o.salesAmt)} + 러닝 ${fmtW(rn.o.salesAmt)}`;
  const den = SEASON === 'ss' ? ssT.qtyIn : SEASON === 'fw' ? by.fw_sub.den : ssT.qtyIn + by.fw_sub.den + xr.qtyIn;      // 판매율 분모 = 표의 소계 행과 같은 기준(26FW 는 리오더 제외 입고 + 7/31 이월)
  return `<div class="kpis">${k(SEASON, SLABEL[SEASON] + ' 원가 총액', fmtW(tot.cost * 1.1), parts)}${k(SEASON, '누계판매량', fmtN(tot.sales) + '<small style="display:inline;font-size:13px"> 장</small>', partsS)}${k(SEASON, '실판매금액', fmtW(tot.salesAmt), partsA)}${k(SEASON, '판매율 (판매 ÷ 투입)', fmtP(tot.sales / den), SEASON === 'ss' ? '12/31 재고 + 생산 기준' : '리오더 제외 입고 + 구품번 7/31 이월 재고 기준(대시보드와 같음)')}</div>`;
}
function costTable(P, season){
  const {R, tot} = rowsFor(P, season), span = 7;
  const rows = R.map(r => { const cv = r.o.cost * 1.1, q = r.o.qtyIn;
    if(r.reo) return `<tr class="${r.cls}"><th>${esc(r.name)}<small>리오더 기획 ${fmtN(r.o.plan)}장 · 입고 ${fmtN(q)}장</small></th><td class="n">${fmtW(cv)}</td><td class="n">${fmtN(q)}</td><td class="n">${q ? fmtN(cv / q) + '원' : '-'}</td><td class="n">-</td><td class="n">-</td><td class="n">${tot.cost ? fmtP(r.o.cost / tot.cost) : '-'}</td></tr>`;
    return `<tr class="${r.cls} exp">${expTh(r, 'c')}<td class="n">${fmtW(cv)}</td><td class="n">${fmtN(q)}</td><td class="n">${q ? fmtN(cv / q) + '원' : '-'}</td><td class="n">${fmtW(r.o.salesCost * 1.1)}</td><td class="n">${r.noRest ? '<span class="dim">이월분 포함 · 산출 안 함</span>' : fmtW((r.o.cost - r.o.salesCost) * 1.1)}</td><td class="n">${tot.cost ? fmtP(r.o.cost / tot.cost) : '-'}</td></tr>` + detRow(r, span, 'c'); }).join('');
  return `<div class="tw"><table><thead><tr><th>구분 <small style="display:inline">▸ 누르면 전년 동시점 비교</small></th><th class="n">원가 총액</th><th class="n">입고(투입)량</th><th class="n">장당 원가</th><th class="n">판매 원가</th><th class="n">남은 재고 원가<small>추정</small></th><th class="n">비중</th></tr></thead><tbody>${rows}</tbody></table></div>`;
}
function saleTable(P, season){
  const {R, tot} = rowsFor(P, season), span = 7;
  const rows = R.map(r => { if(r.reo) return `<tr class="${r.cls}"><th>${esc(r.name)}</th><td class="n dim">차수로 못 나눔</td><td class="n">-</td><td class="n">-</td><td class="n">-</td><td class="n">-</td><td class="p"></td></tr>`;
    const den = r.den != null ? r.den : r.o.qtyIn + (r.carry ? r.o.carryQty : 0), rate = den ? fmtP(r.o.sales / den) + (r.carry ? '<small style="display:inline">※</small>' : '') : '-';
    return `<tr class="${r.cls} exp">${expTh(r, 's')}<td class="n">${fmtN(r.o.sales)}</td><td class="n">${fmtW(r.o.salesAmt)}</td><td class="n">${rate}</td><td class="n">${tot.sales ? fmtP(r.o.sales / tot.sales) : '-'}</td><td class="n">${r.o.wkQty ? fmtN(r.o.wkQty) : '-'}</td><td class="p">${esc(r.period || '')}</td></tr>` + detRow(r, span, 's'); }).join('');
  return `<div class="tw"><table><thead><tr><th>구분 <small style="display:inline">▸ 누르면 전년 동시점 비교</small></th><th class="n">누계판매량</th><th class="n">실판매금액</th><th class="n">판매율<small>판매 ÷ 투입</small></th><th class="n">비중</th><th class="n">이번 주</th><th>기간</th></tr></thead><tbody>${rows}</tbody></table></div>`;
}
function viewOverview(){
  const t = D.total;
  return kpis() + `
  <div class="sec"><h2>1. 원가 총액 <small>입고(투입) 기준 · VAT 포함 · 행을 누르면 전년 동시점 비교가 펼쳐집니다</small></h2><div class="card">${costTable(t, SEASON)}
    <div class="note">26FW 신상·신품번의 기획·입고·원가·TAG 는 리오더 제외(대시보드 26FW 집계표 본 숫자와 같음), 리오더는 ‘└ 리오더’ 행에 따로 · 26SS 러닝 투입 = 12/31 시점재고 + 생산(FW 과거품번 ${D.nSplit}개는 1/1~7/31, 그 밖 ${D.nWhole}개는 1/1~) · 26FW 러닝 구품번 투입 = FW 과거품번의 8/1~ 생산만(7/31 이월 재고 ${fmtN(t.fw_run_old.carryQty)}장 · 원가 ${fmtW(t.fw_run_old.carryCost * 1.1)}은 26SS 투입에 이미 들어 있어 더하지 않음) · 신상·신품번 = 생산 전체(리오더 포함) · 전년 동시점은 신상만 비교(러닝은 전년 자료 없음).</div></div></div>
  <div class="sec"><h2>2. 누계판매 <small>수량 · 실판매금액</small></h2><div class="card">${saleTable(t, SEASON)}
    <div class="note">26SS 러닝 판매 = FW 과거품번은 7/31 시점 누계 − 12/31 시점 누계, 그 밖 구품번은 기준일 누계 − 12/31 · 26FW 러닝 구품번 판매 = FW 과거품번 ${D.nFwOld}개의 기준일 누계 − 7/31 시점 누계(매주 월요일 갱신 · 대시보드 26FW 집계표의 ‘시점재고 합산’과 같은 목록) · 신품번은 신상과 같은 누계 · ※ 판매율은 구품번의 7/31 이월 재고를 공급량에 더해 나눈 값.</div></div></div>
  ${rule()}`;
}
/* 복종 · 성별: 한 지표를 골라 칸 × 행 매트릭스 (+ 행마다 전년 비교 펼치기) */
function matrixHtml(groups, rowLabel, which){
  const cols = SEASON === 'ss' ? [['26SS 신상', 'ss_new'], ['26SS 러닝', 'ss_run'], ['과년차 러닝', 'extra_run']] : SEASON === 'fw' ? [['26FW 신상', 'fw_new_m'], ['26FW 러닝 신품번', 'fw_run_new_m'], ['26FW 과거품번 8/1~', 'fw_run_old'], ['26FW 리오더', 'fw_reo']] : [['26SS 신상', 'ss_new'], ['26SS 러닝', 'ss_run'], ['26FW 신상', 'fw_new_m'], ['26FW 러닝', 'fw_run_m'], ['26FW 리오더', 'fw_reo'], ['과년차 러닝', 'extra_run']];
  const get = (P, k) => k === 'fw_run' ? add(P.fw_run_new, P.fw_run_old) : k === 'fw_new_m' ? mainOf(P, 'fw_new') : k === 'fw_run_new_m' ? mainOf(P, 'fw_run_new') : k === 'fw_run_m' ? add(mainOf(P, 'fw_run_new'), P.fw_run_old) : k === 'fw_reo' ? add(reoOf(P, 'fw_new'), reoOf(P, 'fw_run_new')) : P[k];
  const m = MET[AMET];
  const cells = groups.map(([label, P]) => ({label, P, cs: cols.map(([_, k]) => get(P, k))}));
  cells.forEach(c => { c.rowT = add(...c.cs); });
  const colTot = cols.map((_, i) => add(...cells.map(c => c.cs[i]))), totRow = add(...cells.map(c => c.rowT));
  const pyOf = P => { const a = SEASON !== 'fw' ? P.py_ss_new : null, b = SEASON !== 'ss' ? P.py_fw_new : null; if(!a && !b) return null; return add(a || {}, b || {}); };
  const span = cols.length + 3;
  const body = cells.map(c => { const py = pyOf(c.P), curNew = add(SEASON !== 'fw' ? c.P.ss_new : {}, SEASON !== 'ss' ? c.P.fw_new : {});
    const r = {k: c.label, name: c.label, o: curNew, py, pyl: PYLAB[SEASON] + ' (신상만)', pyNote: '전년은 신상만 · 러닝 제외', nopy: '전년 자료 없음'};
    return `<tr class="l1 exp">${expTh(r, which)}${c.cs.map(x => `<td class="n">${m.f(x)}</td>`).join('')}<td class="n"><b>${m.f(c.rowT)}</b></td><td class="n">${fmtP(m.v(c.rowT) / (m.v(totRow) || 1))}</td></tr>` + detRow(r, span, which); }).join('');
  const chips = Object.keys(MET).map(k => `<button class="chip${AMET === k ? ' on' : ''}" data-amet="${k}">${MET[k].l}</button>`).join('');
  return `<div class="chips"><span class="flabel">지표</span>${chips}</div><div class="tw"><table><thead><tr><th>${esc(rowLabel)} <small style="display:inline">▸ 전년 비교</small></th>${cols.map(c => `<th class="n">${esc(c[0])}</th>`).join('')}<th class="n">합계</th><th class="n">비중</th></tr></thead><tbody>${body}
    <tr class="tot"><th>합계</th>${colTot.map(c => `<td class="n">${m.f(c)}</td>`).join('')}<td class="n">${m.f(totRow)}</td><td class="n">100%</td></tr></tbody></table></div>`;
}
function viewItem(){
  const groups = ['OUTER', 'TOP', 'BOTTOM', 'ACC'].map(it => [it, D.items[it]]);
  return kpis() + `<div class="sec"><h2>${esc(SNAME[SEASON])} · 복종 <small>${MET[AMET].l} · 신상 / 러닝 × OUTER · TOP · BOTTOM · ACC</small></h2><div class="card">${matrixHtml(groups, '복종', 'i')}
    <div class="note">26SS 러닝 = 26SS 러닝 목록 31개(FW 과거품번은 1/1~7/31, 그 밖은 1/1~) · 과년차 러닝 = 목록 밖 구품번 28개(26SS 소계 밖, 연간에만) · 26FW 러닝 = 신품번 전체 + FW 과거품번 8/1~ · 원가는 VAT 포함 · 복종은 ERP 품목(KNIT·INNER → TOP) 기준 · 전년 비교는 신상만(복종별 전년 자료는 신상에만 있음).</div></div></div>${rule()}`;
}
function viewGender(){
  const groups = [['여성', D.gender['여성']], ['유니맨 (남성 + 공용)', D.gender['남성']]];
  return kpis() + `<div class="sec"><h2>${esc(SNAME[SEASON])} · 성별 <small>${MET[AMET].l} · 여성 / 유니맨</small></h2><div class="card">${matrixHtml(groups, '성별', 'g')}
    <div class="note">성별은 ERP 성별 열(여성 / 그 외 = 유니맨). 전년 성별 자료는 기준이 확인되지 않아 비교하지 않습니다(펼쳐도 ‘전년 자료 없음’).</div></div></div>${rule()}`;
}
/* 러닝 — 신품번 / 구품번 그룹(접었다 펴기) */
function viewRun(){
  const P = D.total, allPairs = (D.pairs || []).slice().sort((a, b) => ((b.news || []).length ? 1 : 0) - ((a.news || []).length ? 1 : 0) || (b.fwSales - a.fwSales));
  const EX = new Set(D.extraList || []), pairs = allPairs.filter(r => !EX.has(r.code)), extras = allPairs.filter(r => EX.has(r.code)), xr = P.extra_run || Z();
  const nsum = (r, k) => (r.news || []).reduce((z, n) => z + (n[k] || 0), 0);
  const showSS = SEASON !== 'fw', showFW = SEASON !== 'ss';
  const sr = P.ss_run, fo = P.fw_run_old, fn = P.fw_run_new;
  // 신품번 목록(37): 짝 구품번 + 신품번 자신의 누계
  const newList = []; (D.pairs || []).forEach(r => (r.news || []).forEach(n => { if(!newList.find(x => x.code === n.code)) newList.push(Object.assign({old: r.code, oldName: r.name, oldFw: r.fwSales}, n)); }));
  (D.unpaired || []).forEach(n => newList.push(Object.assign({old: null}, n)));
  newList.sort((a, b) => b.sales - a.sales);
  const newT = newList.reduce((z, n) => ({sales: z.sales + n.sales, stock: z.stock + n.stock, wk: z.wk + (n.wkQty || 0), cost: z.cost + (n.cost || 0), qtyIn: z.qtyIn + (n.qtyIn || 0)}), {sales:0, stock:0, wk:0, cost:0, qtyIn:0});
  const grpOpen = key => OPEN.has('grp:' + key) ? ' open' : '';
  const newGrp = `<details class="grp" data-grp="new"${grpOpen('new')}><summary>신품번 ${D.nNew}개 <small>26FW · GF7·GF8 · 신상과 같은 방식(생산 전체 · 누계)</small><span style="margin-left:auto">판매 <b>${fmtN(fn.sales)}장</b> · 투입 ${fmtN(fn.qtyIn)}장 · 원가 ${fmtW(fn.cost * 1.1)} · 현재고 ${fmtN(fn.stock)}장</span></summary>
    <div class="tw"><table><thead><tr><th>신품번</th><th>복종</th><th>짝 구품번</th><th class="n">입고</th><th class="n">누계판매</th><th class="n">판매율</th><th class="n">현재고</th><th class="n">이번 주</th><th class="n">원가(VAT+)</th></tr></thead><tbody>
    ${newList.map(n => `<tr><th>${esc(n.code)}<small>${esc(n.name)}</small></th><td>${esc(n.item || '')}</td><td>${n.old ? `${esc(n.old)}<small>${esc(n.oldName || '')} · 8/1~ ${fmtN(n.oldFw)}장</small>` : '<span class="dim">없음</span>'}</td><td class="n">${fmtN(n.qtyIn)}</td><td class="n">${fmtN(n.sales)}</td><td class="n">${fmtP(n.qtyIn ? n.sales / n.qtyIn : null)}</td><td class="n">${fmtN(n.stock)}</td><td class="n">${fmtN(n.wkQty)}</td><td class="n">${fmtW((n.cost || 0) * 1.1)}</td></tr>`).join('')}
    <tr class="tot"><th>합계</th><td></td><td></td><td class="n">${fmtN(newT.qtyIn)}</td><td class="n">${fmtN(newT.sales)}</td><td class="n">${fmtP(newT.qtyIn ? newT.sales / newT.qtyIn : null)}</td><td class="n">${fmtN(newT.stock)}</td><td class="n">${fmtN(newT.wk)}</td><td class="n">${fmtW(newT.cost * 1.1)}</td></tr></tbody></table></div></details>`;
  const oldHead = `<tr><th>구품번</th><th>복종</th>${showSS ? '<th class="n">1/1~7/31 판매<small>26SS 몫</small></th><th class="n">1/1~7/31 투입<small>12/31 재고 + 생산</small></th>' : ''}${showFW ? '<th class="n">8/1~ 판매<small>26FW 몫</small></th><th>짝 신품번</th><th class="n">신품번 누계</th><th class="n">26FW 러닝 합<small>구 8/1~ + 신</small></th>' : ''}<th class="n">${showFW ? '현재고' : '7/31 재고'}</th>${showFW ? '<th class="n">이번 주</th>' : ''}</tr>`;
  const oldRow = r => `<tr><th>${esc(r.code)}<small>${esc(r.name)}${r.split ? (r.wasNew ? ' · FW 과거품번 · 7/31 까지는 26SS 신상' : ' · FW 과거품번') : ' · 나누지 않음(1/1~ 전부)'}</small></th><td>${esc(r.item)}</td>${showSS ? `<td class="n">${fmtN(r.ssSales)}</td><td class="n">${fmtN(r.ssIn)}</td>` : ''}${showFW ? `<td class="n">${fmtN(r.fwSales)}</td><td>${(r.news || []).length ? r.news.map(n => `${esc(n.code)}<small>${esc(n.name)}</small>`).join('') : '<span class="dim">짝 없음</span>'}</td><td class="n">${(r.news || []).length ? fmtN(nsum(r, 'sales')) : '-'}</td><td class="n"><b>${fmtN(r.fwSales + nsum(r, 'sales'))}</b></td>` : ''}<td class="n">${showFW ? fmtN(r.stock) : fmtN(r.stock731)}</td>${showFW ? `<td class="n">${fmtN(r.wkQty)}</td>` : ''}</tr>`;
  const oldSum = showSS && !showFW ? `판매 <b>${fmtN(sr.sales)}장</b>(1/1~7/31) · 투입 ${fmtN(sr.qtyIn)}장 · 원가 ${fmtW(sr.cost * 1.1)} · 7/31 재고 ${fmtN(sr.stock)}장` : showFW && !showSS ? `8/1~ 판매 <b>${fmtN(fo.sales)}장</b> · 8/1~ 생산 ${fmtN(fo.qtyIn)}장 · 원가 ${fmtW(fo.cost * 1.1)} · 현재고 ${fmtN(fo.stock)}장` : `1/1~7/31 판매 <b>${fmtN(sr.sales)}장</b> (26SS) · 8/1~ 판매 <b>${fmtN(fo.sales)}장</b> (26FW) · 현재고 ${fmtN(fo.stock)}장`;
  const oldGrp = `<details class="grp" data-grp="old"${grpOpen('old')}><summary>26SS 러닝 구품번 ${D.nSs}개 + GF1 ${(D.gf1Cut || []).length}개 <small>26SS 러닝 목록 ${D.nSs}개(FW 과거품번 ${D.ssSplit}개는 1/1~7/31 → 26SS, 8/1~ → 26FW 러닝 · 그 밖 ${D.ssWhole}개는 1/1~ 전부 26SS) + 26SS 신상 GF1 ${(D.gf1Cut || []).length}개(8/1~ → 26FW 러닝)</small><span style="margin-left:auto">${oldSum}</span></summary>
    <div class="tw"><table><thead>${oldHead}</thead><tbody>${pairs.map(oldRow).join('')}</tbody></table></div></details>`;
  const exSum = showSS && !showFW ? `판매 <b>${fmtN(xr.sales)}장</b> · 투입 ${fmtN(xr.qtyIn)}장 · 원가 ${fmtW(xr.cost * 1.1)} · 재고 ${fmtN(xr.stock)}장` : showFW && !showSS ? `8/1~ 몫은 위 26FW 러닝에 포함(FW 과거품번 ${D.extraSplit}개) · 1/1~7/31 몫 판매 ${fmtN(xr.sales)}장은 연간에만` : `판매 <b>${fmtN(xr.sales)}장</b> · 투입 ${fmtN(xr.qtyIn)}장 · 원가 ${fmtW(xr.cost * 1.1)} (연간에만 더함)`;
  const exGrp = `<details class="grp" data-grp="extra"${grpOpen('extra')}><summary>과년차 러닝 상품 ${D.nExtra}개 <small>26SS 러닝 목록 밖의 과거 시즌(24SS~25FW) 러닝 · FW 과거품번 ${D.extraSplit}개는 1/1~7/31 몫, 그 밖 ${D.extraWhole}개는 1/1~ 전체 · 26SS/26FW 어느 시즌에도 안 넣고 연간에만 별도 칸</small><span style="margin-left:auto">${exSum}</span></summary>
    <div class="tw"><table><thead>${oldHead}</thead><tbody>${extras.map(oldRow).join('')}</tbody></table></div></details>`;
  const sp = D.split || {}, chk = sp.check || {}, ok = Object.values(chk).every(v => Math.abs(v) <= 1);
  const head = SEASON === 'ss' ? `<b>26SS 러닝</b> = 구품번 ${D.nSs}개의 1/1~7/31 몫 — 판매 <b>${fmtN(sr.sales)}장</b> · 투입 ${fmtN(sr.qtyIn)}장 · 원가 ${fmtW(sr.cost * 1.1)}`
    : SEASON === 'fw' ? `<b>26FW 러닝</b> = 신품번 ${D.nNew}개 <b>${fmtN(fn.sales)}장</b> + 구품번 8/1~ <b>${fmtN(fo.sales)}장</b> = <b>${fmtN(fn.sales + fo.sales)}장</b> · 원가 ${fmtW((fn.cost + fo.cost) * 1.1)} · 현재고 ${fmtN(fn.stock + fo.stock)}장`
    : `<b>연간 러닝</b> 판매 <b>${fmtN(sr.sales + fo.sales + fn.sales + xr.sales)}장</b> = 26SS 러닝 ${fmtN(sr.sales)} + 26FW 러닝(신품번 ${fmtN(fn.sales)} + FW 과거품번 8/1~ ${fmtN(fo.sales)}) + 과년차 러닝 ${fmtN(xr.sales)} · 원가 ${fmtW((sr.cost + fo.cost + fn.cost + xr.cost) * 1.1)}`;
  return kpis() + `<div class="sec"><h2>${esc(SNAME[SEASON])} · 러닝 <small>신품번 / 구품번 그룹 — 누르면 품번 목록이 펼쳐집니다 · 짝 ${sp.nPairs || 0}쌍</small></h2>
    <div class="headline">${head}</div>
    ${showFW ? newGrp : ''}${oldGrp}${exGrp}
    <div class="note">짝 = 러닝아이템 탭의 ‘26FW - 과거품번 → 26FW’ · 7/31 기준 = ${esc(sp.tab731 || '시점재고(~26.07.31)')} 탭(칼라 행을 품번으로 합침) · 검산(구품번 1/1~ 전체 − 나눈 합): 판매 ${fmtN(chk.sales || 0)} · 투입 ${fmtN(chk.qtyIn || 0)} · 원가 ${fmtN(chk.cost || 0)} → <span class="${ok ? 'ok' : 'bad'}">${ok ? '일치' : '차이 있음'}</span>${(sp.miss731 || []).length ? ` · ⚠ 7/31 탭에 없는 구품번 ${sp.miss731.length}개(${esc(sp.miss731.join(', '))})` : ''}</div></div>`;
}
/* 온 · 오프 — 허브 MARK 채널 자료(26FW 만) */
const CH = D.chn || {}, chKeys = {sales:['on', 'off', 'etc'], wk:['onW', 'offW', 'etcW'], cost:['costOn', 'costOff', 'costEtc']};
function chSum(rows){ const t = {on:0, onW:0, off:0, offW:0, etc:0, etcW:0, sales:0, wk:0, cost:0, costOn:0, costOff:0, costEtc:0, cards:0, chCards:0}; rows.forEach(r => Object.keys(t).forEach(k => { t[k] += (r[k] || 0); })); return t; }
const chRows = (scope, item) => (CH.rows || []).filter(r => (scope === '전체' || r.g === scope) && (!item || r.bok === item));
const pct = (a, b) => b ? a / b : null;
function stackRows(items){
  return items.map(it => { const tot = it.parts.reduce((a, p) => a + (p.v || 0), 0); if(!tot) return `<div class="st-row"><div class="st-name">${esc(it.name)}</div><div class="dim">값 없음</div></div>`;
    const segs = it.parts.map(p => { const s = p.v / tot * 100; return `<span class="${p.cls}" style="flex:${Math.max(s, 0.4)} 1 0" title="${esc(p.label)} ${s.toFixed(1)}%">${s >= 14 ? s.toFixed(0) + '%' : ''}</span>`; }).join('');
    return `<div class="st-row"><div class="st-name">${esc(it.name)}</div><div class="st-bar">${segs}</div><div class="st-cap">${it.parts.map(p => `${esc(p.label)} ${(p.v / tot * 100).toFixed(1)}% (${it.fmt(p.v)})`).join(' · ')}</div></div>`; }).join('');
}
function viewOnOff(){
  if(SEASON === 'ss') return kpis() + `<div class="sec"><div class="card warn">26SS 는 채널(온라인·오프라인)별 판매 자료가 없습니다 — 허브 MARK 채널 자료에 26SS 가 없어 온·오프는 26FW 에서만 볼 수 있습니다. <button class="chip" data-season="fw">26FW 온·오프 보기</button></div></div>`;
  if(!(CH.rows || []).length) return kpis() + `<div class="sec"><div class="card warn">온/오프 판매 자료를 읽지 못했습니다 — 빌드 PC 에서 허브 MARK 채널 자료(chteam.json)가 열리는지 확인하세요.</div></div>`;
  const scopes = ['전체', '가을', '겨울', '러닝'];
  const chips = `<div class="chips"><span class="flabel">범위</span>${scopes.map(s => `<button class="chip${OSCOPE === s ? ' on' : ''}" data-oscope="${s}">${s === '전체' ? '전체 (신상품+러닝 신품번)' : s === '러닝' ? '러닝 (신품번)' : s}</button>`).join('')}</div>`;
  const rows = chRows(OSCOPE), T = chSum(rows), tot = T.on + T.off + T.etc, totW = T.onW + T.offW + T.etcW;
  const pOn = pct(T.on, tot), pOff = pct(T.off, tot), pOnW = pct(T.onW, totW), shift = (pOn != null && pOnW != null) ? (pOnW - pOn) * 100 : null;
  const lab = OSCOPE === '전체' ? '신상품 + 러닝(신품번)' : OSCOPE === '러닝' ? '러닝(신품번)' : OSCOPE;
  const gapC = T.sales - tot;
  const banner = SEASON === 'all' ? `<div class="note" style="margin:0 0 10px">연간 시즌이지만 이 화면은 <b>26FW 부분만</b> 나옵니다 — 허브 MARK 채널 자료에 26SS 가 없어 SS+FW 합계는 만들 수 없습니다.</div>` : '';
  const head = `<div class="headline">26FW ${esc(lab)} 누계판매 <b>${fmtN(tot)}장</b> 중 <b>온라인 ${fmtN(T.on)}장(${fmtP(pOn)})</b> · <b>오프라인 ${fmtN(T.off)}장(${fmtP(pOff)})</b>${T.etc ? ` · 기타 ${fmtN(T.etc)}장` : ''} <span class="dim">· MARK ${esc(CH.at || '기준일 미확인')}</span>
    <div class="story">이번 주 판매 ${fmtN(totW)}장은 온라인 <b>${fmtN(T.onW)}장</b> · 오프라인 <b>${fmtN(T.offW)}장</b>${T.etcW ? ` · 기타 ${fmtN(T.etcW)}장` : ''} — 온라인 비중 ${fmtP(pOnW)}로, 누계(${fmtP(pOn)})보다 ${shift == null ? '-' : `<b class="${shift >= 0 ? 'up' : 'down'}">${Math.abs(shift).toFixed(1)}%p ${shift >= 0 ? '높습니다' : '낮습니다'}</b>`}.</div>
    <div class="story">판매 원가(스타일 원가 × 판매량, VAT 포함) <b>${fmtW(T.cost)}</b> 중 온라인 ${fmtW(T.costOn)} · 오프라인 ${fmtW(T.costOff)}${T.costEtc ? ` · 기타 ${fmtW(T.costEtc)}` : ''}.${T.sales ? ` 스타일맵 판매 ${fmtN(T.sales)}장 중 채널 확인분 ${fmtN(tot)}장(${fmtP(tot / T.sales)})${gapC ? ` — 차이 ${fmtN(Math.abs(gapC))}장은 원인 미확인` : ''}.` : ''}${CH.at && CH.end && String(CH.at) !== String(CH.end) ? ` <span class="bad">⚠ 채널 자료 기준일(${esc(CH.at)})이 스타일맵(${esc(CH.end)})과 다릅니다.</span>` : ''}</div></div>`;
  const k = (l, v, s) => `<div class="kpi fw"><div class="l">${l}</div><div class="v">${v}</div><div class="s">${s}</div></div>`;
  const kp = `<div class="kpis">${k('판매 원가 · 채널별', fmtW(T.cost), `온라인 ${fmtW(T.costOn)} · 오프라인 ${fmtW(T.costOff)} · 기타 ${fmtW(T.costEtc)}`)}${k('온라인 판매 (누계)', fmtN(T.on) + '<small style="display:inline;font-size:13px"> 장</small>', `전체 대비 ${fmtP(pOn)}`)}${k('오프라인 판매 (누계)', fmtN(T.off) + '<small style="display:inline;font-size:13px"> 장</small>', `전체 대비 ${fmtP(pOff)}`)}${k('온라인 비중 (이번 주)', fmtP(pOnW), `누계 ${fmtP(pOn)} ${shift == null ? '' : `· ${shift >= 0 ? '▲' : '▼'} ${Math.abs(shift).toFixed(1)}%p`}`)}</div>`;
  const groups = OSCOPE === '전체' ? ['가을', '겨울', '러닝'] : [OSCOPE], trs = [];
  const line = (name, t, cls, sub) => { const tt = t.on + t.off + t.etc, tw = t.onW + t.offW + t.etcW; trs.push(`<tr class="${cls}"><th>${esc(name)}${sub ? `<small>${esc(sub)}</small>` : ''}</th><td class="n">${fmtN(t.on)}</td><td class="n">${fmtN(t.off)}</td><td class="n">${fmtN(t.etc)}</td><td class="n">${fmtP(pct(t.on, tt))}</td><td class="n">${fmtN(t.onW)}</td><td class="n">${fmtN(t.offW)}</td><td class="n">${fmtP(pct(t.onW, tw))}</td><td class="n">${fmtW(t.costOn)}</td><td class="n">${fmtW(t.costOff)}</td><td class="n">${fmtW(t.cost)}</td></tr>`); };
  groups.forEach(g => { const gr = chRows(g); if(!gr.length) return; line(g === '러닝' ? '러닝 (26FW 신품번)' : g, chSum(gr), 'sub'); ['OUTER', 'TOP', 'BOTTOM', 'ACC'].forEach(it => { const rr = chRows(g, it); if(rr.length){ const t = chSum(rr); if(t.sales || t.on || t.off) line('　' + it, t, 'l2'); } }); });
  if(OSCOPE === '전체') line('합계 (신상품 + 러닝 신품번)', chSum(CH.rows), 'tot');
  const tbl = `<div class="tw"><table><thead><tr><th rowspan="2">구분</th><th colspan="4" style="text-align:center">누계 판매량</th><th colspan="3" style="text-align:center">이번 주 판매량</th><th colspan="3" style="text-align:center">판매 원가 (VAT 포함)</th></tr><tr><th class="n">온라인</th><th class="n">오프라인</th><th class="n">기타</th><th class="n">온라인 비중</th><th class="n">온라인</th><th class="n">오프라인</th><th class="n">온라인 비중</th><th class="n">온라인</th><th class="n">오프라인</th><th class="n">합계<small>기타·미확인 포함</small></th></tr></thead><tbody>${trs.join('')}</tbody></table></div>`;
  const stacks = `<div class="legend"><span><i class="ch-on"></i>온라인</span><span><i class="ch-off"></i>오프라인</span><span><i class="ch-etc"></i>기타(글로벌 등)</span></div>` + stackRows([['누계 판매량', 'sales', v => fmtN(v) + '장'], ['이번 주 판매량', 'wk', v => fmtN(v) + '장'], ['판매 원가', 'cost', v => fmtW(v)]].map(([n, m, f]) => ({name: n, fmt: f, parts: [{label:'온라인', cls:'ch-on', v:T[chKeys[m][0]]}, {label:'오프라인', cls:'ch-off', v:T[chKeys[m][1]]}, {label:'기타', cls:'ch-etc', v:T[chKeys[m][2]]}]})));
  return kpis() + `<div class="sec"><h2>26FW · 온 · 오프 <small>허브 MARK 채널 자료 · 누계와 이번 주만 있고 전년·SS 비교는 없음</small></h2>${banner}${chips}${head}${kp}
    <div class="card" style="margin-top:12px"><h3>채널 구성</h3>${stacks}</div>
    <div class="card" style="margin-top:12px"><h3>시즌 · 복종별 채널 판매</h3>${tbl}<div class="note">러닝은 26FW 신품번 판매만(구품번 8/1~ 몫은 채널 자료가 없음) · 기준일 MARK ${esc(CH.at || '-')} · 스타일맵 ${esc(CH.end || '-')}.</div></div></div>`;
}
/* 전체 표 — 시즌·복종 / 복종·카테고리 / 시즌별 카테고리 / 성별·복종 / 성별·카테고리 (+ 26FW 러닝 신품번 복종·카테고리) */
const TCOLS = [['styles', 'ST수', (c, p) => fmtN(c.styles)], ['plan', '기획량', c => fmtN(c.plan)], ['qtyIn', '입고량', c => fmtN(c.qtyIn)], ['inRate', '입고율', c => fmtP(c.plan ? c.qtyIn / c.plan : null)], ['costV', '원가(VAT+)', c => fmtW((c.cost || 0) * 1.1)], ['tag', 'TAG 금액', c => fmtW(c.tag)],
  ['sales', '누계판매', c => fmtN(c.sales)], ['sellRate', '판매율', c => fmtP(c.qtyIn ? c.sales / c.qtyIn : null)], ['salesAmt', '실판매금액', c => fmtW(c.salesAmt)], ['disc', '할인율', c => fmtP(c.salesTag ? 1 - (c.salesAmt || 0) / c.salesTag : null)], ['wkQty', '주간판매', c => fmtN(c.wkQty)]];
function genTable(rows, labelFn, keyPrefix){
  const span = TCOLS.length + 1;
  const body = rows.map((r, i) => { const c = r.cur || {}, p = r.py, k = keyPrefix + i, open = OPEN.has(k), cls = /TTL|TOTAL|합계/.test(String(r.kind || r.item || '')) || r.item === 'TTL' || r.item === 'TOTAL' ? 'sub' : 'l1';
    const det = p ? `<tr class="det"${open ? '' : ' hidden'}><td colspan="${span}"><div class="dg"><b>전년 동시점</b>${TCOLS.filter(x => !['styles'].includes(x[0])).map(([key, l, f]) => { const a = f(c), b = f(p); const num = ['plan', 'qtyIn', 'sales', 'costV', 'tag', 'salesAmt', 'wkQty'].includes(key) ? dPct(key === 'costV' ? c.cost : c[key], key === 'costV' ? p.cost : p[key]) : null; return `<span>${l} <b>${a}</b> <span class="dim">vs ${b}</span> ${num == null ? '' : dHtml(num)}</span>`; }).join('')}</div></td></tr>` : `<tr class="det"${open ? '' : ' hidden'}><td colspan="${span}"><span class="dim">전년 자료 없음</span></td></tr>`;
    return `<tr class="${cls} exp"><th data-exp="${esc(k)}"><span class="tg">${open ? '▾' : '▸'}</span>${labelFn(r)}</th>${TCOLS.map(([key, l, f]) => `<td class="n">${f(c)}</td>`).join('')}</tr>${det}`; }).join('');
  return `<div class="tw"><table><thead><tr><th>구분 <small style="display:inline">▸ 전년 비교</small></th>${TCOLS.map(x => `<th class="n">${x[1]}</th>`).join('')}</tr></thead><tbody>${body}</tbody></table></div>`;
}
function viewTables(){
  const T = D.tables || {}, parts = [];
  const sec = (title, sub, html, key) => `<details class="grp" data-grp="${esc(key)}"${OPEN.has('grp:' + key) ? ' open' : ''}><summary>${esc(title)} <small>${esc(sub)}</small></summary>${html}</details>`;
  const lab1 = r => `${esc(r.season || '')}${r.season ? ' · ' : ''}${esc(r.item === 'TTL' ? (r.kind || 'TTL') : r.item || '')}${r.mode === 'run' && r.item !== 'TTL' ? '<small>러닝</small>' : ''}`;
  const lab2 = r => `${esc(r.item || '')}${r.cat ? ` · ${esc(r.cat)}` : ''}`;
  const lab3 = r => `${esc(r.season || '')} · ${esc(r.item || '')}${r.cat ? ` · ${esc(r.cat)}` : ''}`;
  const lab4 = r => `${esc(r.gender === '남성' ? '유니맨' : r.gender || '')} · ${esc(r.item || '')}${r.cat ? ` · ${esc(r.cat)}` : ''}`;
  const one = (S, tag, src) => {
    parts.push(sec(`${tag} ① 시즌 · 복종`, '26SS 탭/연간 기준 · 신상 + 러닝 TTL', genTable(src.block1 || [], lab1, tag + 'b1_'), tag + 'b1'));
    parts.push(sec(`${tag} ② 복종 · 카테고리`, '신상', genTable(src.newCat || [], lab2, tag + 'nc_'), tag + 'nc'));
    parts.push(sec(`${tag} ③ 시즌별 카테고리`, '신상', genTable(src.seasonCat || [], lab3, tag + 'sc_'), tag + 'sc'));
    parts.push(sec(`${tag} ④ 성별 · 복종`, '신상 · 전년 없음', genTable(src.genderItem || [], lab4, tag + 'gi_'), tag + 'gi'));
    parts.push(sec(`${tag} ⑤ 성별 · 카테고리`, '신상 · 전년 없음', genTable(src.genderCat || [], lab4, tag + 'gc_'), tag + 'gc'));
  };
  if(SEASON !== 'fw') one(SEASON, '26SS', T.ss || {});
  if(SEASON !== 'ss'){ one(SEASON, '26FW', T.fw || {}); const fr = T.fwRun || {}; if((fr.runCat || []).length) parts.push(sec('26FW ⑥ 러닝 신품번 · 복종 · 카테고리', '26FW 집계표 기준(최초기획·기획입고) · 전년은 25FW 러닝', genTable((fr.runItem || []).concat(fr.runCat || []), lab2, 'fwrc_'), 'fwrc')); }
  return kpis() + `<div class="sec"><h2>${esc(SNAME[SEASON])} · 전체 표 <small>표를 누르면 펼쳐지고, 행을 누르면 전년 동시점 비교가 나옵니다</small></h2>${parts.join('')}
    <div class="note">26SS 표는 ERP 생판재(구글 시트) 기준 · 26FW 표는 연간과 같은 기준(리오더 포함 · 입고 수량은 주간판매 화면) · 러닝 TTL 행은 26SS 는 구품번 1/1~7/31, 26FW 는 신품번. 원가는 VAT 포함.</div></div>`;
}
function rule(){
  return `<div class="sec"><div class="rule"><b>기준 요약</b><ul>
<li><b>26SS 신상</b> ERP 2026 봄·여름 ${D.nSsNew}품번(사은품·다이어리 제외) · <b>26FW 신상</b> 2026 가을·겨울 ${D.nFwNew}품번(사은품 제외). 기획 = 작지, 입고 = 생산 전체(리오더 포함).</li>
<li><b>26SS 러닝</b> = 러닝아이템 탭 26SS 열의 <b>${D.nSs}개</b>(대시보드 26SS 종합과 같은 목록). 이 중 ‘26FW - 과거품번’에도 있는 ${D.ssSplit}개는 1/1~7/31 몫만, 그 밖 ${D.ssWhole}개는 1/1~ 전체.</li>
<li><b>26FW 러닝 이월</b> = ‘26FW - 과거품번’ <b>${D.nFwOld}개</b>의 8/1~ 몫(대시보드 26FW 집계표의 ‘시점재고 합산’과 같은 목록 · 7/31 시점재고 탭 기준). 26SS 신상인 GF1 ${(D.gf1Cut || []).length}개(${(D.gf1Cut || []).join(', ')})는 7/31 까지 몫이 26SS 신상에 그대로 있습니다. <b>러닝 신품번 ${D.nNew}개</b>(GF7·GF8)는 전부 26FW 이고 구품번과 짝지어 합산합니다.</li>
<li><b>과년차 러닝 상품 ${D.nExtra}개</b> = 러닝아이템 탭 과거 시즌 열(24SS~25FW)에만 있는 구품번. 26SS 러닝 목록에 없어 26SS/26FW 어느 시즌에도 넣지 않고 연간에만 별도 칸으로 더합니다(FW 과거품번 ${D.extraSplit}개는 1/1~7/31 몫만 — 8/1~ 는 26FW 러닝 이월에 이미 있음 · 그 밖 ${D.extraWhole}개는 1/1~ 전체).</li>
<li>한 품번은 한 칸에만 들어가므로 26SS + 26FW + 과년차 러닝 = 연간이 됩니다. 온·오프는 허브 MARK 채널 자료가 26FW 만 있어 26FW 부분만 나옵니다.</li></ul></div></div>`;
}
function render(){
  document.querySelectorAll('#segSeason .opt').forEach(b => b.classList.toggle('on', b.dataset.season === SEASON));
  document.querySelectorAll('#segView .opt').forEach(b => b.classList.toggle('on', b.dataset.view === VIEW));
  const rngTxt = `기준일 ${D.refS}(일) · 원천 구글 시트 ERP 자동 조회 · 원가 VAT 포함 · 만든 시각 ${D.gen}`;
  document.getElementById('rng').textContent = rngTxt;
  if(IN_FRAME){ document.querySelector('.hd1').style.display = 'none'; try{ window.parent.postMessage({type:'page-meta', text: rngTxt}, '*'); }catch(e){} }   // 허브 안에서는 허브 머리글이 제목을 보여 주므로 페이지 제목은 숨기고 기준일만 올려 보낸다
  const V = {overview: viewOverview, item: viewItem, gender: viewGender, run: viewRun, onoff: viewOnOff, tables: viewTables};
  document.getElementById('view').innerHTML = (V[VIEW] || viewOverview)();
  document.title = `${SNAME[SEASON]} 집계표`;
  history.replaceState(null, '', `#season=${SEASON}&view=${VIEW}&metric=${AMET}${VIEW === 'onoff' ? '&oscope=' + encodeURIComponent(OSCOPE) : ''}`);
}
document.getElementById('segSeason').addEventListener('click', e => { const b = e.target.closest('.opt'); if(b && b.dataset.season){ SEASON = b.dataset.season; render(); } });
document.getElementById('segView').addEventListener('click', e => { const b = e.target.closest('.opt'); if(b && b.dataset.view){ VIEW = b.dataset.view; render(); } });
document.getElementById('view').addEventListener('click', e => {
  const c = e.target.closest('[data-amet]'); if(c){ AMET = c.dataset.amet; render(); return; }
  const o = e.target.closest('[data-oscope]'); if(o){ OSCOPE = o.dataset.oscope; render(); return; }
  const sb = e.target.closest('[data-season]'); if(sb){ SEASON = sb.dataset.season; render(); return; }
  const x = e.target.closest('[data-exp]'); if(x){ const k = x.dataset.exp; const tr = x.closest('tr'); let det = tr.nextElementSibling; if(det && det.classList.contains('det')){ const show = det.hasAttribute('hidden'); det.toggleAttribute('hidden', !show); x.querySelector('.tg').textContent = show ? '▾' : '▸'; if(show) OPEN.add(k); else OPEN.delete(k); } }
});
document.getElementById('view').addEventListener('toggle', e => { const d = e.target; if(d.matches && d.matches('details.grp')){ const k = 'grp:' + d.dataset.grp; if(d.open) OPEN.add(k); else OPEN.delete(k); } }, true);
(function(){ const h = new URLSearchParams(location.hash.replace(/^#/, '')); if(['ss','fw','all'].includes(h.get('season'))) SEASON = h.get('season'); if(['overview','item','gender','run','onoff','tables'].includes(h.get('view'))) VIEW = h.get('view'); if(MET[h.get('metric')]) AMET = h.get('metric'); if(['전체','가을','겨울','러닝'].includes(h.get('oscope'))) OSCOPE = h.get('oscope'); })();
render();
</script></body></html>"""
