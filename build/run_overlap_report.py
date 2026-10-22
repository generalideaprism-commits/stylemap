# -*- coding: utf-8 -*-
"""러닝 겹침 점검표(엑셀) — 26SS 러닝과 26FW 러닝(시점재고 합산)이 같은 구품번을 두 번 세는 것을 찾아 뺀 근거.
build/run_annual_snapshot.json(빌드 때 run_annual.py 가 만든 값)을 읽어 바탕화면에 '26FW_러닝_겹침_점검.xlsx' 를 쓴다.
사용:  python build/run_overlap_report.py [저장경로]"""
import json
import os
import sys

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.expanduser('~'), 'Desktop', '26FW_러닝_겹침_점검.xlsx')
R = json.load(open(os.path.join(HERE, 'run_annual_snapshot.json'), encoding='utf-8'))
ck = R['checks']
TOL = ck['tol']
wb = openpyxl.Workbook()
HEAD = PatternFill('solid', fgColor='1F2937')


def sheet(ws, title, header, rows, widths):
    ws.title = title
    ws.append(header)
    for c in ws[1]:
        c.font = Font(bold=True, color='FFFFFF')
        c.fill = HEAD
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    for r in rows:
        ws.append(r)
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'B2'
    ws.row_dimensions[1].height = 34
    for row in ws.iter_rows(min_row=2):
        for c in row:
            if isinstance(c.value, (int, float)):
                c.number_format = '#,##0'
    return ws


ov1, ov2, only, allss = R['overlap'], R['overlapNew'], R['fwOnly'], R['ssAll']
c_all, c1, c2 = ck['carryAll'], ck['carryOverlap'], ck['carryOverlapNew']
ws = wb.active
ws.title = '요약'
lines = [
    ['26FW · 26SS 러닝 겹침 점검', ''],
    ['', ''],
    ['무엇을 점검했나', '26SS 러닝(1/1~)과 26FW 러닝 ‘시점재고 합산’(7/31 시점재고 + 8/1~ 판매)이 같은 구품번을 두 번 세는지, 두 스타일맵이 서로 맞는지(오차 %d장 이내)' % TOL],
    ['SS 러닝 품번', ck['ssN']],
    ['FW 러닝 구품번(이월)', ck['fwOldN']],
    ['① SS 러닝과 겹치는 구품번(W 코드)', ck['overlapN']],
    ['② 26SS 신상품과 겹치는 구품번(GF1 코드)', ck['overlapNewN']],
    ['③ FW 전용 구품번', ck['fwOldN'] - ck['overlapN'] - ck['overlapNewN']],
    ['', ''],
    ['FW 이월 합계 — 7/31 시점재고(장)', c_all['qtyIn']],
    ['  ① 겹침분', c1['qtyIn']],
    ['  ② 겹침분', c2['qtyIn']],
    ['  ③ FW 전용(연간 러닝에 더하는 몫)', c_all['qtyIn'] - c1['qtyIn'] - c2['qtyIn']],
    ['FW 이월 합계 — 8/1~ 판매(장)', c_all['sales']],
    ['  ① 겹침분', c1['sales']],
    ['  ② 겹침분', c2['sales']],
    ['  ③ FW 전용', c_all['sales'] - c1['sales'] - c2['sales']],
    ['', ''],
    ['대조 결과(허용 오차 %d장)' % TOL, ''],
    ['주간판매(52주) 최대 차이 — SS 러닝 %d품번' % ck['ssN'], ck['weeklyMaxDiffAll']],
    ['주간판매(52주) 최대 차이 — 26SS 신상품과 겹침 %d품번' % ck['overlapNewN'], ck['weeklyMaxDiffNew']],
    ['현재고 최대 차이(SS 러닝 / 신상품 겹침)', '%s / %s' % (int(ck['stockMaxDiffAll']), int(ck['stockMaxDiffNew']))],
    ['FW 이월 판매(8/1~)가 SS 주간판매 경계 범위(8/3주~ ~ 7/27주~) 밖인 품번', ', '.join(ck['boundaryFail'] + ck['boundaryFailNew']) or '없음'],
    ['FW 이월 계산 = 종합 시트(CJ~CR)', '일치' if ck.get('carryMatchesSheet') else '불일치'],
    ['SS 러닝 품번이 26SS 신상품에도 있는 경우', ', '.join(ck.get('ssRunInNew') or []) or '없음'],
    ['', ''],
    ['이어붙임(SS 러닝과 겹치는 %d개의 1/1~ 판매)' % ck['overlapN'], ''],
    ['스타일맵 정의(런닝 BACK 판매량 = 올해재고 − 현재고)', ck['overlapSsSales']],
    ['판매기록 기준(런닝 주마감 누계 − 12월기초재고 누계)', ck['overlapSsSalesRec']],
    ['차이(품번별 최대 / %d장 넘는 품번)' % TOL, '%s장 / %s' % (int(ck['ssSalesDiffMax']), ', '.join(ck['ssSalesDiffOver']) or '없음')],
    ['SS 러닝 31품번 전체: 스타일맵 정의 / 판매기록', '%s / %s' % (int(ck['ssSalesBackTotal']), int(ck['ssSalesRecTotal']))],
    ['해석', '두 파일은 같은 ERP 추출이라 판매기록은 일치. SS 스타일맵의 러닝 판매량은 올해재고−현재고로 계산해 신규발주와 실제 입고의 차이만큼 판매기록보다 적다. 연간에는 스타일맵 정의 그대로 사용.'],
    ['', ''],
    ['연간 러닝에 들어 있지 않은 것', 'FW 전용 %d개의 1/1~7/31 판매 %s장(8/1~ 판매만 포함) · 25SS·25FW 러닝 전년 비교' % (ck['fwOldN'] - ck['overlapN'] - ck['overlapNewN'], int(ck.get('fwOnlyJanJulSales') or 0))],
    ['FW 이월 재고 원가(VAT 미포함, 시점재고 시트 재고 원가금액)', '합계 %s원 · 겹침 ① %s원 · ② %s원 · FW 전용 %s원' % (format(int(c_all['cost']), ','), format(int(c1['cost']), ','), format(int(c2['cost']), ','), format(int(c_all['cost'] - c1['cost'] - c2['cost']), ','))],
    ['', ''],
    ['정의', '‘FW 이월 판매(8/1~)’ = 러닝 ST 시트 누계 − 시점재고(7/31) 시트 누계(날짜 경계 정확). SS 주간판매는 주(월~일) 단위라 8/3 주~ 와 7/27 주~ 두 값의 범위로 본다.'],
    ['연간 러닝', 'SS 러닝 전체 + FW 신품번 + FW 전용 이월(③). ①②의 FW 이월분은 SS 쪽에 이미 들어 있어 뺀다.'],
]
for ln in lines:
    ws.append(ln)
ws['A1'].font = Font(bold=True, size=14)
ws.column_dimensions['A'].width = 62
ws.column_dimensions['B'].width = 90
for row in ws.iter_rows(min_row=3):
    row[0].font = Font(bold=row[0].value in ('무엇을 점검했나', '대조 결과(허용 오차 %d장)' % TOL, '정의', '연간 러닝'))
    if isinstance(row[1].value, (int, float)):
        row[1].number_format = '#,##0'
        row[1].alignment = Alignment(horizontal='left')


def ov_rows(lst):
    out = []
    for r in lst:
        out.append([r['code'], r.get('name', ''), r.get('item', ''), r.get('ssIn'), r.get('ssSales'), r.get('ssSalesRec'), (r.get('ssSales') or 0) - (r.get('ssSalesRec') or 0) if r.get('ssSalesRec') is not None else None,
                    r.get('fwCarryStock'), r.get('fwAugSales'), r.get('ssWk0803'), r.get('ssWk0727'),
                    r.get('ssStock'), r.get('fwStock'), (r.get('ssStock') or 0) - (r.get('fwStock') or 0), '○' if r.get('boundaryOk') else '× 확인'])
    return out


hdr = ['품번', '스타일명', '복종', '입고', '1/1~ 판매(스타일맵 정의)', '판매기록 기준(주마감−12월말)', '차이', 'FW 이월 재고(7/31)', 'FW 이월 판매(8/1~)', 'SS 주간판매 8/3주~', 'SS 주간판매 7/27주~', 'SS 현재고', 'FW 현재고', '현재고 차이', '경계 안(±%d)' % TOL]
w13 = [14, 30, 9, 12, 14, 14, 9, 14, 14, 14, 14, 11, 11, 11, 12]
sheet(wb.create_sheet(), '① SS러닝과 겹침', hdr, ov_rows(ov1), w13)
sheet(wb.create_sheet(), '② SS신상품과 겹침', hdr, ov_rows(ov2), w13)
sheet(wb.create_sheet(), '③ FW 전용 이월', ['품번', '스타일명', '복종', 'FW 이월 재고(7/31)', 'FW 이월 판매(8/1~)', '1/1~7/31 판매(연간에 미포함)', 'FW 현재고'],
      [[r['code'], r['name'], r['item'], r['fwCarryStock'], r['fwAugSales'], r.get('janJulSales'), r['fwStock']] for r in only], [14, 34, 9, 16, 16, 20, 12])
sheet(wb.create_sheet(), 'SS 러닝 전 품번 대조', ['품번', '겹침(FW 구품번)', '스타일맵 정의 판매', '판매기록 기준 판매', 'SS 주간판매 7/27주~', 'FW 주간판매 7/27주~', '52주 최대 차이', 'SS 현재고', 'FW 현재고', '현재고 차이'],
      [[r['code'], '○' if r['overlap'] else '', r.get('ssSalesBack'), r.get('ssSalesRec'), r['ssWk0727'], r['fwWk0727'], r['weeklyMaxDiff'], r['ssStock'], r['fwStock'], r['ssStock'] - r['fwStock']] for r in allss], [14, 14, 16, 16, 18, 18, 14, 12, 12, 12])
wb.save(OUT)
print('저장:', OUT)
