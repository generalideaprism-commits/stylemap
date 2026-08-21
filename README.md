# stylemap — 26FW 판매자료 카드 · 시즌 집계표

`26FW 스타일맵.xlsx` 한 파일에서 **공유용 HTML 2종**을 만들어 두는 저장소입니다.
production-hub 의 `PIP 공유파일`(`/share/`) 에 그대로 얹어 쓰는 것을 전제로 합니다.

```
share/26fw-sales-cards.html      판매자료 카드 (품번별 카드 204장)
share/26fw-season-summary.html   시즌 집계표 (5종 · 지표 8종 전환)
build/                           빌드 스크립트 · 템플릿 · 이미지 저장소
data/                            원본 스타일맵 엑셀 (재생성용 스냅샷)
```

## 1. 주간 업데이트 방법

1. 새 `YYMMDD_26FW 스타일맵.xlsx` 를 `data/` 에 넣는다 (파일명 앞 6자리가 화면의 *업데이트 날짜*).
2. 빌드한다. → `share/` 의 HTML 2개가 새로 만들어진다.

```bash
python build/build_salecards.py data/260818_26FW\ 스타일맵.xlsx
```

3. 커밋 & 푸시. production-hub 쪽 `share/index.html` 의 해당 항목 `updated` / `history` 도 같이 고쳐 준다.

새 품번이 추가돼 도식화 이미지가 비어 있으면, 빌드 전에 한 번 돌린다 (게시된 구글 시트에서 이미지 수집).

```bash
python build/fetch_pubhtml_images.py 1620257578
```

## 2. 데이터가 만들어지는 규칙

**판매 카드** — `생판재-26FW(ST)` / `생판재-러닝(ST)` / `주간판매-26FW(CO)` / `주간판매-러닝(CO)` 시트

| 항목 | 출처 |
|---|---|
| TAG가 / 실판가 / 원가 | ST 시트 Z / Y / AA열, 배수 = TAG ÷ 원가 ÷ 1.1 |
| 시즌 · 성별 · 품목 · 생산처 | ST 시트 D / E / F / P열 (러닝 시즌의 `(러닝)` 표기는 제거) |
| 메인 / 러닝 | 26FW 시트에 있으면 메인, 없으면 러닝 |
| 입고여부 | CO 시트 V열(입고) 합계 > 0 이면 `입고` |
| 컬러별 기획 / 입고 / 전주 / 2주전 / 누계 / 판매율 | CO 시트 R / V / CE / CD / CG / CH열 |
| 입고일 · 출고일 | CO 시트 U · Y열 중 가장 빠른 날짜 |
| TAG판매금액 / 실판매금액 | ST 시트 BH / BI열, 할인율 = 1 − 실판매 ÷ TAG판매 |
| 품번 통합 | CO 러닝 시트 **CU열** 기준으로 기존품번 + 러닝 신규품번을 한 장으로 합산 (카드 클릭 시 신규품번 단독 실적) |

**집계표** — `종합` 시트 (① 시즌·복종별 7~23행, ②③④ 30~64행, ⑤ 전년비교 AD~AI열)
막대 길이는 비율 지표면 0~100% 절대 기준, 수량·금액 지표면 표 안 최대 항목(TTL 제외) 대비.

## 3. 이미지 처리

1. **Supabase 갤러리에 있으면 URL 링크** — `…/product-gallery/{품번}/cut_0.jpg` (파일 용량 절약)
2. 없으면 `build/gs_images.json` 의 **base64 도식화 이미지를 HTML 안에 정적 삽입**
3. 러닝 품번은 `품번 → CU열 신품번 → 러닝 매핑표 짝 품번` 순으로 후보를 시도

`build/supabase_cache.json` 은 Supabase 존재 여부 캐시라 지우면 빌드 때 다시 조회한다.

## 4. production-hub 에 등록하기

`share/index.html` 의 `FILES` 배열에 아래를 추가하면 좌측 메뉴·메타·이력·링크복사가 자동 생성된다.

```js
,{
  key:'cards26fw', sec:'gi26fw',
  label:'26FW 판매자료 카드', en:'26FW SALES CARDS',
  file:'26fw-sales-cards.html',
  team:['디자인실','영업팀'], owner:'PIP팀',
  updated:'2026-08-18', status:'live',
  note:'품번별 판매 카드. 시즌·메인/러닝·입고여부·성별·품목 필터와 품번 검색. 러닝은 CU열 기준 구/신 품번 통합.',
  history:[{d:'2026-08-18', t:'최초 등록 — 카드 204장'}]
}
,{
  key:'sum26fw', sec:'gi26fw',
  label:'26FW 시즌 집계표', en:'26FW SEASON SUMMARY',
  file:'26fw-season-summary.html',
  team:['디자인실','영업팀'], owner:'PIP팀',
  updated:'2026-08-18', status:'live',
  note:'종합 시트 기반 집계 5종과 전년 동시점 비교.',
  history:[{d:'2026-08-18', t:'최초 등록'}]
}
```
