# -*- coding: utf-8 -*-
"""판매자료 카드 — 온라인/오프라인 판매 (MD 2026-10-19 "온,오프 판매도 볼 수 있어야 되겠다").

원천: 허브 tools/review-build/data/chteam.json (style → 컬러 → 팀 → {주간·누적·재고·일간}) 와 markteam.json 의 '_at'(MARK 기준일).
      신상 판매 리뷰의 온라인/오프라인 버튼과 같은 MARK 팀 정의다 — 온라인 = 온라인팀(몰·자사몰·풀필먼트·벤더·반품 마이너스 포함),
      오프라인 = 오프라인팀(직영·백화점·아울렛·면세·한컬렉션), 기타 = 글로벌팀 + 기타팀. 스타일맵(NAS)에는 채널별 판매 열이 없다.
스냅샷: 카드는 일요일 마감 스냅샷이라 MARK 기준일이 카드 마감일(W1 마지막 날)과 같아야 숫자가 맞는다.
        허브 값의 기준일이 다르면(예: 화요일 재빌드) 같은 날짜로 저장해 둔 build/chan_snapshot.json 을 우선 쓴다.
값: 컬러별 [온라인 누계, 온라인 전주, 오프라인 누계, 오프라인 전주, 기타 누계, 기타 전주] — 2주전은 원천이 없어 나누지 않는다.
"""
import io
import json
import os


def _g(teams, name, field):
    v = (teams.get(name) or {}).get(field) or 0
    try:
        return int(round(float(str(v).replace(',', ''))))
    except (TypeError, ValueError, OverflowError):
        return 0


def _from_hub(hub_dir):
    ct = json.load(io.open(os.path.join(hub_dir, 'chteam.json'), encoding='utf-8'))
    at = (json.load(io.open(os.path.join(hub_dir, 'markteam.json'), encoding='utf-8')) or {}).get('_at') or ''
    return ct, at


def _vec(teams):
    return [_g(teams, '온라인', '누적'), _g(teams, '온라인', '주간'),
            _g(teams, '오프라인', '누적'), _g(teams, '오프라인', '주간'),
            _g(teams, '글로벌', '누적') + _g(teams, '기타', '누적'), _g(teams, '글로벌', '주간') + _g(teams, '기타', '주간')]


def load(build_dir, hub_dir, want_at, cards, log=print):
    """cards = {키(품번, 러닝은 신품번): [컬러코드…]}.  반환 (채널 dict, MARK 기준일, 출처 문구)
       채널 dict: 키 → {'t':[6], 'c':{컬러:[6]}} — 카드 컬러만 합산해 카드 합계와 같은 범위로 만든다."""
    snap_p = os.path.join(build_dir, 'chan_snapshot.json')
    snap = None
    if os.path.exists(snap_p):
        try:
            snap = json.load(io.open(snap_p, encoding='utf-8'))
        except Exception:
            snap = None
    hub = None
    try:
        hub = _from_hub(hub_dir)
    except Exception as e:
        log('[경고] 온/오프 판매 원천(chteam.json)을 못 읽음 — %s' % e)
    if hub and hub[1] and hub[1] == want_at:
        pass                                                # 기준일이 카드 마감일과 같은 허브 값 — 가장 좋다
    elif snap and snap.get('at') and snap.get('at') == want_at:
        log('온/오프: 허브 값의 MARK 기준일(%s)이 카드 마감일(%s)과 달라 같은 날짜의 지난 스냅샷을 씁니다' % (hub[1] if hub else '-', want_at))
        return snap['data'], snap['at'], '지난 빌드 스냅샷 · MARK %s' % snap['at']
    if not hub:
        if snap:
            log('[경고] 온/오프: 허브 자료가 없어 지난 스냅샷(MARK %s)을 씁니다 — 카드 마감일과 다릅니다' % snap.get('at'))
            return snap['data'], snap['at'], '지난 빌드 스냅샷 · MARK %s' % snap['at']
        return {}, '', ''
    ct, at = hub
    out = {}
    for key, colors in cards.items():
        try:                                                # 카드 하나가 이상해도 나머지는 계속
            by = ct.get(key)
            if not by:
                continue
            tot = [0] * 6
            cs = {}
            for code in colors:
                v = _vec(by.get(code) or {})
                if any(v):
                    cs[code] = v
                    tot = [a + b for a, b in zip(tot, v)]
            if any(tot):
                out[key] = {'t': tot, 'c': cs}
        except Exception as e:
            log('[경고] 온/오프: %s 처리 실패 — %s' % (key, e))
    if at and at == want_at and out:
        tmp = snap_p + '.tmp'
        with io.open(tmp, 'w', encoding='utf-8') as f:
            json.dump({'at': at, 'data': out}, f, ensure_ascii=False, separators=(',', ':'))
        os.replace(tmp, snap_p)                             # 쓰다 끊겨도 기존 스냅샷이 남는다
    elif not out:
        log('[경고] 온/오프: 허브 자료에 이 카드들의 값이 없어 스냅샷을 덮지 않습니다')
    else:
        log('[경고] 온/오프: MARK 기준일(%s)이 카드 마감일(%s)과 다릅니다 — 화면에 날짜를 밝힙니다' % (at, want_at))
    return out, at, '허브 MARK 채널 자료'
