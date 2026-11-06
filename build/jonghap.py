# -*- coding: utf-8 -*-
"""'종합' 시트를 엑셀 없이 계산한다 — 스타일맵 종합(2) 시트의 수식을 규칙으로 삼아 그대로 평가.

  왜: 2026-11 부터 원자료가 구글 시트(ERP 직접 조회)로 들어온다. 구글 시트에는 원자료 탭만 있고 종합 시트가 없으므로,
      엑셀이 하던 집계(시즌·복종·카테고리·성별 × 26FW/25FW/러닝/시점재고)를 빌드가 직접 한다.
  어떻게: 종합(2) 시트의 셀(수식·상수)을 틀(jonghap_template.json)로 떠 두고, 수식을 파싱해 원자료 격자에 대해 계산한다.
      수식이 쓰는 것은 SUMIFS · COUNTIFS · SUM · IFERROR 와 사칙연산·셀 참조뿐이다(2026-11-06 전수 확인: 7,868개).
      규칙을 손으로 옮기지 않으므로 '25FW 신상품에서 무엇을 빼는가' 같은 세부가 엑셀과 달라질 일이 없다.
  검증: 엑셀 파일의 원자료로 계산한 결과가 엑셀이 저장해 둔 값과 전 셀에서 같아야 한다 (selftest).

  엑셀 조건 함수의 동작 중 여기서 지키는 것
    · 조건 값이 숫자든 글자든 '같은 글자 모양'이면 일치로 본다 (년도 칸의 "2025" 글자와 조건 2025 숫자).
    · 글자 비교는 대소문자를 가리지 않는다. 조건의 * ? 는 와일드카드.
    · ">0" 같은 비교 조건은 숫자 칸에만 맞는다.
    · 빈 칸을 가리키는 조건은 0 과 같다. 합계는 숫자 칸만 더한다.
    · IFERROR 는 0 으로 나누기·글자 연산 같은 오류를 두 번째 인자로 바꾼다.
"""
import json, os, re

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, 'jonghap_template.json')


class XlError(Exception):
    pass


def col2idx(s):
    n = 0
    for ch in s: n = n * 26 + (ord(ch) - 64)
    return n - 1


def idx2col(i):
    s = ''; i += 1
    while i: i, r = divmod(i - 1, 26); s = chr(65 + r) + s
    return s


# ---------------------------------------------------------------- 수식 파서
_TOK = re.compile(r'''\s*(?:
    (?P<num>\d+\.?\d*(?:[eE][+-]?\d+)?) |
    (?P<str>"(?:[^"]|"")*") |
    (?P<ref>(?:'(?:[^']|'')+'!|[A-Za-z가-힣0-9_.()~,\s]+?!)?\$?[A-Z]{1,3}\$?\d*(?::\$?[A-Z]{1,3}\$?\d*)?(?![A-Za-z(])) |
    (?P<func>[A-Z][A-Z0-9.]*)\( |
    (?P<op>[-+*/^(),])
)''', re.X)


def tokenize(f):
    out, p = [], 0
    f = f.strip()
    if f.startswith('='): f = f[1:]
    while p < len(f):
        m = _TOK.match(f, p)
        if not m or m.end() == p:
            if f[p:].strip() == '': break
            raise ValueError('수식을 읽을 수 없음 (위치 %d): %s' % (p, f[max(0, p - 20):p + 40]))
        kind = m.lastgroup; out.append((kind, m.group(kind))); p = m.end()
    return out


_REF = re.compile(r"^(?:'((?:[^']|'')+)'!|([^!]+)!)?\$?([A-Z]{1,3})\$?(\d*)(?::\$?([A-Z]{1,3})\$?(\d*))?$")


def parse_ref(text):
    m = _REF.match(text.strip())
    if not m: raise ValueError('참조를 읽을 수 없음: %s' % text)
    sheet = (m.group(1) or m.group(2) or '').replace("''", "'").strip() or None
    c1, r1, c2, r2 = m.group(3), m.group(4), m.group(5), m.group(6)
    if c2 is None: return ('cell', sheet, col2idx(c1), int(r1))
    return ('range', sheet, col2idx(c1), int(r1) if r1 else None, col2idx(c2), int(r2) if r2 else None)


class Parser:
    def __init__(self, toks): self.t, self.p = toks, 0
    def peek(self): return self.t[self.p] if self.p < len(self.t) else (None, None)
    def take(self): v = self.t[self.p]; self.p += 1; return v
    def expr(self):
        n = self.term()
        while self.peek() in (('op', '+'), ('op', '-')):
            op = self.take()[1]; n = ('bin', op, n, self.term())
        return n
    def term(self):
        n = self.unary()
        while self.peek() in (('op', '*'), ('op', '/')):
            op = self.take()[1]; n = ('bin', op, n, self.unary())
        return n
    def unary(self):
        if self.peek() in (('op', '+'), ('op', '-')):
            op = self.take()[1]; return ('neg', self.unary()) if op == '-' else ('pos', self.unary())
        return self.primary()
    def primary(self):
        k, v = self.take()
        if k == 'num': return ('num', float(v))
        if k == 'str': return ('str', v[1:-1].replace('""', '"'))
        if k == 'ref': return parse_ref(v)
        if k == 'func':
            args = []
            if self.peek() != ('op', ')'):
                args.append(self.expr())
                while self.peek() == ('op', ','): self.take(); args.append(self.expr())
            if self.take() != ('op', ')'): raise ValueError('괄호가 맞지 않음')
            return ('func', v, args)
        if (k, v) == ('op', '('):
            n = self.expr()
            if self.take() != ('op', ')'): raise ValueError('괄호가 맞지 않음')
            return n
        raise ValueError('예상 밖의 토큰: %s %s' % (k, v))


def parse(formula):
    p = Parser(tokenize(formula)); n = p.expr()
    if p.p != len(p.t): raise ValueError('수식 끝에 남는 내용: %s' % (p.t[p.p:],))
    return n


# ---------------------------------------------------------------- 평가기
def _num(v):
    """산술용 숫자. 빈칸 0 · 숫자 모양 글자는 숫자 · 그 밖의 글자는 #VALUE!"""
    if v is None or v == '':
        if v == '': raise XlError('#VALUE!')          # 수식이 낸 빈 글자("")는 엑셀에서 산술 오류
        return 0.0
    if isinstance(v, bool): return float(v)
    if isinstance(v, (int, float)): return float(v)
    try: return float(str(v).replace(',', ''))
    except ValueError: raise XlError('#VALUE!')


def _crit_text(v):
    """조건·칸 값을 비교용 글자로 (숫자 1.0 → '1')."""
    if isinstance(v, bool): return 'TRUE' if v else 'FALSE'
    if isinstance(v, (int, float)):
        return str(int(v)) if float(v) == int(v) else repr(float(v))
    return str(v).strip().lower()


class Evaluator:
    def __init__(self, template, sheets, self_name):
        """template: {'A1': '=수식' | 값} · sheets: {시트이름: 격자(행 리스트, 0행=엑셀 1행)} · self_name: 틀 시트 이름"""
        self.t, self.sheets, self.self_name = template, sheets, self_name
        self.ast, self.val, self.stack = {}, {}, set()
        self._colcache, self._maskcache = {}, {}
        self.aliases = {}

    def grid(self, name):
        g = self.sheets.get(self.aliases.get(name, name))
        if g is None: raise KeyError('원자료 시트가 없습니다: %s (있는 것: %s)' % (name, ', '.join(self.sheets)))
        return g

    # --- 셀 ---
    def cell(self, sheet, c, r):
        if sheet in (None, self.self_name): return self.value('%s%d' % (idx2col(c), r))
        g = self.grid(sheet)
        if r - 1 < len(g) and c < len(g[r - 1]): return g[r - 1][c]
        return None

    def value(self, addr):
        if addr in self.val:
            v = self.val[addr]
            if isinstance(v, XlError): raise v
            return v
        raw = self.t.get(addr)
        if not (isinstance(raw, str) and raw.startswith('=')):
            self.val[addr] = raw; return raw
        if addr in self.stack: raise XlError('#순환!')
        self.stack.add(addr)
        try:
            if addr not in self.ast: self.ast[addr] = parse(raw)
            v = self.ev(self.ast[addr])
        except XlError as e:
            self.val[addr] = e; raise
        finally:
            self.stack.discard(addr)
        self.val[addr] = v
        return v

    # --- 범위 ---
    def column(self, sheet, c):
        key = (sheet, c)
        if key not in self._colcache:
            if sheet in (None, self.self_name):
                rows = max((int(re.sub(r'^[A-Z]+', '', a)) for a in self.t), default=0)
                self._colcache[key] = None, rows                      # 틀 시트는 셀마다 평가
            else:
                g = self.grid(sheet)
                self._colcache[key] = [row[c] if c < len(row) else None for row in g], len(g)
        return self._colcache[key]

    def range_values(self, node):
        _, sheet, c1, r1, c2, r2 = node
        col, n = self.column(sheet, c1)
        r1 = r1 or 1; r2 = min(r2 or n, max(n, 1))
        out = []
        for c in range(c1, c2 + 1):
            if sheet in (None, self.self_name):
                for r in range(r1, (node[5] or n) + 1):
                    try: out.append(self.value('%s%d' % (idx2col(c), r)))
                    except XlError as e: out.append(e)
            else:
                cv, _ = self.column(sheet, c)
                out.extend(cv[r1 - 1:r2])
        return out

    def mask(self, node, crit):
        """조건 범위 × 조건 → 맞는 행 번호(엑셀 행) 집합. 같은 (범위, 조건)은 한 번만 계산한다."""
        _, sheet, c1, r1, c2, r2 = node
        if c1 != c2: raise ValueError('조건 범위는 한 열이어야 함')
        ck = crit if not isinstance(crit, float) else ('#', crit)
        key = (sheet, c1, r1, r2, ck if not isinstance(crit, str) else crit.lower())
        if key in self._maskcache: return self._maskcache[key]
        cv, n = self.column(sheet, c1)
        r1_ = r1 or 1; r2_ = min(r2 or n, n)
        rows = range(r1_, r2_ + 1)
        op, val = '=', crit
        if isinstance(crit, str):
            m = re.match(r'^(<>|>=|<=|>|<|=)(.*)$', crit)
            if m: op, val = m.group(1), m.group(2)
        if crit is None: val = 0.0                                   # 빈 칸을 가리키는 조건 = 0
        res = set()
        if op in ('>', '<', '>=', '<='):
            try: x = float(val)
            except (TypeError, ValueError): x = None
            if x is not None:
                for r in rows:
                    v = cv[r - 1]
                    if isinstance(v, (int, float)) and not isinstance(v, bool):
                        if (op == '>' and v > x) or (op == '<' and v < x) or (op == '>=' and v >= x) or (op == '<=' and v <= x): res.add(r)
        else:
            want = _crit_text(val)
            if isinstance(val, str) and ('*' in val or '?' in val):
                rx = re.compile('^' + re.escape(want).replace(r'\*', '.*').replace(r'\?', '.') + '$', re.S)
                hit = lambda v: v is not None and rx.match(_crit_text(v)) is not None
            elif want == '':
                hit = lambda v: v is None or v == ''
            else:
                hit = lambda v: v is not None and v != '' and _crit_text(v) == want
            for r in rows:
                if hit(cv[r - 1]) != (op == '<>'): res.add(r)
        self._maskcache[key] = res
        return res

    # --- 식 ---
    def ev(self, n):
        k = n[0]
        if k == 'num': return n[1]
        if k == 'str': return n[1]
        if k == 'cell': return self.cell(n[1], n[2], n[3])
        if k == 'range': raise ValueError('범위를 값으로 쓸 수 없음')
        if k == 'neg': return -_num(self.ev(n[1]))
        if k == 'pos': return self.ev(n[1])
        if k == 'bin':
            a, b = _num(self.ev(n[2])), _num(self.ev(n[3]))
            if n[1] == '+': return a + b
            if n[1] == '-': return a - b
            if n[1] == '*': return a * b
            if b == 0: raise XlError('#DIV/0!')
            return a / b
        if k == 'func': return self.func(n[1], n[2])
        raise ValueError(k)

    def func(self, name, args):
        if name == 'IFERROR':
            try: return self.ev(args[0])
            except XlError: return self.ev(args[1])
        if name == 'SUM':
            tot = 0.0
            for a in args:
                if a[0] == 'range':
                    for v in self.range_values(a):
                        if isinstance(v, XlError): raise v
                        if isinstance(v, (int, float)) and not isinstance(v, bool): tot += v
                else:
                    tot += _num(self.ev(a))
            return tot
        if name in ('SUMIFS', 'COUNTIFS'):
            pairs = args[1:] if name == 'SUMIFS' else args
            rows = None
            for i in range(0, len(pairs), 2):
                m = self.mask(pairs[i], self.ev(pairs[i + 1]))
                rows = m if rows is None else rows & m
                if not rows: break
            if name == 'COUNTIFS': return float(len(rows or ()))
            sr = args[0]
            if sr[2] != sr[4]: raise ValueError('합계 범위는 한 열이어야 함')
            if (sr[3] or 1) != (pairs[0][3] or 1): raise ValueError('합계 범위와 조건 범위의 시작 행이 다름')
            cv, n = self.column(sr[1], sr[2])
            tot = 0.0
            for r in rows or ():
                v = cv[r - 1] if r - 1 < len(cv) else None
                if isinstance(v, (int, float)) and not isinstance(v, bool): tot += v
            return tot
        raise ValueError('지원하지 않는 함수: %s' % name)

    def compute_all(self):
        """틀의 모든 셀을 계산 → {'A1': 값}. 오류 셀은 '#…' 글자로 둔다(엑셀 표시와 같게)."""
        out = {}
        for addr in self.t:
            try: out[addr] = self.value(addr)
            except XlError as e: out[addr] = str(e)
        return out


# ---------------------------------------------------------------- 틀 뜨기 / 격자 만들기
def extract_template(xlsx, sheet):
    """엑셀의 종합 시트에서 수식·상수를 떠 {'sheet': 이름, 'cells': {'A1': …}} 로 돌려준다."""
    import openpyxl
    wb = openpyxl.load_workbook(xlsx, read_only=True, data_only=False)
    cells = {}
    for i, row in enumerate(wb[sheet].iter_rows(values_only=True), start=1):
        for j, v in enumerate(row):
            if v is None: continue
            if not isinstance(v, (int, float, str)): v = str(v)
            cells['%s%d' % (idx2col(j), i)] = v
    return {'sheet': sheet, 'source': os.path.basename(xlsx), 'cells': cells}


def load_template(path=TEMPLATE):
    return json.load(open(path, encoding='utf-8'))


def to_rows(computed, nrow=None, ncol=None):
    """{'A1': 값} → 행 리스트(빌드가 종합 시트를 읽던 모양: rows[r-1][c])."""
    mr = mc = 0
    for a in computed:
        m = re.match(r'^([A-Z]+)(\d+)$', a); mr = max(mr, int(m.group(2))); mc = max(mc, col2idx(m.group(1)) + 1)
    nrow, ncol = nrow or mr, ncol or mc
    rows = [[None] * ncol for _ in range(nrow)]
    for a, v in computed.items():
        m = re.match(r'^([A-Z]+)(\d+)$', a)
        rows[int(m.group(2)) - 1][col2idx(m.group(1))] = v
    return rows


def compute(sheets, template=None):
    """원자료 격자들로 종합 시트를 계산 → 행 리스트."""
    t = template or load_template()
    ev = Evaluator(t['cells'], sheets, t['sheet'])
    return to_rows(ev.compute_all())


# ---------------------------------------------------------------- 자체 검증
def selftest(xlsx, sheet='종합(2)', verbose=True):
    """엑셀 파일의 원자료로 계산한 값 == 엑셀이 저장해 둔 값 인지 전 셀 대조. 돌려주는 값: (수식 셀 수, 불일치 목록)"""
    import openpyxl
    t = extract_template(xlsx, sheet)
    wv = openpyxl.load_workbook(xlsx, read_only=True, data_only=True)
    need = set()
    for v in t['cells'].values():
        if isinstance(v, str) and v.startswith('='):
            for a, b in re.findall(r"'((?:[^']|'')+)'!|([A-Za-z가-힣0-9_.()~]+)!", v):
                nm = (a or b).replace("''", "'")
                if nm != sheet: need.add(nm)
    sheets = {nm: [list(r) for r in wv[nm].iter_rows(values_only=True)] for nm in need}
    ev = Evaluator(t['cells'], sheets, sheet)
    got = ev.compute_all()
    cached = {}
    for i, row in enumerate(wv[sheet].iter_rows(values_only=True), start=1):
        for j, v in enumerate(row):
            if v is not None: cached['%s%d' % (idx2col(j), i)] = v
    bad, nf = [], 0
    for addr, raw in t['cells'].items():
        if not (isinstance(raw, str) and raw.startswith('=')): continue
        nf += 1
        g, c = got.get(addr), cached.get(addr)
        if isinstance(g, (int, float)) and isinstance(c, (int, float)):
            if abs(g - c) > 1e-6 * max(1.0, abs(c)): bad.append((addr, c, g))
        elif (g in (None, '') and c in (None, '')) or g == c: pass
        elif isinstance(g, (int, float)) and g == 0 and c in (None, ''): pass
        else: bad.append((addr, c, g))
    if verbose:
        print('%s · %s: 수식 %d개 계산 · 엑셀 저장값과 불일치 %d개 · 참조 시트 %s' % (os.path.basename(xlsx), sheet, nf, len(bad), sorted(need)))
        for b in bad[:25]: print('   %s  엑셀=%r  계산=%r  | %s' % (b[0], b[1], b[2], str(t['cells'][b[0]])[:110].replace('\n', ' ')))
    return nf, bad


if __name__ == '__main__':
    import sys
    try: sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception: pass
    if len(sys.argv) >= 3 and sys.argv[1] == 'selftest':
        selftest(sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else '종합(2)')
    elif len(sys.argv) >= 3 and sys.argv[1] == 'extract':
        t = extract_template(sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else '종합(2)')
        json.dump(t, open(TEMPLATE, 'w', encoding='utf-8'), ensure_ascii=False)
        print('틀 저장: %s (%d셀, 수식 %d개)' % (TEMPLATE, len(t['cells']), sum(1 for v in t['cells'].values() if isinstance(v, str) and v.startswith('='))))
    else:
        print('사용: python jonghap.py selftest <xlsx> [시트] | extract <xlsx> [시트]')
