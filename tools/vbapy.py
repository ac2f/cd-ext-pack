#!/usr/bin/env python3
"""ac2f pack - VBA kaynagini Python'a ceviren kucuk derleyici.

Amac, CorelDRAW olmadan modulleri calistirip sinamak:

  * ac2fNestCore gibi saf hesap modulleri dogrudan,
  * ac2fNest / ac2fStart gibi belgeye dokunan moduller, tools/cdrmock.py
    icindeki CorelDRAW nesne modeli taklidiyle (Shape, Curve, SubPath ...).

VBA anlami olabildigince korunur:

  - Option Explicit: tanimsiz bir ad derleme hatasidir
  - Long/Integer/Byte atamada bankaci yuvarlamasi ve tasma denetimi
  - diziler sinir denetimlidir (a(-1) sondan okumaz)
  - And / Or kisa devre yapmaz
  - ByRef skaler parametreler cagirandaki degiskene geri yazilir
  - a = b dizi atamasi kopyalar
  - On Error Resume Next / On Error GoTo etiket / GoTo / Err nesnesi
  - nesnede olmayan bir uye cagrisi VBA'daki gibi hata 438 olur

Desteklenmeyen bir yapi acik hata verir; o yordam cagrildiginda
NotImplementedError atilir (diger yordamlar yine calisir).

Kullanim:
    from vbapy import Project, load_module
    m = load_module('src/ac2fNestCore.bas')          # tek modul
    p = Project(['src/ac2fCore.bas', ...], host)       # coklu + ana bilgisayar
    p.m.ac2fnest.f_ac2fnest()
"""
import math, re, time, types

# =====================================================================
# runtime
# =====================================================================
class VBError(Exception):
    def __init__(self, msg, number=5):
        Exception.__init__(self, msg)
        self.number = number


class _Goto(Exception):
    def __init__(self, label):
        Exception.__init__(self, label)
        self.label = label


class NullStr(str):
    """vbNullString: InputBox'ta Iptal. StrPtr bunun icin 0 dondurur."""
    pass


class Ref(object):
    """Uye cagrilarinda ByRef arguman kutusu."""
    __slots__ = ('v',)

    def __init__(self, v):
        self.v = v


def V(x):
    return x.v if isinstance(x, Ref) else x


def _round_even(x):
    if isinstance(x, bool):
        return -1 if x else 0
    if isinstance(x, int):
        return x
    if isinstance(x, str):
        x = vb_cdbl(x)
    f = math.floor(x)
    d = x - f
    if d > 0.5:
        return int(f) + 1
    if d < 0.5:
        return int(f)
    return int(f) if int(f) % 2 == 0 else int(f) + 1


def _num(x):
    if isinstance(x, bool):
        return -1 if x else 0
    if x is None:
        return 0
    if isinstance(x, str):
        return vb_cdbl(x)
    return x


def vb_cdbl(x):
    if isinstance(x, str):
        t = x.strip()
        try:
            return float(t)
        except ValueError:
            raise VBError('Type mismatch: %r is not a number' % x, 13)
    return float(_num(x))


def conv_long(x):
    v = _round_even(_num(x))
    if v < -2147483648 or v > 2147483647:
        raise VBError('Overflow (Long): %r' % (x,), 6)
    return v


def conv_integer(x):
    v = _round_even(_num(x))
    if v < -32768 or v > 32767:
        raise VBError('Overflow (Integer): %r' % (x,), 6)
    return v


def conv_byte(x):
    v = _round_even(_num(x))
    if v < 0 or v > 255:
        raise VBError('Overflow (Byte): %r' % (x,), 6)
    return v


def conv_double(x):
    return vb_cdbl(x)


def conv_boolean(x):
    if isinstance(x, bool):
        return x
    if isinstance(x, str):
        if x.lower() == 'true':
            return True
        if x.lower() == 'false':
            return False
    return _num(x) != 0


def conv_string(x):
    return vb_cstr(x)


def conv_variant(x):
    return x


def conv_object(x):
    return x


CONV = {'long': conv_long, 'integer': conv_integer, 'byte': conv_byte,
        'double': conv_double, 'single': conv_double, 'boolean': conv_boolean,
        'string': conv_string, 'variant': conv_variant, 'currency': conv_double,
        'object': conv_object}
DEFAULT = {'long': 0, 'integer': 0, 'byte': 0, 'double': 0.0, 'single': 0.0,
           'boolean': False, 'string': '', 'variant': None, 'currency': 0.0,
           'object': None}
SCALAR_TYPES = set(CONV)


def norm_type(t):
    t = (t or 'variant').lower()
    if t in SCALAR_TYPES:
        return t
    if t.startswith('cdr'):
        return 'long'          # CorelDRAW enumerations
    return 'object'


class VArr(object):
    __slots__ = ('lo', 'data', 'typ', 'conv')

    def __init__(self, typ, lo=None, hi=None):
        self.typ = typ
        self.conv = CONV[typ]
        self.lo = 0
        self.data = None
        if lo is not None:
            self.redim(lo, hi, False)

    def redim(self, lo, hi, preserve):
        lo = conv_long(lo)
        hi = conv_long(hi)
        n = hi - lo + 1
        if n < 0:
            raise VBError('Subscript out of range in ReDim (%d To %d)' % (lo, hi), 9)
        dv = DEFAULT[self.typ]
        if preserve and self.data is not None:
            if lo != self.lo:
                raise VBError('ReDim Preserve cannot change the lower bound', 9)
            old = self.data
            if n <= len(old):
                self.data = old[:n]
            else:
                self.data = old + [dv] * (n - len(old))
        else:
            self.data = [dv] * n
        self.lo = lo

    def _ix(self, i):
        if not isinstance(i, int) or isinstance(i, bool):
            i = conv_long(i)
        if self.data is None:
            raise VBError('Subscript out of range: array not allocated', 9)
        j = i - self.lo
        if j < 0 or j >= len(self.data):
            raise VBError('Subscript out of range: %d not in %d..%d'
                          % (i, self.lo, self.lo + len(self.data) - 1), 9)
        return j

    def __getitem__(self, i):
        return self.data[self._ix(i)]

    def __setitem__(self, i, v):
        self.data[self._ix(i)] = self.conv(v)

    def ubound(self, dim=1):
        if self.data is None:
            raise VBError('Subscript out of range: UBound of unallocated array', 9)
        return self.lo + len(self.data) - 1

    def lbound(self, dim=1):
        if self.data is None:
            raise VBError('Subscript out of range: LBound of unallocated array', 9)
        return self.lo

    def assign(self, other):
        if not isinstance(other, VArr):
            raise VBError('Type mismatch: array assignment from non-array', 13)
        if other.typ != self.typ and not (self.typ == 'variant' or other.typ == 'variant'):
            raise VBError('Type mismatch: array of %s into %s' % (other.typ, self.typ), 13)
        self.lo = other.lo
        self.data = None if other.data is None else [self.conv(x) for x in other.data]

    def items(self):
        return [] if self.data is None else list(self.data)


class VArr2(object):
    """Iki boyutlu Variant dizi (GetAllSettings)."""

    def __init__(self, rows):
        self.rows = rows

    def item(self, i, j):
        return self.rows[conv_long(i)][conv_long(j)]

    def ubound(self, dim=1):
        return len(self.rows) - 1 if dim == 1 else len(self.rows[0]) - 1

    def lbound(self, dim=1):
        return 0


def vb_cstr(x):
    if isinstance(x, bool):
        return 'True' if x else 'False'
    if x is None:
        return ''
    if isinstance(x, float):
        if x == int(x) and abs(x) < 1e15:
            return str(int(x))
        r = repr(x)
        if 'e' in r:
            return r
        return r
    return str(x)


def vb_and(a, b):
    if isinstance(a, bool) and isinstance(b, bool):
        return a and b
    return _round_even(_num(a)) & _round_even(_num(b))


def vb_or(a, b):
    if isinstance(a, bool) and isinstance(b, bool):
        return a or b
    return _round_even(_num(a)) | _round_even(_num(b))


def vb_xor(a, b):
    if isinstance(a, bool) and isinstance(b, bool):
        return a != b
    return _round_even(_num(a)) ^ _round_even(_num(b))


def vb_not(a):
    if isinstance(a, bool):
        return not a
    return ~_round_even(_num(a))


def vb_idiv(a, b):
    a = _round_even(_num(a))
    b = _round_even(_num(b))
    if b == 0:
        raise VBError('Division by zero', 11)
    q = abs(a) // abs(b)
    return q if (a >= 0) == (b >= 0) else -q


def vb_mod(a, b):
    a = _round_even(_num(a))
    b = _round_even(_num(b))
    if b == 0:
        raise VBError('Division by zero', 11)
    return a - b * vb_idiv(a, b)


def vb_div(a, b):
    b = _num(b)
    if b == 0:
        raise VBError('Division by zero', 11)
    return _num(a) / b


def vb_add(a, b):
    if isinstance(a, str) and isinstance(b, str):
        return a + b
    return _num(a) + _num(b)


def vb_sub(a, b):
    return _num(a) - _num(b)


def vb_mul(a, b):
    return _num(a) * _num(b)


def vb_neg(a):
    return -_num(a)


def vb_cat(a, b):
    return vb_cstr(a) + vb_cstr(b)


def vb_eq(a, b):
    if isinstance(a, str) or isinstance(b, str):
        return vb_cstr(a) == vb_cstr(b)
    return _num(a) == _num(b)


def vb_ne(a, b):
    return not vb_eq(a, b)


def vb_cmp(a, op, b):
    if isinstance(a, str) and isinstance(b, str):
        x, y = a, b
    else:
        x, y = _num(a), _num(b)
    if op == '<':
        return x < y
    if op == '>':
        return x > y
    if op == '<=':
        return x <= y
    return x >= y


def vb_int(x):
    x = _num(x)
    return math.floor(x) if isinstance(x, float) else x


def vb_fix(x):
    x = _num(x)
    return math.trunc(x) if isinstance(x, float) else x


def vb_sqr(x):
    x = _num(x)
    if x < 0:
        raise VBError('Invalid procedure call: Sqr of negative', 5)
    return math.sqrt(x)


def vb_iif(c, a, b):
    return a if conv_boolean(c) else b


def vb_timer():
    t = time.time()
    return t - 86400.0 * math.floor(t / 86400.0)


def vb_val(s):
    s = vb_cstr(s)
    m = re.match(r'\s*[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?', s)
    return float(m.group(0)) if m else 0.0


def vb_ubound(a, dim=1):
    if not hasattr(a, 'ubound'):
        raise VBError('Type mismatch: UBound of a non-array', 13)
    return a.ubound(dim)


def vb_lbound(a, dim=1):
    if not hasattr(a, 'lbound'):
        raise VBError('Type mismatch: LBound of a non-array', 13)
    return a.lbound(dim)


def vb_mid(s, start, n=None):
    s = vb_cstr(s)
    start = conv_long(start)
    if start < 1:
        raise VBError('Invalid procedure call: Mid start < 1', 5)
    r = s[start - 1:]
    return r if n is None else r[:conv_long(n)]


def vb_left(s, n):
    n = conv_long(n)
    if n < 0:
        raise VBError('Invalid procedure call: Left with negative length', 5)
    return vb_cstr(s)[:n]


def vb_right(s, n):
    s = vb_cstr(s)
    n = conv_long(n)
    if n < 0:
        raise VBError('Invalid procedure call: Right with negative length', 5)
    return s[len(s) - n:] if n else ''


def vb_instr(*a):
    if len(a) == 2:
        start, s, t = 1, a[0], a[1]
    else:
        start, s, t = a[0], a[1], a[2]
    if len(a) == 4 and conv_long(a[3]) == 1:
        return vb_cstr(s).lower().find(vb_cstr(t).lower(), conv_long(start) - 1) + 1
    return vb_cstr(s).find(vb_cstr(t), conv_long(start) - 1) + 1


def vb_replace(s, a, b):
    return vb_cstr(s).replace(vb_cstr(a), vb_cstr(b))


def vb_split(s, sep=' '):
    arr = VArr('string')
    parts = vb_cstr(s).split(vb_cstr(sep))
    arr.redim(0, len(parts) - 1, False)
    for i, p in enumerate(parts):
        arr[i] = p
    return arr


def vb_array(*items):
    arr = VArr('variant')
    arr.redim(0, len(items) - 1, False)
    for i, x in enumerate(items):
        arr[i] = x
    return arr


def vb_trim(s):
    return vb_cstr(s).strip(' ')


def vb_strcomp(a, b, mode=0):
    a, b = vb_cstr(a), vb_cstr(b)
    if conv_long(mode) == 1:
        a, b = a.lower(), b.lower()
    return (a > b) - (a < b)


def vb_isnumeric(x):
    if isinstance(x, (int, float)) and not isinstance(x, bool):
        return True
    try:
        float(vb_cstr(x).strip())
        return vb_cstr(x).strip() != ''
    except ValueError:
        return False


def vb_isarray(x):
    return isinstance(x, (VArr, VArr2))


def vb_strptr(s):
    return 0 if isinstance(s, NullStr) else 1


def vb_ascw(s):
    s = vb_cstr(s)
    if not s:
        raise VBError('Invalid procedure call: AscW of empty string', 5)
    c = ord(s[0])
    return c - 65536 if c > 32767 else c     # VBA returns an Integer


def vb_string(n, ch):
    return vb_cstr(ch)[:1] * conv_long(n)


def vb_space(n):
    return ' ' * conv_long(n)


def vb_str(x):
    v = _num(x)
    s = vb_cstr(float(v)) if isinstance(v, float) else vb_cstr(v)
    if s.startswith('0.'):
        s = s[1:]
    elif s.startswith('-0.'):
        s = '-' + s[2:]
    return (' ' + s) if v >= 0 else s


def vb_format(v, fmt=''):
    if isinstance(v, str):
        try:
            v = float(v)
        except ValueError:
            return v
    v = _num(v)
    fmt = vb_cstr(fmt)
    if fmt == '':
        return vb_cstr(v)
    grouping = ',' in fmt.split('.')[0]
    if '.' in fmt:
        dec = fmt.split('.')[1]
        req = dec.count('0')
        opt = dec.count('#')
        s = ('{:,.%df}' if grouping else '{:.%df}') % (req + opt)
        s = s.format(v)
        if opt:
            whole, frac = s.split('.')
            frac = frac.rstrip('0')
            if len(frac) < req:
                frac = frac + '0' * (req - len(frac))
            s = whole + ('.' + frac if frac else '')
        return s
    s = ('{:,.0f}' if grouping else '{:.0f}').format(_round_even(v) if isinstance(v, float) else v)
    return s


def vb_item(obj, *args):
    """obj(args): dizi elemani, varsayilan uye ya da Variant icindeki dizi."""
    if isinstance(obj, VArr):
        if len(args) != 1:
            raise VBError('Subscript count mismatch', 9)
        return obj[args[0]]
    if isinstance(obj, VArr2):
        return obj.item(*args)
    if obj is None:
        raise VBError('Object variable not set (91)', 91)
    f = getattr(obj, 'item', None)
    if f is None:
        raise VBError('Object does not support this property or method (438): item', 438)
    return f(*[V(a) for a in args])


def vb_callmem(obj, name, *args):
    """obj.name veya obj.name(args): yontem ise cagirir, ozellikse okur
    (ve argumanlar varsa donen degerin elemanini alir)."""
    if obj is None:
        raise VBError('Object variable not set (91): .%s' % name, 91)
    try:
        a = getattr(obj, name)
    except AttributeError:
        raise VBError('Object does not support this property or method (438): %s.%s'
                      % (type(obj).__name__, name), 438)
    if callable(a) and not isinstance(a, type):
        return a(*args)
    if args:
        return vb_item(a, *[V(x) for x in args])
    return a


def vb_setmem(obj, name, value):
    if obj is None:
        raise VBError('Object variable not set (91): .%s =' % name, 91)
    if not hasattr(obj, name):
        raise VBError('Object does not support this property or method (438): %s.%s ='
                      % (type(obj).__name__, name), 438)
    setattr(obj, name, value)


def vb_iter(x):
    if isinstance(x, VArr):
        return x.items()
    if x is None:
        raise VBError('Object variable not set (91) in For Each', 91)
    if hasattr(x, 'items_list'):
        return x.items_list()
    raise VBError('For Each over a non-collection', 451)


BUILTINS = {
    'abs': 'abs', 'int': 'vb_int', 'fix': 'vb_fix', 'sqr': 'vb_sqr',
    'sin': 'math.sin', 'cos': 'math.cos', 'tan': 'math.tan', 'atn': 'math.atan',
    'exp': 'math.exp', 'log': 'math.log',
    'clng': 'conv_long', 'cint': 'conv_integer', 'cbyte': 'conv_byte',
    'cdbl': 'conv_double', 'csng': 'conv_double', 'cbool': 'conv_boolean',
    'cstr': 'vb_cstr', 'iif': 'vb_iif', 'val': 'vb_val',
    'ubound': 'vb_ubound', 'lbound': 'vb_lbound', 'len': '(lambda s: len(vb_cstr(s)))',
    'mid': 'vb_mid', 'mid$': 'vb_mid', 'left$': 'vb_left', 'left': 'vb_left',
    'right$': 'vb_right', 'right': 'vb_right',
    'trim$': 'vb_trim', 'trim': 'vb_trim',
    'ucase$': '(lambda s: vb_cstr(s).upper())', 'lcase$': '(lambda s: vb_cstr(s).lower())',
    'ucase': '(lambda s: vb_cstr(s).upper())', 'lcase': '(lambda s: vb_cstr(s).lower())',
    'instr': 'vb_instr', 'replace$': 'vb_replace', 'replace': 'vb_replace',
    'split': 'vb_split', 'array': 'vb_array', 'strcomp': 'vb_strcomp',
    'isnumeric': 'vb_isnumeric', 'isarray': 'vb_isarray', 'strptr': 'vb_strptr',
    'ascw': 'vb_ascw', 'string$': 'vb_string', 'space$': 'vb_space', 'str$': 'vb_str',
    'format$': 'vb_format', 'format': 'vb_format',
}
BUILTIN_NOARG = {'timer': 'vb_timer()', 'true': 'True', 'false': 'False',
                 'vbcrlf': '"\\r\\n"', 'vblf': '"\\n"', 'vbcr': '"\\r"',
                 'vbtab': '"\\t"', 'nothing': 'None', 'vbnullstring': 'NullStr("")',
                 'vbtextcompare': '1', 'vbbinarycompare': '0',
                 'vbyes': '6', 'vbno': '7', 'vbok': '1', 'vbcancel': '2',
                 'vbquestion': '32', 'vbyesno': '4', 'vbinformation': '64',
                 'vbexclamation': '48', 'vbcritical': '16', 'vbokonly': '0',
                 'vbokcancel': '1', 'vbyesnocancel': '3', 'vbdefaultbutton2': '256',
                 'vbdirectory': '16', 'vbnormal': '0'}
# Ana bilgisayara (host) giden VBA islevleri: H.<ad>(...)
HOST_FUNCS = {'inputbox', 'msgbox', 'getsetting', 'savesetting', 'deletesetting',
              'getallsettings', 'environ$', 'environ', 'dir$', 'dir', 'createobject',
              'freefile', 'doevents', 'getattr', 'filelen'}

# =====================================================================
# source preprocessing
# =====================================================================

def strip_comment(ln):
    out, inq = [], False
    for ch in ln:
        if ch == '"':
            inq = not inq
        elif ch == "'" and not inq:
            break
        out.append(ch)
    s = ''.join(out)
    if re.match(r'^\s*Rem\b', s, re.I):
        return ''
    return s


def logical_lines(src):
    out, buf, anchor = [], None, None
    for i, ln in enumerate(src.split('\n'), 1):
        code = strip_comment(ln.rstrip('\r'))
        r0 = code.rstrip()
        cont = len(r0) > 1 and r0[-1] == '_' and r0[-2].isspace()
        body = r0[:-1] + ' ' if cont else code
        if buf is None:
            if cont:
                buf, anchor = body, i
            else:
                out.append((i, code))
        else:
            buf += body
            if not cont:
                out.append((anchor, buf))
                buf, anchor = None, None
    if buf is not None:
        out.append((anchor, buf))
    return out


def split_colons(code):
    parts, cur, inq, depth = [], [], False, 0
    for ch in code:
        if ch == '"':
            inq = not inq
        if not inq:
            if ch == '(':
                depth += 1
            elif ch == ')':
                depth -= 1
            elif ch == ':' and depth == 0:
                parts.append(''.join(cur))
                cur = []
                continue
        cur.append(ch)
    parts.append(''.join(cur))
    return [p.strip() for p in parts if p.strip()]


def split_commas(text):
    parts, cur, depth, inq = [], [], 0, False
    for ch in text:
        if ch == '"':
            inq = not inq
        if not inq:
            if ch == '(':
                depth += 1
            elif ch == ')':
                depth -= 1
            elif ch == ',' and depth == 0:
                parts.append(''.join(cur))
                cur = []
                continue
        cur.append(ch)
    if ''.join(cur).strip():
        parts.append(''.join(cur))
    return parts


# =====================================================================
# expressions
# =====================================================================
TOK_RE = re.compile(r'''
    (?P<ws>\s+)
  | (?P<num>&H[0-9A-Fa-f]+&?|\d+\.\d*(?:[eE][-+]?\d+)?[#!@]?|\.\d+(?:[eE][-+]?\d+)?[#!@]?|\d+(?:[eE][-+]?\d+)?[#!&%@]?)
  | (?P<str>"(?:[^"]|"")*")
  | (?P<id>[A-Za-z_][A-Za-z0-9_]*\$?)
  | (?P<op><>|<=|>=|:=|[-+*/\\^&=<>(),.\#])
''', re.X)

KEYWORD_OPS = {'and', 'or', 'not', 'xor', 'mod', 'is', 'to', 'eqv', 'imp'}


def tokenize(s):
    pos, out = 0, []
    while pos < len(s):
        m = TOK_RE.match(s, pos)
        if not m:
            raise VBError('cannot tokenize %r at %d' % (s, pos))
        pos = m.end()
        kind = m.lastgroup
        if kind == 'ws':
            continue
        val = m.group(kind)
        if kind == 'id' and val.lower() in KEYWORD_OPS:
            out.append(('op', val.lower()))
        else:
            out.append((kind, val))
    return out


BIN_BP = {'imp': 1, 'eqv': 2, 'xor': 3, 'or': 4, 'and': 5,
          '=': 7, '<>': 7, '<': 7, '>': 7, '<=': 7, '>=': 7, 'is': 7,
          '&': 8, '+': 9, '-': 9, 'mod': 10, '\\': 11, '*': 12, '/': 12, '^': 14}


class CompileError(Exception):
    pass


class ExprCtx(object):
    def __init__(self, comp, proc):
        self.comp = comp            # ModuleCompiler
        self.proc = proc
        self.pre = []

    def resolve(self, name):
        """-> (kind, pyref, type, extra)"""
        n = name.lower()
        p = self.proc
        c = self.comp
        if p is not None:
            if n == p.name and p.is_func:
                return 'ret', '__ret', p.ret_type, p
            if n in p.vars:
                t, isarr = p.vars[n]
                return ('arr' if isarr else 'var'), 'v_' + n, t, None
        mod = c.mod
        if n in mod.vars:
            t, isarr = mod.vars[n]
            return ('arr' if isarr else 'var'), 'M.v_' + n, t, None
        if n in mod.consts:
            return 'const', 'M.v_' + n, None, None
        if n in mod.procs:
            return 'proc', 'f_' + n, None, mod.procs[n]
        g = c.project.publics.get(n)
        if g is not None:
            kind, owner, info = g
            base = 'P.m.%s' % owner
            if kind == 'proc':
                return 'proc', '%s.f_%s' % (base, n), None, info
            if kind == 'const':
                return 'const', '%s.M.v_%s' % (base, n), None, None
            t, isarr = info
            return ('arr' if isarr else 'var'), '%s.M.v_%s' % (base, n), t, None
        if n in BUILTINS or n in BUILTIN_NOARG:
            return 'builtin', n, None, None
        if n in HOST_FUNCS:
            return 'hostfunc', 'H.' + re.sub(r'\$$', '', n), None, None
        if n == 'err':
            return 'host', 'H.err', None, None
        if c.project.host_names is not None and n in c.project.host_names:
            return 'host', 'H.' + n, None, None
        raise CompileError('Variable not defined: %r' % name)


class Parser(object):
    def __init__(self, toks, ctx):
        self.t = toks
        self.i = 0
        self.ctx = ctx

    def peek(self, k=0):
        j = self.i + k
        return self.t[j] if j < len(self.t) else (None, None)

    def take(self):
        tok = self.peek()
        self.i += 1
        return tok

    def expect(self, val):
        k, v = self.take()
        if v is None or v.lower() != val:
            raise CompileError('expected %r got %r' % (val, v))

    def at_end(self):
        return self.i >= len(self.t)

    def expr(self, rbp=0):
        left = self.prefix()
        while True:
            k, v = self.peek()
            if k != 'op':
                break
            op = v.lower()
            if op not in BIN_BP:
                break
            bp = BIN_BP[op]
            if bp <= rbp:
                break
            self.take()
            right = self.expr(bp)
            left = self.binop(op, left, right)
        return left

    def binop(self, op, a, b):
        m = {'+': 'vb_add', '-': 'vb_sub', '*': 'vb_mul', '/': 'vb_div', '\\': 'vb_idiv',
             'mod': 'vb_mod', '&': 'vb_cat', '=': 'vb_eq', '<>': 'vb_ne',
             'and': 'vb_and', 'or': 'vb_or', 'xor': 'vb_xor'}
        if op in m:
            return '%s(%s, %s)' % (m[op], a, b)
        if op == '^':
            return '(float(_num(%s)) ** _num(%s))' % (a, b)
        if op in ('<', '>', '<=', '>='):
            return 'vb_cmp(%s, %r, %s)' % (a, op, b)
        if op == 'is':
            return '(%s is %s)' % (a, b)
        raise CompileError('operator %r not supported' % op)

    def prefix(self):
        k, v = self.take()
        if k is None:
            raise CompileError('unexpected end of expression')
        if k == 'num':
            return self.postfix(self.number(v), None)
        if k == 'str':
            return repr(v[1:-1].replace('""', '"'))
        if k == 'op':
            if v == '(':
                e = self.expr()
                self.expect(')')
                return self.postfix('(%s)' % e, None)
            if v == '-':
                return 'vb_neg(%s)' % self.expr(13)
            if v == '+':
                return self.expr(13)
            if v == 'not':
                return 'vb_not(%s)' % self.expr(6)
            if v == '.':
                raise CompileError('With blocks are not supported')
            raise CompileError('unexpected operator %r' % v)
        if k == 'id':
            return self.name(v)
        raise CompileError('unexpected token %r' % v)

    def number(self, v):
        if v.lower().startswith('&h'):
            n = int(v[2:].rstrip('&'), 16)
            if n >= 0x8000 and not v.endswith('&') and n <= 0xFFFF:
                n -= 0x10000
            return str(n)
        s = v.rstrip('#!&%@')
        if v[-1] in '#!@' or '.' in s or 'e' in s.lower():
            return repr(float(s))
        return str(int(s))

    def args(self):
        out = []
        if self.peek()[1] == ')':
            self.take()
            return out
        while True:
            out.append(self.arg_expr())
            k, v = self.take()
            if v == ')':
                return out
            if v != ',':
                raise CompileError('expected , or ) in argument list, got %r' % v)

    def arg_expr(self):
        start = self.i
        if self.peek()[1] == ',' or self.peek()[1] == ')':
            return ('None', [])          # omitted optional argument
        e = self.expr()
        return (e, self.t[start:self.i])

    def postfix(self, base, kind):
        """.uye ve (argumanlar) zinciri."""
        while True:
            k, v = self.peek()
            if v == '.':
                self.take()
                k2, mem = self.take()
                if k2 != 'id' and k2 != 'op':
                    raise CompileError('member name expected after .')
                mem = mem.lower().rstrip('$')
                if self.peek()[1] == '(':
                    self.take()
                    a = self.args()
                    base = 'vb_callmem(%s, %r%s)' % (base, mem, ''.join(', Ref(%s)' % x[0] for x in a))
                else:
                    base = 'vb_callmem(%s, %r)' % (base, mem)
            elif v == '(' and kind == 'obj':
                self.take()
                a = self.args()
                base = 'vb_item(%s, %s)' % (base, ', '.join(x[0] for x in a))
            else:
                return base
            kind = 'obj'

    def name(self, v):
        n = v.lower()
        ctx = self.ctx
        kind, ref, t, info = ctx.resolve(n)
        has_paren = self.peek()[1] == '('
        if kind == 'arr':
            if has_paren:
                self.take()
                a = self.args()
                if len(a) != 1:
                    raise CompileError('only one-dimensional arrays are supported: %r' % v)
                return self.postfix('%s[%s]' % (ref, a[0][0]), 'obj')
            return self.postfix(ref, 'obj')
        if kind in ('var', 'const', 'ret'):
            if kind == 'ret' and has_paren:
                return self.postfix(self.call(n, info), 'obj')
            if has_paren:
                if t in ('variant', 'object'):
                    self.take()
                    a = self.args()
                    return self.postfix('vb_item(%s, %s)' % (ref, ', '.join(x[0] for x in a)), 'obj')
                raise CompileError('%r is not an array or function' % v)
            return self.postfix(ref, 'obj')
        if kind == 'proc':
            return self.postfix(self.call(n, info), 'obj')
        if kind == 'builtin':
            if n in BUILTIN_NOARG:
                if has_paren and n == 'timer':
                    self.take()
                    self.expect(')')
                return BUILTIN_NOARG[n]
            if not has_paren:
                raise CompileError('builtin %r needs arguments' % v)
            self.take()
            a = self.args()
            return '%s(%s)' % (BUILTINS[n], ', '.join(x[0] for x in a))
        if kind == 'hostfunc':
            a = []
            if has_paren:
                self.take()
                a = self.args()
            return '%s(%s)' % (ref, ', '.join(x[0] for x in a))
        if kind == 'host':
            return self.postfix(ref, 'obj')
        raise CompileError('cannot resolve %r' % v)

    def call(self, n, proc):
        a = []
        if self.peek()[1] == '(':
            self.take()
            a = self.args()
        return self.ctx.comp.emit_call(self.ctx, n, proc, a)


def parse_expr(text, ctx):
    p = Parser(tokenize(text), ctx)
    e = p.expr()
    if not p.at_end():
        raise CompileError('trailing tokens in expression %r' % text)
    return e


# =====================================================================
# module / procedure model
# =====================================================================
TYPE_RE = r'([A-Za-z_]\w*)'


class Proc(object):
    def __init__(self, name, is_func, params, ret_type, body, line, public):
        self.name = name
        self.is_func = is_func
        self.params = params        # (name, type, isarr, byref, default)
        self.ret_type = ret_type
        self.body = body
        self.line = line
        self.public = public
        self.vars = {}
        for p in params:
            self.vars[p[0]] = (p[1], p[2])

    @property
    def byref_scalars(self):
        return [i for i, p in enumerate(self.params) if p[3] and not p[2]]


class Module(object):
    def __init__(self, src, name):
        self.name = name.lower()
        self.vars = {}
        self.consts = {}
        self.const_order = []
        self.procs = {}
        self.public_vars = set()
        self.public_consts = set()
        self.parse(src)

    def parse_decl(self, text, target, line):
        out = []
        for part in split_commas(text):
            m = re.match(r'^(\w+)\s*(\(([^)]*)\))?\s*(?:As\s+(?:New\s+)?' + TYPE_RE + r')?\s*$', part.strip(), re.I)
            if not m:
                raise CompileError('line %d: cannot parse declaration %r' % (line, part))
            t = norm_type(m.group(4))
            name = m.group(1).lower()
            target[name] = (t, bool(m.group(2)))
            out.append((name, t, m.group(3)))
        return out

    def parse(self, src):
        lines = logical_lines(src)
        i = 0
        while i < len(lines):
            ln, code = lines[i]
            s = code.strip()
            i += 1
            if not s or re.match(r'^(Attribute|Option)\b', s, re.I):
                continue
            m = re.match(r'^(?:(Public|Private|Global)\s+)?Const\s+(\w+)\s*(?:As\s+' + TYPE_RE + r')?\s*=\s*(.+)$', s, re.I)
            if m:
                n = m.group(2).lower()
                self.consts[n] = (m.group(4), norm_type(m.group(3)))
                self.const_order.append(n)
                if (m.group(1) or '').lower() in ('public', 'global'):
                    self.public_consts.add(n)
                continue
            m = re.match(r'^(?:(Public|Private|Friend)\s+)?(?:Static\s+)?(Sub|Function)\s+(\w+)\s*\((.*)\)\s*(?:As\s+' + TYPE_RE + r'(\(\))?)?\s*$', s, re.I)
            if m:
                name = m.group(3).lower()
                is_func = m.group(2).lower() == 'function'
                params = self.parse_params(m.group(4), ln)
                ret = norm_type(m.group(5)) if is_func else None
                body = []
                end_re = re.compile(r'^End\s+' + m.group(2) + r'\s*$', re.I)
                while i < len(lines) and not end_re.match(lines[i][1].strip()):
                    body.append(lines[i])
                    i += 1
                if i >= len(lines):
                    raise CompileError('line %d: unterminated %s' % (ln, name))
                i += 1
                pub = (m.group(1) or 'public').lower() == 'public'
                self.procs[name] = Proc(name, is_func, params, ret, body, ln, pub)
                continue
            m = re.match(r'^(?:(Public|Private)\s+)?(Type|Enum)\s+(\w+)\s*$', s, re.I)
            if m:
                # user defined type: skipped; procedures that touch one fail to
                # compile and become stubs, the rest of the module still runs
                end_re = re.compile(r'^End\s+' + m.group(2) + r'\s*$', re.I)
                while i < len(lines) and not end_re.match(lines[i][1].strip()):
                    i += 1
                i += 1
                continue
            m = re.match(r'^(Public|Private|Dim|Global)\s+(.+)$', s, re.I)
            if m:
                if re.match(r'^(Declare)\b', m.group(2), re.I):
                    continue
                decl = self.parse_decl(m.group(2), self.vars, ln)
                if m.group(1).lower() in ('public', 'global'):
                    for d in decl:
                        self.public_vars.add(d[0])
                continue
            raise CompileError('line %d: unsupported module level statement %r' % (ln, s))

    def parse_params(self, text, line):
        out = []
        for part in split_commas(text):
            p = part.strip()
            m = re.match(r'^(Optional\s+)?(ByVal\s+|ByRef\s+)?(\w+)\s*(\(\s*\))?\s*(?:As\s+' + TYPE_RE + r')?\s*(?:=\s*(.+))?$', p, re.I)
            if not m:
                raise CompileError('line %d: cannot parse parameter %r' % (line, p))
            byref = not (m.group(2) and m.group(2).strip().lower() == 'byval')
            out.append((m.group(3).lower(), norm_type(m.group(5)), bool(m.group(4)), byref, m.group(6)))
        return out


# =====================================================================
# code generation
# =====================================================================
class ModuleCompiler(object):
    def __init__(self, project, mod):
        self.project = project
        self.mod = mod
        self.out = []
        self.warnings = []

    def emit(self, ind, text):
        self.out.append('    ' * ind + text)

    def compile(self):
        self.out = []
        ctx = ExprCtx(self, None)
        for n in self.mod.const_order:
            expr, t = self.mod.consts[n]
            e = parse_expr(expr, ctx)
            self.emit(0, 'M.v_%s = %s' % (n, e))
        for n, (t, isarr) in self.mod.vars.items():
            if isarr:
                self.emit(0, 'M.v_%s = VArr(%r)' % (n, t))
            else:
                self.emit(0, 'M.v_%s = %r' % (n, DEFAULT[t]))
        for p in self.mod.procs.values():
            start = len(self.out)
            try:
                self.compile_proc(p)
            except (CompileError, VBError) as ex:
                del self.out[start:]
                msg = '%s.%s: %s' % (self.mod.name, p.name, ex)
                self.warnings.append(msg)
                self.emit(0, '')
                self.emit(0, 'def f_%s(*a, **k):' % p.name)
                self.emit(1, 'raise NotImplementedError(%r)' % msg)
        return '\n'.join(self.out) + '\n'

    # ---- procedures ----
    def compile_proc(self, p):
        self.cur = p
        self.loops = []
        self.tmpn = 0
        p.vars = {}
        for pp in p.params:
            p.vars[pp[0]] = (pp[1], pp[2])
        fixed = {}
        body_lines = []
        for ln, code in p.body:
            for st in self.statements(code):
                body_lines.append((ln, st))
        for ln, st in body_lines:
            m = re.match(r'^Dim\s+(.+)$', st, re.I)
            if m:
                for (name, t, dims) in self.mod.parse_decl(m.group(1), p.vars, ln):
                    if dims is not None and dims.strip():
                        fixed[name] = dims
            m = re.match(r'^Const\s+(\w+)\s*(?:As\s+' + TYPE_RE + r')?\s*=\s*(.+)$', st, re.I)
            if m:
                p.vars[m.group(1).lower()] = (norm_type(m.group(2)), False)
        self.labels = {}
        self.has_eh = False
        for ln, st in body_lines:
            if re.match(r'^[A-Za-z_]\w*:$', st):
                self.labels[st[:-1].lower()] = len(self.labels) + 1
            if re.match(r'^(On\s+Error|GoTo|Resume)\b', st, re.I):
                self.has_eh = True
        if self.labels:
            self.has_eh = True

        args = []
        for (n, t, isarr, byref, dflt) in p.params:
            if dflt is not None:
                args.append('v_%s=%s' % (n, parse_expr(dflt, ExprCtx(self, None))))
            else:
                args.append('v_' + n)
        self.emit(0, '')
        self.emit(0, 'def f_%s(%s):' % (p.name, ', '.join(args)))
        for (n, t, isarr, byref, dflt) in p.params:
            if not isarr:
                self.emit(1, 'v_%s = conv_%s(V(v_%s))' % (n, t, n))
        if p.is_func:
            self.emit(1, '__ret = %r' % DEFAULT[p.ret_type])
        pnames = {x[0] for x in p.params}
        for n, (t, isarr) in p.vars.items():
            if n in pnames:
                continue
            if isarr:
                if n in fixed:
                    md = re.match(r'^(.+?)\s+To\s+(.+)$', fixed[n].strip(), re.I)
                    lo, hi = (md.group(1), md.group(2)) if md else ('0', fixed[n])
                    ctx = ExprCtx(self, p)
                    self.emit(1, 'v_%s = VArr(%r, %s, %s)' % (n, t, parse_expr(lo, ctx), parse_expr(hi, ctx)))
                else:
                    self.emit(1, 'v_%s = VArr(%r)' % (n, t))
            else:
                self.emit(1, 'v_%s = %r' % (n, DEFAULT[t]))
        if not self.has_eh:
            self.compile_block(body_lines, 0, len(body_lines), 1)
            self.emit(1, self.ret_stmt())
            return
        # error handling: segments between labels, run by a dispatcher
        self.emit(1, '__eh = None')
        self.emit(1, '__inh = False')
        self.emit(1, '__pc = 0')
        self.emit(1, 'while True:')
        self.emit(2, 'try:')
        segs, cur, start = [], 0, 0
        for idx, (ln, st) in enumerate(body_lines):
            if re.match(r'^[A-Za-z_]\w*:$', st):
                segs.append((cur, start, idx))
                cur = self.labels[st[:-1].lower()]
                start = idx + 1
        segs.append((cur, start, len(body_lines)))
        for (num, a, b) in segs:
            self.emit(3, 'if __pc <= %d:' % num)
            self.compile_block(body_lines, a, b, 4)
        self.emit(3, 'break')
        self.emit(2, 'except _Goto as __g:')
        self.emit(3, '__pc = %r[__g.label]' % self.labels)
        self.emit(2, 'except VBError as __e:')
        self.emit(3, 'H.err.set(__e)')
        self.emit(3, 'if __eh is None or __eh == "resume" or __inh: raise')
        self.emit(3, '__pc = %r[__eh]' % self.labels)
        self.emit(3, '__inh = True')
        self.emit(1, self.ret_stmt())

    def ret_stmt(self):
        p = self.cur
        br = p.byref_scalars
        r = '__ret' if p.is_func else 'None'
        if br:
            return 'return (%s, %s)' % (r, ', '.join('v_' + p.params[i][0] for i in br))
        return 'return %s' % r

    def statements(self, code):
        s = code.strip()
        if not s or re.match(r'^Attribute\s', s, re.I):
            return []
        if re.match(r'^If\b', s, re.I) and not re.search(r'\bThen\s*$', s, re.I):
            return [s]
        # a label "Name:" alone, or followed by a statement
        m = re.match(r'^([A-Za-z_]\w*):(?!=)\s*(.*)$', s)
        if m and not re.match(r'^(Case|Else)\b', s, re.I) and m.group(1).lower() not in ('case',):
            rest = m.group(2)
            if not rest or not re.match(r'^[^"]*=', m.group(1)):
                out = [m.group(1) + ':']
                if rest:
                    out += split_colons(rest)
                return out
        return split_colons(s)

    def new_ctx(self):
        return ExprCtx(self, self.cur)

    def expr(self, text, ind, line):
        ctx = self.new_ctx()
        try:
            e = parse_expr(text, ctx)
        except (CompileError, VBError) as ex:
            raise CompileError('line %d: %s' % (line, ex))
        for st in ctx.pre:
            self.emit(ind, st)
        return e

    def emit_call(self, ctx, name, proc, a):
        if len(a) > len(proc.params):
            raise CompileError('too many arguments to %s' % name)
        args = []
        for i, (e, toks) in enumerate(a):
            pr = proc.params[i]
            if e == 'None' and not toks:
                if pr[4] is None:
                    raise CompileError('argument %d of %s is not optional' % (i + 1, name))
                args.append(parse_expr(pr[4], ExprCtx(self, None)))
            else:
                args.append(e)
        k0, ref, _, _ = ctx.resolve(name)
        if k0 == 'ret':
            ref = 'f_' + name          # a function calling itself
        call = '%s(%s)' % (ref, ', '.join(args))
        br = proc.byref_scalars
        if not br:
            return call
        self.tmpn += 1
        t = '__t%d' % self.tmpn
        ctx.pre.append('%s = %s' % (t, call))
        for j, pi in enumerate(br):
            if pi >= len(a):
                continue
            target = self.lvalue_from_tokens(a[pi][1], ctx)
            if target is not None:
                ctx.pre.append('%s = %s[%d]' % (target, t, j + 1))
        return '%s[0]' % t

    def lvalue_from_tokens(self, toks, ctx):
        if len(toks) == 1 and toks[0][0] == 'id':
            k, ref, t, _ = ctx.resolve(toks[0][1])
            if k in ('var', 'ret'):
                return ref
            return None
        if len(toks) >= 4 and toks[0][0] == 'id' and toks[1][1] == '(' and toks[-1][1] == ')':
            k, ref, t, _ = ctx.resolve(toks[0][1])
            if k == 'arr':
                depth = 0
                for j, tk in enumerate(toks[1:], 1):
                    if tk[1] == '(':
                        depth += 1
                    elif tk[1] == ')':
                        depth -= 1
                        if depth == 0 and j != len(toks) - 1:
                            return None
                sub = Parser(toks[2:-1], ctx)
                return '%s[%s]' % (ref, sub.expr())
        return None

    # ---- statements ----
    def wrap(self, ind):
        """Hata yakalanan yordamlarda her yalin deyim kendi try'inda."""
        if not self.has_eh:
            return ind
        self.emit(ind, 'try:')
        return ind + 1

    def unwrap(self, ind):
        if not self.has_eh:
            return
        self.emit(ind - 1, 'except VBError as __e:')
        self.emit(ind, 'if __eh != "resume": raise')
        self.emit(ind, 'H.err.set(__e)')

    def assign(self, lhs, rhs, ind, line, is_set=False):
        ctx = self.new_ctx()
        lt = tokenize(lhs)
        k, ref, t, _ = ctx.resolve(lt[0][1])
        rt = tokenize(rhs)
        if k == 'arr' and len(lt) == 1:
            if len(rt) == 1 and rt[0][0] == 'id':
                k2, ref2, _, _ = ctx.resolve(rt[0][1])
                if k2 == 'arr':
                    self.emit(ind, '%s.assign(%s)' % (ref, ref2))
                    return
            e = self.expr(rhs, ind, line)
            self.emit(ind, '%s.assign(%s)' % (ref, e))
            return
        e = self.expr(rhs, ind, line)
        # member target: a.b.c = x  or  arr(i).Name = x
        dot = None
        depth = 0
        for j, tk in enumerate(lt):
            if tk[1] == '(':
                depth += 1
            elif tk[1] == ')':
                depth -= 1
            elif tk[1] == '.' and depth == 0:
                dot = j
        if dot is not None:
            base_toks = lt[:dot]
            mem = lt[dot + 1][1].lower()
            if len(lt) != dot + 2:
                raise CompileError('line %d: unsupported assignment target %r' % (line, lhs))
            p = Parser(base_toks, ctx)
            b = p.expr()
            for st in ctx.pre:
                self.emit(ind, st)
            self.emit(ind, 'vb_setmem(%s, %r, %s)' % (b, mem, e))
            return
        if k == 'arr':
            sub = Parser(lt[2:-1], ctx)
            ix = sub.expr()
            for st in ctx.pre:
                self.emit(ind, st)
            self.emit(ind, '%s[%s] = %s' % (ref, ix, e))
        elif k in ('var', 'ret'):
            if len(lt) != 1:
                raise CompileError('line %d: bad assignment target %r' % (line, lhs))
            tt = t if k != 'ret' else self.cur.ret_type
            if tt == 'object' and not is_set and rt and rt[0][1].lower() != 'nothing':
                pass
            self.emit(ind, '%s = conv_%s(%s)' % (ref, tt, e))
        else:
            raise CompileError('line %d: cannot assign to %r' % (line, lhs))

    def call_stmt(self, st, ind, line):
        m = re.match(r'^Call\s+(.+)$', st, re.I)
        if m:
            st = m.group(1)
        toks = tokenize(st)
        ctx = self.new_ctx()
        if not toks or toks[0][0] != 'id':
            raise CompileError('line %d: unsupported statement %r' % (line, st))
        n = toks[0][1].lower()
        kind, ref, t, info = ctx.resolve(n)
        if kind == 'proc' and (len(toks) == 1 or toks[1][1] != '.'):
            # VBA procedure: "Name a, b" or "Name(a, b)"
            rest = toks[1:]
            if rest and rest[0][1] == '(' and rest[-1][1] == ')':
                depth = 0
                whole = True
                for j, tk in enumerate(rest):
                    if tk[1] == '(':
                        depth += 1
                    elif tk[1] == ')':
                        depth -= 1
                        if depth == 0 and j != len(rest) - 1:
                            whole = False
                if whole and len(info.params) != 1:
                    rest = rest[1:-1]
            a = []
            if rest:
                p = Parser(rest, ctx)
                while True:
                    a.append(p.arg_expr())
                    if p.at_end():
                        break
                    if p.take()[1] != ',':
                        raise CompileError('line %d: bad argument list in %r' % (line, st))
            call = self.emit_call(ctx, n, info, a)
            for s in ctx.pre:
                self.emit(ind, s)
            if not ctx.pre:
                self.emit(ind, call)
            return
        if kind == 'host' and (len(toks) == 1 or toks[1][1] != '.'):
            # a host-provided procedure (a stub for a module not loaded)
            a = []
            if len(toks) > 1:
                pp = Parser(toks[1:], ctx)
                while True:
                    a.append(pp.arg_expr())
                    if pp.at_end():
                        break
                    pp.take()
            for s in ctx.pre:
                self.emit(ind, s)
            self.emit(ind, '%s(%s)' % (ref, ', '.join(x[0] for x in a)))
            return
        # member call: chain then optional unparenthesised arguments
        p = Parser(toks, ctx)
        k0, v0 = p.take()
        base_kind, base_ref, _, _ = ctx.resolve(v0)
        if base_kind == 'hostfunc':
            a = []
            rest = toks[1:]
            if rest:
                pp = Parser(rest, ctx)
                if rest[0][1] == '(' and rest[-1][1] == ')':
                    pp.take()
                    a = pp.args()
                else:
                    while True:
                        a.append(pp.arg_expr())
                        if pp.at_end():
                            break
                        pp.take()
            for s in ctx.pre:
                self.emit(ind, s)
            self.emit(ind, '%s(%s)' % (base_ref, ', '.join(x[0] for x in a)))
            return
        if base_kind == 'arr':
            if p.peek()[1] == '(':
                p.take()
                a = p.args()
                base = '%s[%s]' % (base_ref, a[0][0])
            else:
                base = base_ref
        elif base_kind in ('var', 'host', 'const', 'ret'):
            base = base_ref
            if p.peek()[1] == '(' and t in ('variant', 'object', None):
                p.take()
                a = p.args()
                base = 'vb_item(%s, %s)' % (base, ', '.join(x[0] for x in a))
        elif base_kind == 'proc':
            base = p.call(v0, info) if False else self.emit_call(ctx, v0.lower(), info, [])
        else:
            raise CompileError('line %d: unsupported statement %r' % (line, st))
        # walk the chain up to the last member
        chain = []
        while p.peek()[1] == '.':
            p.take()
            mem = p.take()[1].lower().rstrip('$')
            pargs = None
            if p.peek()[1] == '(':
                save = p.i
                p.take()
                pargs = p.args()
                # "obj.Method (a), b": parenthesised first argument
                if not p.at_end() and p.peek()[1] == ',':
                    p.i = save
                    pargs = None
            chain.append((mem, pargs))
        if not chain:
            raise CompileError('line %d: unsupported statement %r' % (line, st))
        for mem, pargs in chain[:-1]:
            if pargs is None:
                base = 'vb_callmem(%s, %r)' % (base, mem)
            else:
                base = 'vb_callmem(%s, %r%s)' % (base, mem, ''.join(', Ref(%s)' % x[0] for x in pargs))
        mem, pargs = chain[-1]
        a = pargs
        if a is None:
            a = []
            if not p.at_end():
                while True:
                    a.append(p.arg_expr())
                    if p.at_end():
                        break
                    if p.take()[1] != ',':
                        raise CompileError('line %d: bad argument list in %r' % (line, st))
        elif not p.at_end():
            raise CompileError('line %d: trailing tokens in %r' % (line, st))
        for s in ctx.pre:
            self.emit(ind, s)
        self.tmpn += 1
        rv = '__r%d' % self.tmpn
        self.emit(ind, '%s = [%s]' % (rv, ', '.join('Ref(%s)' % x[0] for x in a)))
        self.emit(ind, 'vb_callmem(%s, %r, *%s)' % (base, mem, rv))
        for j, (e, tk) in enumerate(a):
            target = self.lvalue_from_tokens(tk, ctx)
            if target is not None:
                k1 = ctx.resolve(tk[0][1])
                tt = k1[2] if k1[0] == 'var' else None
                if tt and tt != 'object':
                    self.emit(ind, '%s = conv_%s(%s[%d].v)' % (target, tt, rv, j))
                else:
                    self.emit(ind, '%s = %s[%d].v' % (target, rv, j))

    def find_assign(self, st):
        depth, inq = 0, False
        for i, ch in enumerate(st):
            if ch == '"':
                inq = not inq
            if inq:
                continue
            if ch == '(':
                depth += 1
            elif ch == ')':
                depth -= 1
            elif ch == '=' and depth == 0:
                prev = st[i - 1] if i else ''
                nxt = st[i + 1] if i + 1 < len(st) else ''
                if prev in '<>:' or nxt in '<>=':
                    continue
                return i
        return -1

    def compile_block(self, L, a, b, ind):
        i = a
        start_len = len(self.out)
        while i < b:
            ln, st = L[i]
            if re.match(r'^[A-Za-z_]\w*:$', st):
                i += 1
                continue
            m = re.match(r'^If\s+(.+?)\s+Then\s*(.*)$', st, re.I)
            if m and m.group(2):
                self.single_if(m.group(1), m.group(2), ind, ln)
                i += 1
                continue
            if m:
                i = self.block_if(L, i, b, ind)
                continue
            m = re.match(r'^For\s+Each\s+(\w+)\s+In\s+(.+)$', st, re.I)
            if m:
                j = self.find_end(L, i, b, r'^For\b', r'^Next\b')
                ctx = self.new_ctx()
                ref = ctx.resolve(m.group(1))[1]
                coll = self.expr(m.group(2), ind, ln)
                self.emit(ind, 'for %s in vb_iter(%s):' % (ref, coll))
                self.loops.append('for')
                self.compile_block(L, i + 1, j, ind + 1)
                self.loops.pop()
                i = j + 1
                continue
            m = re.match(r'^For\s+(\w+)\s*=\s*(.+?)\s+To\s+(.+?)(?:\s+Step\s+(.+))?$', st, re.I)
            if m:
                j = self.find_end(L, i, b, r'^For\b', r'^Next\b')
                self.for_loop(m, L, i, j, ind, ln)
                i = j + 1
                continue
            if re.match(r'^Do\b', st, re.I):
                j = self.find_end(L, i, b, r'^Do\b', r'^Loop\b')
                self.do_loop(L, i, j, ind)
                i = j + 1
                continue
            m = re.match(r'^Select\s+Case\s+(.+)$', st, re.I)
            if m:
                j = self.find_end(L, i, b, r'^Select\s+Case\b', r'^End\s+Select\b')
                self.select_case(m.group(1), L, i, j, ind, ln)
                i = j + 1
                continue
            self.simple(st, ind, ln)
            i += 1
        if len(self.out) == start_len:
            self.emit(ind, 'pass')

    def find_end(self, L, i, b, open_re, close_re):
        depth = 0
        for j in range(i, b):
            s = L[j][1]
            if re.match(open_re, s, re.I):
                depth += 1
            elif re.match(close_re, s, re.I):
                depth -= 1
                if depth == 0:
                    return j
        raise CompileError('line %d: unterminated block' % L[i][0])

    def simple(self, st, ind, ln):
        if re.match(r'^Dim\b', st, re.I):
            return
        m = re.match(r'^Const\s+(\w+)\s*(?:As\s+' + TYPE_RE + r')?\s*=\s*(.+)$', st, re.I)
        if m:
            e = self.expr(m.group(3), ind, ln)
            self.emit(ind, 'v_%s = %s' % (m.group(1).lower(), e))
            return
        m = re.match(r'^On\s+Error\s+(Resume\s+Next|GoTo\s+(\w+|-?\d+))\s*$', st, re.I)
        if m:
            if m.group(1).lower().startswith('resume'):
                self.emit(ind, '__eh = "resume"')
            elif m.group(2) in ('0', '-1'):
                self.emit(ind, '__eh = None')
                if m.group(2) == '-1':
                    self.emit(ind, '__inh = False')
            else:
                lab = m.group(2).lower()
                if lab not in self.labels:
                    raise CompileError('line %d: unknown label %r' % (ln, lab))
                self.emit(ind, '__eh = %r' % lab)
            return
        m = re.match(r'^GoTo\s+(\w+)$', st, re.I)
        if m:
            lab = m.group(1).lower()
            if lab not in self.labels:
                raise CompileError('line %d: unknown label %r' % (ln, lab))
            self.emit(ind, 'raise _Goto(%r)' % lab)
            return
        if re.match(r'^Resume\b', st, re.I):
            raise CompileError('line %d: Resume is not supported' % ln)
        m = re.match(r'^ReDim\s+(Preserve\s+)?(.+)$', st, re.I)
        if m:
            pres = bool(m.group(1))
            ind2 = self.wrap(ind)
            for part in split_commas(m.group(2)):
                mm = re.match(r'^\s*(\w+)\s*\((.+)\)\s*(?:As\s+' + TYPE_RE + r')?\s*$', part, re.I)
                if not mm:
                    raise CompileError('line %d: cannot parse ReDim %r' % (ln, part))
                dims = mm.group(2)
                if len(split_commas(dims)) > 1:
                    raise CompileError('line %d: multi-dimensional ReDim not supported' % ln)
                md = re.match(r'^(.+?)\s+To\s+(.+)$', dims, re.I)
                lo, hi = (md.group(1), md.group(2)) if md else ('0', dims)
                ctx = self.new_ctx()
                ref = ctx.resolve(mm.group(1))[1]
                elo = self.expr(lo, ind2, ln)
                ehi = self.expr(hi, ind2, ln)
                self.emit(ind2, '%s.redim(%s, %s, %s)' % (ref, elo, ehi, pres))
            self.unwrap(ind2)
            return
        if re.match(r'^Exit\s+For$', st, re.I):
            if not self.loops or self.loops[-1] != 'for':
                raise CompileError('line %d: Exit For not directly inside a For loop' % ln)
            self.emit(ind, 'break')
            return
        if re.match(r'^Exit\s+Do$', st, re.I):
            if not self.loops or self.loops[-1] != 'do':
                raise CompileError('line %d: Exit Do not directly inside a Do loop' % ln)
            self.emit(ind, 'break')
            return
        if re.match(r'^Exit\s+(Function|Sub)$', st, re.I):
            self.emit(ind, self.ret_stmt())
            return
        m = re.match(r'^Open\s+(.+?)\s+For\s+(Output|Input|Append|Binary)\s+As\s+#?(.+)$', st, re.I)
        if m:
            ind2 = self.wrap(ind)
            path = self.expr(m.group(1), ind2, ln)
            fn = self.expr(m.group(3), ind2, ln)
            self.emit(ind2, 'H.file_open(%s, %r, %s)' % (path, m.group(2).lower(), fn))
            self.unwrap(ind2)
            return
        m = re.match(r'^Print\s+#(\w+)\s*,\s*(.*)$', st, re.I)
        if m:
            ind2 = self.wrap(ind)
            fn = self.expr(m.group(1), ind2, ln)
            e = self.expr(m.group(2), ind2, ln) if m.group(2).strip() else "''"
            self.emit(ind2, 'H.file_print(%s, %s)' % (fn, e))
            self.unwrap(ind2)
            return
        m = re.match(r'^Close(?:\s+#?(.+))?$', st, re.I)
        if m:
            ind2 = self.wrap(ind)
            fn = self.expr(m.group(1), ind2, ln) if m.group(1) else 'None'
            self.emit(ind2, 'H.file_close(%s)' % fn)
            self.unwrap(ind2)
            return
        m = re.match(r'^Kill\s+(.+)$', st, re.I)
        if m:
            ind2 = self.wrap(ind)
            self.emit(ind2, 'H.kill(%s)' % self.expr(m.group(1), ind2, ln))
            self.unwrap(ind2)
            return
        if re.match(r'^DoEvents$', st, re.I):
            self.emit(ind, 'H.doevents()')
            return
        if re.match(r'^(GoSub|With|End\s+With|Erase|Debug\.)\b', st, re.I):
            raise CompileError('line %d: unsupported statement %r' % (ln, st))
        m = re.match(r'^Set\s+(.+)$', st, re.I)
        is_set = bool(m)
        if m:
            st = m.group(1)
        k = self.find_assign(st)
        ind2 = self.wrap(ind)
        if k > 0:
            self.assign(st[:k].strip(), st[k + 1:].strip(), ind2, ln, is_set)
        else:
            if is_set:
                raise CompileError('line %d: bad Set statement' % ln)
            self.call_stmt(st, ind2, ln)
        self.unwrap(ind2)

    def single_if(self, cond, rest, ind, ln):
        m = re.match(r'^(.*?)\s+Else\s+(.*)$', rest, re.I)
        then_part, else_part = (m.group(1), m.group(2)) if m else (rest, None)
        c = self.expr(cond, ind, ln)
        self.emit(ind, 'if conv_boolean(%s):' % c)
        L = [(ln, s) for s in split_colons(then_part)]
        self.compile_block(L, 0, len(L), ind + 1)
        if else_part is not None:
            self.emit(ind, 'else:')
            L = [(ln, s) for s in split_colons(else_part)]
            self.compile_block(L, 0, len(L), ind + 1)

    def block_if(self, L, i, b, ind):
        branches = []
        depth = 0
        cur_cond = re.match(r'^If\s+(.+?)\s+Then\s*$', L[i][1], re.I).group(1)
        cur_start = i + 1
        cur_line = L[i][0]
        j = i + 1
        while j < b:
            s = L[j][1]
            if re.match(r'^If\b.*\bThen\s*$', s, re.I):
                depth += 1
            elif re.match(r'^End\s+If\b', s, re.I):
                if depth == 0:
                    branches.append((cur_cond, cur_start, j, cur_line))
                    break
                depth -= 1
            elif depth == 0 and re.match(r'^ElseIf\b', s, re.I):
                branches.append((cur_cond, cur_start, j, cur_line))
                cur_cond = re.match(r'^ElseIf\s+(.+?)\s+Then\s*$', s, re.I).group(1)
                cur_start = j + 1
                cur_line = L[j][0]
            elif depth == 0 and re.match(r'^Else\s*$', s, re.I):
                branches.append((cur_cond, cur_start, j, cur_line))
                cur_cond = None
                cur_start = j + 1
                cur_line = L[j][0]
            j += 1
        else:
            raise CompileError('line %d: unterminated If' % L[i][0])
        first = True
        cur = ind
        for (c, s0, e0, ln) in branches:
            if c is None:
                self.emit(cur, 'else:')
            else:
                ctx = self.new_ctx()
                try:
                    ce = parse_expr(c, ctx)
                except (CompileError, VBError) as ex:
                    raise CompileError('line %d: %s' % (ln, ex))
                if ctx.pre and not first:
                    # an ElseIf whose test calls with ByRef: else + nested if
                    self.emit(cur, 'else:')
                    cur += 1
                    first = True
                for st in ctx.pre:
                    self.emit(cur, st)
                self.emit(cur, '%s conv_boolean(%s):' % ('if' if first else 'elif', ce))
            self.compile_block(L, s0, e0, cur + 1)
            first = False
        return j + 1

    def for_loop(self, m, L, i, j, ind, ln):
        var = m.group(1)
        ctx = self.new_ctx()
        k, ref, t, _ = ctx.resolve(var)
        a = self.expr(m.group(2), ind, ln)
        bnd = self.expr(m.group(3), ind, ln)
        stp = self.expr(m.group(4), ind, ln) if m.group(4) else '1'
        self.tmpn += 1
        e, s = '__e%d' % self.tmpn, '__s%d' % self.tmpn
        self.emit(ind, '%s = conv_%s(%s)' % (ref, t, a))
        self.emit(ind, '%s = _num(%s); %s = _num(%s)' % (e, bnd, s, stp))
        self.emit(ind, 'while (%s <= %s) if %s >= 0 else (%s >= %s):' % (ref, e, s, ref, e))
        self.loops.append('for')
        self.compile_block(L, i + 1, j, ind + 1)
        self.loops.pop()
        self.emit(ind + 1, '%s = conv_%s(%s + %s)' % (ref, t, ref, s))

    def do_loop(self, L, i, j, ind):
        head = L[i][1]
        tail = L[j][1]
        mh = re.match(r'^Do(?:\s+(While|Until)\s+(.+))?\s*$', head, re.I)
        mt = re.match(r'^Loop(?:\s+(While|Until)\s+(.+))?\s*$', tail, re.I)
        if not mh or not mt:
            raise CompileError('line %d: cannot parse Do/Loop' % L[i][0])
        self.emit(ind, 'while True:')
        if mh.group(1):
            ctx = self.new_ctx()
            c = parse_expr(mh.group(2), ctx)
            for st in ctx.pre:
                self.emit(ind + 1, st)
            neg = mh.group(1).lower() == 'while'
            self.emit(ind + 1, 'if %sconv_boolean(%s): break' % ('not ' if neg else '', c))
        self.loops.append('do')
        self.compile_block(L, i + 1, j, ind + 1)
        self.loops.pop()
        if mt.group(1):
            ctx = self.new_ctx()
            c = parse_expr(mt.group(2), ctx)
            for st in ctx.pre:
                self.emit(ind + 1, st)
            neg = mt.group(1).lower() == 'while'
            self.emit(ind + 1, 'if %sconv_boolean(%s): break' % ('not ' if neg else '', c))

    def select_case(self, sel, L, i, j, ind, ln):
        self.tmpn += 1
        sv = '__sel%d' % self.tmpn
        e = self.expr(sel, ind, ln)
        self.emit(ind, '%s = %s' % (sv, e))
        cases = []
        k = i + 1
        depth = 0
        cur = None
        while k < j:
            s = L[k][1]
            if re.match(r'^Select\s+Case\b', s, re.I):
                depth += 1
            elif re.match(r'^End\s+Select\b', s, re.I):
                depth -= 1
            elif depth == 0 and re.match(r'^Case\b', s, re.I):
                if cur:
                    cases.append(cur + [k])
                cur = [s[4:].strip(), k + 1, L[k][0]]
            k += 1
        if cur:
            cases.append(cur + [j])
        first = True
        any_branch = False
        for (spec, s0, cl, e0) in cases:
            if spec.lower() == 'else':
                if first:
                    self.emit(ind, 'if True:')
                else:
                    self.emit(ind, 'else:')
            else:
                conds = []
                ctx = self.new_ctx()
                for part in split_commas(spec):
                    part = part.strip()
                    mm = re.match(r'^Is\s*(<>|<=|>=|<|>|=)\s*(.+)$', part, re.I)
                    mt = re.match(r'^(.+?)\s+To\s+(.+)$', part, re.I)
                    if mm:
                        op = mm.group(1)
                        rhs = parse_expr(mm.group(2), ctx)
                        if op == '=':
                            conds.append('vb_eq(%s, %s)' % (sv, rhs))
                        elif op == '<>':
                            conds.append('vb_ne(%s, %s)' % (sv, rhs))
                        else:
                            conds.append('vb_cmp(%s, %r, %s)' % (sv, op, rhs))
                    elif mt:
                        conds.append('(vb_cmp(%s, ">=", %s) and vb_cmp(%s, "<=", %s))'
                                     % (sv, parse_expr(mt.group(1), ctx), sv, parse_expr(mt.group(2), ctx)))
                    else:
                        conds.append('vb_eq(%s, %s)' % (sv, parse_expr(part, ctx)))
                if ctx.pre:
                    raise CompileError('line %d: ByRef call in Case' % cl)
                self.emit(ind, '%s %s:' % ('if' if first else 'elif', ' or '.join(conds)))
            self.compile_block(L, s0, e0, ind + 1)
            first = False
            any_branch = True


# =====================================================================
# project
# =====================================================================
class _NS(object):
    pass


class Project(object):
    def __init__(self, paths, host=None):
        self.host = host
        self.host_names = None if host is None else set(host.names())
        self.mods = []
        for p in paths:
            name = re.sub(r'\.bas$', '', p.replace('\\', '/').split('/')[-1])
            self.mods.append(Module(open(p, encoding='utf-8').read(), name))
        self.publics = {}
        for md in self.mods:
            for n, pr in md.procs.items():
                if pr.public:
                    self.publics[n] = ('proc', md.name, pr)
            for n in md.public_vars:
                self.publics[n] = ('var', md.name, md.vars[n])
            for n in md.public_consts:
                self.publics[n] = ('const', md.name, None)
        self.m = _NS()
        self.sources = {}
        self.warnings = []
        compiled = []
        for md in self.mods:
            mc = ModuleCompiler(self, md)
            src = mc.compile()
            self.warnings += mc.warnings
            self.sources[md.name] = src
            compiled.append((md, src))
        # module state first, so constants that refer to other modules work
        for md, src in compiled:
            ns = types.ModuleType('vba_' + md.name)
            ns.__dict__.update(_RUNTIME)
            ns.M = _NS()
            ns.P = self
            ns.H = host
            setattr(self.m, md.name, ns)
        pending = list(compiled)
        for _ in range(len(pending) + 1):
            still = []
            for md, src in pending:
                ns = getattr(self.m, md.name)
                try:
                    exec(compile(src, md.name + '.bas.py', 'exec'), ns.__dict__)
                except AttributeError:
                    still.append((md, src))
            if not still:
                break
            pending = still
        else:
            raise CompileError('could not initialise modules (cyclic constants?)')


_RUNTIME = {k: v for k, v in globals().items()
            if k.startswith('vb_') or k.startswith('conv_') or k in (
                'VArr', 'VArr2', 'VBError', '_Goto', 'NullStr', 'Ref', 'V', '_num', 'math')}


def load_module(path, name=None):
    p = Project([path])
    if p.warnings:
        raise CompileError('\n'.join(p.warnings))
    return getattr(p.m, p.mods[0].name)


def compile_source(src):
    md = Module(src, 'mod')

    class _P(object):
        publics = {}
        host_names = None
    mc = ModuleCompiler(_P(), md)
    out = mc.compile()
    if mc.warnings:
        raise CompileError('\n'.join(mc.warnings))
    return out


if __name__ == '__main__':
    import sys
    sys.stdout.write(compile_source(open(sys.argv[1], encoding='utf-8').read()))
