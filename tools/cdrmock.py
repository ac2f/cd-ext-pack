#!/usr/bin/env python3
"""ac2f pack - CorelDRAW nesne modelinin sinamalar icin kucuk taklidi.

tools/vbapy.py ile Python'a cevrilen modullere "ana bilgisayar" (host)
olarak verilir. Yalniz paketin kullandigi yuzey vardir: Document, Page,
Layer, Shape, ShapeRange, Curve, SubPath, Segment, Node; kayit defteri,
dosyalar, InputBox / MsgBox, WScript.Shell ve ADODB.Stream.

Geometri gercektir: egriler kubik bezier, dondurme ve aynalama sinirlayici
kutunun merkezine gore (CorelDRAW'in varsayilani), GetBoundingBox egrinin
kendisinden olculur. Paket kodunun belgeye gercekten ne yaptigi boylece
sinanabilir.

Bilinen varsayim: Node.BreakApart + StartNode.JoinWith EndNode, kirilan
dugumu yeni baslangic yapar (CorelDRAW topluluk makrolarinin dayandigi
davranis). SubPath.ReverseDirection istenirse kapatilabilir
(Host.reverse_supported = False).
"""
import copy, fnmatch, math, os, re

from vbapy import V, VBError, NullStr, VArr2

# ---------------------------------------------------------------------
# constants
# ---------------------------------------------------------------------
CONSTS = {
    'cdrrectangleshape': 1, 'cdrellipseshape': 2, 'cdrcurveshape': 3,
    'cdrpolygonshape': 4, 'cdrbitmapshape': 5, 'cdrtextshape': 6,
    'cdrgroupshape': 7, 'cdrolecobjectshape': 12, 'cdroleobjectshape': 12,
    'cdrmillimeter': 3, 'cdrinch': 1,
    'cdrfliphorizontal': 1, 'cdrflipvertical': 2, 'cdrflipboth': 3,
    'cdrrelativesegmentoffset': 1, 'cdrabsolutesegmentoffset': 2,
    'cdrlinesegment': 0, 'cdrcurvesegment': 1,
}


# ---------------------------------------------------------------------
# geometry
# ---------------------------------------------------------------------
def bez(p0, c1, c2, p1, t):
    u = 1 - t
    return (u * u * u * p0[0] + 3 * u * u * t * c1[0] + 3 * u * t * t * c2[0] + t * t * t * p1[0],
            u * u * u * p0[1] + 3 * u * u * t * c1[1] + 3 * u * t * t * c2[1] + t * t * t * p1[1])


class Path(object):
    """Bir alt yol: dugumler ve her segmentin kontrol noktalari (None = cizgi)."""

    def __init__(self, nodes, ctrls=None, closed=True):
        self.nodes = [tuple(map(float, p)) for p in nodes]
        n = len(self.nodes)
        nseg = n if closed else n - 1
        self.ctrls = list(ctrls) if ctrls is not None else [None] * nseg
        while len(self.ctrls) < nseg:
            self.ctrls.append(None)
        self.closed = closed

    def nseg(self):
        return len(self.nodes) if self.closed else len(self.nodes) - 1

    def seg_pts(self, k):
        a = self.nodes[k]
        b = self.nodes[(k + 1) % len(self.nodes)]
        return a, self.ctrls[k], b

    def point(self, k, t):
        a, c, b = self.seg_pts(k)
        if c is None:
            return (a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t)
        return bez(a, c[0], c[1], b, t)

    def seg_len(self, k):
        a, c, b = self.seg_pts(k)
        if c is None:
            return math.hypot(b[0] - a[0], b[1] - a[1])
        L, prev = 0.0, a
        for i in range(1, 201):
            q = bez(a, c[0], c[1], b, i / 200.0)
            L += math.hypot(q[0] - prev[0], q[1] - prev[1])
            prev = q
        return L

    def param_at_fraction(self, k, f):
        a, c, b = self.seg_pts(k)
        if c is None or f <= 0:
            return max(0.0, f)
        if f >= 1:
            return 1.0
        N = 2000
        cum, prev = [0.0], a
        for i in range(1, N + 1):
            q = bez(a, c[0], c[1], b, i / float(N))
            cum.append(cum[-1] + math.hypot(q[0] - prev[0], q[1] - prev[1]))
            prev = q
        target = f * cum[-1]
        lo, hi = 0, N
        while hi - lo > 1:
            mid = (lo + hi) // 2
            if cum[mid] < target:
                lo = mid
            else:
                hi = mid
        span = cum[hi] - cum[lo]
        u = 0.0 if span <= 0 else (target - cum[lo]) / span
        return (lo + u) / float(N)

    def samples(self, per=40):
        out = []
        for k in range(self.nseg()):
            a, c, b = self.seg_pts(k)
            if c is None:
                out.append(a)
            else:
                for i in range(per):
                    out.append(self.point(k, i / float(per)))
        if not self.closed:
            out.append(self.nodes[-1])
        return out

    def transform(self, f):
        self.nodes = [f(p) for p in self.nodes]
        self.ctrls = [None if c is None else (f(c[0]), f(c[1])) for c in self.ctrls]


def rect_path(x0, y0, x1, y1):
    return Path([(x0, y0), (x1, y0), (x1, y1), (x0, y1)])


def ellipse_path(cx, cy, rx, ry):
    k = 0.5522847498
    pts = [(cx + rx, cy), (cx, cy + ry), (cx - rx, cy), (cx, cy - ry)]
    ctr = [((cx + rx, cy + k * ry), (cx + k * rx, cy + ry)),
           ((cx - k * rx, cy + ry), (cx - rx, cy + k * ry)),
           ((cx - rx, cy - k * ry), (cx - k * rx, cy - ry)),
           ((cx + k * rx, cy - ry), (cx + rx, cy - k * ry))]
    return Path(pts, ctr, True)


# ---------------------------------------------------------------------
# collections
# ---------------------------------------------------------------------
class Coll(object):
    def __init__(self, items):
        self._items = list(items)

    @property
    def count(self):
        return len(self._items)

    def item(self, i):
        i = int(V(i))
        if i < 1 or i > len(self._items):
            raise VBError('Index out of range: %d' % i, 9)
        return self._items[i - 1]

    def items_list(self):
        return list(self._items)


class ShapeRange(Coll):
    def __init__(self, doc, items):
        Coll.__init__(self, items)
        self.doc = doc

    def group(self):
        if len(self._items) < 2:
            raise VBError('Grouping needs at least two shapes', 5)
        return self.doc.group_shapes(self._items)

    def add(self, s):
        self._items.append(V(s))

    def getboundingbox(self, x, y, w, h, *rest):
        bb = None
        for s in self._items:
            b = s.bbox()
            bb = b if bb is None else (min(bb[0], b[0]), min(bb[1], b[1]), max(bb[2], b[2]), max(bb[3], b[3]))
        if bb is None:
            bb = (0, 0, 0, 0)
        x.v, y.v, w.v, h.v = bb[0], bb[1], bb[2] - bb[0], bb[3] - bb[1]

    def delete(self):
        for s in list(self._items):
            s.delete()


# ---------------------------------------------------------------------
# curve objects (live views on a shape's paths)
# ---------------------------------------------------------------------
class Node(object):
    def __init__(self, shape, pi, k):
        self._shape, self._pi, self._k = shape, pi, k

    @property
    def _path(self):
        return self._shape.paths[self._pi]

    @property
    def positionx(self):
        return self._path.nodes[self._k][0]

    @property
    def positiony(self):
        return self._path.nodes[self._k][1]

    @property
    def subpath(self):
        return SubPath(self._shape, self._pi)

    def breakapart(self):
        p = self._path
        k = self._k
        if not p.closed:
            raise VBError('BreakApart on an open path is not modelled', 5)
        n = len(p.nodes)
        p.nodes = p.nodes[k:] + p.nodes[:k] + [p.nodes[k]]
        p.ctrls = p.ctrls[k:] + p.ctrls[:k]
        p.closed = False
        self._shape.doc.log.append(('breakapart', self._pi, k))

    def joinwith(self, other):
        other = V(other)
        p = self._path
        if other._pi != self._pi:
            raise VBError('JoinWith across subpaths is not modelled', 5)
        a, b = self.positionx, self.positiony
        if p.closed or len(p.nodes) < 3:
            raise VBError('JoinWith needs an open path', 5)
        last = p.nodes[-1]
        first = p.nodes[0]
        if math.hypot(last[0] - first[0], last[1] - first[1]) > 1e-6:
            raise VBError('JoinWith of distant ends is not modelled', 5)
        p.nodes = p.nodes[:-1]
        p.closed = True
        self._shape.doc.log.append(('joinwith', self._pi))


class Segment(object):
    def __init__(self, shape, pi, k):
        self._shape, self._pi, self._k = shape, pi, k

    @property
    def _path(self):
        return self._shape.paths[self._pi]

    @property
    def startnode(self):
        return Node(self._shape, self._pi, self._k)

    @property
    def endnode(self):
        return Node(self._shape, self._pi, (self._k + 1) % len(self._path.nodes))

    @property
    def length(self):
        return self._path.seg_len(self._k)

    @property
    def type(self):
        return 0 if self._path.ctrls[self._k] is None else 1

    def getpointpositionat(self, x, y, offset=None, otype=None):
        if not self._shape.doc.host.sampling_supported:
            raise VBError('Object does not support this property or method (438)', 438)
        t = float(V(offset)) if offset is not None else 0.5
        if otype is not None and int(V(otype)) != CONSTS['cdrrelativesegmentoffset']:
            raise VBError('absolute offsets are not modelled', 5)
        mode = getattr(self._shape.doc.host, 'sampling_mode', 'param')
        if mode == 'garbage':
            x.v, y.v = 12345.0, -777.0
            return
        if mode == 'arclen':
            # CorelDRAW's relative offset is a fraction of the ARC LENGTH, not of
            # the bezier parameter: walk a fine table to find that point
            t = self._path.param_at_fraction(self._k, t)
        px, py = self._path.point(self._k, t)
        x.v, y.v = px, py

    def addnodeat(self, offset, otype=None):
        t = float(V(offset))
        p = self._path
        k = self._k
        a, c, b = p.seg_pts(k)
        if c is None:
            m = (a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t)
            p.nodes.insert(k + 1, m)
            p.ctrls[k:k + 1] = [None, None]
        else:
            c1, c2 = c
            # de Casteljau split
            def lerp(u, v):
                return (u[0] + (v[0] - u[0]) * t, u[1] + (v[1] - u[1]) * t)
            p01, p12, p23 = lerp(a, c1), lerp(c1, c2), lerp(c2, b)
            p012, p123 = lerp(p01, p12), lerp(p12, p23)
            m = lerp(p012, p123)
            p.nodes.insert(k + 1, m)
            p.ctrls[k:k + 1] = [(p01, p012), (p123, p23)]
        return Node(self._shape, self._pi, k + 1)


class SubPath(object):
    def __init__(self, shape, pi):
        self._shape, self._pi = shape, pi

    @property
    def _path(self):
        return self._shape.paths[self._pi]

    @property
    def segments(self):
        return Coll([Segment(self._shape, self._pi, k) for k in range(self._path.nseg())])

    @property
    def nodes(self):
        return Coll([Node(self._shape, self._pi, k) for k in range(len(self._path.nodes))])

    @property
    def startnode(self):
        return Node(self._shape, self._pi, 0)

    @property
    def endnode(self):
        p = self._path
        return Node(self._shape, self._pi, 0 if p.closed else len(p.nodes) - 1)

    @property
    def length(self):
        return sum(self._path.seg_len(k) for k in range(self._path.nseg()))

    def _get_closed(self):
        return self._path.closed

    def _set_closed(self, v):
        p = self._path
        v = bool(V(v))
        if v and not p.closed:
            p.closed = True
            p.ctrls.append(None)
        elif not v and p.closed:
            p.closed = False
            p.ctrls = p.ctrls[:-1]
    closed = property(_get_closed, _set_closed)

    def reversedirection(self):
        if not self._shape.doc.host.reverse_supported:
            raise VBError('Object does not support this property or method (438)', 438)
        p = self._path
        if p.closed:
            # keep the start node, walk the other way
            n = len(p.nodes)
            nodes = [p.nodes[0]] + p.nodes[1:][::-1]
            ctrls = []
            for k in range(n):
                # new segment k goes nodes[k] -> nodes[k+1] (new order) = old segment reversed
                old = (n - 1 - k) % n
                c = p.ctrls[old]
                ctrls.append(None if c is None else (c[1], c[0]))
            p.nodes, p.ctrls = nodes, ctrls
        else:
            p.nodes = p.nodes[::-1]
            p.ctrls = [None if c is None else (c[1], c[0]) for c in p.ctrls[::-1]]


class Curve(object):
    def __init__(self, shape):
        self._shape = shape

    @property
    def subpaths(self):
        return Coll([SubPath(self._shape, i) for i in range(len(self._shape.paths))])


# ---------------------------------------------------------------------
# shapes
# ---------------------------------------------------------------------
class _Color(object):
    def rgbassign(self, r, g, b):
        self.rgb = (V(r), V(g), V(b))


class _Outline(object):
    def __init__(self):
        self.width = 0.0
        self.color = _Color()


class _Fill(object):
    def applynofill(self):
        self.none = True


class Shape(object):
    _ids = [0]

    def __init__(self, doc, typ, paths=None, children=None, name=''):
        self.doc = doc
        self.type = typ
        self.paths = paths or []
        self.children = children or []
        self.name = name
        self.layer = None
        self.parent = None
        self.fill = _Fill()
        self.outline = _Outline()
        self.powerclip = None
        self.locked = False
        Shape._ids[0] += 1
        self.staticid = Shape._ids[0]
        self.deleted = False

    # geometry
    def all_paths(self):
        if self.type == CONSTS['cdrgroupshape']:
            out = []
            for c in self.children:
                out += c.all_paths()
            return out
        return list(self.paths)

    def bbox(self):
        pts = [q for p in self.all_paths() for q in p.samples()]
        if not pts:
            return (0.0, 0.0, 0.0, 0.0)
        xs = [p[0] for p in pts]
        ys = [p[1] for p in pts]
        return min(xs), min(ys), max(xs), max(ys)

    def _xform(self, f):
        for p in self.all_paths():
            p.transform(f)

    # VBA surface
    def getboundingbox(self, x, y, w, h, *rest):
        b = self.bbox()
        x.v, y.v, w.v, h.v = b[0], b[1], b[2] - b[0], b[3] - b[1]

    def _unlocked(self):
        if self.locked:
            raise VBError('The object is locked', 5)

    def move(self, dx, dy):
        self._unlocked()
        dx, dy = float(V(dx)), float(V(dy))
        self._xform(lambda p: (p[0] + dx, p[1] + dy))

    def rotate(self, ang):
        self._unlocked()
        a = math.radians(float(V(ang)))
        b = self.bbox()
        cx, cy = (b[0] + b[2]) / 2, (b[1] + b[3]) / 2
        c, s = math.cos(a), math.sin(a)
        self._xform(lambda p: (cx + (p[0] - cx) * c - (p[1] - cy) * s,
                               cy + (p[0] - cx) * s + (p[1] - cy) * c))

    def flip(self, axes):
        self._unlocked()
        axes = int(V(axes))
        b = self.bbox()
        cx, cy = (b[0] + b[2]) / 2, (b[1] + b[3]) / 2
        fx = axes in (1, 3)
        fy = axes in (2, 3)
        self._xform(lambda p: (2 * cx - p[0] if fx else p[0], 2 * cy - p[1] if fy else p[1]))

    @property
    def displaycurve(self):
        if self.type in (CONSTS['cdrtextshape'], CONSTS['cdrgroupshape'], CONSTS['cdrbitmapshape']):
            return None
        c = copy.deepcopy(self)       # DisplayCurve is a copy in CorelDRAW
        return Curve(c)

    @property
    def curve(self):
        if self.type != CONSTS['cdrcurveshape']:
            return None
        return Curve(self)

    @property
    def shapes(self):
        return Coll(self.children)

    def converttocurves(self):
        if self.type in (CONSTS['cdrrectangleshape'], CONSTS['cdrellipseshape'],
                         CONSTS['cdrpolygonshape'], CONSTS['cdrtextshape']):
            self.type = CONSTS['cdrcurveshape']

    def duplicate(self, dx=0, dy=0):
        d = self.doc.clone_shape(self)
        d.move(V(dx), V(dy))
        self.layer._add(d)
        return d

    def delete(self):
        self.doc.remove_shape(self)

    def addtoselection(self):
        if self not in self.doc.selection:
            self.doc.selection.append(self)

    def ungroup(self):
        if self.type != CONSTS['cdrgroupshape']:
            raise VBError('Ungroup on a non-group', 5)
        ly = self.layer
        i = ly.items.index(self)
        ly.items[i:i + 1] = self.children
        for c in self.children:
            c.layer = ly
            c.parent = None
        self.children = []
        self.deleted = True


class Layer(object):
    def __init__(self, doc, name):
        self.doc = doc
        self.name = name
        self.items = []

    def _add(self, s):
        s.layer = self
        self.items.append(s)
        return s

    @property
    def shapes(self):
        return _LayerShapes(self)

    def createrectangle(self, x1, y1, x2, y2, *rest):
        x1, y1, x2, y2 = [float(V(v)) for v in (x1, y1, x2, y2)]
        return self._add(Shape(self.doc, CONSTS['cdrrectangleshape'],
                               [rect_path(min(x1, x2), min(y1, y2), max(x1, x2), max(y1, y2))]))

    def createellipse2(self, cx, cy, r1, r2, *rest):
        return self._add(Shape(self.doc, CONSTS['cdrellipseshape'],
                               [ellipse_path(float(V(cx)), float(V(cy)), float(V(r1)), float(V(r2)))]))

    def createlinesegment(self, x1, y1, x2, y2):
        return self._add(Shape(self.doc, CONSTS['cdrcurveshape'],
                               [Path([(float(V(x1)), float(V(y1))), (float(V(x2)), float(V(y2)))], closed=False)]))

    def activate(self):
        self.doc.active_layer = self


class _LayerShapes(Coll):
    def __init__(self, layer):
        Coll.__init__(self, layer.items)
        self.layer = layer

    @property
    def all(self):
        return ShapeRange(self.layer.doc, self.layer.items)


class Collection(object):
    """VBA Collection: Add(item, key), Item(index | key), Count."""

    def __init__(self):
        self._items, self._keys = [], []

    @property
    def count(self):
        return len(self._items)

    def add(self, item, key=None):
        k = None if key is None else V(key)
        if k is not None and k.lower() in self._keys:
            raise VBError('This key is already associated with an element of this collection', 457)
        self._items.append(V(item))
        self._keys.append(None if k is None else k.lower())

    def item(self, i):
        i = V(i)
        if isinstance(i, str):
            if i.lower() not in self._keys:
                raise VBError('Invalid procedure call or argument', 5)
            return self._items[self._keys.index(i.lower())]
        i = int(i)
        if i < 1 or i > len(self._items):
            raise VBError('Subscript out of range', 9)
        return self._items[i - 1]


class Page(object):
    def __init__(self, doc, w, h):
        self.doc = doc
        self.sizewidth = w
        self.sizeheight = h

    @property
    def shapes(self):
        # Object Manager: on plan one (front) first; the mock stacks last-added on top
        return Coll(list(reversed(self.doc.all_shapes())))

    @property
    def layers(self):
        return Coll(self.doc.layers)

    def createlayer(self, name):
        ly = Layer(self.doc, V(name))
        self.doc.layers.append(ly)
        return ly


class Document(object):
    def __init__(self, host, w=3000.0, h=2000.0):
        self.host = host
        self.unit = CONSTS['cdrmillimeter']
        self.layers = [Layer(self, 'Layer 1')]
        self.active_layer = self.layers[0]
        self.page = Page(self, w, h)
        self.selection = []
        self.undo_stack = []
        self.cmd_depth = 0
        self.log = []

    # helpers for tests
    def add(self, shape, select=True):
        self.active_layer._add(shape)
        if select:
            self.selection.append(shape)
        return shape

    def all_shapes(self):
        return [s for ly in self.layers for s in ly.items]

    def clone_shape(self, s):
        memo = {id(self): self, id(self.host): self.host}
        for ly in self.layers:
            memo[id(ly)] = ly
        d = copy.deepcopy(s, memo)
        Shape._ids[0] += 1
        d.staticid = Shape._ids[0]
        return d

    def remove_shape(self, s):
        for ly in self.layers:
            if s in ly.items:
                ly.items.remove(s)
        if s in self.selection:
            self.selection.remove(s)
        s.deleted = True

    def group_shapes(self, shapes):
        shapes = list(shapes)
        if not shapes:
            raise VBError('Group of nothing', 5)
        ly = shapes[0].layer
        g = Shape(self, CONSTS['cdrgroupshape'], children=shapes)
        pos = min(ly.items.index(s) for s in shapes if s in ly.items) if any(s in ly.items for s in shapes) else len(ly.items)
        for s in shapes:
            for l2 in self.layers:
                if s in l2.items:
                    l2.items.remove(s)
            s.parent = g
            s.layer = ly
        ly.items.insert(pos, g)
        g.layer = ly
        self.selection = [g]
        return g

    # VBA surface
    def begincommandgroup(self, name=''):
        if self.cmd_depth == 0:
            self.undo_stack.append(self._snapshot())
        self.cmd_depth += 1

    def endcommandgroup(self):
        if self.cmd_depth == 0:
            raise VBError('EndCommandGroup without BeginCommandGroup', 5)
        self.cmd_depth -= 1

    def undo(self):
        if not self.undo_stack:
            raise VBError('Nothing to undo', 5)
        self._restore(self.undo_stack.pop())

    def clearselection(self):
        self.selection = []

    def _snapshot(self):
        return [(ly, list(ly.items), [copy.deepcopy(s.__dict__, {id(self): self, id(self.host): self.host,
                                                                   id(ly): ly}) for s in ly.items])
                for ly in self.layers]

    def _restore(self, snap):
        for ly, items, states in snap:
            ly.items = items
            for s, st in zip(items, states):
                s.__dict__.update(st)
        self.selection = []


class Application(object):
    def __init__(self):
        self.optimization = False
        self.refreshes = 0

    def refresh(self):
        self.refreshes += 1


# ---------------------------------------------------------------------
# host services
# ---------------------------------------------------------------------
class ErrObj(object):
    def __init__(self):
        self.number = 0
        self.description = ''

    def set(self, e):
        self.number = getattr(e, 'number', 5)
        self.description = str(e)

    def clear(self):
        self.number = 0
        self.description = ''


class _File(object):
    def __init__(self, path, mode):
        self.path, self.mode, self.lines = path, mode, []


class ADODBStream(object):
    def __init__(self, host):
        self.host = host
        self.type = 1
        self.charset = ''
        self._text = ''

    def open(self):
        self._text = ''

    def writetext(self, t, *rest):
        self._text += V(t)

    def savetofile(self, path, mode=2):
        with open(self.host.fs(V(path)), 'w', encoding='utf-8-sig' if True else 'utf-8', newline='') as f:
            f.write(self._text)

    def loadfromfile(self, path):
        with open(self.host.fs(V(path)), 'r', encoding='utf-8', newline='') as f:
            self._text = f.read()

    def readtext(self, n=-1):
        return self._text

    def close(self):
        pass


class ServerXMLHTTP(object):
    """MSXML2.ServerXMLHTTP: host.http_routes[url] -> (status, body), bir
    istisna ya da callable(url, headers). Yol yoksa sunucu yanit vermez."""

    def __init__(self, host):
        self.host = host
        self.status = 0
        self.responsetext = ''
        self._url = ''
        self.headers = {}

    def settimeouts(self, *a):
        pass

    def open(self, method, url, async_=False):
        self._url = V(url)

    def setrequestheader(self, k, v):
        self.headers[V(k).lower()] = V(v)

    def send(self, body=None):
        self.host.http_log.append(self._url)
        r = self.host.http_routes.get(self._url)
        if callable(r):
            r = r(self._url, self.headers)
        if r is None:
            raise VBError('The server name or address could not be resolved', -2147012889)
        if isinstance(r, Exception):
            raise r
        self.status, self.responsetext = r


class WScriptShell(object):
    def __init__(self, host):
        self.host = host

    def run(self, cmd, style=1, wait=False):
        cmd = V(cmd)
        self.host.commands.append(cmd)
        m_job = re.search(r'-Path "([^"]+)"', cmd)
        if m_job and '-Target "' in cmd:
            # ac2fPlot sender: remember what would have gone down the socket
            self.host.sent.append(open(self.host.fs(m_job.group(1)), 'rb').read())
            return self.host.send_rc
        handler = self.host.window_handler
        if handler is None:
            return 1
        m_in = re.search(r'-In "([^"]+)"', cmd)
        m_out = re.search(r'-Out "([^"]+)"', cmd)
        m_ps = re.search(r'-File "([^"]+)"', cmd)
        return handler(self.host.fs(m_ps.group(1)), self.host.fs(m_in.group(1)), self.host.fs(m_out.group(1)))


class Host(object):
    def __init__(self, tmpdir, page=(3000.0, 2000.0)):
        self.tmpdir = tmpdir
        self.registry = {}
        self.inputs = []
        self.messages = []
        self.commands = []
        self.files = {}
        self.window_handler = None
        self.http_routes = {}
        self.http_log = []
        self.sampling_supported = True
        self.sent = []
        self.send_rc = 0
        self.reverse_supported = True
        self.err = ErrObj()
        self.application = Application()
        self.doc = Document(self, page[0], page[1])
        self.stubs = {}
        self.doevents_calls = 0

    # name table for the compiler
    def names(self):
        base = ['activedocument', 'activeselectionrange', 'activelayer', 'activepage',
                'application', 'err'] + list(CONSTS) + list(self.stubs)
        return base

    def __getattr__(self, n):
        if n in CONSTS:
            return CONSTS[n]
        st = self.__dict__.get('stubs', {})
        if n in st:
            return st[n]
        raise AttributeError(n)

    @property
    def activedocument(self):
        return self.doc

    @property
    def activeselectionrange(self):
        return ShapeRange(self.doc, list(self.doc.selection))

    @property
    def activelayer(self):
        return self.doc.active_layer

    @property
    def activepage(self):
        return self.doc.page

    # VBA functions
    def fs(self, p):
        p = p.replace('\\', '/')
        if not p.startswith('/'):
            p = os.path.join(self.tmpdir, p)
        return p

    def inputbox(self, prompt, title='', default=''):
        self.messages.append(('input', prompt))
        if not self.inputs:
            raise AssertionError('unexpected InputBox: %s' % prompt[:200])
        a = self.inputs.pop(0)
        if callable(a):
            a = a(prompt, default)
        return NullStr('') if a is None else a

    def msgbox(self, prompt, buttons=0, title=''):
        self.messages.append(('msg', prompt))
        # vbOKCancel answers OK; every other button set answers Yes
        return 1 if (int(V(buttons)) & 15) == 1 else 6

    def getsetting(self, app, sec, key, default=''):
        return self.registry.get((app, sec, key.lower()), default)

    def savesetting(self, app, sec, key, value):
        self.registry[(app, sec, key.lower())] = '%s' % (value,)

    def deletesetting(self, app, sec, key=None):
        for k in list(self.registry):
            if k[0] == app and k[1] == sec and (key is None or k[2] == key.lower()):
                del self.registry[k]

    def getallsettings(self, app, sec):
        rows = [[k[2], v] for k, v in self.registry.items() if k[0] == app and k[1] == sec]
        return VArr2(rows) if rows else None

    def environ(self, name):
        return self.tmpdir if name.upper() in ('TEMP', 'TMP') else ''

    def dir(self, path='', attrs=0):
        """Dir$: tam yol (var mi) ya da joker (*, ?) ile listeleme; Dir$()
        listenin sonrakini verir. Buyuk/kucuk harf duyarsiz (Windows gibi)."""
        if not path:
            return self._dir_left.pop(0) if getattr(self, '_dir_left', None) else ''
        p = self.fs(path)
        base = os.path.basename(p)
        if '*' in base or '?' in base:
            parent = os.path.dirname(p)
            try:
                names = sorted(os.listdir(parent))
            except OSError:
                names = []
            pat = base.lower()
            hits = [n for n in names if fnmatch.fnmatchcase(n.lower(), pat)]
            self._dir_left = hits[1:]
            return hits[0] if hits else ''
        self._dir_left = []
        return base if os.path.exists(p) else ''

    def filelen(self, path):
        p = self.fs(path)
        if not os.path.exists(p):
            raise VBError('File not found', 53)
        return os.path.getsize(p)

    def kill(self, path):
        p = self.fs(path)
        if not os.path.exists(p):
            raise VBError('File not found', 53)
        os.remove(p)

    def new_object(self, cls):
        if cls == 'shaperange':
            return ShapeRange(self.doc, [])
        if cls == 'collection':
            return Collection()
        raise VBError('Invalid use of New keyword: %s' % cls, 429)

    def createobject(self, prog):
        if prog.lower() == 'adodb.stream':
            return ADODBStream(self)
        if prog.lower() == 'wscript.shell':
            return WScriptShell(self)
        if prog.lower() in ('msxml2.serverxmlhttp.6.0', 'msxml2.serverxmlhttp'):
            return ServerXMLHTTP(self)
        raise VBError('ActiveX component cannot create object: %s' % prog, 429)

    def freefile(self):
        n = 1
        while n in self.files:
            n += 1
        return n

    def file_open(self, path, mode, n):
        self.files[n] = _File(self.fs(path), mode)

    def file_print(self, n, text):
        self.files[n].lines.append(text)

    def file_close(self, n=None):
        ks = [n] if n is not None else list(self.files)
        for k in ks:
            f = self.files.pop(k)
            if f.mode == 'output':
                with open(f.path, 'w', encoding='latin-1', newline='') as fh:
                    fh.write(''.join(l + '\r\n' for l in f.lines))

    def doevents(self):
        self.doevents_calls += 1
        return 0

    def getattr(self, p):
        p = self.fs(p)
        if not os.path.exists(p):
            raise VBError('File not found', 53)
        return 16 if os.path.isdir(p) else 0
