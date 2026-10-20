#!/usr/bin/env python3
"""ac2fNestCore sinamalari: src/ac2fNestCore.bas, tools/vbapy.py ile
Python'a cevrilip GERCEK geometriyle denetlenir.

Her senaryoda yerlesen her parcanin cokgeni geri kurulur ve sunlara
bakilir:
  - iki parca arasindaki en kisa mesafe >= parca araligi
  - hicbir parca digerinin malzemesinin icinde degil (deliginde olabilir)
  - her parca tabakanin kullanilabilir alaninin icinde

Kullanim:  python3 tools/test_nest.py [-v]
"""
import math, os, sys, time

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from vbapy import load_module

SRC = os.path.join(HERE, '..', 'src', 'ac2fNestCore.bas')
VERBOSE = '-v' in sys.argv

# --------------------------------------------------------------------
# shapes (lists of contours; a contour is (kind, [(x, y), ...]))
# --------------------------------------------------------------------
CLOSED, OPEN, SOLID = 0, 1, 2


def rect(w, h, x=0.0, y=0.0):
    return [(CLOSED, [(x, y), (x + w, y), (x + w, y + h), (x, y + h)])]


def circle_pts(r, cx=0.0, cy=0.0, n=48):
    return [(cx + r * math.cos(2 * math.pi * i / n), cy + r * math.sin(2 * math.pi * i / n))
            for i in range(n)]


def circle(r, n=48):
    return [(CLOSED, circle_pts(r, n=n))]


def ring(ro, ri, n=48):
    return [(CLOSED, circle_pts(ro, n=n)), (CLOSED, circle_pts(ri, n=n)[::-1])]


def lshape(a, t):
    return [(CLOSED, [(0, 0), (a, 0), (a, t), (t, t), (t, a), (0, a)])]


def tshape(a, t):
    return [(CLOSED, [(0, 0), (a, 0), (a, t), ((a + t) / 2, t), ((a + t) / 2, a),
                      ((a - t) / 2, a), ((a - t) / 2, t), (0, t)])]


def star(ro, ri, k=5):
    pts = []
    for i in range(2 * k):
        r = ro if i % 2 == 0 else ri
        a = math.pi / 2 + math.pi * i / k
        pts.append((r * math.cos(a), r * math.sin(a)))
    return [(CLOSED, pts)]


def rotate(shape, deg):
    a = math.radians(deg)
    c, s = math.cos(a), math.sin(a)
    return [(k, [(x * c - y * s, x * s + y * c) for (x, y) in pts]) for (k, pts) in shape]


def translate(shape, dx, dy):
    return [(k, [(x + dx, y + dy) for (x, y) in pts]) for (k, pts) in shape]


# --------------------------------------------------------------------
# geometry checks
# --------------------------------------------------------------------

def seg_dist(a, b, c, d):
    def pt_seg(p, q, r):
        qx, qy = r[0] - q[0], r[1] - q[1]
        L = qx * qx + qy * qy
        if L == 0:
            return math.hypot(p[0] - q[0], p[1] - q[1])
        t = max(0.0, min(1.0, ((p[0] - q[0]) * qx + (p[1] - q[1]) * qy) / L))
        return math.hypot(p[0] - q[0] - t * qx, p[1] - q[1] - t * qy)

    def ccw(p, q, r):
        return (r[1] - p[1]) * (q[0] - p[0]) - (q[1] - p[1]) * (r[0] - p[0])

    d1, d2 = ccw(a, b, c), ccw(a, b, d)
    d3, d4 = ccw(c, d, a), ccw(c, d, b)
    if ((d1 > 0) != (d2 > 0)) and ((d3 > 0) != (d4 > 0)):
        return 0.0
    return min(pt_seg(a, c, d), pt_seg(b, c, d), pt_seg(c, a, b), pt_seg(d, a, b))


def edges(shape):
    for (k, pts) in shape:
        n = len(pts)
        m = n if k != OPEN else n - 1
        for i in range(m):
            yield pts[i], pts[(i + 1) % n]


def pip(p, pts):
    inside = False
    j = len(pts) - 1
    for i in range(len(pts)):
        xi, yi = pts[i]
        xj, yj = pts[j]
        if (yi > p[1]) != (yj > p[1]):
            if p[0] < xi + (p[1] - yi) * (xj - xi) / (yj - yi):
                inside = not inside
        j = i
    return inside


def in_material(p, shape, holes):
    closed = [pts for (k, pts) in shape if k != OPEN]
    n = sum(1 for pts in closed if pip(p, pts))
    if not holes:
        return n > 0
    return n % 2 == 1


def bbox(shape):
    xs = [x for (_, pts) in shape for (x, _) in pts]
    ys = [y for (_, pts) in shape for (_, y) in pts]
    return min(xs), min(ys), max(xs), max(ys)


def min_dist(A, B):
    best = float('inf')
    for a, b in edges(A):
        for c, d in edges(B):
            v = seg_dist(a, b, c, d)
            if v < best:
                best = v
    return best


# --------------------------------------------------------------------
# driver
# --------------------------------------------------------------------
class Job(object):
    def __init__(self, res=1.0, gap=4.0, holes=False, tol=0.0):
        self.m = load_module(SRC)
        self.res, self.gap, self.holes = res, gap, holes
        self.grow = gap / 2.0 + tol
        self.m.f_ac2fncinit(res, self.grow, holes)
        self.classes = []        # list of [variant shapes by variant index]
        self.vshape = {}         # variant index -> shape in the variant frame
        self.parts = []          # class per part
        self.sheets = []         # (kind, data)

    def feed(self, shape):
        m = self.m
        m.f_ac2fncgeoclear()
        for (k, pts) in shape:
            m.f_ac2fncgeobegin(k)
            for (x, y) in pts:
                m.f_ac2fncgeopt(x, y)
            m.f_ac2fncgeoend()

    def add_class(self, shape, angles=(0,), copies=1):
        m = self.m
        c = m.f_ac2fncaddclass()
        for i, a in enumerate(angles):
            s = rotate(shape, a)
            self.feed(s)
            v = m.f_ac2fncaddvariant(i)
            self.vshape[v] = s
        for _ in range(copies):
            self.parts.append(c)
            m.f_ac2fncaddpart(c)
        return c

    def add_rect_sheet(self, w, h, qty=-1, margins=(0, 0, 0, 0)):
        L, R, B, T = margins
        t = self.m.f_ac2fncaddrectsheet(w - L - R, h - B - T, qty)
        self.sheets.append(('rect', (w - L - R, h - B - T)))
        return t

    def add_shape_sheet(self, shape, margin):
        self.feed(shape)
        t = self.m.f_ac2fncaddshapesheet(margin)
        self.sheets.append(('shape', (shape, margin)))
        return t

    def run(self, steps=None, max_inst=20, seconds=None, seed=7):
        m = self.m
        m.f_ac2fncprepare(max_inst)
        m.f_ac2fncoptbegin(seed)
        t0 = time.time()
        n = 0
        fixed = m.f_ac2fncoptfixedsteps()
        while True:
            m.f_ac2fncoptstep()
            n += 1
            if steps is not None and n >= steps:
                break
            if seconds is not None and time.time() - t0 >= seconds and n >= fixed:
                break
            if steps is None and seconds is None and n >= fixed:
                break
        self.elapsed = time.time() - t0
        self.steps = n
        return self.collect()

    def collect(self):
        m = self.m
        out = []
        for p in range(len(self.parts)):
            k, v, cx, cy = m.f_ac2fncresult(p, 0, 0, 0)
            if k < 0:
                out.append(None)
                continue
            r = m.f_ac2fncvariantinfo(v, 0, 0, 0.0, 0.0, 0, 0)
            _, w, h, ox, oy, cells, tag = r
            t = m.f_ac2fncinstancetype(k)
            info = m.f_ac2fncsheettypeinfo(t, 0, 0, 0.0, 0.0, 0)
            _, sw, sh, sox, soy, free = info
            dx = sox + cx * self.res - ox
            dy = soy + cy * self.res - oy
            out.append((k, t, translate(self.vshape[v], dx, dy)))
        self.result = out
        return out

    def check(self):
        """Returns a list of problems; empty means the layout is sound."""
        probs = []
        res = self.result
        byk = {}
        for p, r in enumerate(res):
            if r is None:
                continue
            k, t, s = r
            byk.setdefault(k, []).append((p, t, s))
            kind, data = self.sheets[t]
            x0, y0, x1, y1 = bbox(s)
            if kind == 'rect':
                W, H = data
                if x0 < -1e-6 or y0 < -1e-6 or x1 > W + 1e-6 or y1 > H + 1e-6:
                    probs.append('part %d outside sheet: %.3f %.3f %.3f %.3f (W=%g H=%g)'
                                 % (p, x0, y0, x1, y1, W, H))
            else:
                sshape, margin = data
                for (_, pts) in s:
                    for q in pts:
                        if not in_material(q, sshape, True):
                            probs.append('part %d point outside drawn sheet' % p)
                            break
                dm = min_dist(s, sshape)
                if dm < margin - 1e-6:
                    probs.append('part %d only %.3f from the drawn sheet edge' % (p, dm))
        for k, items in byk.items():
            for i in range(len(items)):
                for j in range(i + 1, len(items)):
                    pa, _, A = items[i]
                    pb, _, B = items[j]
                    ax0, ay0, ax1, ay1 = bbox(A)
                    bx0, by0, bx1, by1 = bbox(B)
                    g = self.gap
                    if ax1 + g < bx0 or bx1 + g < ax0 or ay1 + g < by0 or by1 + g < ay0:
                        continue
                    d = min_dist(A, B)
                    if d < g - 1e-6:
                        probs.append('parts %d and %d only %.3f apart (gap %g)' % (pa, pb, d, g))
                    elif in_material(A[0][1][0], B, self.holes) or in_material(B[0][1][0], A, self.holes):
                        probs.append('parts %d and %d overlap (one inside the other)' % (pa, pb))
        return probs

    def stats(self):
        m = self.m
        placed = sum(1 for r in self.result if r is not None)
        used = {}
        for r in self.result:
            if r is not None:
                used[r[0]] = r[1]
        return placed, len(self.result), len(used), m.f_ac2fncbestmaterial() * self.res * self.res


# --------------------------------------------------------------------
# scenarios
# --------------------------------------------------------------------
FAILS = []


def expect(cond, msg):
    print(('  ok   ' if cond else '  FAIL ') + msg)
    if not cond:
        FAILS.append(msg)


def scenario(name):
    print('\n' + name)


def t_rectangles():
    scenario('Dikdortgenler, tek tabaka, 90 derece')
    j = Job(res=1.0, gap=4.0)
    sizes = [(300, 200), (250, 120), (180, 180), (400, 90), (120, 60), (90, 90), (200, 50), (60, 300)]
    for w, h in sizes:
        j.add_class(rect(w, h), angles=(0, 90), copies=2)
    j.add_rect_sheet(1220, 800, qty=-1)
    j.run()
    probs = j.check()
    expect(not probs, 'gecerli yerlesim' + ('' if not probs else ': ' + probs[0]))
    placed, total, sheets, mat = j.stats()
    expect(placed == total, 'hepsi yerlesti (%d/%d), %d tabaka' % (placed, total, sheets))
    area = sum(w * h for w, h in sizes) * 2
    print('       doluluk %.1f %% (malzeme %.0f mm2), %.2f s' % (100.0 * area / mat, mat, j.elapsed))


def t_lshapes_interlock():
    scenario('L parcalar ic ice girer (sinirlayici kutu degil)')
    j = Job(res=1.0, gap=3.0)
    j.add_class(lshape(200, 60), angles=(0, 90, 180, 270), copies=10)
    j.add_rect_sheet(1000, 1000, qty=1)
    j.run(steps=12)
    probs = j.check()
    expect(not probs, 'gecerli yerlesim' + ('' if not probs else ': ' + probs[0]))
    placed, total, sheets, mat = j.stats()
    expect(placed == total, 'hepsi yerlesti (%d/%d)' % (placed, total))
    bbox_mat = 10 * 203 * 203
    print('       malzeme %.0f mm2, sinirlayici kutularla en az %.0f' % (mat, bbox_mat))
    expect(mat < bbox_mat * 0.85, 'kutu yerlesiminden en az %15 az malzeme')


def t_holes():
    scenario('Kucuk parcalar buyuk parcanin deligine girer')
    for holes in (True, False):
        j = Job(res=1.0, gap=4.0, holes=holes)
        j.add_class(ring(200, 170), angles=(0,), copies=1)
        j.add_class(rect(80, 80), angles=(0,), copies=6)
        j.add_rect_sheet(410, 1000, qty=1)
        j.run(steps=6)
        probs = j.check()
        expect(not probs, 'holes=%s gecerli yerlesim' % holes + ('' if not probs else ': ' + probs[0]))
        ring_pos = j.result[0][2]
        cx = (bbox(ring_pos)[0] + bbox(ring_pos)[2]) / 2
        cy = (bbox(ring_pos)[1] + bbox(ring_pos)[3]) / 2
        inside = 0
        for r in j.result[1:]:
            if r is None:
                continue
            x0, y0, x1, y1 = bbox(r[2])
            if math.hypot((x0 + x1) / 2 - cx, (y0 + y1) / 2 - cy) < 170:
                inside += 1
        if holes:
            expect(inside >= 2, 'delikte %d parca' % inside)
        else:
            expect(inside == 0, 'delik kapaliyken delikte parca yok (%d)' % inside)


def t_shape_sheet():
    scenario('Cizili tabaka (daire artik) ve kenar payi')
    j = Job(res=1.0, gap=4.0)
    j.add_class(rect(100, 60), angles=(0, 90), copies=12)
    j.add_shape_sheet(circle(320, n=96), margin=8.0)
    j.run(steps=8)
    probs = j.check()
    expect(not probs, 'gecerli yerlesim' + ('' if not probs else ': ' + probs[0]))
    placed, total, sheets, mat = j.stats()
    print('       %d/%d parca daireye sigdi' % (placed, total))
    expect(placed >= 12, 'en az 12 parca yerlesti')


def t_quantities():
    scenario('Tabaka adetleri: biri 1 adet, sonra 2 adet, fazlasi disarida kalir')
    j = Job(res=2.0, gap=4.0)
    j.add_class(rect(280, 280), angles=(0,), copies=12)
    j.add_rect_sheet(600, 600, qty=1)     # 4 parca alir
    j.add_rect_sheet(600, 300, qty=2)     # 2'ser parca
    j.run()
    probs = j.check()
    expect(not probs, 'gecerli yerlesim' + ('' if not probs else ': ' + probs[0]))
    placed, total, sheets, mat = j.stats()
    expect(placed == 8, '8 parca yerlesti (%d), 4 disarida' % placed)
    expect(sheets == 3, '3 tabaka kullanildi (%d)' % sheets)


def t_angles():
    scenario('45 derecelik donme ve yildizlar')
    j = Job(res=1.0, gap=3.0)
    j.add_class(star(80, 35), angles=(0, 45, 90, 135, 180, 225, 270, 315), copies=14)
    j.add_class(tshape(150, 40), angles=(0, 90, 180, 270), copies=6)
    j.add_rect_sheet(900, 600, qty=-1)
    j.run(steps=10)
    probs = j.check()
    expect(not probs, 'gecerli yerlesim' + ('' if not probs else ': ' + probs[0]))
    placed, total, sheets, mat = j.stats()
    expect(placed == total, 'hepsi yerlesti (%d/%d), %d tabaka, %.2f s' % (placed, total, sheets, j.elapsed))


def t_open_lines():
    scenario('Acik yollar (cizgiler) de bosluk payi ile yerlesir')
    j = Job(res=1.0, gap=5.0)
    zig = [(OPEN, [(0, 0), (50, 80), (100, 0), (150, 80)])]
    j.add_class(zig, angles=(0, 180), copies=8)
    j.add_rect_sheet(500, 500, qty=1)
    j.run(steps=6)
    probs = j.check()
    expect(not probs, 'gecerli yerlesim' + ('' if not probs else ': ' + probs[0]))


def t_optimiser():
    scenario('Optimizasyon en iyi sonucu tutar ve kotulestirmez')
    shapes = [lshape(220, 70), tshape(180, 50), rect(160, 90), star(70, 30), circle(55)]
    j = Job(res=2.0, gap=4.0)
    for i, s in enumerate(shapes):
        j.add_class(s, angles=(0, 90, 180, 270), copies=4 + i)
    j.add_rect_sheet(1000, 700, qty=-1)
    m = j.m
    m.f_ac2fncprepare(10)
    m.f_ac2fncoptbegin(3)
    mats = []
    for n in range(40):
        m.f_ac2fncoptstep()
        mats.append(m.f_ac2fncbestmaterial())
    j.collect()
    probs = j.check()
    expect(not probs, 'en iyi yerlesim gecerli' + ('' if not probs else ': ' + probs[0]))
    mono = all(mats[i + 1] <= mats[i] + 1e-9 for i in range(len(mats) - 1))
    expect(mono, 'en iyi malzeme hic artmadi')
    print('       ilk gecis %.0f, 6 gecis %.0f, 40 gecis %.0f hucre' % (mats[0], mats[5], mats[-1]))


def t_copies_speed():
    scenario('Cok kopya: 150 ayni yildiz')
    j = Job(res=1.0, gap=3.0)
    j.add_class(star(40, 18), angles=(0, 90, 180, 270), copies=150)
    j.add_rect_sheet(1220, 2440, qty=-1)
    t0 = time.time()
    j.run(steps=1)
    dt = time.time() - t0
    probs = j.check() if VERBOSE else []
    placed, total, sheets, mat = j.stats()
    expect(placed == total and not probs, '150/%d yerlesti, %d tabaka, tek gecis %.2f s' % (placed, sheets, dt))


def t_containment():
    scenario('Ic/dis: icteki sekil distakine baglanir')
    m = load_module(SRC)
    m.f_ac2fncinit(1.0, 2.0, False)
    m.f_ac2fncpoolclear()

    def add(shape):
        m.f_ac2fncgeoclear()
        for (k, pts) in shape:
            m.f_ac2fncgeobegin(k)
            for (x, y) in pts:
                m.f_ac2fncgeopt(x, y)
            m.f_ac2fncgeoend()
        return m.f_ac2fncpooladd()

    a = add(rect(300, 300))                        # frame
    b = add(rect(100, 100, 50, 50))                # inside a
    c = add(rect(30, 30, 60, 60))                  # inside b -> goes with a
    d = add(rect(100, 100, 400, 0))                # separate
    e = add(rect(300, 300, 1000, 0))               # same size twins: not nested
    f = add(rect(300, 300, 1000, 0))
    m.f_ac2fncpoolresolve()
    par = [m.f_ac2fncpoolparent(i) for i in range(6)]
    expect(par == [-1, a, a, -1, -1, -1], 'ebeveynler %s' % par)


def t_signature():
    scenario('Kopya tanima: ayni sekil ayni imza, ayna farkli')
    m = load_module(SRC)
    m.f_ac2fncinit(1.0, 2.0, False)

    def sig(shape):
        m.f_ac2fncgeoclear()
        for (k, pts) in shape:
            m.f_ac2fncgeobegin(k)
            for (x, y) in pts:
                m.f_ac2fncgeopt(x, y)
            m.f_ac2fncgeoend()
        return m.f_ac2fncgeosig()

    L = lshape(100, 30)
    mirrored = [(k, [(-x, y) for (x, y) in pts]) for (k, pts) in L]
    expect(sig(L) == sig(translate(L, 123.4, -55)), 'tasinmis kopya ayni imza')
    expect(sig(L) != sig(mirrored), 'ayna goruntusu farkli imza')


def main():
    t0 = time.time()
    t_signature()
    t_containment()
    t_rectangles()
    t_lshapes_interlock()
    t_holes()
    t_shape_sheet()
    t_quantities()
    t_angles()
    t_open_lines()
    t_optimiser()
    t_copies_speed()
    print('\n%s (%.1f s)' % ('HATA: %d' % len(FAILS) if FAILS else 'Hepsi gecti', time.time() - t0))
    sys.exit(1 if FAILS else 0)


if __name__ == '__main__':
    main()
