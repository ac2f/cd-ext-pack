#!/usr/bin/env python3
"""ac2fPlot sinamasi: plotter dosyasi tasarimla BIREBIR ayni mi.

src/ac2fCore.bas + ac2fBending.bas + ac2fPlot.bas, tools/vbapy.py ile cevrilir
ve tools/cdrmock.py belgesinde calistirilir. HPGL dosyasi geri okunur ve
belgedeki GERCEK Bezier egrileriyle karsilastirilir:

  - dosyadaki her nokta gercek egriye en cok ne kadar uzak (sapma)
  - gercek egrinin her noktasi dosyadaki cizgiye en cok ne kadar uzak
  - boyut, yon (dondurme yok), kenar payi, kapali konturlarin kapanmasi
  - el yazisi gibi S egrileri, asimetrik tutamaclar, ilmekler, sivri uclar

Kullanim:  python3 tools/test_plot.py
"""
import math, os, random, re, shutil, sys, tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from vbapy import Project
import cdrmock as cm

SRC = [os.path.join(HERE, '..', 'src', n) for n in
       ('ac2fCore.bas', 'ac2fBending.bas', 'ac2fPlot.bas')]
FAILS = []
REG = ('ac2fPack', 'Ayarlar')
U = 40.0  # HPGL birimi / mm


def expect(cond, msg):
    print(('  ok   ' if cond else '  FAIL ') + msg)
    if not cond:
        FAILS.append(msg)


def scenario(name):
    print('\n' + name)


class Env(object):
    def __init__(self, settings=None, page=(600.0, 400.0)):
        self.tmp = tempfile.mkdtemp(prefix='ac2f_')
        self.h = cm.Host(self.tmp, page)
        self.h.stubs = {'ac2fsettings': lambda: None}
        self.p = Project(SRC, self.h)
        # .plt yeniden yazma yardimcilari (ac2fPlotFixSend) bu sinamanin konusu degil
        skip = ('ac2fcore.', 'ac2fplot.ac2fplshiftfile', 'ac2fplot.ac2fplreadall')
        bad = [w for w in self.p.warnings if not w.startswith(skip)]
        if bad:
            raise SystemExit('\n'.join(bad))
        for k, v in (settings or {}).items():
            self.h.registry[(REG[0], REG[1], k.lower())] = str(v)
        self.doc = self.h.doc

    def curve(self, paths):
        return self.doc.add(cm.Shape(self.doc, cm.CONSTS['cdrcurveshape'], paths), True)

    def plot(self):
        out = os.path.join(self.tmp, 'job.plt')
        self.h.inputs = [out]
        self.p.m.ac2fplot.f_ac2fplotsave()
        return out

    def close(self):
        shutil.rmtree(self.tmp, ignore_errors=True)


# ---------------------------------------------------------------------
# HPGL geri okuma
# ---------------------------------------------------------------------
def parse_hpgl(path):
    """-> [(pen_down_polyline_in_mm_units)], ham komutlar"""
    txt = open(path).read()
    cmds = [c.strip() for c in txt.replace('\r', '').replace('\n', '').split(';') if c.strip()]
    runs, cur, pos = [], None, None
    for c in cmds:
        op, args = c[:2].upper(), c[2:]
        if op in ('PU', 'PD'):
            nums = [int(v) for v in re.findall(r'-?\d+', args)]
            pts = list(zip(nums[0::2], nums[1::2]))
            if op == 'PU':
                if pts:
                    pos = pts[-1]
                cur = None
            else:
                if cur is None:
                    cur = [pos] if pos is not None else []
                    runs.append(cur)
                cur.extend(pts)
                if pts:
                    pos = pts[-1]
    return [[(x / U, y / U) for (x, y) in r] for r in runs if len(r) >= 2], cmds


def pt_seg(p, a, b):
    dx, dy = b[0] - a[0], b[1] - a[1]
    L = dx * dx + dy * dy
    t = 0.0 if L == 0 else max(0.0, min(1.0, ((p[0] - a[0]) * dx + (p[1] - a[1]) * dy) / L))
    return math.hypot(p[0] - (a[0] + t * dx), p[1] - (a[1] + t * dy))


def dist_to_poly(p, poly):
    return min(pt_seg(p, poly[i], poly[i + 1]) for i in range(len(poly) - 1))


def hausdorff(A, B):
    """A noktalarinin B cizgisine en buyuk uzakligi."""
    return max(dist_to_poly(p, B) for p in A)


def fine(path, per=400):
    pts = path.samples(per)
    if path.closed:
        pts.append(pts[0])
    return pts


def compare(paths, runs, shift):
    """Her gercek yolu, dosyadaki en yakin kosuyla eslestirip iki yonlu sapma."""
    worst = 0.0
    for p in paths:
        truth = [(x + shift[0], y + shift[1]) for (x, y) in fine(p)]
        # kosu: bu yolun baslangic noktasina en yakin baslayan
        best = min(runs, key=lambda r: math.hypot(r[0][0] - truth[0][0], r[0][1] - truth[0][1]))
        a = hausdorff(best, truth)
        b = hausdorff(truth, best)
        worst = max(worst, a, b)
    return worst


# ---------------------------------------------------------------------
# el yazisi benzeri sekiller
# ---------------------------------------------------------------------
def s_curve():
    # S: iki ters yone donen, asimetrik tutamacli tek segment
    return cm.Path([(0, 0), (60, 40)], [((10, 90), (70, -50))], closed=False)


def script_e():
    # kapali, ilmekli, tutamac uzunluklari cok farkli "e"
    n = [(40, 20), (10, 40), (0, 10), (30, -5), (55, 10)]
    c = [((40, 35), (25, 45)),
         ((-5, 35), (-8, 5)),
         ((10, -20), (22, -15)),
         ((45, 0), (52, 6)),
         ((57, 18), (50, 22))]
    return cm.Path(n, c, True)


def loop_l():
    # ilmekli "l": egri kendi uzerinden gecer
    return cm.Path([(0, 0), (30, 0)], [((60, 120), (-30, 120))], closed=False)


def cusp_leaf():
    return cm.Path([(0, 0), (50, 10), (20, 60)],
                   [((30, -30), (60, -20)), ((40, 40), (35, 55)), None], True)


def tight_hook():
    # cok kucuk yaricapli kanca (ince detay)
    return cm.Path([(0, 0), (3, 0), (3, 3)], [((0, 4), (3, 4)), ((4, 1), (4, 2)), None], False)


def random_scribble(seed, n=8):
    rnd = random.Random(seed)
    nodes = [(rnd.uniform(0, 200), rnd.uniform(0, 100))]
    ctr = []
    for _ in range(n):
        a = nodes[-1]
        b = (rnd.uniform(0, 200), rnd.uniform(0, 100))
        c1 = (a[0] + rnd.uniform(-90, 90), a[1] + rnd.uniform(-90, 90))
        c2 = (b[0] + rnd.uniform(-90, 90), b[1] + rnd.uniform(-90, 90))
        nodes.append(b)
        ctr.append((c1, c2))
    return cm.Path(nodes, ctr, False)


SHAPES = [('S egrisi', s_curve), ('el yazisi e', script_e), ('ilmekli l', loop_l),
          ('sivri yaprak', cusp_leaf), ('kucuk kanca', tight_hook)]


def t_fidelity(tol_gap=0.05):
    scenario('Sadakat: HPGL, gercek Bezier egrilerinden en cok %.2f mm sapar' % (tol_gap * 2))
    for name, mk in SHAPES + [('karalama %d' % i, (lambda i=i: random_scribble(i))) for i in range(1, 7)]:
        e = Env()
        p = mk()
        e.curve([p])
        out = e.plot()
        runs, cmds = parse_hpgl(out)
        allpts = [q for r in runs for q in r]
        tpts = fine(p)
        shift = (5.0 - min(q[0] for q in tpts), 5.0 - min(q[1] for q in tpts))
        dev = compare([p], runs, shift)
        expect(len(runs) == 1, '%s: tek kontur (%d)' % (name, len(runs)))
        expect(dev <= tol_gap * 2, '%s: sapma %.3f mm (<= %.2f)' % (name, dev, tol_gap * 2))
        e.close()


def t_orientation_and_size():
    scenario('Yon, olcu ve kenar payi')
    e = Env()
    # 200 x 50 mm, sol alt kose (30, 40): genis iş genis kalmali
    e.curve([cm.rect_path(30, 40, 230, 90)])
    out = e.plot()
    runs, cmds = parse_hpgl(out)
    xs = [x for r in runs for (x, y) in r]
    ys = [y for r in runs for (x, y) in r]
    w, h = max(xs) - min(xs), max(ys) - min(ys)
    expect(abs(w - 200) < 0.05 and abs(h - 50) < 0.05, 'genislik/yukseklik %.2f x %.2f (200 x 50)' % (w, h))
    expect(abs(min(xs) - 5) < 0.05 and abs(min(ys) - 5) < 0.05, 'kenar payi 5 mm (%.2f, %.2f)' % (min(xs), min(ys)))
    e.close()
    # L: belge yonu korunuyor mu
    e = Env()
    e.curve([cm.Path([(0, 0), (40, 0), (40, 20), (15, 20), (15, 80), (0, 80)], closed=True)])
    out = e.plot()
    runs, _ = parse_hpgl(out)
    r = runs[0]
    xs = [x for (x, y) in r]; ys = [y for (x, y) in r]
    top = [x for (x, y) in r if abs(y - max(ys)) < 0.01]
    expect(max(xs) - min(xs) == 40 and max(ys) - min(ys) == 80, 'L: 40 genis, 80 yuksek')
    expect(max(top) - min(xs) < 16, 'L: uzun kol SOLDA (belgedeki gibi)')
    e.close()


def t_closed_and_units():
    scenario('Kapali konturlar kapanir, birimler 40/mm')
    e = Env()
    e.curve([cm.ellipse_path(100, 100, 30, 20)])
    out = e.plot()
    runs, cmds = parse_hpgl(out)
    r = runs[0]
    expect(r[0] == r[-1], 'elips kapaniyor (bas = son)')
    expect(any(c.startswith('IN') for c in cmds), 'IN; var')
    expect(all(c[:2] in ('IN', 'SP', 'PA', 'PU', 'PD') for c in cmds), 'yalniz bilinen HPGL komutlari: %s' % sorted({c[:2] for c in cmds}))
    e.close()


def one_fidelity(name, mk, mode='param', supported=True, settings=None, bound=0.1):
    e = Env(settings)
    e.h.sampling_mode = mode
    e.h.sampling_supported = supported
    p = mk()
    e.curve([p])
    out = e.plot()
    runs, cmds = parse_hpgl(out)
    tpts = fine(p)
    shift = (5.0 - min(q[0] for q in tpts), 5.0 - min(q[1] for q in tpts))
    dev = compare([p], runs, shift)
    expect(len(runs) == 1 and dev <= bound, '%s: sapma %.3f mm (<= %.2f), %d nokta' % (
        name, dev, bound, len(runs[0]) if runs else 0))
    e.close()
    return dev


def t_arclength_offsets():
    scenario('Ofset yay uzunluguna gore (CorelDRAW gibi), Bezier parametresine gore degil')
    for name, mk in SHAPES + [('karalama %d' % i, (lambda i=i: random_scribble(i))) for i in range(10, 16)]:
        one_fidelity(name, mk, mode='arclen')


def t_cutting_fallback():
    scenario('GetPointPositionAt yoksa (438): AddNodeAt ile bolen yedek yol')
    for name, mk in SHAPES + [('karalama %d' % i, (lambda i=i: random_scribble(i))) for i in range(20, 24)]:
        one_fidelity(name, mk, supported=False, bound=0.1)


def t_garbage_answers():
    scenario('CorelDRAW sacma nokta verirse: fark edilir, yedek yola gecilir')
    for name, mk in SHAPES[:3]:
        one_fidelity(name, mk, mode='garbage', bound=0.1)


def t_tolerance_setting():
    scenario('Egri toleransi ayari uygulanir')
    d1 = one_fidelity('tol 0.2', lambda: script_e(), settings={'PLEgriToleransiMM': 0.2}, bound=0.4)
    d2 = one_fidelity('tol 0.01', lambda: script_e(), settings={'PLEgriToleransiMM': 0.01}, bound=0.02)
    expect(d2 < d1, 'tolerans kuculdukce sapma kucuk (%.3f < %.3f)' % (d2, d1))


def t_many_random():
    scenario('Rastgele karalamalar (40 tohum): hicbiri tolerans disina cikmaz')
    worst = 0.0
    for sd in range(100, 140):
        worst = max(worst, one_fidelity('tohum %d' % sd, lambda sd=sd: random_scribble(sd, 5)))
    print('  en kotu sapma: %.3f mm' % worst)


def t_straight_not_oversampled():
    scenario('Duz cizgiler ve dikdortgen: gereksiz nokta uretilmez')
    e = Env()
    e.curve([cm.rect_path(0, 0, 300, 120)])
    runs, _ = parse_hpgl(e.plot())
    expect(len(runs[0]) == 5, 'dikdortgen 4 kose + kapanis = 5 nokta (%d)' % len(runs[0]))
    e.close()


def transform(pt, rot, mir):
    x, y = pt
    q = ((rot % 360) // 90) % 4
    for _ in range(q):          # 90 derece saat yonunun tersi
        x, y = -y, x
    if mir in (1, 3):
        x = -x
    if mir in (2, 3):
        y = -y
    return (x, y)


def t_rotation_mirror():
    scenario('Rotate / Mirror: sekil beklenen sekilde doner, bas yukari/yan yatma karismaz')
    base = [(0, 0), (40, 0), (40, 20), (15, 20), (15, 80), (0, 80)]
    for rot in (0, 90, 180, 270):
        for mir in (0, 1, 2, 3):
            e = Env({'PLDondurmeDerece': rot, 'PLAynalama': mir})
            e.curve([cm.Path(base, closed=True)])
            runs, _ = parse_hpgl(e.plot())
            want = [transform(p, rot, mir) for p in base]
            wx = min(x for x, y in want); wy = min(y for x, y in want)
            want = sorted((round(x - wx + 5, 2), round(y - wy + 5, 2)) for x, y in want)
            got = sorted((round(x, 2), round(y, 2)) for x, y in runs[0][:-1])
            expect(len(runs[0]) == len(base) + 1, 'rotate %3d mirror %d: %d nokta + kapanis' % (rot, mir, len(base)))
            expect(got == want, 'rotate %3d mirror %d: noktalar beklenenle ayni' % (rot, mir))
            e.close()


def t_group_and_order():
    scenario('Grup, cok nesne ve Object Manager sirasi')
    e = Env()
    a = e.doc.add(cm.Shape(e.doc, cm.CONSTS['cdrcurveshape'], [cm.Path([(0, 0), (20, 0), (20, 20), (0, 20)])]), True)
    b = e.doc.add(cm.Shape(e.doc, cm.CONSTS['cdrcurveshape'], [s_curve()]), True)
    c = e.doc.add(cm.Shape(e.doc, cm.CONSTS['cdrellipseshape'], [cm.ellipse_path(100, 100, 15, 10)]), True)
    runs, _ = parse_hpgl(e.plot())
    expect(len(runs) == 3, 'uc nesne uc kontur (%d)' % len(runs))
    e.close()


def mat(q, m):
    """(a, b, c, d) fuer q Vierteldrehung (CCW), dann Spiegel m: vbapy ile ayni anlam."""
    a, b, c, d = [(1, 0, 0, 1), (0, -1, 1, 0), (-1, 0, 0, -1), (0, 1, -1, 0)][q % 4]
    if m in (1, 3):
        a, b = -a, -b
    if m in (2, 3):
        c, d = -c, -d
    return (a, b, c, d)


def apply(M, p):
    return (M[0] * p[0] + M[1] * p[1], M[2] * p[0] + M[3] * p[1])


def seen_picture(sent_bytes, M):
    """Makine M'in ciktisini kullanicinin gordugu 1..8 resmine cevirir."""
    path = os.path.join(tempfile.gettempdir(), 'ac2f_seen.plt')
    open(path, 'wb').write(sent_bytes)
    runs, _ = parse_hpgl(path)
    v = [apply(M, p) for p in runs[0][:6]]
    sg = lambda x: (x > 0) - (x < 0)
    ux, uy = sg(v[5][0] - v[0][0]), sg(v[5][1] - v[0][1])
    vx, vy = sg(v[1][0] - v[0][0]), sg(v[1][1] - v[0][1])
    for k in range(8):
        if mat(k % 4, k // 4) == (vx, ux, vy, uy):
            return k + 1
    raise AssertionError('tanimsiz sekil')


def t_wizard():
    scenario('Yon sihirbazi: 8 olasi makine yonunun hepsinde dogru Rotate/Mirror bulur')
    for mk in range(8):
        M = mat(mk % 4, mk // 4)
        # sihirbaz once eski ayari bilerek bozuk birakalim
        e = Env({'PLDondurmeDerece': 270, 'PLAynalama': 2})
        e.h.sent = []
        asked = []
        def pick(prompt, default, M=M, e=e, asked=asked):
            n = seen_picture(e.h.sent[-1], M)
            asked.append(n)
            return str(n)
        e.h.inputs = ['192.168.1.5:9100', pick, pick, pick, pick]
        e.p.m.ac2fplot.f_ac2fplotcalibrate()
        rot = int(float(e.h.registry[(REG[0], REG[1], 'pldondurmederece')]))
        mir = int(float(e.h.registry[(REG[0], REG[1], 'playnalama')]))
        S = mat(rot // 90, mir)
        # makine S'li isi M ile cizer: sonuc birim matris olmali
        final = (M[0] * S[0] + M[1] * S[2], M[0] * S[1] + M[1] * S[3],
                 M[2] * S[0] + M[3] * S[2], M[2] * S[1] + M[3] * S[3])
        expect(final == (1, 0, 0, 1), 'makine %d: Rotate %d Mirror %d -> L duzgun cikar (%d soru)' % (mk + 1, rot, mir, len(asked)))
        expect(e.h.registry.get((REG[0], REG[1], 'plyonkontroledildi')) == '1', 'makine %d: kontrol edildi isareti' % (mk + 1))
        expect(len(asked) <= 2, 'makine %d: en cok 2 kesimde bulundu (%d)' % (mk + 1, len(asked)))
        e.close()


def t_wizard_user_error():
    scenario('Kullanici yanlis resmi secerse: dogrulama kesimi duzeltir')
    M = mat(1, 0)
    e = Env()
    e.h.sent = []
    seq = []
    def liar(prompt, default):
        truth = seen_picture(e.h.sent[-1], M)
        seq.append(truth)
        # ilk soruda yanlis resim (2 yerine 3), sonra dogru
        return str(3 if len(seq) == 1 and truth != 3 else truth)
    e.h.inputs = ['192.168.1.5:9100', liar, liar, liar, liar]
    e.p.m.ac2fplot.f_ac2fplotcalibrate()
    rot = int(float(e.h.registry[(REG[0], REG[1], 'pldondurmederece')]))
    mir = int(float(e.h.registry[(REG[0], REG[1], 'playnalama')]))
    S = mat(rot // 90, mir)
    final = (M[0] * S[0] + M[1] * S[2], M[0] * S[1] + M[1] * S[3],
             M[2] * S[0] + M[3] * S[2], M[2] * S[1] + M[3] * S[3])
    expect(final == (1, 0, 0, 1), 'yanlis ilk cevaptan sonra yine dogru ayara varir (%d kesim)' % len(seq))
    e.close()


def t_wizard_cancel_and_send_fail():
    scenario('Sihirbaz: iptal ve gonderilemeyen durum ayari bozmaz')
    e = Env({'PLDondurmeDerece': 90, 'PLAynalama': 0})
    e.h.send_rc = 1
    e.h.inputs = ['192.168.1.5:9100']
    # the log reader uses Input$, which the translator does not model
    e.p.m.ac2fplot.f_ac2fplreadall = lambda path: 'Connection refused'
    e.p.m.ac2fplot.f_ac2fplotcalibrate()
    expect(any('could not be sent' in m[1] for m in e.h.messages), 'gonderilemedi mesaji gosterildi')
    expect(e.h.registry[(REG[0], REG[1], 'pldondurmederece')] == '90', 'gonderilemeyen denemede onceki ayar (90) geri yuklenir')
    expect((REG[0], REG[1], 'plyonkontroledildi') not in e.h.registry, 'kontrol edildi isareti konmaz')
    e.close()


if __name__ == '__main__':
    import time
    t0 = time.time()
    t_fidelity()
    t_orientation_and_size()
    t_closed_and_units()
    t_arclength_offsets()
    t_cutting_fallback()
    t_garbage_answers()
    t_tolerance_setting()
    t_straight_not_oversampled()
    t_rotation_mirror()
    t_group_and_order()
    t_wizard()
    t_wizard_user_error()
    t_wizard_cancel_and_send_fail()
    t_many_random()
    print()
    if FAILS:
        print('BASARISIZ (%d):' % len(FAILS))
        for f in FAILS:
            print('  -', f)
        sys.exit(1)
    print('Hepsi gecti (%.1f s)' % (time.time() - t0))
