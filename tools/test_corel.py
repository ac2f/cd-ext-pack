#!/usr/bin/env python3
"""ac2fNest ve ac2fStart'in CorelDRAW taklidi uzerinde uctan uca sinamasi.

src/ac2fCore.bas + ac2fNestCore.bas + ac2fNest.bas + ac2fStart.bas,
tools/vbapy.py ile cevrilir ve tools/cdrmock.py belgesinde calistirilir.
Makro bittikten sonra belgedeki GERCEK sekillere bakilir:

  - yerlesen parcalar birbirinden en az parca araligi kadar uzak mi
  - hepsi kendi tabakasinin kenar paylari icinde mi
  - icteki sekil distakiyle birlikte mi gitti, kopyalar olustu mu
  - rapor modu belgeyi hic degistirmedi mi
  - pencere protokolu (giris/cikis dosyalari, hazir ayarlar) dogru mu
  - baslangic noktalari istenen yere geldi mi

Kullanim:  python3 tools/test_corel.py
"""
import math, os, shutil, sys, tempfile, time

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import vbapy
from vbapy import Project
import cdrmock as cm
from test_nest import seg_dist, pip

SRC = [os.path.join(HERE, '..', 'src', n) for n in
       ('ac2fCore.bas', 'ac2fNestCore.bas', 'ac2fNest.bas', 'ac2fStart.bas')]
FAILS = []
REG = ('ac2fPack', 'Ayarlar')


def expect(cond, msg):
    print(('  ok   ' if cond else '  FAIL ') + msg)
    if not cond:
        FAILS.append(msg)


def scenario(name):
    print('\n' + name)


# --------------------------------------------------------------------
# set-up
# --------------------------------------------------------------------
class Env(object):
    def __init__(self, settings=None, sheets='1;Sheet;600;400;0', page=(600.0, 400.0)):
        self.tmp = tempfile.mkdtemp(prefix='ac2f_')
        self.h = cm.Host(self.tmp, page)
        self.h.stubs = {'ac2fsettings': lambda: None,
                        'ac2f_k_jb_list': 'JBListePenceresi', 'ac2f_def_jb_list': 1}
        self.p = Project(SRC, self.h)
        bad = [w for w in self.p.warnings if not w.startswith('ac2fcore.')]
        if bad:
            raise SystemExit('\n'.join(bad))
        base = {'NSParcaAraligiMM': 4, 'NSKenarSolMM': 10, 'NSKenarSagMM': 10,
                'NSKenarUstMM': 10, 'NSKenarAltMM': 10, 'NSDonmeAdimiDerece': 90,
                'NSHassasiyetMM': 2, 'NSOptimizeSaniye': 0, 'NSGrupla': 1,
                'NSTabakaCizgisi': 1, 'JBListePenceresi': 0}
        base.update(settings or {})
        for k, v in base.items():
            self.set(k, v)
        if sheets is not None:
            self.set('NSTabakaListesi', sheets)
        self.doc = self.h.doc

    def set(self, key, v):
        self.h.registry[(REG[0], REG[1], key.lower())] = str(v)

    def get(self, key):
        return self.h.registry.get((REG[0], REG[1], key.lower()))

    def add(self, shape, select=True):
        return self.doc.add(shape, select)

    def rect(self, x, y, w, h, name=''):
        return self.add(cm.Shape(self.doc, cm.CONSTS['cdrrectangleshape'],
                                 [cm.rect_path(x, y, x + w, y + h)], name=name))

    def ellipse(self, cx, cy, r, name=''):
        return self.add(cm.Shape(self.doc, cm.CONSTS['cdrellipseshape'],
                                 [cm.ellipse_path(cx, cy, r, r)], name=name))

    def curve(self, paths, name=''):
        return self.add(cm.Shape(self.doc, cm.CONSTS['cdrcurveshape'], paths, name=name))

    def lshape(self, x, y, a, t):
        return self.curve([cm.Path([(x, y), (x + a, y), (x + a, y + t), (x + t, y + t),
                                    (x + t, y + a), (x, y + a)])])

    def run(self, macro):
        getattr(getattr(self.p.m, 'ac2fnest' if 'nest' in macro else 'ac2fstart'), 'f_' + macro)()
        return self.h.messages[-1][1] if self.h.messages else ''

    def close(self):
        shutil.rmtree(self.tmp, ignore_errors=True)


def leaves(s):
    if s.type == cm.CONSTS['cdrgroupshape']:
        out = []
        for c in s.children:
            out += leaves(c)
        return out
    return [s]


def polys(s):
    out = []
    for leaf in leaves(s):
        for p in leaf.paths:
            out.append((p.closed, p.samples(24)))
    return out


def min_dist(A, B):
    best = float('inf')
    for (ca, pa) in A:
        na = len(pa)
        ea = [(pa[i], pa[(i + 1) % na]) for i in range(na if ca else na - 1)]
        for (cb, pb) in B:
            nb = len(pb)
            eb = [(pb[i], pb[(i + 1) % nb]) for i in range(nb if cb else nb - 1)]
            for a, b in ea:
                for c, d in eb:
                    v = seg_dist(a, b, c, d)
                    if v < best:
                        best = v
    return best


def bbox_of(P):
    xs = [x for (_, pts) in P for (x, _) in pts]
    ys = [y for (_, pts) in P for (_, y) in pts]
    return min(xs), min(ys), max(xs), max(ys)


def inside_material(pt, P):
    n = sum(1 for (c, pts) in P if c and pip(pt, pts))
    return n % 2 == 1


def outlines(env):
    out = []

    def walk(s):
        if s.name.lower().startswith('ac2f sheet'):
            out.append(s)
        for c in s.children:
            walk(c)
    for s in env.doc.all_shapes():
        walk(s)
    return out


def check_layout(env, parts, gap, margins=(10, 10, 10, 10), offcut=None, tol=0.15):
    """parts: the top level shapes after nesting. Returns problems."""
    probs = []
    sheets = []
    for o in outlines(env):
        x0, y0, x1, y1 = o.bbox()
        sheets.append((x0, y0, x1, y1))
    L, R, T, B = margins
    P = [polys(s) for s in parts]
    boxes = [bbox_of(p) for p in P]
    for i, b in enumerate(boxes):
        ok = False
        for (x0, y0, x1, y1) in sheets:
            if b[0] >= x0 + L - tol and b[2] <= x1 - R + tol and b[1] >= y0 + B - tol and b[3] <= y1 - T + tol:
                ok = True
        if offcut is not None:
            oc, om = offcut
            OP = polys(oc)
            if all(inside_material(q, OP) for (_, pts) in P[i] for q in pts) and min_dist(P[i], OP) >= om - tol:
                ok = True
        if not ok:
            probs.append('part %d at %s is not inside any sheet margin' % (i, tuple(round(v, 1) for v in b)))
    for i in range(len(P)):
        for j in range(i + 1, len(P)):
            a, b = boxes[i], boxes[j]
            if a[2] + gap < b[0] or b[2] + gap < a[0] or a[3] + gap < b[1] or b[3] + gap < a[1]:
                continue
            d = min_dist(P[i], P[j])
            if d < gap - tol:
                probs.append('parts %d and %d only %.2f apart (gap %g)' % (i, j, d, gap))
            elif inside_material(P[i][0][1][0], P[j]) or inside_material(P[j][0][1][0], P[i]):
                probs.append('parts %d and %d overlap' % (i, j))
    return probs


def top_parts(env):
    return [s for s in env.doc.all_shapes() if not s.name.lower().startswith('ac2f sheet')
            and not s.name.lower().startswith('offcut')]


# --------------------------------------------------------------------
# scenarios
# --------------------------------------------------------------------
def t_basic():
    scenario('Nesting: dikdortgen, daire, L; pencere yok, son ayarlar')
    e = Env()
    for i, (w, h) in enumerate([(120, 80), (90, 60), (200, 50), (60, 60)]):
        e.rect(700 + i * 10, 500 + i * 7, w, h)
    e.ellipse(900, 900, 40)
    e.lshape(1000, 800, 150, 45)
    e.lshape(1200, 800, 150, 45)
    msg = e.run('ac2fnestagain')
    parts = top_parts(e)
    probs = check_layout(e, parts, 4)
    expect('NESTED' in msg, 'rapor: ' + msg.split('\n')[1].strip())
    expect(not probs, 'gercek sekiller cakismiyor, kenar paylari icinde' + ('' if not probs else ': ' + probs[0]))
    expect(len(outlines(e)) >= 1, '%d tabaka cizgisi cizildi' % len(outlines(e)))
    expect(e.doc.cmd_depth == 0, 'komut grubu kapandi')
    e.close()


def t_corner_columns():
    scenario('Sag ust kose + sutun yonu')
    e = Env({'NSBaslangicKosesi': 4, 'NSDoldurmaYonu': 1})
    for i in range(5):
        e.rect(700 + i * 30, 500, 100, 70)
    e.run('ac2fnestagain')
    parts = top_parts(e)
    probs = check_layout(e, parts, 4)
    expect(not probs, 'gecerli' + ('' if not probs else ': ' + probs[0]))
    o = outlines(e)[0].bbox()
    cx = sum(bbox_of(polys(s))[2] for s in parts) / len(parts)
    cy = sum(bbox_of(polys(s))[3] for s in parts) / len(parts)
    expect(abs(max(bbox_of(polys(s))[2] for s in parts) - (o[2] - 10)) < 3,
           'en sagdaki parca sag kenar payina dayali')
    expect(abs(max(bbox_of(polys(s))[3] for s in parts) - (o[3] - 10)) < 3,
           'en ustteki parca ust kenar payina dayali')
    e.close()


def t_inner_outer():
    scenario('Ic/dis: cerceve icindeki sekil onunla gider; gruplama yok')
    e = Env({'NSGrupla': 0, 'NSIcDisBirlikte': 1})
    outer = e.rect(1000, 1000, 200, 150)
    inner = e.ellipse(1100, 1075, 30)
    other = e.rect(700, 700, 120, 90)
    e.run('ac2fnestagain')
    ob = outer.bbox()
    ib = inner.bbox()
    expect(abs((ib[0] + ib[2]) / 2 - (ob[0] + ob[2]) / 2) < 0.5 and abs((ib[1] + ib[3]) / 2 - (ob[1] + ob[3]) / 2) < 0.5,
           'icteki daire hala cercevenin ortasinda')
    groups = [s for s in e.doc.all_shapes() if s.type == cm.CONSTS['cdrgroupshape']]
    expect(not groups, 'yapilan grup geri cozuldu (%d grup)' % len(groups))
    probs = check_layout(e, [outer, other], 4)
    expect(not probs, 'gecerli' + ('' if not probs else ': ' + probs[0]))
    e.close()


def t_copies_mirror_groupsheet():
    scenario('Kopya (3), ayna serbest, tabaka basina grup')
    e = Env({'NSKopyaAdedi': 3, 'NSAynaSerbest': 1, 'NSGrupla': 2})
    e.lshape(700, 700, 120, 40)
    e.rect(900, 700, 80, 50)
    e.run('ac2fnestagain')
    groups = [s for s in e.doc.all_shapes() if s.type == cm.CONSTS['cdrgroupshape']]
    expect(len(groups) >= 1, '%d tabaka grubu' % len(groups))
    parts = []
    for g in groups:
        parts += [c for c in g.children if not c.name.lower().startswith('ac2f sheet')]
    expect(len(parts) == 6, '6 parca (2 x 3 kopya), bulunan %d' % len(parts))
    probs = check_layout(e, parts, 4)
    expect(not probs, 'gecerli' + ('' if not probs else ': ' + probs[0]))
    e.close()


def t_offcut():
    scenario('Cizili artik tabaka (daire) kendi yerinde kullanilir')
    e = Env({'NSArtikKenarMM': 6}, sheets='0;Sheet;600;400;0')
    oc = e.ellipse(300, 300, 220, name='offcut 1')
    for i in range(4):
        e.rect(800 + i * 10, 800, 90, 60)
    e.run('ac2fnestagain')
    parts = top_parts(e)
    probs = check_layout(e, parts, 4, offcut=(oc, 6))
    expect(not probs, 'parcalar dairenin icinde, 6 mm payla' + ('' if not probs else ': ' + probs[0]))
    expect(abs(oc.bbox()[0] - 80) < 0.01, 'artik yerinden oynamadi')
    expect(not outlines(e), 'artik icin tabaka cizgisi cizilmedi')
    e.close()


def snapshot(doc):
    return sorted((s.type, tuple(round(c, 6) for p in s.all_paths() for q in p.nodes for c in q))
                  for s in doc.all_shapes())


def window_handler(env, action='NEST', change=None, presets=None, record=None, preview=None):
    """Answers the two windows the way a user would. action: the settings
    window's button; preview: the preview window's buttons in turn
    (APPLY / SETTINGS / CANCEL); change: edits made in the settings window,
    a list of them when the window opens more than once."""
    state = {'n': 0, 'p': 0}
    if record is not None:
        record.setdefault('previews', [])

    def settings(ps, inp, out):
        text = open(inp, encoding='utf-8-sig').read()
        if record is not None:
            record['in'] = text
            record['ps'] = open(ps, encoding='latin-1').read()
        ch = change
        if isinstance(change, list):
            ch = change[min(state['n'], len(change) - 1)]
        acts = action if isinstance(action, list) else [action]
        act = acts[min(state['n'], len(acts) - 1)]
        state['n'] += 1
        lines = ['A|' + act]
        vals = {}
        sheets = []
        for ln in text.splitlines():
            f = ln.split('|')
            if f[0] == 'K':
                vals[f[1]] = f[2]
            elif f[0] == 'S':
                sheets.append(f[1:])
        if ch:
            vals.update(ch.get('vals', {}))
            if 'sheets' in ch:
                sheets = ch['sheets'](sheets)
        for k, v in vals.items():
            lines.append('K|%s|%s' % (k, v))
        for s in sheets:
            lines.append('S|' + '|'.join(s))
        for k, v in (presets or {}).items():
            lines.append('R|%s|%s' % (k, v))
        with open(out, 'w', encoding='utf-8-sig', newline='') as fh:
            fh.write('\r\n'.join(lines) + '\r\n')
        return 0

    def prev(ps, inp, out):
        if record is not None:
            record['previews'].append(open(inp, encoding='latin-1').read())
            record['pps'] = open(ps, encoding='latin-1').read()
        acts = preview or ['CANCEL']
        act = acts[min(state['p'], len(acts) - 1)]
        state['p'] += 1
        with open(out, 'w', encoding='utf-8-sig', newline='') as fh:
            fh.write('A|%s\r\n' % act)
        return 0

    def h(ps, inp, out):
        if os.path.basename(ps) == 'ac2f_nest_preview.ps1':
            return prev(ps, inp, out)
        return settings(ps, inp, out)
    return h


def t_report_unchanged():
    scenario('Rapor modu belgeyi degistirmez (ic/dis grubu, donmeler, metin kopyasi dahil)')
    e = Env({'JBListePenceresi': 1, 'NSAynaSerbest': 1, 'NSDonmeAdimiDerece': 45})
    e.rect(1000, 1000, 200, 150)
    e.ellipse(1100, 1075, 30)
    e.lshape(700, 700, 120, 40)
    e.add(cm.Shape(e.doc, cm.CONSTS['cdrtextshape'], [cm.rect_path(1300, 1300, 1380, 1330)], name='text'))
    before = snapshot(e.doc)
    e.h.window_handler = window_handler(e, 'REPORT')
    msg = e.run('ac2fnestreport')
    expect('REPORT ONLY' in msg, 'rapor gosterildi')
    expect(snapshot(e.doc) == before, 'belge birebir ayni')
    e.close()


def t_window_protocol():
    scenario('Pencere: giris dosyasi, degisen ayar ve tabaka sirasi, hazir ayar')
    e = Env({'JBListePenceresi': 1}, sheets='1;Big;800;500;0|1;Small;300;200;2')
    e.rect(700, 700, 100, 60)
    oc = e.rect(2000, 2000, 400, 300, name='offcut A')
    rec = {}

    def reorder(rows):
        # offcut first, Small unticked, Big qty 1
        d = [r for r in rows if r[5].startswith('D')]
        b = [r for r in rows if r[1] == 'Big']
        s = [r for r in rows if r[1] == 'Small']
        b[0][4] = '1'
        s[0][0] = '0'
        return d + b + s

    e.h.window_handler = window_handler(e, 'NEST', {'vals': {'gap': '6', 'time': '0'}, 'sheets': reorder},
                                        presets={'Laser 3mm': 'gap=3;mL=5'}, record=rec)
    e.run('ac2fnest')
    expect('K|gap|4' in rec['in'] and 'S|1|Big|800|500|0|L' in rec['in'], 'giris dosyasinda ayarlar ve tabakalar')
    expect('|D' in rec['in'] and 'offcut A' in rec['in'], 'cizili artik giris dosyasinda')
    expect(e.get('NSParcaAraligiMM') == '6', 'yeni aralik kaydedildi')
    expect(e.get('NSTabakaListesi') == '1;Big;800;500;1|0;Small;300;200;2', 'tabaka listesi: %s' % e.get('NSTabakaListesi'))
    expect(e.h.registry.get(('ac2fPack', 'NestPresets', 'laser 3mm')) == 'gap=3;mL=5', 'hazir ayar kaydedildi')
    src = open(os.path.join(HERE, 'nest_window.ps1'), encoding='utf-8').read().replace('\r\n', '\n')
    expect(rec['ps'].replace('\r\n', '\n') == src, 'yazilan betik nest_window.ps1 ile birebir ayni')
    parts = top_parts(e)
    probs = check_layout(e, [s for s in parts if s is not oc], 6, offcut=(oc, 5))
    expect(not probs, 'parca once artiga yerlesti' + ('' if not probs else ': ' + probs[0]))
    e.close()


def t_text_fallback():
    scenario('Pencere acilamazsa metin sorusu (Enter = nest)')
    e = Env({'JBListePenceresi': 1})
    e.rect(700, 700, 100, 60)
    e.rect(800, 700, 100, 60)
    e.h.window_handler = None          # powershell fails
    e.h.inputs = ['T 500x300x1', '']
    msg = e.run('ac2fnest')
    expect(e.get('NSTabakaListesi') == '1;Sheet 500 x 300;500;300;1', 'yazilan tabaka kaydedildi')
    expect('NESTED' in msg, 'nest calisti')
    probs = check_layout(e, top_parts(e), 4)
    expect(not probs, 'gecerli' + ('' if not probs else ': ' + probs[0]))
    e.close()


def t_safe_model():
    scenario('GetPointPositionAt yoksa guvenli model (egriler kutu ile)')
    e = Env()
    e.h.sampling_supported = False
    e.ellipse(900, 900, 50)
    e.ellipse(1000, 900, 50)
    e.ellipse(1100, 900, 35)
    msg = e.run('ac2fnestagain')
    expect('safe model' in msg, 'raporda guvenli model notu')
    probs = check_layout(e, top_parts(e), 4)
    expect(not probs, 'gecerli' + ('' if not probs else ': ' + probs[0]))
    e.close()


def t_unplaced():
    scenario('Sigmayan parca yerinde kalir, tabaka adedi siniri')
    e = Env(sheets='1;Small;300;200;1')
    big = e.rect(5000, 5000, 400, 400)
    a = e.rect(700, 700, 100, 60)
    msg = e.run('ac2fnestagain')
    expect(abs(big.bbox()[0] - 5000) < 1e-6, 'buyuk parca yerinde kaldi')
    expect('Did not fit' in msg, 'raporda sigmayan var')
    e.close()


def preview_polys(text):
    """P and S lines of a preview input -> parts {id: [contours]} and sheets, in mm."""
    out, sheets = {}, {}
    for ln in text.splitlines():
        f = ln.split('|')
        if f[0] == 'P':
            v = [int(x) / 10.0 for x in f[5].split()]
            out.setdefault(int(f[2]), []).append(list(zip(v[0::2], v[1::2])))
        elif f[0] == 'S':
            sheets[int(f[1])] = [int(x) / 10.0 for x in f[4:12]]
    return out, sheets


def t_preview_apply():
    scenario('Onizleme: Preview -> Apply; cizilen yerlesim belgedekiyle ayni')
    e = Env({'JBListePenceresi': 1, 'NSBaslangicKosesi': 3, 'NSDoldurmaYonu': 1})
    shapes = [e.rect(700 + i * 300, 500, w, h) for i, (w, h) in enumerate([(120, 80), (90, 60), (200, 50)])]
    shapes.append(e.lshape(1000, 800, 150, 45))
    shapes.append(e.ellipse(1300, 900, 40))
    rec = {}
    e.h.window_handler = window_handler(e, 'PREVIEW', preview=['APPLY'], record=rec)
    n_msgs = len(e.h.messages)
    e.run('ac2fnest')
    expect(len(rec['previews']) == 1, 'onizleme bir kez acildi')
    src = open(os.path.join(HERE, 'nest_preview.ps1'), encoding='utf-8').read().replace('\r\n', '\n')
    expect(rec.get('pps', '').replace('\r\n', '\n') == src, 'yazilan betik nest_preview.ps1 ile birebir ayni')
    text = rec['previews'][0] if rec['previews'] else ''
    expect('R|PREVIEW - nothing has moved yet' in text, 'onizlemede rapor')
    pp, sheets = preview_polys(text)
    expect(len(pp) == 5 and len(sheets) >= 1, '5 parca, %d tabaka cizildi' % len(sheets))
    probs = check_layout(e, top_parts(e), 4)
    expect(not probs, 'uygulanan yerlesim gecerli' + ('' if not probs else ': ' + probs[0]))
    real = sorted(bbox_of(polys(sh)) for sh in shapes)
    drawn = sorted((min(x for c in v for x, _ in c), min(y for c in v for _, y in c),
                    max(x for c in v for x, _ in c), max(y for c in v for _, y in c)) for v in pp.values())
    worst = max([max(abs(a - b) for a, b in zip(r, d)) for r, d in zip(real, drawn)] or [99])
    expect(len(real) == len(drawn) and worst < 1.0,
           'cizilen parcalar gercekte gittikleri yerde (en buyuk fark %.2f mm)' % worst)
    if sheets and outlines(e):
        ol = outlines(e)[0].bbox()
        expect(max(abs(a - b) for a, b in zip(ol, sheets[1][:4])) < 0.2, 'tabaka 1 onizlemedeki yerinde')
    expect(len(e.h.messages) == n_msgs, 'onizlemeden sonra ayrica mesaj kutusu yok')
    e.close()


def t_preview_back_cancel():
    scenario('Onizleme: Change settings -> yeni aralikla tekrar -> Cancel; belge degismedi')
    e = Env({'JBListePenceresi': 1, 'NSIcDisBirlikte': 1})
    e.rect(1000, 1000, 200, 150)
    e.ellipse(1100, 1075, 30)
    e.lshape(700, 700, 120, 40)
    before = snapshot(e.doc)
    rec = {}
    e.h.window_handler = window_handler(e, 'PREVIEW', change=[{}, {'vals': {'gap': '9'}}],
                                        preview=['SETTINGS', 'CANCEL'], record=rec)
    e.run('ac2fnest')
    expect(len(rec['previews']) == 2, 'onizleme iki kez acildi')
    expect(e.get('NSParcaAraligiMM') == '9', 'ikinci turdaki aralik kaydedildi')
    expect(len(rec['previews']) == 2 and 'Gap 9.00 mm' in rec['previews'][1], 'ikinci onizleme yeni aralikla')
    expect(snapshot(e.doc) == before, 'belge birebir ayni (ic/dis grubu ve donmeler geri alindi)')
    expect(e.doc.cmd_depth == 0, 'komut grubu kapandi')
    e.close()


def t_locked():
    scenario('Kilitli nesne yerinde kalir ve raporda yazar')
    e = Env()
    lk = e.rect(5000, 5000, 100, 60)
    lk.locked = True
    e.rect(700, 700, 100, 60)
    msg = e.run('ac2fnestagain')
    expect(abs(lk.bbox()[0] - 5000) < 1e-6, 'kilitli nesne oynamadi')
    expect('Locked            : 1' in msg, 'raporda kilitli satiri')
    e.close()
    e = Env()
    only = e.rect(700, 700, 100, 60)
    only.locked = True
    msg = e.run('ac2fnestagain')
    expect('locked' in msg.lower() and 'Unlock' in msg, 'yalniz kilitli secildiyse uyari')
    e.close()


def t_real_windows():
    scenario('Gercek pencereler (PowerShell + WinForms taklidi): VBA -> Preview -> Apply')
    import test_windows as tw
    pwsh = tw.pwsh_path()
    if not pwsh:
        print('  --   pwsh bulunamadi (AC2F_PWSH), atlandi')
        return
    win = tw.Win(pwsh)
    e = Env({'JBListePenceresi': 1})
    for i, (w, h) in enumerate([(120, 80), (90, 60), (200, 50)]):
        e.rect(700 + i * 300, 500, w, h)
    e.lshape(1000, 800, 150, 45)
    oc = e.rect(3000, 3000, 300, 250, name='offcut A')
    acts = {'ac2f_nest.ps1': r"""$OnShow = { param($f)
  $main = $f.FirstOf('TabControl').TabPages[0]
  $main.AllOf('TextBox')[0].Text = '5'
  Click $f 'Preview' }""",
            'ac2f_nest_preview.ps1': r"""$OnShow = { param($f)
  if ($f.FirstOf('PictureBox').Image -eq $null) { throw 'nothing drawn' }
  if ($f.FirstOf('ListBox').Items.Count -lt 2) { throw 'no sheet in the list' }
  Click $f 'Apply' }"""}
    runs = []

    def h(ps, inp, out):
        act = os.path.join(win.tmp, 'act.ps1')
        with open(act, 'w') as fh:
            fh.write(acts[os.path.basename(ps)])
        r = win.run_files(ps, inp, out, act)
        runs.append((os.path.basename(ps), r))
        return r['rc']
    e.h.window_handler = h
    e.run('ac2fnest')
    win.close()
    expect([n for n, _ in runs] == ['ac2f_nest.ps1', 'ac2f_nest_preview.ps1'], 'iki pencere sirayla acildi')
    expect(runs and all(r['rc'] == 0 for _, r in runs), 'ikisi de hatasiz calisti')
    expect(len(runs) == 2 and runs[1][1]['stats'] and runs[1][1]['stats'][1] >= 4,
           'onizlemede parcalar cizildi: %s' % (runs[1][1]['stats'] if len(runs) == 2 else None))
    expect(e.get('NSParcaAraligiMM') == '5', 'pencerede yazilan aralik kaydedildi')
    probs = check_layout(e, [s for s in top_parts(e) if s is not oc], 5, offcut=(oc, 5))
    expect(not probs, 'uygulanan yerlesim gecerli' + ('' if not probs else ': ' + probs[0]))


def area(pts):
    a = 0
    for i in range(len(pts)):
        x1, y1 = pts[i]
        x2, y2 = pts[(i + 1) % len(pts)]
        a += x1 * y2 - x2 * y1
    return a / 2


def t_start_points():
    scenario('Baslangic noktalari: 7 (sol ust), 8 tam nokta, yon kurali')
    e = Env()
    r = e.rect(100, 100, 200, 100)
    ring = e.curve([cm.Path([(400, 100), (600, 100), (600, 300), (400, 300)]),
                    cm.Path([(450, 150), (450, 250), (550, 250), (550, 150)])])
    e.h.inputs = ['7 E X K Q']
    e.run('ac2fstartpoints')
    expect(r.type == cm.CONSTS['cdrcurveshape'], 'dikdortgen egriye cevrildi')
    expect(r.paths[0].nodes[0] == (100.0, 200.0), 'dikdortgen sol ustten basliyor: %s' % (r.paths[0].nodes[0],))
    expect(ring.paths[1].nodes[0] == (450.0, 250.0), 'delik kendi kutusunun sol ustunden: %s' % (ring.paths[1].nodes[0],))

    e.doc.selection = [r, ring]
    e.h.inputs = ['8 W']
    e.run('ac2fstartpoints')
    n0 = r.paths[0].nodes[0]
    expect(abs(n0[0] - 200) < 0.01 and abs(n0[1] - 200) < 0.01, 'ust kenarin ortasina dugum eklendi: %s' % (n0,))
    expect(len(r.paths[0].nodes) == 5, 'sekil ayni, bir dugum fazla (%d)' % len(r.paths[0].nodes))
    expect(area(ring.paths[0].nodes) < 0 and area(ring.paths[1].nodes) > 0,
           'dis saat yonunde, delik ters yonde')
    expect(area(r.paths[0].nodes) < 0, 'dikdortgen saat yonunde')

    # circle: exact point at the top centre of the selection
    e2 = Env()
    c = e2.ellipse(300, 300, 100)
    e2.h.inputs = ['2 S X']
    e2.run('ac2fstartpoints')
    n0 = c.paths[0].nodes[0]
    expect(abs(n0[0] - 300) < 0.05 and abs(n0[1] - 200) < 0.05, 'daire alt ortadan: %s' % (tuple(round(v, 3) for v in n0),))

    # no reverse support: reported, rest still done
    e3 = Env()
    e3.h.reverse_supported = False
    q = e3.rect(0, 0, 100, 50)
    e3.h.inputs = ['9 A']
    msg = e3.run('ac2fstartpoints')
    expect('NOT changed' in msg and q.paths[0].nodes[0] == (100.0, 50.0), 'ters cevrilemiyorsa soyler, baslangic yine tasinir')
    e.close(); e2.close(); e3.close()


def t_start_markers():
    scenario('Isaretler: ciz, tekrar calistirinca sil')
    e = Env()
    e.rect(100, 100, 200, 100)
    e.ellipse(500, 500, 50)
    e.doc.selection[0].converttocurves()
    e.doc.selection[1].converttocurves()
    e.run('ac2fstartpointsshow')
    ly = [l for l in e.doc.layers if l.name == 'ac2f start points']
    expect(ly and len(ly[0].items) == 4, 'isaret katmani, 2 yol x (daire + cizgi) = %d' % (len(ly[0].items) if ly else 0))
    e.run('ac2fstartpointsshow')
    expect(len(ly[0].items) == 0, 'ikinci calistirma isaretleri sildi')
    e.close()


def t_nest_then_start():
    scenario('Nesting sonrasi baslangic noktalari (ayar acik)')
    e = Env({'NSBaslangicNoktasi': 1})
    e.set('SPKonum', 1)
    a = e.rect(700, 700, 100, 60)
    b = e.rect(900, 700, 80, 80)
    e.run('ac2fnestagain')
    ok = True
    for s in (a, b):
        x0, y0, x1, y1 = s.bbox()
        n0 = s.paths[0].nodes[0]
        if abs(n0[0] - x0) > 0.01 or abs(n0[1] - y0) > 0.01:
            ok = False
    expect(ok, 'her parca sol alttan basliyor')
    e.close()


PING = '#ac2f-corel 1\r\nok\tAcme Reklam\r\n'


class JobEnv(object):
    """ac2fCore + ac2fJob (StockTrack baglantisi) icin ayri derleme."""

    def __init__(self):
        self.tmp = tempfile.mkdtemp(prefix='ac2f_job_')
        self.h = cm.Host(self.tmp)
        self.h.stubs = {'ac2fsettings': lambda: None}
        self.p = Project([os.path.join(HERE, '..', 'src', n) for n in ('ac2fCore.bas', 'ac2fJob.bas')], self.h)
        # Gun klasoru (CDR Auto Renamer ayari) bu senaryonun konusu degil.
        self.p.m.ac2fjob.f_ac2fjbdaydir = lambda *a, **k: 'D:\\Isler\\21-10-2026'
        # Cevrimdisi kopya (diske yazma) de konusu degil: yok sayilir.
        self.p.m.ac2fjob.f_ac2fjbcachesave = lambda *a, **k: None
        self.p.m.ac2fjob.f_ac2fjbcacherows = lambda *a, **k: -1

    def set(self, key, v):
        self.h.registry[(REG[0], REG[1], key.lower())] = str(v)

    def get(self, key):
        return self.h.registry.get((REG[0], REG[1], key.lower()))

    def close(self):
        shutil.rmtree(self.tmp, ignore_errors=True)


def t_job_local_server():
    scenario('Job: ana sunucu yoksa yerel sunucu (5), bir sure once yerel, gercek hata gizlenmez')
    env = JobEnv()
    try:
        env.set('JBSunucuAdresi', 'http://main/api/v1')
        env.set('JBAnahtar', 'k')
        env.set('JBGunlukKlasorKoku', env.tmp)
        h = env.h
        job = env.p.m.ac2fjob
        main, local = 'http://main/api/v1/corel/ping', 'http://192.168.1.9/api/v1/corel/ping'

        def setup(*answers):
            h.messages[:] = []
            h.inputs[:] = list(answers) + ['']
            job.f_ac2fjobsetup()
            return [m for k, m in h.messages if k == 'msg']

        setup('5=192.168.1.9')
        expect(env.get('JBYerelSunucu') == 'http://192.168.1.9/api/v1',
               '5= adresi tamamlar: %s' % env.get('JBYerelSunucu'))
        h.http_routes = {local: (200, PING)}
        msgs = setup('T')
        expect(any('Acme Reklam' in m and '(local server)' in m for m in msgs),
               'ana sunucu yanit vermeyince yerel sunucu: %s' % msgs[-1:])
        expect(h.http_log == [main, local], 'once ana, sonra yerel: %s' % h.http_log)
        h.http_log[:] = []
        setup('T')
        expect(h.http_log == [local], 'gecisten sonra once yerel sorulur: %s' % h.http_log)

        vbapy.CLOCK['offset'] = 6 * 60
        h.http_log[:] = []
        h.http_routes = {main: (200, PING), local: (200, PING)}
        msgs = setup('T')
        expect(h.http_log == [main] and not any('(local server)' in m for m in msgs),
               'sure dolunca yine ana sunucu: %s' % h.http_log)

        h.http_log[:] = []
        h.http_routes = {main: (401, '{"message":"Corel anahtari gecersiz"}'), local: (200, PING)}
        msgs = setup('T')
        expect(h.http_log == [main] and any('401' in m and 'gecersiz' in m for m in msgs),
               'gercek hata (401) yerel sunucuyla ortulmez: %s' % msgs[-1:])

        h.http_log[:] = []
        h.http_routes = {main: (503, 'down'), local: (200, PING)}
        msgs = setup('T')
        expect(h.http_log == [main, local] and any('(local server)' in m for m in msgs),
               '503 (gecit: sunucu kapali) yerel sunucuya gecer')

        h.http_routes = {}
        msgs = setup('L')
        expect(any('local server did not answer' in m for m in msgs), 'L: yalniz yereli dener')
        h.http_routes = {local: (200, PING)}
        msgs = setup('L')
        expect(any('Local server answers:  Acme Reklam' in m for m in msgs), 'L: yerel yanit')

        setup('5=')
        expect((env.get('JBYerelSunucu') or '') == '', '5= bos: yerel sunucu yok')
    finally:
        vbapy.CLOCK['offset'] = 0.0
        env.close()


def main():
    t0 = time.time()
    only = [a for a in sys.argv[1:] if a.startswith('t_')]
    if only:
        for n in only:
            globals()[n]()
        print('\n%s' % ('HATA: %d' % len(FAILS) if FAILS else 'Hepsi gecti'))
        sys.exit(1 if FAILS else 0)
    t_basic()
    t_corner_columns()
    t_inner_outer()
    t_copies_mirror_groupsheet()
    t_offcut()
    t_report_unchanged()
    t_window_protocol()
    t_text_fallback()
    t_safe_model()
    t_unplaced()
    t_preview_apply()
    t_preview_back_cancel()
    t_locked()
    t_real_windows()
    t_start_points()
    t_start_markers()
    t_nest_then_start()
    t_job_local_server()
    print('\n%s (%.1f s)' % ('HATA: %d' % len(FAILS) if FAILS else 'Hepsi gecti', time.time() - t0))
    sys.exit(1 if FAILS else 0)


if __name__ == '__main__':
    main()
