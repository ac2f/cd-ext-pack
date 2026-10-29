#!/usr/bin/env python3
"""ac2fPanel (ACP V derzi + plakalara bolme) CorelDRAW taklidi uzerinde sinamasi.

src/ac2fCore.bas + ac2fPanel.bas, tools/vbapy.py ile cevrilir ve
tools/cdrmock.py belgesinde calistirilir. Makro bittikten sonra belgedeki
GERCEK sekillere bakilir:

  - tek panel: 110x210, dort derz, adlar, baslangic noktalari
  - dis kesim "ACP Cut", derzler "ACP Grooves" katmaninda; her katmanda grup
  - 4 m tabela, 50 mm kat, 30 mm derz payi, 3.20 m plaka: 3200 + 960
  - plakalar YERINDE: yuzler bitmis olcuyu tam kaplar, derzler asil
    dikdortgenin ustunde, derz paylari disarida ve komsuyla ust uste biner
  - her plakanin dort kenari dogru pay mi (uc kenar 50, birlesim kenari 30;
    ortadakilerin uzun kenarlari 50, uc kenarlari 30)
  - yuzler toplami tam bitmis olcu mu (derz paylari dusulunce net olcu)
  - "sep": plakalar ayri dizilir; "in" yonu, uzun-dik panel, birden cok
    dortgen, iptal, hatali giris, katman yeniden kullanimi
  - hicbir InputBox 1000 karakteri asmiyor mu (VBA ~1024'te keser)

Kullanim:  python3 tools/test_panel.py
"""
import os, shutil, sys, tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from vbapy import Project
import cdrmock as cm

SRC = [os.path.join(HERE, '..', 'src', n) for n in ('ac2fCore.bas', 'ac2fPanel.bas')]
FAILS = []
REG = ('ac2fPack', 'Ayarlar')


def expect(cond, msg):
    print(('  ok   ' if cond else '  FAIL ') + msg)
    if not cond:
        FAILS.append(msg)


def scenario(name):
    print('\n' + name)


class Env(object):
    def __init__(self, page=(9000.0, 9000.0), settings=None):
        self.tmp = tempfile.mkdtemp(prefix='ac2f_')
        self.h = cm.Host(self.tmp, page)
        self.h.stubs = {'ac2fsettings': lambda: None}
        self.p = Project(SRC, self.h)
        bad = [w for w in self.p.warnings if not w.startswith('ac2fcore.')]
        if bad:
            raise SystemExit('\n'.join(bad))
        self.doc = self.h.doc
        for k, v in (settings or {}).items():
            self.set(k, v)

    def set(self, key, v):
        self.h.registry[(REG[0], REG[1], key.lower())] = str(v)

    def get(self, key):
        return self.h.registry.get((REG[0], REG[1], key.lower()))

    def rect(self, x, y, w, h, select=True):
        return self.doc.add(cm.Shape(self.doc, cm.CONSTS['cdrrectangleshape'],
                                     [cm.rect_path(x, y, x + w, y + h)]), select)

    def run(self, macro, *answers):
        """Makroyu verilen InputBox yanitlariyla calistirir; (sorular, son mesaj)."""
        self.h.inputs[:] = list(answers)
        self.h.messages[:] = []
        getattr(self.p.m.ac2fpanel, 'f_' + macro)()
        asked = [m for k, m in self.h.messages if k == 'input']
        msgs = [m for k, m in self.h.messages if k == 'msg']
        return asked, msgs

    def close(self):
        shutil.rmtree(self.tmp, ignore_errors=True)


# --------------------------------------------------------------------
# reading the drawing back
# --------------------------------------------------------------------
CUT, GRV = 'ACP Cut', 'ACP Grooves'


def layer(env, name):
    for ly in env.doc.layers:
        if ly.name == name:
            return ly
    return None


def walk(s):
    """Bir sekil ve (grupsa) altindaki her sey."""
    yield s
    for c in s.children:
        for x in walk(c):
            yield x


def panels(env):
    """Cizilen plakalar, sirayla: [(anahtar, kesim kutusu, derz cizgileri)].
    Kesim "ACP Cut" katmaninda, derzler "ACP Grooves" katmanindadir; ikisi
    ayni anahtarla ("1 of 2 3200x1100" ya da "110x210") eslesir."""
    cut, grv = layer(env, CUT), layer(env, GRV)
    sheets, grooves = {}, {}
    if cut is not None:
        for top in cut.items:
            for s in walk(top):
                if s.name.startswith('ac2f ACP sheet '):
                    sheets[s.name[len('ac2f ACP sheet '):]] = s.bbox()
    if grv is not None:
        for top in grv.items:
            for s in walk(top):
                if s.name.startswith('ac2f ACP grooves ') and s.children and \
                        all(c.name.startswith('ac2f groove ') for c in s.children):
                    grooves[s.name[len('ac2f ACP grooves '):]] = {
                        c.name.replace('ac2f groove ', ''): (c.paths[0].nodes[0], c.paths[0].nodes[1])
                        for c in s.children}
    import re

    def order(k):
        m = re.match(r'(\d+) of (\d+) ', k)
        return int(m.group(1)) if m else 0
    return [(k, sheets[k], grooves[k]) for k in sorted(sheets, key=order) if k in grooves]


def faces(p):
    """Bir panelin derz cizgilerinin kesistigi yuz (genislik, yukseklik)."""
    _, _, lines = p
    xl, xr = lines['1 left'][0][0], lines['3 right'][0][0]
    yt, yb = lines['2 top'][0][1], lines['4 bottom'][0][1]
    return xr - xl, yt - yb


def allowances(p):
    """(sol, alt, sag, ust) derz paylari: derzin plaka kenarina uzakligi."""
    _, (x0, y0, x1, y1), lines = p
    return (lines['1 left'][0][0] - x0, lines['4 bottom'][0][1] - y0,
            x1 - lines['3 right'][0][0], y1 - lines['2 top'][0][1])


def close(a, b, tol=0.05):
    return abs(a - b) <= tol


def near(t, ref, tol=0.05):
    return all(close(a, b, tol) for a, b in zip(t, ref))


# --------------------------------------------------------------------
# scenarios
# --------------------------------------------------------------------
def t_single_unchanged():
    scenario('Tek panel eskisi gibi: 100x200 + 5mm out, plaka sorusu yok')
    e = Env()
    try:
        e.rect(0, 0, 100, 200)
        asked, msgs = e.run('ac2fpanelgroove', '5mm out')
        ps = panels(e)
        expect(len(asked) == 1, 'yalniz kat sorusu soruldu (%d soru)' % len(asked))
        expect(len(ps) == 1 and ps[0][0] == '110x210', 'plaka anahtari: %s' % [p[0] for p in ps])
        cut, grv = layer(e, CUT), layer(e, GRV)
        expect(cut is not None and grv is not None, 'iki katman olustu: %s' % [l.name for l in e.doc.layers])
        expect([s.name for s in cut.items] == ['ac2f ACP sheet 110x210'], 'kesim katmaninda yalniz plaka dis hatti: %s' % [s.name for s in cut.items])
        expect([s.name for s in grv.items] == ['ac2f ACP grooves 110x210'] and len(grv.items[0].children) == 4,
               'derz katmaninda tek grup, 4 cizgi: %s' % [s.name for s in grv.items])
        expect(all(not s.name.startswith('ac2f') for s in e.doc.layers[0].items), 'ilk katmana (kaynagin katmani) dokunulmadi')
        expect(e.doc.active_layer is e.doc.layers[0], 'aktif katman geri yuklendi')
        expect(near(ps[0][1], (-5, -5, 105, 205)), 'plaka -5,-5 .. 105,205: %s' % (ps[0][1],))
        expect(near(faces(ps[0]), (100, 200)), 'derzler 100x200 uzerinde kesisir')
        expect(sorted(ps[0][2]) == ['1 left', '2 top', '3 right', '4 bottom'], 'dort derz: %s' % sorted(ps[0][2]))
        l = ps[0][2]
        expect(near(l['1 left'][0], (0, -5)) and near(l['1 left'][1], (0, 205)), '1 sol alttan baslar')
        expect(near(l['2 top'][0], (-5, 200)), '2 ust soldan baslar')
        expect(near(l['3 right'][0], (100, 205)), '3 sag ustten baslar')
        expect(near(l['4 bottom'][0], (105, 0)), '4 alt sagdan baslar')
        expect(e.get('ACPDerzMM') == '5', 'kat saklandi')
    finally:
        e.close()


def t_user_example():
    scenario('4 m tabela, kat 50, derz payi 30, plaka 3.20 m -> 3200 + 960')
    e = Env()
    try:
        e.rect(0, 0, 4000, 1000)
        asked, msgs = e.run('ac2fpanelgroove', '5cm out', '3200 j30', '1')
        ps = panels(e)
        expect(len(asked) == 3, 'uc soru: kat, plaka+derz, secenek (%d)' % len(asked))
        expect(len(ps) == 2, 'iki plaka cizildi (%d)' % len(ps))
        names = [p[0] for p in ps]
        expect(names == ['1 of 2 3200x1100', '2 of 2 960x1100'], 'plaka anahtarlari: %s' % names)
        cut, grv = layer(e, CUT), layer(e, GRV)
        # Her plaka AYRI gruplanir: tum plakalari saran tek grup yok.
        expect([s.name for s in cut.items] == ['ac2f ACP sheet 1 of 2 3200x1100', 'ac2f ACP sheet 2 of 2 960x1100'],
               'kesim katmani: her plakanin dis hatti ayri nesne, ortak grup yok: %s' % [s.name for s in cut.items])
        expect(all(s.type == cm.CONSTS['cdrrectangleshape'] for s in cut.items), 'dis hatlar grup degil, dortgen')
        expect([s.name for s in grv.items] == ['ac2f ACP grooves 1 of 2 3200x1100', 'ac2f ACP grooves 2 of 2 960x1100'],
               'derz katmani: her plakanin derzi kendi grubunda, ortak grup yok: %s' % [s.name for s in grv.items])
        expect(all(s.type == cm.CONSTS['cdrgroupshape'] and len(s.children) == 4 for s in grv.items), 'her derz grubunda dort cizgi')
        expect(all(s.parent is None for s in cut.items + grv.items), 'hicbiri baska bir grubun icinde degil')
        src = [s for s in e.doc.layers[0].items if s.type == cm.CONSTS['cdrrectangleshape']]
        expect(len(src) == 1 and near(src[0].bbox(), (0, 0, 4000, 1000)), 'kaynak dikdortgen yerinde, ilk katmanda')
        a, b = ps
        expect(close(a[1][2] - a[1][0], 3200) and close(a[1][3] - a[1][1], 1100), 'plaka 1 kesim 3200x1100')
        expect(close(b[1][2] - b[1][0], 960) and close(b[1][3] - b[1][1], 1100), 'plaka 2 kesim 960x1100')
        # Her kenarin payi: bastaki uc kenar 50, birlesim kenari 30.
        expect(near(allowances(a), (50, 50, 30, 50)), 'plaka 1 paylari sol/alt/sag/ust = %s' % (allowances(a),))
        expect(near(allowances(b), (30, 50, 50, 50)), 'plaka 2 paylari sol/alt/sag/ust = %s' % (allowances(b),))
        fa, fb = faces(a), faces(b)
        expect(close(fa[0], 3120) and close(fb[0], 880), 'yuz uzunluklari 3120 ve 880: %s %s' % (fa[0], fb[0]))
        expect(close(fa[0] + fb[0], 4000), 'yuzler toplami tam 4000: %s' % (fa[0] + fb[0]))
        expect(close(fa[1], 1000) and close(fb[1], 1000), 'yukseklik net 1000')
        # Plakalar YERINDE: yuzler 4000x1000'i tam kaplar, derzler asil dikdortgenin ustunde.
        expect(near(a[1], (-50, -50, 3150, 1050)), 'plaka 1 kutusu -50..3150: %s' % (a[1],))
        expect(near(b[1], (3090, -50, 4050, 1050)), 'plaka 2 kutusu 3090..4050: %s' % (b[1],))
        expect(close(a[1][2] - b[1][0], 60), 'birlesimde iki derz payi 60 mm ust uste biner: %s' % (a[1][2] - b[1][0]))
        la, lb = a[2], b[2]
        expect(close(la['1 left'][0][0], 0) and close(lb['3 right'][0][0], 4000), 'ilk sol derz x=0, son sag derz x=4000')
        expect(close(la['3 right'][0][0], 3120) and close(lb['1 left'][0][0], 3120), 'birlesim derzleri ayni x=3120 uzerinde')
        expect(close(la['4 bottom'][0][1], 0) and close(la['2 top'][0][1], 1000), 'alt/ust derz y=0 ve y=1000 (asil dikdortgen)')
        expect(e.get('ACPPlakaListesi') == '3200' and e.get('ACPBirlesimMM') == '30', 'plaka listesi ve derz payi saklandi')
        expect(all(len(m) <= 1000 for m in asked), 'hicbir soru 1000 karakteri asmiyor: %s' % [len(m) for m in asked])
        expect(len(msgs) == 1 and len(msgs[0]) <= 1000, 'rapor 1000 karakteri asmiyor (%d)' % len(msgs[0]))
        expect('faces add up to 4000' in msgs[0], 'rapor net olcuyu soyluyor')
        expect('1. 3200 + 960  faces 3120 + 880  (left 2240)' in asked[2], 'secenek 1 yuzleri de gosterir: %s' % asked[2].split('\n')[4])
        expect('2. 2080 + 2080  faces 2000 + 2000  (left 1120+1120)' in asked[2], 'secenek 2: 2080+2080, yuzler 2000+2000')
        expect('Plates sit in place' in msgs[0] and 'layer ACP Cut' in msgs[0] and 'layer ACP Grooves' in msgs[0], 'rapor yerinde yerlesimi ve katmanlari soyler')
    finally:
        e.close()


def t_equal_option():
    scenario('Secenek 2: esit yuzler, esit aralikli birlesim (2080 + 2080)')
    e = Env()
    try:
        e.rect(0, 0, 4000, 1000)
        e.run('ac2fpanelgroove', '5cm out', '3200 j30', '2')
        ps = panels(e)
        expect(len(ps) == 2, 'iki plaka')
        expect(close(ps[0][1][2] - ps[0][1][0], 2080) and close(ps[1][1][2] - ps[1][1][0], 2080), 'ikisi de 2080 kesim')
        expect(close(faces(ps[0])[0], 2000) and close(faces(ps[1])[0], 2000), 'yuzler 2000 + 2000')
    finally:
        e.close()


def t_three_pieces():
    scenario('3 plaka: orta plakanin iki ucu derz payi, uzun kenarlar kat')
    e = Env()
    try:
        e.rect(0, 0, 7000, 1200)
        e.run('ac2fpanelgroove', '5cm out', '3200 j30', '1')
        ps = panels(e)
        expect(len(ps) == 3, 'uc plaka (%d)' % len(ps))
        lens = [round(p[1][2] - p[1][0], 1) for p in ps]
        expect(lens == [3200, 3200, 820], 'kesim boylari %s' % lens)
        expect(near(allowances(ps[1]), (30, 50, 30, 50)), 'ortadaki: iki paralel kenar 50, diger iki kenar 30: %s' % (allowances(ps[1]),))
        expect(near(allowances(ps[0]), (50, 50, 30, 50)) and near(allowances(ps[2]), (30, 50, 50, 50)), 'bastaki ve sondaki: uc kenar 50, bir kenar 30')
        tot = sum(faces(p)[0] for p in ps)
        expect(close(tot, 7000), 'yuzler toplami 7000: %s' % tot)
        expect(close(ps[0][2]['3 right'][0][0], ps[1][2]['1 left'][0][0]) and close(ps[1][2]['3 right'][0][0], ps[2][2]['1 left'][0][0]),
               'komsu plakalarin birlesim derzleri ayni x uzerinde (yuzler bitisik)')
        expect(close(ps[0][2]['1 left'][0][0], 0) and close(ps[2][2]['3 right'][0][0], 7000), 'ilk sol derz 0, son sag derz 7000 (asil olcu)')
        expect(close(ps[0][1][2] - ps[1][1][0], 60) and close(ps[1][1][2] - ps[2][1][0], 60), 'her birlesimde 60 mm ust uste')
    finally:
        e.close()


def t_custom_lengths():
    scenario('Kendi uzunluklarim: kullanilacak plaka boylari')
    e = Env()
    try:
        e.rect(0, 0, 4000, 1000)
        asked, msgs = e.run('ac2fpanelgroove', '5cm out', '3200 j30', '3000 1160')
        ps = panels(e)
        expect(len(ps) == 2 and close(ps[0][1][2] - ps[0][1][0], 3000) and close(ps[1][1][2] - ps[1][1][0], 1160),
               'kesim 3000 + 1160')
        expect(close(faces(ps[0])[0] + faces(ps[1])[0], 4000), 'net 4000')
    finally:
        e.close()

    e = Env()
    try:
        e.rect(0, 0, 4000, 1000)
        asked, msgs = e.run('ac2fpanelgroove', '5cm out', '3200 j30', '3000 1000', '3000 *')
        ps = panels(e)
        expect(len(asked) == 4, 'toplam tutmayinca ayni soru yeniden soruldu (%d)' % len(asked))
        expect(any('faces add up to 3840' in m.lower() and 'finished size is 4000' in m and 'short by 160' in m for m in msgs), 'uyari yuzler toplami ve farki soyluyor: %s' % [m[:120] for m in msgs])
        expect(len(ps) == 2 and close(ps[1][1][2] - ps[1][1][0], 1160), '* kalani aldi: 1160')
    finally:
        e.close()

    e = Env()
    try:
        e.rect(0, 0, 4000, 1000)
        e.run('ac2fpanelgroove', '5cm out', '3200 3050 j30', '3200:3200 3050:*')
        ps = panels(e)
        expect(len(ps) == 2, 'plaka:uzunluk yazimi, iki farkli plaka')
        if len(ps) == 2:
            expect(close(ps[0][1][2] - ps[0][1][0], 3200) and close(ps[1][1][2] - ps[1][1][0], 960), '3200 + 960')
    finally:
        e.close()

    e = Env()
    try:
        e.rect(0, 0, 4000, 1000)
        asked, msgs = e.run('ac2fpanelgroove', '5cm out', '3200 j30', '3200 960 500', '3.2m 960')
        expect(msgs and 'add up to 4440' in msgs[0] and 'finished size is 4000' in msgs[0] and 'over by 440' in msgs[0], 'uc parca (3200 960 500): yuzler 4440 eder, bitmis olcu 4000: 440 fazla')
        ps = panels(e)
        expect(len(ps) == 2, 'metre yazimi (3.2m) okundu, 3200 + 960')
    finally:
        e.close()


def t_joint_values():
    scenario('Derz payi 20 mm ve yazim bicimleri')
    e = Env()
    try:
        e.rect(0, 0, 4000, 1000)
        e.run('ac2fpanelgroove', '5cm out', '3200 j20', '1')
        ps = panels(e)
        expect(len(ps) == 2, 'iki plaka')
        expect(near(allowances(ps[0]), (50, 50, 20, 50)), 'birlesim payi 20: %s' % (allowances(ps[0]),))
        expect(close(ps[0][1][2] - ps[0][1][0] + ps[1][1][2] - ps[1][1][0], 4000 + 100 + 40), 'toplam kesim 4140')
        expect(e.get('ACPBirlesimMM') == '20', 'derz payi 20 saklandi')
    finally:
        e.close()
    for text, plate, joint in (('3200 joint 25', '3200', '25'), ('320cm d2cm', '3200', '20'),
                               ('3.2m j3cm', '3200', '30'), ('3,2m j2,5cm', '3200', '25'),
                               ('3200, 3050 j30', '3050 3200', '30')):
        e = Env()
        try:
            e.rect(0, 0, 4000, 1000)
            e.run('ac2fpanelgroove', '5cm out', text, '1')
            expect(e.get('ACPPlakaListesi') == plate and e.get('ACPBirlesimMM') == joint,
                   '"%s" -> plaka %s, derz %s (okunan: %s, %s)' % (text, plate, joint, e.get('ACPPlakaListesi'), e.get('ACPBirlesimMM')))
        finally:
            e.close()


def t_sliver():
    scenario('Son plaka ince bir serit olmasin: 100 mm yuz alt siniri')
    e = Env()
    try:
        e.rect(0, 0, 3110, 800)
        e.run('ac2fpanelgroove', '5cm out', '3200 j30', '1')
        ps = panels(e)
        fs = [round(faces(p)[0], 1) for p in ps]
        expect(len(ps) == 2 and min(fs) >= 100 and close(sum(fs), 3110), 'yuzler %s, toplam 3110' % fs)
    finally:
        e.close()


def t_in_direction():
    scenario('"in": secim plakanin kendisi (4160); derz payi dusulunce net 4000')
    e = Env()
    try:
        e.rect(0, 0, 4160, 1000)
        asked, msgs = e.run('ac2fpanelgroove', '5cm in', '3200 j30', '1')
        ps = panels(e)
        expect(len(ps) == 2, 'iki plaka')
        expect(close(ps[0][1][3] - ps[0][1][1], 1000), 'plaka yuksekligi buyumedi (1000)')
        tot = sum(faces(p)[0] for p in ps)
        expect(close(tot, 4000), 'net 4000 (4160 - 100 - 60): %s' % tot)
        expect(close(sum(p[1][2] - p[1][0] for p in ps), 4160), 'kesimler toplami 4160')
        expect('Sheet 4160 - folds 100 - joints 60 = 4000 finished' in asked[2], 'soru net olcuyu gosteriyor')
    finally:
        e.close()


def t_tall_panel():
    scenario('Dik (uzun) panel: Y ekseninde bolunur, plakalar sagda yan yana')
    e = Env()
    try:
        e.rect(0, 0, 1000, 4000)
        e.run('ac2fpanelgroove', '5cm out', '3200 j30', '1')
        ps = panels(e)
        expect(len(ps) == 2, 'iki plaka')
        if len(ps) == 2:
            a, b = ps
            expect(close(a[1][3] - a[1][1], 3200) and close(b[1][3] - b[1][1], 960), 'boylar 3200 ve 960 (Y yonunde)')
            expect(close(a[1][2] - a[1][0], 1100), 'genislik 1100')
            expect(near(a[1], (-50, 850, 1050, 4050)) and near(b[1], (-50, -50, 1050, 910)), 'yerinde: ust plaka 850..4050, alt plaka -50..910 (60 ust uste): %s %s' % (a[1], b[1]))
            expect(close(a[2]['4 bottom'][0][1], 880) and close(b[2]['2 top'][0][1], 880), 'birlesim derzleri y=880 uzerinde ust uste')
            expect(close(a[2]['2 top'][0][1], 4000) and close(b[2]['4 bottom'][0][1], 0), 'ilk ust derz y=4000, son alt derz y=0')
            # ust plaka: ust kenar kat(50), alt kenar derz payi (30)
            expect(near(allowances(a), (50, 30, 50, 50)), 'ust plaka sol/alt/sag/ust = %s' % (allowances(a),))
            expect(near(allowances(b), (50, 50, 50, 30)), 'alt plaka sol/alt/sag/ust = %s' % (allowances(b),))
            expect(close(faces(a)[1] + faces(b)[1], 4000), 'yuzler toplami 4000')
    finally:
        e.close()


def t_no_split_when_it_fits():
    scenario('Plakaya sigan panel bolunmez; menu 28 (ac2fPanelSplit) her zaman sorar')
    e = Env()
    try:
        e.rect(0, 0, 2000, 1000)
        asked, msgs = e.run('ac2fpanelgroove', '5cm out')
        expect(len(asked) == 1 and len(panels(e)) == 1, 'sigan panel: tek soru, tek plaka')
    finally:
        e.close()
    e = Env()
    try:
        e.rect(0, 0, 2000, 1000)
        asked, msgs = e.run('ac2fpanelsplit', '5cm out', '3200 j30', '2')
        expect(len(asked) == 3, 'ac2fPanelSplit sigsa da plaka sorusunu sorar (%d)' % len(asked))
        expect('1. 2100' in asked[2] and '(1 plate' in asked[2], 'secenek 1: tek plaka 2100')
        expect(len(panels(e)) == 1 or len(panels(e)) == 2, 'secime gore cizildi (%d plaka)' % len(panels(e)))
    finally:
        e.close()


def t_zero_and_cancel():
    scenario('0 = bolme yok; iptal hicbir sey cizmez ve birimi geri yukler')
    e = Env()
    try:
        e.rect(0, 0, 4000, 1000)
        e.run('ac2fpanelgroove', '5cm out', '0')
        ps = panels(e)
        expect(len(ps) == 1 and ps[0][0] == '4100x1100', 'plaka sorusunda 0: tek parca 4100x1100')
        expect(e.get('ACPPlakaListesi') is None, '0 saklanmadi')
    finally:
        e.close()
    e = Env()
    try:
        e.rect(0, 0, 4000, 1000)
        e.run('ac2fpanelgroove', '5cm out', '3200 j30', '0')
        expect(len(panels(e)) == 1, 'secenek sorusunda 0: tek parca')
    finally:
        e.close()
    for answers, what in ((('5cm out', None), 'plaka sorusunda iptal'),
                          (('5cm out', '3200 j30', None), 'secenek sorusunda iptal')):
        e = Env()
        try:
            e.rect(0, 0, 4000, 1000)
            e.doc.unit = cm.CONSTS['cdrinch']
            before = len(e.doc.all_shapes())
            e.run('ac2fpanelgroove', *answers)
            expect(len(e.doc.all_shapes()) == before, '%s: belge degismedi' % what)
            expect(e.doc.unit == cm.CONSTS['cdrinch'], '%s: birim geri yuklendi' % what)
            expect(not e.h.application.optimization, '%s: optimizasyon kapali' % what)
            expect(e.doc.cmd_depth == 0 and not e.doc.undo_stack, '%s: komut grubu acilmadi' % what)
        finally:
            e.close()
    e = Env()
    try:
        e.rect(0, 0, 4000, 1000)
        e.doc.unit = cm.CONSTS['cdrinch']
        e.run('ac2fpanelgroove', '5cm out', '3200 j30', '1')
        expect(e.doc.unit == cm.CONSTS['cdrinch'] and e.doc.cmd_depth == 0, 'cizimden sonra birim geri yuklendi, komut grubu kapandi')
        expect(len(e.doc.undo_stack) == 1, 'tek geri alma adimi (%d)' % len(e.doc.undo_stack))
    finally:
        e.close()


def t_bad_input():
    scenario('Hatali giris yeniden sorulur')
    e = Env()
    try:
        e.rect(0, 0, 4000, 1000)
        asked, msgs = e.run('ac2fpanelgroove', '5cm out', 'abc', '3200 j30', '9', '1')
        expect(any('Could not read a plate length' in m for m in msgs), 'plaka okunamadi uyarisi')
        expect(any('There is no option 9' in m for m in msgs), 'olmayan secenek uyarisi')
        expect(len(panels(e)) == 2, 'sonunda iki plaka cizildi')
    finally:
        e.close()
    e = Env()
    try:
        e.rect(0, 0, 4000, 1000)
        asked, msgs = e.run('ac2fpanelgroove', '5cm out', '3200 j30', '2500', '4000:3900 *')
        expect(any('add up' in m or 'need' in m for m in msgs), 'tek uzunluk toplami tutmaz')
    finally:
        e.close()
    e = Env()
    try:
        e.rect(0, 0, 4000, 1000)
        asked, msgs = e.run('ac2fpanelgroove', '5cm out', '3200 j30', '3500 660', '3200 960')
        expect(any('No plate in your list is long enough for 3500' in m for m in msgs), 'listedeki hicbir plaka 3500 icin yetmez uyarisi')
        expect(len(panels(e)) == 2, '3200 960 ile bitti')
    finally:
        e.close()
    e = Env()
    try:
        e.rect(0, 0, 4000, 1000)
        asked, msgs = e.run('ac2fpanelgroove', '5cm out', '3200 j30', '3200:3200 3200:*')
        # 3200 + 960 degil: ikinci plaka * = 960 oldugu icin tamam
        expect(len(panels(e)) == 2, 'ikinci plaka * ile 960')
    finally:
        e.close()


def t_multi_plate_lengths():
    scenario('Iki plaka boyu (3200, 2440): secenekler ikisi icin de sunulur')
    e = Env()
    try:
        e.rect(0, 0, 4000, 1000)
        asked, msgs = e.run('ac2fpanelgroove', '5cm out', '3200 2440 j30', '1')
        q = asked[2]
        expect('of 3200' in q and 'of 2440' in q, 'secenek satirlarinda plaka boyu yazili')
        expect(all(len(m) <= 1000 for m in asked), 'sorular 1000 karakteri asmiyor: %s' % [len(m) for m in asked])
        expect(len(panels(e)) == 2, 'iki plaka')
        expect(e.get('ACPPlakaListesi') == '2440 3200', 'liste artan sirada saklandi: %s' % e.get('ACPPlakaListesi'))
    finally:
        e.close()


def t_many_rectangles():
    scenario('Birden cok dortgen: plaka sorusu bir kez, bolunecek olana secenek')
    e = Env()
    try:
        e.rect(0, 5000, 1000, 500)
        e.rect(0, 0, 4000, 1000)
        asked, msgs = e.run('ac2fpanelgroove', '5cm out', '3200 j30', '1')
        ps = panels(e)
        expect(len(asked) == 3, 'kat + plaka + tek secenek sorusu (%d)' % len(asked))
        expect(len(ps) == 3, 'biri tek parca, biri iki plaka = 3 grup (%d)' % len(ps))
        expect(sum(1 for p in ps if 'of' in p[0]) == 2, 'iki plaka bolunmus panelden')
        expect(any(p[0] == '1100x600' for p in ps), 'kisa panel tek parca cizildi')
    finally:
        e.close()


def t_default_text():
    scenario('Hatirlanan degerler soruda hazir gelir')
    e = Env(settings={'ACPPlakaListesi': '3050 3200', 'ACPBirlesimMM': '20'})
    try:
        e.rect(0, 0, 4000, 1000)
        seen = []
        e.h.inputs[:] = ['5cm out', lambda p, d: (seen.append(d), d)[1], '1']
        e.h.messages[:] = []
        e.p.m.ac2fpanel.f_ac2fpanelgroove()
        expect(seen and seen[0] == '3050 3200 j20', 'plaka sorusunun hazir metni: %r' % (seen,))
        expect(len(panels(e)) == 2, 'iki plaka')
    finally:
        e.close()


def t_prompt_limits():
    scenario('Cok secenekte soru 1000 karakteri asmaz; gorunmeyen numara secilemez')
    e = Env()
    try:
        e.rect(0, 0, 25000, 1000)
        asked, msgs = e.run('ac2fpanelgroove', '5cm out', '3200 3050 2440 2000 1500 1220 j30', '9', '1')
        q = asked[2]
        import re
        nums = [int(n) for n in re.findall(r'^(\d+)\. ', q, re.M)]
        expect(len(q) <= 1000 and len(asked[3]) <= 1000, 'sorular 1000 karakteri asmiyor (%d, %d)' % (len(q), len(asked[3])))
        expect(nums == list(range(1, len(nums) + 1)) and len(nums) >= 3, 'gorunen secenekler 1..%d ardisik' % len(nums))
        expect(any('There is no option 9' in m for m in msgs) or len(nums) >= 9, '9. secenek gorunmuyorsa secilemez')
        expect(len(msgs[-1]) <= 1000, 'rapor 1000 karakteri asmiyor (%d)' % len(msgs[-1]))
    finally:
        e.close()


def t_waste_order():
    scenario('Ayni plaka sayisinda en az fire veren secenek once; ayni kesimler tek satir')
    e = Env()
    try:
        e.rect(0, 0, 4000, 1000)
        asked, msgs = e.run('ac2fpanelgroove', '5cm out', '3200 3050 2440 j30', '1')
        q = asked[2]
        expect('1. 2080 + 2080  faces 2000 + 2000  (of 2440, left 360+360)' in q, 'ilk secenek 2 x 2440 plakadan 2080 + 2080: %s' % q.split('\n')[4])
        expect(q.count('2080 + 2080') == 1, '3200 ve 3050 plakadan ayni 2080 + 2080 tekrar listelenmedi')
        ps = panels(e)
        expect(len(ps) == 2 and close(ps[0][1][2] - ps[0][1][0], 2080), 'secim cizildi: 2080 + 2080')
    finally:
        e.close()


def t_in_place_in_direction():
    scenario('"in" yerinde: secim toplam acinim; net dikdortgen ic kaydirilmis yerde')
    e = Env()
    try:
        e.rect(100, 200, 4160, 1000)
        e.run('ac2fpanelgroove', '5cm in', '3200 j30', '1')
        ps = panels(e)
        expect(len(ps) == 2, 'iki plaka')
        a, b = ps
        expect(near(a[1], (100, 200, 3300, 1200)), 'plaka 1 kutusu secimin sol kenarindan baslar: %s' % (a[1],))
        expect(close(a[2]['1 left'][0][0], 150) and close(b[2]['3 right'][0][0], 4150), 'net 150..4150 (4000)')
        expect(close(a[2]['4 bottom'][0][1], 250) and close(a[2]['2 top'][0][1], 1150), 'net yukseklik 250..1150 (900): kat kadar ice')
        expect(close(a[2]['3 right'][0][0], b[2]['1 left'][0][0]), 'birlesim derzleri ayni x')
        expect(close(a[1][2] - b[1][0], 60), 'birlesimde 60 mm ust uste')
    finally:
        e.close()


def t_sep_layout():
    scenario('sep: plakalar ayri dizilir (kaynagin altina / sagina), sakli degildir')
    e = Env()
    try:
        e.rect(0, 0, 4000, 1000)
        asked, msgs = e.run('ac2fpanelgroove', '5cm out', '3200 j30 sep', '1')
        ps = panels(e)
        expect(len(ps) == 2, 'iki plaka')
        a, b = ps
        expect(near(a[1], (-50, -1250, 3150, -150)), 'plaka 1 kaynagin altinda (aralik 100): %s' % (a[1],))
        expect(b[1][3] <= a[1][1] - 100 + 0.05 and close(b[1][0], -50), 'plaka 2, plaka 1in altinda, sol kenar hizali: %s' % (b[1],))
        expect(near(allowances(a), (50, 50, 30, 50)) and near(allowances(b), (30, 50, 50, 50)), 'paylar yerindekiyle ayni')
        expect('laid apart' in msgs[0], 'rapor ayri dizildigini soyler')
        expect(e.get('ACPPlakaListesi') == '3200', 'sep listeye karismadi')
    finally:
        e.close()
    e = Env()
    try:
        e.rect(0, 0, 1000, 4000)
        e.run('ac2fpanelgroove', '5cm out', '3200 j30 sep', '1')
        ps = panels(e)
        a, b = ps
        expect(a[1][0] >= 1000 + 50 + 100 - 0.05 and b[1][0] >= a[1][2] + 100 - 0.05, 'dik panel: sagda yan yana, aralik >= 100')
        expect(close(a[1][3], 4050) and close(b[1][3], 4050), 'ust kenarlar hizali')
    finally:
        e.close()
    e = Env()
    try:
        src = e.rect(0, 0, 4000, 1000)
        asked, msgs = e.run('ac2fpanelgroove', '5cm out', '3200 j30 sep', '1')
        expect(near(panels(e)[0][1], (-50, -1250, 3150, -150)), 'sep yazilinca ayri dizilir')
        # Ikinci calisma: yalniz kaynak secili, ayni belge (eski cizim duruyor).
        e.doc.selection[:] = [src]
        asked2, msgs2 = e.run('ac2fpanelgroove', '5cm out', '3200 j30', '1')
        yeni = [p for p in panels(e) if near(p[1], (-50, -50, 3150, 1050))]
        expect(len(yeni) == 1 and 'laid apart' not in msgs2[0], 'onceki calismadaki sep sonrakini etkilemez: yerinde cizildi')
    finally:
        e.close()


def t_layers_reused():
    scenario('Katmanlar yeniden kullanilir; var olan katmanin kilidi acilir')
    e = Env()
    try:
        e.rect(0, 0, 4000, 1000)
        e.run('ac2fpanelgroove', '5cm out', '3200 j30', '1')
        names1 = [l.name for l in e.doc.layers]
        e.doc.selection[:] = []
        e.rect(0, 5000, 4000, 1000)
        e.run('ac2fpanelgroove', '5cm out', '3200 j30', '1')
        names2 = [l.name for l in e.doc.layers]
        expect(names1 == names2 == ['Layer 1', 'ACP Cut', 'ACP Grooves'], 'ikinci calismada yeni katman acilmadi: %s' % names2)
        expect(len(layer(e, CUT).items) == 4 and len(layer(e, GRV).items) == 4, 'iki calismanin (2 x 2 plaka) nesneleri ayni katmanlarda')
        expect(e.doc.active_layer is e.doc.layers[0], 'aktif katman geri yuklendi')
    finally:
        e.close()
    e = Env()
    try:
        ly = e.doc.page.createlayer('ACP Grooves')
        ly.editable = False
        ly.visible = False
        e.rect(0, 0, 100, 200)
        e.run('ac2fpanelgroove', '5mm out')
        expect(len(ps_ := panels(e)) == 1 and ly.editable is True and ly.visible is True, 'kilitli/gizli derz katmani acildi ve kullanildi')
    finally:
        e.close()


def t_net_and_cut_pieces():
    scenario('10 m tabela, kat 7.5 cm: parcalar n (yuz) / k (kesim) ile, bukum paylari otomatik')
    e = Env()
    try:
        e.rect(0, 0, 10000, 1200)
        asked, msgs = e.run('ac2fpanelgroove', '7.5cm out', '3200 j30', 'n200cm k320cm n250cm *')
        ps = panels(e)
        expect(len(ps) == 4, 'dort plaka (%d)' % len(ps))
        cuts = [round(p[1][2] - p[1][0], 1) for p in ps]
        expect(cuts == [2105.0, 3200.0, 2560.0, 2465.0], 'kesim boylari (yuz + paylar otomatik): %s' % cuts)
        fs = [round(faces(p)[0], 1) for p in ps]
        expect(fs == [2000.0, 3140.0, 2500.0, 2360.0], 'yuzler 2000 / 3140 (k320 = tam plaka) / 2500 / kalan 2360: %s' % fs)
        expect(close(sum(fs), 10000), 'yuzler toplami tam 10000')
        xs = [round(p[2]['1 left'][0][0], 1) for p in ps]
        expect(xs == [0.0, 2000.0, 5140.0, 7640.0], 'birlesim konumlari 2000 / 5140 / 7640 (asil olcu uzerinde): %s' % xs)
        expect(near(allowances(ps[0]), (75, 75, 30, 75)) and near(allowances(ps[3]), (30, 75, 75, 75)), 'bastaki/sondaki: uc kenar 75 kat, birlesim 30')
        expect(len(asked) == 3 and any(m.startswith('Pieces, left to right') for m in msgs), 'yazilan plan parca parca gosterilip onaylandi')
        conf = [m for m in msgs if m.startswith('Pieces, left to right')][0]
        expect('1.  cut 2105   face 0-2000 (2000)   folds 75 | 30' in conf and '2.  cut 3200   face 2000-5140 (3140)   folds 30 | 30' in conf,
               'onay metni her parcanin kesim, yuz aralik ve paylarini yazar: %s' % conf.split('\r\n')[1:3])
        expect(len(conf) <= 1000, 'onay metni 1000 karakteri asmiyor (%d)' % len(conf))
    finally:
        e.close()

    cases = (
        ('net 2000 3100 * *', [2105.0, 3160.0, 2510.0, 2555.0], [2000.0, 3100.0, 2450.0, 2450.0]),
        ('n2000 cut 3200 3200 *', [2105.0, 3200.0, 3200.0, 1825.0], [2000.0, 3140.0, 3140.0, 1720.0]),
        ('3050:n2000 k3200 *', [2105.0, 3200.0, 4685.0 - 2105.0 - 0.0 if False else None], None),
    )
    for text, want_cuts, want_faces in cases[:2]:
        e = Env()
        try:
            e.rect(0, 0, 10000, 1200)
            e.run('ac2fpanelgroove', '7.5cm out', '3200 j30', text)
            ps = panels(e)
            fs = [round(faces(p)[0], 1) for p in ps]
            cuts = [round(p[1][2] - p[1][0], 1) for p in ps]
            expect(fs == want_faces and cuts == want_cuts, '"%s" -> yuzler %s, kesimler %s' % (text, fs, cuts))
        finally:
            e.close()
    e = Env()
    try:
        e.rect(0, 0, 10000, 1200)
        e.run('ac2fpanelgroove', '7.5cm out', '3200 j30', '3050:n2000 k3200 n2500 *')
        ps = panels(e)
        expect(len(ps) == 4 and close(faces(ps[0])[0], 2000), 'plaka:n birlesimi: 3050 plakasindan 2000 yuz')
        expect(e.get('ACPPlakaListesi') == '3200', 'plaka listesi degismedi')
    finally:
        e.close()


def t_positions():
    scenario('at: birlesimleri net olcu uzerindeki konumla yaz (dekupe bosluklarina denk getir)')
    e = Env()
    try:
        e.rect(0, 0, 10000, 1200)
        asked, msgs = e.run('ac2fpanelgroove', '7.5cm out', '3200 j30', 'at 2000 5100 7500')
        ps = panels(e)
        fs = [round(faces(p)[0], 1) for p in ps]
        expect(len(ps) == 4 and fs == [2000.0, 3100.0, 2400.0, 2500.0], 'konumlardan yuzler: %s' % fs)
        xs = [round(p[2]['1 left'][0][0], 1) for p in ps]
        expect(xs == [0.0, 2000.0, 5100.0, 7500.0], 'birlesim derzleri tam yazilan konumlarda: %s' % xs)
        expect(close(ps[3][2]['3 right'][0][0], 10000), 'son derz 10000 (net olcu)')
        cuts = [round(p[1][2] - p[1][0], 1) for p in ps]
        expect(cuts == [2105.0, 3160.0, 2460.0, 2605.0], 'kesimler (paylar otomatik): %s' % cuts)
    finally:
        e.close()
    e = Env()
    try:
        e.rect(0, 0, 10000, 1200)
        e.run('ac2fpanelgroove', '7.5cm out', '3200 j30', 'at 2000 5140 7640')
        cuts = [round(p[1][2] - p[1][0], 1) for p in panels(e)]
        expect(cuts == [2105.0, 3200.0, 2560.0, 2465.0], 'belgedeki ornek: at 2000 5140 7640 = n200cm k320cm n250cm * (%s)' % cuts)
    finally:
        e.close()
    e = Env()
    try:
        e.rect(0, 0, 1200, 10000)
        e.run('ac2fpanelgroove', '7.5cm out', '3200 j30', 'at 2000 5100 7500')
        ps = panels(e)
        ys = [round(p[2]['2 top'][0][1], 1) for p in ps]
        expect(len(ps) == 4 and ys == [10000.0, 8000.0, 4900.0, 2500.0], 'dik panelde konum ustten olculur: ust derzler %s' % ys)
    finally:
        e.close()
    for text, frag in (('at 2000 1000', 'does not fit'), ('at 2000 10000', 'does not fit'),
                       ('at 2000 5100 *', 'plain positions'), ('at', 'Type an option')):
        e = Env()
        try:
            e.rect(0, 0, 10000, 1200)
            asked, msgs = e.run('ac2fpanelgroove', '7.5cm out', '3200 j30', text, '0')
            expect(any(frag in m for m in msgs), '"%s" -> uyari "%s"' % (text, frag))
        finally:
            e.close()
    e = Env()
    try:
        e.rect(0, 0, 4160, 1000)
        e.run('ac2fpanelgroove', '5cm in', '3200 j30', 'at 3120')
        ps = panels(e)
        expect(len(ps) == 2 and close(ps[0][2]['3 right'][0][0], 3170) and close(ps[1][2]['3 right'][0][0], 4050),
               '"in": konum net dikdortgenin basindan (x=50) olculur: birlesim 3170, son 4050')
    finally:
        e.close()


def t_confirm_no_and_errors():
    scenario('Onayda Hayir yeniden sorar (yazi hazir gelir); hatali parcalar uyarilir')
    e = Env()
    try:
        e.rect(0, 0, 4000, 1000)
        answers = iter([7, 6])

        def mb(prompt, buttons=0, title=''):
            e.h.messages.append(('msg', prompt))
            return next(answers) if (int(buttons) & 15) == 4 else 1
        e.h.msgbox = mb
        seen = []
        e.h.inputs[:] = ['5cm out', '3200 j30', 'n3000 *', lambda p, d: (seen.append(d), 'n2000 *')[1]]
        e.h.messages[:] = []
        e.p.m.ac2fpanel.f_ac2fpanelgroove()
        ps = panels(e)
        expect(seen == ['1'] or seen == ['n3000 *'], 'Hayir sonrasi ayni soru yeniden: hazir yazi %r' % (seen,))
        expect(len(ps) == 2 and close(faces(ps[0])[0], 2000), 'ikinci yanitla cizildi: ilk yuz 2000')
    finally:
        e.close()
    for text, frag in (('n3500 *', 'No plate in your list is long enough'), ('k50 *', 'no longer than its own folds'),
                       ('n3000 n900', 'The faces add up to 3900'), ('n4000 *', 'nothing is left'),
                       ('x2000 *', 'Could not read')):
        e = Env()
        try:
            e.rect(0, 0, 4000, 1000)
            asked, msgs = e.run('ac2fpanelgroove', '5cm out', '3200 j30', text, '0')
            expect(any(frag in m for m in msgs), '"%s" -> uyari "%s": %s' % (text, frag, [m[:70] for m in msgs[:1]]))
        finally:
            e.close()
    e = Env()
    try:
        e.rect(0, 0, 4000, 1000)
        asked, msgs = e.run('ac2fpanelgroove', '5cm out', '3200 j30', 'n3000 *')
        ps = panels(e)
        expect(len(ps) == 2 and close(faces(ps[0])[0], 3000) and close(faces(ps[1])[0], 1000), 'n3000 *: yuzler 3000 + kalan 1000')
        expect(all(len(m) <= 1000 for m in asked), 'sorular 1000 karakteri asmiyor (%s)' % [len(m) for m in asked])
    finally:
        e.close()


def main():
    t_single_unchanged()
    t_user_example()
    t_equal_option()
    t_three_pieces()
    t_custom_lengths()
    t_joint_values()
    t_sliver()
    t_in_direction()
    t_tall_panel()
    t_no_split_when_it_fits()
    t_zero_and_cancel()
    t_bad_input()
    t_multi_plate_lengths()
    t_many_rectangles()
    t_default_text()
    t_prompt_limits()
    t_waste_order()
    t_in_place_in_direction()
    t_sep_layout()
    t_layers_reused()
    t_net_and_cut_pieces()
    t_positions()
    t_confirm_no_and_errors()
    print('\n%s' % ('HATA: %d' % len(FAILS) if FAILS else 'Hepsi gecti'))
    sys.exit(1 if FAILS else 0)


if __name__ == '__main__':
    main()
