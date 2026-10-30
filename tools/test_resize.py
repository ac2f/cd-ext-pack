#!/usr/bin/env python3
"""ac2fResize (her nesneyi ayri ayri olcuye getirme) CorelDRAW taklidi uzerinde sinamasi.

src/ac2fCore.bas + ac2fResize.bas, tools/vbapy.py ile cevrilir ve
tools/cdrmock.py belgesinde calistirilir. Makro bittikten sonra belgedeki
GERCEK sekillere bakilir:

  - WxH: her nesne kendi ayri olcusune gelir (secim tek blok olarak degil)
  - w / h: tek kenar verilir, oran korunur; fit: kutuya sigdirir
  - birimler (mm cm m in), bosluklu ve * ile yazimlar, ondalik virgul
  - @1..@9: o nokta yerinde kalir; taklit CorelDRAW'in yeniden boyutlandirmayi
    baska bir noktadan yaptigi ve konturu olcuye kattigi durumlarda da
  - grup tek nesnedir; "deep" ile uyeleri tek tek boyutlanir
  - duz cizgi: tek kenari olceklenir, olmayan kenari istenirse dokunulmaz
  - kilitli nesne hata sayilir, digerleri yine boyutlanir; tek Geri Al
  - hatali giris yeniden sorulur, ayni metin varsayilan olarak geri gelir
  - son yanit hatirlanir (ac2fResizeEachAgain); InputBox 1000 karakteri asmaz

Kullanim:  python3 tools/test_resize.py
"""
import os, shutil, sys, tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from vbapy import Project
import cdrmock as cm

SRC = [os.path.join(HERE, '..', 'src', n) for n in ('ac2fCore.bas', 'ac2fResize.bas')]
FAILS = []
REG = ('ac2fPack', 'Ayarlar')


def expect(cond, msg):
    print(('  ok   ' if cond else '  FAIL ') + msg)
    if not cond:
        FAILS.append(msg)


def scenario(name):
    print('\n' + name)


def near(a, b, tol=0.01):
    return abs(a - b) <= tol


class Env(object):
    def __init__(self):
        self.tmp = tempfile.mkdtemp(prefix='ac2f_')
        self.h = cm.Host(self.tmp, (9000.0, 9000.0))
        self.h.stubs = {'ac2fsettings': lambda: None}
        self.p = Project(SRC, self.h)
        bad = [w for w in self.p.warnings if not w.startswith('ac2fcore.')]
        if bad:
            raise SystemExit('\n'.join(bad))
        self.doc = self.h.doc

    def rect(self, x, y, w, h, select=True):
        return self.doc.add(cm.Shape(self.doc, cm.CONSTS['cdrrectangleshape'],
                                     [cm.rect_path(x, y, x + w, y + h)]), select)

    def ellipse(self, cx, cy, rx, ry, select=True):
        return self.doc.add(cm.Shape(self.doc, cm.CONSTS['cdrcurveshape'],
                                     [cm.ellipse_path(cx, cy, rx, ry)]), select)

    def line(self, x0, y0, x1, y1, select=True):
        return self.doc.add(cm.Shape(self.doc, cm.CONSTS['cdrcurveshape'],
                                     [cm.Path([(x0, y0), (x1, y1)], closed=False)]), select)

    def run(self, macro, *answers):
        """Makroyu verilen InputBox yanitlariyla calistirir; (sorular, mesajlar)."""
        self.h.inputs[:] = list(answers)
        self.h.messages[:] = []
        getattr(self.p.m.ac2fresize, 'f_' + macro)()
        asked = [m for k, m in self.h.messages if k == 'input']
        msgs = [m for k, m in self.h.messages if k == 'msg']
        return asked, msgs

    def each(self, *answers):
        return self.run('ac2fresizeeach', *answers)

    def last(self):
        return self.h.registry.get((REG[0], REG[1], 'rzson'))

    def close(self):
        shutil.rmtree(self.tmp, ignore_errors=True)


def box(s):
    b = s.bbox()
    return b[0], b[1], b[2] - b[0], b[3] - b[1]


def size(s):
    b = box(s)
    return b[2], b[3]


def centre(s):
    b = s.bbox()
    return (b[0] + b[2]) / 2.0, (b[1] + b[3]) / 2.0


def sizes_are(shapes, w, h):
    return all(near(size(s)[0], w) and near(size(s)[1], h) for s in shapes)


# --------------------------------------------------------------------
def t_exact_each():
    scenario('50x30: her nesne ayri olcuye gelir, merkezleri yerinde')
    e = Env()
    a = e.rect(100, 100, 100, 50)
    b = e.rect(400, 300, 30, 30)
    c = e.rect(900, 50, 200, 80)
    before = [centre(s) for s in (a, b, c)]
    asked, msgs = e.each('50x30')
    expect(sizes_are([a, b, c], 50, 30),
           'uc nesne de 50x30: %s' % [tuple(round(v, 2) for v in size(s)) for s in (a, b, c)])
    after = [centre(s) for s in (a, b, c)]
    expect(all(near(p[0], q[0], 0.02) and near(p[1], q[1], 0.02) for p, q in zip(before, after)),
           'merkezler yerinde kaldi')
    expect(len(e.doc.selection) == 3, 'secim korundu')
    expect(any('Resized            : 3 object(s)' in m for m in msgs), 'rapor: 3 nesne boyutlandi')
    expect(e.last() == '50x30', 'son yanit saklandi: %r' % e.last())
    expect(e.doc.cmd_depth == 0, 'komut grubu kapandi')
    e.close()


def t_one_side():
    scenario('w50 ve h30: tek kenar verilir, oran korunur')
    e = Env()
    a = e.rect(0, 0, 100, 50)
    b = e.rect(500, 0, 30, 60)
    ra, rb = 100 / 50.0, 30 / 60.0
    e.each('w50')
    expect(near(size(a)[0], 50) and near(size(a)[1], 25), 'a: 50 x 25')
    expect(near(size(b)[0], 50) and near(size(b)[1], 100), 'b: 50 x 100')
    expect(near(size(a)[0] / size(a)[1], ra) and near(size(b)[0] / size(b)[1], rb), 'oranlar korundu')
    e.each('h30')
    expect(near(size(a)[1], 30) and near(size(a)[0], 60), 'a: 60 x 30')
    expect(near(size(b)[1], 30) and near(size(b)[0], 15), 'b: 15 x 30')
    e.each('50x')
    expect(near(size(a)[0], 50) and near(size(b)[0], 50), '"50x": yalniz genislik')
    e.each('x20')
    expect(near(size(a)[1], 20) and near(size(b)[1], 20), '"x20": yalniz yukseklik')
    e.each('w50 h30')
    expect(sizes_are([a, b], 50, 30), '"w50 h30" = 50x30')
    e.close()


def t_fit():
    scenario('fit: sekil korunur, kutuya sigar')
    e = Env()
    a = e.rect(0, 0, 200, 50)        # genislik sinirlar
    b = e.rect(500, 0, 20, 100)      # yukseklik sinirlar
    c = e.ellipse(900, 900, 40, 40)  # kare gibi
    e.each('fit 50x30')
    wa, ha = size(a)
    wb, hb = size(b)
    expect(near(wa, 50) and near(ha, 12.5), 'a genislige dayandi: %.2f x %.2f' % (wa, ha))
    expect(near(hb, 30) and near(wb, 6), 'b yukseklige dayandi: %.2f x %.2f' % (wb, hb))
    expect(all(size(s)[0] <= 50.01 and size(s)[1] <= 30.01 for s in (a, b, c)), 'hepsi kutunun icinde')
    expect(near(size(c)[0], size(c)[1]) and near(size(c)[1], 30), 'daire 30 x 30')
    e.each('fit 80')
    expect(near(max(size(a)), 80) and near(max(size(b)), 80) and near(max(size(c)), 80),
           '"fit 80": en uzun kenar 80')
    e.each('fit w40')
    expect(near(size(a)[0], 40) and near(size(b)[0], 40), '"fit w40": yalniz genislik siniri')
    e.close()


def t_units_and_forms():
    scenario('birimler ve yazim bicimleri')
    forms = [('12cm x 4cm', 120, 40), ('12cm*4cm', 120, 40), ('50 mm x 30 mm', 50, 30),
             ('50 x 30', 50, 30), ('50 x30', 50, 30), ('50x 30', 50, 30),
             ('12,5x30', 12.5, 30), ('12.5x30', 12.5, 30), ('0,5mx200', 500, 200),
             ('W50 H30', 50, 30), ('2in x 1in', 50.8, 25.4), ('  50X30  ', 50, 30),
             ('5 cm x 30', 50, 30), ('50;30'.replace(';', 'x'), 50, 30)]
    for text, w, h in forms:
        e = Env()
        a = e.rect(0, 0, 77, 33)
        e.each(text)
        expect(sizes_are([a], w, h), '%-16r -> %.2f x %.2f (aranan %g x %g)' % ((text,) + size(a) + (w, h)))
        e.close()
    e = Env()
    a = e.rect(0, 0, 77, 33)
    e.each('fit 5cm')
    expect(near(max(size(a)), 50), '"fit 5cm": en uzun kenar 50')
    e.each('h 3 cm')
    expect(near(size(a)[1], 30), '"h 3 cm" bosluklu birim')
    e.close()


def t_anchor():
    scenario('@1..@9: o nokta yerinde kalir (taklit CorelDRAW baska noktadan boyutlar)')
    pos = {1: (0, 0), 2: (.5, 0), 3: (1, 0), 4: (0, .5), 5: (.5, .5), 6: (1, .5),
           7: (0, 1), 8: (.5, 1), 9: (1, 1)}
    for taklit in ('tl', 'br', 'c'):
        bad = []
        for p, (fx, fy) in pos.items():
            e = Env()
            a = e.rect(100, 200, 80, 40)
            a.size_anchor = taklit
            x, y, w, h = box(a)
            keep = (x + fx * w, y + fy * h)
            e.each('50x30 @%d' % p)
            x2, y2, w2, h2 = box(a)
            got = (x2 + fx * w2, y2 + fy * h2)
            if not (near(got[0], keep[0], 0.02) and near(got[1], keep[1], 0.02) and near(w2, 50) and near(h2, 30)):
                bad.append((p, got, keep))
            e.close()
        expect(not bad, 'taklit %s: dokuz noktanin dokuzu yerinde kaldi %s' % (taklit, bad[:2]))
    e = Env()
    a = e.rect(100, 200, 80, 40)
    e.each('50x30 @7')
    expect(near(box(a)[0], 100, 0.02) and near(box(a)[1] + box(a)[3], 240, 0.02), '@7: sol ust (100, 240) sabit')
    e.each('60x20')                      # @ yazilmazsa varsayilan @5, onceki @7 degil
    cx, cy = centre(a)
    expect(near(cx, 125, 0.02) and near(cy, 225, 0.02), '@ yazilmayinca merkez sabit: (%.2f, %.2f)' % (cx, cy))
    e.close()


def t_outline_pad():
    scenario('CorelDRAW konturu olcuye katarsa: okunup duzeltilir')
    e = Env()
    a = e.rect(0, 0, 100, 50)
    b = e.rect(500, 0, 30, 30)
    for s in (a, b):
        s.size_pad = 0.6
        s.size_anchor = 'br'
    asked, msgs = e.each('50x30 @5')
    expect(sizes_are([a, b], 50, 30), 'gercek (sinir kutusu) olcu 50x30: %s %s' % (size(a), size(b)))
    expect(not any('Not exactly' in m for m in msgs), 'raporda "ulasilamadi" yok')
    e.each('w80')
    expect(near(size(a)[0], 80) and near(size(b)[0], 80), 'w80 de kontur payli ayni sekilde')
    e.close()


def t_groups():
    scenario('grup tek nesnedir; deep ile uyeleri tek tek boyutlanir')
    e = Env()
    a = e.rect(0, 0, 40, 20, select=False)
    b = e.rect(100, 0, 20, 60, select=False)
    g = e.doc.group_shapes([a, b])
    gb = box(g)
    e.each('w100')
    nb = box(g)
    expect(near(nb[2], 100) and near(nb[3], gb[3] * 100 / gb[2]), 'grup tek blok: %.2f x %.2f' % (nb[2], nb[3]))
    expect(near(size(a)[0] / size(b)[0], 2.0), 'uyelerin birbirine orani degismedi')

    e2 = Env()
    a = e2.rect(0, 0, 40, 20, select=False)
    b = e2.rect(100, 0, 20, 60, select=False)
    c = e2.rect(300, 0, 10, 10, select=False)
    g = e2.doc.group_shapes([a, b])
    e2.doc.selection = [g, c]
    asked, msgs = e2.each('50x30 deep')
    expect(sizes_are([a, b, c], 50, 30), 'deep: grup uyeleri ve disaridaki nesne 50x30')
    expect(len(e2.doc.selection) == 2, 'secim hala grup + nesne')
    expect(any('inside groups' in m for m in msgs), 'raporda deep belirtildi')
    e.close()
    e2.close()


def t_flat_lines():
    scenario('duz cizgi: tek kenari olceklenir, olmayan kenar istenirse dokunulmaz')
    e = Env()
    ln = e.line(0, 100, 200, 100)        # yatay: yuksekligi yok
    r = e.rect(500, 0, 40, 20)
    asked, msgs = e.each('w50')
    expect(near(size(ln)[0], 50) and near(size(ln)[1], 0), 'yatay cizgi: boy 50, yukseklik 0')
    expect(near(size(r)[0], 50), 'dikdortgen de 50')
    asked, msgs = e.each('h30')
    expect(near(size(ln)[0], 50), 'h30: cizgiye dokunulmadi')
    expect(near(size(r)[1], 30), 'h30: dikdortgen 30')
    expect(any('Left alone         : 1' in m for m in msgs), 'rapor: 1 nesneye dokunulmadi')
    asked, msgs = e.each('40x40')
    expect(near(size(ln)[0], 50) and near(size(r)[0], 40) and near(size(r)[1], 40),
           '40x40: cizgi ayni, dikdortgen 40x40')
    vl = e.line(700, 0, 700, 80)
    e.doc.selection = [vl]
    e.each('h20')
    expect(near(size(vl)[1], 20) and near(size(vl)[0], 0), 'dikey cizgi: yukseklik 20')
    e.close()


def t_same_size():
    scenario('zaten o olcude olan nesneye dokunulmaz')
    e = Env()
    a = e.rect(0, 0, 50, 30)
    b = e.rect(200, 0, 60, 30)
    asked, msgs = e.each('50x30')
    expect(any('Already that size  : 1' in m for m in msgs), 'rapor: 1 nesne zaten o olcude')
    expect(any('Resized            : 1 object(s)' in m for m in msgs), 'rapor: 1 nesne boyutlandi')
    asked, msgs = e.each('50x30')
    expect(any('Resized            : 0 object(s)' in m for m in msgs) and
           any('Already that size  : 2' in m for m in msgs), 'ikinci calistirma: hicbiri degismedi')
    e.close()


def t_locked_and_undo():
    scenario('kilitli nesne hata sayilir, digerleri boyutlanir; tek Geri Al')
    e = Env()
    a = e.rect(0, 0, 100, 50)
    b = e.rect(300, 0, 100, 50)
    c = e.rect(600, 0, 100, 50)
    b.locked = True
    asked, msgs = e.each('40x40')
    expect(sizes_are([a, c], 40, 40) and near(size(b)[0], 100), 'a ve c boyutlandi, kilitli b ayni')
    expect(any('Could not resize   : 1 object(s): The object is locked' in m for m in msgs),
           'rapor: kilitli nesne nedeniyle birlikte')
    expect(e.doc.cmd_depth == 0, 'komut grubu kapandi')
    e.doc.undo()
    expect(sizes_are([a, b, c], 100, 50), 'tek Geri Al hepsini geri aldi')
    e.close()


def t_unit_restored():
    scenario('belge birimi eski haline doner')
    e = Env()
    e.doc.unit = cm.CONSTS['cdrinch']
    e.rect(0, 0, 100, 50)
    e.each('50x30')
    expect(e.doc.unit == cm.CONSTS['cdrinch'], 'birim inch olarak geri geldi')
    expect(e.h.application.optimization is False, 'Optimization eski haline dondu')
    e.close()


def t_bad_input():
    scenario('hatali giris: nedeni soylenir, ayni metin geri gelir, duzeltince calisir')
    cases = [('50', 'lone number'), ('abc', 'cannot read'), ('w0', 'above zero'),
             ('50x30 @0', 'digit from 1 to 9'), ('50x30 @55', 'digit from 1 to 9'),
             ('50x30 @', 'digit from 1 to 9'), ('x', 'width, a height'), ('w', 'above zero'),
             ('50xabc', 'above zero'), ('deep', 'Type a size'), ('fit', 'Type a size'),
             ('-5x10', 'above zero'), ('5cmx', None)]
    for text, why in cases:
        e = Env()
        a = e.rect(0, 0, 100, 50)
        seen = []

        def second(prompt, default, seen=seen):
            seen.append(default)
            return '40x20'
        asked, msgs = e.each(text, second)
        if why is None:
            expect(near(size(a)[0], 50) and not seen, '%r gecerli (yalniz genislik 5cm)' % text)
        else:
            expect(len(asked) == 2 and seen == [text] and sizes_are([a], 40, 20)
                   and any(why in m for m in msgs),
                   '%-12r -> uyari "%s", ayni metin geri geldi, duzeltince 40x20' % (text, why))
        e.close()
    e = Env()
    a = e.rect(0, 0, 100, 50)
    e.each('10', '7 7', '40x20')
    expect(sizes_are([a], 40, 20), 'ust uste iki hatadan sonra da dogru calisti')
    e.close()


def t_cancel_and_empty():
    scenario('iptal / bos yanit / secim yok: hicbir sey degismez')
    e = Env()
    a = e.rect(0, 0, 100, 50)
    e.each(None)
    expect(sizes_are([a], 100, 50) and e.last() is None, 'iptal: degismedi, yanit saklanmadi')
    e.each('   ')
    expect(sizes_are([a], 100, 50), 'bos yanit: degismedi')
    e.doc.selection = []
    asked, msgs = e.each('50x30')
    expect(not asked and any('Select the objects first' in m for m in msgs),
           'secim yokken sorulmadan uyarilir')
    e.close()


def t_again_and_memory():
    scenario('son yanit hatirlanir; Again sormadan uygular')
    e = Env()
    a = e.rect(0, 0, 100, 50)
    asked, msgs = e.run('ac2fresizeeachagain')
    expect(not asked and any('no earlier size' in m for m in msgs), 'yanit yokken Again uyarir')
    expect(sizes_are([a], 100, 50), 've hicbir sey degismez')
    e.each('h20 @7')
    b = e.rect(700, 700, 90, 90, select=False)
    e.doc.selection = [b]
    top_before = box(b)[1] + box(b)[3]
    asked, msgs = e.run('ac2fresizeeachagain')
    expect(not asked and near(size(b)[1], 20) and near(size(b)[0], 20), 'Again: sormadan h20 uyguladi')
    expect(near(box(b)[1] + box(b)[3], top_before, 0.02) and near(box(b)[0], 700, 0.02),
           'Again: @7 (sol ust) da hatirlandi')
    seen = []
    e.each(lambda prompt, default: (seen.append(default), 'w10')[1])
    expect(seen == ['h20 @7'], 'InputBox varsayilani son yanit: %r' % seen)
    e.close()
    e = Env()
    e.rect(0, 0, 100, 50)
    seen = []
    e.each(lambda prompt, default: (seen.append(default), None)[1])
    expect(seen == ['50x50'], 'ilk soruda varsayilan 50x50: %r' % seen)
    e.close()


def t_prompt_limits():
    scenario('InputBox metni 1000 karakteri asmaz, nesne sayisini soyler')
    e = Env()
    for i in range(7):
        e.rect(i * 100, 0, 50, 50)
    asked, msgs = e.each(None)
    expect(len(asked[0]) < 1000, 'soru metni %d karakter' % len(asked[0]))
    expect(asked[0].startswith('7 object(s) selected'), 'baslik: 7 nesne secili')
    e.close()


def t_many_and_rotated_like():
    scenario('kalabalik secim: 200 nesne, hepsi dogru; ikinci calistirma degistirmez')
    e = Env()
    shapes = [e.rect((i % 20) * 120, (i // 20) * 120, 20 + i % 17, 10 + i % 23) for i in range(200)]
    e.each('fit 60x40 @5')
    ok = all(size(s)[0] <= 60.01 and size(s)[1] <= 40.01 and
             (near(size(s)[0], 60) or near(size(s)[1], 40)) for s in shapes)
    expect(ok, '200 nesne de kutuya sigdi ve bir kenari kutuya dayandi')
    snapshot = [box(s) for s in shapes]
    asked, msgs = e.each('fit 60x40 @5')
    expect(all(near(a, b, 0.001) for s, bx in zip(shapes, snapshot) for a, b in zip(box(s), bx)),
           'ikinci calistirma hicbir seyi oynatmadi')
    e.close()


def main():
    t_exact_each()
    t_one_side()
    t_fit()
    t_units_and_forms()
    t_anchor()
    t_outline_pad()
    t_groups()
    t_flat_lines()
    t_same_size()
    t_locked_and_undo()
    t_unit_restored()
    t_bad_input()
    t_cancel_and_empty()
    t_again_and_memory()
    t_prompt_limits()
    t_many_and_rotated_like()
    print('\n%s' % ('HATA: %d' % len(FAILS) if FAILS else 'Hepsi gecti'))
    sys.exit(1 if FAILS else 0)


if __name__ == '__main__':
    main()
