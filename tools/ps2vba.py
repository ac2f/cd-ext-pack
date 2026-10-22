#!/usr/bin/env python3
"""ac2f pack - nesting pencerelerinin PowerShell betiklerini VBA'ya gomer.

tools/nest_window.ps1 (ayarlar) ya da tools/nest_preview.ps1 (onizleme)
duzenlenir, sonra:

    python3 tools/ps2vba.py

src/ac2fNest.bas'in sonundaki uretilmis bolumu yeniden yazar: her betik
satiri bir Print #f olur, betik ~8 KB'lik yordamlara bolunur (VBA tek
yordamin derlenmis halini 64 KB ile sinirlar) ve ac2fNSWriteScript hangi
pencere istendiyse onun parcalarini sirayla yazar.
"""
import os, sys

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = (('Set', os.path.join(HERE, 'nest_window.ps1')),
           ('Prev', os.path.join(HERE, 'nest_preview.ps1')))
BAS = os.path.join(HERE, '..', 'src', 'ac2fNest.bas')
MARK = ("'---------------------------------------------------------------------\n"
        "' The window scripts, written out each run so they always match this\n")
CHUNK = 8000


def proc(name, chunk):
    out = ['Private Function %s(ByVal f As Integer) As Boolean' % name,
           '    On Error GoTo Fail']
    for ln in chunk:
        if any(ord(c) > 126 for c in ln):
            sys.exit('%s: ASCII disi karakter: %r' % (name, ln))
        if len(ln) > 900:
            sys.exit('%s: satir VBA icin fazla uzun (%d)' % (name, len(ln)))
        out.append('    Print #f, "%s"' % ln.replace('"', '""'))
    out += ['    %s = True' % name, 'Fail:', 'End Function']
    return '\n'.join(out)


def chunks(lines):
    """Pieces of about CHUNK characters, cut at blank lines when possible."""
    out, cur, size = [], [], 0
    for ln in lines:
        cur.append(ln)
        size += len(ln) + 16
        if size >= CHUNK and not ln.strip():
            out.append(cur)
            cur, size = [], 0
        elif size >= CHUNK * 1.5:
            out.append(cur)
            cur, size = [], 0
    if cur:
        out.append(cur)
    return out


def main():
    gen = []
    disp = ['Private Function ac2fNSWriteScript(ByVal which As Long, ByVal f As Integer) As Boolean',
            '    Dim ok As Boolean']
    total = {}
    for i, (tag, path) in enumerate(SCRIPTS, 1):
        lines = open(path, encoding='utf-8').read().replace('\r\n', '\n').rstrip('\n').split('\n')
        total[tag] = len(lines)
        names = []
        for k, ch in enumerate(chunks(lines), 1):
            n = 'ac2fNSPs%s%d' % (tag, k)
            names.append(n)
            gen.append(proc(n, ch))
        disp.append('    %s which = %d Then' % ('If' if i == 1 else 'ElseIf', i))
        disp.append('        ok = %s(f)' % names[0])
        for n in names[1:]:
            disp.append('        If ok Then ok = %s(f)' % n)
    disp += ['    End If', '    ac2fNSWriteScript = ok', 'End Function']
    head = ("' module (1 settings, 2 preview). Generated from tools/nest_window.ps1\n"
            "' and tools/nest_preview.ps1 by tools/ps2vba.py.\n"
            "'---------------------------------------------------------------------\n")
    src = open(BAS, encoding='utf-8').read().replace('\r\n', '\n')
    i = src.index(MARK) + len(MARK)
    src = src[:i] + head + '\n'.join(disp) + '\n\n' + '\n\n'.join(gen) + '\n'
    open(BAS, 'w', encoding='utf-8', newline='\n').write(src)
    print('ac2fNest.bas: ayar penceresi %d, onizleme %d satir' % (total['Set'], total['Prev']))


if __name__ == '__main__':
    main()
