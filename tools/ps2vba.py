#!/usr/bin/env python3
"""ac2f pack - nesting penceresinin PowerShell betigini VBA'ya gomer.

tools/nest_window.ps1 duzenlenir, sonra:

    python3 tools/ps2vba.py

src/ac2fNest.bas icindeki ac2fNSScriptA / ac2fNSScriptB yordamlarini
betigin satirlarindan yeniden uretir (her satir bir Print #f). Betik iki
yordama bolunur ki tek yordam VBA'nin boyut sinirina yaklasmasin.
"""
import os, re, sys

HERE = os.path.dirname(os.path.abspath(__file__))
PS = os.path.join(HERE, 'nest_window.ps1')
BAS = os.path.join(HERE, '..', 'src', 'ac2fNest.bas')
MARK = "' module. Generated from tools/nest_window.ps1 by tools/ps2vba.py.\n'---------------------------------------------------------------------\n"


def proc(name, chunk):
    out = ['Private Function %s(ByVal f As Integer) As Boolean' % name,
           '    On Error GoTo Fail']
    for ln in chunk:
        if any(ord(c) > 126 for c in ln):
            sys.exit('nest_window.ps1: ASCII disi karakter: %r' % ln)
        out.append('    Print #f, "%s"' % ln.replace('"', '""'))
    out += ['    %s = True' % name, 'Fail:', 'End Function']
    return '\n'.join(out)


def main():
    lines = open(PS, encoding='utf-8').read().replace('\r\n', '\n').rstrip('\n').split('\n')
    half = len(lines) // 2
    for k in range(half, len(lines)):
        if not lines[k].strip():
            half = k
            break
    gen = proc('ac2fNSScriptA', lines[:half]) + '\n\n' + proc('ac2fNSScriptB', lines[half:]) + '\n'
    src = open(BAS, encoding='utf-8').read().replace('\r\n', '\n')
    i = src.index(MARK) + len(MARK)
    src = src[:i] + gen
    open(BAS, 'w', encoding='utf-8', newline='\n').write(src)
    print('ac2fNest.bas: %d betik satiri yazildi' % len(lines))


if __name__ == '__main__':
    main()
