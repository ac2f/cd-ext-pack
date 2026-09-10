#!/usr/bin/env python3
"""ac2f pack - VBA kaynak denetleyicisi.

Blok dengesi, yanlis yerlestirilmis Attribute satirlari, VBA global ad
alaninda cakisan yordam adlari ve tanimsiz ac2f* sembolleri icin src/*.bas
dosyalarini denetler. VBA derleyicisinin yerini tutmaz; ucuz bir on kontroldur.

Kullanim:  python3 tools/lint.py
"""
import re, glob, sys, collections

proc_re  = re.compile(r'^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(Sub|Function|Property\s+\w+)\s+([A-Za-z_]\w*)', re.I)
const_re = re.compile(r'^\s*(?:Public\s+|Private\s+)?Const\s+([A-Za-z_]\w*)', re.I)
type_re  = re.compile(r'^\s*(?:Public\s+|Private\s+)?Type\s+([A-Za-z_]\w*)', re.I)
attr_re  = re.compile(r'^\s*Attribute\s+([A-Za-z_]\w*)\.VB_\w+\s*=', re.I)
decl_re  = re.compile(r'^\s*(?:Public|Private|Dim|Global)\s+(?:WithEvents\s+)?([A-Za-z_]\w*)\s*\(?', re.I)


def strip_comment(ln):
    """Dize disindaki ilk tirnaktan itibarasini atar. VBA'da dize iceren
    satirlarda kesme isareti (ornegin "3'lu") yorum baslatmaz."""
    out, inq = [], False
    for ch in ln:
        if ch == '"':
            inq = not inq
        elif ch == "'" and not inq:
            break
        out.append(ch)
    return ''.join(out)


def preprocess(raw):
    """Yorumlari atar ve VBA satir devamlarini (" _") tek satirda birlestirir.
    Donen her oge (ilk satir numarasi, birlesik kod)."""
    out, buf, anchor = [], None, None
    for i, ln in enumerate(raw, 1):
        code = strip_comment(ln)
        # VBA satir devami BOSLUK + alt cizgi ile biter. Salt "_" ile biten
        # bir tanimlayici (ornegin CAPTION_) devam isareti degildir.
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


OPENERS = {'sub': 'Sub', 'function': 'Function', 'type': 'Type'}


def main():
    files = sorted(glob.glob('src/*.bas'))
    errors, defs, calls = [], {}, collections.defaultdict(list)
    known_syms = set()

    for f in files:
        raw = open(f, encoding='utf-8').read().split('\n')
        mod = f.split('/')[-1][:-4]
        known_syms.add(mod)
        if not raw[0].startswith('Attribute VB_Name'):
            errors.append(f'{f}:1 ilk satir "Attribute VB_Name" olmali')

        stack, prev_proc = [], None
        for i, code in preprocess(raw):
            s = code.strip()
            if not s:
                continue
            low = s.lower()

            m = proc_re.match(code)
            if m:
                kind, name = m.group(1).split()[0].capitalize(), m.group(2)
                if name in defs:
                    errors.append(f'{f}:{i} "{name}" zaten {defs[name]} icinde tanimli '
                                  f'(VBA global ad alani)')
                defs[name] = mod
                stack.append((kind, name, i))
                prev_proc = name
                continue

            am = attr_re.match(code)
            if am:
                if am.group(1) != prev_proc:
                    errors.append(f'{f}:{i} Attribute "{am.group(1)}" yanlis yerde '
                                  f'(beklenen {prev_proc})')
                continue

            cm = const_re.match(code)
            if cm:
                known_syms.add(cm.group(1))
            tm = type_re.match(code)
            if tm:
                known_syms.add(tm.group(1))
                stack.append(('Type', tm.group(1), i))
                continue
            dm = decl_re.match(code)
            if dm:
                known_syms.add(dm.group(1))

            if low.startswith('end sub'):
                if not stack or stack[-1][0] != 'Sub':
                    errors.append(f'{f}:{i} eslesmeyen End Sub')
                else:
                    stack.pop()
            elif low.startswith('end function'):
                if not stack or stack[-1][0] != 'Function':
                    errors.append(f'{f}:{i} eslesmeyen End Function')
                else:
                    stack.pop()
            elif low.startswith('end type'):
                if not stack or stack[-1][0] != 'Type':
                    errors.append(f'{f}:{i} eslesmeyen End Type')
                else:
                    stack.pop()
            elif re.match(r'^if\b.*\bthen$', low):        # yalnizca blok If
                stack.append(('If', 'if', i))
            elif low.startswith('end if'):
                if not stack or stack[-1][0] != 'If':
                    errors.append(f'{f}:{i} eslesmeyen End If')
                else:
                    stack.pop()
            elif re.match(r'^select\s+case\b', low):
                stack.append(('Select', 'select', i))
            elif low.startswith('end select'):
                if not stack or stack[-1][0] != 'Select':
                    errors.append(f'{f}:{i} eslesmeyen End Select')
                else:
                    stack.pop()
            elif re.match(r'^for\b', low):
                stack.append(('For', 'for', i))
            elif re.match(r'^next\b', low):
                if not stack or stack[-1][0] != 'For':
                    errors.append(f'{f}:{i} eslesmeyen Next')
                else:
                    stack.pop()
            elif re.match(r'^do\b', low) or re.match(r'^while\b', low):
                stack.append(('Do', 'do', i))
            elif re.match(r'^(loop|wend)\b', low):
                if not stack or stack[-1][0] != 'Do':
                    errors.append(f'{f}:{i} eslesmeyen Loop')
                else:
                    stack.pop()

            for name in re.findall(r'\bac2f[A-Za-z_]\w*', code):
                calls[name].append(f'{f}:{i}')

        if stack:
            errors.append(f'{f}: kapanmamis blok(lar): '
                          + ', '.join(f'{k} {n} (satir {l})' for k, n, l in stack))

    known = set(defs) | known_syms
    for name, where in sorted(calls.items()):
        if name not in known and not name.startswith('AC2F_'):
            errors.append(f'{where[0]} tanimsiz ac2f sembolu: {name}')

    print(f'{len(files)} dosya, {len(defs)} yordam')
    if errors:
        print('\nSORUNLAR:')
        for e in errors:
            print('  ' + e)
        return 1
    print('Sorun bulunamadi.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
