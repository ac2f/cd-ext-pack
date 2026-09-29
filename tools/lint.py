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


SHEET_LIMIT = 1000   # VBA InputBox prompt tops out around 1024 characters


def settings_sheet_length():
    """Rebuilds the ac2fSettings sheet the way ac2fSheet does and returns
    its length. The sheet is passed to InputBox as the prompt, so it must
    stay under the VBA limit; adding a setting is what pushes it over."""
    src = 'src/ac2fSettings.bas'
    try:
        raw = open(src, encoding='utf-8').read()
    except OSError:
        return None, 0
    joined = ' '.join(c for _, c in preprocess(raw.split('\n')))
    lab_w = int(re.search(r'LAB_W\s+As\s+Long\s*=\s*(\d+)', joined).group(1))
    val_w = int(re.search(r'VAL_W\s+As\s+Long\s*=\s*(\d+)', joined).group(1))

    calls = re.findall(
        r'ac2fAddSetting\s+\w+\s*,\s*"([^"]*)"\s*,\s*"([^"]*)"\s*,\s*"([^"]*)"', joined)
    if not calls:
        return None, 0

    # Mirrors ac2fSheet: one group per page, two columns, no unit column.
    # The sheet is paged, so what must stay under the limit is the
    # LARGEST page, not the whole table.
    groups = []
    for grp, label, _unit in calls:
        if not groups or groups[-1][0] != grp:
            groups.append((grp, []))
        groups[-1][1].append(label)

    index = '  '.join('%d %s' % (i, g) for i, (g, _) in enumerate(groups, 1))
    worst, n = 0, 0
    for gi, (grp, labels) in enumerate(groups, 1):
        lines = ['SETTINGS  [' + 'M' * 12 + ']', '',
                 '[%s]  page %d/%d' % (grp, gi, len(groups))]
        pending = ''
        for label in labels:
            cell = '%2d %s %s' % (99, label.ljust(lab_w)[:lab_w], '9' * val_w)
            if not pending:
                pending = cell
            else:
                lines.append(pending.ljust(lab_w + val_w + 5) + cell)
                pending = ''
        if pending:
            lines.append(pending)
        lines += ['', index, '',
                  'N=val  ?N=help  #n=page  P=profiles  R=reset  Enter=close']
        worst = max(worst, len('\r\n'.join(lines)))
        n += len(labels)
    return worst, n


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

            # Strings hold file names like "ac2f_plot.plt"; they are data,
            # not symbol references, so scan only the code outside them.
            outside = re.sub(r'"[^"]*"', '""', code)
            for name in re.findall(r'\bac2f[A-Za-z_]\w*', outside):
                calls[name].append(f'{f}:{i}')

        if stack:
            errors.append(f'{f}: kapanmamis blok(lar): '
                          + ', '.join(f'{k} {n} (satir {l})' for k, n, l in stack))

    known = set(defs) | known_syms
    for name, where in sorted(calls.items()):
        if name not in known and not name.startswith('AC2F_'):
            errors.append(f'{where[0]} tanimsiz ac2f sembolu: {name}')

    sheet_len, n_set = settings_sheet_length()
    if sheet_len is not None:
        if sheet_len > SHEET_LIMIT:
            errors.append(f'ac2fSettings en buyuk sayfasi {sheet_len} karakter '
                          f'({n_set} ayar) - InputBox siniri asilir, '
                          f'grubu bolun veya etiketleri kisaltin')

    print(f'{len(files)} dosya, {len(defs)} yordam', end='')
    if sheet_len is not None:
        print(f', en buyuk ayar sayfasi {sheet_len} karakter ({n_set} ayar)', end='')
    print()
    if errors:
        print('\nSORUNLAR:')
        for e in errors:
            print('  ' + e)
        return 1
    print('Sorun bulunamadi.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
