#!/usr/bin/env python3
"""ac2f pack - VBA kaynak denetleyicisi.

Blok dengesi, yanlis yerlestirilmis Attribute satirlari, VBA global ad
alaninda cakisan yordam adlari ve tanimsiz ac2f* sembolleri icin src/*.bas
dosyalarini denetler. VBA derleyicisinin yerini tutmaz; ucuz bir on kontroldur.

Kullanim:  python3 tools/lint.py
"""
import re, glob, sys, collections

files = sorted(glob.glob('src/*.bas'))
errors, defs, calls = [], {}, collections.defaultdict(list)
consts = {}

proc_re = re.compile(r'^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(Sub|Function|Property\s+\w+)\s+([A-Za-z_]\w*)', re.I)
const_re = re.compile(r'^\s*(?:Public\s+|Private\s+)?Const\s+([A-Za-z_]\w*)', re.I)
type_re  = re.compile(r'^\s*(?:Public\s+|Private\s+)?Type\s+([A-Za-z_]\w*)', re.I)
attr_re  = re.compile(r'^\s*Attribute\s+([A-Za-z_]\w*)\.VB_\w+\s*=', re.I)

for f in files:
    lines = open(f, encoding='utf-8').read().split('\n')
    mod = f.split('/')[-1][:-4]
    if not lines[0].startswith('Attribute VB_Name'):
        errors.append(f'{f}:1 ilk satir "Attribute VB_Name" olmali')
    stack = []
    prev_proc = None
    for i, ln in enumerate(lines, 1):
        s = ln.strip()
        if s.startswith("'") or not s:
            continue
        code = re.sub(r"'.*$", '', ln)  # naive comment strip (no string-aware)
        low = code.strip().lower()

        m = proc_re.match(code)
        if m:
            kind = m.group(1).split()[0].capitalize()
            name = m.group(2)
            if name in defs:
                errors.append(f'{f}:{i} "{name}" zaten {defs[name]} icinde tanimli (VBA global ad alani)')
            defs[name] = f'{mod}'
            stack.append((kind, name, i))
            prev_proc = name
            continue

        am = attr_re.match(code)
        if am:
            if am.group(1) != prev_proc:
                errors.append(f'{f}:{i} Attribute "{am.group(1)}" yanlis yerde (beklenen {prev_proc})')
            continue

        cm = const_re.match(code)
        if cm: consts[cm.group(1)] = mod
        tm = type_re.match(code)
        if tm:
            consts[tm.group(1)] = mod
            stack.append(('Type', tm.group(1), i)); continue

        if low.startswith('end sub'):
            if not stack or stack[-1][0] != 'Sub': errors.append(f'{f}:{i} eslesmeyen End Sub')
            else: stack.pop()
        elif low.startswith('end function'):
            if not stack or stack[-1][0] != 'Function': errors.append(f'{f}:{i} eslesmeyen End Function')
            else: stack.pop()
        elif low.startswith('end type'):
            if not stack or stack[-1][0] != 'Type': errors.append(f'{f}:{i} eslesmeyen End Type')
            else: stack.pop()

        # block If: "If ... Then" with nothing after Then
        if re.match(r'^\s*if\b.*\bthen\s*$', low): stack.append(('If', 'if', i))
        elif low.startswith('end if'):
            if not stack or stack[-1][0] != 'If': errors.append(f'{f}:{i} eslesmeyen End If')
            else: stack.pop()
        elif re.match(r'^\s*select\s+case\b', low): stack.append(('Select', 'select', i))
        elif low.startswith('end select'):
            if not stack or stack[-1][0] != 'Select': errors.append(f'{f}:{i} eslesmeyen End Select')
            else: stack.pop()
        elif re.match(r'^\s*for\b', low) and not low.startswith('format'): stack.append(('For', 'for', i))
        elif re.match(r'^\s*next\b', low):
            if not stack or stack[-1][0] != 'For': errors.append(f'{f}:{i} eslesmeyen Next')
            else: stack.pop()
        elif re.match(r'^\s*(do|while)\b', low): stack.append(('Do', 'do', i))
        elif re.match(r'^\s*(loop|wend)\b', low):
            if not stack or stack[-1][0] != 'Do': errors.append(f'{f}:{i} eslesmeyen Loop')
            else: stack.pop()

        for name in re.findall(r'\bac2f[A-Za-z_]\w*', code):
            calls[name].append(f'{f}:{i}')
    if stack:
        errors.append(f'{f}: kapanmamis blok(lar): {stack}')

modules = {f.split('/')[-1][:-4] for f in files}
known = set(defs) | set(consts) | modules
for name, where in sorted(calls.items()):
    if name not in known and not name.startswith('AC2F_'):
        errors.append(f'{where[0]} tanimsiz ac2f sembolu: {name}')

print(f'{len(files)} dosya, {len(defs)} yordam, {len(consts)} sabit/tip')
print('\n'.join(f'  {n:26s} {m}' for n, m in sorted(defs.items())))
print()
if errors:
    print('SORUNLAR:'); [print('  ' + e) for e in errors]; sys.exit(1)
print('Sorun bulunamadi.')
