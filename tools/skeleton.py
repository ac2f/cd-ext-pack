#!/usr/bin/env python3
"""ac2f pack - control-flow skeleton extractor.

Prints, for every procedure, the sequence of control-flow keywords it
contains. Renaming identifiers, translating strings and rewriting comments
must NOT change this sequence, so comparing the output before and after a
large mechanical rewrite catches dropped or duplicated logic.

Usage:  python3 tools/skeleton.py [src_dir] > before.txt
"""
import re, glob, sys

sys.path.insert(0, 'tools')
from lint import preprocess, proc_re

KEYWORDS = [
    (re.compile(r'^if\b.*\bthen$', re.I),        'IF'),
    (re.compile(r'^if\b', re.I),                 'IF1'),      # single-line If
    (re.compile(r'^elseif\b', re.I),             'ELIF'),
    (re.compile(r'^else$', re.I),                'ELSE'),
    (re.compile(r'^end if\b', re.I),             '/IF'),
    (re.compile(r'^select\s+case\b', re.I),      'SEL'),
    (re.compile(r'^case\b', re.I),               'CASE'),
    (re.compile(r'^end select\b', re.I),         '/SEL'),
    (re.compile(r'^for\b', re.I),                'FOR'),
    (re.compile(r'^next\b', re.I),               '/FOR'),
    (re.compile(r'^(do|while)\b', re.I),         'DO'),
    (re.compile(r'^(loop|wend)\b', re.I),        '/DO'),
    (re.compile(r'^exit\s+(sub|function)\b', re.I), 'EXIT'),
    (re.compile(r'^exit\s+(for|do)\b', re.I),    'BREAK'),
    (re.compile(r'^on\s+error\b', re.I),         'ONERR'),
    (re.compile(r'^goto\b', re.I),               'GOTO'),
    (re.compile(r'^resume\b', re.I),             'RESUME'),
    (re.compile(r'^redim\b', re.I),              'REDIM'),
    (re.compile(r'^end (sub|function)\b', re.I), 'END'),
]


def skeleton(path):
    raw = open(path, encoding='utf-8').read().split('\n')
    out, cur = [], None
    for _, code in preprocess(raw):
        s = code.strip()
        if not s:
            continue
        m = proc_re.match(code)
        if m:
            cur = []
            out.append((m.group(2), cur))
            continue
        if cur is None:
            continue
        for rx, tok in KEYWORDS:
            if rx.match(s):
                cur.append(tok)
                break
    return out


def main():
    d = sys.argv[1] if len(sys.argv) > 1 else 'src'
    for f in sorted(glob.glob(f'{d}/*.bas')):
        for name, toks in skeleton(f):
            print(f'{name}\t{" ".join(toks)}')


if __name__ == '__main__':
    main()
