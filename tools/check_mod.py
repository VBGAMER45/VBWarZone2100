"""
check_mod.py - finds VB Mod operations that a later * / or \\ takes apart.

  python -I tools/check_mod.py

In VB, Mod binds looser than *, / and \\: a Mod b * c is a Mod (b * c), where
C's a % b * c is (a % b) * c (% is at the level of * and /). So C's
left-to-right arithmetic must be written with parentheses: (a Mod b) * c.
This reports every Mod that is followed, at the same parenthesis depth and
before a + - comparison And Or Mod or the end of the expression, by a *, /
or \\. (feature.cpp's psStats->baseWidth % 2 * TILE_UNITS / 2 was such a
bug.) The left side needs nothing: a * b Mod c is (a * b) Mod c in both.
"""
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

STOP = re.compile(r'\s*(\+|-|=|<|>|<>|<=|>=|\bAnd\b|\bOr\b|\bXor\b|\bMod\b|\bThen\b|\bTo\b|\bStep\b|\bElse\b|,|:|\))', re.I)


def code_of(line):
    """the line without strings and the comment"""
    out = []
    i = 0
    while i < len(line):
        ch = line[i]
        if ch == '"':
            j = line.find('"', i + 1)
            while j != -1 and j + 1 < len(line) and line[j + 1] == '"':
                j = line.find('"', j + 2)
            if j == -1:
                break
            out.append('""')
            i = j + 1
            continue
        if ch == "'":
            break
        out.append(ch)
        i += 1
    return ''.join(out)


def check_line(code):
    hits = []
    for m in re.finditer(r'\bMod\b', code, re.I):
        depth = 0
        i = m.end()
        while i < len(code):
            ch = code[i]
            if ch == '(':
                depth += 1
            elif ch == ')':
                if depth == 0:
                    break
                depth -= 1
            elif depth == 0:
                if ch in '*/\\':
                    hits.append(m.start())
                    break
                if STOP.match(code, i) and code[i] not in ' ':
                    break
            i += 1
    return hits


def main():
    problems = 0
    for d, _, files in sorted(os.walk(os.path.join(ROOT, 'src'))):
        for name in sorted(files):
            if not name.lower().endswith(('.bas', '.cls', '.frm')):
                continue
            path = os.path.join(d, name)
            with open(path, encoding='latin-1') as f:
                for no, line in enumerate(f, 1):
                    code = code_of(line.rstrip('\r\n'))
                    if check_line(code):
                        print('%s:%d: %s' % (os.path.relpath(path, ROOT), no, line.strip()))
                        problems += 1
    print('%d problem(s)' % problems)
    return 1 if problems else 0


if __name__ == '__main__':
    sys.exit(main())
