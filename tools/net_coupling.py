"""How tightly is lib/netplay coupled to the rest of Warzone? (input for WzNet's INetGame)

up:   names declared in src/, gamelib and widget headers that netplay's .cpp files use
      (what INetGame must offer; netlobby and port mapping left out)
down: names declared in netplay headers that src/, gamelib and widget use (WzNet's API)

The match is by name, so common member names (run, initialize, size) show up as noise.

usage: python -I tools\\net_coupling.py ref\\warzone2100
"""
import os
import re
import sys

root = sys.argv[1]
TOK = re.compile(r'\b[A-Za-z_][A-Za-z0-9_]*\b')
DECL = re.compile(r'(?:^|[\s*&:])([A-Za-z_]\w*)\s*\(', re.M)          # function declarations
EXTERN = re.compile(r'^\s*extern\s+[^;(]*?\b([A-Za-z_]\w*)\s*(?:\[[^\]]*\])?\s*;', re.M)
TYPE = re.compile(r'^\s*(?:struct|class|enum(?:\s+class)?)\s+([A-Za-z_]\w*)', re.M)
SKIP = set('if while for switch return sizeof defined static_assert decltype alignof assert ASSERT'.split())
NOT_PORTED = ('netlobby', 'port_mapping', 'stdinreader', 'urlrequest')


def clean(t):
    t = re.sub(r'/\*.*?\*/', ' ', t, flags=re.S)
    t = re.sub(r'//[^\n]*', ' ', t)
    return re.sub(r'"(?:[^"\\\n]|\\.)*"', '""', t)


def read(p):
    with open(p, encoding='utf-8', errors='replace') as f:
        return clean(f.read())


def declared(text):
    names = set(DECL.findall(text)) | set(EXTERN.findall(text)) | set(TYPE.findall(text))
    return {n for n in names if n not in SKIP}


def files(top, exts, skip=()):
    out = []
    for r, _, fs in os.walk(os.path.join(root, top)):
        if '3rdparty' in r:
            continue
        out += [os.path.join(r, f) for f in fs if f.endswith(exts) and not any(s in f for s in skip)]
    return out


net_cpp = files('lib/netplay', ('.cpp',), NOT_PORTED)
net_tokens = set()
for p in net_cpp:
    net_tokens |= set(TOK.findall(read(p)))

print("== up: what netplay (without netlobby, port mapping) uses from src/gamelib/widget ==")
for p in sorted(set(files('src', ('.h',)) + files('lib/gamelib', ('.h',)) + files('lib/widget', ('.h',)))):
    used = sorted(declared(read(p)) & net_tokens)
    if used:
        print("  %-40s %s" % (os.path.relpath(p, root).replace(os.sep, '/'), ", ".join(used)))

net_api = set()
for p in files('lib/netplay', ('.h',)):
    net_api |= declared(read(p))
users = files('src', ('.cpp', '.h')) + files('lib/gamelib', ('.cpp', '.h')) + files('lib/widget', ('.cpp', '.h'))
down = {}
for p in users:
    for n in set(TOK.findall(read(p))) & net_api:
        down[n] = down.get(n, 0) + 1
print("\n== down: netplay names used by src/gamelib/widget: %d distinct ==" % len(down))
print("  " + ", ".join("%s(%d)" % (n, c) for n, c in sorted(down.items(), key=lambda x: -x[1])[:80]))
