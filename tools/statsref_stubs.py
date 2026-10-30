"""The functions statsref defines itself (its stand-ins: defined in tools/statsref/*.cpp outside a #line'd
copy of the game's code) whose VB port now lives outside the stubs modules without a vbDisplayOff gate
in its first lines (the vbtest runs must skip those: CONVENTIONS, "Tests of game state").
python -I tools\statsref_stubs.py"""
import glob
import os
import re

ROOT = r"C:\Users\Owner\desktop\websites\warzone"
names = set()
for p in glob.glob(os.path.join(ROOT, "tools", "statsref", "*.cpp")):
    inref = False
    for l in open(p, encoding="utf-8", errors="replace"):
        m = re.match(r'#line \d+ "([^"]+)"', l)
        if m:
            inref = not m.group(1).startswith("tools/statsref")
            continue
        if inref:
            continue
        m = re.match(r"^[A-Za-z_][\w:<>\*&, ]*?[\s\*&]([a-zA-Z_]\w*)\s*\([^;]*$", l)
        if m and m.group(1) not in ("if", "while", "for", "switch", "return", "main"):
            names.add(m.group(1))
defs = {}
for p in glob.glob(os.path.join(ROOT, "src", "game", "*.bas")):
    lines = open(p, encoding="latin-1").read().split("\n")
    for i, l in enumerate(lines):
        m = re.match(r"^(?:Public |Private |Friend )?(?:Sub|Function) (\w+)\b", l)
        if m and m.group(1) in names:
            gated = any("vbDisplayOff" in x or "vbEffectsOff" in x for x in lines[i + 1:i + 6])
            defs.setdefault(m.group(1), []).append((os.path.basename(p), gated))
for n in sorted(defs):
    real = [(f, g) for f, g in defs[n] if not f.startswith("stubs_") and f not in ("vbtest_vb.bas",)]
    if real and not any(g for f, g in real):
        print("%-40s %s" % (n, " ".join(f for f, g in real)))
