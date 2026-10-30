"""VB6 name-table probe: does VB6 compile a project with N distinct identifiers?

A port of Warzone 2100 (about 45,000 distinct names in its C++, case-folded)
would need far more names than QuakeVB (about 5,600). VB6 keeps one
case-insensitive name table per project. This writes a synthetic project with
N distinct names spread over M modules, mixed roughly like ported code
(public functions with a parameter and a local, constants, UDT fields,
module variables), compiles it with VB6 /make and runs it.

usage: python -I names_gen.py <work dir> <names> <modules> [<names> <modules> ...]
"""
import os
import shutil
import subprocess
import sys
import time

VB6 = r"C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE"

VBP_HEAD = """Type=Exe
Reference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\SysWOW64\\stdole2.tlb#OLE Automation
"""
VBP_TAIL = """Startup="Sub Main"
ExeName32="names.exe"
Name="names"
CompilationType=0
OptimizationType=2
BoundsCheck=0
OverflowCheck=-1
FlPointCheck=-1
FDIVCheck=-1
UnroundedFP=-1
"""


def module_source(m, quota, counter):
    """One module with about `quota` new distinct names; returns (text, name of its first function)."""
    nfunc = max(1, quota // 5)          # each function brings 3 names: itself, a parameter, a local
    nconst = quota * 15 // 100
    nfield = quota * 15 // 100
    nvar = quota - 3 * nfunc - nconst - nfield
    if nvar < 0:
        nvar = 0
    out = ['Attribute VB_Name = "m%d"' % m, "Option Explicit", ""]
    consts = []
    for _ in range(nconst):
        counter[0] += 1
        c = "kc%d" % counter[0]
        consts.append(c)
        out.append("Public Const %s As Long = %d" % (c, counter[0] % 1000))
    if not consts:
        consts = ["0"]
    counter[0] += 1
    tname = "ty%d" % counter[0]
    fields = []
    out.append("Public Type %s" % tname)
    for _ in range(max(1, nfield)):
        counter[0] += 1
        f = "fd%d" % counter[0]
        fields.append(f)
        out.append("    %s As Long" % f)
    out.append("End Type")
    for _ in range(nvar):
        counter[0] += 1
        out.append("Private mv%d As Long" % counter[0])
    out.append("")
    first = None
    for i in range(nfunc):
        counter[0] += 1
        fn = "fn%d" % counter[0]
        counter[0] += 1
        arg = "ar%d" % counter[0]
        counter[0] += 1
        loc = "lo%d" % counter[0]
        first = first or fn
        c = consts[i % len(consts)]
        f = fields[i % len(fields)]
        out += ["Public Function %s(ByVal %s As Long) As Long" % (fn, arg),
                "    Dim %s As Long, v As %s" % (loc, tname),
                "    v.%s = %s + %s" % (f, arg, c),
                "    %s = v.%s * 2" % (loc, f),
                "    %s = %s" % (fn, loc),
                "End Function", ""]
    return "\r\n".join(out) + "\r\n", first


def build(work, names, modules):
    folder = os.path.join(work, "names_%d_%d" % (names, modules))
    if os.path.isdir(folder):
        shutil.rmtree(folder)
    os.makedirs(folder)
    counter = [0]
    vbp = [VBP_HEAD]
    calls = []
    quota = names // modules
    for m in range(modules):
        text, first = module_source(m, quota, counter)
        with open(os.path.join(folder, "m%d.bas" % m), "w", newline="") as f:
            f.write(text)
        vbp.append("Module=m%d; m%d.bas\n" % (m, m))
        calls.append("    s = s + %s(%d)" % (first, m))
    main = ['Attribute VB_Name = "probemain"', "Option Explicit", "",
            "Sub Main()", "    Dim s As Long, f As Integer"] + calls + [
            "    f = FreeFile",
            '    Open App.Path & "\\names.log" For Output As #f',
            '    Print #f, "ran, checksum " & s',
            "    Close #f", "End Sub", ""]
    with open(os.path.join(folder, "probemain.bas"), "w", newline="") as f:
        f.write("\r\n".join(main))
    vbp.append("Module=probemain; probemain.bas\n")
    vbp.append(VBP_TAIL)
    with open(os.path.join(folder, "names.vbp"), "w", newline="\r\n") as f:
        f.write("".join(vbp))
    t0 = time.time()
    try:
        subprocess.run([VB6, "/make", "names.vbp", "/out", "build.txt"], cwd=folder, timeout=1800)
    except subprocess.TimeoutExpired:
        print("%6d names, %4d modules: VB6 /make TIMED OUT" % (counter[0], modules))
        return
    secs = time.time() - t0
    with open(os.path.join(folder, "build.txt"), encoding="latin-1") as f:
        result = " ".join(f.read().split())
    ran = ""
    if "succeeded" in result:
        subprocess.run([os.path.join(folder, "names.exe")], cwd=folder, timeout=120)
        with open(os.path.join(folder, "names.log"), encoding="latin-1") as f:
            ran = f.read().strip()
        shutil.rmtree(folder)       # keep only failures for inspection
    print("%6d names, %4d modules: %s (%.0f s) %s" % (counter[0], modules, result, secs, ran))
    sys.stdout.flush()


def main(argv):
    work = os.path.abspath(argv[0])
    pairs = argv[1:]
    for i in range(0, len(pairs), 2):
        build(work, int(pairs[i]), int(pairs[i + 1]))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
