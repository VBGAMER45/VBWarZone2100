"""Compile and run the Chakra and OpenGL feasibility probes.

usage: python -I run_probes.py <probe dir> <warzone source dir>

The probes build like QuakeVB's default: native code, no optimization,
bounds checks on.
"""
import os
import subprocess
import sys

VB6 = r"C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE"

VBP = r"""Type=Exe
Reference=*\G{{00020430-0000-0000-C000-000000000046}}#2.0#0#C:\Windows\SysWOW64\stdole2.tlb#OLE Automation
Module={name}; {name}.bas
Startup="Sub Main"
ExeName32="{name}.exe"
Name="probe_{name}"
CompilationType=0
OptimizationType=2
FavorPentiumPro(tm)=0
CodeViewDebugInfo=0
NoAliasing=0
BoundsCheck=0
OverflowCheck=-1
FlPointCheck=-1
FDIVCheck=-1
UnroundedFP=-1
"""


def build_and_run(folder, name, arg):
    path = os.path.join(folder, name + ".bas")
    with open(path, "rb") as f:
        data = f.read()
    with open(path, "wb") as f:
        f.write(data.replace(b"\r\n", b"\n").replace(b"\n", b"\r\n"))
    with open(os.path.join(folder, name + ".vbp"), "w", newline="\r\n") as f:
        f.write(VBP.format(name=name))
    subprocess.run([VB6, "/make", name + ".vbp", "/out", name + "_build.txt"], cwd=folder)
    with open(os.path.join(folder, name + "_build.txt"), encoding="latin-1") as f:
        result = f.read().strip()
    print("== %s: %s" % (name, " ".join(result.split())))
    if "succeeded" not in result:
        return 1
    subprocess.run([os.path.join(folder, name + ".exe"), arg], cwd=folder, timeout=300)
    with open(os.path.join(folder, name + ".log"), encoding="latin-1") as f:
        print(f.read())
    return 0


def main(folder, wz):
    folder = os.path.abspath(folder)
    data = os.path.join(os.path.abspath(wz), "data")
    scripts = []
    for root, _, files in os.walk(data):
        scripts += [os.path.join(root, n) for n in files if n.endswith(".js")]
    listfile = os.path.join(folder, "wz_scripts.txt")
    with open(listfile, "w", newline="\r\n") as f:
        f.write("\n".join(sorted(scripts)) + "\n")
    rc = build_and_run(folder, "chakra", listfile)
    rc |= build_and_run(folder, "glprobe", os.path.join(data, "base", "shaders"))
    return rc


if __name__ == "__main__":
    sys.exit(main(sys.argv[1], sys.argv[2]))
