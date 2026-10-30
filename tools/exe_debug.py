"""Rebuild only Warzone.exe with CodeView debug info and no optimisation (a .pdb beside the exe), for
stack traces with cdb: python -I tools\exe_debug.py (tools\build.py rebuilds the release exe)"""
import os
import sys

sys.path.insert(0, r"C:\Users\Owner\desktop\websites\warzone\tools")
import build  # noqa: E402

build.write_autorevision()
for name, kind, sub, refs in build.PROJECTS:
    if kind != "exe":
        continue
    folder = os.path.join(build.SRC, sub)
    resfile = build.write_manifest([r for r in refs])
    if resfile is None:
        sys.exit(1)
    vbp = build.write_vbp(name, kind, folder, refs, False, resfile)
    text = open(vbp, encoding="latin-1").read()
    text = text.replace("CodeViewDebugInfo=0", "CodeViewDebugInfo=-1")   # (OptimizationType=2: none already)
    open(vbp, "w", encoding="latin-1").write(text)
    sys.exit(0 if build.make(name, folder, vbp) else 1)
