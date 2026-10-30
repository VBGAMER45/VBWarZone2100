"""
draw_test.py - cluster D's screenshot test: Warzone.exe --vbdrawtest draws boxes, lines, text in every
font, a WzText, rotated text and the frontend image sheet over the title screen, saves a screenshot
through screen.cpp's screenDumpToDisk (the GL backend's glReadPixels, iV_saveImage_PNG on GDI+) and quits.

  python -I tools\\draw_test.py [name] [--keep]

* Runs bin\\Warzone.exe --configdir=build/draw_test/config --vbdrawtest=<name> (default "drawtest").
* Copies the screenshot to build/draw_test/<name>.png and prints its size and the log's errors.
* If tools/draw_test/<name>.png exists (a screenshot checked by eye), compares the two pixel by pixel
  (the same GPU and driver give the same pixels) and prints how many differ.
"""
import glob
import os
import shutil
import subprocess
import sys
import time

ROOT = r"C:\Users\Owner\desktop\websites\warzone"
args = [a for a in sys.argv[1:] if not a.startswith("--")]
name = args[0] if args else "drawtest"
work = os.path.join(ROOT, "build", "draw_test")
cfg = os.path.join(work, "config")
if os.path.exists(cfg) and "--keep" not in sys.argv:
    shutil.rmtree(cfg)
os.makedirs(cfg, exist_ok=True)
t0 = time.time()
p = subprocess.run([os.path.join(ROOT, "bin", "Warzone.exe"), "--configdir=" + cfg, "--nosound", "--vbdrawtest=" + name],
                   cwd=os.path.join(ROOT, "bin"), timeout=120)
print("exit code %d after %.1f s" % (p.returncode, time.time() - t0))
logs = sorted(glob.glob(os.path.join(cfg, "logs", "WZlog-*.txt")))
if logs:
    for l in open(logs[-1], encoding="utf-8", errors="replace").read().splitlines():
        if l.startswith(("error", "fatal")) or "assert" in l.lower():
            print("  | " + l[:200])
shots = sorted(glob.glob(os.path.join(cfg, "screenshots", "*-%s.png" % name)))
if not shots:
    print("no screenshot")
    sys.exit(1)
out = os.path.join(work, name + ".png")
shutil.copyfile(shots[-1], out)
print("screenshot: %s (%d bytes)" % (out, os.path.getsize(out)))
ref = os.path.join(ROOT, "tools", "draw_test", name + ".png")
if os.path.exists(ref):
    try:
        from PIL import Image
    except ImportError:
        print("(no PIL: not compared with %s)" % ref)
        sys.exit(0)
    a, b = Image.open(out).convert("RGB"), Image.open(ref).convert("RGB")
    if a.size != b.size:
        print("DIFFERENT size %s, reference %s" % (a.size, b.size))
        sys.exit(1)
    pa, pb = a.load(), b.load()
    diff = sum(1 for y in range(a.size[1]) for x in range(a.size[0]) if pa[x, y] != pb[x, y])
    print("%d of %d pixels differ from %s" % (diff, a.size[0] * a.size[1], ref))
    sys.exit(1 if diff else 0)
