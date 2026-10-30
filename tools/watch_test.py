"""
watch_test.py - cluster E's test: an AI skirmish in a window (Warzone.exe --skirmish=<test> --autogame
--vbwatchshot=<gameTime>,<name>), which logs the frame rate every 100 frames, saves a screenshot through
screen.cpp's screenDumpToDisk at the game time and quits.

  python -I tools\\watch_test.py [gameTime ms] [test] [timeout s] [extra Warzone.exe arguments...]

* gameTime: default 30000 (30 s of game time); test: a tools/skirmish_tests file (default
  ai4_bloat_spectator.json); timeout: default 600.
* --realtime: the game at its normal speed (--autogame runs it 500 times faster than real time as far
  as the frame rate lets it, levels.cpp's gameTimeSetMod(Rational(500)); the run sets it back to 1 on
  its first frame).
* --player=N: the camera to player N's home (multiplay.cpp's cameraToHome) on the first frame.
* Copies the screenshot to build/watch_test/<test>-<gameTime>.png and prints the log's frame-rate
  lines, errors and asserts.
"""
import glob
import os
import shutil
import subprocess
import sys
import time

ROOT = r"C:\Users\Owner\desktop\websites\warzone"
args = sys.argv[1:]
shot = int(args[0]) if len(args) > 0 else 30000
test = args[1] if len(args) > 1 else "ai4_bloat_spectator.json"
limit = int(args[2]) if len(args) > 2 else 600
extra = args[3:]
rt = ",rt" if "--realtime" in extra else ""
rt += "".join(",p" + a.split("=")[1] for a in extra if a.startswith("--player="))
extra = [a for a in extra if a != "--realtime" and not a.startswith("--player=")]
name = "%s-%d" % (os.path.splitext(test)[0], shot)
work = os.path.join(ROOT, "build", "watch_test")
cfg = os.path.join(work, "config")
if os.path.exists(cfg):
    shutil.rmtree(cfg)
os.makedirs(os.path.join(cfg, "tests"))
shutil.copyfile(os.path.join(ROOT, "tools", "skirmish_tests", test), os.path.join(cfg, "tests", test))
t0 = time.time()
p = subprocess.Popen([os.path.join(ROOT, "bin", "Warzone.exe"), "--configdir=" + cfg, "--skirmish=" + test,
                      "--autogame", "--nosound", "--vbwatchshot=%d,%s%s" % (shot, name, rt)] + extra,
                     cwd=os.path.join(ROOT, "bin"))
try:
    p.wait(timeout=limit)
    print("exit code %d after %.0f s" % (p.returncode, time.time() - t0))
except subprocess.TimeoutExpired:
    p.kill()
    p.wait()
    print("stopped after %.0f s (limit)" % (time.time() - t0))
logs = sorted(glob.glob(os.path.join(cfg, "logs", "WZlog-*.txt")))
if logs:
    lines = open(logs[-1], encoding="utf-8", errors="replace").read().splitlines()
    n = 0
    for l in lines:
        if "(VB) watch" in l or l.startswith(("error", "fatal")) or "assert" in l.lower():
            n += 1
            if n <= 60:
                print("  | " + l[:220])
    print("log: %d lines" % len(lines))
shots = sorted(glob.glob(os.path.join(cfg, "screenshots", "*-%s.png" % name)))
if not shots:
    print("no screenshot")
    sys.exit(1)
out = os.path.join(work, name + ".png")
shutil.copyfile(shots[-1], out)
print("screenshot: %s" % out)
