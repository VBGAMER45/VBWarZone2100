"""
ui_test.py - runs a UI test script in Warzone.exe (cluster F and G's tests): --vbscript=<file> runs the
script's commands on the game's frames, one at a time (drawtest_vb's vbscript_frame):

  wait <frames>        move <x> <y>        click <x> <y>        rclick <x> <y>
  key <virtual key>    text <characters>   shot <name>          log <text>         quit
  widgets              (the widget tree in the log: id, class, screen geometry, hidden, children)
  # a comment

The mouse and the keys are window messages posted to the game's window (client coordinates, the game's
screen at 100%); shot takes the game's own screenshot (screen.cpp's screenDumpToDisk).

  python -I tools\\ui_test.py <script> [timeout s] [--sound] [extra Warzone.exe arguments...]

* script: a file in tools/ui_tests (or a path); the run's folder is build/ui_test/<script name>, with a
  fresh config folder each run (--keep-config keeps the last one: a test of what the previous run saved;
  --config-of=<script name> runs on that script's last config folder, as it was left: options_check.txt
  after options.txt). The game runs with --nosound (silent) unless --sound is given (then --sound: the
  config may say sound=0 from an earlier run).
* Copies the screenshots to build/ui_test/<script name>/<shot name>.png, prints the exit code, the
  log's errors and asserts and the script's log lines.
"""
import glob
import os
import shutil
import subprocess
import sys
import time

ROOT = r"C:\Users\Owner\desktop\websites\warzone"
args = sys.argv[1:]
keep = "--keep-config" in args
sound = "--sound" in args
configOf = [a.split("=", 1)[1] for a in args if a.startswith("--config-of=")]
args = [a for a in args if a not in ("--keep-config", "--sound") and not a.startswith("--config-of=")]
script = args[0]
limit = int(args[1]) if len(args) > 1 else 300
extra = args[2:]
if not os.path.isabs(script) and not os.path.exists(script):
    script = os.path.join(ROOT, "tools", "ui_tests", script)
script = os.path.abspath(script)                    # the game runs in bin\
name = os.path.splitext(os.path.basename(script))[0]
work = os.path.join(ROOT, "build", "ui_test", name)
cfg = os.path.join(work, "config")
if configOf:
    cfg = os.path.join(ROOT, "build", "ui_test", os.path.splitext(configOf[0])[0], "config")
elif os.path.exists(cfg) and not keep:
    shutil.rmtree(cfg)
os.makedirs(cfg, exist_ok=True)
for old in glob.glob(os.path.join(work, "*.png")):
    os.remove(old)
t0 = time.time()
p = subprocess.Popen([os.path.join(ROOT, "bin", "Warzone.exe"), "--configdir=" + cfg, "--vbscript=" + script,
                      "--sound" if sound else "--nosound"] + extra,
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
    for l in lines:
        if "(VB) script" in l or "[dumpWidgets:" in l or l.startswith(("error", "fatal")) or "assert" in l.lower():
            print("  | " + l[:220])
    print("log: %d lines (%s)" % (len(lines), logs[-1]))
for shot in sorted(glob.glob(os.path.join(cfg, "screenshots", "*.png"))):
    if os.path.getmtime(shot) < t0:
        continue                                    # an earlier run's (a kept config folder)
    base = os.path.basename(shot)
    # wz2100-<date>_<time>-<name>.png
    shotname = base.split("-", 2)[2][:-4] if base.count("-") >= 2 else base[:-4]
    out = os.path.join(work, shotname + ".png")
    os.makedirs(work, exist_ok=True)                # (with --config-of, the config folder is another's)
    shutil.copyfile(shot, out)
    print("screenshot: %s" % out)
