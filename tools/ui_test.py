"""
ui_test.py - runs a UI test script in Warzone.exe (cluster F and G's tests): --vbscript=<file> runs the
script's commands on the game's frames, one at a time (drawtest_vb's vbscript_frame):

  wait <frames>        move <x> <y>        click <x> <y>        rclick <x> <y>
  key <virtual key>    text <characters>   shot <name>          log <text>         quit
  widgets              (the widget tree in the log: id, class, screen geometry, hidden, children)
  novideo              (waits while a full-screen video plays)   embark (the units board the transporter)
  quiet <frames>       (waits until that many frames in a row passed without a full-screen video)
  waitlevel <name>     (waits until that level is loaded and the game runs it)
  key <vk> ext         (an extended key: keypad Enter is "key 13 ext")
  keydown <vk>, keyup <vk>  (a key held: "keydown 17", "key 86", "keyup 17" is Ctrl+V)
  closeguide           (closes the Game Guide if a topic popped it up; nothing otherwise)
  # a comment

The mouse and the keys are window messages posted to the game's window (client coordinates, the game's
screen at 100%); shot takes the game's own screenshot (screen.cpp's screenDumpToDisk).

  python -I tools\\ui_test.py <script> [timeout s] [--sound] [--real-network] [extra Warzone.exe arguments...]

* A fresh config folder gets portMapping=0 and lobbyserver=http://127.0.0.1:9 (nothing answers): a test never
  maps a port on the router nor lists a game on the public lobby. --real-network leaves 4.7.0's defaults.

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
TEST_LOBBY = "http://127.0.0.1:9"                   # the discard port: nothing answers
args = sys.argv[1:]
keep = "--keep-config" in args
sound = "--sound" in args
real_network = "--real-network" in args
configOf = [a.split("=", 1)[1] for a in args if a.startswith("--config-of=")]
args = [a for a in args if a not in ("--keep-config", "--sound", "--real-network") and not a.startswith("--config-of=")]
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
# a test never maps a port on the router nor lists a game on the public lobby: a config folder without a config
# file gets port mapping off and a lobby address where nothing listens (--real-network: 4.7.0's defaults)
if not real_network and not os.path.exists(os.path.join(cfg, "config")):
    with open(os.path.join(cfg, "config"), "w", newline="\r\n") as f:
        f.write("[General]\nportMapping=0\nlobbyserver=%s\n" % TEST_LOBBY)
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
