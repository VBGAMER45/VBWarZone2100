"""
ai_skirmish.py - plays a headless AI skirmish in Warzone.exe (cluster C's milestone) and summarises it.

  python -I tools\\ai_skirmish.py [test] [timeout seconds] [extra Warzone.exe arguments...]

* test: a test file in tools/skirmish_tests (default ai4_bloat_spectator.json: Nexus, NullBot,
  BoneCrusher and Cobra on Bloat, free for all, player 0 a spectator so that its loss does not end
  the game). The format is the game's tests/<name>.json (multiint.cpp's loadMapChallengeSettings and
  loadMapPlayerSettings): a "challenge" group (map, maxPlayers, scavengers, alliances, powerLevel,
  bases counted from 1) and player_N groups (team, ai: the js name as multiplay/skirmish/*.json give
  it, difficulty, spectator). Under --skirmish ... --autogame the local player 0 gets no AI.
* Runs Warzone.exe --configdir=build/ai_skirmish/<test>/config --skirmish=<test> --autogame --headless --nosound
  (the test copied into the config's tests folder) until the game ends (rules.js's end conditions call
  gameOverMessage, which quits an autogame) or the timeout.
* Prints the exit code and time, the log's lines by level and function (script errors are
  "jsNativeFunction" VB errors and "callFunction" uncaught exceptions), the log's last lines and the
  scripts' performance logs (logs/<script>.<player>.log, written at a clean quit).
"""
import glob
import os
import shutil
import subprocess
import sys
import time

ROOT = r"C:\Users\Owner\desktop\websites\warzone"
SP = os.path.join(os.path.dirname(os.path.abspath(__file__)), "skirmish_tests")
test = sys.argv[1] if len(sys.argv) > 1 else "ai4_bloat_spectator.json"
limit = int(sys.argv[2]) if len(sys.argv) > 2 else 1800
extra = sys.argv[3:]
work = os.path.join(ROOT, "build", "ai_skirmish", os.path.splitext(os.path.basename(test))[0])
if os.path.exists(work):
    shutil.rmtree(work)
cfg = os.path.join(work, "config")
os.makedirs(os.path.join(cfg, "tests"))
shutil.copyfile(os.path.join(SP, test) if not os.path.isabs(test) else test, os.path.join(cfg, "tests", os.path.basename(test)))
t0 = time.time()
out = open(os.path.join(work, "stdout.txt"), "wb")
p = subprocess.Popen([os.path.join(ROOT, "bin", "Warzone.exe"), "--configdir=" + cfg, "--skirmish=" + os.path.basename(test),
                      "--autogame", "--headless", "--nosound"] + extra, stdout=out, stderr=subprocess.STDOUT)
try:
    p.wait(timeout=limit)
    print("exit code %d after %.0f s" % (p.returncode, time.time() - t0))
except subprocess.TimeoutExpired:
    p.kill()
    p.wait()
    print("stopped after %.0f s (limit)" % (time.time() - t0))
out.close()
logs = sorted(glob.glob(os.path.join(cfg, "logs", "WZlog-*.txt")))
if logs:
    lines = open(logs[-1], encoding="utf-8", errors="replace").read().splitlines()
    print("log: %d lines" % len(lines))
    import collections
    kinds = collections.Counter()
    for l in lines:
        if "|" in l and "[" in l:
            lvl = l.split("|")[0].strip()
            fn = l[l.index("[") + 1:l.index("]")] if "]" in l else "?"
            kinds[(lvl, fn.split(":")[0])] += 1
    for (lvl, fn), n in kinds.most_common(25):
        print("  %6d  %-7s %s" % (n, lvl, fn))
    print("--- last lines:")
    for l in lines[-20:]:
        print("  | " + l[:220])
print("--- stdout:")
print(open(os.path.join(work, "stdout.txt"), "rb").read().decode("utf-8", "replace")[-2500:])
print("--- script logs:", [os.path.basename(f) for f in glob.glob(os.path.join(cfg, "logs", "*.log"))])
