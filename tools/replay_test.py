"""
replay_test.py - plays an official 4.7.0 replay in Warzone.exe headless and compares
its sync log, tick by tick, with the official build's (cluster B's milestone).

  python -I tools\\replay_test.py [name] [--until GAMETIME] [--timeout SECONDS] [--keep]

* name: a replay in ref/official/replays (without .wzrp; default: every one that
  has reference logs). Its reference logs are ref/official/synclogs/<name>/
  sync<gameTime>_p23.txt, written by the official build with
  warzone2100.exe --loadreplay=<file> --autogame --headless
                  --debug-verbose-sync-logs-until=<gameTime>
  (the replay viewer is player 23: MAX_CONNECTED_PLAYERS).
* Warzone.exe runs the same command line in build/replay_tests/<name>/config
  (the replay copied into its replay/skirmish folder) and writes its own
  logs/sync<gameTime>_p23.txt; each tick's file is compared line by line.
* Reports how many ticks match, and for the first that does not, the first
  differing entry with some context (the BEGIN line's entry count and CRC
  differ whenever an entry does, so the entries are compared first).
  --until defaults to the last reference tick. Warzone.exe is stopped once
  it has written its log of that tick (or after --timeout seconds, 600).
"""
import os
import shutil
import subprocess
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OFFICIAL = os.path.join(ROOT, "ref", "official")
WORK = os.path.join(ROOT, "build", "replay_tests")
EXE = os.path.join(ROOT, "bin", "Warzone.exe")
VIEWER = 23                                              # MAX_CONNECTED_PLAYERS: the replay viewer's slot


def ref_ticks(name):
    d = os.path.join(OFFICIAL, "synclogs", name)
    ticks = []
    for fn in os.listdir(d):
        if fn.startswith("sync") and fn.endswith("_p%d.txt" % VIEWER):
            ticks.append(int(fn[4:fn.index("_")]))
    return sorted(ticks)


def read_lines(path):
    with open(path, "rb") as f:
        return f.read().replace(b"\r\n", b"\n").split(b"\n")


def run(name, until, keep, timeout):
    replay = os.path.join(OFFICIAL, "replays", name + ".wzrp")
    ticks = [t for t in ref_ticks(name) if until is None or t <= until]
    if not ticks:
        print("%s: no reference logs" % name)
        return False
    until = ticks[-1]
    work = os.path.join(WORK, name)
    if os.path.exists(work):
        shutil.rmtree(work)
    cfg = os.path.join(work, "config")
    os.makedirs(os.path.join(cfg, "replay", "skirmish"))
    shutil.copyfile(replay, os.path.join(cfg, "replay", "skirmish", name + ".wzrp"))
    t0 = time.time()
    logs = os.path.join(cfg, "logs")
    last = os.path.join(logs, "sync%d_p%d.txt" % (until, VIEWER))
    # the game goes on after the replay's end (headless: until a time limit or a victory), so it is
    # stopped once its last log is written
    p = subprocess.Popen([EXE, "--configdir=" + cfg, "--loadreplay=" + name, "--autogame", "--headless", "--nosound",
                          "--debug-verbose-sync-logs-until=%d" % until],
                         stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    while p.poll() is None and not os.path.exists(last) and time.time() - t0 < timeout:
        time.sleep(0.5)
    if p.poll() is None:
        time.sleep(1)
        p.kill()
        p.wait()
        print("%s: Warzone.exe stopped after %.1f s" % (name, time.time() - t0))
    else:
        print("%s: Warzone.exe %.1f s, exit code %d" % (name, time.time() - t0, p.returncode))
    matched = 0
    ok = True
    for t in ticks:
        fn = "sync%d_p%d.txt" % (t, VIEWER)
        ref = read_lines(os.path.join(OFFICIAL, "synclogs", name, fn))
        path = os.path.join(logs, fn)
        if not os.path.exists(path):
            print("  gameTime %d: no log from Warzone.exe" % t)
            ok = False
            break
        vb = read_lines(path)
        if ref == vb:
            matched += 1
            continue
        # the BEGIN / END lines carry the entry count and the CRC of them all: the entries first
        rb, vbb = ref[1:-2], vb[1:-2]
        if rb == vbb:
            k = 0
        else:
            k = next(i for i in range(max(len(rb), len(vbb))) if i >= len(rb) or i >= len(vbb) or rb[i] != vbb[i]) + 1
        print("  gameTime %d: line %d of %d / %d differs (%d entries match)" % (t, k + 1, len(ref), len(vb), max(0, k - 1)))
        for i in range(max(0, k - 3), k + 4):
            print("    ref %s" % (ref[i].decode("utf-8", "replace")[:200] if i < len(ref) else "<end>"))
            print("    vb  %s" % (vb[i].decode("utf-8", "replace")[:200] if i < len(vb) else "<end>"))
        ok = False
        break
    print("%s: %d of %d ticks match" % (name, matched, len(ticks)))
    if ok and not keep:
        shutil.rmtree(work)
    return ok


def main(argv):
    keep = "--keep" in argv
    until = None
    if "--until" in argv:
        until = int(argv[argv.index("--until") + 1])
    timeout = 600
    if "--timeout" in argv:
        timeout = int(argv[argv.index("--timeout") + 1])
    names = [a for a in argv if not a.startswith("--") and not a.isdigit()]
    if not names:
        names = sorted(n for n in os.listdir(os.path.join(OFFICIAL, "synclogs")))
    return 0 if all([run(n, until, keep, timeout) for n in names]) else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
