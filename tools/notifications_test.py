"""
notifications_test.py - the in-game notifications (src/notifications.cpp: notifications_c) and the help overlay
(src/screens/helpscreen.cpp: helpscreen_c), cluster S (fork S4), in Warzone.exe with tools/ui_tests' sS4_* scripts:

  python -I tools\\notifications_test.py [--exe PATH] [look] [help] [ignore] [vote]

* help: a skirmish; "help" in the chat box (cheat.cpp: intShowWidgetHelp) opens the help overlay over the game; the
  mouse rests on the reticule's research button (its help panel), then Escape closes the overlay.
* ignore: the main menu with the campaign videos missing (a data folder of hard links to data\\ without
  sequences.wz) raises frontend.cpp's "Campaign videos are missing" notification (makeIgnorable(..., 2)): four runs
  on one config folder: dismissed twice (notifications.json's "seen" 1, then 2; no "Do not show again" box), shown
  the third time with the box, which is ticked before the dismissal ("skip": true), and not shown the fourth time
  ("Ignoring notification"). The notification's "Open wz2100.net" action is never clicked (it opens a browser).
* vote: two copies on 127.0.0.1 (game port 2193): the client joins the host's game through the game
  browser's IP box; the host clicks its map button, which with two humans asks the client for a vote
  (multivote.cpp's setupLobbyChangeVoteChoice): the client's "Vote" notification, its "Allow" action clicked (the
  host logs the vote).
* look: one run of the main menu without videos (screenshots only).
Default: help and ignore. --exe: the Warzone.exe to run (default bin\\Warzone.exe; a fork's private build, e.g.
build\\j_S4\\bin\\Warzone.exe). Every config has portMapping=0 and lobbyserver=http://127.0.0.1:9 (no router mapping,
no public lobby). Results in build\\notifications_test\\<run>\\ (screenshots, the log's lines).
"""
import glob
import json
import os
import shutil
import subprocess
import sys
import time

ROOT = r"C:\Users\Owner\desktop\websites\warzone"
WORK = os.path.join(ROOT, "build", "notifications_test")
SCRIPTS = os.path.join(ROOT, "tools", "ui_tests")
NOVIDEO = os.path.join(WORK, "novideo_data")
VOTE_PORT = 2193                        # (tools/lan_test.py's 2100 and the other forks' tests stay free)
CHECKS = []


def check(ok, what):
    CHECKS.append((bool(ok), what))
    print("  [%s] %s" % ("ok" if ok else "FAIL", what))


def novideo_data():
    """data\\ as hard links, without sequences.wz (the videos missing)."""
    src = os.path.join(ROOT, "data")
    for dp, _, fs in os.walk(src):
        rel = os.path.relpath(dp, src)
        if rel.split(os.sep)[0] == "sequences.wz":
            continue
        os.makedirs(os.path.join(NOVIDEO, rel), exist_ok=True)
        for f in fs:
            if rel == "." and f == "sequences.wz":
                continue
            d = os.path.join(NOVIDEO, rel, f)
            if not os.path.exists(d):
                os.link(os.path.join(dp, f), d)
    return NOVIDEO


def make_config(name, fresh=True):
    cfg = os.path.join(WORK, name, "config")
    if fresh and os.path.exists(cfg):
        shutil.rmtree(cfg)
    os.makedirs(cfg, exist_ok=True)
    if not os.path.exists(os.path.join(cfg, "config")):
        with open(os.path.join(cfg, "config"), "w", newline="\r\n") as f:
            f.write("[General]\nportMapping=0\nlobbyserver=http://127.0.0.1:9\n")
    return cfg


def novideo_exe(exe):
    """A copy of the exe and its DLLs in build\notifications_test\bin: the game also searches the data folder
    beside its bin folder (<prefix>\data, as 4.7.0 does), which for bin\Warzone.exe is data\ with the videos."""
    out = os.path.join(WORK, "bin")
    os.makedirs(out, exist_ok=True)
    folder = os.path.dirname(exe)
    for n in os.listdir(folder):
        if n.lower() == "warzone.exe" or n.lower().endswith(".dll"):
            shutil.copyfile(os.path.join(folder, n), os.path.join(out, n))
    return os.path.join(out, os.path.basename(exe))


def start(exe, cfg, script, data, extra=()):
    args = [exe, "--configdir=" + cfg, "--vbscript=" + os.path.join(SCRIPTS, script), "--nosound",
            "--datadir=" + data, "--debug=gui"] + list(extra)
    return subprocess.Popen(args, cwd=os.path.dirname(exe))


def finish(p, limit, t0):
    try:
        p.wait(timeout=max(1, limit - (time.time() - t0)))
    except subprocess.TimeoutExpired:
        p.kill()
        p.wait()
        return "killed at the time limit"
    return "exit code %d" % p.returncode


def log_text(cfg, since):
    out = []
    for f in sorted(glob.glob(os.path.join(cfg, "logs", "WZlog-*.txt"))):
        if os.path.getmtime(f) >= since:
            out += open(f, encoding="utf-8", errors="replace").read().splitlines()
    return out


def copy_shots(cfg, dest, since):
    got = []
    os.makedirs(dest, exist_ok=True)
    for shot in sorted(glob.glob(os.path.join(cfg, "screenshots", "*.png"))):
        if os.path.getmtime(shot) < since:
            continue
        base = os.path.basename(shot)
        out = os.path.join(dest, (base.split("-", 2)[2] if base.count("-") >= 2 else base))
        shutil.copyfile(shot, out)
        got.append(out)
    return got


def run(exe, name, script, data, limit, cfg=None, extra=()):
    """One run: its exit, its screenshots, its log lines."""
    cfg = cfg or make_config(name)
    for old in glob.glob(os.path.join(cfg, "screenshots", "*.png")):
        os.remove(old)                  # (a config folder used again: the earlier runs' screenshots)
    t0 = time.time()
    p = start(exe, cfg, script, data, extra)
    res = finish(p, limit, t0)
    shots = copy_shots(cfg, os.path.join(WORK, name), t0 - 1)
    lines = log_text(cfg, t0 - 1)
    print("%s (%s): %s, %.0f s, %d log lines" % (name, script, res, time.time() - t0, len(lines)))
    for s in shots:
        print("  screenshot: " + s)
    bad = [l for l in lines if l.startswith(("error", "fatal")) or "assert" in l.lower()]
    for l in bad[:10]:
        print("  | " + l[:220])
    return res, lines, bad


def prefs(cfg):
    p = os.path.join(cfg, "notifications.json")
    if not os.path.exists(p):
        return None
    return json.load(open(p, encoding="utf-8"))


def mode_help(exe):
    print("== help")
    res, lines, bad = run(exe, "help", "sS4_help.txt", os.path.join(ROOT, "data"), 240)
    check(res == "exit code 0", "the game quit by itself (%s)" % res)
    check(not bad, "no error or assert in the log (%d)" % len(bad))
    check(any("(VB) script: help overlay open" in l for l in lines), "the script reached the overlay")
    for s in ("help_overlay", "help_panel", "help_closed"):
        check(os.path.exists(os.path.join(WORK, "help", s + ".png")), "screenshot %s (look at it)" % s)


def mode_ignore(exe):
    print("== ignore")
    data = novideo_data()
    exe = novideo_exe(exe)
    cfg = make_config("ignore")
    ident = "campaignVideoNotification"
    expected = [(1, None, "sS4_notify_dismiss.txt"), (2, None, "sS4_notify_dismiss.txt"),
                (3, True, "sS4_notify_dontshow.txt"), (3, True, "sS4_notify_none.txt")]
    for i, (seen, skip, script) in enumerate(expected, 1):
        res, lines, bad = run(exe, "ignore%d" % i, script, data, 120, cfg=cfg)
        check(res == "exit code 0", "run %d: the game quit by itself (%s)" % (i, res))
        check(not bad, "run %d: no error or assert in the log (%d)" % (i, len(bad)))
        pr = prefs(cfg)
        entry = ((pr or {}).get("notifications") or {}).get(ident, {})
        check(entry.get("seen") == seen and entry.get("skip") == skip,
              "run %d: notifications.json %s = %s (want seen %s, skip %s)" % (i, ident, entry, seen, skip))
        if i == 3:
            check(any("Do Not Show Notification Again: " + ident in l for l in lines), "run 3: the box ticked was stored")
        if i == 4:
            check(any("Ignoring notification: " + ident in l for l in lines), "run 4: the notification was ignored")


def main(argv):
    exe = os.path.join(ROOT, "bin", "Warzone.exe")
    if "--exe" in argv:
        exe = os.path.abspath(argv[argv.index("--exe") + 1])
    modes = [a for a in argv if a in ("look", "help", "ignore", "vote")] or ["help", "ignore"]
    os.makedirs(WORK, exist_ok=True)
    print("exe %s" % exe)
    if "look" in modes:
        run(novideo_exe(exe), "look", "sS4_notify_look.txt", novideo_data(), 90)
    if "help" in modes:
        mode_help(exe)
    if "ignore" in modes:
        mode_ignore(exe)
    if "vote" in modes:
        mode_vote(exe)
    failed = [w for ok, w in CHECKS if not ok]
    print("%d of %d checks passed" % (len(CHECKS) - len(failed), len(CHECKS)))
    return 1 if failed else 0


def mode_vote(exe):
    print("== vote")
    data = os.path.join(ROOT, "data")
    hcfg = make_config("vote_host")
    with open(os.path.join(hcfg, "config"), "a", newline="\r\n") as f:
        f.write("gameserver_port=%d\n" % VOTE_PORT)
    os.makedirs(os.path.join(hcfg, "autohost"), exist_ok=True)
    lan = open(os.path.join(ROOT, "tools", "lan_test.py"), encoding="utf-8").read()
    with open(os.path.join(hcfg, "autohost", "lan_ai.json"), "w") as f:
        f.write(lan.split('AUTOHOST = """')[1].split('"""')[0])
    jcfg = make_config("vote_join")
    t0 = time.time()
    hp = start(exe, hcfg, "sS4_vote_host.txt", data, ["--autohost=lan_ai.json", "--debug=net"])
    time.sleep(2)
    jp = start(exe, jcfg, "sS4_vote_join.txt", data, ["--debug=net"])
    jres, hres = finish(jp, 180, t0), finish(hp, 180, t0)
    for name, cfg, res in (("vote_join", jcfg, jres), ("vote_host", hcfg, hres)):
        print("%s: %s" % (name, res))
        for s in copy_shots(cfg, os.path.join(WORK, name), t0 - 1):
            print("  screenshot: " + s)
    hlog, jlog = log_text(hcfg, t0 - 1), log_text(jcfg, t0 - 1)
    check(hres == "exit code 0" and jres == "exit code 0", "both quit by themselves (%s, %s)" % (hres, jres))
    for name, lines in (("host", hlog), ("client", jlog)):
        bad = [l for l in lines if l.startswith(("error", "fatal")) or "assert" in l.lower()]
        for l in bad[:6]:
            print("  | %s: %s" % (name, l[:200]))
        check(not bad, "%s: no error or assert in the log (%d)" % (name, len(bad)))
    check(any("has joined, IP is: 127.0.0.1" in l for l in hlog), "the client joined the host")
    # recvLobbyChangeVote logs the total when a client's vote arrives: the host's own and the client's
    check(any("[recvLobbyChangeVote:802] total votes: 2/2" in l for l in hlog),
          "the host got the client's vote (the notification's Allow action): total votes 2/2")
    for s in ("vote_join/vj_vote", "vote_join/vj_after", "vote_host/vh_asked", "vote_host/vh_after"):
        check(os.path.exists(os.path.join(WORK, s + ".png")), "screenshot %s (look at it)" % s)


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
