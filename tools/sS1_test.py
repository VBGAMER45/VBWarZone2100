"""
sS1_test.py - cluster S, fork S1's tests: the in-game multiplayer menu (src/multimenu.cpp: intAddMultiMenu, the
players grid, alliances, gifts, the mute and channel buttons, the Team Strategy tab) and the structure limits
screen (src/multilimit.cpp: WzMultiLimitTitleUI), driven by the UI scripts tools/ui_tests/sS1_*.txt.

  python -I tools\\sS1_test.py [--exe PATH] [teams] [alliance] [limits]      (default: all three)

* Each scenario runs Warzone.exe (default bin\\Warzone.exe; --exe: a fork's private build, e.g.
  build\\j_SS1\\bin\\Warzone.exe) on a fresh config folder build/sS1_test/<scenario>/config that has
  portMapping=0 and lobbyserver=http://127.0.0.1:9 (nothing answers: no router, no public lobby), with
  --vbscript=tools/ui_tests/sS1_<scenario>.txt and --nosound.
* teams: --skirmish=sS1_teams.json (Sk-Rush, alliances 2 = ALLIANCES_TEAMS: player 0 and a NullBot in team 0,
  two NullBots in team 1): the menu opened with the keypad's Enter (kf_addMultiMenu), the grid's rows, the
  power gift to the team mate, the Team Strategy tab, the menu closed.
* alliance: --skirmish=sS1_alliance.json (alliances 1 = ALLIANCES): the alliance button towards player 1
  (requestAlliance), the gift buttons appearing once allied, the mute button.
* limits: Single Player > Start Skirmish Game > the options' Limits button > WzMultiLimitTitleUI: a slider
  moved, Accept Settings, the lobby's limits info (printStructureLimitsInfo), the game started.
Prints each run's exit code, the log's errors and asserts, the script's log lines and the lines the checks look
for; copies the screenshots to build/sS1_test/<scenario>/<shot>.png.
"""
import glob
import json
import os
import shutil
import subprocess
import sys
import time

ROOT = r"C:\Users\Owner\desktop\websites\warzone"
WORK = os.path.join(ROOT, "build", "sS1_test")
TEST_LOBBY = "http://127.0.0.1:9"                   # the discard port: nothing answers


def nb(team):
    return {"team": team, "ai": "nb_generic.js", "difficulty": "Medium"}


SKIRMISH = {
    "teams": {"challenge": {"map": "Sk-Rush", "maxPlayers": 4, "scavengers": 0, "alliances": 2, "bases": 2,
                            "powerLevel": 1, "version": 2},
              "player_0": {"team": 0}, "player_1": nb(0), "player_2": nb(1), "player_3": nb(1)},
    "alliance": {"challenge": {"map": "Sk-Rush", "maxPlayers": 4, "scavengers": 0, "alliances": 1, "bases": 2,
                               "powerLevel": 1, "version": 2},
                 "player_0": {"team": 0}, "player_1": nb(1), "player_2": nb(2), "player_3": nb(3)},
}
KEYS = ("(VB) script", "error", "assert", "gift", "alliance", "LimitSet", "limit", "multimenu", "Alliance")


def run(exe, scenario, limit):
    work = os.path.join(WORK, scenario)
    cfg = os.path.join(work, "config")
    if os.path.exists(work):
        shutil.rmtree(work)
    os.makedirs(cfg)
    with open(os.path.join(cfg, "config"), "w", newline="\r\n") as f:
        f.write("[General]\nportMapping=0\nlobbyserver=%s\n" % TEST_LOBBY)
    args = [exe, "--configdir=" + cfg, "--vbscript=" + os.path.join(ROOT, "tools", "ui_tests", "sS1_%s.txt" % scenario),
            "--nosound", "--datadir=" + os.path.join(ROOT, "data")]
    if scenario in SKIRMISH:
        os.makedirs(os.path.join(cfg, "tests"))
        with open(os.path.join(cfg, "tests", "sS1_%s.json" % scenario), "w") as f:
            json.dump(SKIRMISH[scenario], f, indent=1)
        args.append("--skirmish=sS1_%s.json" % scenario)
    t0 = time.time()
    p = subprocess.Popen(args, cwd=os.path.dirname(exe))
    try:
        p.wait(timeout=limit)
        result = "exit code %d after %.0f s" % (p.returncode, time.time() - t0)
    except subprocess.TimeoutExpired:
        p.kill()
        p.wait()
        result = "stopped after %.0f s (limit)" % (time.time() - t0)
    print("== %s: %s" % (scenario, result))
    lines = []
    for log in sorted(glob.glob(os.path.join(cfg, "logs", "WZlog-*.txt"))):
        lines += open(log, encoding="utf-8", errors="replace").read().splitlines()
    for ln in lines:
        if any(k.lower() in ln.lower() for k in KEYS) and "khr_callback" not in ln:
            print("  | " + ln[:230])
    for shot in sorted(glob.glob(os.path.join(cfg, "screenshots", "*.png"))):
        base = os.path.basename(shot)
        out = os.path.join(work, (base.split("-", 2)[2] if base.count("-") >= 2 else base))
        shutil.copyfile(shot, out)
        print("  screenshot: " + out)
    return p.returncode == 0, lines


def count(scenario, shot, box, test):
    """pixels of box (x0, y0, x1, y1; every 2nd pixel) in the shot for which test(r, g, b) holds; -1: no shot"""
    from PIL import Image
    p = os.path.join(WORK, scenario, shot + ".png")
    if not os.path.exists(p):
        return -1
    px = Image.open(p).convert("RGB").load()
    return sum(1 for x in range(box[0], box[2], 2) for y in range(box[1], box[3], 2) if test(*px[x, y]))


def menu_blue(r, g, b):
    return b > 80 and r < 60 and g < 60


def bright(r, g, b):
    return r > 200 and g > 200 and b > 200


CHECKS = []


def check(ok, what):
    CHECKS.append(ok)
    print("  [%s] %s" % ("ok" if ok else "FAIL", what))


def main(argv):
    exe = os.path.join(ROOT, "bin", "Warzone.exe")
    if "--exe" in argv:
        exe = os.path.abspath(argv[argv.index("--exe") + 1])
    scenarios = [a for a in argv if a in ("teams", "alliance", "limits")] or ["teams", "alliance", "limits"]
    for s in scenarios:
        good, lines = run(exe, s, 300)
        check(good, "%s: exit code 0" % s)
        bad = [ln for ln in lines if ln.startswith(("error", "fatal")) or "assert" in ln.lower()]
        check(not bad, "%s: no error or assert in the log %s" % (s, bad[:2]))
        menu = (230, 40, 800, 190)                  # the menu's grid (the teams game's layout)
        if s == "teams":
            for shot, up in (("menu_open", True), ("power_given", True), ("menu_closed", False),
                             ("menu_reopened", True), ("menu_closed_key", False), ("intel_overlay", True)):
                n = count(s, shot, menu, menu_blue)
                check((n > 300) == up and n >= 0, "teams: the menu %s in %s (%d)" % ("up" if up else "closed", shot, n))
        if s == "alliance":
            button = (378, 95, 414, 119)            # player 1's alliance button
            a, b, c = (count(s, n, button, bright) for n in ("menu_open", "alliance_requested", "alliance_withdrawn"))
            check(a >= 0 and b != a and c == a, "alliance: the button offered, then back (%d, %d, %d)" % (a, b, c))
        if s == "limits":
            knob = (557, 300, 600, 318)             # the Factory slider's left end
            a, b = count(s, "limits_screen", knob, bright), count(s, "limits_set", knob, bright)
            check(b > a >= 0, "limits: the Factory slider moved to 0 (%d -> %d)" % (a, b))
            factory = (630, 350, 870, 520)          # the starting factory (F1) right of the HQ
            n = count(s, "game_base", factory, lambda r, g, b: abs(r - g) < 12 and abs(g - b) < 12 and r > 150)
            check(0 <= n < 20, "limits: the forced limit removed the starting factory (grey pixels %d; a base with it: about 80)" % n)
    failed = CHECKS.count(False)
    print("%d of %d checks passed" % (len(CHECKS) - failed, len(CHECKS)))
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
