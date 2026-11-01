"""
guide_test.py - the Game Guide (cluster S, fork S3: src/screens/guidescreen.cpp) through UI scripts
(tools/ui_tests/sS3_*.txt), on bin\\Warzone.exe or a fork's private build:

  python -I tools\\guide_test.py [--exe PATH] [skirmish] [campaign] [save]      (default: all three)

* skirmish (sS3_skirmish.txt): the skirmish rules add every topic (addGuideTopic("wz2100::**")); the in-game menu's
  View Guide opens the guide with its sidebar; a topic picked in the sidebar, the back and forward buttons, the
  sidebar toggled, the options pop-over (Disable Topic Pop-ups), a linked topic ("See Also"), Escape closes it.
* campaign (sS3_campaign.txt): CAM_1A from the main menu; its script adds "wz2100::structures::hq" with
  SHOWTOPIC_FIRSTADD 3 s in: the guide pops up by itself (no sidebar) and pauses the game; the close button
  closes it and the game runs again.
* save (sS3_save.txt, then sS3_load.txt on the same config): CAM_1A's topic pop-up closed (the UI script's
  closeguide), a save from the in-game menu, the main menu's Load Campaign Game of it, View Guide (the topics
  restored), a second save: both saves' guidetopics.json hold the same loaded_topics.
Every run's config is test-safe (portMapping=0, lobbyserver=http://127.0.0.1:9) and logs --debug=wz (the guide's
registry lines). Results in build/guide_test/<scenario>/ (screenshots, the log's guide and error lines).
"""
import glob
import json
import os
import shutil
import subprocess
import sys
import time

ROOT = r"C:\Users\Owner\desktop\websites\warzone"
WORK = os.path.join(ROOT, "build", "guide_test")
SCRIPTS = os.path.join(ROOT, "tools", "ui_tests")
CHECKS = []


def check(ok, what):
    CHECKS.append((bool(ok), what))
    print("  [%s] %s" % ("ok" if ok else "FAIL", what))


def make_config(name):
    cfg = os.path.join(WORK, name, "config")
    if os.path.exists(cfg):
        shutil.rmtree(cfg)
    os.makedirs(cfg)
    with open(os.path.join(cfg, "config"), "w", newline="\r\n") as f:
        f.write("[General]\nportMapping=0\nlobbyserver=http://127.0.0.1:9\n")
    return cfg


def run_script(exe, name, cfg, script, limit):
    """Runs a UI script; returns (exit text, log lines, screenshots copied)."""
    t0 = time.time()
    p = subprocess.Popen([exe, "--configdir=" + cfg, "--vbscript=" + os.path.join(SCRIPTS, script), "--nosound",
                          "--debug=wz", "--datadir=" + os.path.join(ROOT, "data")], cwd=os.path.dirname(exe))
    try:
        p.wait(timeout=limit)
        how = "exit code %d after %.0f s" % (p.returncode, time.time() - t0)
    except subprocess.TimeoutExpired:
        p.kill()
        p.wait()
        how = "killed at the time limit (%d s)" % limit
    lines = []
    for f in sorted(glob.glob(os.path.join(cfg, "logs", "WZlog-*.txt")), key=os.path.getmtime)[-1:]:
        lines = open(f, encoding="utf-8", errors="replace").read().splitlines()
    shots = []
    for shot in sorted(glob.glob(os.path.join(cfg, "screenshots", "*.png"))):
        if os.path.getmtime(shot) < t0:
            continue
        base = os.path.basename(shot)
        out = os.path.join(WORK, name, (base.split("-", 2)[2] if base.count("-") >= 2 else base))
        shutil.copyfile(shot, out)
        shots.append(out)
    return how, lines, shots


def report(name, how, lines, shots):
    print("%s: %s" % (name, how))
    for s in shots:
        print("  screenshot: " + s)
    keys = ("guide", "topic", "(VB) script", "error", "assert", "category", "parent:")
    for l in lines:
        if any(k in l.lower() for k in keys) and "khr_callback" not in l:
            print("  | " + l[:200])


def problems(lines):
    return [l for l in lines if l.startswith(("error", "fatal")) or "assert" in l.lower()]


def shot_names(shots):
    return {os.path.splitext(os.path.basename(s))[0] for s in shots}


def guide_topics(cfg, save):
    p = os.path.join(cfg, "savegames", "campaign", save, "guidetopics.json")
    if not os.path.exists(p):
        return None
    return json.load(open(p, encoding="utf-8"))


def main(argv):
    exe = os.path.join(ROOT, "bin", "Warzone.exe")
    if "--exe" in argv:
        exe = os.path.abspath(argv[argv.index("--exe") + 1])
    modes = [a for a in argv if a in ("skirmish", "campaign", "save")] or ["skirmish", "campaign", "save"]
    os.makedirs(WORK, exist_ok=True)
    print("exe: " + exe)

    if "skirmish" in modes:
        cfg = make_config("skirmish")
        how, lines, shots = run_script(exe, "skirmish", cfg, "sS3_skirmish.txt", 300)
        report("skirmish", how, lines, shots)
        check(how.startswith("exit code 0"), "skirmish: the game ran the script and quit")
        check(not problems(lines), "skirmish: no error or assert in the log: %s" % problems(lines)[:3])
        check(any("Loaded category: wz2100::general" in l for l in lines), "skirmish: the rules' wildcard loaded the categories")
        check({"guide_open", "guide_topic", "guide_back", "guide_options", "guide_linked", "guide_closed"} <= shot_names(shots),
              "skirmish: the guide's screenshots were taken (look at them)")

    if "campaign" in modes:
        cfg = make_config("campaign")
        how, lines, shots = run_script(exe, "campaign", cfg, "sS3_campaign.txt", 400)
        report("campaign", how, lines, shots)
        check(how.startswith("exit code 0"), "campaign: the game ran the script and quit")
        check(not problems(lines), "campaign: no error or assert in the log: %s" % problems(lines)[:3])
        check(any("Loaded category: wz2100::structures" in l for l in lines), "campaign: the HQ topic's category loaded")
        check({"popup", "popup_closed"} <= shot_names(shots), "campaign: the pop-up's screenshots were taken (look at them)")

    if "save" in modes:
        cfg = make_config("save")
        how, lines, shots = run_script(exe, "save", cfg, "sS3_save.txt", 400)
        report("save", how, lines, shots)
        check(how.startswith("exit code 0") and not problems(lines), "save: saved without errors: %s" % problems(lines)[:3])
        how, lines, shots = run_script(exe, "save", cfg, "sS3_load.txt", 400)
        report("load", how, lines, shots)
        check(how.startswith("exit code 0") and not problems(lines), "load: loaded and saved again without errors: %s" % problems(lines)[:3])
        # the two saves the scripts named (the game's autosave folder, "auto", aside)
        saves = [d for n in ("guidesave", "again") for d in sorted(os.listdir(os.path.join(cfg, "savegames", "campaign")))
                 if d.endswith(n) and os.path.isdir(os.path.join(cfg, "savegames", "campaign", d))]
        print("  saves: %s" % saves)
        topics = [guide_topics(cfg, s) for s in saves]
        for s, t in zip(saves, topics):
            print("  %s guidetopics.json: %s" % (s, json.dumps(t)))
        check(len(topics) == 2 and topics[0] is not None and topics[0].get("loaded_topics"),
              "save: the first save has guide topics")
        check(len(topics) == 2 and topics[0] == topics[1], "save: the second save (after the load) has the same topics and options")

    failed = [w for ok, w in CHECKS if not ok]
    print("%d of %d checks passed" % (len(CHECKS) - len(failed), len(CHECKS)))
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
