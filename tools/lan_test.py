"""
lan_test.py - cluster J's milestone test: two copies of Warzone.exe play a multiplayer skirmish together over
the network (127.0.0.1), started from the menus as players do.

  python -I tools\\lan_test.py [--no-ai]

* The host: bin\\Warzone.exe --autohost=lan_ai.json (tools/ui_tests/lan_host_ai.txt waits and takes
  screenshots): Sk-Rush with the host, an open slot and two NullBot AIs (the AIs run on the host; their
  orders reach the client as network messages). --no-ai: lan_host.txt instead, the host going through
  Multi Player > Host Game > Start Hosting and Ready by mouse, an open slot and no AIs.
* The client, 2 s later: tools/ui_tests/lan_join.txt: Multi Player > Join Game, the game browser's
  "GameID or IP:port" box, 127.0.0.1:2100, Enter (the joining screen's handshake: the identity challenge,
  the session keys), the host's lobby, Ready; then about two minutes of play, and it quits.
* Passes when: both exit with 0; the host logs the client's join, startMultiplayerGame and NET_FIREUP sent;
  the client logs NET_FIREUP received; both load the map and start a replay; the host logs the client
  leaving in game (its game time printed); neither writes a desync dump (gtime's checkDebugSync compares every
  player's game state CRC each tick); no assert or error in either log.
* The runs' folders: build/ui_test/lan_host_ai (or lan_host) and build/ui_test/lan_join.
"""
import glob
import os
import re
import shutil
import subprocess
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import build  # noqa: E402

ROOT = build.ROOT
UI = os.path.join(ROOT, "tools", "ui_test.py")
WORK = os.path.join(ROOT, "build", "ui_test")
AUTOHOST = """{
	"challenge": {"map": "Sk-Rush", "maxPlayers": 4, "scavengers": "false", "version": 2},
	"player_0": {"team": 0, "name": "VBhost"},
	"player_1": {"team": 1},
	"player_2": {"ai": "multiplay/skirmish/nb_generic.js", "difficulty": "Medium", "team": 2},
	"player_3": {"ai": "multiplay/skirmish/nb_generic.js", "difficulty": "Medium", "team": 3}
}
"""


def log_text(name):
    logs = glob.glob(os.path.join(WORK, name, "config", "logs", "WZlog-*.txt"))
    return open(logs[0], encoding="utf-8", errors="replace").read() if logs else ""


def main(argv):
    ai = "--no-ai" not in argv
    host = "lan_host_ai" if ai else "lan_host"
    host_args = [sys.executable, "-I", UI, host + ".txt", "900"]
    if ai:
        cfg = os.path.join(WORK, host, "config")
        if os.path.exists(cfg):
            shutil.rmtree(cfg)
        os.makedirs(os.path.join(cfg, "autohost"))
        with open(os.path.join(cfg, "autohost", "lan_ai.json"), "w") as f:
            f.write(AUTOHOST)
        host_args += ["--keep-config", "--autohost=lan_ai.json"]
    t0 = time.time()
    hp = subprocess.Popen(host_args, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
    time.sleep(2)
    cp = subprocess.run([sys.executable, "-I", UI, "lan_join.txt", "600"], capture_output=True, text=True)
    hout, _ = hp.communicate()
    hlog, clog = log_text(host), log_text("lan_join")
    checks = [
        ("host exit code 0", "exit code 0 " in hout),
        ("client exit code 0", "exit code 0 " in cp.stdout),
        ("the host opened its listening socket", "Opening listening socket" in hlog),
        ("the client joined the host", "has joined, IP is: 127.0.0.1" in hlog),
        ("the host started the game", "startMultiplayerGame" in hlog and "Sending NET_FIREUP" in hlog),
        ("the client received NET_FIREUP", "NET_FIREUP was received" in clog),
        ("both loaded the map", "gameLoad" in hlog and "gameLoad" in clog),
        ("both started their replay", "Started writing replay" in hlog and "Started writing replay" in clog),
        ("the host saw the client leave in game", "has dropped, in-game!" in hlog),
    ]
    desyncs = [p for n in (host, "lan_join") for p in glob.glob(os.path.join(WORK, n, "config", "**", "*desync*"),
                                                             recursive=True)]
    checks.append(("no desync dump", not desyncs))
    bad = [ln for ln in (hlog + clog).splitlines() if re.match(r"(error|fatal)\s*\|", ln) or "ASSERT" in ln]
    checks.append(("no error or assert in the logs", not bad))
    m = re.search(r"has dropped, in-game! \(gameTime: (\d+)\)", hlog)
    for name, ok in checks:
        print("  %s %s" % ("ok  " if ok else "FAIL", name))
    for ln in bad[:10]:
        print("    " + ln[:200])
    passed = sum(ok for _, ok in checks)
    print("lan: %d of %d checks pass%s; %s game time %s s together; %.0f s" % (
        passed, len(checks), " (with two NullBots on the host)" if ai else "", host,
        "%.1f" % (int(m.group(1)) / 1000) if m else "?", time.time() - t0))
    return 0 if passed == len(checks) else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
