"""
lobby_commands_test.py - the lobby's slash commands, kicks, mutes, renames and known players (fork S2: src/multilobbycommands.cpp,
multiint.cpp's kick reason popup and kick redirect, multiplay.cpp's mute and rename, multistat.cpp's known players):
two copies of Warzone.exe on 127.0.0.1, as tools/lan_test.py runs them, with test-safe configs (portMapping=0,
lobbyserver=http://127.0.0.1:9) and game port 2160.

  python -I tools\\lobby_commands_test.py [--exe PATH] [plain] [admin] [known]

* plain: the host (--autohost: Sk-Rush, the host, an open slot, two NullBots; --enablelobbyslashcmd) runs
  tools/ui_tests/s2_host.txt: /help, /admin, /me, /team 1 3, /scav 2, /mute 1, /unmute 1, /kick 1 in its chat box;
  the client (s2_join.txt) joins by the browser's IP box and stays not ready: /help, /kick 0 (refused: no admin),
  /me, /team 0 2 (refused), a rename to VBclient, a chat line while muted; then it is kicked.
* admin: the same with s2_join_admin.txt: the client's profile (multiplay/players/Player.sta2) holds a key made
  here, and the host gets its public key through --addlobbyadminpublickey: the client's /team 0 2 is obeyed
  (/kick 0: "Can't kick the host.").
* known: the client joins and readies (the game starts and runs a minute): both copies record each other in
  knownPlayers.json (addKnownPlayer at NET_PLAYERRESPONDING), checked after.
* mute: two sessions on the same config folders: in the first the host clicks the client's name on its chat line
  and "Mute Player" (storePlayerMuteOption: the player_options row, checked); in the second the client rejoins
  (getStoredPlayerOptions). 4.7.0 stores system_clock's 100 ns count and reads it back as seconds, so the int64
  conversion overflows: whether the mute comes back depends on the wrapped value's sign (printed).
Default: plain and admin. --exe: the Warzone.exe to run (default bin\\Warzone.exe). Results in build/lobby_commands_test/
<mode>/ (screenshots, logs); the checks are printed; the screenshots show the chat lines (the room's messages are
not logged).
"""
import base64
import glob
import json
import os
import re
import shutil
import subprocess
import sys
import time

ROOT = r"C:\Users\Owner\desktop\websites\warzone"
WORK = os.path.join(ROOT, "build", "lobby_commands_test")
UI_TESTS = os.path.join(ROOT, "tools", "ui_tests")
GAME_PORT = 2160                    # not 2100: other tests' games may host there meanwhile
TEST_CONFIG = "[General]\nportMapping=0\nlobbyserver=http://127.0.0.1:9\ngameserver_port=%d\n" % GAME_PORT
AUTOHOST = """{
	"challenge": {"map": "Sk-Rush", "maxPlayers": 4, "scavengers": "false", "version": 2},
	"player_0": {"team": 0, "name": "VBhost"},
	"player_1": {"team": 1},
	"player_2": {"ai": "multiplay/skirmish/nb_generic.js", "difficulty": "Medium", "team": 2},
	"player_3": {"ai": "multiplay/skirmish/nb_generic.js", "difficulty": "Medium", "team": 3}
}
"""
KNOWN_JOIN = """# lobby commands test (fork S2), known players: the client joins s2_host_known's game, readies, plays a minute
wait 600
click 512 400
wait 60
click 512 400
wait 120
click 340 57
wait 20
text 127.0.0.1:2160
wait 10
key 13
wait 300
shot j_lobby
click 775 152
wait 3600
shot j_game
log done
wait 5
quit
"""
JOIN_HEAD = ["wait 600", "click 512 400", "wait 60", "click 512 400", "wait 120", "click 340 57", "wait 20",
             "text 127.0.0.1:2160", "wait 10", "key 13", "wait 300"]


def chat_lines(text, after):
    return ["click 515 689", "wait 10", "text " + text, "wait 5", "key 13", "wait %d" % after]


# mute: session 1, the host clicks the client's name on its chat line and "Mute Player" (setPlayerMuted ->
# storePlayerMuteOption); session 2 (the same config folders), the client rejoins (multiStatsSetIdentity ->
# getStoredPlayerOptions) and chats again
MUTE_HOST1 = ["wait 1800", "shot m1_before", "click 287 601", "wait 40", "shot m1_menu", "click 370 634", "wait 400",
              "shot m1_after", "log done", "wait 5", "quit"]
MUTE_JOIN1 = JOIN_HEAD + chat_lines("hi", 900) + chat_lines("after mute", 300) + ["log done", "wait 5", "quit"]
MUTE_HOST2 = ["wait 1800", "shot m2_rejoined", "log done", "wait 5", "quit"]
MUTE_JOIN2 = JOIN_HEAD + chat_lines("hi again", 700) + ["log done", "wait 5", "quit"]

KNOWN_HOST = """# lobby commands test (fork S2), known players: the host (autohost) waits for the client to ready, plays
wait 1800
shot h_lobby
wait 3000
shot h_game
log done
wait 5
quit
"""


def make_config(path, files=None):
    if os.path.exists(path):
        shutil.rmtree(path)
    os.makedirs(path)
    with open(os.path.join(path, "config"), "w", newline="\r\n") as f:
        f.write(TEST_CONFIG)
    for rel, data in (files or {}).items():
        p = os.path.join(path, rel)
        os.makedirs(os.path.dirname(p), exist_ok=True)
        with open(p, "wb") as f:
            f.write(data)


def run(exe, cfg, script, extra, limit):
    args = [exe, "--configdir=" + cfg, "--vbscript=" + script, "--nosound", "--datadir=" + os.path.join(ROOT, "data")] + extra
    return subprocess.Popen(args, cwd=os.path.dirname(exe))


def finish(p, limit, t0):
    try:
        p.wait(timeout=max(1, limit - (time.time() - t0)))
        return p.returncode
    except subprocess.TimeoutExpired:
        p.kill()
        p.wait()
        return None


def log_text(cfg):
    return "".join(open(f, encoding="utf-8", errors="replace").read() for f in sorted(glob.glob(os.path.join(cfg, "logs", "WZlog-*.txt"))))


def known_players(cfg):
    p = os.path.join(cfg, "knownPlayers.json")
    return json.load(open(p, encoding="utf-8")) if os.path.exists(p) else None


def copy_shots(cfg, dest):
    for shot in sorted(glob.glob(os.path.join(cfg, "screenshots", "*.png"))):
        base = os.path.basename(shot)
        shutil.copyfile(shot, os.path.join(dest, base.split("-", 2)[2] if base.count("-") >= 2 else base))


def admin_profile():
    """A profile whose key the host knows: (Player.sta2 bytes, base64 of the public key)."""
    import nacl.signing
    sk = nacl.signing.SigningKey(bytes(range(1, 33)))
    sk64 = bytes(sk) + bytes(sk.verify_key)
    sta = "WZ.STA.v3\n0 0 0 0 0\n%s\n" % base64.b64encode(sk64).decode()
    return sta.encode(), base64.b64encode(bytes(sk.verify_key)).decode()


def mode_run(exe, mode, checks):
    work = os.path.join(WORK, mode)
    os.makedirs(work, exist_ok=True)
    hcfg, ccfg = os.path.join(work, "host_config"), os.path.join(work, "join_config")
    make_config(hcfg, {os.path.join("autohost", "lan_ai.json"): AUTOHOST.encode()})
    host_extra = ["--autohost=lan_ai.json", "--enablelobbyslashcmd"]
    join_files = {}
    if mode == "admin":
        sta, pk = admin_profile()
        join_files[os.path.join("multiplay", "players", "Player.sta2")] = sta
        host_extra.append("--addlobbyadminpublickey=" + pk)
    make_config(ccfg, join_files)
    if mode == "known":
        hs, js = os.path.join(work, "s2_host_known.txt"), os.path.join(work, "s2_join_known.txt")
        open(hs, "w").write(KNOWN_HOST)
        open(js, "w").write(KNOWN_JOIN)
    else:
        hs = os.path.join(UI_TESTS, "s2_host.txt")
        js = os.path.join(UI_TESTS, "s2_join_admin.txt" if mode == "admin" else "s2_join.txt")
    t0 = time.time()
    hp = run(exe, hcfg, hs, host_extra, 300)
    time.sleep(2)
    cp = run(exe, ccfg, js, [], 300)
    cr = finish(cp, 300, t0)
    hr = finish(hp, 300, t0)
    copy_shots(hcfg, work)
    copy_shots(ccfg, work)
    hlog, clog = log_text(hcfg), log_text(ccfg)
    for name, text in (("host", hlog), ("client", clog)):
        for ln in text.splitlines():
            if any(k in ln for k in ("processChatLobbySlashCommands", "kickPlayer", "Kicking", "kick", "admin", "known_player",
                                     "SQLite", "player_options", "ASSERT", "assert")) or re.match(r"(error|fatal)\s*\|", ln):
                if "khr_callback" not in ln:
                    print("  %s | %s" % (name, ln.rstrip()[:200]))

    def check(ok, what):
        checks.append((bool(ok), "%s: %s" % (mode, what)))
        print("  [%s] %s: %s" % ("ok" if ok else "FAIL", mode, what))

    check(hr == 0 and cr == 0, "both exit with 0 (host %s, client %s; %.0f s)" % (hr, cr, time.time() - t0))
    bad = [ln for ln in (hlog + clog).splitlines() if re.match(r"(error|fatal)\s*\|", ln) or "ASSERT" in ln]
    check(not bad, "no error or assert in the logs" + ("" if not bad else ": " + bad[0][:150]))
    hk, ck = known_players(hcfg), known_players(ccfg)
    check(hk is not None and hk.get("user_version") == 1 and ck is not None and ck.get("user_version") == 1,
          "knownPlayers.json made by both (user_version 1)")
    if mode in ("plain", "admin"):
        for cmd in ("/help", "/admin", "/me", "/team 1 3", "/scav 2", "/mute 1", "/unmute 1", "/kick 1"):
            check("] [%s]" % cmd in hlog, "the host ran its %s" % cmd)
        for cmd in ("/help", "/kick 0", "/me", "/team 0 2"):
            check("] [%s]" % cmd in hlog, "the host ran the client's %s" % cmd)
        check("Kicking player 1" in hlog, "the client kicked (kickPlayer)")
        check("You have been kicked" in clog or "kicked" in clog.lower(), "the client saw its kick")
    if mode == "known":
        names_h = [r["name"] for r in (hk or {}).get("known_players", [])]
        names_c = [r["name"] for r in (ck or {}).get("known_players", [])]
        check("Player_2" in names_h, "the host knows the client (Player_2: both start as Player): %s" % names_h)
        check("Player" in names_c, "the client knows the host (Player): %s" % names_c)
        check(all(len(base64.b64decode(r["pk"])) == 32 for r in (hk or {}).get("known_players", []) + (ck or {}).get("known_players", [])),
              "the keys are 32-byte public keys")


def wrapped_ticks(secs):
    """4.7.0's system_clock::time_point(std::chrono::seconds(secs)).time_since_epoch().count() on MSVC: secs * 10^7,
    overflowing int64 (wraps)."""
    v = (secs * 10 ** 7) % 2 ** 64
    return v - 2 ** 64 if v >= 2 ** 63 else v


def mute_run(exe, checks):
    work = os.path.join(WORK, "mute")
    os.makedirs(work, exist_ok=True)
    hcfg, ccfg = os.path.join(work, "host_config"), os.path.join(work, "join_config")
    make_config(hcfg, {os.path.join("autohost", "lan_ai.json"): AUTOHOST.encode()})
    make_config(ccfg)

    def check(ok, what):
        checks.append((bool(ok), "mute: " + what))
        print("  [%s] mute: %s" % ("ok" if ok else "FAIL", what))

    for n, (hl, jl) in enumerate(((MUTE_HOST1, MUTE_JOIN1), (MUTE_HOST2, MUTE_JOIN2)), 1):
        hs, js = os.path.join(work, "host%d.txt" % n), os.path.join(work, "join%d.txt" % n)
        open(hs, "w").write("\n".join(hl) + "\n")
        open(js, "w").write("\n".join(jl) + "\n")
        t0 = time.time()
        hp = run(exe, hcfg, hs, ["--autohost=lan_ai.json"], 200)
        time.sleep(2)
        cp = run(exe, ccfg, js, [], 200)
        cr, hr = finish(cp, 200, t0), finish(hp, 200, t0)
        check(hr == 0 and cr == 0, "session %d: both exit with 0 (host %s, client %s)" % (n, hr, cr))
        if n == 1:
            kp = known_players(hcfg) or {}
            rows = kp.get("player_options", [])
            print("    host's player_options: %s" % rows)
            ok = len(rows) == 1 and rows[0]["name"] == "Player_2" and int(rows[0]["muted"]) > 0 and int(rows[0]["banned"]) == -1
            check(ok, "session 1: the mute stored (Player_2, the client's key, muted > 0, banned -1)")
            ck = known_players(ccfg) or {}
            client_pk = None
            # the client's public key, from its profile (multiplay/players/Player.sta2: the 64-byte secret key's
            # second half)
            sta = os.path.join(ccfg, "multiplay", "players", "Player.sta2")
            if os.path.exists(sta):
                client_pk = base64.b64encode(base64.b64decode(open(sta).read().split()[-1])[32:]).decode()
            check(ok and rows[0]["pk"] == client_pk, "session 1: the stored key is the client's")
            expect_muted = ok and wrapped_ticks(int(rows[0]["muted"])) > 0
        else:
            hlog = log_text(hcfg)
            print("    4.7.0's read-back of the stored mute: %s (the wrapped count's sign)" % ("muted" if expect_muted else "not muted"))
            check("ASSERT" not in hlog, "session 2: no assert")
    copy_shots(hcfg, work)
    copy_shots(ccfg, work)


def main(argv):
    exe = os.path.join(ROOT, "bin", "Warzone.exe")
    if "--exe" in argv:
        exe = os.path.abspath(argv[argv.index("--exe") + 1])
    modes = [a for a in argv if a in ("plain", "admin", "known", "mute")] or ["plain", "admin"]
    checks = []
    for mode in modes:
        print("== %s (%s)" % (mode, exe))
        if mode == "mute":
            mute_run(exe, checks)
        else:
            mode_run(exe, mode, checks)
    failed = [w for ok, w in checks if not ok]
    print("%d of %d checks passed%s" % (len(checks) - len(failed), len(checks), (": failed " + "; ".join(failed)) if failed else ""))
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
