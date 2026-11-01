"""
s5_test.py - cluster S's fork S5: the spectators' stats overlay (spectatorwidgets.cpp), the replay controller
(hci.cpp's ReplayControllerWidget), the blind lobby's waiting room (blindwaitingroom.cpp) and the multiplayer paths
not run before (a spectator joining and moving to a player slot, a map download), in windowed runs.

  python -I tools\\s5_test.py [--exe PATH] [replay] [spectator] [blind] [download] [mod] [ready]

* replay: ref/official/replays/rush_10p_autogame.wzrp played with --loadreplay (the viewer is player 23, a
  spectator): the replay controller over the game (Pause, Resume, Fast-Forward clicked: the game time stops and goes
  on, fast-forward's outline), the "Show Player Stats" button and the stats view (its rows, closed again).
* spectator: two copies on 127.0.0.1 (test-safe configs: portMapping=0, lobbyserver=http://127.0.0.1:9): the host
  (--autohost: Sk-Rush, an open slot, two NullBots, an open spectator slot), a client joining with
  --spectate=127.0.0.1:2100 (the spectators' list), then moving to the open player slot and back; then a client that
  stays a spectator while the game plays: the stats button and the stats view in a networked game.
* blind: the host autohosts in blind mode ("blindMode": "blind_lobby_simple_lobby"), the client joins: its waiting
  room (banner, host, its own row, status, players ready) in place of the players list.
* download: the host autohosts a map the client's data lacks (a map package put in the host's config folder only:
  config/maps/<map>.wz, read by buildMapList through WzMapZipIO and wzmaplib's MapPackage): the client joins and
  downloads it (byte for byte, its check passed), clicks Ready, and both play the level loaded from the package.
* mod: the same with a multiplayer mod (the host's config/mods/4.7.0/multiplay, --mod_mp): the host sends it and the
  client keeps it in mods/downloads (its check passed; fork S9 fixed validateReceivedFile's end-of-file test).
* ready: the host's Ready box with an open player slot: none with 4.7.0's autohost default (locked.readybeforefull),
  and with "locked": {"readybeforefull": false} the box, whose click readies the host and starts the game.
  Default: replay, spectator, blind.
Results in build/s5_test/<scenario>/ (screenshots, logs); the checks print [ok] / [FAIL].
"""
import glob
import os
import shutil
import subprocess
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORK = os.path.join(ROOT, "build", "s5_test")
TEST_CONFIG = "[General]\nportMapping=0\nlobbyserver=http://127.0.0.1:9\n"
CHECKS = []


def check(ok, what):
    CHECKS.append((bool(ok), what))
    print("  [%s] %s" % ("ok" if ok else "FAIL", what))


def make_config(name, extra_config="", autohost=None, files=()):
    cfg = os.path.join(WORK, name, "config")
    if os.path.exists(os.path.dirname(cfg)):
        shutil.rmtree(os.path.dirname(cfg))
    os.makedirs(cfg)
    with open(os.path.join(cfg, "config"), "w", newline="\r\n") as f:
        f.write(TEST_CONFIG + extra_config)     # never a router mapping, never the public lobby
    if autohost:
        os.makedirs(os.path.join(cfg, "autohost"))
        with open(os.path.join(cfg, "autohost", "s5.json"), "w") as f:
            f.write(autohost)
    for src, rel in files:
        os.makedirs(os.path.dirname(os.path.join(cfg, rel)), exist_ok=True)
        shutil.copyfile(src, os.path.join(cfg, rel))
    return cfg


def run_game(exe, name, cfg, lines, extra):
    scr = os.path.join(WORK, name, name + ".txt")
    with open(scr, "w") as f:
        f.write("\n".join(lines) + "\n")
    args = [exe, "--configdir=" + cfg, "--vbscript=" + scr, "--nosound", "--datadir=" + os.path.join(ROOT, "data")] + extra
    return subprocess.Popen(args, cwd=os.path.dirname(exe))


def finish(p, limit, t0):
    try:
        p.wait(timeout=max(1, limit - (time.time() - t0)))
        return p.returncode
    except subprocess.TimeoutExpired:
        p.kill()
        p.wait()
        return None


def log_of(cfg):
    out = ""
    for f in sorted(glob.glob(os.path.join(cfg, "logs", "WZlog-*.txt"))):
        out += open(f, encoding="utf-8", errors="replace").read()
    return out


def report(name, cfg, keys=()):
    log = log_of(cfg)
    shots = []
    for shot in sorted(glob.glob(os.path.join(cfg, "screenshots", "*.png"))):
        base = os.path.basename(shot)
        out = os.path.join(WORK, name, base.split("-", 2)[2] if base.count("-") >= 2 else base)
        shutil.copyfile(shot, out)
        shots.append(out)
    for ln in log.splitlines():
        if ln.startswith(("error", "fatal")) or "assert" in ln.lower() or "(VB) script" in ln or any(k in ln for k in keys):
            if "khr_callback" not in ln:
                print("  | " + ln[:220])
    for s in shots:
        print("  screenshot: " + s)
    return log


# 4.7.0's own log lines in these runs (the same code paths in the C++): a spectator's NET_READY_REQUEST reaching
# the host's in-game recvMessage, which has no case for it; ScrollableTableWidget::setMinimumColumnWidth relayouts
# with the old table width when the stats view's weapon grades column grows (a grade researched while it is open)
KNOWN_UPSTREAM = ("Didn't handle NET_READY_REQUEST message!", "lib/widget/table.cpp:689 (!shrinkableColumnIndexes.empty())")


def no_errors(log, name):
    bad, known = [], []
    for ln in log.splitlines():
        if (ln.startswith(("error", "fatal")) or "assert" in ln.lower()) and "khr_callback" not in ln:
            (known if any(k in ln for k in KNOWN_UPSTREAM) else bad).append(ln)
    check(not bad, "%s: no error or assert in the log (%d; 4.7.0's own: %d)" % (name, len(bad), len(known)))


def scenario_replay(exe):
    name = "replay"
    replay = os.path.join(ROOT, "ref", "official", "replays", "rush_10p_autogame.wzrp")
    cfg = make_config(name, files=[(replay, os.path.join("replay", "skirmish", "rush_10p_autogame.wzrp"))])
    lines = ["wait 600", "shot replay_start", "widgets",
             "click %d %d" % SPEC_BUTTON, "wait 120", "shot stats_view",
             "click %d %d" % STATS_CLOSE, "wait 60", "shot stats_closed",
             "click %d %d" % REPLAY_PAUSE, "wait 120", "log gametime", "shot replay_paused", "wait 120", "log gametime",
             "click %d %d" % REPLAY_PLAY, "wait 120", "shot replay_resumed", "log gametime",
             "click %d %d" % REPLAY_FF, "wait 60", "shot replay_ff1",
             "click %d %d" % REPLAY_FF, "wait 60", "shot replay_ff3",
             "click %d %d" % REPLAY_FF, "wait 60", "shot replay_ff0",
             "log done", "wait 5", "quit"]
    t0 = time.time()
    p = run_game(exe, name, cfg, lines, ["--loadreplay=rush_10p_autogame"])
    code = finish(p, 240, t0)
    print("%s: exit %s after %.0f s" % (name, code, time.time() - t0))
    log = report(name, cfg, ("gametime", "replay", "Replay"))
    check(code == 0, "replay: exit code 0")
    no_errors(log, name)
    return log


SPEC_AUTOHOST = """{
	"challenge": {"map": "Sk-Rush", "maxPlayers": 4, "scavengers": "false", "version": 2, "openSpectatorSlots": 1%s},
	"player_0": {"team": 0, "name": "VBhost"},
	"player_1": {"team": 1},
	"player_2": {"ai": "multiplay/skirmish/nb_generic.js", "difficulty": "Medium", "team": 2},
	"player_3": {"ai": "multiplay/skirmish/nb_generic.js", "difficulty": "Medium", "team": 3}
}
"""


def two_copies(exe, name, host_lines, client_lines, client_args, limit, autohost_extra="", host_files=(), client_cfg=None,
               host_args=()):
    """The host (--autohost=s5.json) and, 3 s later, a client with client_args; both run their scripts."""
    hcfg = make_config(name + "_host", autohost=SPEC_AUTOHOST % autohost_extra, files=host_files)
    t0 = time.time()
    host = run_game(exe, name + "_host", hcfg, host_lines, ["--autohost=s5.json", "--debug=net"] + list(host_args))
    time.sleep(3)
    ccfg = client_cfg or make_config(name + "_client")
    client = run_game(exe, name + "_client", ccfg, client_lines, client_args + ["--debug=net"])
    ccode = finish(client, limit, t0)
    hcode = finish(host, limit, t0)
    print("%s: host exit %s, client exit %s after %.0f s" % (name, hcode, ccode, time.time() - t0))
    return hcfg, ccfg, hcode, ccode


def scenario_probe(exe):
    hcfg, ccfg, hc, cc = two_copies(exe, "probe", ["wait 1500", "shot host_lobby", "quit"],
                                    ["wait 700", "click 730 77", "wait 60", "click 528 109", "wait 60", "click 795 107", "wait 180", "shot client_asked",
                                     "click 585 77", "wait 60", "click 528 147", "wait 60", "shot client_chooser2", "wait 30", "quit"], ["--spectate=127.0.0.1"], 120,
                                    host_args=["--autohost-not-ready"])
    report("probe_client", ccfg, ("dumpWidgets", "spectat", "Spectat"))


def scenario_spectator(exe):
    name = "spectator"
    # the client: Spectators tab, its row's team chooser, "Ask to Play"; then Players tab, its row's chooser, "Move to
    # Spectators"; then (the host ready) the game: the stats button and the stats view
    client = ["wait 700", "click 730 77", "wait 60", "click 528 109", "wait 60", "click 795 107", "wait 180",
              "shot moved_to_players", "click 585 77", "wait 60", "click 528 147", "wait 60", "click 792 145", "wait 180",
              "click 730 77", "wait 60", "shot back_to_spectators", "log done", "wait 5", "quit"]
    host = ["wait 1800", "shot host_lobby", "log done", "wait 5", "quit"]
    hcfg, ccfg, hc, cc = two_copies(exe, name, host, client, ["--spectate=127.0.0.1"], 200, host_args=["--autohost-not-ready"])
    hlog = report(name + "_host", hcfg, ("moved to", "Kicking"))
    clog = report(name + "_client", ccfg, ("Switching from player", "Spectator status changed", "Kicking"))
    check(hc == 0 and cc == 0, "spectator: both exit with 0 (host %s, client %s)" % (hc, cc))
    no_errors(hlog, "spectator host")
    no_errors(clog, "spectator client")
    sw = [ln for ln in clog.splitlines() if "Switching from player" in ln]
    check(len(sw) == 2, "spectator: the client switched slots twice (to Players and back): %s" % [ln.split("] ")[-1] for ln in sw])
    check("Spectator status changed! From true to false" in clog and "Spectator status changed! From false to true" in clog,
          "spectator: its spectator status changed both ways")
    check("Kicking" not in hlog, "spectator: nobody kicked")
    # then a game with a spectator (the host ready: autohost): the stats button and the stats view on the client
    name = "spectator_game"
    client = ["wait 1500", "shot game_spectator", "click %d %d" % SPEC_BUTTON, "wait 300", "shot game_stats_view",
              "wait 600", "shot game_stats_view2", "click %d %d" % STATS_CLOSE_GAME, "wait 60", "shot game_stats_closed",
              "log done", "wait 5", "quit"]
    host = ["wait 2700", "shot host_game", "log done", "wait 5", "quit"]
    hcfg, ccfg, hc, cc = two_copies(exe, name, host, client, ["--spectate=127.0.0.1"], 200)
    hlog = report(name + "_host", hcfg, ("Kicking",))
    clog = report(name + "_client", ccfg, ("Kicking",))
    check(hc == 0 and cc == 0, "spectator game: both exit with 0 (host %s, client %s)" % (hc, cc))
    no_errors(hlog, "spectator game host")
    no_errors(clog, "spectator game client")
    check("gameLoad" in clog, "spectator game: the client loaded the game")


def scenario_blind(exe):
    name = "blind"
    client = ["wait 900", "shot waiting_room", "widgets", "wait 600", "shot waiting_room2", "log done", "wait 5", "quit"]
    host = ["wait 1500", "shot host_lobby", "log done", "wait 5", "quit"]
    hcfg, ccfg, hc, cc = two_copies(exe, name, host, client, ["--join=127.0.0.1"], 150,
                                    autohost_extra=', "blindMode": "blind_lobby_simple_lobby", "spectatorHost": true', host_args=["--autohost-not-ready"])
    hlog = report(name + "_host", hcfg, ("Kicking",))
    clog = report(name + "_client", ccfg, ("WzPlayerBlindWaitingRoom", "WzRoomTitleBanner", "WzBlind", "Kicking"))
    check(hc == 0 and cc == 0, "blind: both exit with 0 (host %s, client %s)" % (hc, cc))
    no_errors(hlog, "blind host")
    no_errors(clog, "blind client")
    for cls in ("WzPlayerBlindWaitingRoom", "WzRoomTitleBanner", "WzBlindRoomHostInfo", "WzBlindRoomPlayerInfo",
                "WzBlindRoomStatusInfo", "WzBlindPlayersReadyInfo"):
        check(cls in clog, "blind: the client's waiting room has a %s" % cls)


def make_test_map(path):
    """A user map the client does not have: 2c-Roughness's files under the name 2c-S5Rough, packed as config/maps
    packs are (a .xplayers.lev at the root, the .gam and the map folder under multiplay/maps)."""
    import zipfile
    src = os.path.join(ROOT, "data", "mp", "multiplay", "maps")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("2c-S5Rough.xplayers.lev", 'level   S5Rough\nplayers 2\ntype    14\ndataset MULTI_CAM_3\n'
                   'game    "multiplay/maps/2c-S5Rough.gam"\n')
        z.write(os.path.join(src, "2c-Roughness.gam"), "multiplay/maps/2c-S5Rough.gam")
        for n in os.listdir(os.path.join(src, "2c-Roughness")):
            z.write(os.path.join(src, "2c-Roughness", n), "multiplay/maps/2c-S5Rough/" + n)


DOWNLOAD_AUTOHOST = """{
	"challenge": {"map": "S5Rough", "maxPlayers": 2, "scavengers": "false", "version": 2},
	"player_0": {"team": 0, "name": "VBhost"},
	"player_1": {"team": 1}
}
"""


def scenario_download(exe):
    name = "download"
    wz = os.path.join(WORK, "2c-S5Rough.wz")
    make_test_map(wz)
    hcfg = make_config(name + "_host", autohost=DOWNLOAD_AUTOHOST, files=[(wz, os.path.join("maps", "2c-S5Rough.wz"))])
    t0 = time.time()
    # (fork S9) after the download the client clicks Ready: the autohost starts the game on the downloaded map
    host = run_game(exe, name + "_host", hcfg, ["wait 2400", "shot host_lobby", "wait 900", "shot host_ingame", "log done", "wait 5", "quit"],
                    ["--autohost=s5.json", "--debug=net", "--debug=wz"])
    time.sleep(3)
    ccfg = make_config(name + "_client")
    client = run_game(exe, name + "_client", ccfg, ["wait 1500", "shot client_lobby", "click 775 152", "wait 1200",
                                                    "shot client_ingame", "log done", "wait 5", "quit"],
                      ["--join=127.0.0.1", "--debug=net", "--debug=wz"])
    cc = finish(client, 200, t0)
    hc = finish(host, 200, t0)
    print("%s: host exit %s, client exit %s after %.0f s" % (name, hc, cc, time.time() - t0))
    hlog = report(name + "_host", hcfg, ("Kicking", "file", "File", "S5Rough", "map package", "levLoadData"))
    clog = report(name + "_client", ccfg, ("Kicking", "file", "File", "download", "Download", "S5Rough", "map package", "levLoadData"))
    check(hc == 0 and cc == 0, "download: both exit with 0 (host %s, client %s)" % (hc, cc))
    no_errors(hlog, "download host")
    no_errors(clog, "download client")
    got = glob.glob(os.path.join(ccfg, "maps", "*.wz"))
    check(len(got) == 1 and open(got[0], "rb").read() == open(wz, "rb").read(),
          "download: the client received the host's map, byte for byte: %s" % [os.path.basename(g) for g in got])
    # (fork S9) the map package (maps/*.wz) read by buildMapList and loaded as the game's level on both sides
    for side, text in (("host", hlog), ("client", clog)):
        check("Data set found is S5Rough type 14" in text, "download %s: the level S5Rough found (a map package)" % side)
        check("Loading scenario file multiplay/maps/2c-S5Rough.gam" in text, "download %s: its scenario loaded from the package" % side)
        check("Started writing replay" in text, "download %s: the game started" % side)


def scenario_mod(exe):
    """A multiplayer mod the client lacks (in the host's config/mods/multiplay only, --mod_mp): the client
    downloads it while joining (4.7.0 keeps downloads in mods/downloads)."""
    import zipfile
    name = "mod"
    wz = os.path.join(WORK, "s5mod.wz")
    with zipfile.ZipFile(wz, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("s5mod/readme.txt", "WarzoneVB fork S5's test mod: no game data.\n" * 50)
    hcfg = make_config(name + "_host", autohost=SPEC_AUTOHOST % "", files=[(wz, os.path.join("mods", "4.7.0", "multiplay", "s5mod.wz"))])
    t0 = time.time()
    host = run_game(exe, name + "_host", hcfg, ["wait 2400", "shot host_lobby", "log done", "wait 5", "quit"],
                    ["--autohost=s5.json", "--autohost-not-ready", "--mod_mp=s5mod.wz", "--debug=net"])
    time.sleep(3)
    ccfg = make_config(name + "_client")
    client = run_game(exe, name + "_client", ccfg, ["wait 1500", "shot client_lobby", "log done", "wait 5", "quit"],
                      ["--join=127.0.0.1", "--debug=net"])
    cc = finish(client, 150, t0)
    hc = finish(host, 150, t0)
    print("%s: host exit %s, client exit %s after %.0f s" % (name, hc, cc, time.time() - t0))
    hlog = report(name + "_host", hcfg, ("Kicking", "s5mod", "file", "File"))
    clog = report(name + "_client", ccfg, ("Kicking", "s5mod", "file", "File", "ownload"))
    check(hc == 0 and cc == 0, "mod: both exit with 0 (host %s, client %s)" % (hc, cc))
    no_errors(hlog, "mod host")
    no_errors(clog, "mod client")
    got = glob.glob(os.path.join(ccfg, "mods", "downloads", "*"))  # 4.7.0 names a download by its hash
    check(any(open(g, "rb").read() == open(wz, "rb").read() for g in got),
          "mod: the client received the host's mod, byte for byte: %s" % [os.path.relpath(g, ccfg) for g in got])


# screen positions (1024x768): the spectator stats button (RET_X, 20), the stats view's close button, the replay
# controller's buttons (top right, left of the in-game options button)
SPEC_BUTTON = (24, 28)
STATS_CLOSE = (913, 16)
STATS_CLOSE_GAME = (913, 16)
REPLAY_PAUSE = (879, 38)
REPLAY_PLAY = (910, 38)
REPLAY_FF = (941, 38)


def ready_widgets(log):
    """The visible Ready buttons in a widgets dump: {player: (x, y)} (WzMultiButton MULTIOP_READY_START + player,
    MULTIOP_READY_START = MULTIOP_READY_FORM_ID (102900) + MAX_CONNECTED_PLAYERS + 1)."""
    import re
    found = {}
    for ln in log.splitlines():
        m = re.search(r"\(VB\) +(\d+) WzMultiButton \((\d+),(\d+) \d+x\d+\)( hidden)?", ln)
        if m and not m.group(4) and READY_START <= int(m.group(1)) < READY_START + 16:
            found[int(m.group(1)) - READY_START] = (int(m.group(2)), int(m.group(3)))
    return found


READY_START = 102900 + 23 + 1      # MAX_CONNECTED_PLAYERS 23: MAX_PLAYER_SLOTS 13 + MAX_SPECTATOR_SLOTS 10


def scenario_ready(exe):
    """(fork S9) The Ready boxes: 4.7.0's autohost locks Ready until the lobby is full (locked.readybeforefull, its
    default with --autohost): with an open player slot no row has a Ready box (what S5 saw). With the autohost
    file's "locked": {"readybeforefull": false} the host's row has its Ready box, and clicking it readies the host."""
    for name, extra, clicks in (("ready_default", "", []),
                                ("ready_unlocked", ', "readybeforefull": false', ["click %d %d", "wait 120", "widgets", "shot ready_clicked"])):
        autohost = SPEC_AUTOHOST.replace('"player_0"', '"locked": {"name": false%s},\n\t"player_0"' % extra) % ""
        cfg = make_config(name, autohost=autohost)
        t0 = time.time()
        lines = ["wait 600", "widgets", "shot ready_lobby"]
        p = run_game(exe, name, cfg, lines + ["log mark", "wait 30", "log done", "wait 5", "quit"], ["--autohost=s5.json", "--autohost-not-ready"])
        code = finish(p, 90, t0)
        log = report(name, cfg, ("Ready", "ready"))
        check(code == 0, "%s: exit code 0 (%s)" % (name, code))
        no_errors(log, name)
        boxes = ready_widgets(log)
        if name == "ready_default":
            check(not boxes, "ready_default: no Ready box while a player slot is open (locked.readybeforefull): %s" % boxes)
        else:
            check(0 in boxes, "ready_unlocked: the host's row has its Ready box: %s" % boxes)
            if 0 in boxes:
                # click it (its centre), then the dump again
                x, y = boxes[0][0] + 20, boxes[0][1] + 15
                cfg2 = make_config(name + "_click", autohost=autohost)
                t0 = time.time()
                p = run_game(exe, name + "_click", cfg2, lines + [c % (x, y) if "%d" in c else c for c in clicks] + ["log done", "wait 5", "quit"],
                             ["--autohost=s5.json", "--autohost-not-ready", "--debug=net"])
                code = finish(p, 90, t0)
                log2 = report(name + "_click", cfg2, ("eady",))
                check(code == 0, "ready_unlocked click: exit code 0 (%s)" % code)
                # the host ready with every present player ready (the AIs are): 4.7.0's click handler starts the game
                check("[startMultiplayerGame:" in log2 and "Sending NET_FIREUP" in log2,
                      "ready_unlocked click: the host ready, the game starts (startMultiplayerGame, NET_FIREUP)")


def main(argv):
    exe = os.path.join(ROOT, "bin", "Warzone.exe")
    if "--exe" in argv:
        exe = os.path.abspath(argv[argv.index("--exe") + 1])
    modes = [a for a in argv if a in ("replay", "spectator", "blind", "download", "mod", "probe", "ready")] or ["replay", "spectator", "blind"]   # download, mod, ready: see the docstring
    os.makedirs(WORK, exist_ok=True)
    for m in modes:
        print("== " + m)
        globals()["scenario_" + m](exe)
    failed = [w for ok, w in CHECKS if not ok]
    print("%d of %d checks passed" % (len(CHECKS) - len(failed), len(CHECKS)))
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
