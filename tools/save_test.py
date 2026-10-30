"""
save_test.py - cluster I's save/load round trip: Warzone.exe starts a level (--game) and saves it as soon as it
starts (--saveandquit: a GTYPE_SAVE_START save, as 4.7.0's), then loads that save and saves it again; the two
save folders are compared file by file (JSON value by value). A save that reloads exactly gives the same
files but for what the clock writes (the save's time, the game info's time), and for what 4.7.0 itself
changes when it reloads a GTYPE_SAVE_START save (VOLATILE, each with the C++ that does it).

  python -I tools\\save_test.py [level] [--keep] [--paths N]

* level: a campaign level (default CAM_1A); the save is build/save_test/config/savegames/campaign/rt1, then rt2
* --paths N: how many differing JSON paths to print per file (default 8)
* Exits 0 when the saves match (the clock's fields and 4.7.0's own changes excepted).

Why 4.7.0 does not give rt1 back: a GTYPE_SAVE_START save restarts its level. levLoadData loads the save
(loadGame of rt1) and then, for GTYPE_SAVE_START, starts the mission from the level's own files
(levels.cpp:1172 -> startMission -> startMissionCampaignStart, mission.cpp:1197: loadGame of
wrf/cam1/cam1a.gam, whose gameLoadV7 sets gameType back to GTYPE_SCENARIO_START and saveGameVersion to 8,
game.cpp:4178). The reload then fires TRIGGER_GAME_LOADED (init.cpp:1779), where --saveandquit writes rt2
(qtscript.cpp:1152), while rt1 was written at TRIGGER_START_LEVEL (main.cpp:851), after processVisibility
(qtscript.cpp:1063) and the level script's eventStartLevel. So what eventStartLevel and that
processVisibility made in rt1 is not in rt2.
"""
import json
import os
import re
import shutil
import struct
import subprocess
import sys

ROOT = r"C:\Users\Owner\desktop\websites\warzone"
EXE = os.path.join(ROOT, "bin", "Warzone.exe")
WORK = os.path.join(ROOT, "build", "save_test")
CFG = os.path.join(WORK, "config")
# Differences that are not the port's, as patterns over walk()'s lines ("<path>: <detail>"; only * is special).
VOLATILE = {
    # what the clock writes, and what a save keeps of the moment it was made
    "save-info.json": ["/epoch: *", "/time: *"],
    "gameinfo.json": ["/GameProperties/current_time: *"],
    # The level's objects are built again from its map with new ids: gameLoadV set synchObjID to the save's
    # objId * 4 (game.cpp:4515), and loadWzMapDroidInit, loadWzMapFeature and loadWzMapStructure number them
    # from there; the counter then ends past the save's objId (writeGameFile game.cpp:5075, writeMainFile
    # game.cpp:4855).
    "droid.json": ["*/id: *"],
    "struct.json": ["*/id: *"],
    "feature.json": [
        "*/id: *",
        # artifact crates are placed by libcampaign's camSetArtifacts in eventStartLevel (cam1a.js)
        "/Crate@*: only in rt1",
        # visibility levels: processVisibility runs before eventStartLevel (qtscript.cpp:1063), not before
        # eventGameLoaded
        "*/visible/*: only in rt1",
    ],
    "labels.json": [
        "*/id: *",
        # loadLabels clears the labels and reads the level map's labels.json (qtscript.cpp:2057); the labels
        # eventStartLevel added (camSetArtifacts' "<pos>_artifact_label", camSetEnemyBases' groups) are not made again
        "/object/*_artifact_label: only in rt1",
        "/group/*: only in rt1",
    ],
    # startMissionCampaignStart frees the messages (mission.cpp:1202) and loadSaveMessage keeps none of a
    # GTYPE_SAVE_START save for an LDS_CAMSTART level (game.cpp:7795); rt1's are eventStartLevel's briefing
    # (camPlayVideos) and the proximity blip of an oil resource processVisibility found (visibility.cpp:799)
    "messtate.json": ["/message_*: only in rt1"],
    # startMission's initEffectsSystem (mission.cpp:439 -> shutdownEffectsSystem, effects.cpp:231) drops the
    # effects loadGame read from fxstate.json; rt1's are the landing lights of eventStartLevel's setNoGoArea
    # (mission.cpp:2798), and the level's loadGame (saveGameVersion 8) sets no landing zone
    "fxstate.json": ["/effect*: only in rt1"],
    "main.json": [
        "/objId: *",
        # Power: the level's loadGame clears it (clearPlayerPower, game.cpp:2884: gameType is
        # GTYPE_SCENARIO_START again); rt1's is eventStartLevel's setPower (cam1a.js, libcampaign's
        # __camAiPowerReset). gameLoadV's setPower (game.cpp:4631) runs after levLoadData has returned, after
        # rt2 is written, and skips GTYPE_SCENARIO_START anyway.
        "/players[*]/power: *",
        # getPlayerNames (game.cpp:2529) sets NetPlay.players[].name from gam.json's playerNames, where
        # writeGameFile put getAIName for an AI (game.cpp:5125): getPlayerName then adds it after the colour
        # (multiplay.cpp:826: "Orange-" becomes "Orange-Nexus")
        "/players[*]/name: *",
        # loadGame's game = saveGameData.sGame (game.cpp:2860) puts the static saveGameData's zeroes over what
        # loadMainFile read (game.cpp:4706, 4714): deserializeMultiplayerGame_json (game.cpp:539) does not read
        # these fields. loadMainFileFinal reads techLevel back (game.cpp:3995), but only after gameLoadV
        # returns, after rt2 is written.
        "/inactivityMinutes: *",
        "/playerLeaveMode: *",
        "/techLevel: *",
    ],
    "gam.json": [
        "/objId: *",
        "/power[*]/currentPower: *",
        # names: as main.json's (rt1's are empty: --game never names the players)
        "/netPlayers[*]/name: *",
    ],
}
# Files whose entries are numbered in list order: they get keys that do not move when rt1 has an entry rt2
# lacks (a feature is its name and position, a label its kind and name).
KEYS = {
    "feature.json": lambda k, v: "%s@%d,%d" % (v["name"], v["position"][0], v["position"][1]),
    "labels.json": lambda k, v: "%s/%s" % (k.split("_")[0], v["label"]),
}

args = sys.argv[1:]
keep = "--keep" in args
paths = 8
if "--paths" in args:
    paths = int(args[args.index("--paths") + 1])
    del args[args.index("--paths"):args.index("--paths") + 2]
args = [a for a in args if a != "--keep"]
level = args[0] if args else "CAM_1A"


def run(extra, label, limit=300):
    r = subprocess.run([EXE, "--configdir=" + CFG, "--nosound"] + extra, cwd=os.path.dirname(EXE),
                       capture_output=True, timeout=limit)
    logs = sorted(os.listdir(os.path.join(CFG, "logs")))
    log = os.path.join(CFG, "logs", logs[-1])
    bad = [l for l in open(log, encoding="utf-8", errors="replace").read().splitlines()
           if l.startswith(("error", "fatal")) or "ssert" in l]
    print("%s: exit code %d, %d error/assert line(s) in the log" % (label, r.returncode, len(bad)))
    for l in bad[:10]:
        print("  | " + l[:200])
    return r.returncode == 0


def walk(x, y, path, out):
    if type(x) != type(y) and not (isinstance(x, (int, float)) and isinstance(y, (int, float))):
        out.append("%s: %r / %r" % (path, x, y))
    elif isinstance(x, dict):
        for k in sorted(set(x) | set(y)):
            if k not in x or k not in y:
                out.append("%s/%s: only in %s" % (path, k, "rt1" if k in x else "rt2"))
            else:
                walk(x[k], y[k], path + "/" + k, out)
    elif isinstance(x, list):
        if len(x) != len(y):
            out.append("%s: %d / %d items" % (path, len(x), len(y)))
        for i, (p, q) in enumerate(zip(x, y)):
            walk(p, q, "%s[%d]" % (path, i), out)
    elif x != y:
        out.append("%s: %r / %r" % (path, x, y))


def rekey(n, obj):
    if n not in KEYS:
        return obj
    out = {}
    for k, v in obj.items():
        key = KEYS[n](k, v)
        dup = 1
        while (key if dup == 1 else "%s#%d" % (key, dup)) in out:
            dup += 1
        out[key if dup == 1 else "%s#%d" % (key, dup)] = v
    return out


def volatile(n, line):
    return any(re.fullmatch(".*".join(re.escape(s) for s in p.split("*")), line, re.S) for p in VOLATILE.get(n, ()))


def map_parts(p):
    # wzmaplib writeMapData (format 40): "map ", version, width, height, then texture and height (ULE16) per
    # tile, then gateway version 1, their number and x1 y1 x2 y2 (bytes) each
    data = open(p, "rb").read()
    w, h = struct.unpack("<II", data[8:16])
    o = 16 + w * h * 4
    n = struct.unpack("<I", data[o + 4:o + 8])[0]
    return data[:o], [data[o + 8 + 4 * i:o + 12 + 4 * i] for i in range(n)], data[o + 8 + 4 * n:]


def binary_diff(n, pa, pb):
    """None when the binary files agree, the difference otherwise (VOLATILE's rules for the binary files)."""
    da, db = open(pa, "rb").read(), open(pb, "rb").read()
    if da == db:
        return None, 0
    if n == "game.map":
        # The tiles must be the same. The gateways: gateway.cpp's list is not cleared between the save's map and
        # the level's map (each mapLoadFromWzMapData adds its map's gateways, map.cpp:1124; only
        # stageOne/TwoShutDown, saveGameReset and campaignReset call gwShutDown), so rt2 holds them twice.
        ta, ga, ra = map_parts(pa)
        tb, gb, rb = map_parts(pb)
        if ta != tb:
            return "tiles differ", 0
        if gb != ga + ga or ra != rb:
            return "gateways %d / %d (not rt1's twice)" % (len(ga), len(gb)), 0
        return None, 1
    if n == "visstate.bjo" and len(da) == len(db) and da[:8] == db[:8]:
        # rt2's explored bits must be rt1's, less those that processVisibility set before eventStartLevel
        # under the objects that became visible (setUnderTilesVis, visibility.cpp:794)
        more = [i for i in range(8, len(da)) if db[i] & ~da[i]]
        if more:
            return "%d tile byte(s) with bits only rt2 has" % len(more), 0
        return None, sum(1 for i in range(8, len(da)) if da[i] != db[i])
    return "binary files differ", 0


def label_ids(b):
    """Object labels in rt2 that name no object of rt2 (the ids are renumbered, so they are checked here)."""
    ids = set()
    for n in ("droid.json", "struct.json", "feature.json"):
        p = os.path.join(b, n)
        if os.path.exists(p):
            ids |= {v.get("id") for v in json.load(open(p, encoding="utf-8")).values()}
    p = os.path.join(b, "labels.json")
    if not os.path.exists(p):
        return []
    return [v["label"] for k, v in json.load(open(p, encoding="utf-8")).items()
            if k.startswith("object") and v.get("id") not in ids]


if os.path.exists(WORK) and not keep:
    shutil.rmtree(WORK)
os.makedirs(CFG, exist_ok=True)
ok = run(["--game=" + level, "--saveandquit=savegames/campaign/rt1.gam"], "save (%s)" % level)
ok = run(["--loadcampaign=rt1", "--saveandquit=savegames/campaign/rt2.gam"], "load and save again") and ok
a = os.path.join(CFG, "savegames", "campaign", "rt1")
b = os.path.join(CFG, "savegames", "campaign", "rt2")
if not (os.path.isdir(a) and os.path.isdir(b)):
    print("no save to compare")
    sys.exit(1)
names = sorted(set(os.listdir(a)) | set(os.listdir(b)))
same = 0
for n in names:
    pa, pb = os.path.join(a, n), os.path.join(b, n)
    if not (os.path.exists(pa) and os.path.exists(pb)):
        print("%-18s only in %s" % (n, "rt1" if os.path.exists(pa) else "rt2"))
        continue
    if n.endswith(".json"):
        out = []
        walk(rekey(n, json.load(open(pa, encoding="utf-8"))), rekey(n, json.load(open(pb, encoding="utf-8"))), "", out)
        excused = len([d for d in out if volatile(n, d)])
        out = [d for d in out if not volatile(n, d)]
        if out:
            print("%-18s %d difference(s): %s" % (n, len(out), "; ".join(out[:paths])))
            continue
    else:
        bad, excused = binary_diff(n, pa, pb)
        if bad:
            print("%-18s %s" % (n, bad))
            continue
    if excused:
        print("%-18s same but for %d difference(s) 4.7.0 makes too" % (n, excused))
    same += 1
lost = label_ids(b)
if lost:
    print("labels.json        %d object label(s) naming no object of rt2: %s" % (len(lost), ", ".join(lost[:paths])))
print("save round trip: %d of %d files the same" % (same, len(names)))
sys.exit(0 if ok and same == len(names) and not lost else 1)
