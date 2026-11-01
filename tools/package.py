"""
package.py - makes a copy of the built game to give to other people: one folder (and optionally a zip) with
the game, its data, a launcher, the licences and the source.

  python -I tools\\package.py [--zip] [--wz] [--no-source] [--out DIR]

* dist\\WarzoneVB-<date>\\ (or --out): bin\\ (Warzone.exe and the eight DLLs only: no test exes, .exp, .lib or
  stray .OBJ), data\\ (real copies of data\\'s files: base, mp, fonts, music, mods, terrain_overrides,
  sequences.wz; --wz takes 4.7.0's own base.wz and mp.wz from the official build instead of the base and mp
  folders, as an installed 4.7.0 has them), Play.cmd, README.txt, 4.7.0's COPYING, COPYING.NONGPL,
  COPYING.README and AUTHORS, libplum's LICENSE (MPL 2.0) as LICENSE.libplum.txt, Basis Universal's LICENSE
  (Apache 2.0) as LICENSE.basisu.txt, zstd's LICENSE (BSD) as LICENSE.zstd.txt, and
  source\\WarzoneVB-source.zip (src, tools, docs: the GPL's corresponding source of this build; --no-source
  leaves it out, and then the source must be offered another way).
* Play.cmd starts bin\\Warzone.exe with --configdir=<the package>\\userdata: the settings, saves and logs stay
  in the package folder. Without it the port uses %APPDATA%\\Warzone 2100 Project\\<version>, the same
  folder as an installed official 4.7.0 (4.7.0's portable mode is not ported).
* bin\\ runs without installing anything: the exe's embedded manifest names the DLLs' classes (reg-free
  COM); MSVBVM60.DLL, OpenGL, GDI+ and chakra.dll come with Windows.
* --zip also writes dist\\WarzoneVB-<date>.zip (stored: the data is compressed already).
* Checks the copy: every file the game needs is there, and the packaged exe reaches the main menu and quits
  (tools/ui_tests/mainmenu.txt, with its own userdata).
"""
import datetime
import os
import shutil
import subprocess
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import build  # noqa: E402

ROOT = build.ROOT
DLLS = [p[0] + ".dll" for p in build.PROJECTS if p[1] == "dll"]
DATA = os.path.join(ROOT, "data")
OFFICIAL_DATA = os.path.join(ROOT, "ref", "official", "wz470", "data")
DATA_ITEMS = ["base", "mp", "fonts", "music", "mods", "terrain_overrides", "sequences.wz"]
LICENCES = ["COPYING", "COPYING.NONGPL", "COPYING.README", "AUTHORS"]
# libplum (WzNet's port mapping, ported from ref/libplum): the Mozilla Public License 2.0
LIBPLUM_LICENCE = os.path.join(ROOT, "ref", "libplum", "LICENSE")
# Basis Universal's transcoder and zstd's decoder (WzBasis, ported from ref/basis_universal): Apache 2.0, BSD
BASISU_LICENCE = os.path.join(ROOT, "ref", "basis_universal", "LICENSE")
ZSTD_LICENCE = os.path.join(ROOT, "ref", "basis_universal", "zstd", "LICENSE")
SOURCE_DIRS = ["src", "tools", "docs"]
SOURCE_SKIP_EXT = (".obj", ".exe", ".dll", ".exp", ".lib", ".log", ".pyc")

PLAY_CMD = r"""@echo off
rem WarzoneVB: starts the game with its settings, saves and logs in this folder (userdata).
rem (The game stops at once when its --configdir folder does not exist, as 4.7.0 does: make it first.)
if not exist "%~dp0userdata\" mkdir "%~dp0userdata"
start "" "%~dp0bin\Warzone.exe" "--configdir=%~dp0userdata" %*
"""

README = """WarzoneVB {date}
================

Warzone 2100 4.7.0 ported to Visual Basic 6 (an unofficial port, not made or supported by the
Warzone 2100 Project).

Start: double-click Play.cmd. The settings, saves and logs go into the userdata folder here.
(bin\\Warzone.exe also runs on its own, but then it uses the same settings folder as an installed
official Warzone 2100 4.7.0: %APPDATA%\\Warzone 2100 Project.)

Needs: Windows 10 or 11, a graphics driver with OpenGL 3.3. Nothing is installed: delete the folder to
remove it.
{network}
Licences: the game and this port are under the GNU General Public License, version 2 or later
(COPYING); the data's other licences are in COPYING.NONGPL and COPYING.README; the authors of
Warzone 2100 are in AUTHORS. WzNet.dll's port mapping is a port of libplum (Copyright Paul-Louis
Ageneau), under the Mozilla Public License 2.0 (LICENSE.libplum.txt); its source is the plum_* modules
of src\\net in the source zip. WzBasis.dll is a port of Basis Universal's transcoder (Copyright
Binomial LLC), under the Apache License 2.0 (LICENSE.basisu.txt), and of Zstandard's decoder (Copyright
Facebook, Inc.), under the BSD licence (LICENSE.zstd.txt); their source is src\\basis in the source zip.
{source}
"""

NETWORK_NOTE = """
Multiplayer over a network: the first time you host, Windows Firewall asks whether Warzone.exe may
use the network; allow it for the networks you play on. As in Warzone 2100 4.7.0, a hosted game is
listed in the internet lobby (Join Game shows the lobby's games), and the game asks your router to
open its port (2100) through UPnP / NAT-PMP / PCP; the lobby lists the game only when other players
can reach that port. Port mapping can be turned off in Options > Defaults > Hosting Options.
"""


def copy_tree(src, dst):
    n = 0
    for folder, _dirs, files in os.walk(src):
        out = os.path.join(dst, os.path.relpath(folder, src))
        os.makedirs(out, exist_ok=True)
        for f in files:
            shutil.copyfile(os.path.join(folder, f), os.path.join(out, f))
            n += 1
    return n


def source_zip(path):
    n = 0
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for d in SOURCE_DIRS:
            for folder, dirs, files in os.walk(os.path.join(ROOT, d)):
                dirs[:] = [x for x in dirs if x != "__pycache__"]
                for f in files:
                    if f.lower().endswith(SOURCE_SKIP_EXT):
                        continue
                    p = os.path.join(folder, f)
                    z.write(p, os.path.relpath(p, ROOT))
                    n += 1
    return n


def check_runs(pkg):
    exe = os.path.join(pkg, "bin", "Warzone.exe")
    cfg = os.path.join(pkg, "userdata")
    os.makedirs(cfg, exist_ok=True)
    script = os.path.join(ROOT, "tools", "ui_tests", "mainmenu.txt")
    r = subprocess.run([exe, "--configdir=" + cfg, "--vbscript=" + script, "--nosound"], cwd=os.path.dirname(exe),
                       timeout=300)
    logs = os.path.join(cfg, "logs")
    found = ""
    for f in sorted(os.listdir(logs)) if os.path.isdir(logs) else []:
        if f.startswith("WZlog"):
            for line in open(os.path.join(logs, f), encoding="utf-8", errors="replace"):
                if "gamedesc.lev found" in line or "error" in line.lower():
                    found += "  " + line.strip()[:160] + "\n"
    shutil.rmtree(cfg)                             # the package ships without a userdata folder
    return r.returncode, found


def main(argv):
    use_wz = "--wz" in argv
    with_source = "--no-source" not in argv
    date = datetime.date.today().isoformat()
    out = os.path.join(ROOT, "dist")
    if "--out" in argv:
        out = argv[argv.index("--out") + 1]
    pkg = os.path.join(out, "WarzoneVB-" + date)
    if os.path.exists(pkg):
        shutil.rmtree(pkg)
    os.makedirs(os.path.join(pkg, "bin"))
    missing = [f for f in ["Warzone.exe"] + DLLS if not os.path.isfile(os.path.join(build.BIN, f))]
    if missing:
        print("not built: " + ", ".join(missing) + " (python -I tools\\build.py)")
        return 1
    for f in ["Warzone.exe"] + DLLS:
        shutil.copyfile(os.path.join(build.BIN, f), os.path.join(pkg, "bin", f))
    files = 0
    os.makedirs(os.path.join(pkg, "data"))
    for item in DATA_ITEMS:
        src = os.path.join(DATA, item)
        if use_wz and item in ("base", "mp"):
            shutil.copyfile(os.path.join(OFFICIAL_DATA, item + ".wz"), os.path.join(pkg, "data", item + ".wz"))
            files += 1
        elif os.path.isdir(src):
            files += copy_tree(src, os.path.join(pkg, "data", item))
        elif os.path.isfile(src):
            shutil.copyfile(src, os.path.join(pkg, "data", item))
            files += 1
        else:
            print("missing data: %s (python -I tools\\setup_data.py)" % src)
            return 1
    for f in LICENCES:
        shutil.copyfile(os.path.join(build.REF, f), os.path.join(pkg, f))
    shutil.copyfile(LIBPLUM_LICENCE, os.path.join(pkg, "LICENSE.libplum.txt"))
    shutil.copyfile(BASISU_LICENCE, os.path.join(pkg, "LICENSE.basisu.txt"))
    shutil.copyfile(ZSTD_LICENCE, os.path.join(pkg, "LICENSE.zstd.txt"))
    with open(os.path.join(pkg, "Play.cmd"), "w", newline="\r\n") as f:
        f.write(PLAY_CMD)
    source_note = ""
    if with_source:
        os.makedirs(os.path.join(pkg, "source"))
        n = source_zip(os.path.join(pkg, "source", "WarzoneVB-source.zip"))
        source_note = ("The port's source (Visual Basic 6) is in source\\WarzoneVB-source.zip; Warzone 2100's own "
                       "source is at\nhttps://github.com/Warzone2100/warzone2100 (tag 4.7.0).")
        print("source: %d files zipped" % n)
    else:
        source_note = "(This copy has no source: offer it to whoever you give the game.)"
    with open(os.path.join(pkg, "README.txt"), "w", newline="\r\n") as f:
        f.write(README.format(date=date, network=NETWORK_NOTE, source=source_note))
    code, lines = check_runs(pkg)
    print("packaged exe: exit code %d (the main menu reached and quit)" % code)
    print(lines, end="")
    size = sum(os.path.getsize(os.path.join(d, f)) for d, _, fs in os.walk(pkg) for f in fs)
    print("%s: %d data files, %.0f MB" % (pkg, files, size / 1e6))
    if code != 0:
        return 1
    if "--zip" in argv:
        zpath = pkg + ".zip"
        with zipfile.ZipFile(zpath, "w", zipfile.ZIP_STORED) as z:
            for d, _, fs in os.walk(pkg):
                for f in fs:
                    p = os.path.join(d, f)
                    z.write(p, os.path.relpath(p, out))
        print("%s: %.0f MB" % (zpath, os.path.getsize(zpath) / 1e6))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
