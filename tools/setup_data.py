"""Puts the game data that 4.7.0's source tree does not hold where 4.7.0 looks for it in a source tree
(ref/warzone2100/data), then mirrors that folder as data/, where the game looks (<prefix>/data):

* music/          the music (4 albums and menu.opus, Opus), from the official 4.7.0 build
                  (ref/official/wz470/data/music: the data-music-opus submodule's files);
* mods/campaign/  the four campaign mods as .wz archives, from the official build (4.7.0's CMake
                  zips the source folders beside them into these; the game skips folders there);
* sequences.wz    the campaign videos (ref/videos/low-quality-en-sequences.wz: wz-sequences release
                  v3, the release 4.7.0's data/CMakeLists.txt downloads; the low-quality English set).

Files are hard links (same volume), copied when a link fails; a file already in place is kept.

data/ is real folders of hard links to ref/warzone2100/data's files, not a junction (an earlier data/ junction
is replaced; its target is left alone): Windows re-parses a junction on every file call, and PhysFS checks
each element of a path with its own call (symbolic links forbidden, as 4.7.0 sets them), so through a junction
each lookup cost about 1.7 times as much (GetFileAttributesExW 53 us against 30 us here). Files in data/ that
ref/warzone2100/data does not have are reported, not deleted.
usage: python -I tools\\setup_data.py [--check]   (--check: report only)
"""
import os
import shutil
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
REF_DATA = os.path.join(ROOT, "ref", "warzone2100", "data")
MIRROR = os.path.join(ROOT, "data")
OFFICIAL = os.path.join(ROOT, "ref", "official", "wz470", "data")
VIDEOS = os.path.join(ROOT, "ref", "videos", "low-quality-en-sequences.wz")


def pairs():
    """(source, destination) of every file to put in place."""
    out = []
    music = os.path.join(OFFICIAL, "music")
    for folder, _dirs, files in os.walk(music):
        for f in files:
            src = os.path.join(folder, f)
            out.append((src, os.path.join(REF_DATA, "music", os.path.relpath(src, music))))
    mods = os.path.join(OFFICIAL, "mods", "campaign")
    for f in sorted(os.listdir(mods)):
        if f.endswith(".wz"):
            out.append((os.path.join(mods, f), os.path.join(REF_DATA, "mods", "campaign", f)))
    out.append((VIDEOS, os.path.join(REF_DATA, "sequences.wz")))
    return out


def place(src, dst):
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    if os.path.exists(dst):
        os.remove(dst)
    try:
        os.link(src, dst)
    except OSError:
        shutil.copyfile(src, dst)


def sync_mirror(check=False):
    """data/ as real folders of hard links to every file of ref/warzone2100/data (see the docstring).
    Returns False when data/ is not (or, with check, would not be) a complete mirror."""
    if os.path.isjunction(MIRROR):
        if check:
            print("data/ is a junction (a mirror of hard links is faster: run without --check)")
            return False
        # rmdir removes the junction itself, never its target's files
        r = subprocess.run(["cmd", "/c", "rmdir", MIRROR], capture_output=True, text=True)
        if r.returncode != 0:
            print("could not remove the data/ junction: " + (r.stdout + r.stderr).strip())
            return False
        print("data/: the junction removed (ref/warzone2100/data untouched)")
    linked = kept = todo = 0
    expected = set()
    for folder, _dirs, files in os.walk(REF_DATA):
        rel = os.path.relpath(folder, REF_DATA)
        for f in files:
            src = os.path.join(folder, f)
            dst = os.path.normpath(os.path.join(MIRROR, rel, f))
            expected.add(os.path.normcase(dst))
            if os.path.isfile(dst) and os.path.samefile(src, dst):
                kept += 1
                continue
            if check:
                todo += 1
                if todo <= 20:
                    print("not mirrored: " + os.path.relpath(dst, ROOT))
                continue
            place(src, dst)
            linked += 1
    extra = []
    for folder, _dirs, files in os.walk(MIRROR):
        for f in files:
            p = os.path.join(folder, f)
            if os.path.normcase(os.path.normpath(p)) not in expected:
                extra.append(os.path.relpath(p, ROOT))
    for p in extra[:20]:
        print("in data/ only (not in ref/warzone2100/data; left alone): " + p)
    print("data/: %d files linked, %d already mirrored, %d not mirrored, %d only in data/"
          % (linked, kept, todo, len(extra)))
    return todo == 0


def main(argv):
    check = "--check" in argv
    missing = [s for s, _ in pairs() if not os.path.isfile(s)]
    if missing:
        print("missing sources:\n  " + "\n  ".join(missing))
        return 1
    placed = present = 0
    for src, dst in pairs():
        if os.path.isfile(dst) and os.path.getsize(dst) == os.path.getsize(src):
            present += 1
            continue
        if check:
            print("not in place: " + os.path.relpath(dst, ROOT))
            continue
        place(src, dst)
        placed += 1
    print("ref/warzone2100/data: %d files placed, %d already in place" % (placed, present))
    return 0 if sync_mirror(check) else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
