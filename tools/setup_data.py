"""Puts the game data that 4.7.0's source tree does not hold where 4.7.0 looks for it in a source tree
(ref/warzone2100/data), then mirrors that folder as data/, where the game looks (<prefix>/data):

* music/          the music (4 albums and menu.opus, Opus), from the official 4.7.0 build
                  (ref/official/wz470/data/music: the data-music-opus submodule's files);
* mods/campaign/  the four campaign mods as .wz archives, from the official build (4.7.0's CMake
                  zips the source folders beside them into these; the game skips folders there);
* terrain_overrides/classic.wz  the Classic terrain quality's textures, from the official build (4.7.0's CMake
                  downloads it; the source tree has an empty classic folder);
* terrain_overrides/high.wz     the Remastered terrain quality's textures (KTX2: UASTC, Zstandard supercompressed;
                  WzBasis transcodes them), from the official build (4.7.0's CMake downloads it too);
* sequences.wz    the campaign videos (ref/videos/low-quality-en-sequences.wz: wz-sequences release
                  v3, the release 4.7.0's data/CMakeLists.txt downloads; the low-quality English set);
* base/texpages/  the 39 textures 4.7.0's build ships as KTX2 instead of PNG (data/CMakeLists.txt's
                  WZ_BASIS_ENCODE_TEXTURES: the 30 terrain pages page-80..109, resized to 1024, and the 9 backdrops
                  in bdrops/), extracted from the official base.wz into ref/official/wz470/base_ktx2. data/ leaves
                  out the source PNGs they replace: a .png anywhere in the search path is taken before a .ktx2
                  (gfx_api.cpp's imageLoadFilenameFromInputFilename), so the source tree's 2048 pages would hide
                  terrain_overrides/high.wz's pages of the same names (Remastered drew the Normal pages, about 1.5
                  times as bright, and Arizona failed on 2048 PNGs beside its 1024 KTX2 cliffs). ref/warzone2100/data
                  keeps its PNGs. (Also not packaged by 4.7.0, but left in: the decal mips tertilesc*hw-16/32/64,
                  which nothing reads.)

Files are hard links (same volume), copied when a link fails; a file already in place is kept.

data/ is real folders of hard links to ref/warzone2100/data's files, not a junction (an earlier data/ junction
is replaced; its target is left alone): Windows re-parses a junction on every file call, and PhysFS checks
each element of a path with its own call (symbolic links forbidden, as 4.7.0 sets them), so through a junction
each lookup cost about 1.7 times as much (GetFileAttributesExW 53 us against 30 us here). Files in data/ that
ref/warzone2100/data does not have are reported, not deleted; the links of the PNGs left out (above) are removed.
usage: python -I tools\\setup_data.py [--check]   (--check: report only)
"""
import os
import shutil
import subprocess
import sys
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
REF_DATA = os.path.join(ROOT, "ref", "warzone2100", "data")
MIRROR = os.path.join(ROOT, "data")
OFFICIAL = os.path.join(ROOT, "ref", "official", "wz470", "data")
VIDEOS = os.path.join(ROOT, "ref", "videos", "low-quality-en-sequences.wz")
BASE_KTX2 = os.path.join(ROOT, "ref", "official", "wz470", "base_ktx2")    # base.wz's KTX2 files, extracted


def base_ktx2_names():
    """The official base.wz's .ktx2 entries ("texpages/page-80-water-1.ktx2", ...)."""
    with zipfile.ZipFile(os.path.join(OFFICIAL, "base.wz")) as z:
        return sorted(i.filename for i in z.infolist() if i.filename.endswith(".ktx2"))


def converted_textures():
    """The source PNGs 4.7.0's build replaces with KTX2 (see the docstring): paths relative to ref/warzone2100/data
    (and to data/), "base\\\\texpages\\\\page-80-water-1.png", ..."""
    out = []
    for n in base_ktx2_names():
        rel = os.path.join("base", *n[:-len(".ktx2")].split("/")) + ".png"
        if os.path.isfile(os.path.join(REF_DATA, rel)):
            out.append(rel)
    return out


def extract_base_ktx2(check=False):
    """base.wz's .ktx2 files into BASE_KTX2 (a file of the right size is kept). Returns how many are missing."""
    missing = 0
    with zipfile.ZipFile(os.path.join(OFFICIAL, "base.wz")) as z:
        for info in z.infolist():
            if not info.filename.endswith(".ktx2"):
                continue
            out = os.path.join(BASE_KTX2, *info.filename.split("/"))
            if os.path.isfile(out) and os.path.getsize(out) == info.file_size:
                continue
            if check:
                missing += 1
                continue
            os.makedirs(os.path.dirname(out), exist_ok=True)
            with z.open(info) as f, open(out, "wb") as g:
                shutil.copyfileobj(f, g)
    return missing


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
    for pack in ("classic.wz", "high.wz"):
        out.append((os.path.join(OFFICIAL, "terrain_overrides", pack),
                    os.path.join(REF_DATA, "terrain_overrides", pack)))
    out.append((VIDEOS, os.path.join(REF_DATA, "sequences.wz")))
    for n in base_ktx2_names():
        out.append((os.path.join(BASE_KTX2, *n.split("/")), os.path.join(REF_DATA, "base", *n.split("/"))))
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
    """data/ as real folders of hard links to every file of ref/warzone2100/data but the converted PNGs (see the
    docstring). Returns False when data/ is not (or, with check, would not be) a complete mirror."""
    left_out = {os.path.normcase(rel) for rel in converted_textures()}
    removed = 0
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
        out = os.path.normpath(os.path.join(MIRROR, rel))
        if not os.path.isdir(out):                      # empty folders too (terrain_overrides\classic, ...)
            if check:
                todo += 1
                print("not mirrored: " + os.path.relpath(out, ROOT) + "\\")
            else:
                os.makedirs(out)
        for f in files:
            src = os.path.join(folder, f)
            dst = os.path.normpath(os.path.join(MIRROR, rel, f))
            if os.path.normcase(os.path.normpath(os.path.join(rel, f))) in left_out:
                if os.path.isfile(dst):
                    if check:
                        todo += 1
                        expected.add(os.path.normcase(dst))     # (reported here, not again below)
                        if todo <= 20:
                            print("to leave out of data/: " + os.path.relpath(dst, ROOT))
                    elif os.path.samefile(src, dst):
                        os.remove(dst)                  # the link only: ref/warzone2100/data keeps the file
                        removed += 1
                    else:
                        expected.add(os.path.normcase(dst))     # not a link of it: reported below
                continue
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
    print("data/: %d files linked, %d already mirrored, %d not mirrored, %d only in data/, %d converted PNGs removed"
          % (linked, kept, todo, len(extra), removed))
    return todo == 0


def main(argv):
    check = "--check" in argv
    if extract_base_ktx2(check):
        print("base.wz's KTX2 files not extracted to " + os.path.relpath(BASE_KTX2, ROOT))
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
