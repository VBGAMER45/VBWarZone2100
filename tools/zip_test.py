"""
zip_test.py - cluster K's test of the ported PhysFS archives (WzFramework's physfs_win.bas, physfs_zip_c.bas,
src/shared/miniz: PhysFS 3.2.0's archiver dispatch, zip archiver and miniz inflate): builds bin/ZipTests.exe
(src/ziptests + those three modules + src/shared's math_vb and utf8_vb, no DLL; compiled as WzFramework is:
no optimisation, bounds checks), runs it on a command file and checks its results against Python's zipfile:

* every entry of the game's archives (data/mods/campaign/*.wz, data/sequences.wz, ref/official/wz470/data/mp.wz;
  base.wz with --base) read through the ported PhysFS - mounted, enumerated recursively, stat'ed, read whole,
  and (the mods, mp.wz and the synthetic archives; all with --all) read again in odd-sized chunks with a seek
  forward and one back, then a seek past the end - equals zipfile's entry (type, size, SHA-1, modification
  time: time.mktime of the DOS time, as PhysFS's mktime); the listing equals zipfile's names with their
  ancestor folders; the time per archive (mount, enumerate + stat, read) and the read speed;
* synthetic archives (written here byte by byte): stored and deflated entries (stored, fixed and dynamic
  Huffman blocks, long matches through the 32 KB window, empty files), zip64 (the zip64 end of central
  directory and extra fields), data before the archive (a self-extractor's), a DOS name with backslashes,
  explicit and implicit folders, symlinks (hidden while symlinks are forbidden, as Warzone sets them; read
  through, "a/../b" included, when permitted; a broken one), "traditional" PKWARE encryption
  ("name$password"), a broken local header, broken deflate data, a duplicate entry, a broken central
  directory, a non-zip file, a missing file;
* the mount points (WZTempProcessCamMod and a nested one), unmount (with a file still open: "files still
  open"; not mounted), the search path, getRealDir, folders and archives merged in one enumeration, seeks
  past the end, PhysFS's error messages, the write directory (mkdir, write, delete), and the cost of a lookup
  over folders (PHYSFS_exists, openRead + close) now that PhysFS's verifyPath checks each element for
  symlinks.

  python -I tools\\zip_test.py [--keep] [--base] [--all]      (run_tests.py calls run(keep))
"""
import hashlib
import os
import shutil
import struct
import subprocess
import sys
import time
import zipfile
import zlib

ROOT = r"C:\Users\Owner\desktop\websites\warzone"
sys.path.insert(0, os.path.join(ROOT, "tools"))
import build  # noqa: E402

FOLDER = os.path.join(ROOT, "src", "ziptests")
WORK = os.path.join(ROOT, "build", "zip_test")
PORTED = [("physfs_win", r"..\framework\physfs_win.bas"), ("physfs_zip_c", r"..\framework\physfs_zip_c.bas"),
          ("miniz_h", r"..\shared\miniz\miniz_h.bas"), ("miniz_tinfl_c", r"..\shared\miniz\miniz_tinfl_c.bas"),
          ("miniz_tdefl_c", r"..\shared\miniz\miniz_tdefl_c.bas"), ("miniz_c", r"..\shared\miniz\miniz_c.bas"),
          ("math_vb", r"..\shared\math_vb.bas"),
          ("utf8_vb", r"..\shared\utf8_vb.bas")]
MODS = ["fractured_kingdom.wz", "reclamation.wz", "wz2100_camclassic.wz", "wz2100_campumpkin.wz"]


def build_exe():
    lines = ["Type=Exe", build.STDOLE]
    for n in build.sources(FOLDER):
        path = os.path.join(FOLDER, n)
        build.normalise(path)
        lines.append("Module=%s; %s" % (build.vb_name(path), n))
    for name, rel in PORTED:
        build.normalise(os.path.normpath(os.path.join(FOLDER, rel)))
        lines.append("Module=%s; %s" % (name, rel))
    lines.append(build.VBP_TAIL.format(name="ZipTests", ext="exe", path=os.path.relpath(build.BIN, FOLDER), startup="Sub Main",
                                       compat="", bounds="0", optimisation="2", threading=""))
    vbp = os.path.join(FOLDER, "ZipTests.vbp")
    with open(vbp, "w", encoding="latin-1", newline="\r\n") as f:
        f.write("\n".join(lines))
    return build.make("ZipTests", FOLDER, vbp)


# ---------------------------------------------------------------------------------------------------------
# a zip writer with every field under control

CRC_TABLE = []
for _n in range(256):
    _c = _n
    for _k in range(8):
        _c = (0xEDB88320 ^ (_c >> 1)) if (_c & 1) else (_c >> 1)
    CRC_TABLE.append(_c)


def dos_time(dt):
    y, mo, d, h, mi, s = dt
    return (((y - 1980) << 25) | (mo << 21) | (d << 16) | (h << 11) | (mi << 5) | (s // 2)) & 0xFFFFFFFF


class Crypto:
    """traditional PKWARE encryption (APPNOTE 6.1)"""
    def __init__(self, password):
        self.k = [0x12345678, 0x23456789, 0x34567890]
        for b in password:
            self.update(b)

    def crc(self, c, b):
        return (c >> 8) ^ CRC_TABLE[(c ^ b) & 0xFF]

    def update(self, b):
        self.k[0] = self.crc(self.k[0], b)
        self.k[1] = (self.k[1] + (self.k[0] & 0xFF)) & 0xFFFFFFFF
        self.k[1] = (self.k[1] * 134775813 + 1) & 0xFFFFFFFF
        self.k[2] = self.crc(self.k[2], self.k[1] >> 24)

    def byte(self):
        t = (self.k[2] | 2) & 0xFFFF
        return ((t * (t ^ 1)) >> 8) & 0xFF

    def encrypt(self, data):
        out = bytearray()
        for p in data:
            out.append(p ^ self.byte())
            self.update(p)
        return bytes(out)

    def decrypt(self, data):
        out = bytearray()
        for c in data:
            p = c ^ self.byte()
            self.update(p)
            out.append(p)
        return bytes(out)


CRYPT_HEADER = bytes((i * 37 + 11) & 0xFF for i in range(11))


class E:
    def __init__(self, name, data=b"", method=0, system=0, attr=0, dt=(2021, 6, 15, 10, 20, 30), password=None,
                 level=6, raw_deflate=None, zip64=False):
        self.name = name if isinstance(name, bytes) else name.encode("utf-8")
        self.data, self.method, self.system, self.attr, self.dt = data, method, system, attr, dt
        self.password, self.level, self.raw_deflate, self.zip64 = password, level, raw_deflate, zip64


def write_zip(path, entries, prefix=b"", zip64=False, bad_local=None, bad_central=False):
    out = bytearray(prefix)
    base = len(prefix)
    central = bytearray()
    for e in entries:
        crc = zlib.crc32(e.data) & 0xFFFFFFFF
        if e.method == 8:
            if e.raw_deflate is not None:
                comp = e.raw_deflate
            else:
                co = zlib.compressobj(e.level, zlib.DEFLATED, -15)
                comp = co.compress(e.data) + co.flush()
        else:
            comp = e.data
        flags = 0
        if e.password is not None:
            flags |= 1
            comp = Crypto(e.password).encrypt(CRYPT_HEADER + bytes([crc >> 24]) + comp)
        offset = len(out) - base
        big = e.zip64 or zip64
        csize, usize = len(comp), len(e.data)
        lextra = struct.pack("<HHQQ", 1, 16, usize, csize) if big else b""
        local = struct.pack("<IHHHIIIIHH", 0x04034B50, 45 if big else 20, flags, e.method, dos_time(e.dt), crc,
                            0xFFFFFFFF if big else csize, 0xFFFFFFFF if big else usize, len(e.name), len(lextra))
        if bad_local is not None and e.name == bad_local:
            local = struct.pack("<I", 0x04034B51) + local[4:]
        out += local + e.name + lextra + comp
        cextra = struct.pack("<HHQQQ", 1, 24, usize, csize, offset) if big else b""
        central += struct.pack("<IHHHHIIIIHHHHHII", 0x02014B50 if not bad_central else 0x02014B52, (e.system << 8) | 20,
                               45 if big else 20, flags, e.method, dos_time(e.dt), crc,
                               0xFFFFFFFF if big else csize, 0xFFFFFFFF if big else usize, len(e.name), len(cextra), 0, 0, 0,
                               e.attr, 0xFFFFFFFF if big else offset) + e.name + cextra
    cd_offset = len(out) - base
    out += central
    n = len(entries)
    if zip64:
        eocd64 = len(out) - base
        out += struct.pack("<IQHHIIQQQQ", 0x06064B50, 44, 45, 45, 0, 0, n, n, len(central), cd_offset)
        out += struct.pack("<IIQI", 0x07064B50, 0, eocd64, 1)
        out += struct.pack("<IHHHHIIH", 0x06054B50, 0, 0, 0xFFFF, 0xFFFF, 0xFFFFFFFF, 0xFFFFFFFF, 0)
    else:
        out += struct.pack("<IHHHHIIH", 0x06054B50, 0, 0, n, n, len(central), cd_offset, 0)
    with open(path, "wb") as f:
        f.write(out)


def pattern(n, seed):
    out = bytearray()
    x = seed
    while len(out) < n:
        x = (x * 1103515245 + 12345) & 0x7FFFFFFF
        out += ("line %d of %d: %s\n" % (len(out), x % 1000, "abc" * (x % 7))).encode()
    return bytes(out[:n])


def noise(n, seed):
    out = bytearray()
    x = seed
    for _ in range(n):
        x = (x * 1664525 + 1013904223) & 0xFFFFFFFF
        out.append(x >> 24)
    return bytes(out)


def wrong_password(crc):
    """a password PhysFS's verifier byte refuses (1 in 256 would pass it)"""
    for cand in (b"wrong", b"wrong1", b"wrong2", b"nope", b"bad"):
        if Crypto(cand).decrypt(Crypto(b"pass").encrypt(CRYPT_HEADER + bytes([crc >> 24])))[11] != crc >> 24:
            return cand.decode()
    return "wrong"


SECRET = b"the secret is " + b"x" * 1000 + b"\n"
SYMLINK = 0o120777 << 16


def make_special(work):
    """the synthetic archives: {name: (path, entries)} of those walked whole, and the paths of the others"""
    walked = {}
    long_text = pattern(3 * 1024 * 1024 + 17, 7)
    entries = [
        E("stored.txt", b"Hello, stored world!\n"),
        E("empty.txt", b""),
        E("empty_deflated.txt", b"", method=8),
        E("dir/fixed.txt", b"short text: fixed Huffman codes\n", method=8),
        E("dir/sub/dynamic.txt", pattern(200000, 1), method=8, level=9),
        E("dir/sub/long.txt", long_text, method=8, level=9),
        E("dir/sub/fast.txt", pattern(150000, 3), method=8, level=1),
        E("dir/noise.bin", noise(300000, 5), method=8, level=6),
        E("dir/storedblocks.bin", noise(70000, 9), method=8, level=0),
        E("explicit/", b"", system=3, attr=(0o40755 << 16), dt=(2019, 1, 2, 3, 4, 6)),
        E("explicit/inside.txt", b"in an explicit folder\n", method=8, system=3, attr=(0o100644 << 16)),
        E(b"dosdir\\dosfile.txt", b"a DOS name\n", system=0),
        E("link", b"dir/fixed.txt", system=3, attr=SYMLINK),
        E("links/up", b"links/../dir/sub/dynamic.txt", system=3, attr=SYMLINK),
        E("Upper.txt", b"sorted before lower case\n", dt=(1999, 12, 31, 23, 59, 58)),
    ]
    p = os.path.join(work, "special.zip")
    write_zip(p, entries)
    walked["special"] = (p, entries)
    # zip64 (end records and extra fields), with data before it (a self-extractor's)
    e64 = [E("big/one.txt", pattern(100000, 11), method=8, zip64=True), E("big/two.bin", noise(5000, 2), zip64=True),
           E("three.txt", b"three\n", method=8)]
    p = os.path.join(work, "zip64.zip")
    write_zip(p, e64, zip64=True)
    walked["zip64"] = (p, e64)
    p = os.path.join(work, "prefixed64.zip")
    write_zip(p, e64, prefix=b"MZ" + b"\x00" * 4000, zip64=True)
    walked["prefixed64"] = (p, e64)
    ep = [E("a.txt", b"prefixed, not zip64\n"), E("b/c.txt", pattern(5000, 4), method=8)]
    p = os.path.join(work, "prefixed32.zip")
    write_zip(p, ep, prefix=b"#!/bin/sh\nexit 0\n" * 50)
    walked["prefixed32"] = (p, ep)
    other = {}
    p = os.path.join(work, "crypt.zip")
    write_zip(p, [E("secret.txt", SECRET, method=8, password=b"pass"), E("secret_stored.txt", b"stored secret\n", password=b"pass"),
                  E("plain.txt", b"plain\n")])
    other["crypt"] = p
    p = os.path.join(work, "brokenlink.zip")
    write_zip(p, [E("fine.txt", b"fine\n"), E("broken", b"nowhere.txt", system=3, attr=SYMLINK)])
    other["brokenlink"] = p
    p = os.path.join(work, "badlocal.zip")
    write_zip(p, [E("ok.txt", b"fine\n"), E("bad/broken.txt", b"broken local header\n")], bad_local=b"bad/broken.txt")
    other["badlocal"] = p
    p = os.path.join(work, "baddeflate.zip")
    write_zip(p, [E("garbage.txt", b"x" * 5000, method=8, raw_deflate=b"\xff" * 40)])
    other["baddeflate"] = p
    p = os.path.join(work, "dupe.zip")
    write_zip(p, [E("same.txt", b"one\n"), E("same.txt", b"two\n")])
    other["dupe"] = p
    p = os.path.join(work, "badcentral.zip")
    write_zip(p, [E("x.txt", b"x\n")], bad_central=True)
    other["badcentral"] = p
    p = os.path.join(work, "notzip.wz")
    with open(p, "wb") as f:
        f.write(b"This is not a zip archive.\n" * 40)
    other["notzip"] = p
    return walked, other


# ---------------------------------------------------------------------------------------------------------
# what PhysFS shows of an archive: {path: ("1" folder | "0" file, size, sha1, modtime)}

def mktime(dt):
    try:
        return int(time.mktime(tuple(dt) + (0, 0, -1)))
    except (OverflowError, ValueError):
        return -1


def add_entry(out, name, isdir, kind, size=None, digest=None, mt=None):
    parts = name.split("/")
    for k in range(1, len(parts)):
        out.setdefault("/".join(parts[:k]), ("1", "0", None, "0"))       # an implicit folder: no time
    if isdir:
        out[name] = ("1", "0", None, str(mt))
    else:
        out[name] = (kind, str(size), digest, str(mt))


def expected_zip(path):
    z = zipfile.ZipFile(path)
    out = {}
    for info in z.infolist():
        raw = info.orig_filename.encode("utf-8" if info.flag_bits & 0x800 else "cp437")
        isdir = raw.endswith(b"/")
        if isdir:
            raw = raw[:-1]
        if info.create_system == 0:
            raw = raw.replace(b"\\", b"/")
        name = raw.split(b"\0")[0].decode("utf-8", "replace")
        symlink = info.create_system not in (0, 1, 2, 4, 6, 11, 13, 14, 15, 18) and info.file_size > 0 and \
            (info.external_attr >> 16) & 0o170000 == 0o120000
        if symlink or info.flag_bits & 1:
            add_entry(out, name, False, "skip")
        else:
            data = b"" if isdir else z.read(info)
            add_entry(out, name, isdir, "0", info.file_size, hashlib.sha1(data).hexdigest(), mktime(info.date_time))
    return {k: v for k, v in out.items() if v[0] != "skip"}


def expected_entries(entries):
    out = {}
    for e in entries:
        nm = e.name.decode("utf-8")
        if e.system == 0:
            nm = nm.replace("\\", "/")
        isdir = nm.endswith("/")
        nm = nm.rstrip("/")
        if (e.attr >> 16) & 0o170000 == 0o120000 or e.password is not None:
            add_entry(out, nm, False, "skip")
        else:
            add_entry(out, nm, isdir, "0", len(e.data), hashlib.sha1(e.data).hexdigest(), mktime(e.dt))
    return {k: v for k, v in out.items() if v[0] != "skip"}


def sha1(b):
    return hashlib.sha1(b).hexdigest()


# ---------------------------------------------------------------------------------------------------------

def check_archive(label, lines, expect, chunked, report):
    """the A/E/T/AU lines of one archive command against the expected tree; returns (ok, details, timing)"""
    ok = True
    details = []
    got = {}
    timing = None
    mounted = lines[0] if lines else None
    unmount = None
    for r in lines[1:]:
        if r[0] == "E":
            got[r[1]] = r
        elif r[0] == "EERR":
            ok = False
            details.append("%s: enumerating %r failed: %s" % (label, r[1], r[2]))
        elif r[0] == "T":
            timing = r
        elif r[0] == "AU":
            unmount = r
    if mounted is None or mounted[0] != "A" or mounted[2] != "1":
        report.append("%s: mount failed - FAILED" % label)
        return False, ["%s: mount failed: %s" % (label, mounted)], None
    missing = sorted(set(expect) - set(got))
    extra = sorted(set(got) - set(expect))
    if missing:
        ok = False
        details.append("%s: %d entries not listed: %s" % (label, len(missing), missing[:5]))
    if extra:
        ok = False
        details.append("%s: %d entries listed that should not be: %s" % (label, len(extra), extra[:5]))
    bad = 0
    for p in sorted(set(expect) & set(got)):
        e, g = expect[p], got[p]
        why = None
        if e[0] == "1":
            if g[2:6] != ["D", "0", e[3], "1"]:
                why = "folder %s: %s, expected modtime %s" % (p, g[2:], e[3])
        else:
            n = e[1]
            if g[2] != "0" or g[3] != n:
                why = "%s: type/size %s %s, expected %s" % (p, g[2], g[3], n)
            elif g[4] != e[3]:
                why = "%s: modtime %s, expected %s" % (p, g[4], e[3])
            elif g[5] != "1" or g[6] != "1":
                why = "%s: readonly %s, opened %s" % (p, g[5], g[6])
            elif g[7] != e[2]:
                why = "%s: read whole: sha1 %s, expected %s" % (p, g[7], e[2])
            elif g[8] != n or g[9] != "1" or g[10] != n or g[11] != "0":
                why = "%s: count %s, eof %s, tell %s, next read %s" % (p, g[8], g[9], g[10], g[11])
            elif chunked and (g[12] != e[2] or g[13] != "1"):
                why = "%s: in chunks with seeks: sha1 %s (expected %s), steps %s" % (p, g[12], e[2], g[13])
            elif chunked and (g[14] != "0" or g[15] != "past end of file"):
                why = "%s: a seek past the end gave %s %s" % (p, g[14], g[15])
        if why:
            bad += 1
            ok = False
            if bad <= 8:
                details.append(label + ": " + why)
    if unmount is None or unmount[2] != "1":
        ok = False
        details.append("%s: unmount: %s" % (label, unmount))
    nfiles = sum(1 for v in expect.values() if v[0] == "0")
    report.append("%s: %d of %d entries as zipfile has them (%d files%s)%s" % (
        label, len(set(expect) & set(got)) - bad, len(expect), nfiles, ", in chunks with seeks too" if chunked else "",
        "" if ok else " - FAILED"))
    return ok, details, timing


def run(keep=False, with_base=False, chunk_all=False):
    os.makedirs(WORK, exist_ok=True)
    if not build_exe():
        print("zip: ZipTests.exe did not build")
        return False
    walked, other = make_special(WORK)
    data = os.path.join(ROOT, "data")
    blocks = []         # (label, path, chunked, mount point, expected tree)
    for m in MODS:
        p = os.path.join(data, "mods", "campaign", m)
        blocks.append(("mods/campaign/" + m, p, True, "", expected_zip(p)))
    p = os.path.join(data, "sequences.wz")
    blocks.append(("sequences.wz", p, chunk_all, "", expected_zip(p)))
    p = os.path.join(ROOT, "ref", "official", "wz470", "data", "mp.wz")
    blocks.append(("mp.wz (4.7.0's)", p, True, "", expected_zip(p)))
    if with_base:
        p = os.path.join(ROOT, "ref", "official", "wz470", "data", "base.wz")
        blocks.append(("base.wz (4.7.0's)", p, chunk_all, "", expected_zip(p)))
    for name, (p, entries) in walked.items():
        blocks.append(("synthetic " + name, p, True, "", expected_entries(entries)))
    fk = os.path.join(data, "mods", "campaign", "fractured_kingdom.wz")
    blocks.append(("fractured_kingdom.wz @ WZTempProcessCamMod", fk, False, "WZTempProcessCamMod", blocks[0][4]))

    cmds = ["init"]
    kinds = ["one"]     # per command: "archive" or "one" (one result line)
    checks = []         # (command index, command, expected fields after the echoed arguments, description)
    block_at = []
    for b in blocks:
        block_at.append(len(cmds))
        cmds.append("archive\t%s\t%s\t%d" % (b[1], b[3], 1 if b[2] else 0))
        kinds.append("archive")

    def fail(err):
        """a read command's result when the file does not open"""
        return ["0", "-2", "", "-2", "-2", "-2", err]

    def c(cmd, expect, what):
        checks.append((len(cmds), cmd, expect, what))
        cmds.append(cmd)
        kinds.append("one")

    classic = os.path.join(data, "mods", "campaign", "wz2100_camclassic.wz")
    pumpkin = os.path.join(data, "mods", "campaign", "wz2100_campumpkin.wz")
    classic_dir = os.path.join(data, "mods", "campaign", "wz2100_camclassic")
    sp = walked["special"][0]
    wdir = os.path.join(WORK, "write")
    if os.path.exists(wdir):
        shutil.rmtree(wdir)
    os.makedirs(wdir)
    fk_info = zipfile.ZipFile(fk).read("mod-info.json")

    # the mount point loadCampaignModsAsyncImpl uses
    c("mount\t%s\tWZTempProcessCamMod\t1" % fk, ["1"], "mount at WZTempProcessCamMod")
    c("exists\tWZTempProcessCamMod/mod-info.json", ["1"], "a file under the mount point")
    c("exists\tmod-info.json", ["0", "not found"], "nothing at the root")
    c("isdir\tWZTempProcessCamMod", ["1"], "the mount point is a folder")
    c("stat\tWZTempProcessCamMod/campaigns", ["1", "1", "0"], "a folder under it")
    c("realdir\tWZTempProcessCamMod/mod-info.json", [fk], "getRealDir gives the archive")
    c("realdir\tnowhere/x.txt", ["(null)", "not found"], "getRealDir of nothing: NULL")
    c("enum\t", ["1", "WZTempProcessCamMod"], "the root lists the mount point")
    c("read\tWZTempProcessCamMod/mod-info.json", ["1", str(len(fk_info)), sha1(fk_info)], "read through the mount point")
    c("keepopen\tWZTempProcessCamMod/mod-info.json", ["1"], "a file kept open")
    c("unmount\t%s" % fk, ["0", "files still open"], "unmount with a file open")
    c("closekept", ["1"], "close it")
    c("closekept", ["0", "invalid argument"], "close it again")
    c("unmount\t%s" % fk, ["1"], "unmount")
    c("unmount\t%s" % fk, ["0", "not mounted"], "unmount again")
    c("exists\tWZTempProcessCamMod/mod-info.json", ["0"], "gone")
    # a nested mount point
    c("mount\t%s\ta/b/c\t1" % classic, ["1"], "a nested mount point")
    c("enum\t", ["1", "a"], "the root lists its first element")
    c("enum\ta", ["1", "b"], "then the next")
    c("isdir\ta/b", ["1"], "part of a mount point is a folder")
    c("stat\ta/b/c/stats/body.json", ["1", "0"], "a file in it")
    c("exists\ta/b/cstats", ["0"], "no partial element match")
    c("unmount\t%s" % classic, ["1"], "unmount")
    # folders and archives merged; the search path
    c("mount\t%s\t\t1" % classic_dir, ["1"], "the classic balance's source folder")
    c("mount\t%s\t\t1" % pumpkin, ["1"], "the pumpkin balance's archive")
    c("mount\t%s\t\t0" % classic, ["1"], "the classic balance's archive, prepended")
    c("searchpath", ["3", "|".join([classic, classic_dir, pumpkin])], "the search path in order")
    c("realdir\tstats/body.json", [classic], "the first entry has it")
    c("realdir\tmod-info.json.in", [classic_dir], "only the folder has it")
    c("mount\t%s\t\t1" % classic, ["1"], "mounting again does nothing")
    c("searchpath", ["3"], "still three")
    want = sorted(set(os.listdir(os.path.join(classic_dir, "stats"))) |
                  {n.split("/", 1)[1] for z in (classic, pumpkin) for n in zipfile.ZipFile(z).namelist() if n.startswith("stats/") and n != "stats/"})
    c("enum\tstats", [str(len(want)), "|".join(want)], "a folder and archives merged, sorted, no duplicates")
    want = sorted(set(os.listdir(classic_dir)) | {n.split("/")[0] for z in (classic, pumpkin) for n in zipfile.ZipFile(z).namelist()})
    c("enum\t", [str(len(want)), "|".join(want)], "their roots merged")
    c("unmount\t%s" % classic, ["1"], "unmount")
    c("unmount\t%s" % pumpkin, ["1"], "unmount")
    c("unmount\t%s" % classic_dir, ["1"], "unmount the folder")
    c("searchpath", ["0"], "nothing mounted")
    c("read\tstats/body.json", fail("not found"), "nothing to read with an empty search path")
    # errors
    c("mount\t%s\t\t1" % os.path.join(WORK, "missing.wz"), ["0", "not found"], "a missing archive")
    c("mount\t%s\t\t1" % other["notzip"], ["0", "unsupported"], "a file that is no zip")
    c("mount\t%s\t\t1" % other["dupe"], ["0", "corrupted"], "a duplicate entry")
    c("mount\t%s\t\t1" % other["badcentral"], ["0", "corrupted"], "a broken central directory")
    c("searchpath", ["0"], "none of them mounted")
    c("mount\t%s\t\t1" % other["badlocal"], ["1"], "a broken local header mounts")
    c("read\tok.txt", ["1", "5", sha1(b"fine\n")], "its good entry reads")
    c("stat\tbad/broken.txt", ["0", "3", "-1", "-1", "1", "corrupted"], "the broken one: stat fails")
    c("read\tbad/broken.txt", fail("corrupted"), "and it does not open")
    c("enum\tbad", ["-1", "", "corrupted"], "enumerating its folder fails (the symlink filter's stat)")
    c("unmount\t%s" % other["badlocal"], ["1"], "unmount")
    c("mount\t%s\t\t1" % other["baddeflate"], ["1"], "broken deflate data mounts")
    c("seekread\tgarbage.txt\t0\t5000", ["1", "no error", "0"], "and reads nothing")
    c("lasterror", ["corrupted"], "zlib's error")
    c("unmount\t%s" % other["baddeflate"], ["1"], "unmount")
    c("mount\t%s\t\t1" % other["brokenlink"], ["1"], "an archive with a broken symlink")
    c("read\tfine.txt", ["1", "5"], "its file reads")
    c("enum\t", ["-1", "", "not found"], "its root does not enumerate (the filter's stat of the link)")
    c("unmount\t%s" % other["brokenlink"], ["1"], "unmount")
    # symlinks, encryption, seeks
    c("mount\t%s\t\t1" % sp, ["1"], "the special archive")
    c("exists\tlink", ["0", "symlinks are forbidden"], "a symlink is hidden while symlinks are forbidden")
    c("read\tlink", fail("symlinks are forbidden"), "and does not open")
    c("enum\tlinks", ["0", "", "no error"], "nor is it listed")
    c("seekread\tdir/sub/long.txt\t3000000\t100", ["1", "no error", "100", sha1(pattern(3 * 1024 * 1024 + 17, 7)[3000000:3000100])],
      "a long forward seek in deflated data")
    c("seekread\tdir/sub/long.txt\t%d\t100" % (3 * 1024 * 1024 + 18), ["0", "past end of file", "100"],
      "a seek past the end fails (the read then starts at 0)")
    c("seekread\tstored.txt\t21\t5", ["1", "no error", "0"], "a seek to the end of a stored entry")
    c("seekread\tstored.txt\t22\t5", ["0", "past end of file", "5"], "and past it")
    c("read\tdosdir/dosfile.txt", ["1", "11", sha1(b"a DOS name\n")], "a DOS name with backslashes")
    c("read\tdosdir\\dosfile.txt", fail("filename is illegal or insecure"), "a backslash in a name")
    c("permitsymlinks\t1", [], "symlinks permitted")
    c("read\tlink", ["1", "32", sha1(b"short text: fixed Huffman codes\n")], "a symlink read through")
    c("stat\tlink", ["1", "2", "0"], "a symlink's stat")
    c("read\tlinks/up", ["1", "200000", sha1(pattern(200000, 1))], "a symlink to links/../dir/sub/dynamic.txt")
    c("enum\tlinks", ["1", "up"], "listed now")
    c("permitsymlinks\t0", [], "symlinks forbidden again")
    c("unmount\t%s" % sp, ["1"], "unmount")
    c("mount\t%s\t\t1" % other["crypt"], ["1"], "an archive with encrypted entries")
    c("read\tsecret.txt$pass", ["1", str(len(SECRET)), sha1(SECRET)], "an encrypted deflated entry with its password")
    c("read\tsecret_stored.txt$pass", ["1", "14", sha1(b"stored secret\n")], "an encrypted stored entry")
    c("seekread\tsecret.txt$pass\t1000\t15", ["1", "not found", "15", sha1(SECRET[1000:1015])],
      "a seek in it (\"not found\": verifyPath's stat of the name with its password, which PhysFS leaves)")
    c("read\tsecret.txt", fail("bad password"), "no password")
    c("read\tsecret.txt$%s" % wrong_password(zlib.crc32(SECRET)), fail("bad password"), "a wrong password")
    c("read\tplain.txt$pass", fail("bad password"), "a password for a plain entry")
    c("read\tplain.txt", ["1", "6"], "the plain entry")
    c("unmount\t%s" % other["crypt"], ["1"], "unmount")
    # the write directory
    c("writedir\t%s" % wdir, ["1", wdir], "set the write dir")
    c("mkdir\tx/y/z", ["1"], "mkdir")
    c("write\tx/y/z/f.txt\thello", ["5", "1"], "write")
    c("mount\t%s\t\t1" % wdir, ["1"], "mount it")
    c("read\tx/y/z/f.txt", ["1", "5", sha1(b"hello")], "read it back")
    c("stat\tx/y", ["1", "1", "0", ], "a folder in it")
    c("seekread\tx/y/z/f.txt\t100\t5", ["1", "no error", "0"], "a seek past the end of a file in a folder succeeds")
    c("keepopen\tx/y/z/f.txt", ["1"], "keep it open")
    c("unmount\t%s" % wdir, ["0", "files still open"], "a folder with a file open does not unmount either")
    c("closekept", ["1"], "close it")
    c("delete\tx/y/z/f.txt", ["1"], "delete")
    c("delete\tx/y/z/f.txt", ["0", "not found"], "delete again")
    c("exists\tx/y/z/f.txt", ["0"], "gone")
    c("read\tx/../x", fail("filename is illegal or insecure"), "a '..' element")
    c("unmount\t%s" % wdir, ["1"], "unmount")
    # the cost of a lookup over folders (the game's search path: the write dir, data/, data/base, data/mp)
    names = []
    for sub in ("stats", "texpages", "script", "messages", "components/bodies", "audio/sfx/weapon"):
        d = os.path.join(data, "base", sub)
        if os.path.isdir(d):
            names += ["%s/%s" % (sub, n) for n in sorted(os.listdir(d))[:40] if os.path.isfile(os.path.join(d, n))]
    names += ["nosuch/%d.txt" % i for i in range(20)]
    folders = (wdir, data, os.path.join(data, "base"), os.path.join(data, "mp"))
    for d in folders:
        cmds.append("mount\t%s\t\t1" % d)
        kinds.append("one")
    cmds.append("lookups\t3\t%s" % "|".join(names))
    kinds.append("lookups")
    cmds.append("rawlookups\t3\t%s\t%s" % ("|".join(names), "|".join(folders)))
    kinds.append("one")
    for d in folders:
        cmds.append("unmount\t%s" % d)
        kinds.append("one")
    # the same over 4.7.0's archives (the write dir, mp.wz, base.wz: as the official build mounts them)
    official = os.path.join(ROOT, "ref", "official", "wz470", "data")
    for d in (wdir, os.path.join(official, "mp.wz"), os.path.join(official, "base.wz")):
        cmds.append("mount\t%s\t\t1" % d)
        kinds.append("one")
    cmds.append("lookups\t3\t%s" % "|".join(names))
    kinds.append("lookups")
    cmds.append("deinit")
    kinds.append("one")

    cmdfile = os.path.join(WORK, "commands.txt")
    resfile = os.path.join(WORK, "results.txt")
    with open(cmdfile, "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(cmds) + "\n")
    if os.path.exists(resfile):
        os.remove(resfile)
    t0 = time.time()
    r = subprocess.run([os.path.join(ROOT, "bin", "ZipTests.exe"), cmdfile, resfile], cwd=WORK, timeout=7200)
    elapsed = time.time() - t0
    if not os.path.exists(resfile):
        print("zip: ZipTests.exe wrote no results (exit %d)" % r.returncode)
        return False
    res = [l.split("\t") for l in open(resfile, encoding="utf-8").read().split("\n") if l]
    # the result lines of each command, in order
    per = []
    i = 0
    for kind in kinds:
        if kind == "archive":
            j = i + 1
            if i < len(res) and res[i][0] == "A" and res[i][2] == "1":
                while j < len(res) and res[j - 1][0] != "AU":
                    j += 1
            per.append(res[i:j])
            i = j
        elif kind == "lookups":
            per.append(res[i:i + 2])
            i += 2
        else:
            per.append(res[i:i + 1])
            i += 1
    stopped = res and res[-1][0] != "done"

    report = []
    details = []
    passed = 0
    total = 0
    timings = []
    for (label, p, chunked, mp, exp), k in zip(blocks, block_at):
        total += 1
        ok, det, tm = check_archive(label, per[k], exp, chunked, report)
        passed += ok
        details += det
        if tm:
            timings.append((label, tm))
    for k, cmd, expect, what in checks:
        total += 1
        got = per[k][0] if per[k] else ["(nothing)"]
        nargs = len(cmd.split("\t"))
        vals = got[nargs:]
        if got[0] == cmd.split("\t")[0] and vals[:len(expect)] == expect:
            passed += 1
        else:
            details.append("%s (%s): got %s, expected %s" % (what, cmd.replace("\t", " ")[:90], vals[:6], expect))

    print("zip: %d of %d checks pass (%d archives entry by entry, %d PhysFS calls); %.0f s%s"
          % (passed, total, len(blocks), len(checks), elapsed, "; ZipTests.exe stopped: %s" % res[-1] if stopped else ""))
    for line in report:
        print("  " + line)
    for d in details[:40]:
        print("  FAIL " + d)
    for label, tm in timings:
        # T path entries bytes mount stat read chunked total
        mb = float(tm[3]) / 1e6
        rd = float(tm[6]) / 1000
        print("  time %-44s %5s entries %7.1f MB: mount %6s ms, enumerate+stat %6s ms, read %6.0f ms (%5.1f MB/s)%s"
              % (label, tm[2], mb, tm[4], tm[5], float(tm[6]), mb / rd if rd > 0 else 0,
                 ", in chunks with seeks %s ms" % tm[7] if float(tm[7]) > 0 else ""))
    for k, kind in enumerate(kinds):
        if kind == "lookups":
            where = "4 folders (write dir, data, data/base, data/mp)" if "rawlookups" in cmds[k + 1] else \
                "the write dir, 4.7.0's mp.wz and base.wz"
            for l in per[k]:
                if l and l[0] == "L":
                    print("  lookups over %s: %s: %s names, %s found, %s us each" % (where, l[1], l[2], l[3], l[4]))
        elif cmds[k].startswith("rawlookups") and per[k]:
            l = per[k][0]
            print("  the same names by GetFileAttributesExW over the folders (the port's PhysFS before, without the symlink"
                  " checks): %s names, %s found, %s us each" % (l[1], l[2], l[3]))
    if not keep:
        for n in os.listdir(WORK):
            q = os.path.join(WORK, n)
            if n.endswith(".zip") or n.endswith(".wz") or n == "write":
                if os.path.isdir(q):
                    shutil.rmtree(q)
                else:
                    os.remove(q)
    return passed == total


if __name__ == "__main__":
    sys.exit(0 if run("--keep" in sys.argv, "--base" in sys.argv, "--all" in sys.argv) else 1)
