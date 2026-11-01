"""
ktx2_test.py - cluster T (fork T2)'s test of WzBasis's KTX2 / UASTC transcoder (src/basis: ktx2_transcoder,
basisu_transcoder_c, basisu_uastc_c, basisu_tables_vb, and the zstd decoder's modules), compiled into test exes
(BasisTests: src/basistests + those modules; built twice, with bounds checks and optimised as WzBasis is), against
Basis Universal's own transcoder as 4.7.0 builds it (ref/basis_universal @ df079ec with its zstddeclib.c, compiled by
g++: tools/ktx2ref.cpp):

* every KTX2 file of the official 4.7.0 terrain_overrides/high.wz (579: the Remastered terrain, UASTC + zstd) and
  base.wz (39) transcoded, every level, to each format the port has (RGBA32, RGB565, BGR565, RGBA4444, ASTC 4x4, BC7,
  BC1, BC3, BC4, BC5), decode_flags 0 as 4.7.0 calls it: the output's size and FNV-1a hash equal the C's (the files of
  a mismatch are written and their first differing block named);
* the formats not ported (ETC1, ETC2, EAC R11 / RG11) fail in the port (the C transcodes them);
* the speed: the C's, the bounds-checked and the optimised exe's time per format (init + transcode, file reads out).

  python -I tools\\ktx2_test.py [--stage DIR] [--formats 13,6,...] [--raw] [--groups N] [--no-bounds] [--keep]

--stage DIR: a src\\ mirror (cluster T's staging): its basis and basistests folders' files are used where they exist.
--raw: the files rewritten without supercompression (tools/ktx2_unzstd.cpp; the C's output on them is checked to equal
       its output on the originals) instead of the originals: the KTX2 path without the zstd decoder. Chosen
       automatically when the basis folder has no zstd decoder but fork T2's stand-in.
--groups N: only the first N folders of files (a quick run); --no-bounds: only the optimised exe; --corrupt-only:
only the broken files (31 variants of one file: the header, level index, DFD, key/value data, blocks), which the port
must accept or refuse as the C does.
The work folder is build\\ktx2_test (the C references are built into its ref folder; its outputs are removed after a
passing run unless --keep).
"""
import os
import shutil
import subprocess
import sys
import time
import zipfile

ROOT = r"C:\Users\Owner\desktop\websites\warzone"
sys.path.insert(0, os.path.join(ROOT, "tools"))
import build  # noqa: E402

WORK = os.path.join(ROOT, "build", "ktx2_test")
REFDIR = os.path.join(WORK, "ref")
BASIS = os.path.join(ROOT, "ref", "basis_universal")
OFFICIAL_DATA = os.path.join(ROOT, "ref", "official", "wz470", "data")
GXX = r"C:\Strawberry\c\bin\g++.exe"
GCC = r"C:\Strawberry\c\bin\gcc.exe"

FORMATS = {13: "RGBA32", 14: "RGB565", 15: "BGR565", 16: "RGBA4444", 10: "ASTC_4x4", 6: "BC7", 2: "BC1", 3: "BC3",
           4: "BC4", 5: "BC5"}
NOT_PORTED = {0: "ETC1", 1: "ETC2", 20: "EAC_R11", 21: "EAC_RG11"}


def tool(name):
    for d in (os.path.dirname(GXX), ""):
        p = os.path.join(d, name) if d else shutil.which(name)
        if p and os.path.exists(p):
            return p
    return None


def build_refs():
    """build\\ktx2_test\\ref\\ktx2ref.exe and ktx2_unzstd.exe (g++ -O2 -DNDEBUG, as 4.7.0's release build compiles the
    transcoder: its debug asserts abort on some broken files the release code refuses or reads; the objects kept)"""
    gxx, gcc = tool("g++.exe"), tool("gcc.exe")
    if not gxx or not gcc:
        sys.exit("ktx2_test: g++ (Strawberry Perl's) not found")
    os.makedirs(REFDIR, exist_ok=True)
    inc = ["-I", os.path.join(BASIS, "transcoder"), "-I", os.path.join(BASIS, "zstd")]
    objs = {"basisu_transcoder.o": [gxx, "-O2", "-w", "-DNDEBUG"] + inc + ["-c", os.path.join(BASIS, "transcoder", "basisu_transcoder.cpp")],
            "zstddeclib.o": [gcc, "-O2", "-w", "-DNDEBUG", "-c", os.path.join(BASIS, "zstd", "zstddeclib.c")]}
    for o, cmd in objs.items():
        out = os.path.join(REFDIR, o)
        if not os.path.exists(out):
            subprocess.run(cmd + ["-o", out], check=True)
    for exe, src, extra in (("ktx2ref.exe", "ktx2ref.cpp", ["basisu_transcoder.o", "zstddeclib.o"]),
                            ("ktx2_unzstd.exe", "ktx2_unzstd.cpp", ["zstddeclib.o"])):
        out = os.path.join(REFDIR, exe)
        srcp = os.path.join(ROOT, "tools", src)
        if not os.path.exists(out) or os.path.getmtime(out) < os.path.getmtime(srcp):
            subprocess.run([gxx, "-O2", "-w", "-static"] + inc + [srcp] + [os.path.join(REFDIR, e) for e in extra] +
                           ["-o", out], check=True)


def data_groups():
    """[(group name, folder, [file paths])]: the KTX2 files of high.wz and base.wz, one group per folder (names repeat
    between folders)"""
    for wz, sub in (("terrain_overrides\\high.wz", "high"), ("base.wz", "base")):
        dest = os.path.join(WORK, "data", sub)
        if not os.path.isdir(dest):
            with zipfile.ZipFile(os.path.join(OFFICIAL_DATA, wz)) as z:
                for n in z.namelist():
                    if n.endswith(".ktx2"):
                        z.extract(n, dest)
    groups = []
    for sub in ("high", "base"):
        top = os.path.join(WORK, "data", sub)
        for dp, _, fs in sorted(os.walk(top)):
            ks = sorted(f for f in fs if f.endswith(".ktx2"))
            if ks:
                groups.append((os.path.relpath(dp, os.path.join(WORK, "data")).replace("\\", "/"), dp,
                               [os.path.join(dp, f) for f in ks]))
    return groups


def rewrite_raw(groups):
    """the groups' files without supercompression (build\\ktx2_test\\raw\\...)"""
    out = []
    for g, folder, files in groups:
        dest = os.path.join(WORK, "raw", g)
        if not os.path.isdir(dest) or len(os.listdir(dest)) != len(files):
            os.makedirs(dest, exist_ok=True)
            for k in range(0, len(files), 80):
                subprocess.run([os.path.join(REFDIR, "ktx2_unzstd.exe"), dest] + files[k:k + 80], check=True)
        out.append((g, dest, [os.path.join(dest, os.path.basename(f)) for f in files]))
    return out


def run_ref(fmt, files, outdir):
    """{name: (bytes, fnv, ms)} or (None) per file from ktx2ref (outdir '-': hashes only)"""
    res = {}
    for k in range(0, len(files), 80):
        r = subprocess.run([os.path.join(REFDIR, "ktx2ref.exe"), str(fmt), outdir] + files[k:k + 80],
                           capture_output=True, text=True)
        for ln in r.stdout.splitlines():
            p = ln.split()
            if len(p) >= 10 and p[2] == "levels":
                res[p[0]] = (int(p[5]), p[7], float(p[8]))
            elif p:
                res[p[0].rstrip(":")] = None
    return res


def source(stage, rel):
    if stage and os.path.exists(os.path.join(stage, rel)):
        return os.path.join(stage, rel)
    return os.path.join(build.SRC, rel)


def folder_files(stage, sub):
    names = set()
    for d in ([os.path.join(stage, sub)] if stage else []) + [os.path.join(build.SRC, sub)]:
        if os.path.isdir(d):
            names |= {n for n in os.listdir(d) if n.lower().endswith((".bas", ".cls"))}
    return sorted(names)


def has_zstd(stage):
    return any(n.lower().startswith("zstd_") and n.lower() != "zstd_standin_t2.bas" for n in folder_files(stage, "basis"))


def build_tests(stage, name, bounds_check, out_bin):
    """BasisTests (with bounds checks) or BasisTestsFast (optimised, no bounds checks) into out_bin"""
    vbp_dir = os.path.join(WORK, "vbp", name)
    if os.path.exists(vbp_dir):
        shutil.rmtree(vbp_dir)
    os.makedirs(vbp_dir)
    os.makedirs(out_bin, exist_ok=True)
    zstd = has_zstd(stage)
    lines = ["Type=Exe", build.STDOLE]
    enums = []
    for n in folder_files(stage, "basis"):
        p = source(stage, os.path.join("basis", n))
        low = n.lower()
        if low == "basisapi.cls":
            # the API class's enums, as a module (an exe's classes are private; the enums must be visible)
            text = open(p, encoding="latin-1").read().replace("\r\n", "\n")
            keep, inside = [], False
            for ln in text.split("\n"):
                if ln.startswith("Public Enum "):
                    inside = True
                if inside:
                    keep.append(ln)
                if ln.startswith("End Enum"):
                    inside = False
            enums = keep
            continue
        if low == "zstd_standin_t2.bas" and zstd:
            continue
        if low.endswith(".cls"):
            # a private copy (VB_Exposed = False)
            text = open(p, encoding="latin-1").read().replace("\r\n", "\n")
            text = text.replace("Attribute VB_Exposed = True", "Attribute VB_Exposed = False")
            q = os.path.join(vbp_dir, n)
            with open(q, "w", encoding="latin-1", newline="\r\n") as f:
                f.write(text)
            lines.append("Class=%s; %s" % (build.vb_name(q), q))
        else:
            build.normalise(p)
            lines.append("Module=%s; %s" % (build.vb_name(p), p))
    q = os.path.join(vbp_dir, "basistests_enums.bas")
    with open(q, "w", encoding="latin-1", newline="\r\n") as f:
        f.write('Attribute VB_Name = "basistests_enums"\n\' (generated by tools/ktx2_test.py) BasisApi.cls\'s enums\n'
                "Option Explicit\n\n" + "\n".join(enums) + "\n")
    lines.append("Module=basistests_enums; %s" % q)
    for n in folder_files(stage, "basistests"):
        p = source(stage, os.path.join("basistests", n))
        build.normalise(p)
        lines.append("Module=%s; %s" % (build.vb_name(p), p))
    for n in ("math_vb.bas", "float_vb.bas"):              # (helpers the zstd modules may use)
        lines.append("Module=%s; %s" % (n[:-4], os.path.join(build.SHARED, n)))
    lines.append(build.VBP_TAIL.format(name=name, ext="exe", path=out_bin, startup="Sub Main", compat="",
                                       bounds="0" if bounds_check else "-1", optimisation="2" if bounds_check else "0",
                                       threading=""))
    vbp = os.path.join(vbp_dir, name + ".vbp")
    with open(vbp, "w", encoding="latin-1", newline="\r\n") as f:
        f.write("\n".join(lines))
    saved = build.BIN
    build.BIN = out_bin
    try:
        return build.make(name, vbp_dir, vbp)
    finally:
        build.BIN = saved


def run_vb(exe, fmt, files, outdir, write):
    """{name: (bytes, fnv, ms) or None}, total ms"""
    os.makedirs(outdir, exist_ok=True)
    lst = os.path.join(outdir, "list.txt")
    with open(lst, "w", newline="\r\n") as f:
        f.write("\n".join(files) + "\n")
    resf = os.path.join(outdir, "results.txt")
    if os.path.exists(resf):
        os.remove(resf)
    args = [exe, str(fmt), outdir, lst] + ([] if write else ["nowrite"])
    subprocess.run(args, timeout=7200)
    res, total = {}, None
    if not os.path.exists(resf):
        return None, None
    for ln in open(resf).read().splitlines():
        p = ln.split("\t")
        if p[0] == "total":
            total = float(p[1])
        elif len(p) >= 7:
            res[p[0]] = (int(p[4]), p[5], float(p[6]))
        elif p and p[0]:
            res[p[0]] = None
    return res, total


def first_difference(a, b, unit):
    n = min(len(a), len(b))
    for i in range(n):
        if a[i] != b[i]:
            return "byte %d (block/pixel %d)" % (i, i // unit)
    return "lengths %d / %d" % (len(a), len(b))


def mode2_pattern31(data, level_ofs):
    """a change that gives the level's first UASTC mode 2 block (code 0x1D, 5 bits) the common pattern 31 (bits 20-24,
    after its 15 hint bits): past the 30 patterns, unpack_uastc refuses it"""
    def f(d):
        for b in range(level_ofs, len(d) - 16, 16):
            if d[b] & 0x1F == 0x1D:
                d[b + 2] |= 0xF0
                d[b + 3] |= 0x01
                break
        return d
    return f


def corrupt_cases(src, exe):
    """broken variants of a KTX2 file without supercompression (the header, the level index, the DFD, the key/value
    data, a block): the port must accept and refuse what the C accepts and refuses, with the same output"""
    import struct
    data = bytearray(open(src, "rb").read())
    dfd = struct.unpack_from("<I", data, 48)[0]
    kvd_ofs, kvd_len = struct.unpack_from("<II", data, 56)
    lvl0_ofs = struct.unpack_from("<Q", data, 80)[0]
    last_ofs = min(struct.unpack_from("<Q", data, 80 + 24 * i)[0] for i in range(struct.unpack_from("<I", data, 40)[0]))

    def put(fmt, ofs, v):
        def f(d):
            struct.pack_into(fmt, d, ofs, v)
            return d
        return f

    cases = [
        ("intact", lambda d: d),
        ("80 bytes", lambda d: d[:80]),
        ("half", lambda d: d[:len(d) // 2]),
        ("identifier", lambda d: d[:5] + b"X" + d[6:]),
        ("vk_format 37", put("<I", 12, 37)),
        ("type_size 2", put("<I", 16, 2)),
        ("width 0", put("<I", 20, 0)),
        ("depth 1", put("<I", 28, 1)),
        ("faces 2", put("<I", 36, 2)),
        ("faces 6 (not square)", put("<I", 36, 6) if struct.unpack_from("<II", data, 20)[0] != struct.unpack_from("<II", data, 20)[1] else put("<I", 36, 6)),
        ("levels 0", put("<I", 40, 0)),
        ("levels 17", put("<I", 40, 17)),
        ("supercompression 3", put("<I", 44, 3)),
        ("supercompression 1 (BasisLZ)", put("<I", 44, 1)),
        ("dfd length 48", put("<I", 52, 48)),
        ("dfd offset 70", put("<I", 48, 70)),
        ("level 0 offset 10", put("<Q", 80, 10)),
        ("level 0 length huge", put("<Q", 88, 0x7FFFFFF0)),
        ("level 0 length 16", put("<Q", 88, 16)),
        ("level 0 uncompressed 0", put("<Q", 96, 0)),
        ("transfer function 0", put("<B", dfd + 14, 0)),
        ("transfer function sRGB", put("<B", dfd + 14, 2)),
        ("color model 160", put("<B", dfd + 12, 160)),
        ("color model ETC1S", put("<B", dfd + 12, 163)),
        ("channel RGBA", put("<B", dfd + 31, 0x30)),
        ("kvd offset 0", put("<I", 56, 0) if kvd_len else put("<I", 60, 8)),
        ("kvd length 3", put("<I", 60, 3)),
        ("key length 1", put("<I", kvd_ofs, 1) if kvd_len else (lambda d: d)),
        ("block mode 19", put("<B", lvl0_ofs, 0x45)),
        ("smallest level mode 19", put("<B", last_ofs, 0x45)),
        ("block pattern 31", mode2_pattern31(data, lvl0_ofs)),
    ]
    os.makedirs(os.path.join(WORK, "corrupt"), exist_ok=True)
    files = []
    for k, (name, f) in enumerate(cases):
        p = os.path.join(WORK, "corrupt", "case%02d.ktx2" % k)
        with open(p, "wb") as fo:
            fo.write(bytes(f(bytearray(data))))
        files.append((name, p))
    bad, c_fails, c_crashes = 0, set(), set()
    for fmt in (13, 6, 10, 3):
        res, _ = run_vb(exe, fmt, [p for _, p in files], os.path.join(WORK, "corrupt", "out%d" % fmt), False)
        for name, p in files:
            n = os.path.basename(p)
            # the C one file at a time: some broken files crash it (its output, buffered, is then lost)
            r = subprocess.run([os.path.join(REFDIR, "ktx2ref.exe"), str(fmt), "-", p], capture_output=True, text=True)
            ref = {}
            for ln in r.stdout.splitlines():
                q = ln.split()
                if len(q) >= 10 and q[2] == "levels":
                    ref[q[0]] = (int(q[5]), q[7])
                elif q:
                    ref[q[0].rstrip(":")] = None
            v = (res or {}).get(n, "missing")
            if n not in ref:
                c_crashes.add(name)
                # the C crashes (undefined behaviour on this file): the port must refuse it
                if v is not None:
                    bad += 1
                    print("ktx2: corrupt %s (%s), format %d: the C crashes (exit %d), the port transcodes it" % (n, name, fmt, r.returncode))
                continue
            rr = ref[n]
            if rr is None:
                c_fails.add(name)
            if (rr is None) != (v is None) or (rr and v and v != "missing" and rr[:2] != v[:2]) or v == "missing":
                bad += 1
                print("ktx2: corrupt %s (%s), format %d: C %s, port %s" % (n, name, fmt, rr, v))
    print("ktx2: %d broken files x 4 formats, %d differ (the C refuses %d of them, crashes on %d: %s)"
          % (len(files), bad, len(c_fails), len(c_crashes), ", ".join(sorted(c_crashes))))
    return bad


def main():
    args = sys.argv[1:]
    stage = None
    if "--stage" in args:
        stage = os.path.abspath(args[args.index("--stage") + 1])
    fmts = list(FORMATS)
    if "--formats" in args:
        fmts = [int(x) for x in args[args.index("--formats") + 1].split(",")]
    keep = "--keep" in args
    raw = "--raw" in args or not has_zstd(stage)
    ngroups = int(args[args.index("--groups") + 1]) if "--groups" in args else None
    exes = [("BasisTestsFast", False)] if "--no-bounds" in args else [("BasisTests", True), ("BasisTestsFast", False)]

    os.makedirs(WORK, exist_ok=True)
    build_refs()
    groups = data_groups()
    if ngroups:
        groups = groups[:ngroups]
    print("ktx2: %d files in %d folders%s" % (sum(len(g[2]) for g in groups), len(groups),
                                             " (rewritten without supercompression: --raw)" if raw else ""))
    failures = 0
    if raw:
        rgroups = rewrite_raw(groups)
        # the rewrite is checked: the C's BC7 output of each raw file equals its output of the original
        for (g, _, files), (_, _, rfiles) in zip(groups, rgroups):
            a, b = run_ref(6, files, "-"), run_ref(6, rfiles, "-")
            bad = [n for n in a if a[n] is None or b.get(n) is None or a[n][:2] != b[n][:2]]
            if bad:
                failures += len(bad)
                print("ktx2: the raw rewrite changes the C's output: %s %s" % (g, bad[:5]))
        groups = rgroups

    out_bin = os.path.join(WORK, "bin")
    for name, bounds in exes:
        if not build_tests(stage, name, bounds, out_bin):
            print("ktx2: %s did not build" % name)
            return 1

    # broken files: one 256x256 file without supercompression (the raw rewrite of a high.wz tile)
    allfiles = [f for g in groups for f in g[2]]
    tile = next((f for f in allfiles if "tile-" in f), allfiles[0])
    if not raw:
        rewrite_dir = os.path.join(WORK, "corrupt_src")
        os.makedirs(rewrite_dir, exist_ok=True)
        subprocess.run([os.path.join(REFDIR, "ktx2_unzstd.exe"), rewrite_dir, tile], check=True)
        tile = os.path.join(rewrite_dir, os.path.basename(tile))
    failures += corrupt_cases(tile, os.path.join(out_bin, exes[-1][0] + ".exe"))
    if "--corrupt-only" in args:
        print("ktx2: %s" % ("PASSED" if failures == 0 else "%d FAILED" % failures))
        return 0 if failures == 0 else 1

    speed = []
    for fmt in fmts + sorted(NOT_PORTED):
        expect_fail = fmt in NOT_PORTED
        fname = FORMATS.get(fmt) or NOT_PORTED[fmt]
        if expect_fail and fmt not in fmts and "--formats" in args:
            continue
        cms, vms, compared, bad = 0.0, {}, 0, 0
        for g, folder, files in groups:
            ref = run_ref(fmt, files, "-")
            cms += sum(v[2] for v in ref.values() if v)
            for exe, _ in exes:
                outdir = os.path.join(WORK, "out", exe, g.replace("/", "_"), str(fmt))
                res, total = run_vb(os.path.join(out_bin, exe + ".exe"), fmt, files, outdir, False)
                if res is None:
                    print("ktx2: %s %s %s wrote no results" % (exe, fname, g))
                    failures += 1
                    continue
                vms[exe] = vms.get(exe, 0.0) + (total or 0.0)
                for f in files:
                    n = os.path.basename(f)
                    r, v = ref.get(n), res.get(n)
                    compared += 1
                    if expect_fail:
                        if v is not None:
                            bad += 1
                            print("ktx2: %s %s/%s: the port transcoded a format it does not have" % (fname, g, n))
                        continue
                    if r is None or v is None or r[:2] != v[:2]:
                        bad += 1
                        if bad <= 5:
                            # write both outputs and find the first difference
                            refout = os.path.join(WORK, "diff", "ref")
                            os.makedirs(refout, exist_ok=True)
                            subprocess.run([os.path.join(REFDIR, "ktx2ref.exe"), str(fmt), refout, f], capture_output=True)
                            run_vb(os.path.join(out_bin, exe + ".exe"), fmt, [f], os.path.join(WORK, "diff", exe), True)
                            pa = os.path.join(refout, "%s.%d" % (n, fmt))
                            pb = os.path.join(WORK, "diff", exe, "%s.%d" % (n, fmt))
                            where = ""
                            if os.path.exists(pa) and os.path.exists(pb):
                                unit = build_unit(fmt)
                                where = first_difference(open(pa, "rb").read(), open(pb, "rb").read(), unit)
                            print("ktx2: %s %s %s/%s differs: C %s, port %s %s" % (exe, fname, g, n, r, v, where))
        failures += bad
        line = "ktx2: %-9s %5d checks, %d failed; C %.0f ms" % (fname, compared, bad, cms)
        for exe, _ in exes:
            if exe in vms:
                line += ", %s %.0f ms (x%.1f)" % (exe, vms[exe], vms[exe] / cms if cms else 0)
        if expect_fail:
            line += " (not ported: the port must fail)"
        print(line)
        speed.append(line)
    print("ktx2: %s" % ("PASSED" if failures == 0 else "%d FAILED" % failures))
    if failures == 0 and not keep:
        shutil.rmtree(os.path.join(WORK, "out"), ignore_errors=True)
    return 0 if failures == 0 else 1


def run(keep=False):
    """run_tests.py's entry (its MEDIA_TESTS): the optimised exe only (--no-bounds, about 6 minutes); True when
    everything passed"""
    old = sys.argv
    sys.argv = [old[0], "--no-bounds"] + (["--keep"] if keep else [])
    try:
        return main() == 0
    finally:
        sys.argv = old


def build_unit(fmt):
    if fmt in (13,):
        return 4
    if fmt in (14, 15, 16):
        return 2
    if fmt in (2, 4):
        return 8
    return 16


if __name__ == "__main__":
    sys.exit(main())
