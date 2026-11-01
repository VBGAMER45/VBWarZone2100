"""
zstd_test.py - cluster T (fork T1)'s test of WzBasis's zstd decoder (src/basis/zstd_*.bas: zstd 1.4.9's decoder
as ref/basis_universal/zstd/zstddeclib.c has it, the single-file decoder Basis Universal ships and 4.7.0 builds,
ported with the 32-bit build's code paths), compiled into test exes (src/zstdtests; ZstdTests with array bounds
checks, ZstdTestsFast optimised as WzBasis is), against two C references built by gcc:

* z32.dll: zstddeclib.c itself compiled with gcc -m32 (tools/zstdref: a freestanding DLL, its few runtime functions
  supplied there, as this gcc has no 32-bit libraries), so its code paths are the 32-bit build's. The test exe
  loads it and runs it on every frame beside the port: the same result (size or error code) and the same bytes;
  the port's writes stay inside dstCapacity (guard bytes around the destination).
* zstdref64.dll: ref/basis_universal/zstd/zstd.c (the whole 1.4.9 library, compressor included) compiled for
  x64: it makes the frames (levels -5 to 22, strategies, window logs up to 27 with long distance matching,
  checksums, unknown content sizes, raw literals, small blocks) and decompresses every valid one; the port's
  output hash must equal it.

Frames: inputs of many shapes (empty, 1 byte, text, source code, a DLL, random bytes, runs, mixtures, > 128 KB
and multi-megabyte ones, 80 MB with 40 MB long-distance repeats for the 32-bit long-offset path and the
prefetching sequence decoder), several frames in a row and skippable frames, crafted headers (window log 30
and 31: the 32-bit limit is 30; a dictionary ID; reserved bit; huge content size), broken frames (cut short,
bytes flipped) and too-small destinations (their errors must equal z32's; the 64-bit build's are listed when
they differ: the window log 31 frames), and every level of the 579 KTX2 files in the official high.wz (5391
frames, 99 MB in, 169 MB out), each with the uncompressed size its level index gives.

  python -I tools\\zstd_test.py [--stage DIR] [--keep] [--quick] [--no-ktx2]

--stage DIR: a src\\ mirror (cluster T's staging): its basis\\zstd_*.bas and zstdtests\\ are used where they exist.
--keep: keep build\\zstd_test\\vectors (else regenerated). --quick: fewer generated frames. --no-ktx2: skip high.wz.
--only A,B: only the generated frames whose names start with A or B.
Everything is built into build\\zstd_test (bin\\ and src\\ are not touched).
"""
import ctypes
import os
import random
import shutil
import struct
import subprocess
import sys
import zipfile
import zlib

ROOT = r"C:\Users\Owner\desktop\websites\warzone"
sys.path.insert(0, os.path.join(ROOT, "tools"))
import build  # noqa: E402

WORK = os.path.join(ROOT, "build", "zstd_test")
BIN = os.path.join(WORK, "bin")
VEC = os.path.join(WORK, "vectors")
ZSTD_DIR = os.path.join(ROOT, "ref", "basis_universal", "zstd")
HIGH_WZ = os.path.join(ROOT, "ref", "official", "wz470", "data", "terrain_overrides", "high.wz")
GCC = r"C:\Strawberry\c\bin\gcc.exe"
MODULES = ["zstd_entropy_c", "zstd_decompress_block_c", "zstd_xxhash_c", "zstd_decompress_c"]

ZSTD_c_compressionLevel, ZSTD_c_windowLog, ZSTD_c_strategy = 100, 101, 107
ZSTD_c_enableLongDistanceMatching, ZSTD_c_contentSizeFlag, ZSTD_c_checksumFlag = 160, 200, 201
ZSTD_c_literalCompressionMode, ZSTD_c_targetCBlockSize = 1002, 1003


def source(stage, rel):
    """rel (a path under src) from the stage when it has it, else from src"""
    if stage and os.path.exists(os.path.join(stage, rel)):
        return os.path.join(stage, rel)
    return os.path.join(build.SRC, rel)


# --- references -----------------------------------------------------------------------------------------------

def build_refs():
    os.makedirs(BIN, exist_ok=True)
    ref_dir = os.path.join(WORK, "ref")
    os.makedirs(ref_dir, exist_ok=True)
    dll64 = os.path.join(ref_dir, "zstdref64.dll")
    if not os.path.exists(dll64):
        r = subprocess.run([GCC, "-O2", "-w", "-shared", "-static-libgcc", "-DZSTD_DLL_EXPORT=1",
                            os.path.join(ZSTD_DIR, "zstd.c"), "-o", dll64], capture_output=True, text=True)
        if r.returncode != 0:
            sys.exit("zstdref64.dll: " + r.stderr[-2000:])
    z32 = os.path.join(BIN, "z32.dll")
    if not os.path.exists(z32):
        zref = os.path.join(ROOT, "tools", "zstdref")
        objs = []
        for name, extra in (("z32", ["-w", "-I", ZSTD_DIR]), ("rt32", [])):
            obj = os.path.join(ref_dir, name + ".o")
            r = subprocess.run([GCC, "-m32", "-O2", "-ffreestanding", "-fno-builtin"] + extra + ["-c",
                                os.path.join(zref, name + ".c"), "-o", obj], capture_output=True, text=True)
            if r.returncode != 0:
                sys.exit("z32.dll: " + r.stderr[-2000:])
            objs.append(obj)
        r = subprocess.run([GCC, "-m32", "-nostdlib", "-shared", "-o", z32] + objs +
                           ["-Wl,-e,_DllMainCRTStartup@12", "-Wl,--kill-at"], capture_output=True, text=True)
        if r.returncode != 0:
            sys.exit("z32.dll: " + r.stderr[-2000:])
    lib = ctypes.CDLL(dll64)
    lib.ZSTD_createCCtx.restype = ctypes.c_void_p
    lib.ZSTD_freeCCtx.argtypes = [ctypes.c_void_p]
    lib.ZSTD_CCtx_reset.argtypes = [ctypes.c_void_p, ctypes.c_int]
    lib.ZSTD_CCtx_reset.restype = ctypes.c_size_t
    lib.ZSTD_CCtx_setParameter.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_int]
    lib.ZSTD_CCtx_setParameter.restype = ctypes.c_size_t
    lib.ZSTD_compress2.argtypes = [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_size_t, ctypes.c_void_p, ctypes.c_size_t]
    lib.ZSTD_compress2.restype = ctypes.c_size_t
    lib.ZSTD_compressBound.argtypes = [ctypes.c_size_t]
    lib.ZSTD_compressBound.restype = ctypes.c_size_t
    lib.ZSTD_decompress.argtypes = [ctypes.c_void_p, ctypes.c_size_t, ctypes.c_void_p, ctypes.c_size_t]
    lib.ZSTD_decompress.restype = ctypes.c_size_t
    lib.ZSTD_isError.argtypes = [ctypes.c_size_t]
    lib.ZSTD_isError.restype = ctypes.c_uint
    lib.ZSTD_getErrorCode.argtypes = [ctypes.c_size_t]
    lib.ZSTD_getErrorCode.restype = ctypes.c_int
    return lib


class Ref:
    def __init__(self, lib):
        self.lib = lib
        self.cctx = lib.ZSTD_createCCtx()

    def compress(self, data, params):
        lib = self.lib
        lib.ZSTD_CCtx_reset(self.cctx, 3)
        for k, v in params.items():
            r = lib.ZSTD_CCtx_setParameter(self.cctx, k, v)
            if lib.ZSTD_isError(r):
                raise RuntimeError("setParameter %d %d: error %d" % (k, v, lib.ZSTD_getErrorCode(r)))
        cap = lib.ZSTD_compressBound(len(data)) + 1024
        out = ctypes.create_string_buffer(cap)
        src = ctypes.create_string_buffer(bytes(data), len(data) + 1)
        r = lib.ZSTD_compress2(self.cctx, out, cap, src, len(data))
        if lib.ZSTD_isError(r):
            raise RuntimeError("compress2: error %d" % lib.ZSTD_getErrorCode(r))
        return out.raw[:r]

    def decompress(self, frame, capacity):
        """(the 64-bit C's result as a VB Long: a size or -errorCode, the output)"""
        lib = self.lib
        out = ctypes.create_string_buffer(max(capacity, 1))
        src = ctypes.create_string_buffer(bytes(frame), len(frame) + 1)
        r = lib.ZSTD_decompress(out, capacity, src, len(frame))
        if lib.ZSTD_isError(r):
            return -lib.ZSTD_getErrorCode(r), b""
        return r, out.raw[:r]


# --- the VB test exes -----------------------------------------------------------------------------------------

def build_exes(stage):
    vbp_dir = os.path.join(WORK, "vbp")
    os.makedirs(vbp_dir, exist_ok=True)
    built = []
    for name, bounds_check in (("ZstdTests", True), ("ZstdTestsFast", False)):
        lines = ["Type=Exe", build.STDOLE]
        paths = [source(stage, os.path.join("zstdtests", "zstdtests_main.bas"))]
        paths += [source(stage, os.path.join("basis", m + ".bas")) for m in MODULES]
        for p in paths:
            build.normalise(p)
            lines.append("Module=%s; %s" % (build.vb_name(p), p))
        lines.append(build.VBP_TAIL.format(name=name, ext="exe", path=BIN, startup="Sub Main", compat="",
                                           bounds="0" if bounds_check else "-1", optimisation="0", threading=""))
        vbp = os.path.join(vbp_dir, name + ".vbp")
        with open(vbp, "w", encoding="latin-1", newline="\r\n") as f:
            f.write("\n".join(lines))
        log = os.path.join(vbp_dir, name + "_build.log")
        if os.path.exists(log):
            os.remove(log)
        exe = os.path.join(BIN, name + ".exe")
        if os.path.exists(exe):
            os.remove(exe)
        subprocess.run([build.VB6, "/make", vbp, "/outdir", BIN, "/out", log], cwd=vbp_dir)
        text = open(log, encoding="latin-1").read().strip() if os.path.exists(log) else ""
        print(text or "%s: VB6 stopped without a message" % name)
        if not os.path.exists(exe):
            sys.exit("%s was not built" % name)
        built.append(exe)
    return built


# --- the frames -------------------------------------------------------------------------------------------------

def inputs(rng, quick):
    """(name, bytes) of many shapes"""
    out = [("empty", b""), ("one", b"A"), ("two", b"AB"), ("zero100k", bytes(100000))]
    text = open(os.path.join(ZSTD_DIR, "zstd.h"), "rb").read()
    out.append(("text_small", text[:3000]))
    out.append(("text", text))
    out.append(("csource", open(os.path.join(ZSTD_DIR, "zstd.c"), "rb").read()))
    out.append(("dll", open(os.path.join(ROOT, "build", "zstd_test", "ref", "zstdref64.dll"), "rb").read()[:600000]))
    for n in (1, 7, 100, 5000, 70000, 200000):
        out.append(("random%d" % n, bytes(rng.getrandbits(8) for _ in range(n))))
    runs = bytearray()
    while len(runs) < 300000:
        runs += bytes([rng.randrange(4)]) * rng.randrange(1, 300)
    out.append(("runs", bytes(runs)))
    mix = bytearray()
    while len(mix) < 500000:
        k = rng.randrange(3)
        if k == 0:
            mix += bytes(rng.getrandbits(8) for _ in range(rng.randrange(1, 64)))
        elif k == 1 and len(mix) > 100:
            off = rng.randrange(1, min(len(mix), 70000))
            ln = rng.randrange(3, 200)
            for _ in range(ln):
                mix.append(mix[-off])
        else:
            mix += text[rng.randrange(len(text) - 200):][:rng.randrange(1, 200)]
    out.append(("mixture", bytes(mix)))
    # small alphabets (Huffman weights, X1 / X2 choices)
    for alpha in (2, 3, 17, 60):
        out.append(("alpha%d" % alpha, bytes(rng.randrange(alpha) * 3 for _ in range(150000))))
    skew = bytes(min(255, int(rng.expovariate(0.15))) for _ in range(200000))
    out.append(("skewed", skew))
    if not quick:
        big = bytearray()
        while len(big) < 6000000:
            big += text[rng.randrange(len(text) - 5000):][:rng.randrange(100, 5000)]
            big += bytes(rng.getrandbits(8) for _ in range(rng.randrange(0, 300)))
        out.append(("big6m", bytes(big)))
    return out


def long_input(rng):
    """80 MB: 40 MB of random-ish data twice (offsets of 40 MB: the 32-bit long-offset path)"""
    a = bytearray(os.urandom(40 * 1024 * 1024))
    return bytes(a + a)


def param_sets(quick):
    sets = [
        ("l1", {ZSTD_c_compressionLevel: 1}),
        ("l3", {ZSTD_c_compressionLevel: 3}),
        ("l9", {ZSTD_c_compressionLevel: 9}),
        ("l19", {ZSTD_c_compressionLevel: 19}),
        ("l3ck", {ZSTD_c_compressionLevel: 3, ZSTD_c_checksumFlag: 1}),
        ("l3nocs", {ZSTD_c_compressionLevel: 3, ZSTD_c_contentSizeFlag: 0}),
        ("l5ck_nocs", {ZSTD_c_compressionLevel: 5, ZSTD_c_checksumFlag: 1, ZSTD_c_contentSizeFlag: 0}),
        ("rawlit", {ZSTD_c_compressionLevel: 3, ZSTD_c_literalCompressionMode: 2}),
        ("hufflit", {ZSTD_c_compressionLevel: 1, ZSTD_c_literalCompressionMode: 1}),
        ("smallblk", {ZSTD_c_compressionLevel: 6, ZSTD_c_targetCBlockSize: 1340}),
        ("neg5", {ZSTD_c_compressionLevel: -5}),
    ]
    if not quick:
        sets += [
            ("l22", {ZSTD_c_compressionLevel: 22}),
            ("l15w17", {ZSTD_c_compressionLevel: 15, ZSTD_c_windowLog: 17, ZSTD_c_contentSizeFlag: 0}),
            ("l12w10", {ZSTD_c_compressionLevel: 12, ZSTD_c_windowLog: 10, ZSTD_c_checksumFlag: 1}),
            ("neg1", {ZSTD_c_compressionLevel: -1, ZSTD_c_checksumFlag: 1}),
            ("strat2", {ZSTD_c_compressionLevel: 4, ZSTD_c_strategy: 2}),
            ("strat5", {ZSTD_c_compressionLevel: 9, ZSTD_c_strategy: 5}),
            ("strat9", {ZSTD_c_compressionLevel: 19, ZSTD_c_strategy: 9}),
        ]
    return sets


def skippable(payload, magic_low=0):
    return struct.pack("<II", 0x184D2A50 | magic_low, len(payload)) + payload


def craft_window(frame, wlog):
    """the frame with its window descriptor set to 2^wlog (a frame that has one: not single-segment)"""
    fhd = frame[4]
    if (fhd >> 5) & 1:
        return None
    b = bytearray(frame)
    b[5] = ((wlog - 10) << 3)
    return bytes(b)


def make_vectors(ref, rng, quick):
    """[(name, frame bytes, dstCapacity, expected (result, hash) or None for an expected error, kind)]"""
    vec = []
    ins = inputs(rng, quick)
    by_name = dict(ins)
    for iname, data in ins:
        for pname, params in param_sets(quick):
            if iname in ("big6m", "csource") and pname in ("l22", "l19", "strat9"):
                continue
            frame = ref.compress(data, params)
            vec.append(("%s_%s" % (iname, pname), frame, len(data), "valid"))
    # several frames in a row, with skippable frames
    a = ref.compress(by_name["text_small"], {ZSTD_c_compressionLevel: 3})
    b = ref.compress(by_name["mixture"], {ZSTD_c_compressionLevel: 5, ZSTD_c_checksumFlag: 1})
    c = ref.compress(b"", {ZSTD_c_compressionLevel: 1})
    total = len(by_name["text_small"]) + len(by_name["mixture"])
    vec.append(("multi_ab", a + b, total, "valid"))
    vec.append(("multi_skip", skippable(b"hello") + a + skippable(b"", 15) + b + c, total, "valid"))
    vec.append(("only_skip", skippable(bytes(100), 3), 0, "valid"))
    vec.append(("multi_trailing_garbage", a + b"\x01\x02\x03\x04\x05\x06\x07\x08\x09", total, "error"))
    vec.append(("multi_trailing_4", a + b"\x28\xb5\x2f", total, "error"))
    vec.append(("skip_truncated", skippable(bytes(100))[:50], 0, "error"))
    vec.append(("bad_magic", b"\x00\x01\x02\x03" + a[4:], 3000, "error"))
    # crafted headers
    nocs = ref.compress(by_name["mixture"], {ZSTD_c_compressionLevel: 3, ZSTD_c_contentSizeFlag: 0})
    for wlog in (25, 26, 30, 31):
        f = craft_window(nocs, wlog)
        if f:
            vec.append(("window%d" % wlog, f, len(by_name["mixture"]), "valid" if wlog <= 30 else "error"))
    r = bytearray(a)
    r[4] |= 0x08
    vec.append(("reserved_bit", bytes(r), 3000, "error"))
    d = bytearray(a)          # dictID: a 1-byte field (after the window descriptor, which a single segment lacks)
    if (d[4] & 3) == 0:
        pos = 5 if (d[4] >> 5) & 1 else 6
        d[4] |= 1
        vec.append(("dictid", bytes(d[:pos] + bytes([7]) + d[pos:]), 3000, "error"))
    # too small destinations, zero capacity
    for iname in ("text", "runs", "mixture", "zero100k"):
        f = ref.compress(by_name[iname], {ZSTD_c_compressionLevel: 3})
        n = len(by_name[iname])
        for cap in (0, 1, n // 2, n - 1):
            vec.append(("%s_cap%d" % (iname, cap), f, cap, "error"))
        vec.append(("%s_capplus" % iname, f, n + 1000, "valid"))
    # broken frames: cut short and flipped
    bases = [("text", ref.compress(by_name["text"], {ZSTD_c_compressionLevel: 3})),
             ("mixture", ref.compress(by_name["mixture"], {ZSTD_c_compressionLevel: 9, ZSTD_c_checksumFlag: 1})),
             ("alpha3", ref.compress(by_name["alpha3"], {ZSTD_c_compressionLevel: 19})),
             ("skewed", ref.compress(by_name["skewed"], {ZSTD_c_compressionLevel: 5})),
             ("random5000", ref.compress(by_name["random5000"], {ZSTD_c_compressionLevel: 1})),
             ("smallblk", ref.compress(by_name["mixture"], {ZSTD_c_compressionLevel: 6, ZSTD_c_targetCBlockSize: 1340}))]
    sizes = {"text": len(by_name["text"]), "mixture": len(by_name["mixture"]), "alpha3": len(by_name["alpha3"]),
             "skewed": len(by_name["skewed"]), "random5000": 5000, "smallblk": len(by_name["mixture"])}
    flips = 40 if quick else 150
    for bname, f in bases:
        for k in range(12):
            cut = rng.randrange(1, len(f))
            vec.append(("%s_cut%d" % (bname, cut), f[:cut], sizes[bname], "broken"))
        for k in range(flips):
            g = bytearray(f)
            for _ in range(rng.choice((1, 1, 2, 4))):
                pos = rng.randrange(len(g))
                if rng.random() < 0.5:
                    g[pos] ^= 1 << rng.randrange(8)
                else:
                    g[pos] = rng.getrandbits(8)
            vec.append(("%s_flip%d" % (bname, k), bytes(g), sizes[bname], "broken"))
    # literals that are all one byte (set_rle literals): a random vocabulary, then copies of it with 'A' between
    vocab = bytes(rng.getrandbits(8) for _ in range(131072))
    litrle = bytearray(vocab)
    while len(litrle) < 1000000:
        p = rng.randrange(len(vocab) - 300)
        litrle += vocab[p:p + rng.randrange(20, 300)] + b"A"
    vec.append(("litrle_l3", ref.compress(bytes(litrle), {ZSTD_c_compressionLevel: 3}), len(litrle), "valid"))
    vec.append(("litrle_l19", ref.compress(bytes(litrle), {ZSTD_c_compressionLevel: 19}), len(litrle), "valid"))
    if not quick:
        # fresh offsets of 20 to 70 MB in every sequence: the prefetching sequence decoder (window > 16 MB, long
        # offsets in more than 20/256 of the OF table) and the 32-bit long-offset reads (window > 32 MB)
        for total_mb, a_mb, wlog in ((44, 36, 26), (72, 68, 27)):
            a = bytearray(os.urandom(a_mb * 1024 * 1024))
            scatter = bytearray()
            while len(a) + len(scatter) < total_mb * 1024 * 1024:
                p = rng.randrange(0, a_mb * 1024 * 1024 // 2)
                scatter += a[p:p + rng.randrange(100, 2000)]
                scatter += bytes(rng.getrandbits(8) for _ in range(rng.randrange(1, 8)))
            data = bytes(a + scatter)
            vec.append(("scatter%dm" % total_mb, ref.compress(data, {ZSTD_c_compressionLevel: 3, ZSTD_c_windowLog: wlog,
                                                                      ZSTD_c_enableLongDistanceMatching: 1}), len(data), "valid"))
        big = long_input(rng)
        for pname, params in (("w27ldm", {ZSTD_c_compressionLevel: 3, ZSTD_c_windowLog: 27,
                                          ZSTD_c_enableLongDistanceMatching: 1, ZSTD_c_checksumFlag: 1}),
                              ("w26ldm_nocs", {ZSTD_c_compressionLevel: 1, ZSTD_c_windowLog: 26,
                                               ZSTD_c_enableLongDistanceMatching: 1, ZSTD_c_contentSizeFlag: 0})):
            vec.append(("long80m_" + pname, ref.compress(big, params), len(big), "valid"))
    return vec


def ktx2_frames():
    """[(name, file, offset, length, uncompressed length)] for every level of high.wz's KTX2 files"""
    out = []
    folder = os.path.join(WORK, "ktx2")
    with zipfile.ZipFile(HIGH_WZ) as z:
        for n in z.namelist():
            if not n.endswith(".ktx2"):
                continue
            path = os.path.join(folder, n.replace("/", "_"))
            if not os.path.exists(path):
                os.makedirs(folder, exist_ok=True)
                with open(path, "wb") as f:
                    f.write(z.read(n))
            with open(path, "rb") as f:
                d = f.read()
            levels, scheme = struct.unpack_from("<II", d, 40)
            if scheme != 2:
                continue
            for lv in range(levels):
                off, ln, ul = struct.unpack_from("<QQQ", d, 80 + 24 * lv)
                out.append(("%s#%d" % (n, lv), path, off, ln, ul))
    return out


def run_exe(exe, listfile, outdir):
    os.makedirs(outdir, exist_ok=True)
    r = subprocess.run([exe, listfile, outdir], cwd=BIN, timeout=7200)
    res = {}
    total = None
    with open(os.path.join(outdir, "results.txt")) as f:
        for line in f:
            p = line.rstrip("\n").split("\t")
            if p[0] == "total":
                total = p
            else:
                res[p[0]] = p
    return r.returncode, res, total


def main(argv):
    stage = None
    if "--stage" in argv:
        stage = os.path.abspath(argv[argv.index("--stage") + 1])
    keep, quick, no_ktx2 = "--keep" in argv, "--quick" in argv, "--no-ktx2" in argv
    only = argv[argv.index("--only") + 1].split(",") if "--only" in argv else None
    rng = random.Random(20261101)
    lib = build_refs()
    ref = Ref(lib)
    exes = build_exes(stage)

    if os.path.exists(VEC) and not keep:
        shutil.rmtree(VEC)
    os.makedirs(VEC, exist_ok=True)
    print("making frames ...")
    vec = make_vectors(ref, rng, quick)
    if only:
        vec = [v for v in vec if any(v[0].startswith(o) for o in only)]
    lines, expect = [], {}
    for name, frame, cap, kind in vec:
        path = os.path.join(VEC, name + ".zst")
        with open(path, "wb") as f:
            f.write(frame)
        r, out = ref.decompress(frame, cap)
        expect[name] = (kind, r, zlib.crc32(out) & 0xFFFFFFFF if r > 0 else None)
        lines.append("%s\t%s\t0\t%d\t%d" % (name, path, len(frame), cap))
    gen_list = os.path.join(WORK, "frames.txt")
    with open(gen_list, "w", newline="\n") as f:
        f.write("\n".join(lines) + "\n")

    ktx_list = None
    if not no_ktx2:
        print("KTX2 frames of high.wz ...")
        kf = ktx2_frames()
        klines = []
        for name, path, off, ln, ul in kf:
            klines.append("%s\t%s\t%d\t%d\t%d" % (name, path, off, ln, ul))
            expect[name] = ("ktx2", None, None)
        ktx_list = os.path.join(WORK, "ktx2.txt")
        with open(ktx_list, "w", newline="\n") as f:
            f.write("\n".join(klines) + "\n")
        # the 64-bit C's output CRCs of the KTX2 levels
        for i, (name, path, off, ln, ul) in enumerate(kf):
            with open(path, "rb") as f:
                d = f.read()
            r, out = ref.decompress(d[off:off + ln], ul)
            expect[name] = ("ktx2", r, zlib.crc32(out) & 0xFFFFFFFF)

    failed = 0
    for exe in exes:
        ename = os.path.splitext(os.path.basename(exe))[0]
        for lname, lfile in (("frames", gen_list), ("ktx2", ktx_list)):
            if lfile is None:
                continue
            rc, res, total = run_exe(exe, lfile, os.path.join(WORK, "out_" + ename + "_" + lname))
            bad = []
            notes = []
            for name, p in res.items():
                kind, r64, h64 = expect[name]
                vb, c32, same, guards = int(p[1]), int(p[2]), p[3] == "1", p[4] == "1"
                if not same:
                    bad.append("%s: VB %d, C (32-bit) %d" % (name, vb, c32))
                    continue
                if not guards:
                    bad.append("%s: wrote outside dstCapacity" % name)
                    continue
                if kind in ("valid", "ktx2") and r64 is not None:
                    if vb != r64:
                        bad.append("%s: VB %d, C (64-bit) %d" % (name, vb, r64))
                    elif vb > 0 and h64 is not None and int(p[5], 16) != h64:
                        bad.append("%s: output differs from the 64-bit C's" % name)
                if kind == "error" and vb >= 0:
                    bad.append("%s: expected an error, got %d" % (name, vb))
                if kind in ("error", "broken") and r64 is not None and vb != r64:
                    notes.append("%s: VB = C 32-bit %d, C 64-bit %d" % (name, vb, r64))
            if len(res) != sum(1 for _ in open(lfile)):
                bad.append("results for %d of %d frames (exit %d)" % (len(res), sum(1 for _ in open(lfile)), rc))
            mbs = float(total[3]) / 1e6 / (float(total[4]) / 1000) if total and float(total[4]) > 0 else 0
            cmbs = float(total[3]) / 1e6 / (float(total[5]) / 1000) if total and float(total[5]) > 0 else 0
            nerr = sum(1 for p in res.values() if int(p[1]) < 0)
            print("%s %s: %d frames (%d errors as C's), %d problems; %.1f MB out: VB %.0f ms (%.1f MB/s), C 32-bit "
                  "%.0f ms (%.1f MB/s)" % (ename, lname, len(res), nerr, len(bad), float(total[3]) / 1e6 if total else 0,
                                           float(total[4]) if total else 0, mbs, float(total[5]) if total else 0, cmbs))
            for b in bad[:30]:
                print("   ", b)
            if notes and ename == "ZstdTestsFast":
                print("   errors where the 64-bit build differs from the 32-bit (%d), e.g.:" % len(notes))
                for n in notes[:8]:
                    print("     ", n)
            failed += len(bad)
    print("zstd_test: %s" % ("all passed" if failed == 0 else "%d problems" % failed))
    return 1 if failed else 0


def run(keep=False):
    """run_tests.py's entry (its MEDIA_TESTS): the whole test on src; True when everything passed"""
    return main(["--keep"] if keep else []) == 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
