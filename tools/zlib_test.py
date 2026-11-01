"""
zlib_test.py - cluster J's test of the zlib streams (src/shared/miniz: miniz.c v1.11's deflate, tdefl, and inflate,
tinfl, with zlib's stream API), compiled into a test exe (ZlibTests: src/zlibtests + the miniz modules + math_vb;
built as WzNet is: no optimisation, bounds checks):

* the port's deflate against miniz.c v1.11 itself, byte for byte: the C (ref/miniz_v111/miniz.c) compiled with gcc
  into two DLLs, as it is (x64: MINIZ_HAS_64BIT_REGISTERS) and with the x86 32-bit build's code paths
  (MINIZ_HAS_64BIT_REGISTERS 0, the port's), driven through ctypes with the same calls; every level (0 to 10),
  strategy (filtered, Huffman only, RLE, fixed), raw deflate, pieces and small output room, the flushes (partial,
  sync, full) and zlib_compression_adapter.cpp's pattern (each network message compressed, then a
  Z_PARTIAL_FLUSH); the output lengths after each piece too;
* what the port deflates is read back by Python's zlib and by the official 4.7.0 build's libzlib1.dll (ctypes): the
  whole stream, and after each flush exactly the input so far (what the network's peer needs);
* the port's inflate reads zlib's streams (Python's: every level, strategy, memLevel, window size; the official
  libzlib1.dll's netplay pattern at level 6 with Z_PARTIAL_FLUSH, fed segment by segment and byte by byte) and its
  own; broken streams fail with Z_DATA_ERROR as zlib's do;
* mz_compress2 one-shot; the speed.

  python -I tools\\zlib_test.py [--keep] [--stage DIR] [--zip] [--only deflate,inflate,receive]
                                                       (run_tests.py may call run(keep))

--only: just those groups of runs (deflate: the port's deflate against miniz.c and zlib; inflate: zlib's streams;
receive: the network's receive pattern).

--stage DIR: a src\\ mirror (cluster J's staging): its shared\\miniz and zlibtests folders (and framework's
physfs_zip_c, for --zip) are used where they exist, and the test exes build into build\\zlib_test\\bin, not bin\\.
--zip: also runs tools/zip_test.py's checks with a ZipTests built from these miniz modules (PhysFS's zip
archiver reads through miniz_tinfl_c).
Without gcc (Strawberry Perl's, or any on PATH) the byte-for-byte comparison with the C is skipped (said so).
"""
import ctypes
import os
import random
import shutil
import subprocess
import sys
import time
import zlib

ROOT = r"C:\Users\Owner\desktop\websites\warzone"
sys.path.insert(0, os.path.join(ROOT, "tools"))
import build  # noqa: E402

WORK = os.path.join(ROOT, "build", "zlib_test")
MINIZ_C = os.path.join(ROOT, "ref", "miniz_v111", "miniz.c")
OFFICIAL_ZLIB = os.path.join(ROOT, "ref", "official", "wz470", "bin", "libzlib1.dll")
MINIZ_MODULES = ["miniz_h", "miniz_tinfl_c", "miniz_tdefl_c", "miniz_c"]

Z_NO_FLUSH, Z_PARTIAL_FLUSH, Z_SYNC_FLUSH, Z_FULL_FLUSH, Z_FINISH = 0, 1, 2, 3, 4
Z_OK, Z_STREAM_END, Z_BUF_ERROR, Z_STREAM_ERROR, Z_DATA_ERROR = 0, 1, -5, -2, -3
NETPLAY = 9


def source(stage, rel):
    """rel (a path under src) from the stage when it has it, else from src"""
    if stage and os.path.exists(os.path.join(stage, rel)):
        return os.path.join(stage, rel)
    return os.path.join(build.SRC, rel)


# ---------------------------------------------------------------------------------------------------------
# The test exe

def build_exe(stage, out_bin):
    files = [source(stage, os.path.join("shared", "miniz", m + ".bas")) for m in MINIZ_MODULES]
    tests = os.path.dirname(source(stage, os.path.join("zlibtests", "zlibtests_main.bas")))
    files += [os.path.join(tests, n) for n in build.sources(tests)]
    files.append(os.path.join(build.SHARED, "math_vb.bas"))
    vbp_dir = os.path.join(WORK, "vbp")
    os.makedirs(vbp_dir, exist_ok=True)
    os.makedirs(out_bin, exist_ok=True)
    lines = ["Type=Exe", build.STDOLE]
    for p in files:
        build.normalise(p)
        lines.append("Module=%s; %s" % (build.vb_name(p), p))
    lines.append(build.VBP_TAIL.format(name="ZlibTests", ext="exe", path=out_bin, startup="Sub Main", compat="",
                                       bounds="0", optimisation="2", threading=""))
    vbp = os.path.join(vbp_dir, "ZlibTests.vbp")
    with open(vbp, "w", encoding="latin-1", newline="\r\n") as f:
        f.write("\n".join(lines))
    saved = build.BIN
    build.BIN = out_bin
    try:
        return build.make("ZlibTests", vbp_dir, vbp)
    finally:
        build.BIN = saved


def run_exe(exe, cmds):
    cmdfile = os.path.join(WORK, "commands.txt")
    resfile = os.path.join(WORK, "results.txt")
    with open(cmdfile, "w", newline="\r\n") as f:
        f.write("\n".join("\t".join(str(x) for x in c) for c in cmds) + "\n")
    if os.path.exists(resfile):
        os.remove(resfile)
    t0 = time.time()
    r = subprocess.run([exe, cmdfile, resfile], cwd=WORK, timeout=7200)
    dt = time.time() - t0
    if r.returncode != 0 or not os.path.exists(resfile):
        raise RuntimeError("ZlibTests.exe failed (exit %d)" % r.returncode)
    out = [ln.split("\t") for ln in open(resfile).read().splitlines() if ln]
    return out, dt


# ---------------------------------------------------------------------------------------------------------
# The C: miniz.c v1.11 (gcc) and the official libzlib1.dll, through ctypes

class z_stream(ctypes.Structure):
    """zlib's and miniz's stream (x64 Windows: uLong is 32 bits)"""
    _fields_ = [("next_in", ctypes.c_void_p), ("avail_in", ctypes.c_uint), ("total_in", ctypes.c_ulong),
                ("next_out", ctypes.c_void_p), ("avail_out", ctypes.c_uint), ("total_out", ctypes.c_ulong),
                ("msg", ctypes.c_char_p), ("state", ctypes.c_void_p), ("zalloc", ctypes.c_void_p),
                ("zfree", ctypes.c_void_p), ("opaque", ctypes.c_void_p), ("data_type", ctypes.c_int),
                ("adler", ctypes.c_ulong), ("reserved", ctypes.c_ulong)]


def find_gcc():
    for c in ("gcc", r"C:\Strawberry\c\bin\gcc.exe"):
        p = shutil.which(c) if not os.path.isabs(c) else (c if os.path.exists(c) else None)
        if p:
            return p
    return None


def build_miniz_refs():
    """the two miniz DLLs: (x64 as it is, the x86 32-bit code paths), or None without gcc"""
    gcc = find_gcc()
    if gcc is None:
        return None
    ref = os.path.join(WORK, "miniz_ref")
    os.makedirs(ref, exist_ok=True)
    text = open(MINIZ_C, encoding="latin-1").read()
    # (MinGW's own __forceinline is "extern __inline__ ...", which "static" contradicts: a plain inline hint here)
    assert "static __forceinline" in text
    text = text.replace("static __forceinline", "static inline")
    variants = {"miniz_x64": text}
    old = "\n#define MINIZ_HAS_64BIT_REGISTERS 1"           # (the line, not the header comment's example)
    assert text.count(old) == 1
    variants["miniz_x86paths"] = text.replace(old, "\n#define MINIZ_HAS_64BIT_REGISTERS 0")
    dlls = []
    for name, src in variants.items():
        c = os.path.join(ref, name + ".c")
        dll = os.path.join(ref, name + ".dll")
        with open(c, "w", encoding="latin-1") as f:
            f.write(src)
        if not os.path.exists(dll) or os.path.getmtime(dll) < os.path.getmtime(c):
            r = subprocess.run([gcc, "-O2", "-shared", "-DMINIZ_NO_ARCHIVE_APIS", "-DMINIZ_NO_STDIO", "-DMINIZ_NO_TIME",
                                "-DMINIZ_NO_ZLIB_COMPATIBLE_NAMES", "-o", dll, c], capture_output=True, text=True)
            if r.returncode != 0:
                print("zlib: gcc failed on %s: %s" % (name, r.stderr[:500]))
                return None
        dlls.append(dll)
    return [MinizLib(d) for d in dlls]


class MinizLib:
    def __init__(self, path):
        self.lib = ctypes.CDLL(path)
        self.name = os.path.basename(path)
        L = self.lib
        L.mz_deflateInit2.argtypes = [ctypes.POINTER(z_stream)] + [ctypes.c_int] * 5
        L.mz_deflate.argtypes = [ctypes.POINTER(z_stream), ctypes.c_int]
        L.mz_deflateEnd.argtypes = [ctypes.POINTER(z_stream)]
        L.mz_inflateInit2.argtypes = [ctypes.POINTER(z_stream), ctypes.c_int]
        L.mz_inflate.argtypes = [ctypes.POINTER(z_stream), ctypes.c_int]
        L.mz_inflateEnd.argtypes = [ctypes.POINTER(z_stream)]

    def deflateInit2(self, s, level, wbits, strategy):
        return self.lib.mz_deflateInit2(ctypes.byref(s), level, 8, wbits, 9, strategy)

    def deflate(self, s, flush):
        return self.lib.mz_deflate(ctypes.byref(s), flush)

    def deflateEnd(self, s):
        return self.lib.mz_deflateEnd(ctypes.byref(s))

    def inflateInit2(self, s, wbits):
        return self.lib.mz_inflateInit2(ctypes.byref(s), wbits)

    def inflate(self, s, flush):
        return self.lib.mz_inflate(ctypes.byref(s), flush)

    def inflateEnd(self, s):
        return self.lib.mz_inflateEnd(ctypes.byref(s))


class ZlibLib:
    """the official 4.7.0 build's libzlib1.dll"""
    def __init__(self, path):
        if hasattr(os, "add_dll_directory"):
            self._dir = os.add_dll_directory(os.path.dirname(path))
        self.lib = ctypes.CDLL(path)
        self.name = "libzlib1.dll"
        L = self.lib
        L.zlibVersion.restype = ctypes.c_char_p
        self.version = L.zlibVersion()
        L.deflateInit2_.argtypes = [ctypes.POINTER(z_stream)] + [ctypes.c_int] * 5 + [ctypes.c_char_p, ctypes.c_int]
        L.deflate.argtypes = [ctypes.POINTER(z_stream), ctypes.c_int]
        L.deflateEnd.argtypes = [ctypes.POINTER(z_stream)]
        L.inflateInit2_.argtypes = [ctypes.POINTER(z_stream), ctypes.c_int, ctypes.c_char_p, ctypes.c_int]
        L.inflate.argtypes = [ctypes.POINTER(z_stream), ctypes.c_int]
        L.inflateEnd.argtypes = [ctypes.POINTER(z_stream)]

    def deflateInit2(self, s, level, wbits, strategy):
        return self.lib.deflateInit2_(ctypes.byref(s), level, 8, wbits, 9, strategy, self.version, ctypes.sizeof(z_stream))

    def deflate(self, s, flush):
        return self.lib.deflate(ctypes.byref(s), flush)

    def deflateEnd(self, s):
        return self.lib.deflateEnd(ctypes.byref(s))

    def inflateInit2(self, s, wbits):
        return self.lib.inflateInit2_(ctypes.byref(s), wbits, self.version, ctypes.sizeof(z_stream))

    def inflate(self, s, flush):
        return self.lib.inflate(ctypes.byref(s), flush)

    def inflateEnd(self, s):
        return self.lib.inflateEnd(ctypes.byref(s))


def c_deflate(L, data, level, strategy, wbits, sizes, chunk, flush_mode):
    """the C's run of ZlibTests' def command, call for call: (output, marks)"""
    s = z_stream()
    if L.deflateInit2(s, level, wbits, strategy) != Z_OK:
        raise RuntimeError("deflateInit2")
    inbuf = ctypes.create_string_buffer(bytes(data), max(len(data), 1))
    base = ctypes.addressof(inbuf)
    out = bytearray()
    marks = []
    pos = 0

    def call(room, flush, with_input):
        ob = ctypes.create_string_buffer(max(room, 1))
        s.next_out = ctypes.addressof(ob)
        s.avail_out = room
        if not with_input:
            s.next_in = None
            s.avail_in = 0
        st = L.deflate(s, flush)
        out.extend(ob.raw[:room - s.avail_out])
        return st

    for n in sizes:
        if flush_mode == NETPLAY:
            s.next_in = base + pos
            s.avail_in = n
            while True:
                st = call(n + 20, Z_NO_FLUSH, True)
                if s.avail_out != 0:
                    break
            while True:
                st = call(1000, Z_PARTIAL_FLUSH, False)
                if s.avail_out != 0:
                    break
        else:
            s.next_in = base + pos
            s.avail_in = n
            while True:
                st = call(chunk, flush_mode, True)
                if st < 0 and st != Z_BUF_ERROR:
                    raise RuntimeError("deflate %d" % st)
                if s.avail_in == 0 and (flush_mode == Z_NO_FLUSH or s.avail_out != 0):
                    break
        pos += n
        marks.append(len(out))
    if not (flush_mode == NETPLAY and chunk == 0):
        room = chunk if chunk > 0 else 4096
        while True:
            st = call(room, Z_FINISH, False)
            if st == Z_STREAM_END:
                break
            if st < 0 and st != Z_BUF_ERROR:
                raise RuntimeError("finish %d" % st)
    L.deflateEnd(s)
    return bytes(out), marks


def c_inflate(L, data, wbits, sizes, chunk, flush=Z_NO_FLUSH):
    """ZlibTests' inf command through the C: (output, marks, last status)"""
    s = z_stream()
    if L.inflateInit2(s, wbits) != Z_OK:
        raise RuntimeError("inflateInit2")
    inbuf = ctypes.create_string_buffer(bytes(data), max(len(data), 1))
    base = ctypes.addressof(inbuf)
    out = bytearray()
    marks = []
    pos = 0
    st = Z_OK
    for n in sizes:
        s.next_in = base + pos
        s.avail_in = n
        while True:
            ob = ctypes.create_string_buffer(chunk)
            s.next_out = ctypes.addressof(ob)
            s.avail_out = chunk
            st = L.inflate(s, flush)
            out.extend(ob.raw[:chunk - s.avail_out])
            if st < 0 or st == Z_STREAM_END:
                break
            if s.avail_in == 0 and s.avail_out != 0:
                break
        pos += n
        marks.append(len(out))
        if (st < 0 and st != Z_BUF_ERROR) or st == Z_STREAM_END:
            break
    L.inflateEnd(s)
    return bytes(out), marks, st


def c_adapter(L, data, max_size, received, segs):
    """ZlibTests' adapt command through the C: the network's receive (readNoInt over ZlibCompressionAdapter):
    (output, calls [(produced, avail_in, avail_out, status)], calls with room and input left, last status)"""
    s = z_stream()
    if L.inflateInit2(s, 15) != Z_OK:
        raise RuntimeError("inflateInit")
    inbuf = ctypes.create_string_buffer(max_size + 1000)
    ob = ctypes.create_string_buffer(max_size)
    out = bytearray()
    calls = []
    violations = 0
    need = True
    recv = 0
    seg_i = 0
    seg_left = 0
    st = Z_OK
    while True:
        if need:
            if recv >= len(data):
                break
            if received == 1:
                if seg_left == 0:
                    seg_left = segs[seg_i]
                    seg_i += 1
                size = min(seg_left, max_size + 1000)
                seg_left -= size
            else:
                size = min(len(data) - recv, max_size + 1000)
            ctypes.memmove(inbuf, data[recv:recv + size], size)
            recv += size
            s.next_in = ctypes.addressof(inbuf)
            s.avail_in = size
            need = False
        s.next_out = ctypes.addressof(ob)
        s.avail_out = max_size
        st = L.inflate(s, Z_NO_FLUSH)
        if st == Z_STREAM_ERROR:
            break
        produced = max_size - s.avail_out
        out.extend(ob.raw[:produced])
        calls.append((produced, s.avail_in, s.avail_out, st))
        if s.avail_out != 0:
            need = True
            if s.avail_in != 0:
                violations += 1
        if st in (Z_STREAM_END, Z_DATA_ERROR):
            break
    L.inflateEnd(s)
    return bytes(out), calls, violations, st


# ---------------------------------------------------------------------------------------------------------
# The data

def corpus(work):
    """{name: bytes}: the game's text and binary data, and made-up shapes"""
    rnd = random.Random(1234)
    data = os.path.join(ROOT, "data")
    out = {"empty": b"", "one": b"A", "tiny": b"hello hello hello world"}
    stats = os.path.join(data, "mp", "stats")
    txt = b""
    for n in sorted(os.listdir(stats)):
        if n.endswith(".json"):
            txt += open(os.path.join(stats, n), "rb").read()
    out["stats_json"] = txt[:600000]
    js = b""
    camp = os.path.join(data, "base", "script", "campaign")
    for n in sorted(os.listdir(camp))[:40]:
        js += open(os.path.join(camp, n), "rb").read()
    out["scripts_js"] = js[:400000]
    maps = os.path.join(data, "mp", "multiplay", "maps")
    for root, _dirs, files in os.walk(maps):
        for n in files:
            if n == "game.map" and "map" not in out:
                out["map"] = open(os.path.join(root, n), "rb").read()
    for root, _dirs, files in os.walk(os.path.join(data, "base", "texpages")):
        for n in sorted(files):
            if n.endswith(".png") and os.path.getsize(os.path.join(root, n)) > 50000 and "png" not in out:
                out["png"] = open(os.path.join(root, n), "rb").read()[:200000]
    out["noise"] = bytes(rnd.getrandbits(8) for _ in range(150000))
    out["zeros"] = bytes(300000)
    block = bytes(rnd.getrandbits(8) for _ in range(20000))
    far = bytearray()
    for i in range(12):                       # repeats 9 to 30 KB back, with a few changes: long distances
        b = bytearray(block)
        for _ in range(30):
            b[rnd.randrange(len(b))] = rnd.getrandbits(8)
        far += b + bytes(rnd.getrandbits(8) for _ in range(rnd.randrange(1000, 9000)))
    out["far_repeats"] = bytes(far)
    tri = bytearray()                         # 3-byte matches far apart (the 8 KB rule) and short ones near
    for i in range(60000):
        tri += bytes([97 + rnd.randrange(3)])
    out["trigrams"] = bytes(tri)
    runs = bytearray()                        # runs of all lengths: RLE matches, 258-byte matches
    for i in range(2000):
        runs += bytes([rnd.getrandbits(8)]) * rnd.randrange(1, 600)
    out["runs"] = bytes(runs)
    return out


def replay_messages():
    """a real network stream cut into message-sized pieces: (bytes, sizes)"""
    rnd = random.Random(99)
    reps = os.path.join(ROOT, "ref", "official", "replays")
    blob = b""
    if os.path.isdir(reps):
        for n in sorted(os.listdir(reps)):
            if n.endswith(".wzrp"):
                blob += open(os.path.join(reps, n), "rb").read()
    if not blob:
        blob = bytes(rnd.getrandbits(8) for _ in range(100000))
    blob = blob[:300000]
    sizes = []
    left = len(blob)
    while left > 0:
        n = min(left, rnd.choice([1, 2, 5, 9, 17, 30, 60, 120, 250, 400, 1400]))
        sizes.append(n)
        left -= n
    return blob, sizes


# ---------------------------------------------------------------------------------------------------------

def run(keep=False, stage=None, with_zip=False, only=None):
    groups = set(only) if only else {"deflate", "inflate", "receive"}
    os.makedirs(WORK, exist_ok=True)
    out_bin = os.path.join(WORK, "bin") if stage else build.BIN
    if not build_exe(stage, out_bin):
        print("zlib: ZlibTests.exe did not build")
        return False
    exe = os.path.join(out_bin, "ZlibTests.exe")
    refs = build_miniz_refs()
    official = ZlibLib(OFFICIAL_ZLIB) if os.path.exists(OFFICIAL_ZLIB) else None
    files = os.path.join(WORK, "files")
    if os.path.exists(files):
        shutil.rmtree(files)
    os.makedirs(files)
    data = corpus(WORK)
    for name, b in data.items():
        open(os.path.join(files, name), "wb").write(b)
    net, net_sizes = replay_messages()
    open(os.path.join(files, "netplay"), "wb").write(net)
    open(os.path.join(files, "netplay.sizes"), "w").write("\n".join(map(str, net_sizes)) + "\n")

    # the deflate runs: (label, file, level, strategy, wbits, piece, out chunk, flush)
    big = ["stats_json", "scripts_js", "map", "png", "noise", "zeros", "far_repeats", "trigrams", "runs"]
    small = ["empty", "one", "tiny"]
    defs = []
    for name in small:
        for level in (0, 1, 6, 9, 10):
            defs.append(("%s L%d" % (name, level), name, level, 0, 15, 0 if not data[name] else len(data[name]), 64, 0))
    for name in big:
        # (every level on the made-up shapes and the stats; the rest at levels 1 (tdefl_compress_fast) and 6)
        for level in (0, 1, 2, 3, 6, 9) if name in ("stats_json", "far_repeats", "trigrams", "runs") else (1, 6):
            defs.append(("%s L%d" % (name, level), name, level, 0, 15, len(data[name]), 1 << 20, 0))
    for level in (4, 5, 7, 8, 10):
        defs.append(("trigrams L%d" % level, "trigrams", level, 0, 15, len(data["trigrams"]), 1 << 20, 0))
    for strategy, sname in ((1, "filtered"), (2, "huffman only"), (3, "rle"), (4, "fixed")):
        for name in ("stats_json", "map", "runs"):
            defs.append(("%s %s L6" % (name, sname), name, 6, strategy, 15, len(data[name]), 1 << 20, 0))
    for name in ("stats_json", "runs"):
        defs.append(("%s raw L6" % name, name, 6, 0, -15, len(data[name]), 1 << 20, 0))
    for flush, fname in ((0, "none"), (1, "partial"), (2, "sync"), (3, "full")):
        defs.append(("stats_json pieces 1000 room 300 %s" % fname, "stats_json", 6, 0, 15, 1000, 300, flush))
    # (with a flush, a room that divides 6 never ends with miniz v1.11, in C as here: a flush with nothing new is
    # a 6-byte empty block (an empty static block and the sync's stored one), and "call again while the room ran
    # out" starts another each time the room was filled exactly; zlib refuses such a flush with Z_BUF_ERROR)
    defs.append(("tiny pieces 3 room 1", "tiny", 6, 0, 15, 3, 1, 0))
    defs.append(("tiny pieces 3 room 4 sync", "tiny", 6, 0, 15, 3, 4, 2))
    defs.append(("far_repeats pieces 7000 room 50 L9 partial", "far_repeats", 9, 0, 15, 7000, 50, 1))
    defs.append(("netplay adapter L6", "netplay", 6, 0, 15, 0, 0, NETPLAY))
    defs.append(("netplay adapter L1", "netplay", 1, 0, 15, 0, 0, NETPLAY))
    defs.append(("netplay adapter L6 finished", "netplay", 6, 0, 15, 0, 4096, NETPLAY))
    if "deflate" not in groups:
        defs = []

    def sizes_of(name, piece):
        if piece == 0:
            return net_sizes if name == "netplay" else ([len(data[name])] if data[name] else [])
        b = data[name]
        return [min(piece, len(b) - i) for i in range(0, len(b), piece)]

    cmds = []
    for i, (label, name, level, strategy, wbits, piece, chunk, flush) in enumerate(defs):
        src = os.path.join(files, name)
        if piece == 0 and name != "netplay":
            with open(src + ".sizes", "w") as f:
                f.write("".join("%d\n" % n for n in sizes_of(name, 0)))
        cmds.append(("def", src, os.path.join(files, "d%03d" % i), level, strategy, wbits, piece, chunk, flush))

    # the port's inflate: zlib's streams (Python's, every level / strategy / memLevel / window; the official
    # netplay pattern), its own deflate's
    infs = []                                                   # (label, compressed bytes, wbits, sizes, chunk, expect)
    for name in ["tiny", "stats_json", "map", "runs", "noise", "far_repeats"]:
        for level in range(0, 10):
            if name in ("noise",) and level not in (0, 6):
                continue
            infs.append(("zlib L%d %s" % (level, name), zlib.compress(data[name], level), 15, None, 4096, data[name]))
    for strat, sname in ((zlib.Z_FILTERED, "filtered"), (zlib.Z_HUFFMAN_ONLY, "huffman"), (zlib.Z_RLE, "rle"),
                         (zlib.Z_FIXED, "fixed")):
        for name in ("stats_json", "runs"):
            co = zlib.compressobj(6, zlib.DEFLATED, 15, 9, strat)
            infs.append(("zlib %s %s" % (sname, name), co.compress(data[name]) + co.flush(), 15, None, 4096, data[name]))
    for mem in (1, 9):
        co = zlib.compressobj(9, zlib.DEFLATED, 15, mem)
        infs.append(("zlib memLevel %d" % mem, co.compress(data["stats_json"]) + co.flush(), 15, None, 4096,
                     data["stats_json"]))
    for wb in (9, 12):
        co = zlib.compressobj(6, zlib.DEFLATED, wb)
        infs.append(("zlib window %d bits" % wb, co.compress(data["scripts_js"]) + co.flush(), 15, None, 4096,
                     data["scripts_js"]))
    co = zlib.compressobj(6, zlib.DEFLATED, -15)
    infs.append(("zlib raw", co.compress(data["map"]) + co.flush(), -15, None, 4096, data["map"]))
    # sync / partial flushes after each piece (Python's zlib), fed segment by segment
    for fl, fname in ((zlib.Z_SYNC_FLUSH, "sync"), (zlib.Z_PARTIAL_FLUSH, "partial")):
        co = zlib.compressobj(6)
        z = b""
        segs = []
        pos = 0
        for n in net_sizes:
            seg = co.compress(net[pos:pos + n]) + co.flush(fl)
            segs.append(len(seg))
            z += seg
            pos += n
        infs.append(("zlib %s-flushed netplay pieces" % fname, z, 15, segs, 4096, net))
    if official is not None:
        z, marks = c_deflate(official, net, 6, 0, 15, net_sizes, 0, NETPLAY)
        segs = [b - a for a, b in zip([0] + marks[:-1], marks)]
        infs.append(("official libzlib1 netplay adapter L6", z, 15, segs, 4096, net))
        small_net = net[:3000]
        z2, m2 = c_deflate(official, small_net, 6, 0, 15, [n for n in net_sizes if n][:60], 0, NETPLAY)
        infs.append(("official libzlib1 netplay, byte by byte", z2, 15, [1] * len(z2), 7, None))
    bad = bytearray(zlib.compress(data["stats_json"], 6))
    bad[len(bad) // 2] ^= 0x55
    infs.append(("broken stream", bytes(bad), 15, None, 4096, "error"))
    bad2 = bytearray(zlib.compress(data["map"], 6))
    bad2[-2] ^= 0x01
    infs.append(("broken adler-32", bytes(bad2), 15, None, 4096, "error"))
    if "inflate" not in groups:
        infs = []
    inf_base = len(cmds)
    for i, (label, z, wbits, segs, chunk, expect) in enumerate(infs):
        p = os.path.join(files, "z%03d" % i)
        open(p, "wb").write(z)
        if segs is not None:
            with open(p + ".sizes", "w") as f:
                f.write("".join("%d\n" % n for n in segs))
        cmds.append(("inf", p, p + ".out", wbits, 0 if segs is not None else max(len(z), 1), chunk, 0))
    # the network's receive pattern (readNoInt over ZlibCompressionAdapter) on streams far over 32 KB: zlib's inflate
    # goes on until the input or the room runs out, and the adapter relies on it (with room left, all input taken)
    adapts = []                                                 # (label, stream, segments or None, max size, received, data)
    if official is not None:
        z, marks = c_deflate(official, net, 6, 0, 15, net_sizes, 0, NETPLAY)
        segs = [b - a for a, b in zip([0] + marks[:-1], marks)]
        for max_size, received in ((8192, 1), (1000, 1), (8192, 0), (65536, 0), (300, 0)):
            adapts.append(("official netplay stream, room %d, %s" % (max_size, "as sent" if received else "full receives"),
                           z, segs, max_size, received, net))
    zz = zlib.compress(data["stats_json"], 6)
    adapts.append(("zlib stats_json stream, room 4096", zz, None, 4096, 0, data["stats_json"]))
    zr = zlib.compress(data["runs"], 9)
    adapts.append(("zlib runs stream (long matches), room 1000", zr, None, 1000, 0, data["runs"]))
    if "receive" not in groups:
        adapts = []
    adapt_base = len(cmds)
    for i, (label, z, segs, max_size, received, want) in enumerate(adapts):
        p = os.path.join(files, "a%03d" % i)
        open(p, "wb").write(z)
        if segs is not None:
            with open(p + ".sizes", "w") as f:
                f.write("".join("%d\n" % n for n in segs))
        cmds.append(("adapt", p, p + ".out", max_size, received))
    cmp_at = len(cmds)
    cmds.append(("cmp2", os.path.join(files, "stats_json"), os.path.join(files, "cmp2"), 6))
    speed_at = len(cmds)
    cmds.append(("speed", os.path.join(files, "stats_json"), 6, 2))

    results, seconds = run_exe(exe, cmds)
    if len(results) != len(cmds):
        print("zlib: %d results for %d commands" % (len(results), len(cmds)))
        return False
    ok = True
    passed = 0
    total = 0
    exact = 0
    exact_total = 0
    x64_same = 0

    # deflate: the C byte for byte; zlib reads it; after each flush the input so far
    for i, (label, name, level, strategy, wbits, piece, chunk, flush) in enumerate(defs):
        total += 1
        res = results[i]
        if res[1] != "ok":
            print("  FAIL %s: %s" % (label, " ".join(res[1:])))
            ok = False
            continue
        out = open(os.path.join(files, "d%03d" % i), "rb").read()
        marks = [int(x) for x in open(os.path.join(files, "d%03d.marks" % i)).read().split()]
        src = data[name] if name != "netplay" else net
        sizes = sizes_of(name, piece)
        problems = []
        if refs is not None:
            exact_total += 1
            cout, cmarks = c_deflate(refs[1], src, level, strategy, wbits, sizes, chunk, flush)
            if cout == out and cmarks == marks:
                exact += 1
            else:
                k = next((j for j in range(min(len(cout), len(out))) if cout[j] != out[j]), min(len(cout), len(out)))
                problems.append("not miniz's bytes (%d against %d bytes, first difference at %d; marks %s)"
                                % (len(out), len(cout), k, "equal" if cmarks == marks else "differ"))
            xout, xmarks = c_deflate(refs[0], src, level, strategy, wbits, sizes, chunk, flush)
            if xout == cout and xmarks == cmarks:
                x64_same += 1
        # zlib reads it
        try:
            d = zlib.decompressobj(wbits if wbits < 0 else 15)
            if flush in (Z_PARTIAL_FLUSH, Z_SYNC_FLUSH, Z_FULL_FLUSH, NETPLAY):
                got = b""
                pos = 0
                prev = 0
                for n, m in zip(sizes, marks):
                    got += d.decompress(out[prev:m])
                    pos += n
                    prev = m
                    if got != src[:pos]:
                        problems.append("zlib after the flush at input %d: %d bytes, not the input so far" % (pos, len(got)))
                        break
                got += d.decompress(out[prev:])
            else:
                got = d.decompress(out) + d.flush()
            if got != src:
                problems.append("zlib reads %d bytes, not the %d" % (len(got), len(src)))
        except zlib.error as e:
            problems.append("zlib: %s" % e)
        if official is not None and wbits > 0:
            segs = [b - a for a, b in zip([0] + marks[:-1], marks)] + ([len(out) - marks[-1]] if marks else [len(out)])
            segs = [s for s in segs if s > 0] or [len(out)]
            try:
                dout, dmarks, st = c_inflate(official, out, 15, segs, 65536)
                if dout != src:
                    problems.append("libzlib1.dll reads %d bytes, not the %d (status %d)" % (len(dout), len(src), st))
            except RuntimeError as e:
                problems.append("libzlib1.dll: %s" % e)
        # the port reads its own
        if problems:
            ok = False
            print("  FAIL %s: %s" % (label, "; ".join(problems)))
        else:
            passed += 1

    # inflate: the data, and at each mark the input so far
    for i, (label, z, wbits, segs, chunk, expect) in enumerate(infs):
        total += 1
        res = results[inf_base + i]
        p = os.path.join(files, "z%03d" % i)
        if res[1] != "ok":
            print("  FAIL inflate %s: %s" % (label, " ".join(res[1:])))
            ok = False
            continue
        status = int(res[2])
        got = open(p + ".out", "rb").read()
        if expect == "error":
            if status in (Z_DATA_ERROR,):
                passed += 1
            else:
                ok = False
                print("  FAIL inflate %s: status %d, not Z_DATA_ERROR (zlib: error)" % (label, status))
            continue
        if expect is None:                                  # the byte-by-byte run: as zlib inflates it
            ref_out = zlib.decompressobj().decompress(z)
            if got == ref_out and len(got) > 0:
                passed += 1
            else:
                ok = False
                print("  FAIL inflate %s: %d bytes, zlib %d" % (label, len(got), len(ref_out)))
            continue
        problems = []
        if got != expect:
            problems.append("%d bytes, not the %d" % (len(got), len(expect)))
        if segs is not None:
            marks = [int(x) for x in open(p + ".out.marks").read().split()]
            want = []
            pos = 0
            for n in net_sizes:
                pos += n
                want.append(pos)
            if marks[:len(want)] != want:
                k = next((j for j in range(min(len(marks), len(want))) if marks[j] != want[j]), None)
                problems.append("after segment %s: %s bytes, not %s" % (k, marks[k] if k is not None else "-",
                                                                         want[k] if k is not None else "-"))
            if status != Z_OK and status != Z_BUF_ERROR:
                problems.append("status %d (an unfinished stream: Z_OK)" % status)
        elif status != Z_STREAM_END:
            problems.append("status %d, not Z_STREAM_END" % status)
        if problems:
            ok = False
            print("  FAIL inflate %s: %s" % (label, "; ".join(problems)))
        else:
            passed += 1

    # the receive pattern: the data, no call with room and input left, and call for call zlib's (the official DLL's)
    # produced bytes and avail_out; avail_in too where room was left (zlib may keep input when the room ran out)
    adapt_same = 0
    for i, (label, z, segs, max_size, received, want) in enumerate(adapts):
        total += 1
        res = results[adapt_base + i]
        p = os.path.join(files, "a%03d" % i)
        if res[1] != "ok":
            ok = False
            print("  FAIL receive %s: %s" % (label, " ".join(res[1:])))
            continue
        got = open(p + ".out", "rb").read()
        calls = [tuple(int(x) for x in ln.split()) for ln in open(p + ".out.calls").read().splitlines() if ln.strip()]
        problems = []
        if got != want:
            k = next((j for j in range(min(len(got), len(want))) if got[j] != want[j]), min(len(got), len(want)))
            problems.append("%d bytes, not the %d (first difference at %d)" % (len(got), len(want), k))
        if int(res[3]) != 0:
            problems.append("%s calls with room and input left (the adapter's assert; the input dropped)" % res[3])
        if official is not None:
            zout, zcalls, zviol, zst = c_adapter(official, z, max_size, received, segs)
            if zout != want or zviol:
                problems.append("(the official DLL itself: %d bytes, %d asserts)" % (len(zout), zviol))
            same = len(calls) == len(zcalls) and all(
                c[0] == zc[0] and c[2] == zc[2] and (c[2] == 0 or c[1] == zc[1]) for c, zc in zip(calls, zcalls))
            if same:
                adapt_same += 1
            else:
                k = next((j for j in range(min(len(calls), len(zcalls)))
                          if calls[j][0] != zcalls[j][0] or calls[j][2] != zcalls[j][2]
                          or (calls[j][2] != 0 and calls[j][1] != zcalls[j][1])), min(len(calls), len(zcalls)))
                problems.append("not zlib's calls: %d against %d, first difference at call %d (port %s, zlib %s)" % (
                    len(calls), len(zcalls), k, calls[k] if k < len(calls) else "-", zcalls[k] if k < len(zcalls) else "-"))
        if problems:
            ok = False
            print("  FAIL receive %s: %s" % (label, "; ".join(problems)))
        else:
            passed += 1

    # round trips: the port inflates each of its deflate outputs (a second run)
    rt = []
    for i, (label, name, level, strategy, wbits, piece, chunk, flush) in enumerate(defs):
        if results[i][1] == "ok":
            p = os.path.join(files, "d%03d" % i)
            rt.append(("inf", p, p + ".rt", wbits if wbits < 0 else 15, max(os.path.getsize(p), 1), 4096, 0))
    rres, rsec = run_exe(exe, rt)
    rt_ok = 0
    for j, c in enumerate(rt):
        i = int(os.path.basename(c[1])[1:])
        label, name = defs[i][0], defs[i][1]
        src = data[name] if name != "netplay" else net
        got = open(c[2], "rb").read()
        want_status = Z_OK if (defs[i][7] == NETPLAY and defs[i][6] == 0) else Z_STREAM_END
        if rres[j][1] == "ok" and got == src and int(rres[j][2]) in (want_status, Z_BUF_ERROR if want_status == Z_OK else 99):
            rt_ok += 1
        else:
            ok = False
            print("  FAIL round trip %s: %s, %d bytes of %d" % (label, " ".join(rres[j][1:]), len(got), len(src)))
    total += len(rt)
    passed += rt_ok

    # mz_compress2
    total += 1
    cres = results[cmp_at]
    try:
        if cres[1] == "ok" and int(cres[2]) == Z_OK and zlib.decompress(open(os.path.join(files, "cmp2"), "rb").read()) == data["stats_json"]:
            passed += 1
        else:
            ok = False
            print("  FAIL mz_compress2: %s" % " ".join(cres[1:]))
    except zlib.error as e:
        ok = False
        print("  FAIL mz_compress2: zlib: %s" % e)

    sres = results[speed_at]
    print("zlib: %d of %d checks pass (%d deflate runs, %d inflate runs, %d of the network's receive pattern, "
          "%d round trips, mz_compress2)" % (passed, total, len(defs), len(infs), len(adapts), len(rt)))
    if official is not None:
        print("  the receive pattern: %d of %d runs call for call as zlib's (produced bytes, avail_out; avail_in "
              "with room left)" % (adapt_same, len(adapts)))
    if refs is not None:
        print("  miniz.c v1.11's bytes (gcc, the x86 code paths): %d of %d deflate runs identical, output and "
              "flush points; the x64 build's the same in %d of %d" % (exact, exact_total, x64_same, exact_total))
    else:
        print("  (no gcc: the comparison with miniz.c's own bytes was not made)")
    if official is not None:
        print("  the official libzlib1.dll (zlib %s) reads every deflate output; its netplay-pattern stream "
              "inflates segment by segment" % official.version.decode())
    print("  speed (ZlibTests, WzNet's compile settings): deflate %s MB/s, inflate %s MB/s at level 6 (stats JSON); "
          "%.1f s for the runs" % (sres[2], sres[3], seconds + rsec))
    if with_zip:
        ok = run_zip(stage) and ok
    if not keep:
        shutil.rmtree(files, ignore_errors=True)
    return ok and passed == total


def run_zip(stage):
    """tools/zip_test.py's checks with a ZipTests built from these miniz modules (and the stage's physfs_zip_c)"""
    import zip_test
    out_bin = os.path.join(WORK, "bin") if stage else build.BIN
    ported = [(n, p) for n, p in zip_test.PORTED if n != "physfs_miniz_c"]
    ported = [(n, source(stage, os.path.join("framework", n + ".bas")) if n.startswith("physfs") else
               os.path.normpath(os.path.join(zip_test.FOLDER, p))) for n, p in ported]
    ported += [(m, source(stage, os.path.join("shared", "miniz", m + ".bas"))) for m in MINIZ_MODULES]

    def build_zip():
        lines = ["Type=Exe", build.STDOLE]
        for n in build.sources(zip_test.FOLDER):
            path = os.path.join(zip_test.FOLDER, n)
            build.normalise(path)
            lines.append("Module=%s; %s" % (build.vb_name(path), path))
        for name, path in ported:
            build.normalise(path)
            lines.append("Module=%s; %s" % (name, path))
        lines.append(build.VBP_TAIL.format(name="ZipTests", ext="exe", path=out_bin, startup="Sub Main",
                                           compat="", bounds="0", optimisation="2", threading=""))
        vbp_dir = os.path.join(WORK, "vbp")
        os.makedirs(vbp_dir, exist_ok=True)
        vbp = os.path.join(vbp_dir, "ZipTests.vbp")
        with open(vbp, "w", encoding="latin-1", newline="\r\n") as f:
            f.write("\n".join(lines))
        saved = build.BIN
        build.BIN = out_bin
        try:
            return build.make("ZipTests", vbp_dir, vbp)
        finally:
            build.BIN = saved

    real_run = subprocess.run

    def redirected(args, *a, **k):              # zip_test runs bin\ZipTests.exe: this one instead
        if isinstance(args, list) and args and os.path.basename(args[0]).lower() == "ziptests.exe":
            args = [os.path.join(out_bin, "ZipTests.exe")] + args[1:]
        return real_run(args, *a, **k)

    zip_test.build_exe = build_zip
    zip_test.subprocess = type("sp", (), {"run": staticmethod(redirected), "PIPE": subprocess.PIPE})
    print("zip (ZipTests with src/shared/miniz's inflate):")
    return zip_test.run()


def main(argv):
    keep = "--keep" in argv
    stage = None
    if "--stage" in argv:
        stage = argv[argv.index("--stage") + 1]
    only = argv[argv.index("--only") + 1].split(",") if "--only" in argv else None
    return 0 if run(keep, stage, "--zip" in argv, only) else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
