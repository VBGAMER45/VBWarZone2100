"""
theora_test.py - tests WzSequence's libtheora port (cluster K, fork V1: src/sequence/theora_*.bas) on the
campaign videos.

  python -I tools\\theora_test.py [--keep] [--all] [--fast] [--no-build] [--pp-only] [--no-pp] [name ...]

* Python demuxes each video of data/sequences.wz (zipfile + the Ogg page parser below) into a packet file
  of its Theora stream: per packet a little-endian int32 byte count, int32 b_o_s, int32 e_o_s, then the
  bytes (the three header packets first)
* bin/TheoraTests.exe (src/theoratests: the port's modules compiled in, with bounds checks; built here with
  tools/build.py's write_vbp and make, never build.py itself) decodes them through the legacy theora.h API
  as lib/sequence/sequence.cpp does (theora_decode_header, theora_decode_init, theora_control,
  theora_decode_packetin, theora_decode_YUVout) into raw YUV 4:2:0 frames: the picture region (frame_width x
  frame_height at offset_x, offset_y) of each output, Y then U then V, as ffmpeg's -pix_fmt yuv420p
* at post-processing level 0 the frames must equal ffmpeg's decode (`ffmpeg -threads 1 -i <video> -f
  rawvideo -pix_fmt yuv420p`: its frame threads are not deterministic) byte for byte; ffmpeg has no
  post-processing, so at the maximum level (TH_DECCTL_GET_PPLEVEL_MAX, the level sequence.cpp sets) the
  frames are compared with the official 4.7.0 build's libtheoradec.dll (ref/official/wz470/bin, 64-bit)
  called through ctypes, and the frames the post-processing changed are counted (coverage); without that
  DLL the maximum level is only checked to run
* the decode speed: frames per second of theora_decode_packetin + theora_decode_YUVout (the exe's timing)

By default a subset (every 8th video by name, the longest and one with a cropped picture) is tested; --all
tests all of them; names select videos by substring. --fast builds and runs bin/TheoraTestsFast.exe, compiled
as WzSequence is (optimised, no bounds checks), for the speed WzSequence will have. The work files are in
build/theora_tests (deleted unless --keep). Each video is run on its own (the frames of all of them would
take gigabytes).
"""
import ctypes
import os
import shutil
import struct
import subprocess
import sys
import time
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import build  # noqa: E402

SRC = os.path.join(build.SRC, "theoratests")
WORK = os.path.join(build.BUILD, "theora_tests")
ARCHIVE = os.path.join(build.ROOT, "data", "sequences.wz")
OFFICIAL_BIN = os.path.join(build.ROOT, "ref", "official", "wz470", "bin")
EXE = os.path.join(build.BIN, "TheoraTests.exe")
# the port's modules the test exe compiles in (besides its own folder): src/sequence/theora_*.bas
SEQUENCE = os.path.join(build.SRC, "sequence")
# the modules of src/shared the port uses (math_vb: U32, SAR, ToInt8); the others need names their projects
# define (as tools/run_tests.py's VorbisTests says), so they are left out
SHARED_USED = ("math_vb.bas",)


# ---------------------------------------------------------------------------------------------------------
# Ogg demuxing (RFC 3533): pages, then packets of one logical stream

def ogg_pages(data):
    """Yields (serialno, header_type, granulepos, segment table, body) for each page."""
    pos = 0
    n = len(data)
    while pos + 27 <= n:
        if data[pos:pos + 4] != b"OggS":
            nxt = data.find(b"OggS", pos + 1)
            if nxt < 0:
                return
            pos = nxt
            continue
        header_type = data[pos + 5]
        granulepos, serialno = struct.unpack_from("<qI", data, pos + 6)
        nsegs = data[pos + 26]
        segs = data[pos + 27:pos + 27 + nsegs]
        body_len = sum(segs)
        body_start = pos + 27 + nsegs
        yield serialno, header_type, granulepos, segs, data[body_start:body_start + body_len]
        pos = body_start + body_len


def theora_packets(data):
    """The packets of the first Theora stream: a list of (bytes, b_o_s, e_o_s)."""
    serial = None
    packets = []
    partial = b""
    first = True
    for serialno, header_type, granulepos, segs, body in ogg_pages(data):
        if serial is None:
            if header_type & 2 and body[:7] == b"\x80theora":
                serial = serialno
            else:
                continue
        if serialno != serial:
            continue
        off = 0
        if not header_type & 1:
            partial = b""
        for s in segs:
            partial += body[off:off + s]
            off += s
            if s < 255:
                packets.append([partial, 1 if first else 0, 0])
                first = False
                partial = b""
        if header_type & 4 and packets:
            packets[-1][2] = 1
    return [tuple(p) for p in packets]


def write_packet_file(path, packets):
    with open(path, "wb") as f:
        for pkt, bos, eos in packets:
            f.write(struct.pack("<iii", len(pkt), bos, eos))
            f.write(pkt)


def theora_info(packets):
    """The identification header's fields: (frame w, h, pic w, h, pic x, pic y from the top, fps n, d)."""
    p = packets[0][0]
    fw = int.from_bytes(p[10:12], "big") << 4
    fh = int.from_bytes(p[12:14], "big") << 4
    pw = int.from_bytes(p[14:17], "big")
    ph = int.from_bytes(p[17:20], "big")
    px = p[20]
    py = p[21]
    fn = int.from_bytes(p[22:26], "big")
    fd = int.from_bytes(p[26:30], "big")
    return fw, fh, pw, ph, px, fh - ph - py, fn, fd


def videos():
    with zipfile.ZipFile(ARCHIVE) as z:
        return sorted(n for n in z.namelist() if n.lower().endswith(".ogg"))


def read_video(name):
    with zipfile.ZipFile(ARCHIVE) as z:
        return z.read(name)


# ---------------------------------------------------------------------------------------------------------
# References

def ffmpeg_frames(path, w, h):
    """ffmpeg's decode of the video file at path: a list of yuv420p frames (bytes). One decoder thread: ffmpeg
    8.1's frame-threaded VP3/Theora decoder is not deterministic (measured: cam2dico.ogg's frame 199 and
    titles.ogg's frame 289 differ from run to run by a few chroma values; with -threads 1 they are always
    libtheora's)."""
    r = subprocess.run(["ffmpeg", "-v", "error", "-threads", "1", "-i", path, "-an", "-f", "rawvideo",
                        "-pix_fmt", "yuv420p", "-fps_mode", "passthrough", "-"], capture_output=True)
    if r.returncode != 0:
        raise RuntimeError("ffmpeg: " + r.stderr.decode("latin-1")[:300])
    size = w * h + 2 * ((w + 1) // 2) * ((h + 1) // 2)
    data = r.stdout
    return [data[i:i + size] for i in range(0, len(data) - size + 1, size)]


class ogg_packet(ctypes.Structure):
    _fields_ = [("packet", ctypes.c_void_p), ("bytes", ctypes.c_long), ("b_o_s", ctypes.c_long),
                ("e_o_s", ctypes.c_long), ("granulepos", ctypes.c_int64), ("packetno", ctypes.c_int64)]


class th_info(ctypes.Structure):
    _fields_ = [("version_major", ctypes.c_ubyte), ("version_minor", ctypes.c_ubyte),
                ("version_subminor", ctypes.c_ubyte), ("frame_width", ctypes.c_uint32),
                ("frame_height", ctypes.c_uint32), ("pic_width", ctypes.c_uint32), ("pic_height", ctypes.c_uint32),
                ("pic_x", ctypes.c_uint32), ("pic_y", ctypes.c_uint32), ("fps_numerator", ctypes.c_uint32),
                ("fps_denominator", ctypes.c_uint32), ("aspect_numerator", ctypes.c_uint32),
                ("aspect_denominator", ctypes.c_uint32), ("colorspace", ctypes.c_int), ("pixel_fmt", ctypes.c_int),
                ("target_bitrate", ctypes.c_int), ("quality", ctypes.c_int),
                ("keyframe_granule_shift", ctypes.c_int)]


class th_comment(ctypes.Structure):
    _fields_ = [("user_comments", ctypes.c_void_p), ("comment_lengths", ctypes.c_void_p),
                ("comments", ctypes.c_int), ("vendor", ctypes.c_char_p)]


class th_img_plane(ctypes.Structure):
    _fields_ = [("width", ctypes.c_int), ("height", ctypes.c_int), ("stride", ctypes.c_int),
                ("data", ctypes.c_void_p)]


TH_DECCTL_GET_PPLEVEL_MAX = 1
TH_DECCTL_SET_PPLEVEL = 3


class OfficialDecoder:
    """The official 4.7.0 build's libtheoradec.dll through ctypes (th_ API: the legacy wrapper calls these)."""

    def __init__(self):
        if hasattr(os, "add_dll_directory"):
            self._dirs = os.add_dll_directory(OFFICIAL_BIN)
        self.lib = ctypes.CDLL(os.path.join(OFFICIAL_BIN, "libtheoradec.dll"))
        lib = self.lib
        lib.th_decode_headerin.argtypes = [ctypes.POINTER(th_info), ctypes.POINTER(th_comment),
                                           ctypes.POINTER(ctypes.c_void_p), ctypes.POINTER(ogg_packet)]
        lib.th_decode_alloc.argtypes = [ctypes.POINTER(th_info), ctypes.c_void_p]
        lib.th_decode_alloc.restype = ctypes.c_void_p
        lib.th_decode_ctl.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_void_p, ctypes.c_size_t]
        lib.th_decode_packetin.argtypes = [ctypes.c_void_p, ctypes.POINTER(ogg_packet), ctypes.c_void_p]
        lib.th_decode_ycbcr_out.argtypes = [ctypes.c_void_p, ctypes.POINTER(th_img_plane * 3)]
        lib.th_decode_free.argtypes = [ctypes.c_void_p]
        lib.th_setup_free.argtypes = [ctypes.c_void_p]
        lib.th_info_init.argtypes = [ctypes.POINTER(th_info)]
        lib.th_comment_init.argtypes = [ctypes.POINTER(th_comment)]
        lib.th_comment_clear.argtypes = [ctypes.POINTER(th_comment)]

    def decode(self, packets, pp_max):
        """The picture region of each output frame as yuv420p bytes (pp_max: the maximum pp level)."""
        lib = self.lib
        info = th_info()
        tc = th_comment()
        setup = ctypes.c_void_p()
        lib.th_info_init(ctypes.byref(info))
        lib.th_comment_init(ctypes.byref(tc))
        keep = []

        def packet(data, bos, eos):
            buf = ctypes.create_string_buffer(data, max(len(data), 1))
            keep.append(buf)
            op = ogg_packet()
            op.packet = ctypes.cast(buf, ctypes.c_void_p)
            op.bytes = len(data)
            op.b_o_s = bos
            op.e_o_s = eos
            op.granulepos = -1
            return op

        i = 0
        while i < len(packets):
            op = packet(*packets[i])
            r = lib.th_decode_headerin(ctypes.byref(info), ctypes.byref(tc), ctypes.byref(setup), ctypes.byref(op))
            if r < 0:
                raise RuntimeError("th_decode_headerin: %d" % r)
            i += 1
            if r == 1:      # the setup header: the next packet is data
                break
        dec = lib.th_decode_alloc(ctypes.byref(info), setup)
        lib.th_setup_free(setup)
        level = ctypes.c_int(0)
        lib.th_decode_ctl(dec, TH_DECCTL_GET_PPLEVEL_MAX, ctypes.byref(level), ctypes.sizeof(level))
        if not pp_max:
            level = ctypes.c_int(0)
        lib.th_decode_ctl(dec, TH_DECCTL_SET_PPLEVEL, ctypes.byref(level), ctypes.sizeof(level))
        pw, ph, px, py = info.pic_width, info.pic_height, info.pic_x, info.pic_y
        frames = []
        planes = (th_img_plane * 3)()
        for data, bos, eos in packets[i:]:
            op = packet(data, bos, eos)
            r = lib.th_decode_packetin(dec, ctypes.byref(op), None)
            keep.clear()
            if r < 0:
                raise RuntimeError("th_decode_packetin: %d" % r)
            lib.th_decode_ycbcr_out(dec, ctypes.byref(planes))
            out = bytearray()
            for pli in range(3):
                pl = planes[pli]
                sx = 0 if pli == 0 else 1
                x0, y0 = px >> sx, py >> sx
                w = pw if pli == 0 else (pw + 1) >> 1
                h = ph if pli == 0 else (ph + 1) >> 1
                for row in range(h):
                    out += ctypes.string_at(pl.data + (y0 + row) * pl.stride + x0, w)
            frames.append(bytes(out))
        lib.th_decode_free(dec)
        lib.th_comment_clear(ctypes.byref(tc))
        return int(level.value), frames


# ---------------------------------------------------------------------------------------------------------
# The test exe

def build_exe(fast=False):
    """bin/TheoraTests.exe: src/theoratests, src/sequence/theora_*.bas, src/shared's SHARED_USED and ogg_h.bas
    (no DLL references; bounds checks on, no optimisation). fast: bin/TheoraTestsFast.exe, compiled as WzSequence
    is (optimised for fast code, no bounds checks), for the speed."""
    name = "TheoraTestsFast" if fast else "TheoraTests"
    vbp = build.write_vbp(name, "exe", SRC, [], False, None)
    with open(vbp, "r", encoding="latin-1") as f:
        lines = [l for l in f.read().split("\n")
                 if "..\\shared\\" not in l or any(m in l for m in SHARED_USED)]
    if fast:
        lines = ["BoundsCheck=-1" if l.startswith("BoundsCheck=") else
                 "OptimizationType=0" if l.startswith("OptimizationType=") else l for l in lines]
    k = next(i for i, l in enumerate(lines) if l.startswith("Startup="))
    extra = []
    for n in sorted(os.listdir(SEQUENCE)):
        if n.lower().startswith("theora_") and n.lower().endswith(".bas"):
            path = os.path.join(SEQUENCE, n)
            build.normalise(path)
            extra.append("Module=%s; ..\\sequence\\%s" % (build.vb_name(path), n))
    ogg_h = os.path.join(build.SHARED_OGG, "ogg_h.bas")
    build.normalise(ogg_h)
    extra.append("Module=ogg_h; ..\\shared\\ogg\\ogg_h.bas")
    lines[k:k] = extra
    with open(vbp, "w", encoding="latin-1", newline="\r\n") as f:
        f.write("\n".join(lines))
    return build.make(name, SRC, vbp)


def run_exe(exe, listfile, outdir, pp_max):
    """Runs the exe on a list of packet files: (exit code, seconds, {n: fields}, total fields)."""
    t = time.time()
    r = subprocess.run([exe, listfile, outdir, "max" if pp_max else "0"], capture_output=True)
    elapsed = time.time() - t
    results = {}
    total = None
    path = os.path.join(outdir, "results.txt")
    if os.path.exists(path):
        with open(path, "r", encoding="latin-1") as f:
            for line in f:
                parts = line.rstrip("\r\n").split("\t")
                if parts[0] in ("total", "fatal"):
                    total = parts
                elif parts[0].isdigit():
                    results[int(parts[0])] = parts[1:]
    return r.returncode, elapsed, results, total


def read_frames(path, size):
    with open(path, "rb") as f:
        data = f.read()
    return [data[i:i + size] for i in range(0, len(data) - size + 1, size)]


def first_difference(a, b, w, h):
    """Where two yuv420p frames first differ: plane, x, y, and the number of differing bytes."""
    cw, ch = (w + 1) // 2, (h + 1) // 2
    ndiff = sum(1 for x, y in zip(a, b) if x != y)
    for i, (x, y) in enumerate(zip(a, b)):
        if x != y:
            if i < w * h:
                return "Y (%d,%d) %d/%d, %d bytes differ" % (i % w, i // w, x, y, ndiff)
            i -= w * h
            pl = "U" if i < cw * ch else "V"
            i %= cw * ch
            return "%s (%d,%d) %d/%d, %d bytes differ" % (pl, i % cw, i // cw, x, y, ndiff)
    return "sizes %d/%d" % (len(a), len(b))


def compare(got, ref, w, h):
    if len(got) != len(ref):
        n = min(len(got), len(ref))
        same = sum(1 for a, b in zip(got, ref) if a == b)
        return False, "%d frames vs %d, %d of the first %d equal" % (len(got), len(ref), same, n)
    for i, (a, b) in enumerate(zip(got, ref)):
        if a != b:
            return False, "frame %d differs first: %s" % (i, first_difference(a, b, w, h))
    return True, "bit exact"


def choose(all_videos, names):
    every = videos()
    if names:
        return [v for v in every if any(n.lower() in v.lower() for n in names)]
    if all_videos:
        return every
    sizes = {}
    with zipfile.ZipFile(ARCHIVE) as z:
        for i in z.infolist():
            sizes[i.filename] = i.file_size
    chosen = every[::8]
    longest = max(every, key=lambda v: sizes[v])
    if longest not in chosen:
        chosen.append(longest)
    small = [v for v in every if v not in chosen]
    # one of the 192x176 videos (a picture smaller than the frame: the crop)
    for v in small:
        if theora_info(theora_packets(read_video(v)))[1] != 240:
            chosen.append(v)
            break
    return chosen


def run(keep=False, all_videos=False, names=None, do_build=True, pp0=True, ppmax=True, fast=False):
    """The test (run_tests.py calls it): True when every chosen video decodes bit exact at both levels."""
    names = names or []
    if shutil.which("ffmpeg") is None:
        print("theora: ffmpeg not found (the reference decoder)")
        return False
    if os.path.isdir(WORK):
        shutil.rmtree(WORK)
    os.makedirs(WORK)
    exe = os.path.join(build.BIN, "TheoraTestsFast.exe" if fast else "TheoraTests.exe")
    if do_build or not os.path.exists(exe):
        if not build_exe(fast):
            print("theora: %s did not build" % os.path.basename(exe))
            return False
    chosen = choose(all_videos, names)
    official = None
    if ppmax:
        try:
            official = OfficialDecoder()
        except OSError as e:
            print("theora: the official libtheoradec.dll could not be loaded (%s): the maximum level is only "
                  "checked to run" % e)
    modes = ([False] if pp0 else []) + ([True] if ppmax else [])
    stats = {m: {"passed": 0, "frames": 0, "ms": 0.0, "details": [], "elapsed": 0.0} for m in modes}
    t_start = time.time()
    for n, name in enumerate(chosen):
        data = read_video(name)
        vpath = os.path.join(WORK, "video.ogg")
        with open(vpath, "wb") as f:
            f.write(data)
        packets = theora_packets(data)
        fw, fh, pw, ph, px, py, fn, fd = theora_info(packets)
        ppath = os.path.join(WORK, "video.pkt")
        write_packet_file(ppath, packets)
        listfile = os.path.join(WORK, "list.txt")
        with open(listfile, "w", encoding="latin-1", newline="\r\n") as lf:
            lf.write(ppath + "\n")
        size = pw * ph + 2 * ((pw + 1) // 2) * ((ph + 1) // 2)
        level0 = None
        for pp_max in modes:
            st = stats[pp_max]
            outdir = os.path.join(WORK, "out_max" if pp_max else "out_0")
            shutil.rmtree(outdir, ignore_errors=True)
            os.makedirs(outdir)
            code, elapsed, results, total = run_exe(exe, listfile, outdir, pp_max)
            st["elapsed"] += elapsed
            res = results.get(0)
            if res is None or res[0] == "error":
                st["details"].append("  FAIL %s: %s (exit code %d%s)" % (
                    name, "no result" if res is None else "error " + " ".join(res[1:]), code,
                    ", " + " ".join(total) if total else ""))
                continue
            nframes, level, ms = int(res[0]), int(res[1]), float(res[2].replace(",", "."))
            st["frames"] += nframes
            st["ms"] += ms
            got = read_frames(os.path.join(outdir, "0.yuv"), size)
            if pp_max:
                if official is None:
                    differs = sum(1 for a, b in zip(got, level0 or []) if a != b)
                    good = len(got) == nframes and level == 7
                    why = "level %d, %d frames, %d differ from level 0 (not compared)" % (level, len(got), differs)
                else:
                    olevel, ref = official.decode(packets, True)
                    good, why = compare(got, ref, pw, ph)
                    why += " (pp levels %d/%d" % (level, olevel)
                    if level0 is not None:
                        # coverage: the frames the post-processing changed
                        why += ", %d frames differ from level 0" % sum(1 for a, b in zip(got, level0) if a != b)
                    why += ")"
                    good = good and level == olevel
            else:
                ref = ffmpeg_frames(vpath, pw, ph)
                good, why = compare(got, ref, pw, ph)
                level0 = got
            if good:
                st["passed"] += 1
            st["details"].append("  %s%s: %dx%d (frame %dx%d), %d frames, %.0f frames/s: %s"
                                 % ("" if good else "FAIL ", name, pw, ph, fw, fh, nframes,
                                    nframes * 1000.0 / ms if ms > 0 else 0, why))
            if not keep:
                shutil.rmtree(outdir, ignore_errors=True)
    ok = True
    for pp_max in modes:
        st = stats[pp_max]
        label = "pp max" if pp_max else "pp 0"
        what = ("as ffmpeg decodes them" if not pp_max else
                "as the official libtheoradec.dll decodes them" if official else "decoded (not compared)")
        fps = st["frames"] * 1000.0 / st["ms"] if st["ms"] > 0 else 0
        print("theora (%s, %s): %d of %d videos %s, %d frames, decode %.0f frames/s (packetin + YUVout, %s)"
              % (label, os.path.basename(exe), st["passed"], len(chosen), what, st["frames"], fps,
                 "optimised, no bounds checks" if fast else "bounds checks, no optimisation"))
        for d in st["details"]:
            print(d)
        if st["passed"] != len(chosen):
            ok = False
    print("theora: %.0f s" % (time.time() - t_start))
    if not keep:
        shutil.rmtree(WORK, ignore_errors=True)
    return ok


def main(argv):
    names = [a for a in argv if not a.startswith("--")]
    ok = run(keep="--keep" in argv, all_videos="--all" in argv, names=names, do_build="--no-build" not in argv,
             pp0="--pp-only" not in argv, ppmax="--no-pp" not in argv, fast="--fast" in argv)
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
