"""
opus_test.py - cluster K's test of the music's decoding path (fork O1): libogg's framing (src/shared/ogg),
opusfile (src/sound/opusfile_*), oggopus.cpp's WZOpusDecoder, over the libopus port (O2's CELT and range
decoder, O3's SILK and the top level).

Builds bin/OpusTests.exe (src/opustests with src/shared, src/shared/ogg and the WzSound modules, no DLL,
array bounds checks on) and bin/OpusFast.exe (the same compiled as WzSound is: fast code, no bounds
checks, for the speed; a shorter name: a class's programmatic id, project.class, has 39 characters at most), decodes Ogg Opus files with it and compares with ffmpeg's libopus decoder (the
reference: libopus 1.5's float decoder, with the header gain, after the pre-skip, end-trimmed):
  * the official music: data/music/menu.opus and two album tracks (--all: every track)
  * synthetic files made with ffmpeg's libopus encoder (build/opus_test/synth, cached): SILK (voip, NB / MB /
    WB, 8-16 kbps, 10-60 ms frames, with LBRR), hybrid (SWB / FB speech), CELT (music, 64-256 kbps, mono and
    stereo, 2.5-60 ms frames), 5.1 (a multistream file, mapping family 1), a file with a page missing
Per file: the TOC configurations found (mode, bandwidth, frame size), the sample count (= op_pcm_total and
ffmpeg's), the largest difference of op_read_float's samples from ffmpeg's (float) and of op_read_stereo's
16-bit samples from ffmpeg's 16-bit decode; op_read_stereo against this script's own float32 port of
opusfile's float-to-short filter (noise shaping, dither: bit exact expected) run on the port's float output
(the first seconds); op_read_float_stereo's downmix for 5.1; WZOpusDecoder (both callbacks) against
op_read_stereo's bytes (4.7.0 ends a track one Ogg page early: its SEEK_END callback); op_pcm_seek /
op_pcm_tell / op_raw_seek; the decode speed (x real time).

  python -I tools\\opus_test.py [--keep] [--all] [--only=<part of a file name>] [--nobuild]     (run(keep) -> bool)
"""
import array
import os
import shutil
import struct
import subprocess
import sys

ROOT = r"C:\Users\Owner\desktop\websites\warzone"
sys.path.insert(0, os.path.join(ROOT, "tools"))
import build  # noqa: E402

FOLDER = os.path.join(ROOT, "src", "opustests")
SOUND = os.path.join(ROOT, "src", "sound")
WORK = os.path.join(ROOT, "build", "opus_test")
SYNTH = os.path.join(WORK, "synth")
REF = os.path.join(WORK, "ref")
FFMPEG = shutil.which("ffmpeg") or "ffmpeg"

# the WzSound modules O1 ported (opusfile, oggopus.cpp, codecs.h)
O1_MODULES = ["opusfile_h.bas", "opusfile_internal_c.bas", "opusfile_info_c.bas", "opusfile_stream_c.bas",
              "opusfile_c.bas", "oggopus_c.bas", "WZOpusDecoder.cls", "WZDecoder.cls",
              "WZAudioDataResourceInterface.cls", "WzAudioDataPhysFSHandle.cls", "WZAudioDataMemoryBuffer.cls"]

# a float sample's tolerance against the official build's libopus 1.5.2 (the CELT float paths round a little
# differently from its SSE build; SILK is fixed-point: exact), against ffmpeg's (libopus 1.6.1: its concealment
# and mode switches differ, so only used without the DLL), and a 16-bit one against ffmpeg's 16-bit decode
# (dither, OP_GAIN's 32753/32768, rounding)
TOL_FLOAT = 1e-5
TOL_FFMPEG = 1e-3
TOL_S16 = 40
# op_read_stereo against the official build's 16-bit output (libopus 1.5.2's floats through the same filter): one
# float difference that flips a rounding changes the noise shaper's state (an IIR filter with about 20 dB of
# gain), so the two outputs then differ by a few units for a while
TOL_S16_OFFICIAL = 16

# ff_vorbis_channel_layout_offsets: ffmpeg's libopus decoder puts Vorbis channel offsets[c] at its channel c
VORBIS_OFFSETS = {3: [0, 2, 1], 4: [0, 1, 2, 3], 5: [0, 2, 1, 3, 4], 6: [0, 2, 1, 5, 3, 4],
                  7: [0, 2, 1, 6, 5, 3, 4], 8: [0, 2, 1, 7, 5, 6, 3, 4]}

# float32 constants of opusfile.c's float-to-short filter and stereo downmix
OP_FCOEF_B = [2.2374, -0.7339, -0.1251, -0.6033]
OP_FCOEF_A = [0.9030, 0.0116, -0.5853, -0.2571]
OP_STEREO_DOWNMIX = {
    3: [(0.5858, 0.0), (0.4142, 0.4142), (0.0, 0.5858)],
    4: [(0.4226, 0.0), (0.0, 0.4226), (0.366, 0.2114), (0.2114, 0.336)],
    5: [(0.651, 0.0), (0.46, 0.46), (0.0, 0.651), (0.5636, 0.3254), (0.3254, 0.5636)],
    6: [(0.529, 0.0), (0.3741, 0.3741), (0.0, 0.529), (0.4582, 0.2645), (0.2645, 0.4582), (0.3741, 0.3741)],
    7: [(0.4553, 0.0), (0.322, 0.322), (0.0, 0.4553), (0.3943, 0.2277), (0.2277, 0.3943), (0.2788, 0.2788),
        (0.322, 0.322)],
    8: [(0.3886, 0.0), (0.2748, 0.2748), (0.0, 0.3886), (0.3366, 0.1943), (0.1943, 0.3366), (0.3366, 0.1943),
        (0.1943, 0.3366), (0.2748, 0.2748)]}


def decoder_modules():
    """O2's and O3's modules in src/sound (CELT, the range decoder, SILK, the libopus top level)"""
    out = []
    for n in sorted(os.listdir(SOUND)):
        low = n.lower()
        if not low.endswith((".bas", ".cls")):
            continue
        if low.startswith(("celt_", "silk_", "opus_")) or low == "entcode_c.bas":
            out.append(n)
    return out


def write_vbp(name, extra=None, fast=False, decoders=True):
    """src/opustests/<name>.vbp: its modules, src/shared, src/shared/ogg, the WzSound modules; extra = more
    module paths (absolute) in place of the decoders"""
    lines = ["Type=Exe", build.STDOLE]
    files = [os.path.join(FOLDER, n) for n in build.sources(FOLDER)]
    files += [os.path.join(build.SHARED, n) for n in build.sources(build.SHARED)]
    files += [os.path.join(build.SHARED_OGG, n) for n in build.sources(build.SHARED_OGG)]
    files += [os.path.join(SOUND, n) for n in O1_MODULES]
    if decoders:
        files += [os.path.join(SOUND, n) for n in decoder_modules()]
    files += extra or []
    for path in files:
        build.normalise(path)
        rel = os.path.relpath(path, FOLDER)
        if path.lower().endswith(".cls"):
            lines.append("Class=%s; %s" % (build.vb_name(path), rel))
        else:
            lines.append("Module=%s; %s" % (build.vb_name(path), rel))
    lines.append(build.VBP_TAIL.format(name=name, ext="exe", path=os.path.relpath(build.BIN, FOLDER),
                                       startup="Sub Main", compat="", bounds="-1" if fast else "0",
                                       optimisation="0" if fast else "2", threading=""))
    vbp = os.path.join(FOLDER, name + ".vbp")
    with open(vbp, "w", encoding="latin-1", newline="\r\n") as f:
        f.write("\n".join(lines))
    return vbp


def make(name, vbp, outdir=None):
    """VB6 /make into bin (or outdir); a /make that stops without a message is tried again"""
    outdir = outdir or build.BIN
    log = os.path.join(WORK, name + "_build.log")
    os.makedirs(WORK, exist_ok=True)
    text = ""
    for _ in range(3):
        if os.path.exists(log):
            os.remove(log)
        subprocess.run([build.VB6, "/make", vbp, "/outdir", outdir, "/out", log], cwd=FOLDER)
        text = open(log, encoding="latin-1").read().strip() if os.path.exists(log) else ""
        if text:
            break
    print(text or "%s: VB6 stopped without a message" % name)
    return "succeeded" in text


def build_exes():
    ok = make("OpusTests", write_vbp("OpusTests"))
    ok = make("OpusFast", write_vbp("OpusFast", fast=True)) and ok
    return ok


# ---- Ogg parsing (the reference's view of a file) ----

def ogg_pages(data):
    """(offset, header type, granule position, serial, page number, segment table, body) of each page"""
    pos, out = 0, []
    while pos + 27 <= len(data) and data[pos:pos + 4] == b"OggS":
        htype = data[pos + 5]
        gp = struct.unpack_from("<q", data, pos + 6)[0]
        serial, pageno = struct.unpack_from("<II", data, pos + 14)
        nseg = data[pos + 26]
        segs = list(data[pos + 27:pos + 27 + nseg])
        body = data[pos + 27 + nseg:pos + 27 + nseg + sum(segs)]
        out.append((pos, htype, gp, serial, pageno, segs, body))
        pos += 27 + nseg + sum(segs)
    return out


def ogg_packets(pages):
    """the packets of the first logical stream, in order"""
    out, cur = [], b""
    serial = pages[0][3] if pages else None
    for _, htype, gp, s, _, segs, body in pages:
        if s != serial:
            continue
        p = 0
        for v in segs:
            cur += body[p:p + v]
            p += v
            if v < 255:
                out.append(cur)
                cur = b""
    return out


def toc_summary(packets):
    """mode / bandwidth / frame size of the audio packets' TOC bytes, with counts"""
    hist = {}
    for p in packets[2:]:
        if not p:
            continue
        config = p[0] >> 3
        if config < 12:
            mode, bw = "SILK", ["NB", "MB", "WB"][config // 4]
            fs = [10, 20, 40, 60][config % 4]
        elif config < 16:
            mode, bw = "hybrid", ["SWB", "FB"][(config - 12) // 2]
            fs = [10, 20][config % 2]
        else:
            mode, bw = "CELT", ["NB", "WB", "SWB", "FB"][(config - 16) // 4]
            fs = [2.5, 5, 10, 20][config % 4]
        key = "%s %s %sms%s" % (mode, bw, fs, "" if not p[0] & 4 else " st")
        hist[key] = hist.get(key, 0) + 1
    return ", ".join("%s x%d" % (k, v) for k, v in sorted(hist.items()))


def head_of(packets):
    h = packets[0]
    return {"channels": h[9], "pre_skip": struct.unpack_from("<H", h, 10)[0],
            "gain": struct.unpack_from("<h", h, 16)[0], "family": h[18]}


# ---- references ----

def ffmpeg_decode(path, fmt, extra=()):
    """ffmpeg's libopus decode of path: 'flt' (float32) or 's16'; raw bytes"""
    os.makedirs(REF, exist_ok=True)
    key = os.path.relpath(path, ROOT).replace(os.sep, "_").replace(":", "") + "." + fmt + \
        ("".join(extra).replace(" ", "") if extra else "")
    out = os.path.join(REF, key)
    if not os.path.exists(out) or os.path.getmtime(out) < os.path.getmtime(path):
        cmd = [FFMPEG, "-v", "error", "-y", "-request_sample_fmt", fmt] + list(extra) + \
              ["-c:a", "libopus", "-i", path, "-f", "f32le" if fmt == "flt" else "s16le",
               "-acodec", "pcm_f32le" if fmt == "flt" else "pcm_s16le", out]
        r = subprocess.run(cmd, capture_output=True)
        if r.returncode != 0:
            return None
    with open(out, "rb") as f:
        return f.read()


def floats(b):
    a = array.array("f")
    a.frombytes(b)
    return a


def shorts(b):
    a = array.array("h")
    a.frombytes(b[:len(b) // 2 * 2])
    return a


def reorder_from_ffmpeg(ref, ch):
    """ffmpeg's channel order -> Vorbis order (opusfile's) for 3 to 8 channels"""
    if ch < 3:
        return ref
    offs = VORBIS_OFFSETS[ch]
    out = array.array("f", bytes(len(ref) * 4))
    for c in range(ch):
        out[offs[c]::ch] = ref[c::ch]
    return out


def max_diff(a, b, n=None):
    n = min(len(a), len(b)) if n is None else n
    import numpy as np
    x = np.frombuffer(a, dtype=np.float32 if a.typecode == "f" else np.int16, count=n).astype(np.float64)
    y = np.frombuffer(b, dtype=np.float32 if b.typecode == "f" else np.int16, count=n).astype(np.float64)
    if n == 0:
        return 0.0, 0
    d = np.abs(x - y)
    return float(d.max()), int((d > 0).sum())


def float2short_reference(src, ch, serialno, nsamples):
    """opusfile.c's op_float2short_stereo_filter for 1 or 2 channels in float32 (the soft clip omitted: only
    called on samples within [-1, 1]), from a fresh decoder state: the dither seeded with the serial number"""
    import numpy as np
    f = np.float32
    B = [f(v) for v in OP_FCOEF_B]
    A = [f(v) for v in OP_FCOEF_A]
    gain = f(32753.0)
    prng = f(1.0) / f(float(0xFFFFFFFF))
    seed = serialno & 0xFFFFFFFF
    mute = 65
    da = [f(0)] * (4 * ch)
    db = [f(0)] * (4 * ch)
    out = []
    lo, hi = f(-32768.0), f(32767.0)
    for i in range(nsamples):
        silent = True
        for ci in range(ch):
            s = f(src[ch * i + ci])
            silent = silent and s == 0
            s = f(s * gain)
            err = f(0)
            for j in range(4):
                err = f(err + f(f(B[j] * db[ci * 4 + j]) - f(A[j] * da[ci * 4 + j])))
            for j in (2, 1, 0):
                da[ci * 4 + j + 1] = da[ci * 4 + j]
            for j in (2, 1, 0):
                db[ci * 4 + j + 1] = db[ci * 4 + j]
            da[ci * 4] = err
            s = f(s - err)
            if mute > 16:
                r = f(0)
            else:
                seed = (seed * 96314165 + 907633515) & 0xFFFFFFFF
                r = f(f(seed) * prng)
                seed = (seed * 96314165 + 907633515) & 0xFFFFFFFF
                r = f(r - f(f(seed) * prng))
            x = f(s + r)
            x = x if x < hi else hi
            x = lo if lo > x else x
            si = int(np.rint(x))
            out.append(si)
            if mute > 16:
                db[ci * 4] = f(0)
            else:
                y = f(f(si) - s)
                y = y if y < f(1.5) else f(1.5)
                y = f(-1.5) if f(-1.5) > y else y
                db[ci * 4] = y
        mute += 1
        if not silent:
            mute = 0
    if ch == 1:
        out = [v for v in out for _ in (0, 1)]
    return array.array("h", out)


def downmix_reference(src, ch, nsamples):
    import numpy as np
    x = np.frombuffer(src, dtype=np.float32, count=nsamples * ch).reshape(nsamples, ch)
    out = np.zeros((nsamples, 2), dtype=np.float32)
    for ci in range(ch):
        for k in range(2):
            out[:, k] = (out[:, k] + (np.float32(OP_STEREO_DOWNMIX[ch][ci][k]) * x[:, ci]).astype(np.float32)).astype(np.float32)
    return array.array("f", out.astype(np.float32).tobytes())


# ---- the official build's libopus (1.5.2) ----

OFFICIAL_DLLS = os.path.join(ROOT, "ref", "official", "wz470", "bin")
_lib152 = None


def libopus152():
    """the official 4.7.0 build's libopus.dll (libopus 1.5.2, what 4.7.0 decodes with) through ctypes, or None.
    ffmpeg 8.1 links libopus 1.6.1, whose concealment and mode switches differ from 1.5.2's"""
    global _lib152
    if _lib152 is not None:
        return _lib152 or None
    import ctypes
    try:
        os.add_dll_directory(OFFICIAL_DLLS)
        lib = ctypes.CDLL(os.path.join(OFFICIAL_DLLS, "libopus.dll"))
        lib.opus_get_version_string.restype = ctypes.c_char_p
        if not lib.opus_get_version_string().startswith(b"libopus 1.5.2"):
            raise OSError("not 1.5.2")
    except (OSError, AttributeError):
        _lib152 = False
        return None
    lib.opus_multistream_decoder_create.restype = ctypes.c_void_p
    lib.opus_multistream_decoder_create.argtypes = [ctypes.c_int32, ctypes.c_int, ctypes.c_int, ctypes.c_int,
                                                    ctypes.c_char_p, ctypes.POINTER(ctypes.c_int)]
    lib.opus_multistream_decode_float.restype = ctypes.c_int
    lib.opus_multistream_decode_float.argtypes = [ctypes.c_void_p, ctypes.c_char_p, ctypes.c_int32,
                                                  ctypes.POINTER(ctypes.c_float), ctypes.c_int, ctypes.c_int]
    lib.opus_multistream_decoder_destroy.argtypes = [ctypes.c_void_p]
    lib.opus_multistream_decoder_ctl.restype = ctypes.c_int
    _lib152 = lib
    return lib


def packet_duration(p):
    """opus_packet_get_nb_frames * opus_packet_get_samples_per_frame at 48 kHz (0 for a bad TOC)"""
    if not p:
        return 0
    toc = p[0]
    if toc & 0x80:
        fs = (48000 << ((toc >> 3) & 3)) // 400
    elif (toc & 0x60) == 0x60:
        fs = 960 if toc & 8 else 480
    else:
        a = (toc >> 3) & 3
        fs = 2880 if a == 3 else (48000 << a) // 100
    c = toc & 3
    n = 1 if c == 0 else 2 if c != 3 else (p[1] & 0x3F if len(p) > 1 else 0)
    d = n * fs
    return d if 0 < d <= 5760 else 0


def ogg_stream_packets(pages):
    """the first logical stream's packets as libogg returns them: (packet, after a hole, page index); a gap in
    the page numbers drops the partial packet and the next page's continued segments"""
    serial = pages[0][3]
    out, cur, gap, expected = [], b"", False, None
    for pi, (_, htype, gp, s, pageno, segs, body) in enumerate(pages):
        if s != serial:
            continue
        if expected is not None and pageno != expected:
            cur, gap = b"", True
        expected = pageno + 1
        p, i = 0, 0
        if htype & 1 and not cur:
            while i < len(segs):
                p += segs[i]
                i += 1
                if segs[i - 1] < 255:
                    break
        while i < len(segs):
            cur += body[p:p + segs[i]]
            p += segs[i]
            if segs[i] < 255:
                out.append((cur, gap, pi))
                cur, gap = b"", False
            i += 1
    return out


def ref152_decode(pages, head):
    """the samples opusfile gives over libopus 1.5.2 (op_read_float): every packet decoded in order (no
    concealment), the pre-skip discarded, 80 ms discarded after a hole (opusfile's pre-roll after lost data), the
    last page end-trimmed to its granule position; float32 in Vorbis channel order, or None without the DLL"""
    import ctypes
    lib = libopus152()
    if lib is None:
        return None
    pk = ogg_stream_packets(pages)
    h = pk[0][0]
    ch = h[9]
    if h[18] == 0:
        streams, coupled, mapping = 1, ch - 1, bytes([0, 1][:ch])
    else:
        streams, coupled, mapping = h[19], h[20], h[21:21 + ch]
    err = ctypes.c_int(0)
    st = lib.opus_multistream_decoder_create(48000, ch, streams, coupled, mapping, ctypes.byref(err))
    if not st:
        return None
    if head["gain"]:
        lib.opus_multistream_decoder_ctl(ctypes.c_void_p(st), ctypes.c_int(4034), ctypes.c_int(head["gain"]))
    buf = (ctypes.c_float * (5760 * ch))()
    out = []
    discard = head["pre_skip"]
    audio = [x for x in pk[2:] if packet_duration(x[0]) > 0]
    # end trimming: the samples of the last page's packets past its granule position (from the previous page's)
    last_pi = audio[-1][2]
    prev_gp = None
    for x in reversed(audio):
        if x[2] != last_pi:
            prev_gp = pages[x[2]][2]
            break
    trim = 0
    if prev_gp is not None and pages[last_pi][1] & 4:
        trim = max(0, prev_gp + sum(packet_duration(x[0]) for x in audio if x[2] == last_pi) - pages[last_pi][2])
    for pkt, after_gap, pi in audio:
        if after_gap:
            discard = 80 * 48
        ret = lib.opus_multistream_decode_float(st, pkt, len(pkt), buf, 5760, 0)
        if ret < 0:
            break
        a = array.array("f", bytes(memoryview(buf).cast("B")[:ret * ch * 4]))
        k = min(discard, ret)
        discard -= k
        out.append(a[k * ch:])
    lib.opus_multistream_decoder_destroy(st)
    res = array.array("f")
    for a in out:
        res.extend(a)
    if trim:
        del res[len(res) - trim * ch:]
    return res


# ---- synthetic files ----

SYNTH_FILES = [
    # name, source ("speech" / "music" / "music51"), sample rate, channels, encoder options
    ("silk_nb_8k", "speech", 8000, 1, ["-application", "voip", "-b:a", "8k"]),
    ("silk_mb_12k", "speech", 12000, 1, ["-application", "voip", "-b:a", "12k"]),
    ("silk_wb_16k", "speech", 16000, 1, ["-application", "voip", "-b:a", "16k"]),
    ("silk_nb_10ms", "speech", 8000, 1, ["-application", "voip", "-b:a", "8k", "-frame_duration", "10"]),
    ("silk_wb_40ms", "speech", 16000, 1, ["-application", "voip", "-b:a", "14k", "-frame_duration", "40"]),
    ("silk_wb_60ms_st", "speech", 16000, 2, ["-application", "voip", "-b:a", "16k", "-frame_duration", "60"]),
    ("silk_wb_fec", "speech", 16000, 1, ["-application", "voip", "-b:a", "16k", "-packet_loss", "20", "-fec", "1"]),
    ("hybrid_swb", "speech", 24000, 1, ["-application", "voip", "-b:a", "24k"]),
    ("hybrid_fb", "speech", 48000, 1, ["-application", "voip", "-b:a", "32k"]),
    ("hybrid_fb_st", "speech", 48000, 2, ["-application", "voip", "-b:a", "36k", "-frame_duration", "10"]),
    ("celt_64k_st", "music", 48000, 2, ["-application", "audio", "-b:a", "64k"]),
    ("celt_128k_mono", "music", 48000, 1, ["-application", "audio", "-b:a", "128k"]),
    ("celt_256k_st", "music", 48000, 2, ["-application", "audio", "-b:a", "256k"]),
    ("celt_2.5ms", "music", 48000, 2, ["-application", "lowdelay", "-b:a", "128k", "-frame_duration", "2.5"]),
    ("celt_5ms", "music", 48000, 2, ["-application", "lowdelay", "-b:a", "96k", "-frame_duration", "5"]),
    ("celt_10ms", "music", 48000, 2, ["-application", "audio", "-b:a", "96k", "-frame_duration", "10"]),
    ("celt_40ms", "music", 48000, 2, ["-application", "audio", "-b:a", "96k", "-frame_duration", "40"]),
    ("celt_60ms", "music", 48000, 2, ["-application", "audio", "-b:a", "96k", "-frame_duration", "60"]),
    ("celt_vbr_lowrate", "music", 48000, 2, ["-application", "audio", "-b:a", "32k", "-vbr", "constrained"]),
    ("surround51", "music51", 48000, 6, ["-application", "audio", "-b:a", "256k", "-mapping_family", "1"]),
]


def synth_sources():
    """the encoder inputs (WAV, cached): a voice briefing (speech) and 8 s of the menu music"""
    os.makedirs(SYNTH, exist_ok=True)
    speech = os.path.join(SYNTH, "speech.wav")
    if not os.path.exists(speech):
        subprocess.run([FFMPEG, "-v", "error", "-y", "-i", os.path.join(ROOT, "data", "base", "sequenceaudio", "np1.ogg"),
                        "-t", "10", "-ar", "48000", "-ac", "2", speech], check=True)
    music = os.path.join(SYNTH, "music.wav")
    if not os.path.exists(music):
        subprocess.run([FFMPEG, "-v", "error", "-y", "-ss", "40", "-t", "8", "-i", os.path.join(ROOT, "data", "music", "menu.opus"),
                        "-ar", "48000", "-ac", "2", music], check=True)
    music51 = os.path.join(SYNTH, "music51.wav")
    if not os.path.exists(music51):
        subprocess.run([FFMPEG, "-v", "error", "-y", "-i", music, "-af",
                        "pan=5.1|FL=FL|FR=FR|FC=0.5*FL+0.5*FR|LFE=0.4*FL-0.2*FR|BL=0.7*FL|BR=0.6*FR",
                        music51], check=True)
    return {"speech": speech, "music": music, "music51": music51}


def synth_files():
    src = synth_sources()
    out = []
    for name, kind, rate, ch, opts in SYNTH_FILES:
        path = os.path.join(SYNTH, name + ".opus")
        if not os.path.exists(path):
            cmd = [FFMPEG, "-v", "error", "-y", "-i", src[kind], "-ar", str(rate), "-ac", str(ch),
                   "-c:a", "libopus"] + opts + [path]
            r = subprocess.run(cmd, capture_output=True, text=True)
            if r.returncode != 0:
                print("opus: ffmpeg could not make %s: %s" % (name, r.stderr.strip()[:200]))
                continue
        out.append(path)
    # a page missing: the middle page of celt_64k_st dropped (a hole in the page sequence)
    lossy = os.path.join(SYNTH, "celt_64k_st_page_lost.opus")
    base = os.path.join(SYNTH, "celt_64k_st.opus")
    if os.path.exists(base) and not os.path.exists(lossy):
        data = open(base, "rb").read()
        pages = ogg_pages(data)
        k = len(pages) // 2
        cut0 = pages[k][0]
        cut1 = pages[k + 1][0]
        open(lossy, "wb").write(data[:cut0] + data[cut1:])
    if os.path.exists(lossy):
        out.append(lossy)
    return out


def music_files(all_):
    menu = os.path.join(ROOT, "data", "music", "menu.opus")
    albums = os.path.join(ROOT, "data", "music", "albums")
    tracks = []
    for d in sorted(os.listdir(albums)):
        p = os.path.join(albums, d)
        if os.path.isdir(p):
            tracks += [os.path.join(p, n) for n in sorted(os.listdir(p)) if n.lower().endswith(".opus")]
    if all_:
        return [menu] + tracks
    pick = [t for t in tracks if t.endswith(os.path.join("original_soundtrack", "track1.opus"))]
    pick += [t for t in tracks if t.endswith(os.path.join("aftermath_soundtrack", "track19.opus"))]
    if len(pick) < 2:
        pick = tracks[:2]
    return [menu] + pick


# ---- running ----

def run_exe(exe, path, outdir, noseek=False, env=None):
    if os.path.exists(outdir):
        shutil.rmtree(outdir)
    os.makedirs(outdir)
    lst = os.path.join(outdir, "list.txt")
    with open(lst, "w", encoding="latin-1") as f:
        f.write(path + "\n")
    args = [exe, lst, outdir] + (["noseek"] if noseek else [])
    subprocess.run(args, timeout=3600, env=env)
    res = {}
    rpath = os.path.join(outdir, "results.txt")
    if os.path.exists(rpath):
        for line in open(rpath, encoding="latin-1"):
            parts = line.rstrip("\r\n").split("\t")
            if parts[0] == "0":
                for p in parts[1:]:
                    k, _, v = p.partition("=")
                    res[k] = v
            elif parts[0] == "done":
                res["done"] = "1"
                for p in parts[1:]:
                    k, _, v = p.partition("=")
                    res["done_" + k] = v
    return res


def read_out(outdir, ext):
    p = os.path.join(outdir, "0" + ext)
    return open(p, "rb").read() if os.path.exists(p) else None


def check_file(path, exe, fast_exe, keep, fake=False):
    """one file: returns (passed, report lines, decode seconds, audio seconds)"""
    import numpy as np
    name = os.path.relpath(path, ROOT)
    data = open(path, "rb").read()
    pages = ogg_pages(data)
    packets = ogg_packets(pages)
    head = head_of(packets)
    ch = head["channels"]
    lines = ["%s: %s" % (name, toc_summary(packets))]
    problems = []
    outdir = os.path.join(WORK, "out")
    env = None
    if fake:
        env = dict(os.environ, OPUSFAKE_RAW=os.path.join(REF, "fake_raw.f32"))
        raw = ffmpeg_decode(path, "flt", ("-flags2", "skip_manual"))
        open(env["OPUSFAKE_RAW"], "wb").write(raw)
    res = run_exe(exe, path, outdir, noseek=fake, env=env)
    if "error" in res or "open_error" in res or res.get("done") != "1":
        lines.append("  FAILED: %s" % res)
        return False, lines, 0, 0
    ref = ffmpeg_decode(path, "flt")
    ref_s16 = ffmpeg_decode(path, "s16")
    if ref is None or ref_s16 is None:
        lines.append("  ffmpeg could not decode it")
        return False, lines, 0, 0
    reff = reorder_from_ffmpeg(floats(ref), ch)
    ref_n = len(reff) // ch
    vf = floats(read_out(outdir, ".f32") or b"")
    n = len(vf) // ch
    pcm_total = float(res.get("pcm_total", "-1"))
    lost = "page_lost" in name
    r152 = None if fake else ref152_decode(pages, head)
    hole_at = None
    if lost:
        # the samples before the hole: up to the granule position of the page before the gap in the page numbers
        for i in range(1, len(pages)):
            if pages[i][4] != pages[i - 1][4] + 1:
                hole_at = pages[i - 1][2] - head["pre_skip"]
                break
    dmax, ndiff = max_diff(vf, reff)
    lines.append("  samples %d (op_pcm_total %d, ffmpeg %d%s), %d channel(s), pre-skip %d, gain %s, links %s" % (
        n, pcm_total, ref_n, "" if r152 is None else ", 1.5.2 DLL %d" % (len(r152) // ch), ch, head["pre_skip"],
        head["gain"], res.get("links")))
    if r152 is not None:
        d152, n152 = max_diff(vf, r152)
        lines.append("  op_read_float: max difference from the official libopus 1.5.2 %.3g (%d of %d samples differ); "
                     "from ffmpeg's libopus 1.6.1 %.3g%s" % (d152, n152, len(vf), dmax,
                                                             " (after the hole: no concealment, 80 ms discarded)" if lost else ""))
        if n != len(r152) // ch or (not lost and n != pcm_total):
            problems.append("samples %d, op_pcm_total %d, 1.5.2 %d" % (n, pcm_total, len(r152) // ch))
        if d152 > TOL_FLOAT:
            problems.append("float differs from 1.5.2 by %.6g" % d152)
    else:
        lines.append("  op_read_float: max difference from ffmpeg's libopus %.3g (%d of %d samples differ)%s" % (
            dmax, ndiff, len(vf), "" if fake else " (the 1.5.2 DLL could not be loaded)"))
        if not lost and (n != pcm_total or n != ref_n):
            problems.append("samples %d, op_pcm_total %d, ffmpeg %d" % (n, pcm_total, ref_n))
        if not lost and dmax > (TOL_FLOAT if fake else TOL_FFMPEG):
            problems.append("float differs by %.6g" % dmax)
    if lost:
        holes = int(res.get("float_holes", "0"))
        before = max(0, min(n, ref_n, hole_at or 0))
        d2, _ = max_diff(vf, reff, before * ch)
        lines.append("  page lost: %d hole(s) reported (OP_HOLE) at sample %s; the %d samples before it: max difference "
                     "from ffmpeg %.3g" % (holes, hole_at, before, d2))
        if holes < 1:
            problems.append("no hole reported")
        if r152 is None and d2 > TOL_FFMPEG:
            problems.append("differs before the hole by %.3g" % d2)
    # 16-bit
    vs = shorts(read_out(outdir, ".s16") or b"")
    rs = shorts(ref_s16)
    if ch > 2:
        rs = None                               # (ffmpeg's s16 keeps the channels: no stereo reference)
    else:
        if ch == 1:
            rs = array.array("h", [v for v in rs for _ in (0, 1)])
        d16, n16 = max_diff(vs, rs)
        lines.append("  op_read_stereo: %d samples; against ffmpeg's 16-bit decode: max difference %d (%d differ)"
                     % (len(vs) // 2, d16, n16))
        if not lost and (len(vs) // 2 != n or d16 > TOL_S16):
            problems.append("16-bit: %d samples, max difference %d" % (len(vs) // 2, d16))
        # the float-to-short filter against this script's float32 version, on the port's float output
        if not lost:
            secs = 2
            m = min(n, 48000 * secs)
            src = vf[:m * ch]
            if max((abs(v) for v in src), default=0) <= 1.0:
                serial = struct.unpack("<I", struct.pack("<i", int(res.get("serialno", "0"))))[0]
                exp = float2short_reference(src, ch, serial, m)
                dd, nd = max_diff(vs, exp, m * 2)
                lines.append("  op_read_stereo's dither and noise shaping against a float32 version (first %d s): "
                             "%s" % (secs, "identical" if dd == 0 else "max difference %d (%d differ)" % (dd, nd)))
                if dd != 0:
                    problems.append("float-to-short filter differs by %d" % dd)
                if r152 is not None:
                    exp = float2short_reference(r152[:m * ch], ch, serial, m)
                    dd, nd = max_diff(vs, exp, m * 2)
                    lines.append("  op_read_stereo against the official build's (libopus 1.5.2 and the same filter; "
                                 "first %d s): %s" % (secs, "identical" if dd == 0 else "max difference %d (%d differ)" % (dd, nd)))
                    if dd > TOL_S16_OFFICIAL:
                        problems.append("16-bit differs from the official build's by %d" % dd)
            else:
                lines.append("  (the float output exceeds 1.0 in the first seconds: soft clipping, no filter check)")
    # downmix
    if ch > 2:
        f2 = floats(read_out(outdir, ".f2") or b"")
        exp = downmix_reference(vf, ch, n)
        dd, nd = max_diff(f2, exp)
        lines.append("  op_read_float_stereo (downmix): %d samples, max difference from a float32 downmix of the "
                     "port's own channels %.3g" % (len(f2) // 2, dd))
        if len(f2) // 2 != n or dd > 0:
            problems.append("downmix differs by %.3g" % dd)
        exp = downmix_reference(reff, ch, ref_n)
        dd, _ = max_diff(f2, exp)
        lines.append("  ... from ffmpeg's channels downmixed: %.3g" % dd)
        if r152 is not None:
            exp = downmix_reference(r152, ch, len(r152) // ch)
            dd, _ = max_diff(f2, exp)
            lines.append("  ... from libopus 1.5.2's channels downmixed: %.3g" % dd)
            if dd > TOL_FLOAT:
                problems.append("downmix differs from 1.5.2's by %.3g" % dd)
    # WZOpusDecoder
    if ch <= 2:
        wz = read_out(outdir, ".wz") or b""
        wzm = read_out(outdir, ".wzm") or b""
        s16b = read_out(outdir, ".s16") or b""
        # 4.7.0's SEEK_END callback makes opusfile end the stream one byte early: the last page is not read, the
        # stream ends with the last page before it that has a granule position. A hole ends it (decode's nullopt).
        timed = [p for p in pages[:-1] if p[2] not in (0, -1) and p[3] == pages[0][3]]
        exp_n = None
        if timed and not lost:
            exp_n = timed[-1][2] - head["pre_skip"]
        elif lost and hole_at is not None:
            exp_n = hole_at
        same = wz == s16b[:len(wz)]
        lines.append("  WZOpusDecoder: %d samples (PhysFS stream), %d (memory buffer)%s; %s op_read_stereo's; "
                     "totalTime %s, channels %s, frequency %s" % (
                         len(wz) // 4, len(wzm) // 4,
                         "" if exp_n is None else (", the hole at %d" % exp_n if lost else
                                                   ", expected %d (the last page's %d samples are not read)" % (exp_n, n - exp_n)),
                         "the same bytes as" if same else "NOT the same bytes as",
                         res.get("wz_totalTime"), res.get("wz_channels"), res.get("wz_frequency")))
        if lost:
            # a hole: op_read_stereo's OP_HOLE is decode's nullopt (the stream stops there)
            bad = not same or wz != wzm or exp_n is None or len(wz) // 4 > exp_n
        else:
            bad = not same or wz != wzm or (exp_n is not None and len(wz) // 4 != exp_n)
        if bad:
            problems.append("WZOpusDecoder")
    # seeks
    if not fake and "seek0" in res:
        # after a seek the decoder state differs from the straight decode's (80 ms pre-roll), so the position is
        # checked by the offset that best matches the second half of the 4800 samples read (0 expected); a seek
        # to 0 starts the decoder afresh: identical
        seek_ok = True
        txt = []
        x = np.frombuffer(vf, dtype=np.float32)
        for k in range(5):
            t, ret, tell = res["seek%d" % k].split("/")
            t, ret, tell = int(float(t)), int(ret), int(float(tell))
            sv = np.frombuffer(floats(read_out(outdir, ".seek%d.f32" % k) or b""), dtype=np.float32)
            d0 = best = None
            if len(sv) == 4800 * ch:
                tail = sv[2400 * ch:]
                diffs = {}
                for off in range(-40, 41):
                    a = (t + off + 2400) * ch
                    if a >= 0 and a + 2400 * ch <= len(x):
                        diffs[off] = float(np.abs(tail - x[a:a + 2400 * ch]).max())
                d0 = diffs.get(0)
                best = min(diffs, key=lambda o: (diffs[o], abs(o))) if diffs else None
            m2 = min(len(sv), max(0, len(x) - t * ch))
            first = float(np.abs(sv[:m2] - x[t * ch:t * ch + m2]).max()) if m2 else -1
            txt.append("%d -> %d %s, offset %s, %.3g" % (t, tell, "ok" if ret == 0 else "error %d" % ret, best, first))
            good = ret == 0 and tell == t and d0 is not None and d0 <= diffs[best] and (k != 4 or first == 0)
            if not good and not lost:
                seek_ok = False
        lines.append("  op_pcm_seek (target -> op_pcm_tell, the best matching offset, the 4800 samples' max difference "
                     "from the straight decode): " + "; ".join(txt) + "; op_raw_seek (return/samples/op_pcm_tell): " +
                     res.get("raw_seek", "?") + ("  (informative: the page is missing)" if lost else ""))
        if not seek_ok:
            problems.append("seek")
    # speed: the fast build
    secs = n / 48000.0
    dsecs = float(res.get("float_ms", "0")) / 1000
    if fast_exe and os.path.exists(fast_exe) and not fake:
        resf = run_exe(fast_exe, path, outdir, noseek=True)
        if "float_ms" in resf:
            dsecs = float(resf["float_ms"]) / 1000
            lines.append("  speed: %.1fx real time (float; fast build), %.1fx (op_read_stereo), test build %.1fx" % (
                secs / max(dsecs, 1e-9), secs / max(float(resf.get("s16_ms", "0")) / 1000, 1e-9),
                secs / max(float(res.get("float_ms", "0")) / 1000, 1e-9)))
    if res.get("done_asserts", "0") != "0":
        problems.append("%s ASSERT_FAILUREs (log.txt)" % res["done_asserts"])
    if problems:
        lines.append("  PROBLEMS: " + "; ".join(problems))
    if keep:
        kd = os.path.join(WORK, "kept", os.path.splitext(os.path.basename(path))[0])
        if os.path.exists(kd):
            shutil.rmtree(kd)
        shutil.copytree(outdir, kd)
    return not problems, lines, dsecs, secs


def run(keep=False, all_=False, fake=None, only=None, nobuild=False):
    """fake: a list of module paths standing in for O2's and O3's (a stand-in decoder replaying ffmpeg's raw
    decode): tests the Ogg and opusfile layers alone"""
    if not shutil.which(FFMPEG):
        print("opus: ffmpeg not found (the reference decoder)")
        return False
    os.makedirs(WORK, exist_ok=True)
    if fake:
        ok = nobuild or make("OpusTests", write_vbp("OpusTests", extra=fake, decoders=False))
        fast_exe = None
    else:
        if not decoder_modules():
            print("opus: O2's and O3's decoder modules are not in src/sound yet: not run")
            return False
        ok = nobuild or build_exes()
        fast_exe = os.path.join(build.BIN, "OpusFast.exe")
    if not ok:
        print("opus: OpusTests.exe did not build")
        return False
    exe = os.path.join(build.BIN, "OpusTests.exe")
    files = music_files(all_) + synth_files()
    if only:
        files = [f for f in files if only in os.path.basename(f)]
    passed, total_d, total_s = 0, 0.0, 0.0
    for path in files:
        good, lines, d, s = check_file(path, exe, fast_exe, keep, fake=bool(fake))
        passed += good
        total_d += d
        total_s += s
        for l in lines:
            print(l)
    print("opus: %d of %d files decode as %s does%s" % (
        passed, len(files), "ffmpeg's libopus" if fake or libopus152() is None else "the official build's libopus 1.5.2",
        "" if not total_d else "; %.1f s of audio in %.1f s (%.1fx real time, fast build)" % (
            total_s, total_d, total_s / total_d)))
    if not keep and os.path.exists(os.path.join(WORK, "out")):
        shutil.rmtree(os.path.join(WORK, "out"))
    return passed == len(files)


if __name__ == "__main__":
    only = [a.split("=", 1)[1] for a in sys.argv if a.startswith("--only=")]
    sys.exit(0 if run("--keep" in sys.argv, "--all" in sys.argv, only=only[0] if only else None,
                      nobuild="--nobuild" in sys.argv) else 1)
