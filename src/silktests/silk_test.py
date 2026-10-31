"""
silk_test.py - fork O3's test of WzSound's libopus port (cluster K): the SILK decoder (src/sound/silk_*), the
top level (src/sound/opus_*) and O2's range decoder and CELT (entcode_c, celt_*), against ffmpeg's libopus.

  python -I src\\silktests\\silk_test.py [--keep] [--nobuild] [--fast]

* builds bin\\SilkTests.exe (src/silktests + the decoder modules of src/sound + src/shared's float_vb and
  math_vb; bounds checks on, no optimisation) and bin\\SilkTestsFast.exe (the same, optimised without bounds
  checks, as WzSound is built) with tools/build.py's write_vbp and make
* makes test streams with ffmpeg's libopus encoder from the game's tutorial speech (data/base/audio/tutorial):
  SILK-only (voip, 8-16 kbps, 8/12/16 kHz input, 10 to 60 ms frames, stereo, inband FEC data) and hybrid
  (24-32 kbps speech at 24/48 kHz input)
* extracts the Opus packets (the Ogg pages' packets after the two headers) into packet files
* SilkTests.exe decodes them with opus_decode_float (and opus_multistream_decode_float, opus_decode); the
  float samples after the pre-skip are compared with ffmpeg's libopus decode (-c:a libopus -request_sample_fmt
  flt: libopusdec decodes to s16 unless asked), the 16-bit ones with ffmpeg's s16 decode (soft clipped)
* the reference that counts is the official 4.7.0 libopus.dll (libopus 1.5.2, a float build:
  ref/official/wz470/bin, 64-bit) through ctypes (ffmpeg 8.1 links libopus 1.6.1): every stream as is, with lost
  packets (PLC) and FEC decoding, at the other output rates (8-24 kHz: the SILK resampler's other paths), and a
  5.1 file through the multistream decoder (mapping family 1)
* SILK-only streams must be bit exact everywhere (the SILK decoder is integer code); with CELT (O2's float code)
  the float rounding differs a little (1e-6), and CELT's packet loss concealment amplifies it (reported apart)
run(keep=False) -> bool prints "SILK/Opus: N of M files ..." and the details.
"""
import ctypes
import glob
import os
import shutil
import struct
import subprocess
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, os.path.join(ROOT, "tools"))
import build  # noqa: E402

SOUND = os.path.join(ROOT, "src", "sound")
WORK = os.path.join(ROOT, "build", "silktests")
FFMPEG = "ffmpeg"
DLL = os.path.join(ROOT, "ref", "official", "wz470", "bin", "libopus.dll")

# name, ffmpeg options (after -i source), description
STREAMS = [
    ("silk_nb_8k", ["-ac", "1", "-ar", "8000", "-application", "voip", "-b:a", "8k"], "SILK NB 8 kbps"),
    ("silk_mb_12k", ["-ac", "1", "-ar", "12000", "-application", "voip", "-b:a", "12k"], "SILK MB 12 kbps"),
    ("silk_wb_16k", ["-ac", "1", "-ar", "16000", "-application", "voip", "-b:a", "16k"], "SILK WB 16 kbps"),
    ("silk_wb_10ms", ["-ac", "1", "-ar", "16000", "-application", "voip", "-b:a", "14k", "-frame_duration", "10"],
     "SILK WB 10 ms frames"),
    ("silk_wb_40ms", ["-ac", "1", "-ar", "16000", "-application", "voip", "-b:a", "12k", "-frame_duration", "40"],
     "SILK WB 40 ms frames"),
    ("silk_mb_60ms", ["-ac", "1", "-ar", "12000", "-application", "voip", "-b:a", "10k", "-frame_duration", "60"],
     "SILK MB 60 ms frames"),
    ("silk_nb_cbr", ["-ac", "1", "-ar", "8000", "-application", "voip", "-b:a", "9k", "-vbr", "off"],
     "SILK NB CBR"),
    ("silk_wb_fec", ["-ac", "1", "-ar", "16000", "-application", "voip", "-b:a", "16k", "-fec", "1",
                     "-packet_loss", "15"], "SILK WB with inband FEC (LBRR)"),
    ("silk_wb_stereo", ["-ac", "2", "-ar", "16000", "-application", "voip", "-b:a", "20k"], "SILK WB stereo"),
    ("silk_nb_stereo", ["-ac", "2", "-ar", "8000", "-application", "voip", "-b:a", "12k"], "SILK NB stereo"),
    ("hybrid_swb_24k", ["-ac", "1", "-ar", "24000", "-application", "voip", "-b:a", "24k"], "hybrid SWB 24 kbps"),
    ("hybrid_fb_32k", ["-ac", "1", "-ar", "48000", "-application", "voip", "-b:a", "32k"], "hybrid FB 32 kbps"),
    ("hybrid_fb_stereo", ["-ac", "2", "-ar", "48000", "-application", "voip", "-b:a", "32k"], "hybrid FB stereo"),
    ("mixed_fb_24k", ["-ac", "1", "-ar", "48000", "-application", "audio", "-b:a", "24k"],
     "speech, audio 24 kbps"),
    ("switch_32k", ["-ac", "1", "-ar", "48000", "-application", "audio", "-b:a", "32k"],
     "speech and music, audio 32 kbps (mode switches)"),
    ("switch_20k_10ms", ["-ac", "1", "-ar", "48000", "-application", "audio", "-b:a", "20k", "-frame_duration", "10"],
     "speech and music, audio 20 kbps, 10 ms (mode switches)"),
    ("switch_stereo_48k", ["-ac", "2", "-ar", "48000", "-application", "audio", "-b:a", "48k"],
     "speech and music stereo, audio 48 kbps (mode switches)"),
    ("surround51", ["-af", "pan=5.1|FL=c0|FR=c1|FC=0.5*c0+0.5*c1|LFE=0.3*c0|BL=0.7*c1|BR=0.7*c0", "-ar", "48000",
                    "-application", "audio", "-b:a", "160k"], "5.1 (mapping family 1, 6 channels in 4 streams)"),
]
SWITCH = ("switch_32k", "switch_20k_10ms", "switch_stereo_48k")
# the decoders' other output rates (SILK's resampler: 16 kHz internal -> 8 (down_FIR 1:2), 12 (down_FIR 3:4), 16
# (copy), 24 (IIR_FIR); 12 -> 8 (down_FIR 2:3), 24 (up2_HQ); 8 -> 12 (IIR_FIR), 16 (up2_HQ)), and with losses
RATES = {"silk_wb_16k": (8000, 12000, 16000, 24000), "silk_mb_12k": (8000, 16000, 24000),
         "silk_nb_8k": (12000, 16000), "silk_wb_stereo": (24000, 8000), "hybrid_fb_32k": (24000, 16000),
         "switch_32k": (16000,)}
LOSS_RATES = {"silk_wb_16k": (12000,), "silk_wb_stereo": (8000,), "switch_32k": (16000,)}


def run_cmd(cmd, **kw):
    r = subprocess.run(cmd, capture_output=True, **kw)
    if r.returncode != 0:
        print("FAILED: %s\n%s" % (" ".join(cmd), r.stderr.decode("latin-1")[-2000:]))
    return r.returncode == 0


def vb_name(path):
    with open(path, "r", encoding="latin-1") as f:
        for line in f:
            if line.startswith("Attribute VB_Name"):
                return line.split('"')[1]
    return os.path.splitext(os.path.basename(path))[0]


def decoder_modules():
    """the libopus port's modules in src/sound: O3's silk_* and opus_* (not O1's opusfile_*), O2's entcode_c
    and celt_*"""
    out = []
    for pat in ("silk_*.bas", "opus_*.bas", "entcode_c.bas", "celt_*.bas"):
        out += sorted(glob.glob(os.path.join(SOUND, pat)))
    return out


def build_exe(name, fast):
    vbp = build.write_vbp(name, "exe", HERE, [], False, None)
    with open(vbp, "r", encoding="latin-1") as f:
        lines = f.read().split("\n")
    keep = ("float_vb.bas", "math_vb.bas")
    lines = [l for l in lines if "..\\shared\\" not in l or l.rstrip().endswith(keep)]
    k = next(i for i, l in enumerate(lines) if l.startswith("Startup="))
    for path in decoder_modules():
        build.normalise(path)
        lines.insert(k, "Module=%s; %s" % (vb_name(path), os.path.relpath(path, HERE)))
        k += 1
    if fast:
        lines = [("OptimizationType=0" if l.startswith("OptimizationType=") else
                  "BoundsCheck=-1" if l.startswith("BoundsCheck=") else l) for l in lines]
    with open(vbp, "w", encoding="latin-1", newline="\r\n") as f:
        f.write("\n".join(lines))
    ok = build.make(name, HERE, vbp)
    return ok


def ogg_packets(path):
    """the packets of an Ogg file's first logical stream (page lacing reassembled)"""
    with open(path, "rb") as f:
        b = f.read()
    pos, packets, cur = 0, [], b""
    last_granule = None
    while pos + 27 <= len(b) and b[pos:pos + 4] == b"OggS":
        granule = struct.unpack_from("<q", b, pos + 6)[0]
        nseg = b[pos + 26]
        lac = b[pos + 27:pos + 27 + nseg]
        p = pos + 27 + nseg
        for v in lac:
            cur += b[p:p + v]
            p += v
            if v < 255:
                packets.append(cur)
                cur = b""
        if granule != -1:
            last_granule = granule
        pos = p
    return packets, last_granule


def toc_samples(pkt, fs=48000):
    """opus_packet_get_nb_samples"""
    toc = pkt[0]
    if toc & 0x80:
        spf = (fs << ((toc >> 3) & 3)) // 400
    elif (toc & 0x60) == 0x60:
        spf = fs // 50 if toc & 0x08 else fs // 100
    else:
        a = (toc >> 3) & 3
        spf = fs * 60 // 1000 if a == 3 else (fs << a) // 100
    c = toc & 3
    n = 1 if c == 0 else 2 if c != 3 else pkt[1] & 0x3F
    return n * spf


def write_packet_file(path, channels, records, layout):
    """records: (data, decode_fec, frame_size); layout: (Fs, streams, coupled streams, mapping)"""
    fs, streams, coupled, mapping = layout
    mp = bytes(mapping) + b"\0" * (-len(mapping) % 4)
    out = [struct.pack("<iiiii", channels, len(records), fs, streams, coupled), mp]
    for data, fec, fsz in records:
        out.append(struct.pack("<iii", len(data), fec, fsz))
        out.append(data)
    with open(path, "wb") as f:
        f.write(b"".join(out))


def loss_records(packets, period=7, fs=48000):
    """every period-th packet lost: concealed (PLC), or recovered from the next packet's FEC (every other loss)"""
    recs = []
    for i, p in enumerate(packets):
        if i % period == period - 1 and i + 1 < len(packets):
            n = toc_samples(p, fs)
            if (i // period) % 2 == 0:
                recs.append((b"", 0, n))
            else:
                recs.append((packets[i + 1], 1, n))
        else:
            recs.append((p, 0, 5760))
    return recs


def make_source(work):
    files = sorted(glob.glob(os.path.join(ROOT, "data", "base", "audio", "tutorial", "*.ogg")))
    lst = os.path.join(work, "concat.txt")
    with open(lst, "w") as f:
        for p in files:
            f.write("file '%s'\n" % p.replace("\\", "/"))
    mono = os.path.join(work, "speech_mono.wav")
    stereo = os.path.join(work, "speech_stereo.wav")
    if not os.path.exists(mono):
        if not run_cmd([FFMPEG, "-v", "error", "-y", "-f", "concat", "-safe", "0", "-i", lst, "-ac", "1", "-ar", "48000",
                        "-t", "45", "-c:a", "pcm_s16le", mono]):
            return None, None
    if not os.path.exists(stereo):
        if not run_cmd([FFMPEG, "-v", "error", "-y", "-i", mono, "-filter_complex",
                        "[0:a]asplit[a][b];[b]adelay=23,volume=0.7,highpass=f=300[c];[a][c]amerge=inputs=2[o]",
                        "-map", "[o]", "-c:a", "pcm_s16le", stereo]):
            return None, None
    return mono, stereo


def make_switch_source(work, mono):
    """speech and the menu music in turns (8 s each), so that the encoder switches between SILK, hybrid and CELT:
    the transitions, the redundant CELT frames and smooth_fade"""
    out = os.path.join(work, "speech_music.wav")
    music = os.path.join(ROOT, "data", "music", "menu.opus")
    if not os.path.exists(out) and os.path.exists(music):
        f = ("[0:a]atrim=0:8,asetpts=N/SR/TB[s0];[0:a]atrim=8:16,asetpts=N/SR/TB[s1];"
             "[0:a]atrim=16:24,asetpts=N/SR/TB[s2];"
             "[1:a]aresample=48000,pan=mono|c0=0.5*c0+0.5*c1,atrim=20:28,asetpts=N/SR/TB[m0];"
             "[1:a]aresample=48000,pan=mono|c0=0.5*c0+0.5*c1,atrim=60:68,asetpts=N/SR/TB[m1];"
             "[s0][m0][s1][m1][s2]concat=n=5:v=0:a=1[o]")
        if not run_cmd([FFMPEG, "-v", "error", "-y", "-i", mono, "-i", music, "-filter_complex", f, "-map", "[o]",
                        "-ac", "1", "-ar", "48000", "-c:a", "pcm_s16le", out]):
            return None
    return out if os.path.exists(out) else None


def floats(path):
    with open(path, "rb") as f:
        b = f.read()
    return struct.unpack("<%df" % (len(b) // 4), b)


def shorts(path):
    with open(path, "rb") as f:
        b = f.read()
    return struct.unpack("<%dh" % (len(b) // 2), b)


def max_diff(a, b, mask=None, inside=False):
    """the largest |a[i] - b[i]| and its index; with a mask, over the samples where mask[i] == inside"""
    n = min(len(a), len(b))
    m, at = 0.0, -1
    for i in range(n):
        if mask is not None and mask[i] != inside:
            continue
        d = abs(a[i] - b[i])
        if d > m:
            m, at = d, i
    return m, at


def toc_mode(d):
    t = d[0]
    return "CELT" if t & 0x80 else "hybrid" if (t & 0x60) == 0x60 else "SILK"


def prev_mode(recs, k):
    """the mode of the last packet decoded normally before record k (what the PLC continues)"""
    for j in range(k - 1, -1, -1):
        d, fec, fsz = recs[j]
        if d and not fec:
            return toc_mode(d)
    return "SILK"


def dll_decode(dll, channels, records, layout):
    """the official libopus.dll's opus_decode_float (opus_multistream_decode_float for more than one stream) over
    the records: float samples"""
    fs, streams, coupled, mapping = layout
    err = ctypes.c_int(0)
    multi = streams > 1
    if multi:
        dll.opus_multistream_decoder_create.restype = ctypes.c_void_p
        dll.opus_multistream_decoder_create.argtypes = [ctypes.c_int32, ctypes.c_int, ctypes.c_int, ctypes.c_int,
                                                        ctypes.c_char_p, ctypes.POINTER(ctypes.c_int)]
        dll.opus_multistream_decode_float.restype = ctypes.c_int
        dll.opus_multistream_decode_float.argtypes = [ctypes.c_void_p, ctypes.c_char_p, ctypes.c_int32,
                                                      ctypes.POINTER(ctypes.c_float), ctypes.c_int, ctypes.c_int]
        dll.opus_multistream_decoder_destroy.argtypes = [ctypes.c_void_p]
        st = dll.opus_multistream_decoder_create(fs, channels, streams, coupled, bytes(mapping), ctypes.byref(err))
        decode, destroy = dll.opus_multistream_decode_float, dll.opus_multistream_decoder_destroy
    else:
        dll.opus_decoder_create.restype = ctypes.c_void_p
        dll.opus_decoder_create.argtypes = [ctypes.c_int32, ctypes.c_int, ctypes.POINTER(ctypes.c_int)]
        dll.opus_decode_float.restype = ctypes.c_int
        dll.opus_decode_float.argtypes = [ctypes.c_void_p, ctypes.c_char_p, ctypes.c_int32, ctypes.POINTER(ctypes.c_float),
                                          ctypes.c_int, ctypes.c_int]
        dll.opus_decoder_destroy.argtypes = [ctypes.c_void_p]
        st = dll.opus_decoder_create(fs, channels, ctypes.byref(err))
        decode, destroy = dll.opus_decode_float, dll.opus_decoder_destroy
    if not st:
        return None
    out = []
    buf = (ctypes.c_float * (5760 * channels))()
    for data, fec, fsz in records:
        n = decode(st, data if data else None, len(data), buf, fsz, fec)
        if n < 0:
            destroy(st)
            return None
        out.extend(buf[:n * channels])
    destroy(st)
    return out


def run(keep=False, nobuild=False):
    os.makedirs(WORK, exist_ok=True)
    if not nobuild:
        t = time.time()
        if not build_exe("SilkTests", False):
            print("SILK/Opus: SilkTests.exe did not build")
            return False
        if not build_exe("SilkTestsFast", True):
            print("SILK/Opus: SilkTestsFast.exe did not build")
            return False
        print("built SilkTests.exe and SilkTestsFast.exe (%.0f s)" % (time.time() - t))
    mono, stereo = make_source(WORK)
    if mono is None:
        print("SILK/Opus: no source audio")
        return False
    dll = None
    try:
        dll = ctypes.CDLL(DLL)
    except OSError as e:
        print("(the official libopus.dll does not load: %s; no packet loss comparison)" % e)

    # (name, description, packet file, reference f32 (kind "dll": the records), reference s16, pre_skip, channels,
    #  kind, modes, layout)
    tests = []
    switch = make_switch_source(WORK, mono)
    for name, opts, desc in STREAMS:
        src = stereo if "-ac" in opts and opts[opts.index("-ac") + 1] == "2" else mono
        if name in SWITCH:
            if switch is None:
                continue
            src = switch
        if name == "surround51":
            src = stereo
        opus = os.path.join(WORK, name + ".opus")
        if not os.path.exists(opus):
            if not run_cmd([FFMPEG, "-v", "error", "-y", "-i", src] + opts + ["-c:a", "libopus"] + [opus]):
                continue
        packets, granule = ogg_packets(opus)
        head = packets[0]
        channels, pre_skip = head[9], struct.unpack_from("<H", head, 10)[0]
        if head[18] == 0:
            layout = (48000, 1, channels - 1, [0, 1][:channels])
        else:
            layout = (48000, head[19], head[20], list(head[21:21 + channels]))
        audio = packets[2:]
        modes = {}
        for p in audio:
            m = toc_mode(p)
            modes[m] = modes.get(m, 0) + 1
        if layout[1] > 1:
            # (ffmpeg reorders the channels of mapping family 1: only the DLL, with the same mapping)
            if dll is None:
                continue
            recs = [(p, 0, 5760) for p in audio]
            pf = os.path.join(WORK, name + ".pkt")
            write_packet_file(pf, channels, recs, layout)
            tests.append((name, desc, pf, recs, None, 0, channels, "dll", modes, layout))
            recs = loss_records(audio)
            lf = os.path.join(WORK, name + ".loss.pkt")
            write_packet_file(lf, channels, recs, layout)
            tests.append((name + "_loss", desc + ", every 7th packet lost (PLC / FEC)", lf, recs, None, 0, channels,
                          "dll", modes, layout))
            continue
        pf = os.path.join(WORK, name + ".pkt")
        write_packet_file(pf, channels, [(p, 0, 5760) for p in audio], layout)
        ref = os.path.join(WORK, name + ".ref.f32")
        ref16 = os.path.join(WORK, name + ".ref.s16")
        if not os.path.exists(ref):
            # (libopusdec decodes to s16 unless float is requested)
            run_cmd([FFMPEG, "-v", "error", "-y", "-c:a", "libopus", "-request_sample_fmt", "flt", "-i", opus,
                     "-f", "f32le", "-c:a", "pcm_f32le", ref])
        if not os.path.exists(ref16):
            run_cmd([FFMPEG, "-v", "error", "-y", "-c:a", "libopus", "-request_sample_fmt", "s16", "-i", opus,
                     "-f", "s16le", "-c:a", "pcm_s16le", ref16])
        tests.append((name, desc, pf, ref, ref16, pre_skip, channels, "ffmpeg", modes, layout))
        if dll is not None:
            recs = loss_records(audio)
            lf = os.path.join(WORK, name + ".loss.pkt")
            write_packet_file(lf, channels, recs, layout)
            tests.append((name + "_loss", desc + ", every 7th packet lost (PLC / FEC)", lf, recs, None, 0, channels,
                          "dll", modes, layout))
            # the other output rates (the SILK resampler's other paths, CELT's downsampling)
            for fs in RATES.get(name, ()):
                lay = (fs,) + layout[1:]
                recs = loss_records(audio, fs=fs) if fs in LOSS_RATES.get(name, ()) else [(p, 0, 5760) for p in audio]
                rf = os.path.join(WORK, "%s_%d.pkt" % (name, fs))
                write_packet_file(rf, channels, recs, lay)
                tests.append(("%s_%dk" % (name, fs // 1000), "%s decoded at %d Hz%s" % (
                    desc, fs, ", every 7th packet lost" if fs in LOSS_RATES.get(name, ()) else ""), rf, recs, None, 0,
                    channels, "dll", modes, lay))

    listfile = os.path.join(WORK, "list.txt")
    with open(listfile, "w") as f:
        for t in tests:
            f.write(t[2] + "\n")
    results = {}
    for exe in ("SilkTests", "SilkTestsFast"):
        outdir = os.path.join(WORK, exe)
        if os.path.isdir(outdir):
            shutil.rmtree(outdir)
        os.makedirs(outdir)
        t = time.time()
        r = subprocess.run([os.path.join(build.BIN, exe + ".exe"), listfile, outdir], capture_output=True)
        print("%s.exe: %.1f s, exit code %d" % (exe, time.time() - t, r.returncode))
        lines = {}
        rp = os.path.join(outdir, "results.txt")
        if os.path.exists(rp):
            with open(rp) as f:
                for line in f:
                    parts = line.rstrip("\n").split("\t")
                    lines[parts[0]] = parts
        results[exe] = (outdir, lines)

    passed = 0
    outdir, lines = results["SilkTests"]
    fastdir, fastlines = results["SilkTestsFast"]
    for i, (name, desc, pf, ref, ref16, pre_skip, channels, kind, modes, layout) in enumerate(tests):
        parts = lines.get(str(i))
        mstr = ", ".join("%s %d" % (k, v) for k, v in sorted(modes.items()))
        if parts is None or parts[1] == "error":
            print("  FAIL %-22s %s: %s" % (name, desc, " ".join(parts[2:]) if parts else "no result"))
            continue
        samples = int(parts[2])
        mine = floats(os.path.join(outdir, "%d.f32" % i))
        fast = floats(os.path.join(fastdir, "%d.f32" % i)) if os.path.exists(os.path.join(fastdir, "%d.f32" % i)) else ()
        silk_only = "SILK" in modes and len(modes) == 1
        # SILK-only output is integer arithmetic, bit exact everywhere; with CELT (O2's float code, compiled by VB
        # with unrounded floating point allowed) the builds and the official DLL differ by float rounding only,
        # except in the frames CELT conceals (lost or FEC on a CELT/hybrid packet) and the 40 ms after them, where
        # its PLC amplifies those differences (pitch search, LPC synthesis, then the overlap, the postfilter and the
        # energy prediction carry them on): masked out, their max reported apart. Outside, the CELT frames that
        # follow a concealment keep a decaying trace of it (tol: 1e-4; inside: 0.05; our own two builds differ there
        # by the same order).
        mask = None
        tol = 0 if silk_only else 1e-6
        if kind == "dll" and not silk_only:
            mask, after = [], 0
            for k, (d, fec, fsz) in enumerate(ref):
                n = toc_samples(d, layout[0]) if d and not fec else fsz
                concealed = not d or fec
                celt = concealed and (prev_mode(ref, k) != "SILK" or (fec and toc_mode(d) != "SILK"))
                for _ in range(n):
                    m = celt or after > 0
                    mask.extend([m] * channels)
                    if not celt and after > 0:
                        after -= 1
                if celt:
                    after = 1920
            if "CELT" in modes and any(mask):
                tol = 1e-4
        fmd, fat = max_diff(mine, fast, mask)
        fastok = len(fast) == len(mine) and fmd <= tol
        fparts = fastlines.get(str(i))
        fast_ms = float(fparts[3]) if fparts and fparts[1] != "error" else None
        secs = samples / float(layout[0])
        if kind == "ffmpeg":
            reff = floats(ref)
            mine_t = mine[pre_skip * channels:pre_skip * channels + len(reff)]
            md, at = max_diff(mine_t, reff)
            lendiff = len(mine) - pre_skip * channels - len(reff)
            r16 = shorts(ref16) if os.path.exists(ref16) else ()
            m16 = shorts(os.path.join(outdir, "%d.s16" % i))
            md16, at16 = max_diff(m16[pre_skip * channels:pre_skip * channels + len(r16)], r16)
            # the DLL (a float build) on the same packets
            dmd, dat = -1.0, -1
            if dll is not None:
                with open(pf, "rb") as f:
                    pdata = f.read()
                recs, p = [], 20 + (channels + 3) // 4 * 4
                for _ in range(struct.unpack_from("<i", pdata, 4)[0]):
                    ln, fec, fsz = struct.unpack_from("<iii", pdata, p)
                    recs.append((pdata[p + 12:p + 12 + ln], fec, fsz))
                    p += 12 + ln
                reff_dll = dll_decode(dll, channels, recs, layout)
                if reff_dll is not None and len(reff_dll) == len(mine):
                    dmd, dat = max_diff(mine, reff_dll)
            if silk_only:
                ok = md == 0 and md16 == 0 and dmd in (0, -1.0)
            else:
                ok = md <= 1e-6 and md16 <= 1 and dmd <= 1e-6
            ok = ok and lendiff >= 0 and parts[4] == "same" and fastok
            print("  %s %-22s %s [%s]: %d samples x %d ch; ffmpeg: max |diff| float %.3g%s (sample %d), 16-bit %d; "
                  "official libopus.dll: %s; multistream %s; optimised build: %s; %.1f s audio in %.0f ms (%.0fx real "
                  "time; checked build %.0f ms)" % (
                      "ok  " if ok else "FAIL", name, desc, mstr, len(reff) // channels, channels, md,
                      " (bit exact)" if md == 0 else "", at // channels, md16,
                      "max |diff| %.3g%s" % (dmd, " (bit exact)" if dmd == 0 else "") if dmd >= 0 else "-",
                      parts[4], "same" if fmd == 0 else "max |diff| %.3g" % fmd, secs, fast_ms or -1,
                      secs * 1000.0 / fast_ms if fast_ms else 0, float(parts[3])))
        else:
            reff = dll_decode(dll, channels, ref, layout)
            if reff is None:
                print("  FAIL %-22s the DLL failed" % name)
                continue
            md, at = max_diff(mine, reff, mask)
            extra = ""
            cmd = 0.0
            if mask is not None:
                cmd, cat = max_diff(mine, reff, mask, True)
                fcmd, fcat = max_diff(mine, fast, mask, True)
                extra = " (outside the %d CELT-concealed frames and the 40 ms after them; inside: max |diff| %.3g, " \
                    "optimised build %.3g)" % (sum(1 for k, r in enumerate(ref) if (not r[0] or r[1]) and (
                        prev_mode(ref, k) != "SILK" or (r[1] and toc_mode(r[0]) != "SILK"))), cmd, fcmd)
            ok = len(mine) == len(reff) and md <= tol and cmd <= 0.05 and fastok
            print("  %s %-22s %s: %d samples, official libopus.dll: max |diff| %.3g%s (sample %d)%s; optimised build: %s%s" % (
                "ok  " if ok else "FAIL", name, desc, len(mine) // channels, md, " (bit exact)" if md == 0 else "",
                at // channels, "" if len(mine) == len(reff) else ", %d vs %d samples" % (len(mine), len(reff)),
                "same" if fmd == 0 else "max |diff| %.3g" % fmd, extra))
        if ok:
            passed += 1
    speeds = {}
    for i, t in enumerate(tests):
        p = fastlines.get(str(i))
        if p is None or p[1] == "error" or t[7] != "ffmpeg":
            continue
        cat = "SILK-only" if "SILK" in t[8] and len(t[8]) == 1 else "with CELT"
        s, ms = speeds.get(cat, (0.0, 0.0))
        speeds[cat] = (s + int(p[2]) / 48000.0, ms + float(p[3]))
    print("SILK/Opus: %d of %d files as libopus decodes them; optimised build: %s" % (passed, len(tests), "; ".join(
        "%s %.0f s of audio in %.0f ms (%.0fx real time)" % (k, s, ms, s * 1000.0 / ms if ms else 0)
        for k, (s, ms) in sorted(speeds.items()))))
    if not keep:
        for exe in ("SilkTests", "SilkTestsFast"):
            shutil.rmtree(os.path.join(WORK, exe), ignore_errors=True)
    return passed == len(tests)


if __name__ == "__main__":
    ok = run(keep="--keep" in sys.argv, nobuild="--nobuild" in sys.argv)
    sys.exit(0 if ok else 1)
