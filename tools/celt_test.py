"""
celt_test.py - cluster K (fork O2)'s test of the port of libopus 1.5.2's CELT decoder and range decoder
(src/sound/entcode_c.bas, src/sound/celt_*.bas) before the Opus top level exists: builds bin/CeltTests.exe
(src/celttests + those modules, bounds checks on) and bin/CeltTestsFast.exe (the same compiled as WzSound is:
optimised, no bounds checks; for the speed), makes synthetic CELT-only Opus files with ffmpeg's libopus encoder
(-application lowdelay: mono and stereo, 2.5 / 5 / 10 / 20 / 40 / 60 ms frames, low and high bitrates, VBR and
CBR, the four CELT bandwidths, transients, digital silence) and copies of some with lost packets (a lost packet
is replaced by its TOC byte alone: a 0-byte frame, which libopus conceals as a lost one, data NULL), extracts
their packets from the Ogg pages, decodes them with CeltTests.exe (as opus_decode_native / opus_decode_frame do
for CELT-only packets) and compares with ffmpeg's libopus decode of the same file (float output, after the
pre-skip). ffmpeg 8.1's libopus is 1.6.1, whose packet loss concealment differs from 1.5.2's (the noise-based
PLC): the streams with losses are also decoded with the official Warzone 4.7.0 build's libopus.dll (1.5.2,
vcpkg, through ctypes; ref/official/wz470/bin) and judged against it. Prints "celt: N of M streams decode as
libopus does (...)", the max abs differences and the decode speed (x real time).

  python -I tools\\celt_test.py [--keep]          (run(keep) for run_tests.py)
"""
import array
import os
import shutil
import struct
import subprocess
import sys

import numpy as np

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "tools"))
import build  # noqa: E402

FOLDER = os.path.join(ROOT, "src", "celttests")
SOUND = os.path.join(ROOT, "src", "sound")
WORK = os.path.join(ROOT, "build", "celt_test")
FFMPEG = shutil.which("ffmpeg") or "ffmpeg"
RATE = 48000
# the largest difference from libopus' float output allowed (full scale 1.0): float rounding (VB's x87 against
# libopus' SSE build) gives about 1e-6; a decoding error gives differences of the size of the signal
TOLERANCE = 1e-4
# with lost packets the pitch-based PLC's LPC analysis (24th order) and IIR synthesis amplify the summation
# differences (libopus' SIMD sums against VB's unrounded x87 ones): about 2e-4 against libopus 1.5.2; a wrong pitch
# or a wrong concealment gives differences of the size of the signal
PLC_TOLERANCE = 1e-3
OFFICIAL_DLLS = os.path.join(ROOT, "ref", "official", "wz470", "bin")


def ref152(pkt_path, out_path):
    """decodes a packet file with the official 4.7.0 build's libopus.dll (libopus 1.5.2) through ctypes:
    opus_decoder_create(rate, channels), opus_decode_float per packet (a TOC-only packet: a lost frame, concealed).
    Writes the float samples (from the first one, the pre-skip included); False if the DLL can't be used."""
    import ctypes
    try:
        os.add_dll_directory(OFFICIAL_DLLS)
        lib = ctypes.CDLL(os.path.join(OFFICIAL_DLLS, "libopus.dll"))
        lib.opus_get_version_string.restype = ctypes.c_char_p
        if not lib.opus_get_version_string().startswith(b"libopus 1.5.2"):
            return False
    except (OSError, AttributeError):
        return False
    lib.opus_decoder_create.restype = ctypes.c_void_p
    lib.opus_decoder_create.argtypes = [ctypes.c_int32, ctypes.c_int, ctypes.POINTER(ctypes.c_int)]
    lib.opus_decode_float.restype = ctypes.c_int
    lib.opus_decode_float.argtypes = [ctypes.c_void_p, ctypes.c_char_p, ctypes.c_int32, ctypes.POINTER(ctypes.c_float),
                                      ctypes.c_int, ctypes.c_int]
    lib.opus_decoder_destroy.argtypes = [ctypes.c_void_p]
    with open(pkt_path, "rb") as f:
        data = f.read()
    channels, preskip, count, rate = struct.unpack_from("<iiii", data, 8)
    err = ctypes.c_int(0)
    st = lib.opus_decoder_create(rate, channels, ctypes.byref(err))
    if not st:
        return False
    buf = (ctypes.c_float * (5760 * channels))()
    out, p = [], 24
    for _ in range(count):
        n = struct.unpack_from("<i", data, p)[0]
        ret = lib.opus_decode_float(st, data[p + 4:p + 4 + n], n, buf, 5760, 0)
        p += 4 + n
        if ret < 0:
            lib.opus_decoder_destroy(st)
            return False
        out.append(bytes(memoryview(buf).cast("B")[:ret * channels * 4]))
    lib.opus_decoder_destroy(st)
    with open(out_path, "wb") as f:
        f.write(b"".join(out))
    return True


def build_exe(name, bounds_check):
    """bin/<name>.exe from src/celttests and the CELT modules (no DLL; src/shared is left out: the ported modules
    use none of its names, and its modules need names their projects define)"""
    lines = ["Type=Exe", build.STDOLE]
    for n in build.sources(FOLDER):
        path = os.path.join(FOLDER, n)
        build.normalise(path)
        lines.append("Module=%s; %s" % (build.vb_name(path), n))
    for n in sorted(os.listdir(SOUND)):
        if n.lower().endswith(".bas") and (n.lower().startswith("celt_") or n.lower() == "entcode_c.bas"):
            path = os.path.join(SOUND, n)
            build.normalise(path)
            lines.append("Module=%s; ..\\sound\\%s" % (build.vb_name(path), n))
    lines.append(build.VBP_TAIL.format(name=name, ext="exe", path=os.path.relpath(build.BIN, FOLDER), startup="Sub Main",
                                       compat="", bounds="0" if bounds_check else "-1", optimisation="0", threading=""))
    vbp = os.path.join(FOLDER, name + ".vbp")
    with open(vbp, "w", encoding="latin-1", newline="\r\n") as f:
        f.write("\n".join(lines))
    return build.make(name, FOLDER, vbp)


# --- test signals -------------------------------------------------------------------------------------------

def signal(seconds_each=1.5):
    """a stereo test signal (float, -1..1): music-like tones, clicks (transients), a sweep, digital silence,
    noise; the channels differ (stereo coding)"""
    rng = np.random.default_rng(1234)
    n = int(RATE * seconds_each)
    t = np.arange(n) / RATE
    parts_l, parts_r = [], []
    # music: harmonic tones with vibrato and an envelope
    env = 0.5 + 0.5 * np.sin(2 * np.pi * 1.3 * t)
    l = sum(0.25 / k * np.sin(2 * np.pi * 220 * k * t + 0.3 * np.sin(2 * np.pi * 5 * t)) for k in range(1, 9)) * env
    r = sum(0.25 / k * np.sin(2 * np.pi * 330 * k * t + 0.2 * np.sin(2 * np.pi * 4 * t)) for k in range(1, 7)) * env[::-1]
    parts_l.append(l + 0.01 * rng.standard_normal(n))
    parts_r.append(r + 0.01 * rng.standard_normal(n))
    # clicks: impulses and short bursts on a quiet tone
    l = 0.05 * np.sin(2 * np.pi * 150 * t)
    r = 0.05 * np.sin(2 * np.pi * 175 * t)
    for k in range(0, n, RATE // 8):
        l[k] += 0.9
        r[min(n - 1, k + 37)] -= 0.8
        burst = min(n - k, 200)
        l[k:k + burst] += 0.5 * rng.standard_normal(burst) * np.exp(-np.arange(burst) / 40.0)
    parts_l.append(l)
    parts_r.append(r)
    # sweep 50 Hz -> 20 kHz
    f0, f1 = 50.0, 20000.0
    phase = 2 * np.pi * f0 * seconds_each * ((f1 / f0) ** (t / seconds_each) - 1) / np.log(f1 / f0)
    parts_l.append(0.4 * np.sin(phase))
    parts_r.append(0.4 * np.cos(phase * 0.999))
    # digital silence
    parts_l.append(np.zeros(n // 3))
    parts_r.append(np.zeros(n // 3))
    # noise, then a loud tone
    parts_l.append(0.3 * rng.standard_normal(n // 2))
    parts_r.append(0.3 * rng.standard_normal(n // 2))
    parts_l.append(0.95 * np.sin(2 * np.pi * 1000 * t[:n // 2]))
    parts_r.append(-0.95 * np.sin(2 * np.pi * 1000 * t[:n // 2]))
    left = np.clip(np.concatenate(parts_l), -1, 1)
    right = np.clip(np.concatenate(parts_r), -1, 1)
    return np.stack([left, right], axis=1)


def write_wav(path, x):
    """16-bit PCM WAV of the float signal x (frames x channels)"""
    import wave
    pcm = np.round(x * 32767).astype("<i2")
    with wave.open(path, "wb") as w:
        w.setnchannels(x.shape[1])
        w.setsampwidth(2)
        w.setframerate(RATE)
        w.writeframes(pcm.tobytes())


# --- Ogg ---------------------------------------------------------------------------------------------------

def ogg_packets(data):
    """(packets, granules of the pages each packet ends on, serial) of a one-stream Ogg file"""
    pos, packets, cur, serial = 0, [], b"", None
    while pos + 27 <= len(data):
        if data[pos:pos + 4] != b"OggS":
            raise ValueError("not an Ogg page at %d" % pos)
        serial = struct.unpack_from("<I", data, pos + 14)[0]
        nseg = data[pos + 26]
        lacing = data[pos + 27:pos + 27 + nseg]
        p = pos + 27 + nseg
        for v in lacing:
            cur += data[p:p + v]
            p += v
            if v < 255:
                packets.append(cur)
                cur = b""
        pos = p
    return packets, serial


def last_granule(data):
    pos, granule = 0, None
    while pos + 27 <= len(data) and data[pos:pos + 4] == b"OggS":
        granule = struct.unpack_from("<q", data, pos + 6)[0]
        nseg = data[pos + 26]
        pos += 27 + nseg + sum(data[pos + 27:pos + 27 + nseg])
    return granule


def _crc_table():
    t = []
    for i in range(256):
        r = i << 24
        for _ in range(8):
            r = ((r << 1) ^ 0x04C11DB7) if r & 0x80000000 else (r << 1)
            r &= 0xFFFFFFFF
        t.append(r)
    return t


CRC_TABLE = _crc_table()


def ogg_crc(data):
    crc = 0
    for b in data:
        crc = ((crc << 8) & 0xFFFFFFFF) ^ CRC_TABLE[((crc >> 24) & 0xFF) ^ b]
    return crc


def ogg_page(packet, granule, serial, seq, flags):
    lacing = [255] * (len(packet) // 255) + [len(packet) % 255]
    head = struct.pack("<4sBBqIIIB", b"OggS", 0, flags, granule, serial, seq, 0, len(lacing)) + bytes(lacing)
    page = bytearray(head + packet)
    struct.pack_into("<I", page, 22, ogg_crc(page))
    return bytes(page)


def packet_duration(p):
    """samples (48 kHz) of an Opus packet: frame size x frame count (opus_packet_get_nb_samples)"""
    toc = p[0]
    if toc & 0x80:
        fs = 120 << ((toc >> 3) & 3)
    elif (toc & 0x60) == 0x60:
        fs = 960 if toc & 0x08 else 480
    else:
        fs = (480, 960, 1920, 2880)[(toc >> 3) & 3]
    code = toc & 3
    count = 1 if code == 0 else 2 if code in (1, 2) else (p[1] & 0x3F)
    return fs * count


def write_ogg(path, packets, serial, final_granule):
    """an Ogg Opus file, one packet per page (the two header packets first)"""
    out, granule = [], 0
    for i, p in enumerate(packets):
        if i >= 2:
            granule += packet_duration(p)
        g = 0 if i < 2 else (final_granule if i == len(packets) - 1 else granule)
        flags = 2 if i == 0 else (4 if i == len(packets) - 1 else 0)
        out.append(ogg_page(p, g, serial, i, flags))
    with open(path, "wb") as f:
        f.write(b"".join(out))


# --- the cases ---------------------------------------------------------------------------------------------

# name, channels, frame ms, bitrate (kbps), extra ffmpeg options, lost packets (indexes of audio packets), and
# optionally the decoder's channels and sampling rate when they differ from the stream's / 48000 (then the stream
# is judged against libopus 1.5.2 only: ffmpeg decodes with the header's channels at 48 kHz)
CASES = [
    ("mono_20ms_64k", 1, 20, 64, [], []),
    ("stereo_20ms_128k", 2, 20, 128, [], []),
    ("stereo_20ms_24k", 2, 20, 24, [], []),
    ("stereo_20ms_8k", 2, 20, 8, [], []),
    ("mono_20ms_12k_nb", 1, 20, 12, ["-cutoff", "4000"], []),
    ("stereo_20ms_32k_wb", 2, 20, 32, ["-cutoff", "8000"], []),
    ("stereo_20ms_48k_swb", 2, 20, 48, ["-cutoff", "12000"], []),
    ("stereo_20ms_256k_cbr", 2, 20, 256, ["-vbr", "off"], []),
    ("mono_10ms_32k", 1, 10, 32, [], []),
    ("stereo_10ms_96k", 2, 10, 96, [], []),
    ("mono_5ms_48k", 1, 5, 48, [], []),
    ("stereo_5ms_160k", 2, 5, 160, [], []),
    ("mono_2.5ms_64k", 1, 2.5, 64, [], []),
    ("stereo_2.5ms_256k", 2, 2.5, 256, [], []),
    ("stereo_2.5ms_48k", 2, 2.5, 48, [], []),
    ("stereo_40ms_64k_cbr", 2, 40, 64, ["-vbr", "off"], []),
    ("stereo_60ms_96k", 2, 60, 96, [], []),
    ("mono_40ms_16k", 1, 40, 16, [], []),
    ("stereo_20ms_128k_loss", 2, 20, 128, [], [20, 50, 51, 100, 101, 102, 103, 104, 105, 106, 107, 109, 150]),
    ("mono_10ms_32k_loss", 1, 10, 32, [], [30, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 120]),
    ("stereo_5ms_96k_loss", 2, 5, 96, [], [100, 200, 201, 400]),
    ("mono_2.5ms_64k_loss", 1, 2.5, 64, [], [400, 1000, 1001]),
    ("stereo_to_mono_20ms_96k", 2, 20, 96, [], [], {"dec_channels": 1}),
    ("mono_to_stereo_10ms_48k", 1, 10, 48, [], [40, 41], {"dec_channels": 2}),
    ("stereo_20ms_96k_at24k", 2, 20, 96, [], [], {"rate": 24000}),
    ("mono_10ms_32k_at16k_loss", 1, 10, 32, [], [30, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71], {"rate": 16000}),
    ("stereo_5ms_64k_at8k", 2, 5, 64, [], [], {"rate": 8000, "dec_channels": 1}),
]

# celt_decoder_c's celt_vbCounts
COUNTERS = ["frames", "transient", "anti-collapse", "intensity stereo", "dual stereo", "silence", "post-filter",
            "pitch PLC", "noise PLC", "mono to stereo", "stereo to mono", "downsampled", "spread not normal",
            "dynalloc boost", "skipped bands"]


def case(c):
    name, ch, ms, kbps, extra, lost = c[:6]
    opts = c[6] if len(c) > 6 else {}
    return name, ch, ms, kbps, extra, lost, opts


def run(keep=False):
    if os.path.exists(WORK):
        shutil.rmtree(WORK)
    os.makedirs(WORK)
    if not build_exe("CeltTests", True):
        print("celt: CeltTests.exe did not build")
        return False
    fast = build_exe("CeltTestsFast", False)
    src = os.path.join(WORK, "signal.wav")
    x = signal()
    write_wav(src, x)

    lst, info = [], {}
    for c in CASES:
        name, ch, ms, kbps, extra, lost, opts = case(c)
        opus = os.path.join(WORK, name + ".opus")
        cmd = [FFMPEG, "-v", "error", "-y", "-i", src, "-ac", str(ch), "-c:a", "libopus", "-application", "lowdelay",
               "-frame_duration", str(ms), "-b:a", "%dk" % kbps] + extra + [opus]
        r = subprocess.run(cmd, capture_output=True, text=True)
        if r.returncode != 0:
            print("celt: ffmpeg failed for %s: %s" % (name, r.stderr.strip()))
            return False
        with open(opus, "rb") as f:
            data = f.read()
        packets, serial = ogg_packets(data)
        head = packets[0]
        channels, preskip = head[9], struct.unpack_from("<H", head, 10)[0]
        audio = packets[2:]
        if lost:
            for k in lost:
                audio[k] = bytes([audio[k][0] & 0xFC])      # the TOC alone, code 0: one 0-byte frame
            write_ogg(opus, packets[:2] + audio, serial, last_granule(data))
        tocs = sorted(set(p[0] for p in audio))
        dec_channels = opts.get("dec_channels", channels)
        rate = opts.get("rate", RATE)
        pk = os.path.join(WORK, name + ".pkt")
        with open(pk, "wb") as f:
            f.write(b"CELTPKT1" + struct.pack("<iiii", dec_channels, preskip, len(audio), rate))
            for p in audio:
                f.write(struct.pack("<i", len(p)) + p)
        ref = os.path.join(WORK, name + ".ref.f32")
        r = subprocess.run([FFMPEG, "-v", "error", "-y", "-c:a", "libopus", "-request_sample_fmt", "flt", "-i", opus,
                            "-f", "f32le", "-c:a", "pcm_f32le", ref], capture_output=True, text=True)
        if r.returncode != 0:
            print("celt: ffmpeg could not decode %s: %s" % (name, r.stderr.strip()))
            return False
        info[name] = (dec_channels, preskip, len(audio), tocs, len(lost), rate, bool(opts))
        lst.append("%s\t%s" % (pk, name))
    listfile = os.path.join(WORK, "list.txt")
    with open(listfile, "w", newline="\r\n") as f:
        f.write("\n".join(lst) + "\n")

    def run_exe(exe, outdir):
        os.makedirs(outdir, exist_ok=True)
        r = subprocess.run([os.path.join(build.BIN, exe + ".exe"), listfile, outdir], capture_output=True, text=True)
        res = {}
        path = os.path.join(outdir, "results.txt")
        if not os.path.exists(path):
            print("celt: %s wrote no results (exit %d)" % (exe, r.returncode))
            return None
        with open(path) as f:
            for line in f.read().splitlines():
                parts = line.split("\t")
                res[parts[0]] = parts[1:]
        return res

    res = run_exe("CeltTests", os.path.join(WORK, "out"))
    if res is None:
        return False
    passed, judged, details, worst, worst152, worst_loss = 0, 0, [], 0.0, 0.0, 0.0
    totals = [0] * len(COUNTERS)
    for c in CASES:
        name, ch, ms, kbps, extra, lost, opts = case(c)
        channels, preskip, npk, tocs, nlost, rate, overridden = info[name]
        got = res.get(name)
        if got is None or got[0] == "error":
            judged += 1
            details.append("  %-26s FAILED: %s" % (name, " ".join(got[1:]) if got else "no result"))
            continue
        samples, lost_frames, errors, text = int(got[2]), int(got[4]), int(got[5]), (got[7] if len(got) > 7 else "")
        counts = [int(v) for v in got[8].split(",")] if len(got) > 8 and got[8] else [0] * 16
        for k in range(len(COUNTERS)):
            totals[k] += counts[k]
        mine_all = np.fromfile(os.path.join(WORK, "out", name + ".f32"), dtype="<f4")
        # ffmpeg (libopus 1.6.1): after the pre-skip, up to the end the last granule position gives
        maxd, max16, n, nref, rms_ref = None, None, 0, 0, 0.0
        if not overridden:
            ref = np.fromfile(os.path.join(WORK, name + ".ref.f32"), dtype="<f4")
            mine = mine_all[preskip * channels:]
            n, nref = min(len(mine), len(ref)), len(ref)
            maxd = float(np.abs(mine[:n].astype(np.float64) - ref[:n].astype(np.float64)).max()) if n else 1.0
            rms_ref = float(np.sqrt(np.mean(ref[:n].astype(np.float64) ** 2))) if n else 0.0
            mine16 = np.clip(np.round(mine[:n] * 32768), -32768, 32767)
            ref16 = np.clip(np.round(ref[:n] * 32768), -32768, 32767)
            max16 = int(np.abs(mine16 - ref16).max()) if n else 99999
        # libopus 1.5.2 (the official build's DLL): every sample, the pre-skip included
        r152 = os.path.join(WORK, name + ".ref152.f32")
        d152 = None
        if ref152(os.path.join(WORK, name + ".pkt"), r152):
            ref2 = np.fromfile(r152, dtype="<f4")
            n2 = min(len(mine_all), len(ref2))
            d152 = float(np.abs(mine_all[:n2].astype(np.float64) - ref2[:n2].astype(np.float64)).max()) \
                if n2 == len(ref2) else 1.0
            if overridden:
                rms_ref = float(np.sqrt(np.mean(ref2.astype(np.float64) ** 2)))
        clean = errors == 0 and not text
        if nlost or overridden:
            # judged against 1.5.2 (ffmpeg's 1.6.1 conceals differently, and decodes with the header's channels
            # at 48 kHz); without the DLL: not judged
            if d152 is None:
                ok = None
            else:
                ok = clean and d152 <= (PLC_TOLERANCE if nlost else TOLERANCE)
                if nlost:
                    worst_loss = max(worst_loss, d152)
                else:
                    worst152 = max(worst152, d152)
        else:
            ok = clean and n == nref and maxd <= TOLERANCE
            worst = max(worst, maxd)
            if d152 is not None:
                worst152 = max(worst152, d152)
        if ok is not None:
            judged += 1
            passed += ok
        bw = sorted(set({0: "NB", 1: "WB", 2: "SWB", 3: "FB"}[(t >> 5) & 3] for t in tocs))
        codes = sorted(set(t & 3 for t in tocs))
        stream_ch = sorted(set(2 if t & 4 else 1 for t in tocs))
        used = [COUNTERS[k] for k in range(1, len(COUNTERS)) if counts[k]]
        details.append("  %-26s %s decoder %d ch at %d Hz, stream %s ch, %d packets (codes %s, %s), %d lost frames: "
                       "%s; libopus 1.5.2 dll: %s; rms %.3f%s; paths: %s" %
                       ({True: "ok    ", False: "DIFFER", None: "n/a   "}[ok], name, channels, rate,
                        "/".join(map(str, stream_ch)), npk, ",".join(map(str, codes)), "/".join(bw), lost_frames,
                        ("ffmpeg (1.6.1): %d samples compared of %d, max abs diff %.3g (16-bit: %d)" %
                         (n // channels, nref // channels, maxd, max16)) if maxd is not None else "ffmpeg: n/a",
                        "%.3g" % d152 if d152 is not None else "n/a", rms_ref, "" if not text else " (" + text + ")",
                        ", ".join(used) if used else "-"))
    print("celt: %d of %d streams decode as libopus does (float; without losses: ffmpeg's libopus max abs difference "
          "%.3g, the official 1.5.2 libopus.dll's %.3g, tolerance %g; with losses: 1.5.2's %.3g, tolerance %g)" %
          (passed, judged, worst, worst152, TOLERANCE, worst_loss, PLC_TOLERANCE))
    for d in details:
        print(d)
    print("celt paths exercised (frames): " + ", ".join("%s %d" % (COUNTERS[k], totals[k]) for k in range(len(COUNTERS))))
    if fast:
        resf = run_exe("CeltTestsFast", os.path.join(WORK, "out_fast"))
        if resf:
            total_ms = float(resf["total"][0])
            audio_s = 0.0
            for c in CASES:
                name = c[0]
                got = resf.get(name)
                if got and got[0] != "error":
                    audio_s += int(got[2]) / info[name][5]
            print("celt speed (CeltTestsFast.exe, compiled as WzSound: optimised, no bounds checks): %.1f s of audio "
                  "in %.0f ms, %.1f x real time" % (audio_s, total_ms, audio_s * 1000 / total_ms))
            for name in ("stereo_20ms_128k", "stereo_20ms_256k_cbr", "mono_20ms_64k", "stereo_10ms_96k",
                         "stereo_2.5ms_256k"):
                got = resf.get(name)
                if got and got[0] != "error":
                    sec = int(got[2]) / RATE
                    print("  %-26s %.1f s in %.0f ms: %.1f x real time" % (name, sec, float(got[3]),
                                                                          sec * 1000 / float(got[3])))
    else:
        print("celt speed: CeltTestsFast.exe did not build")
    if not keep:
        shutil.rmtree(os.path.join(WORK, "out_fast"), ignore_errors=True)
    return passed == judged


if __name__ == "__main__":
    sys.exit(0 if run("--keep" in sys.argv) else 1)
