"""
video_audio_test.py - the videos' Vorbis audio through WzSound's libvorbis-style packet API (cluster K).

lib/sequence (sequence.cpp) decodes a video's audio with libvorbis's synthesis API, one Ogg packet at a time:
vorbis_synthesis_headerin for the three header packets, then per audio packet vorbis_synthesis and
vorbis_synthesis_blockin, vorbis_synthesis_pcmout (float per channel) and vorbis_synthesis_read. WzSound
implements that API over its stb_vorbis port (src/sound/stb_vorbis_c.bas, "libvorbis's synthesis API").

  python -I tools\\video_audio_test.py [--keep] [--limit=N]

* Python demuxes each video of data/sequences.wz (Python's zipfile; Theora + Vorbis Ogg files) into its Vorbis
  stream's packets as libogg's ogg_stream_packetout hands them over (b_o_s, e_o_s, the granule position on the
  last packet completed on a page, -1 on the others, packetno counting from 0) and writes them to a packet
  file;
* bin\\VorbisPacketTests.exe (src/vorbispackettests, with stb_vorbis_c.bas compiled in, bounds checks on) runs
  each packet file through the API as sequence.cpp does and writes the samples (float, interleaved), the
  header calls' results, and its checks of libvorbis's bookkeeping: the dsp state's granule position after
  each packet that carries one equals it (vd.granulepos), the samples returned so far equal the granule
  position there, a block is refused while samples are pending (OV_EINVAL), pcmout/read in odd-sized pieces;
* the reference is ffmpeg's own Vorbis decoder on the same file's audio stream (-map 0:a:0), float and 16-bit:
  the samples must agree within 1 in 16 bits (as stb_vorbis's game sounds do with ffmpeg: the x87's float
  evaluation rounds differently), the float difference is reported; the length must be the last granule
  position (libvorbis's count; ffmpeg's decode may keep the last frame's padding or trim it).

Every video is tested (--limit=N: the first N); those without audio (the research videos) are counted.
run(keep) -> bool for run_tests.py's --media.
"""
import os
import shutil
import struct
import subprocess
import sys
import time
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import build  # noqa: E402

ROOT = build.ROOT
SRC = os.path.join(build.SRC, "vorbispackettests")
EXE = os.path.join(build.BIN, "VorbisPacketTests.exe")
WORK = os.path.join(build.BUILD, "video_audio_test")
SEQUENCES = os.path.join(ROOT, "data", "sequences.wz")
FFMPEG = shutil.which("ffmpeg") or "ffmpeg"
TOLERANCE = 1           # 16-bit units
NO_GRANULE = -1


def ogg_vorbis_packets(data):
    """The Vorbis stream's packets of an Ogg file, as libogg's ogg_stream_packetout returns them:
    (bytes, b_o_s, e_o_s, granulepos, packetno), in order; None when the file has no Vorbis stream."""
    pos = 0
    serial = None
    seen_bos = set()
    packets = []
    pending = bytearray()           # the packet being gathered across segments
    pending_bos = False
    packetno = 0
    while pos + 27 <= len(data):
        if data[pos:pos + 4] != b"OggS":
            raise ValueError("lost sync at %d" % pos)
        flags = data[pos + 5]
        granule = struct.unpack_from("<q", data, pos + 6)[0]
        sn = struct.unpack_from("<I", data, pos + 14)[0]
        nseg = data[pos + 26]
        lacing = data[pos + 27:pos + 27 + nseg]
        body = pos + 27 + nseg
        size = sum(lacing)
        if flags & 2 and sn not in seen_bos:
            seen_bos.add(sn)
            if serial is None and nseg and lacing[0] >= 7 and data[body:body + 7] == b"\x01vorbis":
                serial = sn
        if sn == serial:
            # the packets this page completes: libogg puts the page's granule position on the last one
            ends = []
            p = body
            bos = bool(flags & 2)
            first_seg = True
            for i, v in enumerate(lacing):
                if first_seg and not (flags & 1) and pending:
                    raise ValueError("a packet without its continuation at page %d" % pos)
                pending += data[p:p + v]
                if first_seg and bos:
                    pending_bos = True
                first_seg = False
                p += v
                if v < 255:
                    ends.append((bytes(pending), pending_bos, i == nseg - 1))
                    pending = bytearray()
                    pending_bos = False
            eos_page = bool(flags & 4)
            for k, (pkt, pbos, last_seg) in enumerate(ends):
                g = granule if k == len(ends) - 1 else NO_GRANULE
                eos = eos_page and last_seg     # the 0x200 flag sits on the page's last segment
                packets.append((pkt, pbos, eos, g, packetno))
                packetno += 1
        pos = body + size
    return packets if serial is not None else None


def write_packet_file(path, packets):
    """The packets for VorbisPacketTests: a count (int32), then per packet: bytes, b_o_s, e_o_s (int32),
    granulepos, packetno (float64), the bytes."""
    with open(path, "wb") as f:
        f.write(struct.pack("<i", len(packets)))
        for pkt, bos, eos, g, n in packets:
            f.write(struct.pack("<iiidd", len(pkt), 1 if bos else 0, 1 if eos else 0, float(g), float(n)))
            f.write(pkt)


def build_exe(optimised=False):
    """bin/VorbisPacketTests.exe: src/vorbispackettests and WzSound's stb_vorbis port, bounds checks on (as
    run_tests.py builds VorbisTests.exe; src/shared is left out); optimised: as WzSound is built (fast code, no
    bounds checks), to measure the speed the game gets."""
    vbp = build.write_vbp("VorbisPacketTests", "exe", SRC, [], False, None)
    with open(vbp, "r", encoding="latin-1") as f:
        lines = [l for l in f.read().split("\n") if "..\\shared\\" not in l]
    k = next(i for i, l in enumerate(lines) if l.startswith("Startup="))
    lines.insert(k, "Module=stb_vorbis_c; ..\\sound\\stb_vorbis_c.bas")
    if optimised:
        lines = ["OptimizationType=0" if l.startswith("OptimizationType=") else
                 "BoundsCheck=-1" if l.startswith("BoundsCheck=") else l for l in lines]
    with open(vbp, "w", encoding="latin-1", newline="\r\n") as f:
        f.write("\n".join(lines))
    return build.make("VorbisPacketTests", SRC, vbp)


def reference(ogg_path, fmt):
    """ffmpeg's decode of the file's first audio stream: fmt "f32le" or "s16le" (not kept: 400 MB for all)."""
    codec = "pcm_f32le" if fmt == "f32le" else "pcm_s16le"
    r = subprocess.run([FFMPEG, "-v", "error", "-i", ogg_path, "-map", "0:a:0", "-f", fmt, "-acodec", codec, "-"],
                       capture_output=True)
    if r.returncode != 0:
        return None
    return r.stdout


def run(keep=False, limit=None, optimised=False):
    import numpy
    if not os.path.exists(SEQUENCES):
        print("video audio: %s is missing (tools/setup_data.py)" % os.path.relpath(SEQUENCES, ROOT))
        return False
    if os.path.exists(WORK):
        shutil.rmtree(WORK)
    os.makedirs(WORK)
    z = zipfile.ZipFile(SEQUENCES)
    names = sorted(n for n in z.namelist() if n.lower().endswith(".ogg"))
    if limit:
        names = names[:limit]
    videos = []         # (name, ogg path, packets)
    problems = []
    no_audio = []       # the videos without a Vorbis stream (Theora only)
    for i, n in enumerate(names):
        data = z.read(n)
        ogg = os.path.join(WORK, "%d.ogg" % i)
        with open(ogg, "wb") as f:
            f.write(data)
        try:
            packets = ogg_vorbis_packets(data)
        except ValueError as e:
            problems.append("%s: demux: %s" % (n, e))
            packets = None
        if packets is None:
            if not problems or not problems[-1].startswith(n):
                no_audio.append(n)
            continue
        write_packet_file(os.path.join(WORK, "%d.pkt" % i), packets)
        videos.append((i, n, ogg, packets))
    if not build_exe(optimised):
        return False
    listfile = os.path.join(WORK, "files.txt")
    with open(listfile, "w", encoding="utf-8") as f:
        f.write("\n".join(os.path.join(WORK, "%d.pkt" % i) for i, _, _, _ in videos) + "\n")
    t = time.time()
    r = subprocess.run([EXE, listfile, WORK], capture_output=True)
    print("VorbisPacketTests.exe: %.1f s, exit code %d" % (time.time() - t, r.returncode))
    results = {}
    total_ms = None
    resfile = os.path.join(WORK, "results.txt")
    if os.path.exists(resfile):
        for line in open(resfile, encoding="latin-1").read().splitlines():
            parts = line.split("\t")
            if parts[0] == "total":
                total_ms = float(parts[1])
            elif parts[0].isdigit():
                results[int(parts[0])] = parts
    ok = 0
    dist = {}
    worst_float = 0.0
    total_samples = 0
    total_secs = 0.0
    rates = set()
    ffmpeg_longer = []
    for k, (i, n, ogg, packets) in enumerate(videos):
        res = results.get(k)
        if res is None or res[1] == "error":
            problems.append("%s: %s" % (n, "no result" if res is None else " ".join(res[2:])))
            continue
        channels, rate, samples, checks = int(res[1]), int(res[2]), int(res[3]), res[5]
        rates.add((channels, rate))
        total_samples += samples
        total_secs += samples / float(rate)
        info = ""
        if checks != "ok":
            info += " checks: %s" % checks
        last_granule = max(g for _, _, _, g, _ in packets)
        if samples != last_granule:
            info += " samples %d, the last granule position %d" % (samples, last_granule)
        ref_f = reference(ogg, "f32le")
        ref_s = reference(ogg, "s16le")
        if ref_f is None or ref_s is None:
            problems.append("%s: ffmpeg failed" % n)
            continue
        with open(os.path.join(WORK, "%d.f32" % k), "rb") as f:
            vb = numpy.frombuffer(f.read(), dtype="<f4")
        a = numpy.frombuffer(ref_f, dtype="<f4")
        s = numpy.frombuffer(ref_s, dtype="<i2").astype(numpy.int32)
        ffmpeg_samples = len(a) // channels
        if ffmpeg_samples != samples:
            ffmpeg_longer.append(ffmpeg_samples - samples)
        m = min(len(a), len(vb), len(s))
        if m:
            fd = float(numpy.abs(a[:m].astype(numpy.float64) - vb[:m].astype(numpy.float64)).max())
            vb16 = numpy.clip(numpy.rint(vb[:m].astype(numpy.float32) * numpy.float32(32768.0)), -32768, 32767)
            d16 = int(numpy.abs(vb16.astype(numpy.int32) - s[:m]).max())
        else:
            fd, d16 = 0.0, 0
        worst_float = max(worst_float, fd)
        dist[d16] = dist.get(d16, 0) + 1
        if d16 > TOLERANCE:
            info += " a sample differs by %d (16-bit), %.3g (float)" % (d16, fd)
        if info:
            problems.append("%s (%d ch, %d Hz):%s" % (n, channels, rate, info))
        else:
            ok += 1
    print("video audio: %d of %d videos' Vorbis streams decode through the packet API as ffmpeg does (16-bit "
          "samples within %d; videos by largest difference: %s; largest float difference %.3g); %s; %d videos "
          "without audio" % (ok, len(videos), TOLERANCE, ", ".join("%d: %d" % (d, dist[d]) for d in sorted(dist)) or "-",
                             worst_float, ", ".join("%d ch %d Hz" % cr for cr in sorted(rates)) or "-", len(no_audio)))
    if total_ms is not None and rates:
        print("  %d samples (%.0f s of audio) decoded in %.1f s: %.0fx real time (%s)"
              % (total_samples, total_secs, total_ms / 1000.0, total_secs / max(total_ms / 1000.0, 1e-9),
                 "optimised as WzSound: fast code, no bounds checks" if optimised else
                 "test build: bounds checks, no optimisation"))
    print("  lengths: %d of %d at the last granule position; ffmpeg's decode differs in length in %d"
          % (len(videos) - sum(1 for p in problems if "the last granule position" in p), len(videos), len(ffmpeg_longer)))
    if ffmpeg_longer:
        print("  ffmpeg's decode differs in length in %d videos (ffmpeg - port: %d to %d; the port stops at the last "
              "granule position, as libvorbis)" % (len(ffmpeg_longer), min(ffmpeg_longer), max(ffmpeg_longer)))
    for p in problems[:20]:
        print("  " + p)
    passed = ok == len(videos) and not problems
    if passed and not keep:
        shutil.rmtree(WORK)
    return passed


if __name__ == "__main__":
    lim = None
    for a in sys.argv[1:]:
        if a.startswith("--limit="):
            lim = int(a.split("=", 1)[1])
    sys.exit(0 if run("--keep" in sys.argv, lim, "--optimised" in sys.argv) else 1)
