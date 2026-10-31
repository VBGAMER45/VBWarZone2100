"""
mixer_test.py - cluster H's test of WzSound's mixer (src/sound/al_win.bas, the OpenAL subset over waveOut):
builds bin/MixerTests.exe (src/mixertests + al_win.bas + src/shared, no DLL), runs it on
build/mixer_test, and checks its WAV files (synthetic sines mixed in capture mode) and results.txt:
levels and gains, the left/right balance of sources around the listener (OpenAL's default orientation and
WZ's), the inverse distance model, resampling 11025 -> 44100 Hz and the pitch, looping, a queued stream's
buffers (processed, unqueued, the end, played again), pause/resume/stop, clipping, a stereo buffer, error
codes. Prints "mixer: N of M checks pass" and the mixing cost per second of audio.

  python -I tools\\mixer_test.py [--keep]          (run_tests.py --sound calls run(keep))
"""
import array
import math
import os
import shutil
import subprocess
import sys
import wave

ROOT = r"C:\Users\Owner\desktop\websites\warzone"
sys.path.insert(0, os.path.join(ROOT, "tools"))
import build  # noqa: E402

FOLDER = os.path.join(ROOT, "src", "mixertests")
WORK = os.path.join(ROOT, "build", "mixer_test")
RATE = 44100

AL_INITIAL, AL_PLAYING, AL_PAUSED, AL_STOPPED = 0x1011, 0x1012, 0x1013, 0x1014
AL_INVALID_NAME, AL_INVALID_ENUM, AL_INVALID_VALUE, AL_INVALID_OPERATION = 0xA001, 0xA002, 0xA003, 0xA004
SINE_RMS = 16000 / math.sqrt(2)          # a sine of amplitude 16000
CENTRE = math.sqrt(0.5)                  # equal power, centred


def build_exe():
    # its own modules and al_win.bas only (src/shared's call WzFramework's debug_; al_win needs none)
    lines = ["Type=Exe", build.STDOLE]
    for n in build.sources(FOLDER):
        path = os.path.join(FOLDER, n)
        build.normalise(path)
        lines.append("Module=%s; %s" % (build.vb_name(path), n))
    al_win = os.path.join(ROOT, "src", "sound", "al_win.bas")
    build.normalise(al_win)
    lines.append("Module=al_win; ..\\sound\\al_win.bas")
    lines.append(build.VBP_TAIL.format(name="MixerTests", ext="exe", path=os.path.relpath(build.BIN, FOLDER), startup="Sub Main", compat="", bounds="0", optimisation="2",
                                       threading=""))
    vbp = os.path.join(FOLDER, "MixerTests.vbp")
    with open(vbp, "w", encoding="latin-1", newline="\r\n") as f:
        f.write("\n".join(lines))
    return build.make("MixerTests", FOLDER, vbp)


def read_wav(name):
    """the left and right channels of build/mixer_test/<name>.wav"""
    with wave.open(os.path.join(WORK, name + ".wav"), "rb") as w:
        data = w.readframes(w.getnframes())
    a = array.array("h")
    a.frombytes(data)
    return a[0::2], a[1::2]


def rms(x, start=0, end=None):
    seg = x[start:end]
    if not seg:
        return 0.0
    return math.sqrt(sum(float(v) * v for v in seg) / len(seg))


def freq(x, start, end):
    """the frequency of a sine from its zero crossings"""
    seg = x[start:end]
    crossings = sum(1 for i in range(1, len(seg)) if (seg[i - 1] < 0) != (seg[i] < 0))
    return crossings / 2.0 / (len(seg) / RATE)


def run(keep=False):
    if not build_exe():
        print("mixer: MixerTests.exe did not build")
        return False
    if os.path.exists(WORK):
        shutil.rmtree(WORK)
    os.makedirs(WORK)
    r = subprocess.run([os.path.join(build.BIN, "MixerTests.exe"), WORK], timeout=600)
    res = {}
    path = os.path.join(WORK, "results.txt")
    if os.path.exists(path):
        for l in open(path, encoding="latin-1"):
            k, _, v = l.strip().partition(" ")
            res[k] = v
    checks = []

    def check(name, ok, detail=""):
        checks.append((name, bool(ok), detail))

    def near(name, value, expected, tol):
        check(name, abs(value - expected) <= tol * max(abs(expected), 1.0),
              "%.1f (expected %.1f)" % (value, expected))

    def eq(name, key, expected):
        v = res.get(key)
        check(name, v is not None and int(v) == expected, "%s = %s (expected %d)" % (key, v, expected))

    if res.get("done") != "1":
        check("MixerTests.exe ran to its end (exit code %d)" % r.returncode, False)
    else:
        L, R = read_wav("gain")
        near("gain 0.5, centred: left", rms(L), SINE_RMS * 0.5 * CENTRE, 0.02)
        near("gain 0.5, centred: right", rms(R), SINE_RMS * 0.5 * CENTRE, 0.02)
        L, R = read_wav("pan_right")
        near("source to the right: right", rms(R), SINE_RMS, 0.02)
        check("source to the right: left silent", rms(L) < 0.01 * SINE_RMS, "%.1f" % rms(L))
        L, R = read_wav("pan_left")
        near("source to the left: left", rms(L), SINE_RMS, 0.02)
        check("source to the left: right silent", rms(R) < 0.01 * SINE_RMS, "%.1f" % rms(R))
        L, R = read_wav("pan_45")
        pan = math.sqrt(0.5)
        near("source at 45 degrees: left", rms(L), SINE_RMS * math.sqrt((1 - pan) / 2), 0.02)
        near("source at 45 degrees: right", rms(R), SINE_RMS * math.sqrt((1 + pan) / 2), 0.02)
        L, R = read_wav("pan_wz")
        near("WZ's orientation turned by pi / 2, source at +y: right", rms(R), SINE_RMS, 0.02)
        check("WZ's orientation turned by pi / 2, source at +y: left silent", rms(L) < 0.01 * SINE_RMS, "%.1f" % rms(L))
        L, R = read_wav("distance")
        near("inverse distance clamped, 4 ahead: left", rms(L), SINE_RMS * 0.25 * CENTRE, 0.02)
        near("inverse distance clamped, 4 ahead: right", rms(R), SINE_RMS * 0.25 * CENTRE, 0.02)
        L, R = read_wav("resample")
        near("11025 Hz buffer: 441 Hz after resampling", freq(L, 0, RATE), 441.0, 0.01)
        near("11025 Hz buffer: level", rms(L, 0, RATE), SINE_RMS * CENTRE, 0.02)
        check("11025 Hz buffer: lasts 1 s", rms(L, RATE, None) < 1.0 and rms(L, RATE - 4410, RATE) > 0.9 * SINE_RMS * CENTRE,
              "after: %.1f, last 0.1 s: %.1f" % (rms(L, RATE, None), rms(L, RATE - 4410, RATE)))
        eq("11025 Hz buffer: playing at 1 s", "resample_state_1s", AL_PLAYING)
        eq("11025 Hz buffer: stopped after its end", "resample_state_end", AL_STOPPED)
        L, R = read_wav("pitch")
        near("pitch 2: 1 kHz plays at 2 kHz", freq(L, 0, 22050), 2000.0, 0.01)
        check("pitch 2: a 2 s buffer lasts 1 s", rms(L, 44100, None) < 1.0 and rms(L, 39690, 44100) > 0.9 * SINE_RMS * CENTRE,
              "after: %.1f" % rms(L, 44100, None))
        eq("pitch 2: playing at 0.5 s", "pitch_state_1s", AL_PLAYING)
        eq("pitch 2: stopped after its end", "pitch_state_end", AL_STOPPED)
        L, R = read_wav("loop")
        near("looping 1000-frame buffer: still sounding at 1 s", rms(L, RATE - 4410, RATE), SINE_RMS * CENTRE, 0.05)
        eq("looping: playing", "loop_state", AL_PLAYING)
        eq("looping: none processed", "loop_processed", 0)
        L, R = read_wav("queue")
        near("queue: buffer 1 (500 Hz)", freq(L, 441, 3969), 500.0, 0.02)
        near("queue: buffer 2 (1000 Hz)", freq(L, 4851, 8379), 1000.0, 0.02)
        near("queue: buffer 3 (1500 Hz)", freq(L, 9261, 12789), 1500.0, 0.02)
        check("queue: silent after the end", rms(L, 13230 + 10, 19845) < 1.0, "%.1f" % rms(L, 13230 + 10, 19845))
        near("queue: played again (500 Hz)", freq(L, 19845 + 441, 24255 - 441), 500.0, 0.02)
        eq("queue: 3 queued", "queue_queued", 3)
        eq("queue: none processed at 0.05 s", "queue_processed_0.05s", 0)
        eq("queue: 1 processed at 0.15 s", "queue_processed_0.15s", 1)
        eq("queue: the first unqueued first", "queue_unqueued_first", 1)
        eq("queue: 2 queued after", "queue_queued_after_unqueue", 2)
        eq("queue: unqueueing more than processed", "queue_unqueue_too_many_error", AL_INVALID_VALUE)
        eq("queue: stopped at the end", "queue_state_end", AL_STOPPED)
        eq("queue: all processed when stopped", "queue_processed_end", 2)
        eq("queue: playing again", "queue_state_replay", AL_PLAYING)
        eq("queue: no error", "queue_error", 0)
        L, R = read_wav("states")
        seg = [rms(L, i * 4410 + 10, (i + 1) * 4410 - 10) for i in range(5)]
        check("pause / resume / stop / play: sound, silence, sound, silence, sound",
              seg[0] > 7000 and seg[1] < 1 and seg[2] > 7000 and seg[3] < 1 and seg[4] > 7000,
              " ".join("%.0f" % v for v in seg))
        eq("paused", "states_paused", AL_PAUSED)
        eq("resumed", "states_resumed", AL_PLAYING)
        eq("stopped", "states_stopped", AL_STOPPED)
        eq("a stopped static source: its buffer processed", "states_processed_stopped", 1)
        eq("played again", "states_replayed", AL_PLAYING)
        eq("states: no error", "states_error", 0)
        L, R = read_wav("clamp")
        check("two loud sources: clipped to 16 bits", max(L) == 32767 and min(L) == -32768 and max(R) == 32767,
              "%d %d" % (max(L), min(L)))
        L, R = read_wav("stereo")
        near("stereo buffer: left as given", rms(L), SINE_RMS, 0.02)
        check("stereo buffer: not panned (right stays silent)", rms(R) < 1.0, "%.1f" % rms(R))
        eq("invalid source name", "error_invalid_name", AL_INVALID_NAME)
        eq("negative gain", "error_negative_gain", AL_INVALID_VALUE)
        eq("alBufferData on a buffer in use", "error_buffer_in_use", AL_INVALID_OPERATION)
        eq("alDeleteBuffers on a buffer in use", "error_delete_in_use", AL_INVALID_OPERATION)
        eq("bad enum", "error_bad_enum", AL_INVALID_ENUM)
        eq("alGetError clears the error", "error_none_left", 0)
        for k in ("gain_error", "pan_right_error", "pan_left_error", "pan_45_error", "pan_wz_error", "distance_error",
                  "clamp_error", "stereo_error"):
            eq("%s: none" % k, k, 0)
        if res.get("device_open") == "1":
            eq("waveOut device (silent): playing at the start", "device_state_start", AL_PLAYING)
            eq("waveOut device (silent): a 0.1 s buffer ended within 0.4 s", "device_state_0.4s", AL_STOPPED)
            eq("waveOut device: no error", "device_error", 0)
        else:
            print("mixer: no waveOut device opened: its checks skipped")
    failed = [c for c in checks if not c[1]]
    for name, ok, detail in failed:
        print("  FAIL %s: %s" % (name, detail))
    print("mixer: %d of %d checks pass" % (len(checks) - len(failed), len(checks)))
    costs = ["%s sources %s ms" % (n, res["bench_%s_ms_per_s" % n]) for n in (1, 8, 32) if "bench_%s_ms_per_s" % n in res]
    if costs:
        print("mixer: mixing 1 s of audio (44100 Hz stereo): " + ", ".join(costs))
    if not failed and not keep:
        shutil.rmtree(WORK, ignore_errors=True)
    return not failed


if __name__ == "__main__":
    sys.exit(0 if run("--keep" in sys.argv) else 1)
