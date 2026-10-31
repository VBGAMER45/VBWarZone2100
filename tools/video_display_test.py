"""
video_display_test.py - checks how the game shows its videos against a model of 4.7.0's display
(lib/sequence/sequence.cpp's video_write and VideoPSO), in every video size and scanline mode, and in the
intelligence screen.

  python -I tools\\video_display_test.py [--official]

* The model: the official 4.7.0 build's libtheoradec.dll (ref/official/wz470/bin, through ctypes as
  theora_test.py does) decodes the video at its maximum post-processing level, as sequence.cpp asks;
  video_write's YUV -> RGB (its integer formula, chroma per 2x2 block) builds the texture: the legacy
  theora.h API's ti.frame_width x frame_height (the picture size; the coded size is ti.width x height)
  read from the top left of the decoded frame, each row followed by a halved (scanlines 50%) or black
  (scanlines black) row, at the top left of a black 1024x1024 texture; VideoPSO's bilinear sampler draws
  it on seq_SetDisplaySize's quad: 1x 320x240 and 2x 640x480 centred, full screen 1024x768 (4:3 at
  1024x768); scanlines are disabled when the quad is lower than twice the video (1x, the intelligence
  screen).
* The port: bin\\Warzone.exe --game=CAM_2A (its opening video cam2/c002.ogg, 320x240) with a config that
  sets FMVmode (0 full screen, 1 1x, 2 2x) and scanlines (0 off, 1 50%, 2 black), the 9 pairs, a screenshot
  every 150 frames (tools/ui_tests/video_modes.txt); and the intelligence screen's research film
  (video_intel.txt: a topic researched with 'work harder', its message opened: res_weapons.ogg, a 192x168
  picture in a 192x176 frame, on a 192x168 quad; scanlines 50% asked, so disabled).
* Each screenshot that shows the video is matched to its frame (a coarse search, then the model drawn for
  the best candidates) and compared inside the quad (3 pixels in from its edge, the subtitle bands left
  out). Measured here: 1x and the intelligence screen equal the model exactly; 2x and full screen within
  0.6 of a level (of 255) at the 99th percentile with scanlines off, a few levels with scanlines (the GPU's
  bilinear weights round where the rows alternate).
* --official: the official 4.7.0 build (ref/official/wz470/bin/warzone2100.exe, data from build/data_wz
  when there, else data/) shows the same 9 pairs, captured from outside (PrintWindow, client area, once a
  second) and checked against the same model. (Its intelligence screen needs real mouse clicks: SDL3 takes
  the click position from the cursor, not from posted messages; not run.)
* Output: build/video_display/<run>/ (its screenshots and config folder, remade on each run). About 12
  minutes (each run plays the level's start; the official build's another 9).
"""
import ctypes
import glob
import os
import shutil
import subprocess
import sys
import time
from ctypes import wintypes

import numpy as np
from PIL import Image

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import build  # noqa: E402
import theora_test as tt  # noqa: E402

ROOT = build.ROOT
WORK = os.path.join(ROOT, "build", "video_display")
UI_TESTS = os.path.join(ROOT, "tools", "ui_tests")
EXE = os.path.join(build.BIN, "Warzone.exe")
OFFICIAL_EXE = os.path.join(ROOT, "ref", "official", "wz470", "bin", "warzone2100.exe")
TEX = 1024
SW, SH = 1024, 768
SIZES = {0: ("full", (0, 0, 1024, 768)), 1: ("x1", (352, 264, 320, 240)), 2: ("x2", (192, 144, 640, 480))}
SCANS = {0: "off", 1: "s50", 2: "black"}
SUBTITLES = [(0, 0, 1024, 34), (0, 682, 1024, 738)]        # seqdisp's text bands (top and bottom)
VIDEO = "sequences/cam2/c002.ogg"
INTEL_VIDEO = "sequences/res_weapons.ogg"
INTEL_QUAD = (539, 209, 192, 168)       # intelmap.cpp's FLIC view: INTMAP_FLICWIDTH x INTMAP_FLICHEIGHT


# ---------------------------------------------------------------------------------------------------------
# The model

def decoded(name):
    """(width, height, Y, U, V): every frame as video_write reads it (the picture size from the top left of
    the decoded frame), decoded by the official libtheoradec.dll at its maximum post-processing level."""
    packets = tt.theora_packets(tt.read_video(name))
    _fw, _fh, pw, ph = tt.theora_info(packets)[:4]
    dec = tt.OfficialDecoder()
    lib = dec.lib
    info = tt.th_info()
    tc = tt.th_comment()
    setup = ctypes.c_void_p()
    lib.th_info_init(ctypes.byref(info))
    lib.th_comment_init(ctypes.byref(tc))
    keep = []

    def packet(data, bos, eos):
        buf = ctypes.create_string_buffer(data, max(len(data), 1))
        keep.append(buf)
        op = tt.ogg_packet()
        op.packet = ctypes.cast(buf, ctypes.c_void_p)
        op.bytes = len(data)
        op.b_o_s = bos
        op.e_o_s = eos
        op.granulepos = -1
        return op

    i = 0
    while True:
        r = lib.th_decode_headerin(ctypes.byref(info), ctypes.byref(tc), ctypes.byref(setup),
                                   ctypes.byref(packet(*packets[i])))
        i += 1
        if r == 1:
            break
    d = lib.th_decode_alloc(ctypes.byref(info), setup)
    lib.th_setup_free(setup)
    level = ctypes.c_int(0)
    lib.th_decode_ctl(d, tt.TH_DECCTL_GET_PPLEVEL_MAX, ctypes.byref(level), ctypes.sizeof(level))
    lib.th_decode_ctl(d, tt.TH_DECCTL_SET_PPLEVEL, ctypes.byref(level), ctypes.sizeof(level))
    planes = (tt.th_img_plane * 3)()
    out = [[], [], []]
    sizes = [(pw, ph), ((pw + 1) // 2, (ph + 1) // 2), ((pw + 1) // 2, (ph + 1) // 2)]
    for data, bos, eos in packets[i:]:
        lib.th_decode_packetin(d, ctypes.byref(packet(data, bos, eos)), None)
        keep.clear()
        lib.th_decode_ycbcr_out(d, ctypes.byref(planes))
        for k, pl in enumerate(planes):
            w, h = sizes[k]
            rows = [ctypes.string_at(pl.data + r * pl.stride, w) for r in range(h)]
            out[k].append(np.frombuffer(b"".join(rows), np.uint8).reshape(h, w))
    lib.th_decode_free(d)
    lib.th_comment_clear(ctypes.byref(tc))
    return pw, ph, np.stack(out[0]), np.stack(out[1]), np.stack(out[2])


def rgb(y, u, v):
    """video_write's conversion of one frame (h x w x 3)."""
    h, w = y.shape
    Y = y.astype(np.int64) - 16
    U = np.repeat(np.repeat(u.astype(np.int64) - 128, 2, axis=0), 2, axis=1)[:h, :w]
    V = np.repeat(np.repeat(v.astype(np.int64) - 128, 2, axis=0), 2, axis=1)[:h, :w]
    A = 298 * Y
    C = 409 * V
    R = np.clip((A + C + 128) >> 8, 0, 255)
    G = np.clip((A - 100 * U - (C >> 1) + 128) >> 8, 0, 255)
    B = np.clip((A + 516 * U + 128) >> 8, 0, 255)
    return np.stack([R, G, B], axis=-1)


def texture(img, scan):
    """The texture's rows: the image, or each row followed by its halved (50%) or a black row."""
    if scan == 0:
        return img
    h, w, _ = img.shape
    t = np.zeros((h * 2, w, 3), np.int64)
    t[0::2] = img
    if scan == 1:
        t[1::2] = img >> 1
    return t


def render(tex, quad):
    """VideoPSO's draw of the texture (its th x tw corner of the 1024x1024 texture, the rest black) on the
    quad: bilinear, the pixels whose centres are inside the quad. Returns (screen, covered mask)."""
    x, y, w, h = quad
    th, tw, _ = tex.shape
    big = np.zeros((th + 2, tw + 2, 3))
    big[1:th + 1, 1:tw + 1] = tex
    px = np.arange(SW) + 0.5
    py = np.arange(SH) + 0.5
    cols = (px >= x) & (px < x + w)
    rows = (py >= y) & (py < y + h)
    s = (px[cols] - x) / w * tw - 0.5
    t = (py[rows] - y) / h * th - 0.5
    x0 = np.floor(s).astype(int)
    y0 = np.floor(t).astype(int)
    fx = (s - x0)[None, :, None]
    fy = (t - y0)[:, None, None]
    x0 += 1
    y0 += 1
    a = big[y0][:, x0]
    b = big[y0][:, x0 + 1]
    c = big[y0 + 1][:, x0]
    d = big[y0 + 1][:, x0 + 1]
    img = (a * (1 - fx) + b * fx) * (1 - fy) + (c * (1 - fx) + d * fx) * fy
    out = np.zeros((SH, SW, 3))
    mask = np.zeros((SH, SW), bool)
    out[np.ix_(rows, cols)] = img
    mask[np.ix_(rows, cols)] = True
    return out, mask


class Video:
    def __init__(self, name):
        w, h, Y, U, V = decoded(name)
        self.frames = np.stack([rgb(Y[i], U[i], V[i]) for i in range(len(Y))]).astype(np.uint8)
        self._sigs = {}

    def signatures(self, factor):
        if factor not in self._sigs:
            self._sigs[factor] = np.stack([signature(f.astype(np.float64)) * factor for f in self.frames])
        return self._sigs[factor]


def signature(img, gy=12, gx=16):
    h, w = img.shape[0] // gy * gy, img.shape[1] // gx * gx
    return img[:h, :w].reshape(gy, h // gy, gx, w // gx, 3).mean(axis=(1, 3))


def check_shot(video, path, quad, scan, candidates=8):
    """The screenshot's best frame and its differences from the model inside the quad: a dict (frame, mean,
    p99, max, over8: the share of values more than 8 levels off), or None for a screenshot of another size."""
    x, y, w, h = quad
    scr = np.asarray(Image.open(path).convert("RGB")).astype(np.float64)
    if scr.shape[:2] != (SH, SW):
        return None
    sig = signature(scr[y:y + h, x:x + w])
    errs = ((video.signatures((1.0, 0.75, 0.5)[scan]) - sig) ** 2).mean(axis=(1, 2, 3))
    region = np.zeros((SH, SW), bool)
    m = 3 if w > 200 else 1
    region[y + m:y + h - m, x + m:x + w - m] = True
    for bx0, by0, bx1, by1 in SUBTITLES:
        region[by0:by1, bx0:bx1] = False
    best = None
    for i in np.argsort(errs)[:candidates]:
        model, mask = render(texture(video.frames[i].astype(np.int64), scan), quad)
        d = np.abs(model - scr)[region & mask]
        st = dict(frame=int(i), mean=d.mean(), p99=np.percentile(d, 99), max=d.max(), over8=(d > 8).mean())
        if best is None or st["mean"] < best["mean"]:
            best = st
    return best


def summarise(label, video, shots, quad, scan):
    """Prints one line for a run; returns False when no screenshot showed the video or one is far off."""
    good = []
    for s in shots:
        st = check_shot(video, s, quad, scan)
        if st is not None and st["mean"] < 1.0:        # shows the video (other screens match a dark frame ~2.5)
            good.append(st)
    if not good:
        print("%-20s no screenshot shows the video (%d taken)" % (label, len(shots)))
        return False
    worst = max(g["over8"] for g in good)
    print("%-20s scanlines %s: %2d screenshots on frames %d-%d: mean difference <= %.2f, 99th percentile <= %.2f, "
          "> 8 levels in %.3f%%" % (label, ("off", "50%", "black")[scan], len(good), min(g["frame"] for g in good),
                                   max(g["frame"] for g in good), max(g["mean"] for g in good),
                                   max(g["p99"] for g in good), 100 * worst))
    return worst < 0.01 and max(g["mean"] for g in good) < 1.0


# ---------------------------------------------------------------------------------------------------------
# Runs

def fresh_config(run, fmv, scan):
    folder = os.path.join(WORK, run)
    if os.path.exists(folder):
        shutil.rmtree(folder)
    cfg = os.path.join(folder, "config")
    os.makedirs(cfg)
    with open(os.path.join(cfg, "config"), "w", newline="\r\n") as f:
        f.write("[General]\nFMVmode=%d\nscanlines=%d\nwidth=1024\nheight=768\nfullscreen=0\n" % (fmv, scan))
    return folder, cfg


def run_port(run, fmv, scan, script):
    folder, cfg = fresh_config(run, fmv, scan)
    p = subprocess.run([EXE, "--configdir=" + cfg, "--vbscript=" + os.path.join(UI_TESTS, script), "--nosound",
                        "--game=CAM_2A", "--debug=video"], cwd=build.BIN, timeout=900)
    shots = []
    for f in sorted(glob.glob(os.path.join(cfg, "screenshots", "*.png"))):
        dst = os.path.join(folder, os.path.basename(f).split("-", 2)[-1])
        shutil.copy(f, dst)
        shots.append(dst)
    if p.returncode != 0:
        print("%s: exit code %d" % (run, p.returncode))
    return sorted(shots)


user32 = ctypes.WinDLL("user32", use_last_error=True)
gdi32 = ctypes.WinDLL("gdi32")
WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)


class BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [("biSize", wintypes.DWORD), ("biWidth", wintypes.LONG), ("biHeight", wintypes.LONG),
                ("biPlanes", wintypes.WORD), ("biBitCount", wintypes.WORD), ("biCompression", wintypes.DWORD),
                ("biSizeImage", wintypes.DWORD), ("biXPelsPerMeter", wintypes.LONG),
                ("biYPelsPerMeter", wintypes.LONG), ("biClrUsed", wintypes.DWORD), ("biClrImportant", wintypes.DWORD)]


def find_window(pid):
    found = []

    def cb(hwnd, _):
        p = wintypes.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(p))
        if p.value == pid and user32.IsWindowVisible(hwnd):
            found.append(hwnd)
        return True
    user32.EnumWindows(WNDENUMPROC(cb), 0)
    return found[0] if found else None


def capture(hwnd):
    rc = wintypes.RECT()
    user32.GetClientRect(hwnd, ctypes.byref(rc))
    w, h = rc.right, rc.bottom
    hdc = user32.GetDC(hwnd)
    mdc = gdi32.CreateCompatibleDC(hdc)
    bmp = gdi32.CreateCompatibleBitmap(hdc, w, h)
    gdi32.SelectObject(mdc, bmp)
    user32.PrintWindow(hwnd, mdc, 3)            # PW_CLIENTONLY | PW_RENDERFULLCONTENT
    bi = BITMAPINFOHEADER(ctypes.sizeof(BITMAPINFOHEADER), w, -h, 1, 32, 0, 0, 0, 0, 0, 0)
    buf = ctypes.create_string_buffer(w * h * 4)
    gdi32.GetDIBits(mdc, bmp, 0, h, buf, ctypes.byref(bi), 0)
    gdi32.DeleteObject(bmp)
    gdi32.DeleteDC(mdc)
    user32.ReleaseDC(hwnd, hdc)
    return Image.frombuffer("RGBA", (w, h), buf.raw, "raw", "BGRA", 0, 1).convert("RGB")


def run_official(run, fmv, scan, seconds=55):
    folder, cfg = fresh_config(run, fmv, scan)
    data = os.path.join(ROOT, "build", "data_wz")
    if not os.path.isdir(data):
        data = os.path.join(ROOT, "data")
    p = subprocess.Popen([OFFICIAL_EXE, "--configdir=" + cfg, "--datadir=" + data, "--nosound", "--game=CAM_2A"],
                         cwd=os.path.dirname(OFFICIAL_EXE))
    t0 = time.time()
    hwnd = None
    shots = []
    while time.time() - t0 < seconds:
        if hwnd is None:
            hwnd = find_window(p.pid)
        if hwnd is not None:
            path = os.path.join(folder, "o%03d.png" % len(shots))
            capture(hwnd).save(path)
            shots.append(path)
        time.sleep(1.0)
    p.kill()
    p.wait()
    return shots


def main(argv):
    official = "--official" in argv
    os.makedirs(WORK, exist_ok=True)
    print("decoding the model's videos (the official libtheoradec.dll)...", flush=True)
    video = Video(VIDEO)
    ok = True
    for fmv in (0, 2, 1):
        for scan in (0, 1, 2):
            size, quad = SIZES[fmv]
            effective = scan if quad[3] >= 480 else 0      # disabled below twice the video's height
            run = "port_%s_%s" % (size, SCANS[scan])
            shots = run_port(run, fmv, scan, "video_modes.txt")
            ok &= summarise(run, video, shots, quad, effective)
    intel = Video(INTEL_VIDEO)
    shots = run_port("port_intel", 1, 1, "video_intel.txt")
    ok &= summarise("port_intel", intel, [s for s in shots if os.path.basename(s).startswith("f")], INTEL_QUAD, 0)
    if official:
        if not os.path.exists(OFFICIAL_EXE):
            print("--official: %s not found" % OFFICIAL_EXE)
            return 1
        for fmv in (0, 2, 1):
            for scan in (0, 1, 2):
                size, quad = SIZES[fmv]
                effective = scan if quad[3] >= 480 else 0
                run = "official_%s_%s" % (size, SCANS[scan])
                summarise(run, video, run_official(run, fmv, scan), quad, effective)
    print("video display: %s" % ("as the model of 4.7.0's" if ok else "DIFFERENCES (see above)"))
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
