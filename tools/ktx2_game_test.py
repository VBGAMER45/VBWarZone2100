"""
ktx2_game_test.py - cluster T's game tests of the KTX2 textures (fork T3): the WZ side of 4.7.0's BASIS_ENABLED build
(lib/ivis_opengl/gfx_api_image_basis_priv.cpp, iV_BaseImage / iV_CompressedImage, gfx_api.cpp's KTX2 branches) over
WzBasis's transcoder, and the Remastered terrain quality (data/terrain_overrides/high.wz, the official pack).

  python -I tools\\ktx2_game_test.py <test> [--exe EXE] [--levels=CAM_1A,...] [--keep]

* --exe: the Warzone.exe to run (default bin\\Warzone.exe). Another one (a private build's) finds the data as an
  installed game does, in <its folder>\\..\\data: the test makes that a mirror of hard links to data\\'s files (not
  --datadir: that folder would come before the config folder and its terrain pack in init.cpp's search path order).
* Every run gets a fresh config folder build\\ktx2_game_test\\<test>\\<run>\\config (test-safe: portMapping=0,
  lobbyserver=http://127.0.0.1:9; 1024x768 windowed, display scale 100); the Remastered runs put a hard link of the
  official high.wz in it (config\\terrain_overrides\\high.wz: the config folder is a search path, ahead of the
  <prefix>\\data folder, so the pack is found whether data\\ has it or not; data\\terrain_overrides\\high, an empty
  folder from the source tree, would otherwise count as the pack: rebuildSearchPath stops at the first folder whose
  terrain_overrides/high or high.wz mounts, and an empty folder mounts). The script commands are tools/ui_test.py's (--vbscript). Each run prints its
  exit code, time, the log's error / fatal / assert lines and its KTX2 lines.

tests:
  regress   the iV_BaseImage change: the --vbdrawtest screenshot against tools/draw_test/drawtest.png (0 pixels
            should differ), a skirmish from the menu (tools/ui_tests/skirmish.txt), CAM_1A's start (campaign.txt:
            a video), a save and its load (save_menu.txt, then save_load.txt on that config): no errors or asserts
  failpath  a skirmish (s7_terrain.txt) with terrainMode=2 (Remastered) and high.wz: before WzBasis's transcoder is
            merged (its stubs fail every init) the KTX2 loads fail as 4.7.0 logs a failed transcoder ("Failed to
            initialize ktx2_transcoder for file: ...")
  options   Options > Graphics > Terrain Appearance with high.wz: Remastered (HQ) offered (screenshot)
  switch    Options > Graphics > Terrain Appearance > Remastered (HQ), back, a skirmish started (s7's clicks)
  tilesets  CAM_1A, CAM_2A, CAM_3A (Arizona, Urban, Rockies) with terrainMode=2 and terrainShadingQuality 512 and
            1024, in the port (its own screenshot, 10 s after the level runs: the guide's pop-up pauses CAM_1A) and
            in the official build (PrintWindow of its client area, its own data folder ref\\official\\wz470\\data,
            which has high.wz; test-safe config; nothing clicked): build\\ktx2_game_test\\tilesets\\<level>_<size>_
            {port,official}.png and a side-by-side <level>_<size>_both.png
  pngtilesets  tilesets with a PNG copy of the pack (each KTX2 file's top level through ktx2ref's RGBA32 transcode,
            build\\ktx2_game_test\\high_png) as a folder in the config folder: terrain.cpp's NORMAL_MAPPING paths
            without WzBasis (before the merge), next to the official build's KTX2 Remastered
  pngswitch switch with that PNG pack
  normaltilesets  tilesets at the Normal quality (terrainMode=1, no pack): the ground pages page-80..109 are KTX2 in
            4.7.0's base.wz (and in data\\ as tools\\setup_data.py lays it out), so this runs the KTX2 array path
            and the 2D one (the water pages) at Normal
  dump      TEST-ONLY WZ_TEST_KTX2_DUMP: CAM_1A (Arizona) and CAM_3A (Rockies) with Remastered at 1024, with
            textureCompression 1 (BC7 here, R8 / RGB8 uncompressed for the single-channel and normal maps) and 0
            (RGBA32 everywhere): every level the port transcoded is compared byte for byte with
            build\\ktx2ref\\ktx2ref.exe's output for the same file and format, and every image handed to an upload
            (after the channel conversion: WZ_TEST_KTX2_DUMP's .U files) with 4.7.0's conversion of ktx2ref's
            output; no asserts
  perf    CAM_1A: the time from the .gam load to the level running and the process's peak private bytes (32-bit
            process), terrainMode 1 (Normal) against 2 (Remastered) at 512 and 1024
  all       regress, options, switch, tilesets, dump, perf
"""
import ctypes
import ctypes.wintypes as wt
import glob
import os
import shutil
import struct
import subprocess
import sys
import time

ROOT = r"C:\Users\Owner\desktop\websites\warzone"
WORK = os.path.join(ROOT, "build", "ktx2_game_test")
HIGH_WZ = os.path.join(ROOT, "ref", "official", "wz470", "data", "terrain_overrides", "high.wz")
OFFICIAL = os.path.join(ROOT, "ref", "official", "wz470", "bin", "warzone2100.exe")
OFFICIAL_DATA = os.path.join(ROOT, "ref", "official", "wz470", "data")
KTX2REF = os.path.join(ROOT, "build", "ktx2ref", "ktx2ref.exe")
HIGH_DIR = os.path.join(ROOT, "build", "ktx2ref", "high")      # high.wz extracted (cluster T's ktx2ref)
UI_TESTS = os.path.join(ROOT, "tools", "ui_tests")
TEST_CONFIG = "[General]\nportMapping=0\nlobbyserver=http://127.0.0.1:9\nwidth=1024\nheight=768\nfullscreen=0\ndisplayScale=100\n"

args = sys.argv[1:]
exe = os.path.join(ROOT, "bin", "Warzone.exe")
if "--exe" in args:
    exe = os.path.abspath(args[args.index("--exe") + 1])
    del args[args.index("--exe"):args.index("--exe") + 2]
args = [a for a in args if a != "--keep"]
LEVELS = ("CAM_1A", "CAM_2A", "CAM_3A")
for a in list(args):
    if a.startswith("--levels="):
        LEVELS = tuple(a.split("=", 1)[1].split(","))
        args.remove(a)

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import setup_data  # noqa: E402  (the data layout: the base textures 4.7.0 ships as KTX2)
if os.path.normcase(setup_data.ROOT) != os.path.normcase(ROOT):    # (a copy of the two tools outside tools\)
    for _name in ("REF_DATA", "MIRROR", "OFFICIAL", "VIDEOS", "BASE_KTX2"):
        setattr(setup_data, _name, os.path.join(ROOT, os.path.relpath(getattr(setup_data, _name), setup_data.ROOT)))

kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
psapi = ctypes.WinDLL("psapi")


class PROCESS_MEMORY_COUNTERS_EX(ctypes.Structure):
    _fields_ = [("cb", wt.DWORD), ("PageFaultCount", wt.DWORD), ("PeakWorkingSetSize", ctypes.c_size_t),
                ("WorkingSetSize", ctypes.c_size_t), ("QuotaPeakPagedPoolUsage", ctypes.c_size_t),
                ("QuotaPagedPoolUsage", ctypes.c_size_t), ("QuotaPeakNonPagedPoolUsage", ctypes.c_size_t),
                ("QuotaNonPagedPoolUsage", ctypes.c_size_t), ("PagefileUsage", ctypes.c_size_t),
                ("PeakPagefileUsage", ctypes.c_size_t), ("PrivateUsage", ctypes.c_size_t)]


def peak_private(pid):
    """The process's peak commit (private bytes) so far, or None."""
    h = kernel32.OpenProcess(0x1000 | 0x0010, False, pid)  # PROCESS_QUERY_LIMITED_INFORMATION | PROCESS_VM_READ
    if not h:
        return None
    try:
        c = PROCESS_MEMORY_COUNTERS_EX()
        c.cb = ctypes.sizeof(c)
        if psapi.GetProcessMemoryInfo(h, ctypes.byref(c), c.cb):
            return c.PeakPagefileUsage
        return None
    finally:
        kernel32.CloseHandle(h)


def is_bin(x):
    return os.path.normcase(os.path.abspath(x)) == os.path.normcase(os.path.join(ROOT, "bin", "Warzone.exe"))


def ensure_data(x, skip=()):
    """<x's folder>\\..\\data, for an exe outside bin\\ (see --exe): hard links to data\\'s files but the names in
    skip, laid out as tools\\setup_data.py lays out data\\ (the base textures 4.7.0 ships as KTX2 in base.wz: their
    KTX2 files, not the source PNGs), whether data\\ has that layout yet or not. Made once (a marker file)."""
    if is_bin(x):
        return
    target = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(x))), "data")
    marker = os.path.join(target, "mirror_complete.txt")   # made once: a later call keeps its skips
    if os.path.exists(marker):
        return
    src = os.path.join(ROOT, "data")
    left_out = {os.path.normcase(rel) for rel in setup_data.converted_textures()}
    setup_data.extract_base_ktx2()
    n = 0
    for folder, _dirs, files in os.walk(src):
        out = os.path.join(target, os.path.relpath(folder, src))
        os.makedirs(out, exist_ok=True)
        for f in files:
            d = os.path.join(out, f)
            rel = os.path.normcase(os.path.normpath(os.path.join(os.path.relpath(folder, src), f)))
            if f not in skip and rel not in left_out and not os.path.exists(d):
                os.link(os.path.join(folder, f), d)
                n += 1
    for name in setup_data.base_ktx2_names():
        d = os.path.join(target, "base", *name.split("/"))
        if not os.path.exists(d):
            os.makedirs(os.path.dirname(d), exist_ok=True)
            os.link(os.path.join(setup_data.BASE_KTX2, *name.split("/")), d)
            n += 1
    open(marker, "w").write("without: %s\n" % ", ".join(skip))
    if n:
        print("  %s: %d files linked from data\\" % (os.path.relpath(target, ROOT), n))


def exe_copy(x, name, skip=()):
    """A copy of x and the DLLs and manifests beside it in build\\ktx2_game_test\\<name>\\bin, its data a mirror
    (ensure_data, without the names in skip). The copy's path."""
    out = os.path.join(WORK, name, "bin")
    os.makedirs(out, exist_ok=True)
    here = os.path.dirname(os.path.abspath(x))
    for f in glob.glob(os.path.join(here, "*.dll")) + glob.glob(os.path.join(here, "*.manifest")) + [x]:
        shutil.copyfile(f, os.path.join(out, os.path.basename(f)))
    copy = os.path.join(out, os.path.basename(x))
    ensure_data(copy, skip=skip)
    return copy


def exe_tag(x):
    """bin, or the name of the build folder the exe is in (j_T3b for build\\j_T3b\\bin\\Warzone.exe)."""
    return "bin" if is_bin(x) else os.path.basename(os.path.dirname(os.path.dirname(os.path.abspath(x))))


def novideo_exe(x):
    """A copy of x (exe_copy, build\\ktx2_game_test\\novideo_<exe_tag>) whose data mirror has no sequences.wz: the
    campaign levels start without their videos, as the official build does from its own data folder (which has no
    sequences.wz), so the two reach the same moments."""
    copy = exe_copy(x, "novideo_" + exe_tag(x), skip=("sequences.wz",))
    stray = os.path.join(os.path.dirname(os.path.dirname(copy)), "data", "sequences.wz")
    if os.path.exists(stray):
        os.remove(stray)                                # (a hard link: the file itself stays in data\)
    return copy


def run_port(test, run, script_lines=None, script_file=None, config="", high=False, extra=(), timeout=300, env=None,
             cfg=None, watch=(), x=None):
    """Runs the port with a script; returns (exit code, log lines, screenshots {name: path}, peak private bytes,
    {watched text: seconds after the start when its first log line appeared})."""
    work = os.path.join(WORK, test, run)
    if cfg is None:
        cfg = os.path.join(work, "config")
        if os.path.exists(cfg):
            shutil.rmtree(cfg)
        os.makedirs(cfg)
        with open(os.path.join(cfg, "config"), "w", newline="\r\n") as f:
            f.write(TEST_CONFIG + "".join(line + "\n" for line in config.split(";") if line))
    if high == "png":
        # the PNG copy of the pack as a folder (config\terrain_overrides\high\...: hard links)
        src = build_png_pack()
        for folder, _dirs, files in os.walk(src):
            out = os.path.join(cfg, "terrain_overrides", "high", os.path.relpath(folder, src))
            os.makedirs(out, exist_ok=True)
            for f in files:
                if not os.path.exists(os.path.join(out, f)):
                    os.link(os.path.join(folder, f), os.path.join(out, f))
    elif high:
        d = os.path.join(cfg, "terrain_overrides")
        os.makedirs(d, exist_ok=True)
        if not os.path.exists(os.path.join(d, "high.wz")):
            os.link(HIGH_WZ, os.path.join(d, "high.wz"))
    os.makedirs(work, exist_ok=True)
    if script_lines is not None:
        script_file = os.path.join(work, "script.txt")
        with open(script_file, "w", newline="\r\n") as f:
            f.write("\n".join(script_lines) + "\n")
    elif not os.path.isabs(script_file):
        script_file = os.path.join(UI_TESTS, script_file)
    x = x or exe
    cmd = [x, "--configdir=" + cfg, "--vbscript=" + script_file, "--nosound"] + list(extra)
    ensure_data(x)
    for old in glob.glob(os.path.join(work, "*.png")):
        os.remove(old)
    t0 = time.time()
    p = subprocess.Popen(cmd, cwd=os.path.dirname(x), env=env)
    peak = 0
    seen = {}
    logpath = None
    while True:
        try:
            p.wait(timeout=0.2)
            break
        except subprocess.TimeoutExpired:
            pass
        m = peak_private(p.pid)
        if m:
            peak = max(peak, m)
        if watch:
            if logpath is None:
                logs = sorted(glob.glob(os.path.join(cfg, "logs", "WZlog-*.txt")), key=os.path.getmtime)
                logpath = logs[-1] if logs and os.path.getmtime(logs[-1]) >= t0 - 1 else None
            if logpath:
                text = open(logpath, encoding="utf-8", errors="replace").read()
                for w in watch:
                    if w not in seen and w in text:
                        seen[w] = time.time() - t0
        if time.time() - t0 > timeout:
            p.kill()
            p.wait()
            print("  stopped after %.0f s (limit)" % (time.time() - t0))
            break
    logs = sorted(glob.glob(os.path.join(cfg, "logs", "WZlog-*.txt")), key=os.path.getmtime)
    lines = open(logs[-1], encoding="utf-8", errors="replace").read().splitlines() if logs else []
    shots = {}
    for shot in sorted(glob.glob(os.path.join(cfg, "screenshots", "*.png"))):
        if os.path.getmtime(shot) < t0:
            continue
        base = os.path.basename(shot)
        name = base.split("-", 2)[2][:-4] if base.count("-") >= 2 else base[:-4]
        out = os.path.join(work, name + ".png")
        shutil.copyfile(shot, out)
        shots[name] = out
    bad = [ln for ln in lines if ln.startswith(("error", "fatal")) or "assert" in ln.lower()]
    ktx = [ln for ln in lines if ("ktx2" in ln.lower() or "basis" in ln.lower()) and not ln.startswith("--- Starting log")]
    print("  %s/%s: exit %d after %.0f s, %d log lines, %d error/assert lines, %d KTX2 lines, peak private %.0f MB"
          % (test, run, p.returncode, time.time() - t0, len(lines), len(bad), len(ktx), peak / 1048576.0))
    for ln in bad[:12]:
        print("    | " + ln[:200])
    if len(bad) > 12:
        print("    | ... (%d more)" % (len(bad) - 12))
    for ln in ktx[:6]:
        if ln not in bad:
            print("    k " + ln[:200])
    return p.returncode, lines, shots, peak, seen


# ---------------------------------------------------------------------------------------------------- the tests

def regress():
    from PIL import Image, ImageChops
    ok = True

    def differing(a, b):
        if a.size != b.size:
            return -1
        r, g, bl = ImageChops.difference(a, b).split()
        return sum(ImageChops.lighter(ImageChops.lighter(r, g), bl).histogram()[1:])

    # the draw test (tools/draw_test.py's run) with copies of this exe and of bin\Warzone.exe beside the same data
    # layout (exe_copy): tools/draw_test/drawtest.png is older than clusters F to S (161,176 pixels differ for bin\
    # itself), so the check is this exe against bin\'s. bin\Warzone.exe on data\ itself is also run: the backdrop
    # there is the source PNG while data\ has not the KTX2 layout yet (informational)
    shots = {}
    binexe = os.path.join(ROOT, "bin", "Warzone.exe")
    for which, x in (("this", exe_copy(exe, "regress_this")), ("bin", exe_copy(binexe, "regress_bin")),
                     ("bin_on_data", binexe)):
        cfg = os.path.join(WORK, "regress", "drawtest_" + which, "config")
        if os.path.exists(cfg):
            shutil.rmtree(cfg)
        os.makedirs(cfg)
        cmd = [x, "--configdir=" + cfg, "--nosound", "--vbdrawtest=drawtest"]
        subprocess.run(cmd, cwd=os.path.dirname(x), timeout=120)
        found = sorted(glob.glob(os.path.join(cfg, "screenshots", "*-drawtest.png")))
        shots[which] = Image.open(found[-1]).convert("RGB") if found else None
    if shots["this"] is None or shots["bin"] is None:
        print("  drawtest: no screenshot")
        ok = False
    else:
        n = differing(shots["this"], shots["bin"])
        ref = differing(shots["this"], Image.open(os.path.join(ROOT, "tools", "draw_test", "drawtest.png")).convert("RGB"))
        print("  drawtest: %d pixels differ from bin\\Warzone.exe's (%d from the old tools/draw_test/drawtest.png)" % (n, ref))
        if shots["bin_on_data"] is not None:
            mean, big = view_difference(shots["this"], shots["bin_on_data"], (0, 0) + shots["this"].size)
            print("  drawtest: against bin\\Warzone.exe on data\\ itself: mean difference %.2f, %.2f%% of pixels off "
                  "by more than 24 (the backdrop: KTX2 here, the source PNG there while data\\ has its old layout)"
                  % (mean, 100 * big))
        ok = ok and n == 0
    for run, script, t in (("skirmish", "skirmish.txt", 300), ("campaign", "campaign.txt", 400)):
        rc, lines, shots, _, _ = run_port("regress", run, script_file=script, timeout=t)
        ok = ok and rc == 0 and not any(ln.startswith(("error", "fatal")) or "assert" in ln.lower() for ln in lines)
    rc, lines, shots, _, _ = run_port("regress", "save_menu", script_file="save_menu.txt", timeout=400)
    ok = ok and rc == 0
    cfg = os.path.join(WORK, "regress", "save_menu", "config")
    rc, lines, shots, _, _ = run_port("regress", "save_load", script_file="save_load.txt", cfg=cfg, timeout=400)
    ok = ok and rc == 0 and not any(ln.startswith(("error", "fatal")) or "assert" in ln.lower() for ln in lines)
    print("regress: %s" % ("passed" if ok else "FAILED"))
    return ok


def failpath():
    rc, lines, shots, _, _ = run_port("failpath", "skirmish", script_file="s7_terrain.txt", config="terrainMode=2",
                                      high=True, timeout=300)
    fails = [ln for ln in lines if "Failed to initialize ktx2_transcoder for file" in ln]
    print("failpath: %d 'Failed to initialize ktx2_transcoder' lines (first: %s)" % (len(fails), fails[0][:150] if fails else "-"))
    return bool(fails)


OPTIONS_CLICKS = ["wait 90", "click 512 470", "wait 90", "click 530 87", "wait 60", "click 700 201", "wait 30"]


def options():
    rc, lines, shots, _, _ = run_port("options", "choices", script_lines=OPTIONS_CLICKS + ["shot terrain_choices", "quit"],
                                      high=True, timeout=120)
    print("options: screenshot %s (Remastered (HQ) should be enabled)" % shots.get("terrain_choices"))
    return rc == 0


def switch(pack=True, test="switch"):
    script = OPTIONS_CLICKS + ["click 710 280", "wait 120", "shot remastered_chosen", "click 215 87", "wait 90",
                               "click 512 365", "wait 60", "click 512 400", "wait 120", "click 365 455", "wait 120",
                               "click 775 90", "wait 600", "shot ingame", "quit"]
    rc, lines, shots, peak, _ = run_port(test, "to_remastered", script_lines=script, high=pack, timeout=600)
    cfgtext = open(os.path.join(WORK, test, "to_remastered", "config", "config"), encoding="utf-8", errors="replace").read()
    mode = [ln for ln in cfgtext.splitlines() if ln.startswith("terrainMode")]
    print("switch: saved %s; screenshots %s" % (mode, sorted(shots)))
    return rc == 0


def level_script(level, shot):
    return ["waitlevel " + level, "quiet 360", "log level_running", "wait 600", "shot " + shot, "quit"]


def official_capture(outdir, seconds, config, extra):
    """The official build with a test-safe config and its own data folder; PrintWindow captures at the given times."""
    from PIL import Image
    user32 = ctypes.WinDLL("user32")
    gdi32 = ctypes.WinDLL("gdi32")
    cfg = os.path.join(outdir, "official_config")
    if os.path.exists(cfg):
        shutil.rmtree(cfg)
    os.makedirs(cfg)
    with open(os.path.join(cfg, "config"), "w", newline="\r\n") as f:
        f.write(TEST_CONFIG + "".join(line + "\n" for line in config.split(";") if line))
    p = subprocess.Popen([OFFICIAL, "--configdir=" + cfg, "--datadir=" + OFFICIAL_DATA, "--nosound"] + list(extra),
                         cwd=os.path.dirname(OFFICIAL))
    WNDENUMPROC = ctypes.WINFUNCTYPE(wt.BOOL, wt.HWND, wt.LPARAM)

    def find_window(pid):
        found = []

        def cb(hwnd, _):
            q = wt.DWORD()
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(q))
            if q.value == pid and user32.IsWindowVisible(hwnd):
                found.append(hwnd)
            return True
        user32.EnumWindows(WNDENUMPROC(cb), 0)
        return found[0] if found else None

    class BITMAPINFOHEADER(ctypes.Structure):
        _fields_ = [("biSize", wt.DWORD), ("biWidth", wt.LONG), ("biHeight", wt.LONG), ("biPlanes", wt.WORD),
                    ("biBitCount", wt.WORD), ("biCompression", wt.DWORD), ("biSizeImage", wt.DWORD),
                    ("biXPelsPerMeter", wt.LONG), ("biYPelsPerMeter", wt.LONG), ("biClrUsed", wt.DWORD),
                    ("biClrImportant", wt.DWORD)]

    def capture(hwnd):
        rc = wt.RECT()
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

    t0 = time.time()
    out = []
    try:
        for k, t in enumerate(seconds):
            while time.time() - t0 < t:
                time.sleep(0.1)
            h = find_window(p.pid)
            if h:
                path = os.path.join(outdir, "official_%d.png" % k)
                capture(h).save(path)
                out.append(path)
    finally:
        p.kill()
        p.wait()
    logs = sorted(glob.glob(os.path.join(cfg, "logs", "WZlog-*.txt")))
    return out, (open(logs[-1], encoding="utf-8", errors="replace").read().splitlines() if logs else [])


def build_png_pack():
    """build\\ktx2_game_test\\high_png: high.wz with every KTX2 file as a PNG of its top level (ktx2ref's RGBA32
    transcode) and the tileset files as they are: the Remastered pack without KTX2, so that terrain.cpp's
    NORMAL_MAPPING paths can run before WzBasis transcodes (the PNG loaders take it: a .png comes before a .ktx2)."""
    from PIL import Image
    out = os.path.join(WORK, "high_png")
    done = os.path.join(out, "complete.txt")
    if os.path.exists(done):
        return os.path.join(out, "pack")
    tmp = os.path.join(out, "rgba")
    for folder, _dirs, files in os.walk(HIGH_DIR):
        rel = os.path.relpath(folder, HIGH_DIR)
        dst = os.path.join(out, "pack", rel)
        os.makedirs(dst, exist_ok=True)
        ktx = [f for f in files if f.endswith(".ktx2")]
        for f in files:
            if not f.endswith(".ktx2"):
                shutil.copyfile(os.path.join(folder, f), os.path.join(dst, f))
        if not ktx:
            continue
        t = os.path.join(tmp, rel)
        os.makedirs(t, exist_ok=True)
        subprocess.run([KTX2REF, "13", t] + [os.path.join(folder, f) for f in ktx], capture_output=True, check=True)
        for f in ktx:
            hdr = open(os.path.join(folder, f), "rb").read(28)
            w, h = struct.unpack_from("<II", hdr, 20)
            data = open(os.path.join(t, f + ".13"), "rb").read(w * h * 4)
            Image.frombytes("RGBA", (w, h), data).save(os.path.join(dst, f[:-5] + ".png"))
    shutil.rmtree(tmp)
    open(done, "w").write("ok\n")
    return os.path.join(out, "pack")


VIEW_BOX = (0, 60, 600, 620)     # the view left of CAM_1A's guide pop-up, above the bottom bar
# the largest mean difference per channel a tilesets run passes with: the terrain matched, units and effects moving
# (6.7 the largest seen when it matched; 28 to 37 with the wrong ground pages)
TILESET_MAX_MEAN = 10


def view_difference(a, b, box=VIEW_BOX):
    """(mean absolute difference per channel, share of pixels off by more than 24 in a channel) over box."""
    from PIL import ImageChops, ImageStat
    diff = ImageChops.difference(a.crop(box), b.crop(box))
    r, g, bl = diff.split()
    m = ImageChops.lighter(ImageChops.lighter(r, g), bl)
    return sum(ImageStat.Stat(diff).mean) / 3, sum(m.histogram()[25:]) / float(m.width * m.height)


def tilesets(pack=True, test="tilesets", mode=2):
    from PIL import Image
    outdir = os.path.join(WORK, test)
    os.makedirs(outdir, exist_ok=True)
    from PIL import ImageStat
    x = novideo_exe(exe)
    ok = True
    for size in (512, 1024):
        for level in LEVELS:
            tag = "%s_%d" % (level, size)
            config = "terrainMode=%d;terrainShadingQuality=%d" % (mode, size)
            rc, lines, shots, peak, _ = run_port(test, tag, script_lines=level_script(level, "view"), config=config,
                                                 high=pack, extra=["--game=" + level], timeout=600, x=x)
            port = os.path.join(outdir, tag + "_port.png")
            if "view" in shots:
                shutil.copyfile(shots["view"], port)
            offdir = os.path.join(outdir, tag + "_official")
            os.makedirs(offdir, exist_ok=True)
            # several moments: the one closest to the port's shot is kept (the loads take different times)
            caps, olines = official_capture(offdir, [12, 15, 18, 21, 24, 27, 30], config, ["--game=" + level])
            # a capture whose view is black is no reference (the official build sometimes draws no 3D view into
            # PrintWindow for a whole run: run the level again)
            caps = [c for c in caps if max(ImageStat.Stat(Image.open(c).convert("RGB").crop(VIEW_BOX)).mean) >= 5]
            off = os.path.join(outdir, tag + "_official.png")
            result = "no comparison (no official capture with a view drawn: run it again)"
            fine = False
            if caps and os.path.exists(port):
                a = Image.open(port).convert("RGB")
                scored = sorted((view_difference(a, Image.open(c).convert("RGB")), c) for c in caps)
                (mean, big), best = scored[0]
                shutil.copyfile(best, off)
                b = Image.open(off).convert("RGB")
                both = Image.new("RGB", (a.width + b.width, max(a.height, b.height)))
                both.paste(a, (0, 0))
                both.paste(b, (a.width, 0))
                both.save(os.path.join(outdir, tag + "_both.png"))
                result = "view: mean difference %.2f, %.1f%% of pixels off by more than 24 (official capture %s)" % (
                    mean, 100 * big, os.path.basename(best))
                fine = mean <= TILESET_MAX_MEAN
            asserts = [ln for ln in lines if "assert" in ln.lower()]
            fine = fine and rc == 0 and not asserts
            ok = ok and fine
            print("%s %s: %s%s%s" % (test, tag, result, "; %d assert lines" % len(asserts) if asserts else "",
                                      "" if fine else "  FAILED"))
    return ok


def pngtilesets():
    return tilesets("png", "pngtilesets")


def pngswitch():
    return switch("png", "pngswitch")


def normaltilesets():
    return tilesets(False, "normaltilesets", 1)


def ktx2_levels(path, fmt):
    """The byte ranges of each level of ktx2ref's output for this file and format."""
    d = open(path, "rb").read(80)
    w, h = struct.unpack_from("<II", d, 20)
    levels = struct.unpack_from("<I", d, 40)[0]
    bpb = {2: 8, 4: 8, 13: 4}.get(fmt, 16)          # BC1, BC4: 8 bytes a block; RGBA32: 4 a pixel; the rest 16
    out, at = [], 0
    for k in range(max(1, levels)):
        ow, oh = max(1, w >> k), max(1, h >> k)
        n = ow * oh * 4 if fmt == 13 else ((ow + 3) // 4) * ((oh + 3) // 4) * bpb
        out.append((at, n))
        at += n
    return out


# the port's pixel_format (GfxApi's enum) of a compressed upload: the transcoder_texture_format ktx2ref takes
PF_TO_TF = {6: 2, 8: 3, 9: 4, 10: 5, 11: 6, 12: 0, 14: 1, 15: 20, 16: 21, 17: 10}
PF_NAMES = {1: "RGBA8", 3: "RGB8", 5: "R8", 6: "BC1", 8: "BC3", 9: "BC4", 10: "BC5", 11: "BC7", 12: "ETC1", 14: "ETC2_EAC",
            15: "R11_EAC", 16: "RG11_EAC", 17: "ASTC_4x4"}


def ktx2_source(base):
    """The KTX2 file a dumped name stands for: high.wz's (mounted first), else base.wz's (BASE_KTX2)."""
    for root in (HIGH_DIR, setup_data.BASE_KTX2):
        src = os.path.join(root, *base.split("#"))
        if os.path.exists(src):
            return src
    return None


def ktx2_ref_level(src, tf, w, h, refdir):
    """ktx2ref's output for the level of src whose size is w x h, in transcoder format tf (None: no such level)."""
    out = os.path.join(refdir, os.path.basename(src) + "." + str(tf))
    if not os.path.exists(out):
        subprocess.run([KTX2REF, str(tf), refdir, src], capture_output=True)
    ref = open(out, "rb").read()
    hdr = open(src, "rb").read(28)
    W, H = struct.unpack_from("<II", hdr, 20)
    for k, (at, n) in enumerate(ktx2_levels(src, tf)):
        if (max(1, W >> k), max(1, H >> k)) == (w, h):
            return ref[at:at + n]
    return None


def compare_uploads(dd, refdir):
    """The upload dumps (<name>.U<pixel_format>.<w>x<h>.L<level>: what the port handed to upload / upload_layer, after
    iVImage_Basis_Convert_Channels) against 4.7.0's: ktx2ref's transcode of the same level, for an uncompressed
    upload through 4.7.0's conversion (gfx_api_image_basis_priv.cpp:265-299: R8 is channel 0 of RGBA32,
    iV_Image::convert_to_single_channel; RGB8 channels 0, 1, 2, convert_channels({0, 1, 2}))."""
    import re
    same = differ = missing = 0
    formats = {}
    for name in sorted(os.listdir(dd)):
        m = re.match(r"^(.*)\.U(\d+)\.(\d+)x(\d+)\.L(\d+)$", name)
        if not m:
            continue
        base, pf, w, h = m.group(1), int(m.group(2)), int(m.group(3)), int(m.group(4))
        src = ktx2_source(base)
        if src is None:
            missing += 1
            continue
        formats[PF_NAMES.get(pf, pf)] = formats.get(PF_NAMES.get(pf, pf), 0) + 1
        if pf in PF_TO_TF:
            want = ktx2_ref_level(src, PF_TO_TF[pf], w, h, refdir)
        else:
            rgba = ktx2_ref_level(src, 13, w, h, refdir)
            want = None if rgba is None else {1: rgba, 3: bytes(b for i, b in enumerate(rgba) if i % 4 != 3),
                                              5: rgba[0::4]}.get(pf)
        got = open(os.path.join(dd, name), "rb").read()
        if want is not None and got == want:
            same += 1
        else:
            differ += 1
            if differ <= 5:
                print("    upload differs: %s %s %dx%d (%d bytes, 4.7.0's %s)" % (base, PF_NAMES.get(pf, pf), w, h,
                      len(got), "none" if want is None else len(want)))
    return same, differ, missing, formats


def dump():
    ok = True
    for compression in (1, 0):
        for level in ("CAM_1A", "CAM_3A"):
            tag = "%s_tc%d" % (level, compression)
            dd = os.path.join(WORK, "dump", tag, "levels")
            if os.path.exists(dd):
                shutil.rmtree(dd)
            os.makedirs(dd)
            env = dict(os.environ)
            env["WZ_TEST_KTX2_DUMP"] = dd
            config = "terrainMode=2;terrainShadingQuality=1024;textureCompression=%d" % compression
            rc, lines, _, _, _ = run_port("dump", tag, script_lines=["waitlevel " + level, "quiet 120", "quit"],
                                          config=config, high=True, extra=["--game=" + level], timeout=600, env=env)
            asserts = [ln for ln in lines if "assert" in ln.lower()]
            refdir = os.path.join(WORK, "dump", tag, "ref")
            if os.path.exists(refdir):
                shutil.rmtree(refdir)
            os.makedirs(refdir)
            by_file = {}
            for name in sorted(os.listdir(dd)):
                if ".U" in name.rsplit("#", 1)[-1]:
                    continue                            # an upload dump (compare_uploads)
                base, fmt, lev = name.rsplit(".", 2)
                by_file.setdefault((base, int(fmt)), []).append(int(lev[1:]))
            same = differ = missing = 0
            formats = {}
            for (base, fmt), levs in sorted(by_file.items()):
                src = ktx2_source(base)
                if src is None:
                    missing += len(levs)
                    continue
                subprocess.run([KTX2REF, str(fmt), refdir, src], capture_output=True)
                ref = open(os.path.join(refdir, os.path.basename(src) + "." + str(fmt)), "rb").read()
                ranges = ktx2_levels(src, fmt)
                formats[fmt] = formats.get(fmt, 0) + len(levs)
                for lev in levs:
                    got = open(os.path.join(dd, "%s.%d.L%d" % (base, fmt, lev)), "rb").read()
                    at, n = ranges[lev]
                    if got == ref[at:at + n]:
                        same += 1
                    else:
                        differ += 1
                        if differ <= 5:
                            print("    differs: %s format %d level %d (%d bytes, ref %d)" % (base, fmt, lev, len(got), n))
            print("dump %s: %d files, %d levels identical to ktx2ref, %d differ, %d without a source; formats %s"
                  % (tag, len(by_file), same, differ, missing, formats))
            usame, udiffer, umissing, uformats = compare_uploads(dd, refdir)
            print("dump %s uploads: %d identical to 4.7.0's conversion of ktx2ref's output, %d differ, %d without a "
                  "source; formats %s; %d assert lines" % (tag, usame, udiffer, umissing, uformats, len(asserts)))
            ok = ok and differ == 0 and same > 0 and udiffer == 0 and usame > 0 and not asserts
    return ok


def perf():
    results = []
    for mode, size in ((1, 512), (2, 512), (2, 1024)):
        tag = "mode%d_%d" % (mode, size)
        config = "terrainMode=%d;terrainShadingQuality=%d" % (mode, size)
        rc, lines, shots, peak, seen = run_port("perf", tag, script_lines=["waitlevel CAM_1A", "log level_running", "quit"],
                                                config=config, high=True, extra=["--game=CAM_1A"], timeout=600,
                                                watch=("loading: wrf/cam1/cam1a.gam", "level_running"))
        t_load = seen.get("loading: wrf/cam1/cam1a.gam")
        t_run = seen.get("level_running")
        span = (t_run - t_load) if t_load is not None and t_run is not None else None
        results.append((tag, span, peak))
        print("perf %s: .gam load to level running %s s, peak private %.0f MB"
              % (tag, "%.1f" % span if span is not None else "?", peak / 1048576.0))
    return True


TESTS = {"regress": regress, "failpath": failpath, "options": options, "switch": switch, "tilesets": tilesets,
         "pngtilesets": pngtilesets, "pngswitch": pngswitch, "normaltilesets": normaltilesets, "dump": dump, "perf": perf}
if not args:
    print(__doc__)
    sys.exit(2)
names = ["regress", "options", "switch", "tilesets", "dump", "perf"] if args[0] == "all" else args
results = {n: TESTS[n]() for n in names}
print("results: " + ", ".join("%s %s" % (n, "ok" if r else "FAILED") for n, r in results.items()))
sys.exit(0 if all(results.values()) else 1)
