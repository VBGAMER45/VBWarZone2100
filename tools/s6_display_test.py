"""
s6_display_test.py - cluster S fork S6's tests: lib/sdl's display, window, cursor, clipboard and system functions
(WzPlatform's main_win and cursors_win) and src/urlhelpers.cpp, run on a Warzone.exe with UI scripts
(tools/ui_tests/sS6_*.txt, ui_test.py's commands) and checked in the log and the saved config.

  python -I tools\\s6_display_test.py [--exe PATH] [scenario ...]      (default: all)

Nothing here changes the user's desktop: every run has the TEST-ONLY WZ_TEST_NO_DESKTOP_EFFECTS set, so a display
mode change, a fullscreen window, the cursor clip, the clipboard and opening a link or a folder are logged by the
game instead of done (the clipboard is the game's own); the fullscreen modes are listed, never switched to. Each run
has a fresh config with portMapping=0 and lobbyserver=http://127.0.0.1:9 (ui_test.py's test-safe values).
Results in build/s6_display_test/<scenario>/ (log, config, screenshots).
"""
import glob
import os
import shutil
import subprocess
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORK = os.path.join(ROOT, "build", "s6_display_test")
SCRIPTS = os.path.join(ROOT, "tools", "ui_tests")
CHECKS = []


def check(ok, what):
    CHECKS.append((bool(ok), what))
    print("  [%s] %s" % ("ok" if ok else "FAIL", what))


def run(exe, name, script, config_extra="", args=(), limit=240, styles=None):
    work = os.path.join(WORK, name)
    if os.path.exists(work):
        shutil.rmtree(work)
    cfg = os.path.join(work, "config")
    os.makedirs(cfg)
    with open(os.path.join(cfg, "config"), "w", newline="\r\n") as f:
        f.write("[General]\nportMapping=0\nlobbyserver=http://127.0.0.1:9\n" + config_extra)
    env = dict(os.environ)
    env["WZ_TEST_NO_DESKTOP_EFFECTS"] = "1"         # TEST-ONLY: no desktop effects (see the docstring)
    t0 = time.time()
    p = subprocess.Popen([exe, "--configdir=" + cfg, "--vbscript=" + os.path.join(SCRIPTS, script), "--nosound",
                          "--datadir=" + os.path.join(ROOT, "data")] + list(args), cwd=os.path.dirname(exe), env=env)
    done = []
    if styles is not None:
        import threading
        probe = threading.Thread(target=window_styles, args=(p.pid, styles, lambda: bool(done)), daemon=True)
        probe.start()
    try:
        p.wait(timeout=limit)
        result = "exit code %d after %.0f s" % (p.returncode, time.time() - t0)
    except subprocess.TimeoutExpired:
        p.kill()
        p.wait()
        result = "killed at the time limit (%d s)" % limit
    done.append(True)
    print("%s: %s" % (name, result))
    log = ""
    for f in sorted(glob.glob(os.path.join(cfg, "logs", "WZlog-*.txt"))):
        log += open(f, encoding="utf-8", errors="replace").read()
    for shot in sorted(glob.glob(os.path.join(cfg, "screenshots", "*.png"))):
        base = os.path.basename(shot)
        out = os.path.join(work, base.split("-", 2)[2] if base.count("-") >= 2 else base)
        shutil.copyfile(shot, out)
        print("  screenshot: " + out)
    config = open(os.path.join(cfg, "config"), encoding="utf-8", errors="replace").read()
    return p.returncode if p.returncode is not None else -1, log, config


def lines(log, *keys):
    return [ln for ln in log.splitlines() if all(k in ln for k in keys)]


def count_with_repeats(log, key):
    """lines containing key, with the log's "last message repeated N times (total T repeats)" lines after them"""
    n, last, total = 0, False, 0
    for ln in log.splitlines():
        if ln.startswith("last message repeated"):
            if last:
                words = ln.replace("(", " ").replace(")", " ").split()
                total = int(words[words.index("total") + 1]) if "total" in words else int(words[3])
            continue
        if "last message repeated" in ln:
            continue                                # another message's repeats, logged in between
        if last:
            n += total
        last, total = key in ln, 0
        if last:
            n += 1
    if last:
        n += total
    return n


def errors(log):
    return [ln for ln in log.splitlines() if ln.startswith(("error", "fatal")) or "assert" in ln.lower()]


def config_value(config, key):
    for ln in config.splitlines():
        if ln.split("=", 1)[0].strip() == key:
            return ln.split("=", 1)[1].strip()
    return None


def main(argv):
    exe = os.path.join(ROOT, "bin", "Warzone.exe")
    if "--exe" in argv:
        exe = os.path.abspath(argv[argv.index("--exe") + 1])
    want = [a for a in argv if not a.startswith("--") and a != exe and os.path.abspath(a) != exe]
    os.makedirs(WORK, exist_ok=True)
    scenarios = SCENARIOS if not want else [s for s in SCENARIOS if s[0] in want]
    for name, fn in scenarios:
        print("== " + name)
        fn(exe)
    failed = [w for ok, w in CHECKS if not ok]
    print("%d of %d checks passed%s" % (len(CHECKS) - len(failed), len(CHECKS), (": FAILED " + "; ".join(failed)) if failed else ""))
    return 1 if failed else 0


WS_THICKFRAME = 0x40000
WS_POPUP = 0x80000000


def window_styles(pid, styles, stop):
    """polls the game window's style (GWL_STYLE) while the game runs: the set of styles seen"""
    import ctypes
    from ctypes import wintypes
    user32 = ctypes.windll.user32
    EnumProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

    def cb(hwnd, _):
        p = wintypes.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(p))
        buf = ctypes.create_unicode_buffer(64)
        user32.GetClassNameW(hwnd, buf, 64)
        if p.value == pid and buf.value == "WarzoneVB":
            styles.add(user32.GetWindowLongW(hwnd, -16) & 0xFFFFFFFF)
        return True
    while not stop():
        user32.EnumWindows(EnumProc(cb), 0)
        time.sleep(0.25)


def common(log, name):
    errs = errors(log)
    for ln in errs[:10]:
        print("  | " + ln[:200])
    check(not errs, "%s: no error or assert in the log" % name)


def scenario_links(exe):
    code, log, config = run(exe, "links", "sS6_links.txt")
    opened = lines(log, "TEST-ONLY: ShellExecuteEx not called")
    for ln in opened:
        print("  | " + ln[ln.index("[") if "[" in ln else 0:][:160])
    check(code == 0, "exit code 0")
    common(log, "links")
    for url in ("open https://wz2100.net/", "open https://wz2100.net/webchat/", "open http://donations.wz2100.net/"):
        check(any(ln.endswith(url) for ln in opened), "title screen link logged, not opened: " + url)
    check(any(" explore " in ln and ln.rstrip("/\\").lower().endswith(os.path.join("links", "config").lower()) for ln in opened),
          "Open Configuration Directory logged, not opened (explore <the config folder>)")


def scenario_window(exe):
    styles = set()
    code, log, config = run(exe, "window", "sS6_window.txt", config_extra="", args=["--debug=wz", "--debug=gui"], styles=styles)
    check(code == 0, "exit code 0")
    expected = "[makeWindowOptionsForm:218] Failed to set fullscreen mode: Fullscreen"   # 4.7.0's line when the change fails
    common("\n".join(ln for ln in log.splitlines() if expected not in ln), "window (but 4.7.0's own line for the refused fullscreen)")
    modes = lines(log, "wzMainScreenSetup_CreateVideoWindow:2745", "Monitor [")
    check(len(modes) > 0, "the display list: %d fullscreen modes kept (Monitor [id] WxH lines)" % len(modes))
    check(lines(log, "Changing fullscreen mode to [1] 640x480"), "Fullscreen Mode 640 x 480 @ 60: wzChangeFullscreenDisplayMode([1] 640x480)")
    check(config_value(config, "fullscreenWidth") == "640" and config_value(config, "fullscreenHeight") == "480"
          and config_value(config, "fullscreenScreen") == "1" and float(config_value(config, "fullscreenRefreshRate") or 0) == 60.0,
          "the fullscreen mode saved: %s x %s on %s @ %s" % (config_value(config, "fullscreenWidth"), config_value(config, "fullscreenHeight"),
                                                             config_value(config, "fullscreenScreen"), config_value(config, "fullscreenRefreshRate")))
    check(not lines(log, "TEST-ONLY: ChangeDisplaySettingsEx"), "no display mode change asked (the window is windowed)")
    check(config_value(config, "minimizeOnFocusLoss") == "0", "Minimize on Focus Loss Off saved (%s)" % config_value(config, "minimizeOnFocusLoss"))
    check(lines(log, "Changing window mode: Windowed -> Fullscreen"), "Window Mode Fullscreen asked")
    check(lines(log, "TEST-ONLY: not entering fullscreen: display 1 at 640x480 @60 Hz"), "fullscreen refused under the test variable (logged with its mode)")
    check(config_value(config, "fullscreen") == "0", "the window mode saved stays windowed (%s)" % config_value(config, "fullscreen"))
    check(config_value(config, "vsync") == "0", "Vertical Sync Off saved (%s)" % config_value(config, "vsync"))
    check(config_value(config, "cursorScale") == "200", "Cursor Size 200%% saved (%s)" % config_value(config, "cursorScale"))
    scaled = count_with_repeats(log, "Scaling cursor image from (")
    check(scaled == 28, "the 28 cursors made again at 200%%: %d images scaled (32 x 32 to 64 x 64, 33 x 33 to 66 x 66)" % scaled)
    check(styles and all(st & WS_THICKFRAME and not st & WS_POPUP for st in styles),
          "the window stayed a resizable, framed window (styles seen: %s)" % ", ".join("%08X" % st for st in sorted(styles)))


def scenario_scale(exe):
    code, log, config = run(exe, "scale", "sS6_scale.txt", args=["--debug=wz"])
    check(code == 0, "exit code 0")
    common(log, "scale")
    check(config_value(config, "displayScale") == "125", "Display Scale 125%% saved (%s)" % config_value(config, "displayScale"))
    check(config_value(config, "width") == "1024" and config_value(config, "height") == "768",
          "the window kept its size (%s x %s)" % (config_value(config, "width"), config_value(config, "height")))
    check(lines(log, "wzChangeDisplayScale:2072", "New mouse position"), "wzChangeDisplayScale ran (its mouse position lines)")


def scenario_altenter(exe):
    code, log, config = run(exe, "altenter", "sS6_altenter.txt")
    check(code == 0, "exit code 0")
    common(log, "altenter")
    check(lines(log, "Changing window mode: Windowed -> Fullscreen"), "Alt+Enter: wzAltEnterToggleFullscreen asked for fullscreen")
    check(lines(log, "TEST-ONLY: not entering fullscreen: display 1, Desktop Full"), "fullscreen (Desktop Full) refused under the test variable")
    check(config_value(config, "fullscreen") == "0", "the window mode saved stays windowed (%s)" % config_value(config, "fullscreen"))


def scenario_clipboard(exe):
    code, log, config = run(exe, "clipboard", "sS6_clipboard.txt")
    check(code == 0, "exit code 0")
    common(log, "clipboard")
    check(lines(log, "TEST-ONLY: clipboard text set (this process's, not the user's): Press the start hosting button to begin hosting a game."),
          "Copy Text to Clipboard: the chat line set in the game's own clipboard")
    print("  (the paste: look at build/s6_display_test/clipboard/pasted.png, the chat box)")


def scenario_grab(exe):
    code, log, config = run(exe, "grab", "sS6_grab.txt", config_extra="trapCursor=1\n", limit=300)
    check(code == 0, "exit code 0")
    common(log, "grab")
    grab = lines(log, "TEST-ONLY: SDL_SetWindowMouseGrab(true)")
    check(grab, "Trap Cursor: the game's start grabs the mouse (wzGrabMouse)")
    focused = any(ln.rstrip().endswith("(focus: -1)") for ln in grab)
    clipped = lines(log, "TEST-ONLY: ClipCursor not called: (")
    check(clipped if focused else not clipped, "the clip follows the focus (focused: %s, ClipCursor logged: %s)" % (focused, bool(clipped)))
    check(not clipped or lines(log, "TEST-ONLY: ClipCursor(NULL) not called"), "a clip made is released (logged) at the quit")


SCENARIOS = [("links", scenario_links), ("window", scenario_window), ("scale", scenario_scale),
             ("altenter", scenario_altenter), ("clipboard", scenario_clipboard), ("grab", scenario_grab)]

if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
