"""
map_script_test.py - script-generated maps (lib/wzmaplib/src/map_script.cpp: a map folder's game.js run with a seed;
WzMap's map_script_c over chakra.dll in QuickJS's place) against the C++ library with the QuickJS 4.7.0 pins:

  python -I tools\\map_script_test.py [--bin DIR] [--src DIR] [--seeds N,N,...] [--keep]

* the reference: build/map_script_test/ref/mapscriptref.exe (tools/mapscriptref.cpp with g++: lib/wzmaplib and its
  map_script.cpp from ref/warzone2100, QuickJS-wz a57c627 from ref/quickjs-wz with gcc, the defines of its
  CMakeLists.txt);
* the port: build/map_script_test/bin/MapScriptTests.exe (src/mapscripttests) over --bin's WzFramework.dll and
  WzMap.dll (default bin\\; a fork's private build, e.g. build\\j_S8\\bin), with a reg-free manifest;
* each script map of data/ (every map folder with a game.js) loaded with each seed (default 0, 1, 2, 42, 12345,
  2147483647, 2147483648, 4294967295) as a skirmish map, and the probes (build/map_script_test/maps: the API's
  checks and errors, the limited context, gameRand, generateFractalValueNoise's options, and the maps' own
  Math.sqrt / Math.pow calls logged): the dumps (map data, objects, CRCs, log lines) compared byte for byte, but for
  the text after the first line of an uncaught exception or a syntax error (QuickJS's and Chakra's stacks and
  syntax messages differ) and for the probe lines marked "(engine)" (what differs by design: BigInt).
* the logged Math.pow / Math.sqrt lines are compared by value (Chakra's Number-to-string may give another of the
  17-digit strings of the same double: counted and reported); their results are also checked against ucrtbase.dll's (the C runtime of the official
  MSVC build, through ctypes): what the official build computes for these maps.
--src: the folder holding mapscripttests\\ (default src\\, or a fork's stage). Results in build/map_script_test.
"""
import ctypes
import glob
import os
import re
import shutil
import subprocess
import sys

ROOT = r"C:\Users\Owner\desktop\websites\warzone"
sys.path.insert(0, os.path.join(ROOT, "tools"))
import build  # noqa: E402
import typelib  # noqa: E402

WORK = os.path.join(ROOT, "build", "map_script_test")
REF = os.path.join(ROOT, "ref", "warzone2100")
QJS = os.path.join(ROOT, "ref", "quickjs-wz")
WZMAPLIB = os.path.join(REF, "lib", "wzmaplib")
WZMAPLIB_SOURCES = ["map.cpp", "map_io.cpp", "map_terrain_types.cpp", "map_crc.cpp", "map_internal.cpp",
                    "map_debug.cpp", "map_jsonhelpers.cpp", "map_version.cpp", "map_script.cpp"]
QJS_SOURCES = ["cutils.c", "dtoa.c", "libregexp.c", "libunicode.c", "quickjs.c"]
MAP_CONFIG_H = """#pragma once
#define WZMAPLIB_VERSION_STRING "1.4.6"
#define WZMAPLIB_VERSION_MAJOR 1
#define WZMAPLIB_VERSION_MINOR 4
#define WZMAPLIB_VERSION_REV 6
"""
GXX = shutil.which("g++") or r"C:\Strawberry\c\bin\g++.exe"
GCC = os.path.join(os.path.dirname(GXX), "gcc.exe")
SEEDS = [0, 1, 2, 42, 12345, 2147483647, 2147483648, 4294967295]

# Math.sqrt and Math.pow logged with their arguments and results (the maps' floating point the engines compute)
MATH_LOGGER = """(function () {
  var M = Math, sqrt = M.sqrt, pow = M.pow;
  M.sqrt = function (a) { var r = sqrt(a); log('math sqrt ' + a + ' ' + r); return r; };
  M.pow = function (a, b) { var r = pow(a, b); log('math pow ' + a + ' ' + b + ' ' + r); return r; };
})();
"""

PROBES = {
    # the API's checks (every result logged; the errors caught)
    "probe_api": r"""
function t(name, f) { try { log(name + ': ' + String(f())); } catch (e) { log(name + ' threw ' + String(e)); } }
t('rand', function () { return [gameRand(), gameRand(7), gameRand(0), gameRand(-1), gameRand(4294967301), gameRand(2.9), gameRand('5'), gameRand(null), gameRand(undefined)].join(','); });
t('rand2', function () { return gameRand(1, 2) + ',' + gameRand(true); });
t('lognone', function () { return log(); });
t('logtwo', function () { return log('a', 'b'); });
t('logobj', function () { return log({toString: function () { return 'obj!'; }}); });
t('args6', function () { return setMapData(2, 2, [0, 0, 0, 0], [0, 0, 0, 0], [], []); });
t('widthstr', function () { return setMapData('2', 2, [0, 0, 0, 0], [0, 0, 0, 0], [], [], []); });
t('texnotarr', function () { return setMapData(2, 2, 5, [0, 0, 0, 0], [], [], []); });
t('small', function () { return setMapData(1, 2, [0, 0], [0, 0], [], [], []); });
t('wide', function () { return setMapData(300, 2, [], [], [], [], []); });
t('neg', function () { return setMapData(-2, 2, [], [], [], [], []); });
t('texlen', function () { return setMapData(2, 2, [0, 0, 0], [0, 0, 0, 0], [], [], []); });
t('texbig', function () { return setMapData(2, 2, [0, 70000, 0, 0], [0, 0, 0, 0], [], [], []); });
t('structpos', function () { return setMapData(2, 2, [0, 0, 0, 0], [0, 0, 0, 0], [{name: 'A', position: [1, 2, 3]}], [], []); });
t('structdir', function () { return setMapData(2, 2, [0, 0, 0, 0], [0, 0, 0, 0], [{name: 'A', position: [1, 2], direction: 70000}], [], []); });
t('structmod', function () { return setMapData(2, 2, [0, 0, 0, 0], [0, 0, 0, 0], [{name: 'A', position: [1, 2], direction: 1, modules: 255}], [], []); });
t('droiddir', function () { return setMapData(2, 2, [0, 0, 0, 0], [0, 0, 0, 0], [], [{name: 'D', position: [1, 2], direction: -1}], []); });
t('featdir', function () { return setMapData(2, 2, [0, 0, 0, 0], [0, 0, 0, 0], [], [], [{name: 'F', position: [1, 2], direction: 65536}]); });
t('newfn', function () { return new gameRand(); });
t('desc', function () {
  var d = Object.getOwnPropertyDescriptor(this, 'gameRand'), x = Object.getOwnPropertyDescriptor(this, 'XFLIP');
  return [typeof gameRand, d.writable, d.enumerable, d.configurable, x.value, x.writable, x.enumerable, x.configurable,
          XFLIP, YFLIP, ROTMASK, ROTSHIFT, TRIFLIP, preview].join(',');
});
preview = true; XFLIP = 1;
log('readonly ' + preview + ' ' + XFLIP);
var w = 4, h = 3, tex = [], ht = [];
for (var i = 0; i < w * h; ++i) { tex.push(i * 300 + (i % 3 == 0 ? 0.75 : 0)); ht.push(i == 5 ? 600 : (i == 6 ? -1 : i * 40)); }
tex[7] = '12'; ht[8] = {valueOf: function () { return 77; }}; ht[9] = undefined; tex[10] = true;
log('set ' + setMapData(w, h, tex, ht,
  [{name: 'A0CommandCentre', position: [128.9, 256], direction: 0x4000, modules: 0, player: 1},
   {name: 'Factory', position: [-1, 4294967296 + 5], direction: 0, modules: 2, player: -1},
   {position: [3, 4], direction: 10, modules: 1, player: 200},
   {name: 'X', direction: 5, modules: 254, player: 10},
   5],
  [{name: 'ConstructionDroid', position: [64, 64], direction: 65535, player: 3},
   {name: 'D2', position: ['7', null], direction: 0, player: -129}],
  [{name: 'Tree1', position: [1, 1], direction: 12345, player: 4},
   {name: 'Boulder', direction: 0}]));
""",
    # what the limited context has (QuickJS's base objects and Map / Set)
    "probe_context": r"""
var names = ['Object', 'Function', 'Array', 'Number', 'Boolean', 'String', 'Symbol', 'Math', 'Reflect', 'Error', 'TypeError',
  'RangeError', 'ReferenceError', 'SyntaxError', 'URIError', 'EvalError', 'Map', 'Set', 'WeakMap', 'WeakSet', 'parseInt',
  'parseFloat', 'isNaN', 'isFinite', 'escape', 'unescape', 'encodeURIComponent', 'decodeURI', 'eval', 'globalThis',
  'Date', 'RegExp', 'JSON', 'Proxy', 'Promise', 'ArrayBuffer', 'DataView', 'Uint8Array', 'Float64Array', 'Atomics',
  'SharedArrayBuffer', 'WeakRef', 'FinalizationRegistry', 'Intl', 'WebAssembly'];
log('types ' + names.map(function (n) { return n + ':' + typeof this[n]; }, this).join(' '));
log('(engine) BigInt:' + typeof this.BigInt);
log('normalize ' + typeof String.prototype.normalize);
var threw = false; try { eval('1 + 1'); } catch (e) { threw = true; } log('eval threw ' + threw);
threw = false; try { new Function('return 1')(); } catch (e) { threw = true; } log('Function threw ' + threw);
var m = new Map(); m.set('a', 1); var s = new Set([3, 1, 3]); log('mapset ' + m.get('a') + ' ' + s.size);
log('es6 ' + [1, 2, 3].map(x => x * 2).join(',') + ' ' + `${1 + 1}` + ' ' + (2 ** 10) + ' ' + [..."ab"].length);
gameRand(5);   // (a native call: VB's runtime reloads its x87 control word; the port sets 53 bits again)
log('sqrt ' + [10809, 5991.25, 2702.25, 441.14034858729053].map(function (v) { return Math.sqrt(v) === [103.96634070698074, 77.40316530995356, 51.98317035349037, 21.003341367203706][[10809, 5991.25, 2702.25, 441.14034858729053].indexOf(v)]; }).join(','));
setMapData(2, 2, [0, 0, 0, 0], [0, 0, 0, 0], [], [], []);
""",
    # generateFractalValueNoise's options and checks
    "probe_noise": r"""
function t(name, f) { try { var r = f(); log(name + ': ' + (Array.isArray(r) ? r.length + ' [' + r.join(',') + ']' : String(r))); } catch (e) { log(name + ' threw ' + String(e)); } }
t('basic', function () { return generateFractalValueNoise(16, 12, 100, 5, 8); });
t('norm', function () { return generateFractalValueNoise(16, 12, 100, 5, 8, 255); });
t('rows', function () { return generateFractalValueNoise(10, 7, 1000, 7, 4, 0, [], true); });
t('rigged', function () { return generateFractalValueNoise(20, 20, 64, 6, 8, 100,
  [[0, 0, 8, 8, function () { return 0; }], [10, 10, 30, 30, function (x, y, l, r) { return (x + y + l) % r; }],
   [4, 4, 12, 12, function (x, y, l, r) { return r - 1; }]], false); });
t('bigrange', function () { return generateFractalValueNoise(9, 9, 4294967295, 10, 16); });
t('crisp', function () { return generateFractalValueNoise(9, 9, 3000000000, 15, 32, 1000); });
t('hugescale', function () { return generateFractalValueNoise(5, 5, 50, 9, 200000); });   // layerScale^2 wraps (never to 0: C++ divides by it)
t('normbig', function () { return generateFractalValueNoise(4, 4, 40, 5, 4, 4294967295); });
t('norm2_64', function () { return generateFractalValueNoise(4, 4, 40, 5, 4, 18446744073709551616); });
t('normneg', function () { return generateFractalValueNoise(4, 4, 40, 5, 4, -1); });
t('norm2_32', function () { return generateFractalValueNoise(4, 4, 40, 5, 4, 4294967296); });
t('w0', function () { return generateFractalValueNoise(0, 4, 40, 5, 4); });
t('w300', function () { return generateFractalValueNoise(300, 4, 40, 5, 4); });
t('area', function () { return generateFractalValueNoise(256, 256, 40, 5, 4); });
t('s', function () { return generateFractalValueNoise(4, 4, 40, '5', 4); });
t('args4', function () { return generateFractalValueNoise(4, 4, 40, 5); });
t('args9', function () { return generateFractalValueNoise(4, 4, 40, 5, 4, 0, [], false, 1); });
t('regionsnotarr', function () { return generateFractalValueNoise(4, 4, 40, 5, 4, 0, 5); });
t('region4', function () { return generateFractalValueNoise(4, 4, 40, 5, 4, 0, [[0, 0, 1, 1]]); });
t('regionfn', function () { return generateFractalValueNoise(4, 4, 40, 5, 4, 0, [[0, 0, 1, 1, 5]]); });
t('cbrange', function () { return generateFractalValueNoise(8, 8, 40, 5, 4, 0, [[0, 0, 8, 8, function (x, y, l, r) { return r; }]]); });
t('cbthrow', function () { return generateFractalValueNoise(8, 8, 40, 5, 4, 0, [[0, 0, 8, 8, function () { throw new Error('cb'); }]]); });
t('cbstr', function () { return generateFractalValueNoise(8, 8, 40, 5, 4, 0, [[0, 0, 8, 8, function () { return '1'; }]]); });
t('rowsnotbool', function () { return generateFractalValueNoise(4, 4, 40, 5, 4, 0, [], 1); });
t('after', function () { return gameRand(); });
setMapData(2, 2, [0, 0, 0, 0], [0, 0, 0, 0], [], [], []);
""",
    "probe_syntax": "var a = ;\n",
    "probe_throw": "log('before');\nthrow new TypeError('map failed');\n",
    "probe_nomap": "log('no setMapData');\n",
}


def run_cmd(cmd, **kw):
    r = subprocess.run(cmd, capture_output=True, text=True, **kw)
    if r.returncode != 0:
        print("FAILED (%d): %s\n%s%s" % (r.returncode, " ".join(cmd[:3]), r.stdout[-3000:], r.stderr[-3000:]))
    return r.returncode == 0


def build_ref():
    """build/map_script_test/ref/mapscriptref.exe (rebuilt when a source is newer)."""
    out = os.path.join(WORK, "ref")
    exe = os.path.join(out, "mapscriptref.exe")
    sources = ([os.path.join(WZMAPLIB, "src", n) for n in WZMAPLIB_SOURCES] + [os.path.join(QJS, "quickjs", n) for n in QJS_SOURCES]
               + [os.path.join(ROOT, "tools", "mapscriptref.cpp")])
    if os.path.exists(exe) and os.path.getmtime(exe) >= max(os.path.getmtime(p) for p in sources):
        return exe
    print("building mapscriptref.exe (g++ / gcc: lib/wzmaplib with map_script.cpp, QuickJS-wz a57c627) ...")
    os.makedirs(os.path.join(out, "wzmaplib_internal"), exist_ok=True)
    with open(os.path.join(out, "wzmaplib_internal", "map_config_internal.h"), "w", newline="\n") as f:
        f.write(MAP_CONFIG_H)
    qinc = ["-I", os.path.join(QJS, "quickjs"), "-I", os.path.join(QJS, "quickjs-wz-extensions")]
    objs = []
    version = open(os.path.join(QJS, "quickjs", "VERSION.txt")).read().strip()
    for n in QJS_SOURCES:
        o = os.path.join(out, n[:-2] + ".o")
        # quickjs-wz's CMakeLists.txt: _GNU_SOURCE, QJS_DISABLE_ATOMICS, WIN32_LEAN_AND_MEAN, _WIN32_WINNT, CONFIG_VERSION
        if not run_cmd([GCC, "-c", "-O2", "-w", "-D_GNU_SOURCE", "-DQJS_DISABLE_ATOMICS", "-DWIN32_LEAN_AND_MEAN", "-D_WIN32_WINNT=0x0601",
                    '-DCONFIG_VERSION="%s"' % version, "-DQUICKJS_HAVE_SYS_TIME_H"] + qinc + ["-o", o, os.path.join(QJS, "quickjs", n)]):
            return None
        objs.append(o)
    ok = run_cmd([GXX, "-std=c++17", "-O1", "-static", "-w", "-I", os.path.join(REF, "3rdparty", "json", "include"),
              "-I", os.path.join(REF, "3rdparty", "optional-lite", "include"), "-I", out, "-I", os.path.join(WZMAPLIB, "include")]
             + qinc + ["-o", exe, os.path.join(ROOT, "tools", "mapscriptref.cpp")]
             + [os.path.join(WZMAPLIB, "src", n) for n in WZMAPLIB_SOURCES] + objs + ["-lpthread"])
    return exe if ok else None


def build_vb(dll_bin, src_root):
    """build/map_script_test/bin/MapScriptTests.exe over dll_bin's WzFramework.dll and WzMap.dll (copied beside it)."""
    fbin = os.path.join(WORK, "bin")
    os.makedirs(fbin, exist_ok=True)
    for n in ("WzFramework.dll", "WzMap.dll"):
        shutil.copy(os.path.join(dll_bin, n), fbin)
    folder = os.path.join(src_root, "mapscripttests")
    if not os.path.isdir(folder):
        folder = os.path.join(ROOT, "src", "mapscripttests")
    files = [os.path.join(folder, n) for n in build.sources(folder)]
    files += [os.path.join(build.SHARED, n) for n in build.sources(build.SHARED)]
    old = build.BIN, build.BUILD
    build.BIN, build.BUILD = fbin, WORK
    try:
        lines = ["Type=Exe", build.STDOLE]
        for ref in ("WzFramework", "WzMap"):
            libname, tlbid, (major, minor), _ = typelib.library(os.path.join(fbin, ref + ".dll"))
            lines.append("Reference=*\\G%s#%x.%x#0#%s#%s" % (tlbid, major, minor, os.path.join(fbin, ref + ".dll"), libname))
        for p in files:
            build.normalise(p)
            if p.lower().endswith(".cls"):
                lines.append("Class=%s; %s" % (build.vb_name(p), p))
            else:
                lines.append("Module=%s; %s" % (build.vb_name(p), p))
        res = build.write_manifest(["WzFramework", "WzMap"], "MapScriptTests")
        if res is None:
            return None
        lines.append('ResFile32="%s"' % res)
        lines.append(build.VBP_TAIL.format(name="MapScriptTests", ext="exe", path=fbin, startup="Sub Main", compat="",
                                           bounds="0", optimisation="2", threading=""))
        vbp = os.path.join(WORK, "MapScriptTests.vbp")
        with open(vbp, "w", newline="\r\n") as f:
            f.write("\n".join(lines))
        if not build.make("MapScriptTests", WORK, vbp):
            return None
    finally:
        build.BIN, build.BUILD = old
    return os.path.join(fbin, "MapScriptTests.exe")


def script_maps():
    """(name, folder, max players) of every map folder in data/ with a game.js."""
    found = []
    for js in sorted(glob.glob(os.path.join(ROOT, "data", "**", "game.js"), recursive=True)):
        folder = os.path.dirname(js)
        name = os.path.basename(folder)
        m = re.match(r"(\d+)c-", name)
        found.append((name, folder, int(m.group(1)) if m else 8))
    return found


def write_probes(maps):
    """build/map_script_test/maps: the probes, and the script maps with their Math calls logged."""
    out = os.path.join(WORK, "maps")
    if os.path.exists(out):
        shutil.rmtree(out)
    probes = []
    for name, text in PROBES.items():
        os.makedirs(os.path.join(out, name))
        with open(os.path.join(out, name, "game.js"), "w", newline="\n") as f:
            f.write(text)
        probes.append((name, os.path.join(out, name), 2))
    for name, folder, players in maps:
        src = open(os.path.join(folder, "game.js"), encoding="utf-8").read()
        if "Math.sqrt" not in src and "Math.pow" not in src:
            continue
        d = os.path.join(out, "math_" + name)
        os.makedirs(d)
        with open(os.path.join(d, "game.js"), "w", encoding="utf-8", newline="\n") as f:
            f.write(MATH_LOGGER + src)
        probes.append(("math_" + name, d, players))
    return probes


NOISY = re.compile(r"^log \d+ runMapScript:(843|889) ")
MATH_LINE = re.compile(r'^(log \d+ runMap_log:\d+ game\.js: "math (?:sqrt|pow) )(.*)"$')


def normalise(text):
    """The output without what the engines say differently: an uncaught exception's or a syntax error's text past
    its first line (the stacks), a syntax error's message, and the "(engine)" probe lines."""
    out, skipping = [], False
    for line in text.split("\n"):
        if skipping:
            if line.startswith(("log ", "loadFromPath failed", "scriptgenerated ")):
                skipping = False
            else:
                continue
        if NOISY.match(line):
            if line.split(" ", 3)[2].endswith(":843"):
                line = line.split(": ")[0] + ": <syntax error>"
            skipping = True
        if "(engine)" in line:
            continue
        m = MATH_LINE.match(line)
        if m:
            # the numbers by value: Chakra's Number-to-string can give another of the 17-digit strings of a double
            line = m.group(1) + " ".join(repr(float(v)) for v in m.group(2).split(" ")) + '"'
        out.append(line)
    return "\n".join(out)


def check_math(texts):
    """The logged Math.sqrt / Math.pow calls against ucrtbase.dll's sqrt / pow (the official MSVC build's C runtime):
    (calls, differing)."""
    ucrt = ctypes.CDLL("ucrtbase")
    ucrt.pow.restype = ucrt.sqrt.restype = ctypes.c_double
    ucrt.pow.argtypes = [ctypes.c_double, ctypes.c_double]
    ucrt.sqrt.argtypes = [ctypes.c_double]
    calls, differing = 0, []
    for text in texts:
        for line in text.split("\n"):
            m = re.match(r"^log \d+ runMap_log:\d+ game\.js: \"math (sqrt|pow) (.*)\"$", line)
            if not m:
                continue
            vals = [float(v) for v in m.group(2).split(" ")]
            got = vals[-1]
            want = ucrt.sqrt(vals[0]) if m.group(1) == "sqrt" else ucrt.pow(vals[0], vals[1])
            calls += 1
            if not (got == want or (got != got and want != want)):
                differing.append(line)
    return calls, differing


def main(argv):
    dll_bin = build.BIN
    src_root = os.path.join(ROOT, "src")
    seeds = SEEDS
    if "--bin" in argv:
        dll_bin = os.path.abspath(argv[argv.index("--bin") + 1])
    if "--src" in argv:
        src_root = os.path.abspath(argv[argv.index("--src") + 1])
    if "--seeds" in argv:
        seeds = [int(s) for s in argv[argv.index("--seeds") + 1].split(",")]
    os.makedirs(WORK, exist_ok=True)
    ref = build_ref()
    if ref is None:
        return 1
    vb = build_vb(dll_bin, src_root)
    if vb is None:
        return 1
    maps = script_maps()
    probes = write_probes(maps)
    print("script maps: %s" % ", ".join(n for n, _, _ in maps))
    cmds = []
    for name, folder, players in maps:
        for seed in seeds:
            cmds.append(("%s|%d|%d" % (folder, players, seed), "%s_%d.txt" % (name, seed)))
    for name, folder, players in probes:
        for seed in (seeds[:2] if name.startswith("probe") else seeds):
            cmds.append(("%s|%d|%d" % (folder, players, seed), "%s_%d.txt" % (name, seed)))
    for side in ("ref", "vb"):
        d = os.path.join(WORK, "out_" + side)
        if os.path.exists(d):
            shutil.rmtree(d)
        os.makedirs(d)
    cf = os.path.join(WORK, "commands.txt")
    with open(cf, "w", encoding="utf-8", newline="\n") as f:
        f.write("".join("%s\t%s\n" % c for c in cmds))
    import time
    t = time.time()
    if not run_cmd([ref, cf, os.path.join(WORK, "out_ref")]):
        return 1
    tref = time.time() - t
    t = time.time()
    r = subprocess.run([vb, cf, os.path.join(WORK, "out_vb")], timeout=3600)
    tvb = time.time() - t
    print("reference %.1f s, port %.1f s (exit code %d), %d runs" % (tref, tvb, r.returncode, len(cmds)))
    same = 0
    failed = []
    texts = []
    for arg, out in cmds:
        a = open(os.path.join(WORK, "out_ref", out), encoding="utf-8", errors="replace").read()
        p = os.path.join(WORK, "out_vb", out)
        b = open(p, encoding="utf-8", errors="replace").read() if os.path.exists(p) else "(no output)"
        texts += [a, b] if out.startswith("math_") else []
        if normalise(a) == normalise(b):
            same += 1
        else:
            failed.append(out)
        if a.startswith("loadFromPath failed") and not out.startswith("probe"):
            print("  %s: the reference failed to load it" % out)
    for name, folder, players in maps:
        a = open(os.path.join(WORK, "out_ref", "%s_%d.txt" % (name, seeds[0])), encoding="utf-8").read()
        crc = [ln for ln in a.split("\n") if ln.startswith(("mapdata", "crc", "structures", "droids", "features"))]
        print("  %s seed %d: %s" % (name, seeds[0], "; ".join(crc)))
    print("%d of %d dumps identical%s" % (same, len(cmds), (": differ " + ", ".join(failed)) if failed else ""))
    strings = 0
    for arg, out in cmds:
        if out.startswith("math_"):
            a = open(os.path.join(WORK, "out_ref", out), encoding="utf-8").read().split("\n")
            b = open(os.path.join(WORK, "out_vb", out), encoding="utf-8").read().split("\n")
            strings += sum(1 for x, y in zip(a, b) if x != y and MATH_LINE.match(x))
    print("number strings that differ for the same double (Chakra's Number-to-string): %d" % strings)
    fcf = os.path.join(WORK, "fpu_commands.txt")
    with open(fcf, "w", newline="\n") as f:
        f.write("fpu\tfpu.txt\n")
    subprocess.run([vb, fcf, WORK])
    print("the VB process's x87 control word: %s" % open(os.path.join(WORK, "fpu.txt")).read().strip())
    calls, mdiff = check_math(texts)
    print("Math.sqrt / Math.pow calls logged: %d (the reference's and the port's); different from ucrtbase.dll's: %d"
          % (calls, len(mdiff)))
    for line in mdiff[:10]:
        print("  " + line)
    return 0 if not failed and not mdiff and r.returncode == 0 else 1


def run(keep=False):
    """run_tests.py's entry (bin's DLLs, src/mapscripttests, the default seeds): True when every dump matches."""
    return main([]) == 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
