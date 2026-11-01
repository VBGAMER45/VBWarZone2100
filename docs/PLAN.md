# WarzoneVB: feasibility and plan

Plan and status of the Visual Basic 6 port of Warzone 2100 (Windows only),
done the way DoomVB and QuakeVB were: one VB module per C++ file,
translated statement by statement, compiled to native code, tested with
hidden scripted runs. Read this first; QuakeVB's `docs/CONVENTIONS.md`
(`C:\Users\Owner\Desktop\githubprojects\VBQuake`) has the C-to-VB rules
this port starts from.

## Source

* **Warzone 2100 4.7.0** (release tag, commit
  `7d3d8e51a61ceee89e6bf672a41337e1a9cd7392`, 2026-04-06), shallow clone in
  `ref/warzone2100`, with the data submodules `data/base/texpages` (all
  textures) and `data/fonts`. Code and data are GPL v2 or later; the port
  stays GPL.
* Not fetched: the music (`data/music`, Opus), `data/terrain_overrides`,
  the campaign mods, and the C++ libraries (`3rdparty/*`) that the port
  replaces. Fetch one with
  `git -C ref\warzone2100 submodule update --init --depth 1 <path>`.
* GitHub's "Download ZIP" (the first copy, in Downloads) leaves out all 31
  submodules, the textures among them; do not use it.
* The official 4.7.0 Windows build (wz2100.net) is the reference
  implementation for the replay tests (see Testing). It is not downloaded
  yet; it is needed in cluster B.

## Verdict

**Possible.** Every question that could have ruled it out was measured on
this PC (`tools/probe`, results below):

1. **VB6 name table.** A VB6 project holds about **32,100 distinct names**
   (all of them count: procedures, parameters, locals, constants, UDT
   fields). The code to port needs 29,505. That leaves too little room for
   the VB-side names, so the port is split into **seven projects**: six of
   Warzone's own `lib/` folders become DLLs, and the rest goes in the exe.
   VB6 handles that well: a call between projects costs 31 ns, and `bin/`
   runs without registering anything.
2. **JavaScript.** Game rules, campaign and skirmish AIs are 53,900 lines of
   ES2015 JavaScript. Windows' own Chakra engine (`chakra.dll`) runs them
   from VB6: **all 195 scripts parse**, and a call from JS into VB costs
   0.25 us. QuickJS is not ported.
3. **Renderer.** Warzone draws only with OpenGL 3 (or Vulkan) shaders. VB6
   creates an OpenGL 3.3 core context and calls GL through function
   pointers (37 ns per call); **Warzone's own shaders compile and link
   unchanged**. The GL backend is ported; the 96 GLSL files stay as data.

Size: **243,000 code lines of C++17** to port, about 6 times QuakeVB's
39,000 lines of C. Estimate: 250,000 to 320,000 lines of VB.

## Size compared to QuakeVB

| | WinQuake (QuakeVB) | Warzone 2100 4.7.0 |
|---|---|---|
| Source to port | 63,000 lines of C, 39,300 code lines | 664 files (300 `.cpp`), **242,756 code lines** in scope (345,000 lines in src and lib) |
| Language | C89 | C++17 (counts below) |
| Distinct names (case-folded) | about 5,600 (QuakeVB's VB) | 29,505 (the C++ in scope) |
| Game logic | QuakeC VM (`progs.dat`) | C++ simulation (integer, deterministic) plus 53,945 lines of JavaScript in 195 files: campaign 20,398, skirmish AIs 27,406, rules 2,784 |
| Renderer | software spans and edges | OpenGL 3.x / Vulkan, 96 GLSL files |
| Data | 50 MB of paks | 239 MB (1,868 PNG, 560 Ogg Vorbis sounds, 870 PIE models, JSON stats, maps) |
| Libraries | none | SDL3, PhysFS, OpenAL, Vorbis, Opus, Theora, libpng, FreeType, HarfBuzz, FriBiDi, nlohmann/json, fmt, glm, QuickJS, libsodium, SQLite, curl, libzip, zlib, ... |
| VB | about 37,000 lines | estimate 250,000 to 320,000 lines |

What the C++ uses (src and lib, without 3rdparty): smart pointers 2,823,
`std::vector` 1,674, `virtual` 1,514, lambdas 1,308, nlohmann JSON 856,
`int64_t`/`uint64_t` 466, `std::function` 360, `throw`/`catch` 348,
templates 289, threads and mutexes 113.

The simulation is integer arithmetic, for multiplayer sync: `move`,
`action`, `order`, `combat`, `projectile`, `droid`, `ai`, `astar`, `fpath`
and `visibility` contain no `float` at all. VB's `Long` with overflow checks
off behaves like `int32_t`, so a **bit-exact** simulation is a realistic
target.

## Project layout

| Project | Kind | Ports | Names (4.7.0) | Left of ~32,100 |
|---|---|---|---|---|
| `WzFramework.dll` | ActiveX DLL | `lib/framework` (debug log, strings, files, config, JSON, trig, CRC) and the window-free part of `wzapp.h` (ticks, dialogs, threads); new: PhysFS over folders, JSON DOM, zip/inflate | 2,870 | 29,230 |
| `WzMap.dll` | ActiveX DLL | `lib/wzmaplib` (map formats) | 1,510 | 30,590 |
| `WzSound.dll` | ActiveX DLL | `lib/sound`; new: waveOut mixer, stb_vorbis port | 1,156 | 30,944 |
| `WzGfx.dll` | ActiveX DLL | `lib/ivis_opengl` (gfx_api, GL backend, PIE models, textures, text) | 3,354 | 28,746 |
| `WzNet.dll` | ActiveX DLL | `lib/netplay` (connections, message encoding, queues, replays, syncDebug, sessions) | 2,407 | 29,693 |
| `WzPlatform.dll` | ActiveX DLL | `lib/sdl` on Win32: window, input, event loop, the GL backend (WGL) | 1,023 | 31,077 |
| `Warzone.exe` | Standard EXE | `src`, `lib/gamelib`, `lib/widget` | 22,264 | 9,836 |

`src/shared` holds modules compiled into every project, like a statically
linked C runtime: the C library replacements (`printf_vb`, `stdio_win`,
`ctime_win`) and small framework modules without state (`wztime_c`).

Why this split: `framework` and `wzmaplib` depend on nothing in Warzone,
and `sound` and `ivis_opengl` reach up into `src` only for a few small
things (the game clock, a console message, a droid's position for a
tracking sound, profiling). Those few become callbacks into the exe
(30 ns each). `gamelib` (the game clock) and `widget` (subclassed all over
the UI) are woven into the game, so they stay in the exe.

`lib/sdl` gets its own DLL, as it is its own library in Warzone's build: it
sits above the renderer (it implements the renderer's GL backend interface,
`backend_OpenGL_Impl`, and calls `pie_SetVideoBufferWidth` and the context's
`handleWindowSizeChange`) and below the game, which it reaches through
`IPlatformGame` (`mainLoop`, `mainShutdown`, `realTime`, the window
settings `war_*`, `gameScreenSizeDidChange`, `saveGfxConfig`). The few
window functions the framework itself calls (`wzIsFullscreen`,
`wzChangeWindowMode`, `inputInitialise`, the dialog owner) go through
`IFrameworkPlatform`, which WzPlatform registers at startup.

`netplay` is the most entangled of the DLLs (17 includes of `src` headers),
but the coupling is countable (`tools/net_coupling.py` scans both sides' headers,
leaving out `netlobby` and port mapping):

* **netplay calls the game** about 60 times: the game clock (`gameTime`,
  `realTime`, `setGameTime`, `checkPlayerGameTime`, `recvPlayerGameTime`),
  joins and leaves (`MultiPlayerJoin`, `MultiPlayerLeave`), lobby options
  (`changeColour`, `kickPlayer`, `getLockedOptions`, `resetReadyStatus`, ...),
  player names and stats (`getPlayerName`, `setPlayerName`,
  `recvMultiStats`, ...), quick chat, `addConsoleMessage`,
  `version_getVersionString`, `war_getMPcolour`. These become the methods
  of `INetGame`, an interface of WzNet that the exe implements.
* **The game calls netplay** through the `NET*` message API
  (`NETbeginEncode`, `NETuint32_t`, `NETend`, the queues), the `NetPlay`
  player table (63 uses in src: a SAFEARRAY alias in the exe), and
  `syncDebug`.
* Replay files are **not compressed** (a magic number, the settings as
  JSON, the embedded map, then the raw network messages), so cluster B
  reads official replays with WzFramework's JSON and plain binary reads.
* Live connections are zlib streams. Only inflate needs a real port (it is
  needed for `.wz` archives anyway): the sending side can write "stored"
  deflate blocks, which are valid zlib for any reader.
* netplay's three threads (replay saving, pending writes, opening a
  connection) become non-blocking sockets polled every frame, as in
  QuakeVB's `net_wins`.
* `netlobby` (the official lobby server client, 2,594 lines) is optional
  (cluster K); it would need more of the UI (`verifyprompt`).

If the exe still runs short of names, the script bridge (wzapi, qtscript,
Chakra backend) is the next candidate for a DLL.

Names a project uses from a DLL count in that project too (measured: see
below), so DLL interfaces stay narrow. `python -I tools\name_budget.py
ref\warzone2100` recounts the budget.

Folders:

| Folder | Holds |
|---|---|
| `src/framework`, `src/map`, `src/sound`, `src/gfx`, `src/net`, `src/platform`, `src/game` | the seven projects' sources (`.bas`, `.cls`, `.frm`); `build.py` writes their `.vbp` |
| `src/shared` | modules compiled into every project |
| `bin/` | built exe and DLLs, the runtime folder |
| `compat/` | binary compatibility copies of the DLLs (keep them: they fix the GUIDs) |
| `build/` | build logs, the generated manifest and `Warzone.res` |
| `ref/warzone2100` | Warzone 4.7.0 source and data |
| `tools/` | build and analysis scripts; `tools/probe` the feasibility probes |
| `docs/` | this plan |

## Feasibility measurements (tools/probe)

### 1. VB6 name table and module count (`names_gen.py`)

Synthetic projects. The names are mixed like ported code: public
functions (each with a parameter and a local), constants, UDT fields and
module variables. Each project is compiled with `VB6 /make` and run.

| Distinct names | Modules | Result |
|---|---|---|
| 20,100 to 32,100 | 100 | builds and runs |
| 32,400 / 33,400 / 35,100 / 40,100 | 100 | **fails with no message** (empty `/out` file) |
| 20,100 | 300 | `Could not execute LINK.EXE` from a 140-character folder; builds from the project folder |
| 20,400 | 600 | `LINK.EXE` fails from the project folder; builds from a short `subst` drive |
| 4,250 | 850 | builds (short drive) |
| 6,300 | 900 | fails with no message (short drive) |

Findings:

* **About 32,100 names per project**, locals and parameters included.
* **At most about 850 modules** per project.
* VB6 passes every `.OBJ` path to LINK.EXE on one command line, so a
  project's module count times its path length must stay under 32,767
  characters. `build.py` warns when it gets close; the cure is a shorter
  path (`subst`).
* VB6 can fail without a message: `build.py` treats an empty log as a
  failure and says why it usually happens.

### 2. Chakra JavaScript engine (`chakra.bas`)

| Test | Result |
|---|---|
| Runtime and context from VB6 (`SysWOW64\chakra.dll`, declared directly) | works |
| ES2015: class with getter, `let`/`const`, arrow functions, template literal, `includes`, `**` | correct |
| ES2018 object spread and rest | SyntaxError (legacy Chakra lacks it; Warzone does not use it) |
| VB function called from JS (`JsCreateFunction` with `AddressOf`) | correct; 0.25 us per call more than a JS function |
| `JsParseScript` on every 4.7.0 script (1.5 M characters) | **195 of 195 parse**, 62 ms |

### 3. OpenGL from VB6 (`glprobe.bas`)

| Test | Result |
|---|---|
| Hidden window, legacy context | NVIDIA GeForce RTX 3080, OpenGL 4.6.0 |
| `wglCreateContextAttribsARB` through a function pointer (DispCallFunc) | OpenGL 3.3 core, GLSL 3.30 |
| 4.7.0's `rect` and `tcmask` (model) shaders, loaded as `gfx_api_gl.cpp` does (`#version 330 core` plus the source, includes inlined) | **compile and link** |
| `glViewport` declared directly / through its pointer with DispCallFunc / with CallWindowProcW | 0.021 / **0.037** / 0.183 us per call (checked: the viewport really changes) |

### 4. Exe and ActiveX DLLs (`xproj.py`)

| Test | Result |
|---|---|
| Exe calls a GlobalMultiUse function of the DLL (unqualified, like its own) | works, **0.031 us** per call |
| A UDT array owned by the DLL, seen from the exe through a SAFEARRAY alias (QuakeVB's `SetView`) | works both ways; the same 16-byte layout on both sides |
| DLL calls back into the exe (an interface class of the DLL, `Implements` in the exe) | works, **0.030 us** per call |
| `Err.Raise` in the DLL | caught in the exe with number, source and description |
| Exe with 29,100 names of its own calling 4,000 DLL functions / not calling them / 27,100 calling them | fails / builds / builds: **the DLL names a project uses count in its name table** |
| Exe with an external manifest, DLL unregistered | works, but only if the manifest is there before the exe's first run (Windows caches its absence), hence the embedded manifest |

### 5. The build (`build.py`, on the six skeleton projects)

| Test | Result |
|---|---|
| Five DLLs and the exe, DLL to DLL references (`WzGfx`, `WzSound`, `WzNet` use `WzFramework`) | build and run |
| WzNet calls back into the exe through `INetGame` (the version string), DLLs unregistered | works |
| A method name with an underscore in an interface used with `Implements` | compile error "Bad interface for Implements: method has underscore in name" (naming rule below) |
| Embedded manifest of a size that is not a multiple of 4 | the exe does not start: the resource is padded with NULs, and Windows rejects them as invalid XML (event log, source SideBySide). `build.py` pads the manifest with spaces to a multiple of 8 |
| Binary compatibility: second build | same type library and class GUIDs |
| `bin/` with every DLL unregistered (embedded manifest) | runs |
| Incompatible change to a DLL's public interface | `/make` fails with VB6's message ("compatibility error ... 'GfxBuildInfo' in the 'GfxApi' class module"), no dialog |
| `--break-compat WzGfx` | new GUIDs; the exe's reference and manifest follow by themselves |

### 6. First run (cluster 0)

`bin\Warzone.exe --configdir=<dir> --frames=300` (all DLLs unregistered):

* The log goes to stderr and to `<dir>\logs\WZlog-MMDD_HHMMSS.txt`, each
  line as Warzone writes it, with the C++ function and line
  (`wz      |14:35:36: [initialize_ConfigDir:509] Write dir: ...`); the
  first line hides the user folder (`--- Starting log
  [[WIN_USER_FOLDER]\...]---`), as Warzone's does.
* A 1024 x 768 window titled "Warzone 2100" opens with an OpenGL 3.3 core
  context (8-bit RGB, 24-bit depth, 8-bit stencil, no sRGB; NVIDIA RTX
  3080), and each frame is cleared and swapped: a capture of the window
  (`tools\capture_window.ps1`) shows the stand-in's test colour (25, 51, 76).
* 300 frames take 5.4 s with startup: about 60 frames per second (vsync).
* The window's events (shown, focus, moved) reach `handleActiveEvent`
  through the event ring; `--frames` quits through `wzQuit`, the loop calls
  `mainShutdown`, and the process ends with code 0.
* A rebuild must keep the DLLs' GUIDs. VB gives new members of a grown
  interface fresh GUIDs on every build until the compat copy has them, and
  then WzPlatform's compat copy no longer matched WzFramework's enums ("The
  binary compatibility DLL or EXE contains a parameter type or return type
  whose definition can not be found"). `build.py` now refreshes a compat
  copy whenever its DLL's type library version changes.

### 7. The config file (cluster 0)

`configuration.cpp` reads and writes `<configdir>\config` through
`INIReaderWriter` (inih), as in 4.7.0:

* A first run in an empty folder logs "Could not read existing
  configuration file" (`loadConfigIniFile:241`) and writes the 100 keys of
  `[General]` with Warzone's defaults (2010 bytes: `width=1024`,
  `gfxbackend=opengl`, `lobbyserver=https://wzlobby.wz2100.net/lobby` decoded
  from netlobby.h's bytes, `configVersion=4`, ...).
* A second run reads it and writes it back byte for byte.
* Edited values go through the C++ checks: `fog=TRUE` is read as true,
  `mapZoom=2650` (not a multiple of 200) becomes 2600, `radarTerrainMode=9`
  becomes 3, `displayScale=50` becomes 100, `fxvol=abc` logs
  `[iniSectionGetInteger:120] ... invalid stoi argument` and keeps 100, and
  `fullscreenPixelDensity=1.5` is written back as 1 (C++'s quirk: its
  `iniGetFloat` lambda returns `optional<int>`). Bad mouse keys, leave
  modes, backends and terrain values log the warnings at lines 354, 371,
  648, 671, 707, 715 and 743. `configVersion=2` upgrades `trapCursor=true`
  to Enabled (1). Keys the game does not write (`hostConnectionProvider`)
  stay in the file.
* A window resized during a run (`tools\resize_window.ps1`, 900 x 700
  outside) is saved at shutdown as its client area, `width=884`,
  `height=661`.
* Deviation: `wzGetDefaultGfxBackendForCurrentSystem` gives OpenGL (Warzone
  on Windows: Vulkan where it works), the only backend ported.

### 8. Command line, data and language (cluster 0)

* `clparse.cpp` is ported whole, with its own popt replacement. `--help`
  prints the 62 options of the Windows build plus the two VB-only ones
  (`--frames`, `--nodialogs`) to stdout and exits 0; `--version` prints
  "Warzone 2100 - Version: 4.7.0, Built: ..."; an unknown option, a bad
  `--debug` flag or `--configdir` without "=" stops with Warzone's message on
  stderr and exit code 1 (qFatal). `--enableconsole` attaches a console.
  Options whose subsystems come later (`--game`, `--loadskirmish`,
  `--autohost`, the game log, the command interface) set their stand-ins'
  values.
* Checked: `--resolution=100x100` logs "Screen width < 640 unsupported" and
  "Screen height < 480" (`ParseCommandLine:898`, `:903`) and saves 640 x
  480; `--gameport=2222` logs line 1108 and is not saved (the port override);
  `--noshadows --nosound --portmapping=false --host-chat-config=qc` are
  saved as `shadows=0`, `sound=0`, `portMapping=0`, `hostingChatDefault=0`.
* The user folders are made as in 4.7.0 (`autohost`, `mods/4.7.0/global`,
  `savegames/campaign/auto`, `replay/multiplay`, ...).
* The data: `build.py` makes `data/` a mirror of `ref/warzone2100/data`
  (real folders, the files hard links; a junction until section 35),
  so `scanDataDirs` finds it as `<prefix>\data` (priority 3). The search
  path comes out as Warzone's: the write dir, then `data\mp`, `data\`,
  `data\base`; "gamedesc.lev found at ...\data\base". `--datadir` is
  registered at priority 1 and comes first. A missing `--mod=missing.wz`
  logs "The global mod ... doesn't exist!" and is dropped. `.wz` archives
  (zip) are not read yet: a mount of one fails with PhysFS's "unsupported".
* `i18n.cpp`: `initI18n` binds the locale folder (`<prefix>\./locale`, as the
  Windows build) and selects the system locale; `language=de` in the config
  logs "Setting language to "Deutsch" (de)" and "Requested locale "7""
  (`SetThreadLocale`) and is saved back; `language=xx` logs "Requested
  language "xx" not supported." Without gettext the text stays English.

### 9. JSON and WzConfig (cluster 0)

* nlohmann::json (Warzone's 3.11.3) is a class, `nlohmann_json`, in
  WzFramework: one object per JSON value, made by `json_parse`,
  `json_object`, `json_array`, `json_value` (`docs/CONVENTIONS.md`, "JSON").
  The lexer, parser and serializer reproduce the library's behaviour and
  messages; its Grisu2 float formatter (`dtoa_impl`) is ported line by line
  with Decimal for `uint64_t`.
* `tools\run_tests.py` builds `bin\WzTests.exe` (`src/tests`, VB-only) and
  `build\jsonref.exe` (g++ from Strawberry Perl, with the real header from
  `ref/warzone2100/3rdparty/json`), runs 2061 commands through both and
  compares the output byte for byte. All match:
  * the 383 `.json` files of the game data (7 MB), each parsed and dumped
    four ways: `dump(4)`, `ordered_json` `dump(4)`, `dump()`, ASCII;
  * 36 crafted inputs: parse errors with their messages, lines and
    columns ("[json.exception.parse_error.101] parse error at line 1,
    column 10: syntax error while parsing value - invalid literal; last
    read: '"a": tru}'"), BOMs, surrogates, duplicate keys, number limits;
  * 3085 doubles (edge values, 1500 random bit patterns, 1500 everyday
    numbers) through `dump`: the Grisu2 port gives the library's digits;
  * `WzConfig` on the 241 data files made of groups: every key read with
    `value(key).toWzString()`, and each file copied group by group into a
    writable `WzConfig`, whose saved file equals the library's `dump(4)`.
* Known difference: invalid UTF-8 bytes quoted in a parse error's "last
  read" text are U+FFFD in the port (a VB String holds no invalid UTF-8).
* Speed: WzTests takes about 24 s for all of it, about 2 MB/s of JSON
  parsed and dumped. Enough for loading the stats once; to be measured
  again with cluster A's loaders.

### 10. wzmaplib (cluster A, step 1)

* WzMap.dll has `lib/wzmaplib`'s core: `map.cpp` (game.map, the
  structures, droids and features in `.bjo` and JSON v1/v2, the lazy-loading
  `Map` class, the CRCs, `loadedMapFormat`, the writers), `map_io.cpp`
  (`IOProvider`, `BinaryIOStream`, `StdIOProvider` on the Windows file API
  with C stdio's EOF and buffering behaviour), `map_terrain_types.cpp`,
  `map_crc.cpp`, `map_jsonhelpers.cpp` (on WzFramework's `nlohmann_json`),
  `map_internal.cpp` (`printLog_`), `map_debug.cpp`, `map_version.cpp`.
  Missing: `map_package`, `map_levparse`, `map_stats`, `map_preview`;
  `runMapScript` (`map_script.cpp`, QuickJS) is a stand-in that fails,
  so the five game.js maps do not load yet.
* The library's classes are public classes of the DLL (`Map`,
  `MapData`, `Structure`, `Droid`, `Feature`, `TerrainTypeData`,
  `ObjectVector`), its abstract classes interfaces (`LoggingProtocol`,
  `IOProvider`, `BinaryIOStream`): the game will implement them on
  PhysFS, as `WzMapPhysFSIO` does.
* `tools\run_tests.py` builds `build\mapref.exe` (`tools/mapref.cpp` with
  the real wzmaplib sources) and compares, for the 68 map folders of the
  game data (24 campaign, 44 skirmish), a dump of everything loaded (the
  tiles, gateways, objects, terrain types, CRCs, loaded format and every
  log line) twice with different map types and player counts, and an
  export in each of the three output formats (every file written, and
  the log). All 1285 outputs (136 dumps, 1149 files and logs of the
  exports) match on the first run, with the C++ bugs
  kept (`leftPadStrToMinimumLength` pads on the right, so v1 JSON keys
  repeat and objects are lost on export, as in Warzone).
* Speed: the map tests take about 45 s of WzTests' 69 s (340 loads, 189
  exports written).

### 11. Resources and levels (cluster A, step 2)

* WzFramework has the resource system: `frameresource.cpp`
  (`frameresource_c`) and the `.wrf` grammar, its flex lexer
  (`resource_lexer_c`) and bison parser (`resource_parser_c`) written out
  by hand with the generators' rules (longest match, start conditions,
  `yylineno`, bison's "syntax error" and `YYABORT`). The load and release
  functions are the game's (`data.cpp`): the framework calls them by
  number through `IFrameworkGame`, which the exe implements
  (`FrameworkGame.cls`), as WzPlatform calls the game through
  `IPlatformGame`.
* The exe has `levels.cpp` (`levels_c`, the level list and `levLoadData`),
  its lexer (`level_lexer_c`), `data.cpp` (`data_c`, all the resource
  types), and from `init.cpp` `loadLevFile`, the map list (`buildMapList`,
  `CheckForMod`, `CheckForRandom`), `systemInitialise` and
  `systemShutdown`; `realmain` now runs `systemInitialise` after
  `frameInitialise`, and `mainShutdown` ends with `systemShutdown`.
  `wz::crc_update` is in `src/shared/crc_c` (the multiplayer data hash),
  `trimTechLevelFromMapName` in WzMap (`map_package_c`).
* What the loaders and `levLoadData` call (stats, images, sounds, string
  resources, scripts, the init stages, missions, save games) are
  stand-ins: a level's data cannot load yet. Map packages (`maps/*.wz`)
  wait for wzmaplib's `map_package` and zip support.
* Checked: flex and bison are not installed, so there is no C++
  reference for the two grammars. A startup with `--debug=wz` loads
  `gamedesc.lev` and `addon.lev` without errors; the 114 datasets the port
  builds (names, types, players, game index, base and camchange links,
  data files) equal an independent Python reading of the same files; a
  test `.wrf` with both kinds of comment, an empty directory and a
  case-different duplicate loads as the C++ does ("Duplicate file name:
  Body.json (hash 9c2b10e)", the PJW hash recomputed in Python), and a bad
  one gives bison's "syntax error at line 3" and the fatal "Failed to
  parse". The JSON, WzConfig and map tests still match (3346).
* Found on the way: `build.py` wrote the manifest's typelib versions in
  hex, so WzFramework 1.11 became "1.b" and Windows refused to start the
  exe; the manifest now takes decimal.

### 12. Stats (cluster A, step 3)

* The exe has the stats: `stats.cpp` (`stats_c`: the eight component
  types, propulsion types, terrain table, weapon modifiers, propulsion
  sounds, the lookups by name, the upgradeable values), `statsdef.h`
  (`statsdef_h`), the structure stats and strength modifiers of
  `structure.cpp` (`structure_c`, `structuredef_h`), `feature.cpp`'s
  (`feature_c`, `featuredef_h`), `research.cpp`'s loader with its cycle and
  category checks (`research_c`, `researchdef_h`), the template store and
  `templates.json` (`template_c`, `droiddef_h`), the players' component and
  structure lists (`mechanics_c`), and what the templates need from
  `design.cpp` (default names), `droid.cpp`, `objmem.cpp` (object ids),
  `hci.cpp` (the UI's template list) and `text.cpp`. `intfac.h`'s image
  enum is `intfac_h`. data.cpp's loaders now load for real.
* WzFramework has the string resources: `strres.cpp`, its grammar
  (`strres_lexer_c`, `strres_parser_c`) and `treap.cpp`; WzSound has
  `audio_id.cpp` (the sound effects' table, `INGAME_AUDIO` in SoundApi).
* The layout (`docs/CONVENTIONS.md`, "Stats"): one flat UDT per stats type,
  arrays with a count (`asWeaponStats`, `asWeaponStats_size`); a typed stat
  pointer is an index (-1 = nullptr); a `BASE_STATS *` or
  `COMPONENT_STATS *` is a handle (the stat's `StatType` plus its index,
  which equals its `ref` for all but templates), read through
  `BASE_STATS_<field>` and `COMPONENT_STATS_<field>`.
* Stand-ins: `getViewData` gives nullptr (the intelligence screen's view
  data). (Until A4, `modelGet` was one too; see section 13.)
  Not ported yet: whatever needs the objects (`objActiveRadar`, the
  research results and upgrades, `designableTemplate`), the player's stored
  designs (`initTemplates`, `storeTemplates`), the favourites file.
* Checked: `tools/run_tests.py` builds `build\statsref.exe` from the real
  4.7.0 code: `wzconfig.cpp`, `wzstring.cpp`, `debug.cpp`, `strres.cpp`,
  `treap.cpp`, `audio_id.cpp` and the rest whole, and the loaders' part of
  `stats.cpp`, `structure.cpp`, `research.cpp`, `template.cpp` and five
  others by line ranges (with `#line`, so the logs carry the real file names
  and lines), over a PhysFS layer on the same search path. `Warzone.exe
  --vbtest=<file>` (VB-only) runs the same commands before its window opens.
  Both load the base and the mp stats as `stats.wrf` orders them, plus
  seven sets of crafted files for the error paths (bad classes, missing
  ZNULL entries, wrapped integers, strings for numbers, missing models and
  sounds, research cycles and category graphs, templates with unknown
  parts, strres syntax errors), and dump every store with the log lines.
  All 9 outputs match (26,000 lines). The only normalisation is the
  libraries' `std::stoul` exception texts (the port has MSVC's, as the
  official build; statsref has libstdc++'s).
* Found on the way: `ASSERT` prints its expression after macro expansion
  (`disabledWhen <= 0x3f`), `ASSERT_OR_RETURN` as written;
  `check_debug_lines.py` now checks that, and it fixed one A2 call
  (`_stricmp`, MSVC's `strcasecmp`).

### 13. PIE models (cluster A, step 4)

* WzGfx has the model loader: `imdload.cpp` (`imdload_c`) with `imd.h`,
  the model structures of `ivisdef.h` (`ivisdef_h`) and `pienormalize.h`
  (`pienormalize_h`). The models are two pools, `iIMDShape` (the display
  levels, chained through `nxt`) and `iIMDBaseShape` (the part for game
  state); a model pointer is a 1-based handle, 0 = nullptr, and the model
  map is a `Scripting.Dictionary` of base shapes by name. The other
  projects read a base shape through GfxApi's `iIMDBaseShape_<field>`.
  The vertex data that the C++ uploads to GPU buffers while it loads stays
  in the shape (`vbo_*`) for cluster D's renderer. The stats now get real
  models (the `modelGet` stand-in is gone).
* Shared: `sscanf_vb` (C's `sscanf`: `%d %i %u %x %f %s %c %n %[set]`, `*`,
  widths) and `fixedpoint_h` (moved from the exe).
* Float math (`glm_vb`): each float operation is done in Double and
  rounded to Single through memory (`F32`), which gives C's correctly
  rounded result, with glm's own formulas (no SIMD). Two x87 traps (see
  CONVENTIONS.md, "Float arithmetic"): VB's Double to Single conversion
  raises an overflow where C gives an infinity, and the sticky "invalid"
  flag that a new NaN (or a comparison with one) leaves makes VB's next
  float to integer conversion, anywhere, fail with run-time error 6.
* Not ported yet: the textures (`getTextures`, `_imd_load_level_textures`,
  `modelReloadAllModelTextures`, the team colours: cluster D) and the
  shadow edge lists (cluster E). `enumerateLoadedModels` goes in insertion
  order (C++: hash order). The faction models that `levLoadData` preloads
  come with A5.
* Checked: statsref now compiles the real `imdload.cpp` (all but the
  texture code, by line ranges), its GPU buffers replaced by recording ones
  (`STATSREF_SUBSTITUTIONS` in run_tests.py), and both sides dump a model
  the same way (`modelDump`: every field, floats as their bits, and the
  uploaded vertex, normal, texcoord, tangent and index data). The 866 model
  names in `modelGet`'s folders of base and mp, and 52 crafted files (every
  error path, comments and CRLF, PIE 2 texture scaling and animation, PIE 4
  level settings, normals and tangents with a zero tangent, shadow points
  and polygons, both connector formats, object animation, EVENT models,
  `debugReloadDisplayModelsForBaseModel`), match bit for bit: 19 of 19
  outputs with the stats (20 MB).
* Found on the way: `check_debug_lines.py` did not expand a `PRIu32` that
  ends a format string.

### 14. The game's map (cluster A, step 5)

* The exe has `map.cpp` and `map.h` (`map_c`, `map_h`): the tiles
  (`psMapTiles`; the struct is `MAPTILE_`, a `MAPTILE *` a tile index), the
  tileset files (enum names, ground types, ground per tile, decals, read
  from `fileLoadBuffer` as C reads them), the riverbed, the blocking and
  aux maps, the continents, `mapLoad` / `mapLoadFromWzMapData` /
  `loadTerrainTypeMap` / `mapShutdown` / `mapReloadGroundTypes`, and the
  queries `map_Height`, `map_LineIntersect`, `map_Intersect`,
  `getTileMaxMin`, `tileSetFire`, `fireOnLocation`. With it: `gateway.cpp`
  (`gateway_c`), `random.cpp` (`random_c`: the Mersenne twister and the
  game's random numbers), `fpath.cpp`'s tile blocking test (`fpath_c`),
  `game.cpp`'s terrain type overrides (`game_c`), wzmaplib's inline tile
  helpers (`wzmaplib_map_h`), and the classes the game gives wzmaplib:
  `WzMapPhysFSIO`, `WzMapBinaryPhysFSStream`, `WzMapDebugLogger`. WzGfx has
  `gfx_api::imageLoadFilenameFromInputFilename` (`gfx_api_c`, as built
  with `BASIS_ENABLED`). `math_vb` has `MUL32` (uint32_t multiplication).
* Stand-ins: `syncDebug` (WzNet's sync log, cluster B), the lightmap,
  `loadTerrainTextures` and `markTileDirty` (clusters D and E). Not ported
  yet: the danger maps and their thread (`mapInit`, `mapUpdate`,
  `threatUpdate`) and `mapObjIsAboveGround` (they need the objects, B), the
  visibility and map saving (cluster I).
* Checked: statsref now compiles wzmaplib, `map.cpp` (all but the parts
  above), `gateway.cpp`, `random.cpp` and the pieces named above.
  `Warzone.exe --vbtest` and statsref load each of the 68 map folders as
  `loadGame` does (wzmaplib through `WzMapPhysFSIO`, the terrain types,
  `mapLoadFromWzMapData`), each with one of the three tilesets, the null
  tileset or an unsupported one in turn, some with the terrain type
  overrides or a ground type reload; they dump every tile (texture, height,
  water level, ground, flags, continents, blocking bits), the ground types,
  terrain types and gateways, and run 520 queries per map (heights, line
  intersections, tile intersections, blocking tiles, fires). All 68 match,
  with the stats and models: 87 of 87 outputs (48 MB). The five
  script-generated maps (`game.js`) do not load on either side
  (`runMapScript` is a stand-in, as in mapref).
* The factions: `faction.cpp` (`faction_c`: the replacement model tables,
  `getEnabledFactions`, `getFactionModelName`), and `levLoadData`'s
  preload of the faction models (in model load order; C++: hash order).
  WzNet's NetPlay stand-in has the players' `faction`. Checked with
  statsref (`src/faction.cpp` compiled whole): 88 of 88 outputs match.
* Still to come for the cluster A milestone: the structures and droids of
  a map (cluster B's first steps, section 15).

### 15. The objects (cluster B, step 1)

* The object model (`docs/CONVENTIONS.md`, "Objects"): `DROID`,
  `STRUCTURE` and `FEATURE` are UDTs in pools whose arrays start at
  `type * &H10000000 + 1`, so one handle is the typed and the base pointer;
  the `PagedEntityContainer`s (`paged_entity_container_h`) take each slot
  once, in order, as C++ does. The lists are `OBJ_LIST`s over a shared node
  pool (`objlist_vb`).
* Ported: `objmem.cpp` whole (lists, the destroyed list, ids, flag
  positions, the lookups by id, the reference checks), `objects.cpp`,
  `baseobject.cpp` (but `checkObject` and `_syncDebugObject`),
  `group.cpp` (but `orderGroup` and `setSecondary`), `feature.cpp` (the
  features: `buildFeature`, `removeFeature`, `destroyFeature`,
  `featureUpdate`, the sync log; not `featureDamage`), `game.cpp`'s
  `loadWzMapFeature`, the type definitions of `basedef.h`, `featuredef.h`,
  `structuredef.h`, `droiddef.h`, `movedef.h`, `orderdef.h`,
  `actiondef.h`, `weapondef.h`, `displaydef.h`, `positiondef.h`, and from
  other files what these need: DROID's and STRUCTURE's constructors and
  destructors, `droidGetName`, `isTransporter`, the target setters,
  `STRUCTURE::isFactory`, `IsStatExpansionModule`, the structure bounds,
  `hasSharedVision`, the alliances (`ai_c`), the start and derrick
  positions (`qtscript_c`), `getDroidOrderName`, `StatIsStructure`,
  `getTileStructure`/`getTileFeature`, map.h's `TileHas...`, gtime's
  `deltaGameTime` and stop count, macros.h's `MIN`/`MAX`.
* Stand-ins (stubs_game, all doing nothing while their modules have no
  state): effects, scores, messages, sounds, script events, the interface,
  `fpathRemoveDroidData`, `formationLeave`, `visRemoveVisibility`,
  `assignFactoryCommandDroid`, `_syncDebugIntList`.
* Checked: statsref now compiles `objmem.cpp`, `objects.cpp`,
  `group.cpp`, `trig.cpp` whole and the rest by line ranges; after each
  map's queries, `mapfeatures` places the map's features as `loadGame`
  does (with mp's stats) and dumps the id counters, the feature and oil
  lists (ids, stats, positions, rotations, bodies, foundation depths,
  models), the derricks and start positions, and every tile's object and
  blocking bits. All 68 maps match (10,262 features): 89 of 89 outputs.
* Found on the way: VB's `Mod` binds looser than `*`, `/` and `\`, so C's
  `w % 2 * TILE_UNITS / 2` needs parentheses; `tools/check_mod.py` finds
  these. A VB6 job unrelated to the port (another `/make` loop on this PC)
  makes VB6 builds stop without a message now and then: build again.

### 16. The structures (cluster B, step 2)

* Ported: from `structure.cpp` the aux bits, limits and factory flags
  (`initStructLimits`, `structureInitVars`, `resetFactoryNumFlag`), the
  walls (but the blueprint ones), foundations (`alignStructure`),
  `buildStructure` / `buildStructureDir` with modules, `setFunctionality`,
  the assembly points (`findAssemblyPointPosition`, `setAssemblyPoint`,
  `setFlagPositionInc`), `validLocation`, `removeStruct`, the derrick and
  power generator links (`checkForResExtractors`, `checkForPowerGen`,
  `releasePowerGen`, ...), `buildingComplete`, `getModuleStat`,
  `structureBodyBuilt`, `getStructureDamage`, `checkStructure`,
  `holdProduction` / `releaseProduction`; `research.cpp`'s
  `holdResearch` / `releaseResearch`; `visibility.cpp`'s tile visibility
  (`visTilesUpdate` with its wave cast, `visRemoveVisibility`,
  `setUnderTilesVis`, `revealAll`, `objSensorRange`); `wavecast.cpp`
  whole (`wavecast_c`); `game.cpp`'s `loadWzMapStructure` and
  `RemapWzMapPlayerNumber`; `droid.cpp`'s `nextModuleToBuild`;
  `multiplay.cpp`'s `turnOffMultiMsg`, `isHumanPlayer`,
  `scavengerSlot`; `mission.cpp`'s landing zones and
  `missionLimboExpand`. WzNet's NetPlay has the players' `position`.
  `glm_vb`'s float operations moved to `src/shared/float_vb` (the exe
  needs them: `structureBodyBuilt` computes in float).
* Stand-ins (stubs_game): the research and production queues
  (`cancelResearch`, `cancelProduction`), the power queue
  (`delPowerRequest`: power.cpp, with its 32.32 fixed point, comes in its
  own step), `droidRepairStopped`, the interface (`intGetResearchState`,
  blueprints, `cancelDeliveryRepos`, the debug key mappings, Ctrl+Shift),
  `sendStructureInfo`, the script events.
* Checked: statsref now compiles those parts of `structure.cpp`,
  `visibility.cpp`, `mission.cpp`, `multiplay.cpp`, `game.cpp`, and
  `wavecast.cpp` whole (the debug key mappings swapped for a stand-in by
  `STATSREF_SUBSTITUTIONS`). `mapplayers` sets up a skirmish or campaign
  game (type, players, start positions, `bMultiPlayer`), and `mapobjects`
  places each map's features and structures as `loadGame` does and dumps
  every structure (position, rotation, body, modules, weapons,
  functionality: assembly points, power links, walls), the function lists,
  the flags, and every tile's explored bits, watchers, sensors and aux
  bits. All 68 maps match (12,319 structures, 894 with modules): 89 of 89
  outputs. A `%p` in a log line is normalised (`run_tests.py`).
* Found on the way: `TILE_UNITS \ 2 * size` divides by `2 * size` in VB
  (written `(TILE_UNITS \ 2) * size`); `check_debug_lines.py` now joins a
  C++ format string split over lines, and knows
  `ASSERT_NOT_NULLPTR_OR_RETURN`.

### 17. The droids (cluster B, step 3) and the cluster A milestone

* Ported: from `droid.cpp` `reallyBuildDroid`, `droidSetBits`,
  `initDroidMovement`, the sums over a design (`calcDroidWeight`,
  `calcTemplateBody`, `calcDroidBaseBody`, `calcDroidBaseSpeed`,
  `calcTemplateBuild`, `calcDroidPoints`, `calcTemplatePower`,
  `calcDroidPower`; the C++ lambdas are a selector), the component getters
  and flight tests, the recycled experience (`droidInit`,
  `getTopExperience`, `add_to_experience_queue`); `move.cpp`'s `moveInit`
  and `updateDroidOrientation`; gtime's `gameTimeAdjustedIncrement`;
  `game.cpp`'s `loadWzMapDroidInit`.
* Stand-ins: `secondarySetState` (order.cpp, for transporters),
  `missionCanReEnforce`, `intAddTransporterLaunch`.
* Checked: `mapobjects` now loads each map's droids first, as `loadGame`
  does, and dumps them (name, type, position, rotation, body, weight,
  speed, components, weapons, blocking bits, group, visibility). All 68
  maps match (1,191 droids): 89 of 89 outputs, with the `%p` of
  `reallyBuildDroid`'s log line normalised.
* **The cluster A milestone is reached**: a hidden run (`Warzone.exe
  --vbtest`) loads the stats, every map with its droids, features and
  structures, and dumps them, identical to the real 4.7.0 code.
* Next in cluster B: power (`power.cpp`, 32.32 fixed point), the object
  grid (`mapgrid.cpp`, `pointtree.cpp`), research results and upgrades, the
  orders, actions and movement, combat and projectiles, the game loop
  (`gameLoop`, `gtime`'s clock), syncDebug in WzNet, then replays.

### 18. Power, the object grid and the sync log (cluster B, step 4)

* Ported: `power.cpp` whole (`power_c`; the 32.32 fixed-point `int64_t`
  values are Doubles, exact below 2^53, and `vsprintf` now prints a Double
  in full for `%lld`); `pointtree.cpp` (`pointtree_c`: the 64-bit Morton
  numbers as two Longs, `std::stable_sort` as a merge sort, the four-range
  query and the filters) and `mapgrid.cpp` (`mapgrid_c`); structure.h's
  `getBuilding...Points` and `_syncDebugStructure`. In WzNet,
  `sync_debug.cpp`'s log (`sync_debug_c`): each tick's entries and their
  CRC for the last 12 ticks, `nextDebugSync`, `resetSyncDebug`,
  `syncDebugGetCrc`/`SetCrc`, `_syncDebugBacktrace`, the dump and the
  verbose per-tick file. The exe's `syncDebug` calls now reach it
  (`NetApi.syncDebug`, `syncDebugIntList_`). `printf_vb` gained C's rule
  that `%.0d` prints no digits for 0.
* Not yet: sending, receiving and comparing desync logs
  (`DesyncLogOutputter`, `sendDebugSync`, `recvDebugSync`,
  `checkDebugSync`), which need the network.
* Checked: each map block now runs `syncreset`, `mapobjects`, `synclog`
  (the sync log as `dumpLocalDebugSyncLog` prints it, with its CRC), then
  `mapgrid` (80 queries of the five kinds per map, half of them centred on
  an object, with the unseen filter marking objects as seen), `mappower`
  (40 ticks of `updatePlayerPower` with power modifiers, a storage limit,
  power requests whole and precise, added and used power, a derrick of each
  player dying) and `synclog` again. statsref compiles `power.cpp`,
  `mapgrid.cpp`, `pointtree.cpp` and the log part of `sync_debug.cpp`. All
  68 maps match: 89 of 89 outputs, about 200,000 sync log lines with their
  CRCs among them.
* Next in cluster B: research results and upgrades, the orders, actions
  and movement, combat and projectiles, the game loop (`gameLoop`,
  `gtime`'s clock), then replays.

### 19. Path finding (cluster B, step 5)

* Ported: `astar.cpp` whole (`astar_c`: the cached PathfindContexts in LRU
  order, the per-tick blocking maps with their sync log checksums, the
  node heap, the path back) and the rest of `fpath.cpp` (`fpath_c`: the
  jobs and results, `fpathRoute`, `fpathDroidRoute`, `findNonblockingPosition`,
  `fpathExecute`, `fpathCheck`); crc.cpp's `crcSumVector2i` (its step is
  `crc_c`'s `crcSumI32`); baseobject.cpp's `getStructureBounds(BASE_OBJECT *)`
  now casts as the C++ does (a nullptr gave an invalid handle); feature.h's
  `isFeature`, `castFeature`.
* No path threads: a job runs when `fpathRoute` queues it, and its result
  waits for the droid's next call. The C++ sends jobs that could share a
  PathfindContext to the same thread in queue order, and fixes each job's
  blocking map when it is queued, so the paths do not depend on when or
  where a job runs. statsref keeps the real dispatch (two threads on this
  PC, each job run at once by its thread's context) and agrees.
* Checked: `mappaths` on every map: over 6 ticks, 3 destinations a tick,
  each wanted from 4 origins with a random propulsion, player and move type
  (4,536 routes, 1,091 of them failing), then up to 30 droids routed with
  `fpathDroidRoute` to points and structures (956), and 40 `fpathCheck`s;
  the sync log of each (blocking map checksums, the route lines with the
  path CRCs). 89 of 89 outputs.
* Speed: the first version took 93 s for these routes (statsref: about
  3 s). A profile (`prof_vb`) showed the A* exploration, at about 14 ms a
  route. Three changes brought it to a third: `iHypot` had gone through
  Variants (`i64Sqrt`), the context's arrays were indexed through
  `ctxs(c).` with array elements passed ByRef to small functions (VB locks
  the array for each such call), and every node made some 50 procedure
  calls. The A* core now takes the arrays as parameters and writes the
  heap comparisons, the estimates and the blocking test out in place. The
  test build (bounds checks on, no optimisation) runs 2.6 million node
  expansions in 4.2 s; `--release` drops the bounds checks (2.8x on such
  code in QuakeVB's probe). The whole `--vbtest` run went from 140 s to
  79 s.
* Next in cluster B: movement (`move.cpp`), then actions and orders,
  combat and projectiles, research results (`researchResult` and its
  upgrades, `eventResearchedHandleUpgrades`), the game loop, replays.

### 20. Movement (cluster B, step 6)

* Ported: `move.cpp` (`move_c`) but the three functions of the
  GAME_SYNC_OPT_CHANGE message (moveSetFormationSpeedLimiting,
  moveToggleFormationSpeedLimiting, recvSyncOptChange: with the game
  messages); the steering module new in 4.7 (`steering_c`: the manager and
  its force blending; `collision_avoidance_behavior_c`); `formation.cpp`
  whole (`formation_c`, a pool of FORMATIONs with per-player lists);
  `raycast.cpp`'s rayCast (`raycast_c`; getBestPitchToEdgeOfGrid waits for
  the display's clipXY); `action.cpp`'s VTOL landing search
  (`actionVTOLLandingPos`, spiralSearch: `action_c`); `cmddroid.cpp`'s
  hasCommander, cmdGetCommanderLevel, cmdDroidUpdateExperience
  (`cmddroid_c`); and from `droid.cpp` removeDroidBase, destroyDroid,
  vanishDroid, the droid sync log line (`syncDebugDroid_`), the levels and
  experience, calcDroidSpeed, droidOnMap and checkDroid (CHECK_DROID);
  structure.cpp's requestOpenGate, geometry.cpp's calcDirection,
  mission.cpp's missionIsOffworld, multiplay.cpp's responsibility tests,
  gtime.h's quantiseFraction and gameTimeAdjustedAverage.
* Stand-ins (stubs_game, and the same in statsref): actionDroid (the
  actions are next), pickupArtefact, the script triggers (pickup, droid
  moved, rank gained), getExpGain (100, as proj_InitSystem sets it), the
  campaign number (1) and fastExp option, the console, the move sounds
  (audio_PlayObjDynamicTrack fails, as with the sound off), the droid
  death effects and the interface's notices.
* Checked: `mapmove` on every map: every droid stopped, up to 48 sent to
  destinations of a fixed generator (three to a destination, so they form
  formations and push past each other; some without a formation), then 80
  ticks of gridReset and moveUpdateDroid for every droid, with stops,
  turns and new destinations at ticks 30 and 50. Each tick's sync log (the
  droid lines before and after each move, the route and blocking map
  lines) is compared by its CRC, the whole log of ticks 1 and 80 line by
  line, and the droids' movement state every 20 ticks: 1,123 orders and
  1,058 later ones, 5,040 ticks, droids in formations, bumps, pauses and
  shuffles. 89 of 89 outputs; the `--vbtest` run takes about 100 s (the
  movement about 20 s of it). Not reached by the maps' droids yet:
  squashing a person, picking up an oil drum, VTOLs (they come with the
  replays).
* Next in cluster B: actions and orders (`action.cpp`, `order.cpp`), combat
  and projectiles, research results, the game loop (`gameLoop`, `gtime`'s
  clock), then replays.

### 21. Actions and orders (cluster B, step 7)

* Ported whole: `action.cpp` (`action_c`), `order.cpp` (`order_c`),
  `cmddroid.cpp` (`cmddroid_c`), `lib/framework/geometry.cpp` (Affine3F,
  `framework_geometry_c`; in the exe, as only the exe has Vector3i), and
  group.cpp's orderGroup and setSecondary. From `droid.cpp`: the per-tick
  `droidUpdate`, cancelBuild, recycleDroid, droidRemove, the building,
  demolishing, restoring and repairing (droidStartBuild .. droidUpdateDroidRepair),
  the muzzle position, the VTOL rearming tests, the sensor tests, the
  selection (SelectDroid, DeSelectDroid), droidSqDist. From `structure.cpp`:
  structureBuild, structureDemolish, structureRepair, the rearm pads
  (findNearestReArmPad, ensureRearmPadClear), the sensor tests, the muzzle
  position, calcLineBuild, canStructureHaveAModuleAdded. From `ai.cpp`:
  validTarget, checkAnyWeaponsTarget, aiObjectIsProbablyDoomed,
  aiObjectAddExpectedDamage. From `visibility.cpp`: visibleObject (with its
  ray callback), visGetBlockingWall, lineOfFire, areaOfFire, arcOfFire. From
  `projectile.cpp` (`projectile_c`): the weapon ranges, proj_Direct,
  calcDamage, establishTargetHeight. And baseobject.cpp's checkObject and
  _syncDebugObject, stats.cpp's objActiveRadar and objRadarDetector,
  display.cpp's CheckInScrollLimits (`display_c`), mission.cpp's
  getLandingX/Y.
* WzNet 1.9: NetPlay.players[i].difficulty. WzGfx 1.8: iIMDShape's
  objanimpie, objanimcycles, objanimtime (droidUpdate's animations).
* Stand-ins (stubs_game, and the same in statsref): the targeting
  (aiBestNearestTarget finds nothing, aiChooseTarget, aiUpdateDroid), combat
  (combFire, droidDamage, the projectiles' sync log and check), the
  transporters, the missions' reinforcements, the camera, the interface's
  select modes and widgets, the network messages (sendDroidInfo: a ModeQueue
  order is lost until the game messages come), the script events, the
  sounds and effects, the console, lighting.
* Checked: `maporders` on every map, after `mapmove`: 16 structure types made
  buildable for all players, power set, a third of the droids and a quarter
  of the structures damaged, secondary orders (hold, pursue, short range,
  repair level), then for up to 64 droids an order by droid type from a
  fixed generator (build, line build, demolish, repair, droid repair,
  observe, attack, guard, move, scout, patrol, hold, RTB, RTR, stop,
  recycle, a click through chooseOrderObj, queued orders; VTOLs circle,
  patrol, attack, RTB) and 90 ticks of droidUpdate for every droid (each
  player's list, as the game loop iterates it) and updatePlayerPower, with
  new orders at ticks 20 and 40. Each tick's sync log is compared by its
  CRC (the whole log of ticks 1 and 90 line by line), the droids' orders and
  actions every 30 ticks, the structures being built or damaged, the counts
  and power. 89 of 89; `--vbtest` takes about 125 s.
* DROID's constructor leaves actionStarted and actionPoints uninitialised
  (C++ reads them only after setting them); the test zeroes them on both
  sides before it prints them.
* tools/check_debug_lines.py now reads a whole multi-line macro call (GCC's
  `__LINE__` is the macro name's line; `//` comments inside dropped) and
  takes `<stdint.h>`'s limits (`UINT8_MAX` 255) before types.h's fallbacks.
* Next in cluster B: combat (ai.cpp's targeting, combat.cpp, projectile.cpp,
  the damage), research results, the structures' update (factories,
  research, repair facilities, rearm pads), the game loop, replays.

### 22. Combat (cluster B, step 8)

* Ported whole: `projectile.cpp` (`projectile_c`, but objectShimmy and
  justBeenHitByEW, the renderer's shimmer), `combat.cpp` (`combat_c`),
  `ai.cpp` (`ai_c`: the targeting, aiUpdateDroid), `visibility.cpp`
  (`visibility_c`: processVisibility, visUpdateLevel, the spotters),
  `difficulty.cpp` (`difficulty_c`). From `droid.cpp`: droidDamage,
  calcDroidMuzzleLocation; from `structure.cpp`: structureDamage,
  destroyStruct, calcStructureMuzzleLocation, electronicDamage; from
  `feature.cpp`: featureDamage; stats.cpp's objHasWeapon.
* Projectiles are objects too: `projectiledef_h`'s PROJECTILE (SIMPLE_OBJECT's
  fields first), in the pool aProj (paged_entity_container_h's
  ProjectileContainer, which reuses freed slots as the C++ default does);
  basedef_h's SIMPLE_OBJECT_* properties take projectile handles.
* The difficulty's damage modifiers are 0 until setDifficultyLevel or
  resetDamageModifiers runs (as in C++): the game's start sets them, the
  test calls resetDamageModifiers.
* Stand-ins: the structures' own update (structureUpdate, aiUpdateStructure:
  defences do not fire yet), gifts and captures (giftSingleStructure,
  giftSingleDroid), the messages (addMessage), the script events, the
  interface's damage bars, the multiplayer statistics, the effects and
  sounds, the console.
* Checked: `mapcombat` on every map, after `maporders`: up to 40 armed
  droids moved 3 to 6 tiles from an enemy structure and ordered to attack
  it, scout to it, hold or stop, then 60 ticks in the game loop's order
  (visUpdateLevel, gridReset, processVisibility, fpathUpdate,
  cmdDroidUpdate, per player updatePlayerPower and droidUpdate,
  proj_UpdateAll, objmemUpdate), each tick's sync log by its CRC (the
  whole log of ticks 1 and 60), the projectiles in flight, the fighters every
  20 ticks, the damaged structures. 3,780 ticks: 2,071 hits on structures,
  24 on features, 17 droids and 3 structures destroyed, up to 78 projectiles
  in flight. 89 of 89, the whole suite 3346 of 3346; `--vbtest` takes about
  175 s.
* A new test command `syncdumpall 1` makes mapmove, maporders and mapcombat
  dump every tick's whole sync log (for finding a divergence; the scratchpad
  run_subset.py runs one map block that way on both sides).
* Next in cluster B: the structures' update (structureUpdate: factories,
  research, power generators, repair facilities, rearm pads, defences
  firing), research results, the game loop (gameStateUpdate and gtime's
  clock), replays.

### 23. The structures' update, research and the game state's update (cluster B, step 9)

* Ported: the rest of `structure.cpp` that runs in a game (`structure_c`:
  structureUpdate, aiUpdateStructure with the defences, sensor towers,
  rearm pads and the lassat, the repair facility's state machine, the
  factories' production with the production runs (asProductionRun),
  structPlaceDroid and placeDroid, the droid limits, the resistance and
  self-repair, handleAbandonedStructures); `research.cpp` whole
  (`research_c`: researchAvailable, researchResult, the upgrades of
  eventResearchedHandleUpgrades, the component replacement, cancelResearch,
  the rewards); from `wzapi.cpp` what research needs (`wzapi_c`:
  setUpgradeStats, getUpgradeStats, GameEntityRules, getUpgradesObject,
  constructStatsObject: the Stats and Upgrades objects of the scripts,
  without a script engine); `loop.cpp`'s game state (`loop_c`:
  gameStateUpdate, countUpdate, the droid counts, the pause states);
  `map.cpp`'s danger maps (mapInit, mapUpdate, threatUpdate,
  dangerFloodFill: the danger thread's work is a plain call where it is
  woken); from `droid.cpp` buildDroid, vtolReadyToRearm, fillVtolDroid;
  from `mission.cpp` missionDroidUpdate, missionTimerUpdate,
  buildMissionDroid; from `intdisplay.cpp` the factories' pending states;
  `multiplay.cpp`'s getPlayerGenericName and
  findResearchingFacilityByResearchIndex. WzNet 1.10 (the players' names,
  NETisReplay, NETflush), WzGfx 1.9 (a shape's next level, animation
  frames, points), WzFramework (ASSERT_HELPER_FAILURE).
* Research results need no script: in 4.7.0 eventResearchedHandleUpgrades
  applies research.json's results in C++ through the Upgrades object.
* Stand-ins: the scripts (updateScripts, the triggers), the network
  messages (SendDroid, sendQueuedDroidInfo, sendResearchStatus,
  sendPlayerGameTime: a game starts research as recvResearchStatus does,
  through a message), the multiplayer statistics, the interface's research
  notices, the effects, the misc models, the gifts.
* Checked: `mapbase` on every map, after `mapcombat`: research and
  production 50 times faster, every component available, the players'
  power and rules.js's droid limits set, VTOLs sent to rearm, half the
  damaged droids to repair, each idle lab given a topic and each idle
  factory a template, then 100 ticks of gameStateUpdate (new work every 25),
  each tick's sync log by its CRC; the research state and the stats'
  upgrades are saved first and restored after (the maps share the stats).
  63 maps x 100 ticks: 2,425 droids built in factories, 442 topics
  researched (with their upgrades), defences targeting (529 structure
  lines with a target), up to 130 projectiles in flight, the danger maps on
  the skirmish maps. 89 of 89; `--vbtest` takes about 335 s. The repair
  facilities are barely reached (3 repairs) and no map has VTOLs with a
  rearm pad: rearming is not exercised yet.
* Found on the way: VB's assert lines took the caller's line where C++'s
  ASSERT_HELPER logs its own (checkDroid's CHECK_DROID callers); seven C++
  `&&`/`||` evaluated a logging function (isVtolRearming, vtolEmpty,
  actionReachedBuildPos, droidResistance, isAttacking) where C++ does not.
  The research "grade" log lines walk a std::unordered_map (the library's
  hash order): run_tests compares each run of them sorted.
* Next in cluster B: gtime's clock (gameTimeUpdate, the game queues),
  the network messages' local half (netqueue, GAME_DROIDINFO and the other
  messages decoded and run as in a game), then replays.

### 24. The game clock and the game messages (cluster B, step 10)

* Ported: `gtime.cpp` whole (`gtime_c`: gameTimeUpdate with the
  GAME_GAME_TIME synchronisation, the latency, the CRC check of each tick's
  sync log, realTimeUpdate, the clock modifier as a numerator and a
  denominator); `loop.cpp`'s gameLoop (its render part, renderLoop, is a
  stand-in); in WzNet, `netqueue.cpp` and `nettypes.cpp` whole
  (`netqueue_c`, `nettypes_c` and the classes NetMessage,
  NetMessageBuilder, NetQueue, NetQueuePair, MessageWriter, MessageReader;
  the std::list and its iterators are serial numbers), `netplay.cpp`'s
  local half (`netplay_c`: NETrecvGame, messageTypeToString, the connection
  states; NETsend, NETflush and NETrecvNet up to their no-connection
  returns), `netlog`'s packet counters, `sync_debug.cpp`'s CRC check and
  desync dumps; `multibot.cpp` and `multistruct.cpp` whole (`multibot_c`,
  `multistruct_c`: droid and secondary orders, production, research holds,
  the lassat and the debug messages); from `multiplay.cpp` recvMessage and
  the message rules, IdToDroid .. IdToPointer, the research messages, the
  sync requests, canGiveOrdersFor; `move.cpp`'s formation speed option;
  `design.cpp`'s intValidTemplate and `template.cpp`'s designableTemplate.
  WzNet's INetGame grew (the clock's callbacks), so WzNet was built with
  `--break-compat WzNet`; the exe now hands WzNet its INetGame at startup.
* The messages' API is in CONVENTIONS ("The network messages"): NETQUEUE is
  a Long, `auto w = NETbeginEncode(...)` is a MessageWriter, and each
  reader/writer overload pair is one procedure taking either As Object.
* Stand-ins: the connections (cluster J); the chat, beacons, votes, kicks,
  data checks, gifts and alliances, player stats and a player leaving
  (popped undecoded); replay files.
* Checked: `mapnet` on every map, after `mapbase`: a skirmish host without
  a network (the selected player the only human, the game queues
  initialised as stageTwoInitialise does, gameTimeInit), on a clock that
  moves 100 ms a frame (statsref's wzGetTicks, WzFramework's
  wzSetTestTicks). Labs, factories and droids get their work through the
  messages: a topic each idle lab may start (sendResearchStatus), a
  template each idle factory (structSetManufacture, ModeQueue), holds,
  releases and cancellations, up to 24 droid orders a round (queued ones
  too) and secondary orders, the debug messages once. 80 frames of
  realTimeUpdate and gameLoop; each frame prints gameTime and the header
  (time, entries, CRC) of the sync log sendPlayerGameTime closed, so every
  tick's log is compared by its CRC. Over 63 maps: 80 ticks each, 1,433
  production orders, 1,622 holds, releases and cancellations, 4,043 droid
  orders, 2,149 secondary orders, 253 research starts and 97 topics done,
  1,616 droids built; no waiting, no CRC error. 89 of 89; `--vbtest` takes
  about 680 s.
* Found on the way: the exe never called NetSetGame (INetGame was unused
  until now); an earlier block's players mask (two humans) made the clock
  wait for a player nobody sends for, on both sides alike, so the first
  89 of 89 tested nothing: `mapnet` checks that the clock advances.
* Next in cluster B: replays (`netreplay.cpp`'s reading, the options
  handler that sets up the game a replay describes). Playing one needs the
  game started from its options, and rules.js's effects at its start.

### 25. Replay files (cluster B, step 11)

* Ported: `netreplay.cpp` whole in WzNet (`netreplay_c`: the file's
  preamble, the settings JSON, the embedded map, every processed game
  message, the end-of-game JSON; the save thread is plain writes, which
  gives the same bytes), nettypes' NETloadReplay, netplay's
  NET_InitPlayer, the netcode versions and NetPlay.players as a vector (the
  replay viewer's slot 23); nettypes.h's `ReplayOptionsHandler` is an
  interface of WzNet; in the exe, `multiint.cpp`'s replay options
  (`WZGameReplayOptionsHandler`: saveOptions, saveMap, restoreOptions, ...)
  with the JSON of the game, ingame and player structs, `multistat.cpp`'s
  PLAYERSTATS and its JSON (an EcKey is its public key's bytes: no
  libsodium), `multiplay.cpp`'s clearPlayerName; PhysFS's big-endian
  reads and writes and `WZ_PHYSFS_cleanupOldFilesInFolder` in WzFramework.
  WzNet's INetGame grew again (`--break-compat WzNet`).
* The netcode's minor version is the count of annotated version tags,
  which the shallow checkout cannot give: 0 on both sides for now. An
  official replay's "minor" key will give the real one.
* Checked: `maprecord` and `mapplay` on every third skirmish map (15
  maps), on the map loaded afresh each time: `mapnet`'s game for 40
  frames recorded as a replay (gameSRand(4242), NETreplaySaveStart, the
  lobby's options as loadConfig sets them), the file shown as text (its
  settings JSON, the messages' count by type and CRC, the end JSON), then
  played back by a replay viewer (NETloadReplay, restoreOptions without its
  level part: the tests load maps by folder). VB and C++ write the same
  file (329 messages with the same CRC on the 10-player map) and play it
  back alike; 89 of 89. Played back, the game repeats the recording tick for
  tick on 14 of the 15 maps; on the last, the viewer's own queued
  secondary order adds one log line at the end (C++ does the same).
  `--vbtest` takes about 19 minutes now.
* Found on the way: a test that reused a map after `objshutdown` left map
  tiles pointing at freed structures (statsref crashed in moveOpenGates;
  VB read stale handles), and the structures built during a recording
  levelled the ground: each recording and playback now loads the map
  afresh. VB's NetPlay players started at position and difficulty -1;
  C++'s value-initialised PLAYERs start at 0, as VB now does. The debug
  droid's id comes from generateNewObjectId, as buildDroid's.
* Next: official replays need the game started from a replay's options
  (restoreOptions' level part, the level load, the game's start) and
  rules.js, which every skirmish and multiplayer game runs (its start sets
  power, research, limits, bases; its timers drop oil drums with
  syncRandom): cluster C's script engine, at least for rules.js. AI scripts
  are not needed (a replay viewer is responsible for no AI; their orders
  are in the replay).

### 26. The official build's replays and the game's start (cluster B, step 12)

* The official 4.7.0 Windows build (`warzone2100_win_x64_archive.zip`
  from the GitHub release) is in `ref/official/wz470` (not part of the
  port; `ref/official/cfg_rec` is its config folder). `--skirmish=<test>
  --autogame --headless` plays a bot game at 500 times speed and saves its
  replay when the game ends. The first one (`ref/official/replays/
  rush_10p_autogame.wzrp`): Sk-Rush, player 0 and nine Nexus AIs, 30 game
  minutes (18,000 ticks), 20,438 messages (2,056 droid orders, 303
  structure messages, 80 research messages, 17,998 game times). The
  "highground" test was not taken: the build started its default skirmish.
* Its settings give the netcode version: 0x4000.104 (netplay_c and
  statsref's netplay_config.h).
* Checked: `replayinfo` reads an official replay as the game does
  (NETreplayLoadStart with the tests' handler, every message): the options
  as restored and saved again, the messages by player and type with their
  CRC, the first ten in hex. VB and C++ read it alike (a player's identity
  key kept as its bytes on both sides); 90 of 90.
* The official build's own sync logs of the replay's playback, the
  milestone's reference: `ref/official/synclogs/rush_10p_autogame/
  sync<gameTime>_p23.txt`, the first 120 ticks (gameTime 102 to 12002),
  written with `--loadreplay=<file> --autogame --headless
  --debug-verbose-sync-logs-until=12002` (player 23 is the replay viewer).
  The first tick's 1,662 entries: 578 of the map's objects
  (generateSynchronisedObjectId, alignStructure), 1,070 of rules.js's
  eventGameInit for the ten players (setPower, researchResult,
  setUpgradeStats), then gameLoop and gameStateUpdate.
* The game's start, ported (three forks):
  * main.cpp's loops: initSaveGameLoad, the title loop (start, run, stop),
    startGameLoop, runGameLoop, stopGameLoop, mainLoop's switch on the game
    mode and realmain's start-up switch; wrappers.cpp (titleLoop, the
    host-launch and headless flags, the win/lose flags); frontend.cpp's
    changeTitleMode and SPinit; titleui (the `WzTitleUI` interface and
    `WzOldTitleUI`, whose menus other than STARTGAME and LOADSAVEGAME are
    stand-ins: cluster F); init.cpp's three stages with their shutdowns,
    InitialiseGlobals, frontendInitialise and frontendShutdown; WzNet's
    NET_InitPlayers.
  * game.cpp: loadGameInit (with its replay branch), gameLoad of a `.gam`
    (version 8 and below: gameLoadV7; later versions are save games,
    cluster I), loadGame whole (its save-game parts call stand-ins),
    GameLoadDetails' getMap; mission.cpp's startMission with every mission
    kind, initMission, the no-go areas; the init parts of message.cpp,
    transporter.cpp and scores.cpp.
  * The multiplayer start: multiopt.cpp (multiGameInit, gameInit,
    playerResponding), multijoin.cpp (clearPlayer, destroyPlayerResources),
    multigifts.cpp's alliances, multilimit.cpp's applyLimitSet,
    makePlayerSpectator, initTemplates, loadMultiScripts with readAIs (the
    script loading itself is cluster C's).
* Headless mode: gfx_api_null.cpp's null context in WzGfx's stand-in
  context (`gfx_api_context_initialize` takes C++'s backend type:
  `--break-compat WzGfx`), and lib/sdl's sdl_Null_Impl as
  `win_Null_Impl` (no window; swapWindow throttles to 120 frames per
  second).
* Stand-ins changed so that the start runs: iV_LoadImageFile returns an
  empty Collection for an IMAGEFILE (frontend.wrf and vidmem.wrf load
  images), texLoad keeps the tileset's name (the map's ground types follow
  it), and the view data loaders return their file name as the C++ does.
* Found: a vbNullString assigned to a class's Public String field from
  outside the class becomes "" (VB's implicit Property Let copies it):
  every built-in map's LEVEL_DATASET had a non-null realFileName and was
  loaded as a map package. LEVEL_DATASET's nullable strings are now
  property pairs (CONVENTIONS, "VB6 pitfalls").
* Checked: `tools/replay_test.py` plays the official replay in
  Warzone.exe (`--loadreplay --autogame --headless
  --debug-verbose-sync-logs-until`) and compares the sync logs tick by
  tick. The first tick's 578 object entries match the official build's;
  the first difference is rules.js's first setPower. With rules.js's 1,070
  lines set aside, the first tick matches exactly (592 of 592 entries,
  gameStateUpdate's pseudorandom number included). From the second tick
  on, the game state differs: rules.js's power, research, limits and
  timers are missing. The regression suite: 3346 of 3346, and 90 of 90 (`--vbtest` 17 minutes).
* Next: cluster C's script engine for rules.js (eventGameInit, its timers),
  then the replay's playback runs on to the milestone.

### 27. The script engine (cluster C, step 1)

* Ported whole, in two rounds of parallel forks against one design (the
  names and the signature rule: CONVENTIONS, "The scripts"):
  * wzapi.cpp (4,856 lines) -> `wzapi_c`: every API function, the
    Upgrades and Stats objects (B9's part kept), the useful constants,
    the static player data, the map tiles array, the queue of objects
    removeObject removes;
  * qtscript.cpp and qtscript.h -> `qtscript_c` (+ `GROUPMAP`,
    `timerNode`): the engine (instances, timers with updateScripts'
    order, labels in std::map order, groups), every triggerEvent*, the
    label and group API, loadPlayerScript / loadGlobalScript,
    prepareScripts, the script states;
  * quickjs_backend.cpp -> a Chakra backend on the JsRT API of Windows'
    legacy chakra.dll: `jsrt_win` (declarations, helpers),
    `chakra_backend_c` (the object conversions convDroid ... convMax with
    QuickJS's property names, order and flags; callFunction with event
    namespaces; unbox/box; include, includeJSON, setTimer, queue,
    removeTimer, namespace, profile; JSON <-> JS), `chakra_functions_c`
    (the 156 IMPL_JS_FUNC wrappers, the hand-written ones, the Upgrades
    getters/setters, registerFunctions whole: 163 functions),
    `ChakraScriptingInstance` (the instance, its 61 event handlers),
    `wzapi_execution_context`, `ChakraTimerFunc` (`ITimerFunc`).
* Behaviour kept from the official build's QuickJS (Bellard's): a failed
  SCRIPT_ASSERT logs and the script goes on (no exception); one JsRT
  runtime with a context per instance; eval only in campaigns; Promise
  removed; QuickJS's debugger builtins (Nexus and Cobra use
  debugGetCallerFuncObject) as JS functions over Function.caller.
* `tools/check_debug_lines.py` now expands nested macros
  (SCRIPT_ASSERT_PLAYER -> SCRIPT_ASSERT -> debug), reads macro calls
  over several lines, and finds functions defined inside namespaces,
  classes and structs (unbox<T>::operator(), constructors in a class
  body): 2,300 log calls, 0 problems, planted errors still caught.
* Found: VB6 compiles a module-level declaration after the first
  procedure as an error ("Only comments may appear after End Sub"): a
  scan for it before building saves a compile.
* Checked: the official replay with rules.js running in Chakra: **120 of 120
  reference ticks match** the official build's sync logs (gameTime 102 to
  12002, every entry, rules.js's setPower, research and upgrade lines
  included), with no desync reported by the replay's own game-time CRCs in
  those ticks. The whole replay (18,000 ticks, 30 game minutes) does not yet
  stay in sync: the replay's game-time CRCs first differ at gameTime 456,202
  (7.6 minutes, the first combat), and from minute 9 on at every tick (the
  official build plays it in 29 s with none). Found and fixed on the way:
  loop.cpp's renderLoop was a stand-in, so multiplay.cpp's multiPlayerLoop
  never set ingame.TimeEveryoneIsInGame; the replay viewer then waited for
  its own game queue and processed its own GAME_DROIDINFO (a desync at
  431,002), and the headless replay never fast-forwarded (both now ported:
  the replay plays in about 10 minutes instead of 30); WzPlatform's event
  queue was off without a window, so wzQuit's quit event was dropped and a
  headless game never ended (now it quits at the game's end). The next
  difference needs the official build's per-tick logs to tick 500,002.
  `run_tests.py` now ends with the official replays (`--replays` runs only
  them).
* The tests follow: `vbtest_vb` calls the engine's prepareScripts once (no
  script loaded: the triggers run over no instance, as statsref's
  stand-ins do); statsref compiles the real renderLoop and
  multiPlayerLoop, their interface, display and network callees stand-ins
  (`tools/statsref/statsref_render.cpp`, as stubs_game's). The
  regression suite: 3346 of 3346, 90 of 90, the official replay's 120
  ticks.
* Speed (the whole replay headless, 18,000 ticks): 10 minutes with the
  fast-forward; a profile (`prof_vb` timers around the tick's parts) put
  more than half of it in the sync log's int lists, whose CRC shifted
  each byte with `SHRU` -> `BIT_` -> `2 ^ n` (a floating point power).
  `BIT_` is now a table and the CRC's byte step is written out (both in
  `src/shared`): **225 s**, the sync log unchanged (120 of 120). The
  profile now: droid updates 94 s (their sync lines 37 s), visibility
  61 s, structure updates 37 s, the map's update 10 s, the grid 5 s,
  projectiles 5 s, rules.js 1.6 s.
* Next: the remaining desync (the official per-tick logs around gameTime
  456,202: the first structure destroyed in the game, a scavenger
  defence); the playback's speed (225 s against the official 29 s:
  visibility and the droid update next);
  more official replays (other maps, longer games, scavengers; campaigns
  are cluster I); then cluster C's milestone, the bots playing a headless
  skirmish to the end.

### 28. The AI skirmish (cluster C, step 2): cluster C's milestone

* The user's priority (2026-10-28): the game itself before the replays
  (their desync and speed wait).
* The skirmish's start from the command line (`--skirmish=<test>
  --autogame --headless`), ported (fork S): multiint.cpp's
  WzMultiplayerOptionsTitleUI (start whole, its screen cluster F's), the
  test file's settings (loadMapChallengeSettings, loadMapPlayerSettings,
  resolveAIForPlayer, resetPlayerConfiguration), startHost, the ready
  status, startMultiplayerGame, SendFireUp; multiopt.cpp's hostCampaign
  and sendOptions; multilimit.cpp's createLimitSet; WzNet's NEThostGame
  (without the network) and player info broadcasts (`--break-compat
  WzNet`: INetGame grew); component.cpp's player colours;
  wrappers.cpp's displayGameOver.
* What the AIs reach (fork G): multigifts.cpp whole (gifts, captures,
  artefacts), droid.cpp's giftSingleDroid and tile pickers,
  structure.cpp's giftSingleStructure and rewards, chat.cpp's
  InGameChatMessage (chat to the AIs), multiplay.cpp's beacons,
  message.cpp's message and proximity lists.
* Found by the first games:
  * The x87's sticky "invalid" flag crosses the Chakra boundary: a NaN
    made in a script turned VB's next float to integer conversion into
    an overflow (error 6) in correct API code. Cleared after every call
    into Chakra and on entry to every native function (CONVENTIONS, "The
    scripts").
  * pickStructLocation and structureCanFit passed a STRUCTURE_STATS
    index where C++ converts to BASE_STATS * (the handle STAT_STRUCTURE +
    index): the builders of BoneCrusher and Cobra could not place a
    building.
* Checked (`tools/ai_skirmish.py`, the tests in `tools/skirmish_tests`):
  Nexus, NullBot, BoneCrusher and Cobra on Bloat, free for all, player 0
  a spectator: **the game plays to its end** (about 23 game minutes, a
  winner before the 30-minute limit; 1,764 attacks, 484 objects
  destroyed, 380 droids and 128 structures built, 233 topics researched,
  oil drums picked up) in 619 s, without a script error or a VB error;
  the only assert left is gameOverMessage's view data (the intelligence
  screen's, cluster G).
* `tools/check_debug_lines.py` also knows netplay.h's ASSERT_HOST_ONLY and
  a definition's parameter list continued on the next line.
* The official replay's desync (section 27) came from the same gap:
  multigifts.cpp's technologyGiveAway, a stand-in until now, writes "Not
  adding artefact." to the sync log in a multiplayer game whenever a
  structure is destroyed (the first at 456,202). With it ported, **the
  whole official replay (18,000 ticks) plays without a desync**: every
  game-time CRC it recorded matches (214 s headless; the official build
  29 s). **Cluster B's milestone is reached.**
* statsref follows the new real code (fork T): multigifts.cpp whole,
  message.cpp's lists, transporter.cpp's transporterIsEmpty, wider
  structure.cpp and droid.cpp ranges; both harnesses free the messages
  when they reload a map's objects (a level's start does). The
  regression suite: 3346 of 3346, 90 of 90, the official replay's 120
  ticks.
* Next: the renderer (clusters D and E) towards watching and playing a
  game.

### 29. The renderer base (cluster D): cluster D's milestone

* The user's priority (2026-10-29): clusters D and E, towards watching and
  then playing a game.
* lib/ivis_opengl ported into WzGfx by five forks against a contract
  written first (see CONVENTIONS, "The renderer"): gfx_api.cpp and every
  pipeline_state_helper `using` of gfx_api.h (fork A, with gfx_api_null.cpp),
  gfx_api_gl.cpp whole on a generated GL loader (fork B: `tools/gen_gl.py`
  reads glad.h and gfx_api_gl.cpp's calls; 122 entry points, 32 of them
  opengl32.dll's own exports, the others through DispCallFunc), pieimage,
  tex, bitimage (the image sheets' atlas packing), piepalette and png_util
  on GDI+ (fork I), piestate, piemode, piematrix, pieclip, culling,
  piefunc, pielighting, pieblitfunc and screen (fork C, with glm's matrix
  functions in `src/shared/gfx/glm_mat_vb`), textdraw on GDI (fork T). About
  21,000 lines of VB; the merge compiled after three fixes (a parameter
  named `array`, a variable named `attribute`, a declaration after the first
  procedure).
* Choices: one context, so gl_context and null_context are modules and the
  virtual calls a `Select Case` (a VB class's public method takes no module
  UDT, even in a private class: measured); a pipeline_state_helper
  instantiation is a `pso_type` value; constant buffers and vertex structs
  are `src/shared/gfx` UDTs passed by address; PNG through GDI+ with the
  colour-correction chunks dropped (checked on every kind of PNG in the
  data: the same bytes as libpng); text through GDI (`GetGlyphOutlineW`'s
  greyscale glyphs laid into textdraw.cpp's own composition, fonts from
  `data/fonts` with `AddFontMemResourceEx`); no texture compression (as
  built without ETCPAK and basis); stb_image_resize replaced by a box filter
  (halvings) and bilinear.
* The exe: realmain now calls pie_LoadShaders and screenInitialise;
  wrappers.cpp's loading screen (stars, load callback) is ported; data.cpp's
  image loaders take WzGfx's image sheet handles.
* Checked (`tools/draw_test.py`): `--vbdrawtest` draws boxes, lines, text in
  all eight fonts, a WzText, rotated text and the 200-odd images of the
  frontend image sheet over the title backdrop and saves a screenshot
  through screen.cpp's screenDumpToDisk (glReadPixels, PNG through GDI+):
  **images and text on screen**, the same pixels on every run
  (`tools/draw_test/drawtest.png`, checked by eye). Headless runs work on
  the ported null backend. **Cluster D's milestone is reached.**
* `tools/check_debug_lines.py` knows glad's `#define glX glad_glX`, members
  whose return type has a parameter list, and local structs' inline
  constructors.
* Next: cluster E (display3d, terrain, piedraw, effects...) to watch the AI
  skirmish in 3D.

### 30. The 3D display (cluster E): cluster E's milestone

* Ported by five forks (scratchpad E_DESIGN.md; the call paths mapped first):
  lib/ivis_opengl/piedraw.cpp whole and imdload.cpp's GPU side (the shapes'
  buffers and textures) in WzGfx (fork P); terrain.cpp, texture.cpp,
  lighting.cpp, shadowcascades.cpp, advvis.cpp (fork T); display3d.cpp's
  globals, scene and camera parts, bucket3d.cpp, miscimd.cpp, the rest of
  geometry.cpp (fork S); display3d.cpp's structure, feature and delivery
  point rendering, component.cpp, effects.cpp, atmos.cpp (fork O);
  display.cpp, warcam.cpp, animation.cpp, radar.cpp and display3d.cpp's 2D
  overlays (fork V). display3d.cpp is three VB modules by line range
  (display3d_c, display3d_render_c, display3d_overlay_c), each naming the
  C++ file in its first comment line. About 18,000 lines of VB.
* The merge: the forks' stand-ins and header ports in one new module,
  `stubs_display` (the real ports among them name the module they belong
  in); every stubs_game item a new module defines dropped by a script over
  check_names' report (203 of them); `Static` for the 14 KB instanced
  uniform block (VB allows a procedure 64 KB of locals); VB keywords as
  names (`each`, `output`, `object`); one `\` / `*` precedence bug
  (getRecoil) found by check_intdiv.
* Found by the first windowed runs: multiplay.cpp's cameraToHome and
  keybind.cpp's enableGodMode (the spectator sees the whole map) were still
  stand-ins; structure.cpp's buildBlueprint (a spectator sees every player's
  planned buildings) needs a STRUCTURE outside the lists: the structure
  pool's scratch slots (`StructContainer_scratch`, never alive, forgotten
  at clear()).
* Tests: the vbtest runs turn effects and the display's entry points off
  (`vbEffectsOff`, `vbDisplayOff`: statsref's stand-ins do nothing; the
  display's edge scrolling with the mouse at 0, 0 logged warcam lines); statsref compiles lighting.cpp's
  calcDroidIllumination (droidUpdate now calls the real one).
* Checked (`tools/watch_test.py`): Nexus, NullBot, BoneCrusher and Cobra on
  Bloat in a window, the camera on an AI's base: terrain (the Normal
  quality's splatting, decals, water, cliffs), structures with modules and
  team colours, droids, features, oil, effects and shadows.
  `--vbwatchshot` takes the game's own screenshot at a game time. **Cluster
  E's milestone is reached**: the AI skirmish can be watched in 3D, at a
  steady 58-60 fps (vsync) early in the game at real speed; the frame
  rate falls to about 25 fps at 4 game minutes with the autogame's game
  tick in every frame (to measure and speed up).
* Open: the real ports in stubs_display move to their modules as those
  come; the camera's pitch from getBestPitchToEdgeOfGrid is 0 until
  raycast_c gets its tile height callback; IntImages (rank and group
  images) waits for hci (cluster G); the terrain quality is Normal (the C++'s
  32-bit default; the official 64-bit build uses High).
* Next: cluster F (widgets and the front end) and G (the in-game
  interface), towards playing a skirmish from the menu; the frame rate.
  (Done: section 31.)

### 31. The widgets, the front end and the in-game interface (clusters F and G): their milestones

* The contract first (scratchpad F_DESIGN.md, G_DESIGN.md, G_MAP.md; the
  CONVENTIONS section "Widgets"): one VB class `WIDGET` for the 334 C++
  widget classes, the C++ class in `kind`, each class's own fields in a data
  class, methods as module procedures, virtuals through generated
  dispatchers, lambdas as `WzFunc` objects; `tools/gen_widgets.py` writes
  the generated half (`widget_classes_vb.bas`: 134 classes, 124 with data,
  70 virtuals, 587 lambda ids).
* Ported by twelve forks: lib/widget whole (W1: widget, form, button, label,
  bar, tip, checkbox, the screen; W2: slider, edit box, scroll bar,
  scrollable list, list, grid layout, alignment, clip rect, multibutton
  form, drop-down, popover and its menu, paragraph, table); frontend.cpp
  and the title screens (F1), the options screens (F3), the skirmish lobby
  (multiint's title UI and multimenu, F2) and its widgets (F4); hci.cpp,
  intdisplay.cpp and the radar widget (G1), the object and stats forms
  (build, manufacture, research, commander, groups) and edit3d.cpp (G2),
  design.cpp and template.cpp's rest (G3), the input manager, key bindings
  and selection (G4), console.cpp, intorder.cpp, ingameop.cpp and mission.cpp's
  interface parts (G5), intelmap.cpp and the in-game options screen (G6).
  About 59,000 lines of VB (80 modules and 124 data classes); the exe has
  366 modules and 161,000 lines.
* The merge: the forks' stand-ins in one new module, `stubs_ui`; VB's
  limits: class names of 31 characters at most (three data classes got
  shorter names: gen_widgets maps a class to its data class's file), names
  that VB keywords or other names take (`SINGLE`, `ALLIANCES_`).
* Found by the first runs (`tools/ui_test.py`: scripts of clicks and keys
  posted to the window, the game's own screenshots, the widget tree in the
  log):
  * passing `Me` straight to a ParamArray evaluates its default member
    (run-time error 438, from WIDGET's Class_Terminate: VB's error box,
    the window hung; found with cdb on a CodeView build, measured on a
    probe): a local copy (CONVENTIONS);
  * Class_Terminate logs its errors (none may leave it);
  * the UI script's hook was in no loop: mainLoop runs it;
  * still stand-ins: netplay.cpp's NETclose and multiopt.cpp's
    multiGameShutdown (the lobby's host button was gone after a first
    game), multiplay.cpp's getPlayerName (the lobby's rows had no names),
    droid.cpp's getDroidLevelName (the unit's rank in its info line);
  * a log call's argument with a side effect (getPlayerName asserts past
    MAX_PLAYERS) runs only with its log part on, as the C++ macro does.
* Not bugs: the blueprint is red almost everywhere on Rush (the map's rocky
  tile 71 is a cliff in both the map's and 4.7's Arizona terrain types,
  factories need a free tile around them, and the delivery points block);
  the design screen's centre shows only after a template is chosen, as in
  C++; the blank design has the null body's zero stats.
* Tests: the vbtest runs skip the interface's entry points that statsref
  has as stand-ins (`vbDisplayOff`: intRunWidgets, widgRunScreen,
  widgGetFromID, intDisplayWidgets, intRefreshScreen and hci's other
  notifications, updateConsoleMessages, console, selDroidDeselect,
  getPlayerName, kickPlayer, the in-game options; intAddProximityButton
  keeps the button's ID): with the real hci, the headless frames asserted
  on the missing widget screen. `tools\statsref_stubs.py` lists
  statsref's stand-ins whose VB port is real and ungated. The suite:
  3346 of 3346, 90 of 90, the official replay's 120 of 120 ticks; the
  `--vbtest` run takes about 500 s (335 s at cluster B's end; not timed
  in D and E: to look at).
* Checked (`tools/ui_tests`): the main menu; options: Show FPS set through
  its drop-down is saved and shown again by the next run (options.txt,
  then options_check.txt with `--config-of=options`). **Cluster F's
  milestone is reached.** From the main menu, Single Player, Start
  Skirmish Game, Start Hosting, Ready: the skirmish against three Nexus
  AIs starts (skirmish.txt); research started from the research form, a
  truck produced from the manufacture form, the design screen with a
  template's stats, a structure placed from the build form and built by a
  truck (build.txt), a truck selected and moved (orders.txt), the in-game
  menu (paused), Quit to the main menu and a second skirmish (quit.txt).
  **Cluster G's milestone is reached: a skirmish against the AIs is played
  from the menu.**
* Open: the lobby's map preview is blank; the keymap.json round trip;
  the intelligence screen with messages; load and save (cluster I); the
  chat and the network lobby (cluster J); the sound (cluster H); the
  frame rate (about 30 fps in a debug build; to measure in the release
  one).
* Next: cluster H (sound) or I (campaign, save and load). (H done: section 32.)

### 32. The sound (cluster H): cluster H's milestone

* The contract first (scratchpad H_DESIGN.md): TRACK, AUDIO_SAMPLE, AUDIO_STREAM and IStreamFinished (the
  std::function onFinished) as WzSound classes; the decoder's API (stb_vorbis's, on memory) and
  openal_track.cpp's (tracklib.h) fixed as VB signatures; three forks:
  * H1, the decoder: stb_vorbis v1.22 (ref/stb_vorbis, public domain; 4.7.0 uses libvorbis) ported whole
    for the pull API on memory (`stb_vorbis_c`, 3,600 lines; a handle is a pool index + 1).
  * H2, the mixer: `al_win`, the OpenAL subset openal_track.cpp calls (sources, buffers, queues, the
    listener; AL_NONE only: WZ computes every gain itself), mixed in VB at 44100 Hz stereo over waveOut,
    polled from sound_Update (QuakeVB's snd_win), linear resampling with the pitch, OpenAL Soft's
    first-order stereo panning (the direction's component along the listener's right), no HRTF
    (Unsupported, as without ALC_SOFT_HRTF); openal_track.cpp whole, openal_error.cpp, what openal_info.cpp
    reports, oggvorbis.cpp's WZVorbisDecoder over stb_vorbis (Opus is not ported).
  * H3: audio.cpp, track.cpp, cdaudio.cpp and playlist.cpp whole (no music ships: the playlist finds no
    albums, as C++ does without data/music); SoundApi's whole surface; `ISoundGame` (aud.h, the game and
    real time, the AUDIO_CALLBACKs by number), implemented by the exe's `SoundGame` (aud.cpp in `aud_c`;
    droid.cpp's droidAudioTrackStopped); every sound stand-in of the exe gone (an exe stand-in would
    silently shadow SoundApi's name); `--vbsndcapture=<file.wav>` (the mix to a WAV file, no device).
  WzSound is 10,700 lines.
* The merge built at the first try. The test tools run silent (`--nosound`; ui_test's `--sound` turns it
  on).
* Tests (`run_tests.py --sound`, in the full run too):
  * the decoder against ffmpeg's own Vorbis decoder (the auto mode classifier refused to compile and run the
    downloaded stb_vorbis.c as a reference): all 560 game sounds (548 effects, 12 briefing voices) decode
    with every sample within 1 of ffmpeg's (19 files identical), 559 to the last granule position
    (libvorbis's ov_pcm_total, what 4.7.0's decoder reports); shield-hit.ogg is 63 samples short (a stream
    with samples to trim at its start: stb_vorbis's own inference of the last frame); about 5 s for all of
    them in a test build; 1.3 s with VB's fast-code optimisation and no bounds checks, as WzSound is
    built since (the user's choice, 2026-10-29; build.py's OPTIMISED; the other projects keep no
    optimisation);
  * the mixer (`tools/mixer_test.py`, MixerTests.exe on al_win alone, synthetic sources in capture mode):
    66 of 66 (gain, panning around a turned listener, resampling 11025 to 44100 Hz, pitch, looping, a
    queued stream, pause, resume, stop, clipping, a stereo buffer unpanned, the error codes, the real
    device cycling with a silent source); 1 source costs about 10 ms per second of audio, 32 about 76 ms.
* Checked in the game (capture): the menu's clicks, a skirmish's sounds; a truck sent to the screen's right
  then to its left: its engine louder on the right channel, then on the left; the options' Audio page
  (the volumes, HRTF Unsupported). The game's load with the sound on is about 2 s longer (the 367 effects
  of the skirmish's resources decoded when they load, as C++ does). **Cluster H's milestone is reached:
  the sound effects play in game.**
* Open: the music (no files ship; Opus is not ported); the briefings' voice streams (campaign, cluster I);
  the latency of the polled ring (about 93 ms, nothing plays while a load runs no frames); a VB run-time
  error's exit ends in an access violation after the error is logged (with the sound off too).
* Next: cluster I (campaign, save and load) or J (network). (I done: section 33.)

### 33. The campaign, save and load (cluster I): cluster I's milestone

* The contract (scratchpad I_DESIGN.md): the stand-ins' VB signatures are the contract between the forks
  (their callers use them); each fork ports its C++ files whole. Four forks:
  * I1, game.cpp's main save file: the SAVE_GAME versions and their (de)serialisers, saveGame, loadGame's
    user-save and mission branches, gameLoad/gameLoadV, the main and game files (`game_main_c`, `game_c`;
    map.cpp's save parts in `map_save_c`);
  * I2, game.cpp's object files: droids, structures, features, templates, terrain types, component and
    structure type lists, research, messages, limits, fire support, the script state, the guide topics
    (`game_objects_c`);
  * I3, mission.cpp's rest (off-world missions: the home base and limbo droids, the end of a mission and its
    result screen, the timers, the reinforcements) and transporter.cpp whole (`mission_c`, `transporter_c`);
  * I4, the campaign's front end: titleui/campaign.cpp (the campaign selector), campaigninfo.cpp,
    seqdisp.cpp (the briefings' text; no videos), scores.cpp, loadsave.cpp (the load and save screens,
    SaveGamePath_t), challenge.cpp, message.cpp's rest, cheat.cpp, screens/chatscreen.cpp (the in-game chat:
    the campaign's test types its cheats there); quick chat stays a stand-in.
  About 18,000 lines of VB; activity.cpp's game mode (`activity_c`: the music's mode, the end screen and the
  challenges read it; its event sinks stay out).
* The merge: 183 stand-ins dropped by check_names' report, the forks' new ones in `stubs_campaign`; the exe
  built at the first try (the exe now holds about 24,000 distinct names of VB6's 32,100).
* Found by the first runs:
  * a fixed-size array passed to a procedure that ReDims it (VB error 10 in the save's header):
    `tools/check_redim.py` finds them;
  * a Collection item that is an object read with a Let (`v = c.Item(k)`): VB asks the object for its
    default member (error 438), so the chat screen's registry and the campaign list's key test never found
    their entries: `Set`, or `IsObject(c.Item(k))`;
  * legacy Chakra lacks ES2022's `Object.hasOwn` (libcampaign calls it): the backend's built-ins script
    defines it; JsRT's JsGetProperty refuses a primitive where QuickJS boxes it (a script reading `.id` of a
    number): `jsBoxPrimitive`;
  * the UI scripts' keys carried no scan code (WzPlatform reads it, as SDL does): `key` sends it, the press
    and the release a frame apart; the script's hook ran in mainLoop and again in the title and game
    loops' test hooks (its waits counted twice a frame): mainLoop's only; and a wait is N frames and
    at least N / 60 s (a covered window runs hundreds of frames a second);
  * the vbtest runs' getViewData: statsref's stand-in (no view data loaded).
* Not bugs: a GTYPE_SAVE_START save (`--saveandquit`) restarts its level when loaded, so the save written
  after its load differs from it as 4.7.0's does (I1 traced each difference to the C++: the level's objects
  renumbered from the save's objId x 4, the start-level script's objects and labels not made again, power
  cleared, the save's game settings over main.json's, the gateways listed twice); a mid-mission off-world
  launch with nobody aboard (the cheat on a "between" level) loses the mission, as in 4.7.0.
* Checked:
  * `tools/save_test.py`: CAM_1A saved at its start, loaded and saved again: 23 of 23 files the same but
    for the differences 4.7.0 makes too (each named with its C++ line); a probe inside levLoadData (then
    removed) showed the save's 47 droids, 36 features, 28 labels and 9 effects read back exactly with their
    ids;
  * a mid-mission save from the in-game menu (save_menu.txt) loaded from the main menu and saved again at once
    (save_load.txt): the 47 droids, 56 structures and 36 features the same, ids and fields, but for two
    adjoining scavenger structures' height (1 unit: the flattening of their shared tiles in another order,
    not yet checked against the C++); the script state (globals, groups, timers) saved and restored;
  * `tools/ui_tests/campaign.txt`: New Campaign, the selector, Start Game: CAM_1A plays; and
    `campaign_walk.txt`: **campaign 1 from the main menu to CAM_1END**, every level won with the campaign
    scripts' chat cheat ("cheat on", "let me win") and the result screen's Continue Game, and each
    "between" level's transporter launched as a player does (the units board it: the script's `embark`
    command, then the launch button): CAM_1A, CAM_1B, SUB_1_1 to SUB_1_3 off world, CAM_1C, CAM_1CA,
    SUB_1_4A, SUB_1_5, CAM_1A-C, SUB_1_7, SUB_1_D, CAM_1END; no error or assert in the log. **Cluster I's
    milestone is reached.**
* Open: the videos (lib/sequence) and the music; quick chat; campaign mods (modinfo.cpp); the two
  structures' height after a load; the in-game save screen's black background (to check against 4.7.0);
  campaigns 2 and 3 not walked yet.
* Next: cluster J (network) or K (optional). (K: section 34.)

### 34. Archives, mods, music, videos, quick chat (cluster K) and the campaign's open items

* The user's request (2026-10-30): everything section 33 left open. Two choices were the user's: the music
  ported as Opus ("port libopus", not converted to Vorbis) and the videos downloaded and ported (libtheora).
* The data (`tools/setup_data.py`, hard links into ref/warzone2100/data, where 4.7.0's source tree has them):
  the official 4.7.0 build's music (data/music: menu.opus and 4 albums, 110 MB of Opus) and campaign mods
  (data/mods/campaign/*.wz), and the videos (data/sequences.wz: wz-sequences release v3, the one 4.7.0's
  data/CMakeLists.txt downloads, its low-quality English set, 170 MB, in ref/videos). The sources, at the
  versions of 4.7.0's vcpkg baseline (their SHA-512s as vcpkg's): ref/physfs (PhysFS 3.2.0), ref/ogg (libogg
  1.3.6), ref/theora (libtheora 1.2.0), ref/opus (opus 1.5.2); opusfile is 4.7.0's submodule (9d718345).
* The contract first (scratchpad K_DESIGN.md: the forks' files, the VB signatures between them, ogg_h.bas's
  types written first); seven forks, about 42,000 lines of VB:
  * Z: PhysFS 3.2.0's archiver dispatch with its DIR and ZIP archivers and miniz's inflate in WzFramework
    (physfs_win rewritten, physfs_zip_c, physfs_miniz_c: 5,400 lines); modinfo.cpp (`modinfo_c`);
  * O1: libogg's framing.c (`src/shared/ogg`), opusfile whole but HTTP (float, OP_HAVE_LRINTF), oggopus.cpp
    (`WZOpusDecoder`), codecs.h's `WZDecoder` interface (6,500 lines);
  * O2: libopus's CELT decoder and range decoder (5,400 lines); O3: its SILK decoder, opus.c, opus_decoder.c
    and the multistream decoder (6,300 lines);
  * V1: libtheora's decoder with the legacy theora.h API, the post-processing included (6,200 lines);
  * V2: lib/sequence (sequence.cpp, timer.cpp) in a new DLL, **WzSequence** (optimised like WzSound;
    references WzFramework, WzGfx, WzSound), libvorbis's synthesis API over stb_vorbis for the videos' audio,
    GfxApi's GFX handles, al_win's alGetBufferi;
  * Q: quickchat.cpp and chatoptions.cpp whole, lib/widget's paneltabbutton.cpp (4,900 lines; gen_widgets:
    170 classes).
* The merge built at the first try (the exe: 427 modules, 185,000 lines; its LINK.EXE command line at 31,097
  of 32,767 characters: the next clusters build from a shorter path, `subst`). Found by the forks' tests and
  the merge: opus_packet_get_samples_per_frame read SILK's shift table before any decoder made it (opusfile
  asks first: no music); ffmpeg 8.1 links libopus 1.6.1 (other packet-loss concealment): the Opus reference
  is the official build's libopus.dll through ctypes, as libtheoradec.dll is for the post-processed frames;
  ffmpeg's frame-threaded Theora decode is not deterministic (`-threads 1`); the full-screen videos run
  outside the loops that poll al_win, so their audio never played nor ended (seq_Update polls it:
  `al_vbProcessCurrent`); WZOpusDecoder::decode writes past its 16 KiB buffer with 60 ms packets (the port
  grows the buffer; the game's tracks use 20 ms); 4.7.0's oggopus.cpp SEEK_END is off by one byte, so each
  track's last Ogg page never plays (kept). Not bugs: CELT's pitch-based concealment differs from 1.5.2 by
  float rounding the LPC filter amplifies (about 5 units in 16 bits).
* Tests (`run_tests.py --media`, in the full run too; each tools/<name>.py's run(keep)):
  * zip_test: every entry of the four mods, sequences.wz, mp.wz (and base.wz with `--base`) and synthetic
    archives (deflate block types, zip64, data before the archive, symlinks) read through the port's PhysFS
    equals Python zipfile's, whole and in chunks with seeks; 98 single PhysFS calls: 109 of 109; inflate
    about 25 MB/s (WzFramework unoptimised);
  * celt_test (src/celttests): 27 of 27 CELT-only streams (mono, stereo, 2.5 to 60 ms, 8 to 256 kbps, lost
    packets) within 1.2e-6 of libopus 1.5.2 (lost: 1.6e-4); silk_test (src/silktests): 50 of 50 (SILK bit
    exact, hybrid 3e-8, mode switches 7.8e-7, 5.1 1.8e-7); opus_test: 24 of 24 files through opusfile (the
    official music, synthetic streams, a dropped page) as libopus 1.5.2 decodes them, op_read_stereo's
    dithering bit for bit; the music decodes at about 160x real time (WzSound optimised);
  * theora_test: all 179 videos (64,256 frames) bit exact, at post-processing level 0 against ffmpeg and at
    level 7 (sequence.cpp's) against the official libtheoradec.dll; 391 frames/s at level 7 (2.6 ms a frame);
    the default run a subset of 25 (`--all`);
  * video_audio_test: 168 of 168 videos' Vorbis streams through the packet API within 1 of ffmpeg in 16 bits.
  The full suite after the merge: 3362 of 3362, 90 of 90 (the PhysFS rewrite changed no output), the official
  replay's 120 of 120 ticks, 560 of 560 sounds, the mixer's 66 of 66 (its build script was stale: fixed).
* Checked in the game (tools/ui_tests): the campaign's opening video plays full screen with its titles and
  subtitles (campaign.txt); the menu music plays (captured: correlation 0.9991 with ffmpeg's decode of
  menu.opus, at the music volume's 0.5); the campaign selector lists the four mods, Reclamation shows its
  banner and starts its level L1 from its .wz (campaign_mods.txt); quick chat in the lobby and in the game,
  by mouse and by keys (quickchat.txt, `--debug=console`); the mid-mission save and its load (save_menu.txt,
  save_load.txt: the 47 droids, 56 structures, 36 features as before) and the start save's round trip
  (save_test.py: 23 of 23).
* Campaign walks (`tools/make_campaign_walk.py [1|2|3]`, campaign2_walk.txt and campaign3_walk.txt new): the
  UI scripts got `novideo`, `quiet <frames>` and `waitlevel <name>` (a key or click during a full-screen video
  goes to the video; a level's briefing videos start seconds in, SUB_1_1S's when the cheat's transporter
  lands), and CAM_3_4S a minute to board (its landing zone is at the far end of CAM3A-D2's map); Continue Game
  waits out a video the win starts (CAM_2C's cheat destroys its air base, which plays MB2_C_MSG2), and CAM_3A
  plays 40 s before its win (campaign 3 from the menu flies its units in: won before they land, SUB_3_1S has
  no truck nor factory and libcampaign's __camPlayerDead loses it at once, as 4.7.0's would). **Campaign 1
  with its videos: CAM_1A to CAM_1END, 50 videos (18,787 frames, 76 dropped), no error or assert (1,500 s).**
  **Campaign 2: CAM_2A to CAM_2END (whose win loads CAM_3A), 47 videos (12,088 frames, 29 dropped), no error
  or assert (1,431 s); campaign 3: CAM_3A to CAM_3_4's win and the outro (outro.ogg), 42 videos (12,191
  frames, 86 dropped, 38 of them in cam3bg.ogg), no error or assert (1,097 s).**
* Section 33's open items:
  * the in-game save screen's black background is 4.7.0's: addLoadSave sets setAllPauseStates(true), renderLoop
    draws the world only while the game update runs (loop.cpp ~322), and gfx_api_gl's endRenderPass clears
    after each swap; the Esc menu pauses with kf_TogglePauseMode, which leaves the game update on;
  * the two structures' height after a load is 4.7.0's: loadSaveStructure2 rebuilds the structures in the
    save's order and alignStructure recomputes their height ((min + max) / 2 of the footprint's vertices, then
    flattened); a simulation of exactly that from the first save (scratchpad flatten_sim.py) gives the second
    save's tiles: 0 differ;
  * the access violation after a VB run-time error: ExitProcess runs the DLLs' detach code in reverse load
    order, so AppXDeploymentClient (loaded by the NVIDIA GL driver) is torn down before MSVBVM60, whose
    detach calls CoFreeUnusedLibraries, and combase calls the dead DLL's DllCanUnloadNow (cdb); exemain_vb's
    error and non-zero exits and src/shared's exit_ use TerminateProcess: exit code 1 now;
  * the asserts qtscript.cpp:1464 (scriptsReady) a cheat-won CAM_3_4S logged: setUpMission shuts the scripts
    down and the next frame's gameStateUpdate runs before renderLoop returns GAMECODE_NEWLEVEL, as in 4.7.0
    (code read, not run on 4.7.0); the walk with videos, its timing other, logged none.
* Open: the music's in-game album changes (only the menu's checked); the videos' scanline modes and the
  intelligence screen's small video (sequence.cpp's 240x180 path) seen once only; the chat options form (a
  hosted network game: cluster J); team strategy (J); the lookups over folders are about 1.8 times slower
  with PhysFS's per-element checks (startup not measured yet); translations.
* Next: cluster J (network). (First: section 35.)

### 35. Cluster K's open items: the Music Manager, the videos' display, file lookups, the build limit

* The user's request (2026-10-31): section 34's open music, video and lookup items, and what LINK.EXE's
  command-line limit means for the next clusters.
* The Music Manager (fork M; scratchpad k2_M_report.md): musicmanager.cpp and titleui/musicmanagertitleui.cpp
  whole (`musicmanager_c`, `musicmanagertitleui_c`, `WzMusicManagerTitleUI`, `MusicManager_CDAudioEventSink`
  implementing WzSound's ICDAudioEventSink, 4 widget classes: about 1,100 lines); SoundApi's
  PlayList_GetCurrentMusicMode and sound_srand. Found:
  * the first in-game track was the same in every run: each project has its own rand() state and nothing
    seeded WzSound's (PlayList_RandomizeCurrentSong is its one draw): main_c seeds it with main.cpp's time
    seed (`sound_srand`);
  * table.cpp:689's assert on every Music Manager open: relayoutColumns runs at the new table's 1-pixel width,
    where 4.7.0's 32-bit size_t wraps (1 - 135 = 4294967162) and takes the expand branch; table_c compared
    signed. Now unsigned (ULT, udiv32): the columns 215/164 as 4.7.0's (chatoptions' table is made the same way).
  Checked (tools/ui_tests music_manager.txt, music_manager_check.txt, music_ingame.txt; --sound --debug=sound):
  the 39 tracks of the 4 albums with the mode columns; a mode unticked, saved to music.json and still unticked
  after a restart; a track played from the manager is what is heard (correlation 0.9998 to 1.0000 with
  ffmpeg's decode, gain 0.499); in a skirmish only the skirmish column enabled, the first track a random one of
  them; **the OST's last track (track3, 299 s) ended and the next album's first started** ("Now playing
  music/albums/aftermath_soundtrack/menu_enhanced.opus (was playing .../original_soundtrack/track3.opus)"),
  the open form following it.
* The videos' display (`tools/video_display_test.py`, new): a model of sequence.cpp's display (the official
  libtheoradec.dll's frames at the maximum post-processing level, video_write's integer YUV -> RGB with its
  scanline rows, the texture's top left of a black 1024x1024, VideoPSO's bilinear quad) against screenshots of
  `--game=CAM_2A`'s opening video (cam2/c002.ogg) in the 9 pairs of video size (1x, 2x, full screen) and
  scanlines (off, 50%, black), and of the intelligence screen's research film (a topic finished with "work
  harder", its message opened: res_weapons.ogg). The port: 1x and the intelligence screen exactly the model;
  2x and full screen within 0.6 of a level (of 255) at the 99th percentile with scanlines off, 1.8 with them
  (the GPU's bilinear weights where the rows alternate); scanlines off at 1x and in the intelligence screen as
  sequence.cpp's rule says. **The official 4.7.0 build, captured from outside (PrintWindow) with the same
  configs: the same model, the same numbers.** (Its intelligence screen is not run: SDL3 takes a click's
  position from the real cursor, and posted messages do not move it.) Frame counts the same in both builds
  (c002 1064, cam2ap 177, cam2acol 201, res_weapons 157; none dropped). Learnt on the way (the model's, not
  the port's): the legacy theora.h API's ti.frame_width x frame_height is the picture size (192x168 in a
  192x176 frame), read from the frame's top left; a research film ends on an empty view (seq_Shutdown releases
  the texture; VIDEO_HOLD_LAST_FRAME only stops a restart), 4.7.0's code; a screenshot every 15 frames makes
  the film drop frames (34 of 157: the PNG writes; none without them).
* File lookups. An instrumented WzFramework (scratchpad prof_physfs.py: each PhysFS function timed; the
  top-level time is the game's time in PhysFS) under the frozen exe: to the main menu 362 ms of 3.2 s in
  PhysFS (739 files opened, 5,804 GetFileAttributesExW), a campaign start 2.2 s, a skirmish start 3.9 s
  (73,733 stat calls, 6,584 PHYSFS_exists); PhysFS's own VB code about 5% of it, the rest the system calls.
  **The cause was the data/ junction**, not the port: Windows re-parses a junction on every file call
  (GetFileAttributesExW 53 us through it, 30 us on the folder itself), and PhysFS's per-element checks (4.7.0
  forbids symbolic links: main.cpp's PHYSFS_permitSymbolicLinks(0)) make several calls per lookup. `data/` is
  now a mirror of ref/warzone2100/data: real folders, its 4,584 files hard links (`tools/setup_data.py`;
  build.py makes it when data/ is missing or still a junction). Medians of 3, the final exe (the old layout
  reproduced with a junction in build/; 4.7.0's release layout, the official base.wz and mp.wz, beside):

  | | mirror (now) | junction (before) | .wz (release layout) |
  |---|---|---|---|
  | to the main menu | 3.38 s | 3.51 s | 3.61 s |
  | a campaign start (CAM_1A, its video skipped) | 25.3 s | 25.7 s | 23.8 s |
  | a skirmish start (4c-rush) | 30.8 s | 33.2 s | 30.9 s |
  | in PhysFS (menu / campaign / skirmish) | 0.26 / 1.49 / 2.49 s | 0.36 / 2.18 / 3.90 s | |

  The rest of a level start is the game's own code (not profiled yet); the release layout's in-memory lookups
  are paid back by inflating at 25 MB/s (WzFramework unoptimised).
* The build limit (measured: tools/linkwatch.ps1 records LINK.EXE's command lines): VB6 compiles each
  module to bin\<module>.OBJ and passes all of them, quoted, to LINK.EXE on one command line, which
  CreateProcess caps at 32,767 characters; past it VB6 says "Could not execute LINK.EXE" (feasibility probe 1).
  The exe after the Music Manager: **30,145 characters, 456 .OBJ files** (65 characters each, 44 of them
  `C:\Users\Owner\desktop\websites\warzone\bin\`): 2,622 left, about 40 modules. build.py's estimate counted
  the source folders' paths (31,704); it now counts bin\'s (30,145, exact) and warns past 30,000. The exe's
  other ceilings: 456 of about 850 modules, about 25,200 of about 32,100 names (the distinct identifiers of
  its sources: an estimate). The cure for the command line is a shorter output folder (the .OBJ files go to
  /outdir), e.g. build.py building into a `subst` drive mapped to bin\ (`W:\<module>.OBJ`: about 28 characters
  a module, room past the 850-module ceiling); the 850 modules and the names have no such cure (a project
  split).
* Tests: the static checks 0 problems; the full suite over the data mirror: 3362 of 3362, 90 of 90, the official replay's 120 of 120 ticks, 560 of 560 sounds, the mixer's 66 of 66, --media (zip 109 of 109, CELT 27 of 27, SILK/Opus 50 of 50, opusfile 24 of 24, Theora 25 of 25 at both levels, video audio 168 of 168); video_display_test as above.
* Open: a level start's time outside PhysFS (about 28 s of a skirmish start, not profiled); translations; the
  chat options form and team strategy (cluster J). (LINK.EXE's command line: section 36.)
* Next: cluster J (network). (First: section 36.)

### 36. The build through a short drive; cluster J's scope (2026-10-31)

* The user's decisions on section 35's build limit and cluster J: build the exe from a short path (option 1);
  keep J to what its milestone needs and leave the optional network features as stand-ins (option 3); J's
  plumbing in WzNet, only the screens that use the widget library in the exe (option 4); a new DLL subsystem
  to ease the per-project limits later. The goal behind it: **the build and the game on other people's
  computers**.
* `build.py` builds every exe (Warzone.exe and the test exes) through a drive letter that `subst` maps to bin\
  for the build (the first free one from W down; no administrator rights; removed afterwards, a mapping an
  interrupted build left is reused). VB6 writes the .OBJ files to that folder and passes `"W:\<module>.OBJ"`
  to LINK.EXE: **the exe's command line went from 30,145 to 11,408 characters** (456 modules, measured with
  tools/linkwatch.ps1, now a WMI creation event: its polling loop missed short links), about 25 characters a
  module, so the 850-module ceiling now comes first; and the line no longer depends on where the tree is (a
  deeper checkout on another computer would have been over the limit already). The DLLs still build into
  bin\ itself (VB6 registers each DLL it builds at its output path, which must stay valid; the largest DLL's
  line is 7,644 characters). build.py's estimate (`link_estimate`) is the measured formula and is printed past
  30,000. VB6 is found through %VB6_DIR%, else its setup's registry entry (ProductDir), else its default
  folder. Checked: two full builds, the main menu test, the mixer test's exe (66 of 66), the leftover-mapping
  reuse; no mapping left behind.
* Cluster J's scope (the estimate: scratchpad j_budget.py, the exe's names and modules against the C++ files):
  * WzNet: the rest of lib/netplay for direct-IP play: tcp/ (netsocket, the TCP connection provider, client
    connection, listen socket, poll group, address resolver), client_connection, connection_provider_registry,
    wz_connection_provider, pending_writes_manager(_map), polling_util, error_categories, ip_helpers,
    netpermissions, netjoin, zlib_compression_adapter and wz_compression_provider (needs a deflate: miniz's,
    beside WzFramework's inflate), the connection parts of netplay.cpp still standing in (about 3,500 lines of
    C++ besides netplay.cpp's);
  * the exe: multisync.cpp, multivote.cpp, multijoin_helpers.cpp, screens/joiningscreen.cpp,
    screens/netpregamescreen.cpp, hci/teamstrategy.cpp (about 7,000 lines; about 27 modules with their
    widget data classes, about 860 new names: the exe to about 482 modules and 26,000 to 27,000 names), and
    the network stand-ins in multiint, multiplay and stubs_*;
  * stand-ins, as Vulkan and Discord are: the internet lobby (netlobby.cpp, titleui/gamebrowser.cpp,
    widgets/gamebrowserform.cpp), the blind waiting room, the spectator widgets, the lobby's slash commands
    (multilobbycommands.cpp), the command interface (stdinreader.cpp), port mapping (UPnP: miniupnpc,
    libplum), GameNetworkingSockets. About 35 exe modules and 1,400 names not spent; they can come later.
* A DLL split, later: with the command line out of the way, the exe's binding limits are about 850 modules
  (456 now, about 482 after J) and about 32,100 names (about 25,200 now, 26,000 to 27,000 after J). Before
  either gets near, a self-contained part of the exe moves into a new DLL; the DLL's names the exe calls
  still count in the exe's table (feasibility probe 4), so the candidate must have a narrow interface, and
  its widgets come along with lib/widget or not at all. To be studied with a measurement of each candidate's
  interface (the exe's calls into it and its calls back), not before cluster J.
* Other computers (to do when the game is shared): bin\ runs without installing (the exe's reg-free COM
  manifest; MSVBVM60.DLL ships with Windows; OpenGL 3.3); data\ is this tree's hard-link mirror and must be
  copied as files (or 4.7.0's base.wz and mp.wz used); no path of this computer is compiled into the game
  (checked: the sources name none). **`tools/package.py`** makes the shareable copy: dist\WarzoneVB-<date>\
  with bin\ (the exe and the seven DLLs only), data\ as real files (`--wz`: 4.7.0's base.wz and mp.wz),
  Play.cmd (the game with `--configdir=<the folder>\userdata`, made first: the game stops at once when that
  folder is missing, as 4.7.0 does; without it the port shares %APPDATA%\Warzone 2100 Project with an
  installed official 4.7.0), README.txt, 4.7.0's licences and AUTHORS, and source\WarzoneVB-source.zip (src,
  tools, docs: the GPL's corresponding source); `--zip` zips it. Checked: 560 MB, the packaged exe reaches
  the main menu, Play.cmd from a folder with spaces. Found on the way: data\'s mirror lacked 4.7.0's five
  empty folders (terrain_overrides\classic and \high, two campaign mod folders): setup_data.py makes them.
* Next: cluster J (network), in that scope. (Done: section 37.)

### 37. The network (cluster J): cluster J's milestone

* The user's request (2026-10-31): "lets do it" (cluster J in section 36's scope). One addition, told to the
  user: the game browser form (titleui/gamebrowser.cpp, widgets/gamebrowserform.cpp), whose "GameID or IP:port"
  box is 4.7.0's only way to join by IP from the menus; its games list shows 4.7.0's "Failed to connect to lobby
  server" (the lobby server is a stand-in).
* Found before the split: 4.7.0's network signs and encrypts with libsodium (the joining handshake's identity
  challenge: ed25519; the secured messages: X25519 session keys and XChaCha20-Poly1305), compresses its streams
  with zlib, and writes from a background thread; the joining screen (the exe) opens its connection itself, so
  the connection classes are WzNet's public classes.
* The contract (scratchpad J_DESIGN.md: zlib streams, crypto, the connection classes, NetApi and INetGame) and
  **private builds**: each fork staged its files in its own folder and built private copies (projects from src\
  plus its stage, private type libraries, the exe through a subst drive) with scratchpad j_build.py (now
  `tools/private_build.py`), so no fork saw another's half-done work. Six forks, about 37,000 lines of VB:
  * JZ: miniz v1.11's deflate (tdefl) and its zlib-style streams beside the inflate, in **src/shared/miniz**
    (WzFramework, WzNet; physfs_miniz_c moved there): 5,200 lines with its tests;
  * JS: crc.cpp's sha256Sum, EcKey, SessionKeys, base64 and genSecRandomBytes over SHA-256/512, BLAKE2b,
    ChaCha20/HChaCha20, Poly1305, XChaCha20-Poly1305, X25519 and ed25519 (TweetNaCl-style, 16-bit limbs in
    Doubles: products bounded at 2^44.5, measured 2^43.1) with libsodium 1.0.20's acceptance rules, crypto_kx,
    RtlGenRandom, in **src/shared/sodium** (the exe, WzNet): 2,800 lines;
  * JT: lib/netplay's connection layer over Winsock (the client connection with its zlib compression, the
    pending writes, the provider registry, the error codes, select descriptor sets, tcp/*): 4,000 lines;
  * JN: the rest of netplay.cpp (hosting, NETallowJoining, the connections, system messages, player slots and
    spectators, file transfer), netjoin, netpermissions, wzfile, netlog; nettypes' session keys and secured
    messages; the NETPLAY struct for real (stubs_net gone); INetGame whole (62 methods) and NetGame.cls; NetApi
    +117 members; netlobby and port mapping as unreachable: 9,900 lines;
  * JJ: joiningscreen.cpp, gamebrowser.cpp, gamebrowserform.cpp, multijoin_helpers.cpp: 6,000 lines;
  * JL: netpregamescreen.cpp, multisync.cpp, multivote.cpp, teamstrategy.cpp, multistat.cpp rewritten
    (identities, session keys, stats messages), the network parts of multiplay, multijoin and multiopt: 9,400
    lines; 75 stand-ins replaced.
* The forks' rule for threads, as the project's: the pending-writes thread is `PendingWritesManager_pollAll`
  (a pass with a zero select timeout, run after each write and from NETrecvNet, NETrecvGame, NETflush); the
  async connect runs inside the call and calls back before it returns (a host that never answers holds the
  joining screen up to its 15 s timeout). Keys and session keys cross between the exe and WzNet as bytes, a
  Sha256 as its 32 bytes; std::error_code is a Long (category in the high byte), net::result's error a ByRef.
* The merge (scratchpad j_merge1-3.py): the six stages copied in, 126 stand-ins removed by name (each with its
  comment lines) and the forks' own stand-ins reconciled through check_names (48 duplicates); what stays is
  `stubs_lobby` (the lobby server, the activity sinks, the update check, the kick redirect) and four permanent
  stand-ins in stubs_game; findHashOfFile real (`frame_hash_c`: SHA-256 is the exe's, not WzFramework's);
  build.py's shared folders; `--break-compat WzNet`. **Built at the first try**; the static checks 0 problems.
  The exe: 500 modules, its LINK.EXE line 12,696 characters through the subst drive (33,278 from bin\: over
  the limit), about 26,800 names.
* Found by the forks and the merge:
  * miniz's mz_inflate returns after copying out its dictionary with input and room left; zlib's inflate goes
    on, and 4.7.0's ZlibCompressionAdapter relies on it: every compressed stream corrupted after about 32 KB
    (JT's test; JZ's fix loops as zlib does; zlib_test checks the receive pattern against libzlib1.dll);
  * VB clears the last error before every declared call: GetLastError / WSAGetLastError through a Declare is
    always 0 (`Err.LastDllError`; JT wraps each Winsock call);
  * a local `ipv6_v6only` is the constant IPV6_V6ONLY in VB's case-insensitive names (the dual-stack socket
    fell back to two): check_names caught it;
  * a widget's virtual calls made inside WIDGET's constructor reach the most derived class (C++ reaches
    WIDGET's own): two screens failed with error 91 before their fields existed (JJ, JL guard their overrides;
    the library fix, WIDGET_construct setting the base kind until the derived constructor runs, is open);
  * stand-in signatures that were wrong (the join challenge sizes 32/32 for 128/64, NETrecvFile's type);
  * the UI scripts found their window by class name: with two copies, one script drove the other's window
    (drawtest_vb: the window of its own process now);
  * run_tests' C++ reference (statsref) stood in for code the port now runs (the stats outputs fell to 75 of
    90): it extracts multistat.cpp's score updates (its stand-ins of qtscript.cpp's trigger events update the
    stats as 4.7.0's do) and multijoin.cpp's sendDataCheck ("sent hash to host").
* Tests: zlib_test 235 of 235 (deflate byte-identical to miniz v1.11, read by Python's zlib and libzlib1.dll;
  inflate as zlib's on 71 streams; the network's receive pattern), crypto_test 1206 of 1206 (against libsodium
  through PyNaCl, hashlib, cryptography and crc.cpp's own texts), net_test 66 of 66 (loopback and a Python peer,
  plain and compressed, the error texts, the async callback, 8 MB of pending writes against a full socket); the
  full suite 3362 of 3362, the stats, models and maps' 90 of 90, the replay's 120 of 120 ticks, the sounds, the mixer and every media test.
* **The milestone: two copies play a multiplayer skirmish together** (`tools/lan_test.py`): the host started
  with --autohost (Sk-Rush, an open slot, two NullBots), the client through Multi Player > Join Game > the
  browser's IP box > the joining screen's handshake (the identity challenge signed, the session keys agreed) >
  the lobby > Ready; NET_FIREUP, both load and play, the AIs' orders reaching the client over the network:
  11 of 11 checks, **131.6 s of game time together, no desync** (gtime's checkDebugSync compares every
  player's game state CRC each tick). Hosting by mouse (lan_host.txt: Host Game, Start Hosting, Ready) works
  too. **With the official 4.7.0 build as the host** (its autohost, its NullBots) and the port joining through
  its browser: the handshake passes, they play **135.9 s of game time together with no desync on either side**
  (scratchpad j_interop.py): the port's protocol, crypto and simulation are 4.7.0's.
* Open: the official build as the client (its SDL3 takes clicks only from the real mouse: not run; done in
  section 38); the widget constructor's fix; untested paths: map and mod downloads, spectator moves, kick
  redirects (a stand-in), the vote widgets (notifications are a stand-in); the lobby server and port mapping
  (done in section 38), GameNetworkingSockets, the command interface (stand-ins by scope); a DLL split before
  the exe's 850 modules or names get near.
* Next: the user's choice (the remaining optional items of cluster K, the DLL split, or a shared build).

### 38. The internet lobby and port mapping (cluster J part 2)

* The user's request (2026-10-31): "lets try this Official build as client ... and Still stand-ins by design:
  the internet lobby server, UPnP port mapping". The user's choices: port libplum (4.7.0's first port-mapping
  library; miniupnpc, its fallback, stays out); one short hosting test on the real lobby at the merge; port
  mapping tested against a simulated router only.
* **The official 4.7.0 build as the client** (`--join=127.0.0.1`; its Ready box clicked with the real mouse,
  the user's approval): the port hosting (autohost, two NullBots) and the official client played about three
  minutes with no desync on either side (scratchpad j_interop2.py). Interop holds both ways.
* Found meanwhile: section 37's official-host test ran with 4.7.0's defaults, so the official build tried to
  map port 2100 on the router (it failed: no UPnP there; nothing changed) and registered a game ID with the real
  lobby ("awaiting publish"). Since then every test config has `portMapping=0` and a lobbyserver where nothing
  answers (`tools/ui_test.py`; `--real-network` leaves 4.7.0's defaults).
* Three forks (scratchpad L_DESIGN.md), private builds as in section 37:
  * LB: lib/netplay/netlobby.cpp, src/urlrequest.cpp and lib/widget/verifyprompt.cpp (the captcha and slide
    prompts) in the exe, about 4,000 lines; JPEG over GDI+ in WzGfx; WzNet's netlobby_c forwards netplay's calls
    through 14 INetGame methods, the events come back through 4 NetApi members; stubs_lobby's netlobby part,
    urlRequestData, urlRequestSetCancelFlag and urlRequestShutdown replaced;
  * PM: libplum at 4.7.0's commit 47be21b (ref/libplum, MPL 2.0) in 19 modules, port_mapping_manager and its
    libplum implementation (miniupnpc's slot a stand-in whose init fails), netplay's NETaddRedirects,
    NETremRedirects, NETinitPortMapping and NEThostGame's mapping branch as 4.7.0's; libplum's client thread
    and the manager's monitor are polls run from NETflush (PLUM_PENDING waits that resume on the next pass);
  * LT: LB's transport was WinHttp.WinHttpRequest, which cannot choose the address family. This machine has
    global IPv6 and the lobby (Cloudflare) has A and AAAA records, so the requests 4.7.0 sends over IPv4 (the
    listing, the tcp:ip4 registration) would have gone over IPv6 and the lobby would record the wrong address.
    curl's part on Windows is now WzNet's `urltransfer_c` (HTTP/1.1 over Winsock, Schannel through SSPI, curl
    8's defaults: certificate, name and revocation checks, TLS 1.2 at least, ALPN http/1.1, happy eyeballs for
    IP_ANY, gzip/deflate, redirects, MAXFILESIZE, low-speed limit), 1,900 lines, 7 NetApi members.
* Found by the forks: crc_sodium_c's base64Decode left a padded result one byte long, so every lobby join failed
  with "Invalid host identity" (fixed as the C++ resizes); netlobby's isResolvableLobbyError keys on the error
  code, not the data; GameDetails dropped the mods. In libplum, ported as they are (the official build does the
  same): the PCP lifetime written little-endian, the NAT-PMP epoch check comparing the new epoch with itself
  (leaving drops the mapping without unmapping it), PCP's epoch check never armed, a use after free (the port
  returns the text); the one deviation: no jitter where C divides by zero. private_build.py now compiles
  src\shared's miniz and sodium (forks had to pass them).
* The merge (scratchpad m_stage.py, m_stage2.py): LB's stage, PM's files and its edit scripts run over LB's
  netplay files (the diffs add up line for line), LT's files over both, the header comments; built with
  `--break-compat WzNet` (WzGfx 1.2 -> 1.3: one method more). The exe: 512 modules, the link line about 13,000
  characters through the subst drive, about 27,800 names.
* Tests: static checks 0 problems (5,035 log calls); run_tests 3362 of 3362, the stats 90 of 90, the replay 120 of
  120, sounds, mixer, every media test, http_test 46 of 46 (47 with the 60 s low-speed case); lan_test 11 of 11
  (131.3 s together, no desync); lobby_test 21 of 21 against the mock on 127.0.0.1 and ::1 (create, heartbeats,
  DELETE and request-join over IPv4, the tcp:ip6 registration over ::1, the lobby's connection check answered
  on both families, the captcha prompt, the join token checked by the host), `down` 1 of 1 (4.7.0's "Error
  connecting to the lobby server"); portmap_test's six scenarios 0 failed (PCP, NAT-PMP and UPnP mapping and
  unmapping with the fake's external address in the lobby's chat, no router, the 10 s timeout, port mapping off).
  External, read-only: the real lobby's games list over forced IPv4 and IPv6 (Cloudflare's challenge page at
  first, for curl.exe too, then the list), badssl.com's expired, wrong-host and self-signed certificates refused.
* **The one real-lobby hosting test** (scratchpad m_real_lobby.py; bin\Warzone.exe, 4.7.0's default lobby,
  `portMapping=0`, autohost Sk-Rush as "WarzoneVB test - please ignore" with a random password, 167 s): the
  create request over IPv4 got a game ID ("awaiting publish"), the tcp:ip4 registration went over IPv4 and the
  tcp:ip6 one over IPv6, heartbeats over IPv4 (TLS 1.2 each). The lobby's connection check could not reach port
  2100 (no mapping, the router's NAT and firewall), so the lobby answered HOST_NOT_ACCESSIBLE and dropped the
  listing itself; the game showed 4.7.0's texts ("Game not listed in the lobby." ... "Make sure port 2100 can
  receive incoming connections. If you're using a router, configure it to enable UPnP/NAT-PMP/PCP or to forward
  the port to your system."). The game was never listed publicly; with the listing gone, quitting sent nothing
  (the DELETE at quit is lobby_test's). What it shows: the real lobby accepts the port's create request,
  registrations and heartbeats over the families 4.7.0 uses.
* Open: a real router (never tested, the user's choice); a game the real lobby can reach (its connection check
  and the signed LobbyConnectCheckResponse ran only against the mock); TLS 1.3 (Windows 11), renegotiation; the
  slide and drag prompts and JPEG captchas at run time; DNS and the revocation check block a frame; miniupnpc;
  the video downloader (seqdisp's urlRequestData) and the activity sinks stay stand-ins.
* Next: the user's choice.

### 39. The screens and gameplay still standing in (cluster S)

* The user's request (2026-10-31): "work on 3. Gameplay and screens that are still stand-ins" (the in-game
  multiplayer menu, the limits screen, the Game Guide, notifications and overlays, player management, display
  settings, visual details, script-generated maps, the multiplayer paths never run); then "You can use the DLL
  split if needed to seperate a part out" (not needed: below).
* Budget first (scratchpad m3_budget.py): about 56 exe modules and 1,100 C++ names; the merge would split a DLL
  if the exe neared 31,000 names. Eight forks (scratchpad S_DESIGN.md: file ownership, the stand-ins'
  signatures as the contract, edits to shared files as exact pairs that stack, WzFunc id ranges, no side
  effects on the desktop):
  * S1: multimenu.cpp (the in-game multiplayer menu: players, scores, alliances, gifts), multilimit.cpp (the
    structure limits screen), lib/widget/resize.cpp; 15 checks;
  * S2: multilobbycommands.cpp (slash commands, admin identities), the kick reason popup and kick redirects,
    mute, colour names, the known players and stored mute options (4.7.0's SQLite tables as knownPlayers.json);
    34 checks with two copies; 4.7.0's mute time stored in 100 ns units and read as seconds, kept;
  * S3: guidescreen.cpp (the Game Guide: 13 widget classes, topics from data/base/guidetopics, the pop-ups that
    pause a campaign, save/load of the topics); the read state needs 4.7.0's activity database (not ported):
    every game shows the topics as new; 12 checks;
  * S4: notifications.cpp (in-game notifications, "do not show again" in notifications.json, the vote
    notifications), helpscreen.cpp (the help overlay: "help" in the chat), the overlays' z-orders; three
    notifications earlier code left as comments restored; 29 checks;
  * S5: spectatorwidgets.cpp (the stats overlay), hci.cpp's replay controller, blindwaitingroom.cpp; found:
    recvSwapPlayerIndexes stopped half-ported (a spectator moving to a player slot was kicked) and
    getModFilename was a stand-in (a host could never send a mod); 21 checks;
  * S6: lib/sdl's display functions in WzPlatform (resolutions as 4.7.0 filters them, both fullscreen kinds,
    display and cursor scale, the 28 cursors, mouse grab, minimize on focus loss, clipboard), urlhelpers.cpp;
    tested through the TEST-ONLY WZ_TEST_NO_DESKTOP_EFFECTS (nothing on the desktop changes); 4.7.0 never caps
    a new window to the work area (an SDL3 result tested the wrong way round): ported as it is; 38 checks;
  * S7: the droid death animations, the tracking camera's pitch at the map's edge, the terrain qualities
    offered as 4.7.0 offers them (Classic and Remastered need their packs: setup_data now places the official
    classic.wz; Remastered's high.wz is KTX2, not ported, so it stays greyed out), the research model;
  * S8: map_script.cpp (script-generated maps: 10c-WaterLoop, the DustyMazes, 6c-Entropy) in WzMap over
    chakra.dll; byte-identical to 4.7.0's own map_script.cpp with its pinned QuickJS (ref/quickjs-wz, built
    with g++) over 32 seeds each; found: the VB process runs the x87 at 64-bit precision (Math.sqrt was 1 ulp
    off: WzMap sets 53 bits whenever control returns to a script), cdecl functions cannot be Declares
    (DispCallFunc), and Chakra without its background threads ran the maps 14 times slower.
* The merge (scratchpad s_merge.py): each fork's edit scripts first rerun on a pristine src to check they
  reproduce its stage (all did; S5's hand-edited stubs removed by name), then stacked fork after fork on one
  tree (every old text found once; the shared files' diffs add up line for line), the widget classes
  regenerated, `--break-compat WzPlatform` (WzFramework 1.15 -> 1.16). The first build stopped at WzNet with
  an empty VB6 log (no fork changed WzNet); the same build again succeeded: a one-off VB6 crash. The exe:
  563 modules, about 29,200 names (no DLL split needed yet), its link line about 14,400 characters.
* Found at the merge: run_tests' stats, models and maps fell to 52 of 90. The map blocks run one after another
  in one process; the port now generated and played the script maps while statsref (and mapref) could not load
  them, so every later block started from other leftovers (the players' templates). mapref and statsref now
  link map_script.cpp with QuickJS: 90 of 90, WzTests 3437 of 3437. notifications_test's "videos missing"
  runs used bin\Warzone.exe, whose ..\data has the videos (the game searches the data beside its bin, as
  4.7.0): it runs a copy of the exe now.
* Tests on bin\: run_tests whole (3437 of 3437; 90 of 90; the replay 120 of 120; sounds, mixer, media, zlib,
  crypto, net, http, the script maps 68 of 68); sS1 15, lobby_commands 34, guide 12, notifications 29, s5 21,
  s6_display 38, the S7 and S8 UI scripts, save_menu / save_load (with closeguide), lan_test 11 of 11
  (130 s together, no desync), lobby_test 21 of 21 and down, portmap_test 0 failed. Campaign 1 walked from the menu to its end (19 levels, 1,406 s, no error), the
  guide's pop-ups closed by the walk (make_campaign_walk.py's closeguide).
* Open: the guide's sidebar toggle icon draws no pixels (at 100% display scale 4.7.0's own: section 40); the
  Team Strategy tab, completed alliances and the research/radar/droid gifts not run; the kick redirect's host
  side (only the command interface sends one); modal notifications; real fullscreen, mode switches, the real
  clipboard, links and the cursor clip never exercised (the no-desktop rule); Remastered terrain (KTX2: section
  40); the
  game scripts' x87 precision and Chakra's background threads in the exe; a map script that never ends hangs
  the load (4.7.0 stops it after 30 s); a map package in level.json's format, a map-mod package and a downloaded
  mod mounted in a started game not run; the map export, .gam writers, map stats and replays' embedded maps.
* Fork S9, after the merge (S5's open items): **downloads work**. Every received map or mod had failed its check
  at the last chunk and been deleted: validateReceivedFile's C `!PHYSFS_eof(...)` was ported as
  `Not PHYSFS_eof(...)`, and VB's bitwise Not of 1 is -2, true. Map packages: wzphysfszipioprovider.cpp and
  `WzMapZipIO` (libzip's role taken by PhysFS's zip reader, mounted under a private name), map_package.cpp's
  read side and map_levparse.cpp in WzMap, PHYSFS_mountHandle, levParse_JSON, levAddWzMap, buildMapList's
  CheckInMap: a host serves a map pack from maps\ and the client downloads it byte for byte and plays it. The
  missing Ready box was 4.7.0's own: autohost's readybeforefull defaults to true, which hides the boxes while a
  slot is open. s5_test 43 of 43 (download 10, mod 4, ready 8, the rest 21); run_tests whole again (3437, 90,
  120, the media), lan_test 11 of 11. The exe: 567 modules, about 29,200 names, its link line 14,553.
* Next: the user's choice (from the list given 2026-10-31: sharing the build and testing it on another PC,
  the DLL split, performance).

### 40. The guide icon, the Remastered terrain (KTX2, cluster T) and fullscreen

* The user's request (2026-11-01): "Guide icon: the guide's sidebar toggle icon draws nothing, though the
  button works. and Remastered terrain: it needs KTX2 textures, which aren't ported." Then: "Also has full
  screen mode been tested?" (answered: never for real; the user allowed desktop fullscreen only).
* **The guide icon is 4.7.0's own.** The official build, run test-safe into CAM_1A where the guide pops up,
  draws no icon either at 100% display scale: image_sidebar_list.png (64x64, three 4-pixel lines) is drawn
  12x12 from the UI atlas, which has no mipmaps, with bilinear sampling, and every sample falls between the
  lines. At 150% both builds draw it, pixel for pixel the same. Not changed.
* **KTX2 is ported** in a new optimised DLL, WzBasis (src\basis; references WzFramework; WzGfx and the exe
  reference it). high.wz is 579 KTX2 files of UASTC LDR 4x4 data in zstd 1.4.9 frames; 4.7.0 transcodes them at
  load to the GPU's best format (BC7 on this RTX 3080; single-channel maps R8, normal maps RGBA32 then RGB8).
  Contract first (scratchpad T_DESIGN.md: WzBasis's skeleton classes BasisApi, ktx2_transcoder,
  ktx2_image_level_info), the reference `tools/ktx2ref.cpp` (g++ over ref/basis_universal at df079ec, the
  commit 4.7.0 pins), three forks:
  * T1: zstddeclib.c (zstd 1.4.9's decoder) as 4.7.0's 32-bit builds compile it; identical to that C built
    32-bit on 1,401 generated frames (526 of them errors) and on all 5,391 frames of high.wz, about 120 MB/s
    optimised (the C 680). Found: huf_decompress.c's 4-stream X2 loop has clang's branch and gcc/MSVC's `#else`;
    the port takes `#else` (clang's failed 51 frames).
  * T2: basisu_transcoder.cpp's KTX2 container and UASTC LDR transcoder to RGBA32, RGB565, BGR565, RGBA4444,
    ASTC 4x4, BC7, BC1, BC3, BC4, BC5: every level of the 618 KTX2 files of high.wz and base.wz byte-identical
    to the C (built -DNDEBUG, as 4.7.0's release; the assert build aborts on two broken files), 31 broken files
    accepted and refused as the C does; 3-5x the C's time (BC7 about 100 ms for a 1024 page with its mips).
    Found: a Single stored in a scalar variable and reused at once keeps the x87's precision in this build
    (`Fix` of a product exactly 1024.0f gave 1023): the transcoder's float helpers pass results through a ByVal
    Single. float_vb's F32 (an array store) was probed in both build modes and rounds correctly.
  * T3: the WZ side: gfx_api_image_basis_priv.cpp whole, gfx_api.cpp's BASIS_ENABLED branches, iV_BaseImage
    as an interface (iV_Image and iV_CompressedImage implement it), the textures' uploads taking it,
    GFX::loadTexture's .ktx2, setup_data placing the official high.wz. The game side was KTX2-aware already.
* The merge (scratchpad t_merge.py: each fork's pairs rerun on src first) built with `--break-compat WzGfx`
  (WzBasis 1.0 -> 1.1 for ktx2_transcoder.find_key). Remastered then drew wrong with the real data: Urban and
  Rockies far too bright, Arizona's ground black (gfx_api.cpp:1073's assert). T3 again (T3b) found the cause in
  the data: 4.7.0's build ships base.wz's 30 ground pages (resized to 1024) and 9 backdrops as KTX2 without
  their PNGs, and a .png anywhere in the search path is taken before a .ktx2, so the source tree's 2048 PNGs hid
  high.wz's pages of the same names. setup_data now extracts the official base.wz's 39 .ktx2 into
  ref\warzone2100\data and leaves the 39 PNGs out of data\ (Normal quality draws the 1024 KTX2 pages too, as the
  official build; peak memory in CAM_1A 330-384 MB -> 154-174 MB).
* Tests on bin\: zstd_test and ktx2_test (both in `run_tests.py --media`); ktx2_game_test: regress (skirmish,
  campaign with its video, save/load), the options offering Remastered with high.wz, switching to it,
  `tilesets` against the official build's captures (mean difference per channel over the view: CAM_1A 0.72 /
  0.97, CAM_2A 5.18 / 5.11, CAM_3A 2.48 / 2.94 at quality 512 / 1024: moving units and effects), `normaltilesets`
  (0.43-5.59), `dump` (every level the game transcodes, 3,610 per compression setting, identical to ktx2ref, and
  every image handed to an upload identical to 4.7.0's conversion: BC7, RGB8, R8, RGBA8), `perf` (.gam load to
  level running: Normal 10.3 s, Remastered 11.1 s at 512, 11.3 s at 1024). The default stays Normal, as 4.7.0's
  32-bit builds have it.
* **Desktop fullscreen tested for real** (the user's go-ahead for that kind only; no display mode change): the
  TEST-ONLY `WZ_TEST_NO_DESKTOP_EFFECTS=borderless` makes only the borderless window, the config asks for Desktop
  Full (fullscreenWidth/Height 0: 4.7.0 then never sets a mode), and scratchpad fs_test.py read every display's
  mode every 0.25 s (unchanged throughout). The port starts in it (a 1920x1080 popup over the primary display,
  "Changing fullscreen mode to Desktop Full" as 4.7.0 logs it), Alt+Enter goes to the 1024x768 window and back,
  the setting is saved, a level draws at 1920x1080, and a switch in the middle of a level and back draws on. The
  official build's window takes the same rect and the same popup style (but for WS_CLIPCHILDREN, SDL's). A black
  view in the first level runs of both builds was edge scrolling: the real pointer rested on the display's top
  row. Outside resizes in a level (1280x960, 1600x900, 1920x1061) draw on in both builds. Exclusive fullscreen
  (a display mode change) remains untested.
* Side findings:
  * tools/draw_test/drawtest.png is out of date (the title menu and the build date are in today's shot, and the
    backdrop is KTX2 now); the primitives it draws are unchanged.
  * CAM_x started with `--game` gets random starting unit designs in both builds (the official build's differed
    between its own runs): 4.7.0's.
  * A window taller than the screen is clamped by Windows in the port (1061 rows) but not by SDL3 (1152); a new
    window's client area is centred 11 rows lower than SDL3 centres it. Open.
* Open: exclusive fullscreen; the ETC/PVRTC targets (desktop GL never needs them).

## Architecture decisions

### Between projects

* Each DLL exposes one **GlobalMultiUse API class** (`FrameworkApi`,
  `MapApi`, `SoundApi`, `GfxApi`). The C++ files are ported into the DLL's
  private `.bas` modules, and the API class forwards to them. Callers use
  the procedures as if they were their own.
* Calls from a DLL up into the exe (the game clock, console messages, sound
  positions) go through one interface class per DLL (`INetGame` for WzNet).
  The exe implements it (`NetGame.cls`) and hands it over at startup
  (`NetSetGame`). Each method keeps the C++ name of the game function it
  stands for, **without underscores**: VB refuses them in an interface used
  with `Implements`, so each underscore is dropped and the next letter
  capitalised (`version_getVersionString` -> `versionGetVersionString`).
* A DLL that holds a reference to an exe object (the `INetGame`) gets
  `Nothing` before the exe ends.
* Big arrays that both sides read belong to the side that writes them. The
  other side sees them through a SAFEARRAY alias, with the UDT declared on
  both sides. UDT parameters on public methods need the UDT in a public
  class module of the DLL; prefer handles and arrays.
* DLLs build with binary compatibility. Change a DLL's public interface only
  by adding to it; `--break-compat` is for the rare rebuild.
* `bin/` runs reg-free: the exe embeds a manifest generated from the DLLs'
  type libraries. VB6 registers each DLL it builds, and `--unregister`
  undoes that.

### Memory model (QuakeVB's, extended for C++)

| C++ | VB |
|---|---|
| `DROID`, `STRUCTURE`, `FEATURE`, `PROJECTILE` (`BASE_OBJECT` subclasses) | UDT pools whose arrays start at type * &H10000000 + 1: a handle is the typed and the base pointer, 0 = NULL; base fields through a `BASE_OBJECT*` and `virtual` calls become `Select Case` on the handle's pool (CONVENTIONS.md, "Objects") |
| stats (`WEAPON_STATS`, `STRUCTURE_STATS`, `RESEARCH`, ...) | one flat UDT per type in an array with a count; a typed pointer is the index (-1 = NULL); a `BASE_STATS *` is a handle, `StatType` plus index; templates live in a pool (1-based, 0 = NULL) |
| models (`iIMDShape`, `iIMDBaseShape`, in WzGfx) | two pools with free lists; a pointer is a 1-based handle, 0 = NULL; `std::unique_ptr` chains (`next`) become handle chains freed by their owner |
| per-player object lists | index chains through the pools |
| the map (`MAPTILE *`, `psAuxMap[player]`) | a tile index into `psMapTiles`; the aux and blocking maps are 2-D Byte arrays (`psAuxMap(player, tile)`) |
| object `id` lookups | integer hash table |
| widgets (`W_WIDGET` tree, owned by `shared_ptr`) | VB6 class modules: reference counting matches `shared_ptr`; `Implements IWidget`, with base behaviour by containment |
| lambdas, `std::function` | in algorithms: plain loops. As callbacks: a dispatch number plus captured state, or a small handler class |
| `std::vector` | dynamic array plus count |
| `std::map`, `unordered_map` | hash module for Long keys; `Scripting.Dictionary` for string keys off the hot paths |
| `WzString`, `std::string` | VB String (UTF-16); UTF-8 only at file and network boundaries |
| templates | one copy per type used |
| exceptions | `Err.Raise`, or return codes |
| `int64_t` | Double (exact to 2^53) for times and sizes; Currency bit patterns for hashes; checked case by case |
| `uint32_t` arithmetic, shifts | QuakeVB's `ULT`/`SAR`/`BIT_` helpers |
| threads (path-finding jobs, loaders) | none: jobs run synchronously |
| JSON (`nlohmann`) | the `nlohmann_json` class in WzFramework, one object per value (section 9) |

### Scripts: Chakra instead of QuickJS

* Port `wzapi.cpp` (4,856 lines) and `qtscript.cpp` (2,903: events, timers,
  script management). Replace `quickjs_backend.cpp` (3,833) with a Chakra
  backend of about the same size. All of these go in the exe.
* Every JS-callable VB function traps its own errors and reports them with
  `JsSetException`. A VB error must never unwind through Chakra (QuakeVB's
  rule: no engine code inside callbacks without a trap).
* `chakra.dll` is the legacy Edge engine. It ships with Windows 10 and 11
  but is no longer developed. Fallback: ChakraCore (open source, same JsRT
  API, x86 build) beside the exe.

### Renderer: Warzone's GL backend, ported (WzGfx.dll)

* Port the `gfx_api` interface (`gfx_api.h`/`.cpp`, 1,110 lines) and
  `gfx_api_gl.cpp` (5,541), plus `piedraw`, `pieblitfunc`, `tex`, `screen`,
  `imdload`. Keep `data/base/shaders` unchanged.
* GL entry points go into a table filled by `wglGetProcAddress`. Calls go
  through DispCallFunc, using wrappers generated from the GL headers.
* First target: OpenGL 3.3 core without the optional effects; they come
  later.
* Text: GDI (`ExtTextOutW` into a DIB, uploaded as a texture) replaces
  FreeType, HarfBuzz and FriBiDi. Warzone's own fonts (`data/fonts`: DejaVu,
  Noto) load with `AddFontResourceEx` (private to the process). Windows does
  the shaping and BiDi.
* Images: PNG through the GDI+ flat API (`gdiplus.dll`).

### Sound (WzSound.dll)

QuakeVB's waveOut mixer, with stereo panning and distance attenuation in
place of OpenAL's 3D sources. All 560 sound effects are Ogg Vorbis, so
stb_vorbis (about 5,000 lines of C) is ported. The music is Opus, and
porting libopus is out of proportion: convert the music to Vorbis with a
tool, or skip it.

Done (section 32): the OpenAL calls openal_track.cpp makes are implemented by a mixer in VB
(`al_win`: WZ sets AL_NONE and computes the gains itself, so OpenAL only mixes, resamples and
pans), polled from sound_Update over waveOut as QuakeVB's; stb_vorbis v1.22 is the decoder,
checked against ffmpeg's. No music ships with the data.

Done (section 34): the music is ported as Opus after all (the user's choice): opusfile and
libopus 1.5.2's decoder (CELT, SILK) in WzSound, checked against the official build's
libopus.dll; tools/setup_data.py puts the official music in data/music.

### What is not ported

Vulkan and the null backend, Discord, curl's other uses (news, updates,
video downloads: the lobby's requests are ported over WzNet's own
transfers, section 38), GameNetworkingSockets, miniupnpc (libplum is
ported, section 38), the crash handler, stdin, the script debugger
(perhaps later), SQLite player statistics (a JSON file instead),
Basis Universal's ETC1/ETC2/EAC, PVRTC, ATC, FXT1 and HDR targets and its
ETC1S and HDR inputs (KTX2 files of UASTC data are ported, section 40),
the JPEG encoder, Emscripten, translations (later).
(Planned out at first and ported since: libsodium's identities and
session keys, autohost, the internet lobby: sections 37 and 38.)

### Multiplayer

The goal was VB to VB over TCP (LAN or direct IP). Play against official
builds needs a bit-exact simulation, the same protocol and version
string, and libsodium signatures; it was not promised, and it works: the
port plays with the official 4.7.0 build as host and as client with no
desync (sections 37 and 38), and registers its games with the internet
lobby (listed once the lobby can reach the host's port).

## Testing: replays and syncDebug

This is the port's equivalent of DoomVB's tic-by-tic demo test. Warzone
records skirmish and multiplayer games as replays (`lib/netplay/netreplay.cpp`:
the settings plus every network message). For desync hunting it can log
every object's state, tick by tick (`syncDebugDroid`, `syncDebugStructure`,
...). The official 4.7.0 build records replays and syncDebug logs. The VB
port plays the same replays headless and must write the same log, line by
line. Hidden runs and screenshots work as in QuakeVB.

To check in cluster B: AI orders should appear in a replay as network
messages. If so, playback does not depend on Chakra behaving exactly like
QuickJS. If not, use replays of human-only games for the bit-exact test.

## Plan: clusters

Each cluster is compiled and tested before the next. The simulation comes
first and is tested headless against replays, as QuakeVB tested its server
before it had a renderer.

| Cluster | Content | Milestone |
|---|---|---|
| 0 Foundation | **Done:** 4.7.0 source with textures and fonts; probes; `build.py` (seven projects, `src/shared`, binary compatibility, embedded manifest); `docs/CONVENTIONS.md`; the checks (`check_names`, `check_intdiv`, `check_with`, `check_debug_lines`, each tested on planted errors); WzFramework: `debug.cpp`, printf, C stdio and time, `wztime.cpp`, wzapp's ticks and dialogs, PhysFS's write dir; WzPlatform: `main_sdl.cpp`'s input, window, event loop and screen setup, `gfx_api_gl_sdl.cpp` on WGL; exe: `main.cpp`'s realmain and mainLoop in part; `frame.cpp`, `trig.cpp`, `crc.cpp`'s CRC and `Sha256`, `version.cpp`, `warzoneconfig.cpp`, inih and `INIReaderWriter`, `configuration.cpp` (the config file), with stand-ins for the settings it reads in the DLLs that will own them; `clparse.cpp` (the whole command line), `i18n.cpp` (the language table and thread locale; no gettext), `wzpaths.cpp`, `modding.cpp` (without the mod hashes), `init.cpp`'s search path, `main.cpp`'s user folders, `scanDataDirs` and mod checks: the game data in `data/` (a junction to `ref/warzone2100/data`) is mounted as in an installed game; nlohmann::json 3.11.3 (`nlohmann_json` class, `json_hpp`: parser, serializer, Grisu2), `wzconfig.cpp` (`WzConfig`, `json_variant`), `wzjsonhelpers.cpp`, `physfs_ext.cpp`'s enumeration, `file_ext.h`; `tools/run_tests.py` checks them against the real C++ (`tools/jsonref.cpp`). **Cluster 0 is done** (see "First run", "The config file", "Command line, data and language" and "JSON and WzConfig" below). **Next:** cluster A | window opens and clears with GL; config and logs written |
| A Data | stats (weapons, bodies, propulsion, structures, research, features, templates), PIE loader (`imdload`), wzmaplib, `map.cpp`, levels. **Done:** A1, wzmaplib's core in WzMap.dll, checked against the real library on every map of the game data (`tools/mapref.cpp`; see "wzmaplib"); A2, the resource system and the `.wrf` grammar (WzFramework), `levels.cpp` and its lexer, `data.cpp`, `systemInitialise` / `systemShutdown` (see "Resources and levels"); A3, the stats, structure, feature and research stats, templates, string resources and sound ids, checked against the real loaders (`statsref`; see "Stats"); A4, `imdload` in WzGfx, every PIE model checked bit for bit against the real loader (see "PIE models"); A5, the game's `map.cpp` with gateways, continents and the map queries, every map folder checked against the real code (see "The game's map"), and the factions; the milestone with cluster B's first steps (see "The droids"). **Cluster A is done.** | a hidden run loads a skirmish map and dumps its objects and stats |
| B Simulation (headless) | objmem, baseobject, droid, structure, feature, move, action, order, combat, projectile, ai, astar/fpath, visibility, power, research, gtime, random, multiplay (local); in WzNet: nettypes (message decoding), netqueue, netreplay (reading), sync_debug. **Done:** B1, the object model, objmem and the features; B2, the structures with walls, modules, flags, power links and tile visibility; B3, the droids; every map's droids, features and structures checked against the real code (see "The objects", "The structures", "The droids"); B4, power, the object grid and WzNet's sync log, every map's sync log checked line by line with its CRC (see "Power, the object grid and the sync log"); B5, path finding (see "Path finding"); B6, movement, steering and formations (see "Movement"); B7, actions and orders (see "Actions and orders"); B8, combat: projectiles, targeting, visibility, damage (see "Combat"); B9, the structures' update, research, production and gameStateUpdate (see "The structures' update"); B10, the game clock and the game messages (see "The game clock and the game messages"); B11, replay files, recorded and played back (see "Replay files"); B12, the game's start from a replay: an official 4.7.0 replay starts headless and its first tick's sync log matches the official build's but for rules.js's lines (see "The official build's replays and the game's start"). **Done (with cluster C):** the official replay plays to its end with identical sync (see "The AI skirmish"). **Cluster B's milestone is reached.** **Next:** the playback's speed (214 s against the official 29 s) | a 4.7.0 replay plays headless with **identical syncDebug**; speed measured on a large late-game replay |
| C Scripts | Chakra backend, wzapi, qtscript (events, timers), rules.js. **Done:** C1, the script engine whole (wzapi.cpp, qtscript.cpp, the Chakra backend in place of quickjs_backend.cpp): rules.js runs, and the official replay's 120 reference ticks match (see "The script engine"); C2, the skirmish's start and what the AIs reach: **four AIs play a headless skirmish to its end** (see "The AI skirmish"). **Cluster C's milestone is reached.** | headless skirmish where the bots (Nexus, NullBot, BoneCrusher, Cobra) play a game to the end |
| D Renderer base | gfx_api, GL backend, textures, pieblit, screen, text via GDI, image sheets. **Done:** lib/ivis_opengl but piedraw.cpp in WzGfx (gfx_api, the GL and null backends, images, PNG on GDI+, textures, image sheets, palette, the 2D blits, screen, text on GDI) and the exe's loading screen (see "The renderer base"). **Cluster D's milestone is reached.** | images and text on screen; screenshot test |
| E 3D display | display3d, terrain, piedraw, effects, atmos, lighting, warcam, radar, bucket3d. **Done:** piedraw and the models' GPU side in WzGfx; terrain, texture, lighting, shadow cascades, advvis, display3d, bucket3d, miscimd, component, effects, atmos, display, warcam, animation, radar in the exe (see "The 3D display"). **Cluster E's milestone is reached.** | watch an AI skirmish in 3D |
| F Widgets and front end | lib/widget, frontend, titleui, options, key bindings. **Done:** lib/widget whole, the title screens, the options, the skirmish lobby (see "The widgets, the front end and the in-game interface"). **Cluster F's milestone is reached.** | main menu; options saved |
| G In-game UI | hci, intdisplay, intorder, design, intelmap, console, chat, in-game options, load/save screens. **Done:** hci, intdisplay, the object and stats forms, edit3d, design, the input manager and key bindings, selection, console, intorder, the in-game options, intelmap (load/save screens: cluster I; the chat: J). **Cluster G's milestone is reached.** | **play a skirmish against AI from the menu** |
| H Sound | mixer, stb_vorbis port, audio, track, audio_id. **Done:** stb_vorbis, an OpenAL subset mixed over waveOut, openal_track, audio, track, cdaudio, playlist, the game's side (see "The sound"). **Cluster H's milestone is reached.** | sound effects in game |
| I Campaign | levels, mission (off-world), transporter, campaign scripts, save/load (`game.cpp`), briefings without video. **Done:** game.cpp's save and load, mission, transporter, the campaign selector, seqdisp, scores, loadsave, challenge, message, cheat, the chat screen (see "The campaign, save and load"). **Cluster I's milestone is reached.** | campaign 1 from start to finish; save/load round trip |
| J Network | the rest of netplay in WzNet (connections over Winsock TCP, zlib streams, sessions) with the full `INetGame`; direct-IP joining; the multiplayer parts of multiint. Scope (the user's, section 36): what the milestone needs (in the exe multisync, multivote, multijoin_helpers, the joining and pre-game screens, team strategy); the internet lobby and game browser, blind waiting room, spectator widgets, lobby commands, stdinreader, port mapping and GNS stay stand-ins. **Done** (section 37): WzNet's connections, zlib streams, crypto, netplay whole but the lobby server and port mapping; the joining screen and the game browser's IP box; the pre-game screen, sync, votes, team strategy; two copies play over the network, and the port plays with an official 4.7.0 host. **Cluster J's milestone is reached.** Part 2 (section 38): the official build as the client; the internet lobby (netlobby, urlrequest over WzNet's Winsock + Schannel transfers, the verify prompts) and port mapping (libplum) ported, tested against a mock lobby and a fake router, and once on the real lobby | two copies play a skirmish over LAN |
| K Optional | `.wz` archives (zip, inflate) for maps and mods, music, translations, shadows and other effects, videos, interop with official builds. **Done:** PhysFS's zip archives, the campaign mods, the music (opusfile, libopus), the videos (libtheora, lib/sequence in WzSequence), quick chat and chat options; campaigns 2 and 3 walked (see "Archives, mods, music, videos, quick chat"); the Music Manager, the videos' display checked against 4.7.0's in every size and scanline mode, data/ a mirror (see "Cluster K's open items"). Left: translations, shadows and the other effects, interop | |
| S Stand-ins of screens and gameplay | what the earlier clusters left standing in that a player meets: the in-game multiplayer menu, the limits screen, the Game Guide, notifications, the help overlay, player management and lobby commands, the spectator and replay overlays, the blind waiting room, the display settings, death animations, the camera's edge pitch, the terrain qualities, script-generated maps, map and mod downloads. **Done** (section 39) | every screen 4.7.0 shows a player is real |
| T Textures | KTX2 (Basis Universal's UASTC transcoder, zstd's decoder) in a new DLL WzBasis, the WZ side in WzGfx, the data laid out as 4.7.0's base.wz; the Remastered terrain. **Done** (section 40) | Remastered terrain drawn as the official build draws it |

## Risks

1. **Volume and C++.** 243,000 code lines of modern C++. The UI (frontend,
   titleui, hci, multiint, design) is built on lambdas, `shared_ptr` and
   class hierarchies, and is the slowest part to translate.
2. **Name budget of the exe.** It starts at 22,264 of about 32,100 names,
   before VB-side names (dispatch constants, helpers, `INetGame_` and other
   callback procedures). Run `name_budget.py` every cluster; the script
   bridge can still move out to a DLL.
3. **Speed.** VB6 native code without optimisation (DoomVB's rule), but for
   WzSound (fast code, no bounds checks: its decoder and mixer are arithmetic the game's state
   does not depend on; the user's choice, 2026-10-29). The
   simulation runs 10 ticks per second; visibility and path-finding with
   hundreds of units are the hot spots. Measure in cluster B. The GPU does
   the drawing, and a GL call costs 37 ns.
4. **chakra.dll** is a legacy Windows component (fallback: ChakraCore).
5. **OpenGL 3.3** is required, as for Warzone itself. It was tested here
   only on NVIDIA; Intel and AMD drivers are still to check.
6. **VB6 build limits:** 850 modules and about 32,100 names per project,
   and LINK.EXE's 32,767-character command line (each module's quoted
   <outdir>\<module>.OBJ); `build.py` reports each. The exes build
   through a `subst` drive (section 36): after cluster S the exe's line is
   about 14,400 characters for 563 modules (37,600 from bin\: over), so
   the 850 modules and the names come first: about 29,200 names of 32,100
   (src/shared/sodium counted). A DLL split (section 36; approved by the
   user, 2026-10-31) when the exe nears 31,000 names.

## Tools

| Command | Does |
|---|---|
| `python -I tools\build.py [--release] [--break-compat NAME] [--unregister]` | builds the DLLs and `Warzone.exe` into `bin/` (see its docstring) |
| `python -I tools\check_names.py` | VB name clashes per project, with the globals of the DLLs it uses |
| `python -I tools\check_intdiv.py` | `\` that a later `*` or `/` takes apart (QuakeVB's) |
| `python -I tools\check_with.py` | `GoTo` out of a `With` on an array element (QuakeVB's) |
| `python -I tools\check_mod.py` | `Mod` that a later `*`, `/` or `\` takes apart |
| `python -I tools\check_debug_lines.py` | each log call's C++ line, format and function against `ref/warzone2100` |
| `powershell -NoProfile -ExecutionPolicy Bypass -File tools\capture_window.ps1 -exe bin\Warzone.exe -arguments "--frames=300" -out shot.png` | runs the game, captures its window, reports the centre pixel and the run time |
| `powershell -NoProfile -File tools\resize_window.ps1 -exe bin\Warzone.exe -arguments "--configdir=<dir> --frames=240" -width 900 -height 700` | runs the game and resizes its window part way (the saved size is then in `<dir>\config`) |
| `python -I tools\draw_test.py [name] [--keep]` | cluster D's screenshot test: `Warzone.exe --vbdrawtest` draws text, images and boxes over the title screen, saves a screenshot and quits; compared pixel by pixel with `tools/draw_test/<name>.png` |
| `python -I tools\watch_test.py [gameTime] [test] [timeout] [--realtime] [--player=N]` | cluster E's test: an AI skirmish in a window; the game's own screenshot at the game time (`--vbwatchshot`), the frame rate every 100 frames |
| `python -I tools\ui_test.py <script> [timeout] [--keep-config] [--config-of=<script>] [args...]` | clusters F and G's tests: a `tools/ui_tests` script of clicks, keys and text posted to the game's window (`--vbscript`), the game's own screenshots, the widget tree in the log; prints the log's errors and asserts |
| `python -I tools\gen_widgets.py` | writes `src/game/widget_classes_vb.bas` from the widget class declarations (see CONVENTIONS, "Widgets") and prints its problems |
| `python -I tools\exe_debug.py` | rebuilds `Warzone.exe` with CodeView info and no optimisation (a .pdb, for cdb); `tools\build.py` rebuilds the release one |
| `python -I tools\stack_scan.py [bytes]` | the return addresses into the game's modules on the running Warzone.exe's main thread (cdb, non-invasive) |
| `powershell -NoProfile -File tools\dialog_text.ps1 [process]` | the text of a running process's dialog boxes (VB's run-time error box) |
| `python -I tools\gen_gl.py` | writes `src/gfx/gl_win.bas`, the GL entry points and constants gfx_api_gl.cpp uses, from glad's header and loader |
| `python -I tools\ai_skirmish.py [test] [timeout] [args...]` | plays a headless AI skirmish (`tools/skirmish_tests`) to its end and summarises the log and the scripts' performance logs |
| `python -I tools\replay_test.py [name] [--until GAMETIME] [--timeout SECONDS] [--keep]` | plays an official replay (`ref/official/replays`) in `Warzone.exe` headless and compares its sync logs tick by tick with the official build's (`ref/official/synclogs`) |
| `python -I tools\run_tests.py [--keep] [--stats]` | builds `bin\WzTests.exe`, `build\jsonref.exe`, `build\mapref.exe` and `build\statsref.exe` (g++), compares the port's JSON, WzConfig, wzmaplib, stats, model and map output with the real C++'s (see "JSON and WzConfig", "wzmaplib", "Stats", "PIE models", "The game's map"); `--stats` runs only the statsref tests |
| `python -I tools\run_tests.py --sound [--keep]` | cluster H's tests (also in the full run): builds `bin\VorbisTests.exe` (src/vorbistests) and decodes the 560 game sounds against ffmpeg (its decodes cached in `build/sound_ref`), then `tools\mixer_test.py` |
| `python -I tools\mixer_test.py` | builds `bin\MixerTests.exe` (src/mixertests: al_win alone) and checks the mixer's capture of synthetic sources: gain, panning, resampling, looping, queues, states, clipping |
| `python -I tools\ui_test.py <script> --sound --vbsndcapture=<file.wav>` | a UI test with the sound on, mixed into a WAV file instead of the device |
| `python -I tools\wav_activity.py <file.wav> [window s]` | a captured WAV's loudness per window and channel (where the sound plays, left against right) |
| `python -I tools\check_redim.py [project]` | a fixed-size array passed to a procedure that ReDims or Erases that parameter (VB's run-time error 10) |
| `python -I tools\save_test.py [level] [--keep]` | cluster I's save round trip: a level saved at its start (`--saveandquit`), loaded and saved again, the two saves compared (but for what 4.7.0 changes too, each named with its C++ line) |
| `python -I tools\ui_test.py campaign_walk.txt 1800` | campaign 1 from the main menu to CAM_1END (cheats and transporter launches; `tools\make_campaign_walk.py` writes the script); `save_menu.txt`, then `save_load.txt --config-of=save_menu`: a mid-mission save from the in-game menu, loaded and saved again |
| `python -I tools\setup_data.py [--check]` | puts the data 4.7.0's source tree lacks where 4.7.0 looks for it: the official build's music and campaign mods, the videos (ref/videos); then mirrors ref/warzone2100/data as data/ (real folders, hard links: no junction); the official terrain_overrides\classic.wz and high.wz |
| `python -I tools\zstd_test.py [--quick] [--no-ktx2]` | zstd's decoder (WzBasis) against zstddeclib.c built 32-bit and the 64-bit library: generated frames and every frame of high.wz's KTX2 files, ZstdTests (bounds-checked) and ZstdTestsFast (optimised); in `run_tests.py --media` |
| `python -I tools\ktx2_test.py [--no-bounds] [--formats 13,6,...] [--raw]` | the KTX2/UASTC transcoder (WzBasis) against Basis Universal's C (-DNDEBUG, as 4.7.0's release): every level of the official high.wz and base.wz KTX2 files in 14 formats (4 must fail: not ported), 31 broken files; in `run_tests.py --media` (`--no-bounds`) |
| `python -I tools\ktx2_game_test.py [--exe PATH] regress options switch tilesets dump perf` | the Remastered terrain in the game: regressions, the options offering it, switching to it, three tilesets against the official build's captures, every transcoded level the game loads against the C (`WZ_TEST_KTX2_DUMP`), load time and memory |
| `python -I tools\gen_basisu_tables.py` | writes `src/basis/basisu_tables_vb.bas` from basisu_transcoder's tables |
| `build\ktx2ref\ktx2ref.exe <format> <outdir> <files>` | Basis Universal's transcoder as 4.7.0 builds it (`tools/ktx2ref.cpp`, g++): each file's levels in a format, their FNV-1a hashes |
| `python -I tools\lan_test.py [--no-ai]` | cluster J's milestone: two Warzone.exe on 127.0.0.1, the host (--autohost, two NullBots) and the client joining through the game browser's IP box; the checks and the game time played together, no desync |
| `python -I tools\zlib_test.py`, `crypto_test.py`, `net_test.py`, `http_test.py` | miniz's streams against zlib, the crypto against libsodium (PyNaCl), WzNet's connections (loopback, a Python peer), urltransfer_c against a local HTTP server on 127.0.0.1 and ::1 (the address family, curl's request bytes, bodies, redirects, limits, failures); in `run_tests.py --media`. `http_test.py --external`: a few read-only HTTPS GETs (the real lobby's games list per family, badssl.com's bad certificates refused) |
| `python -I tools\lobby_test.py [--exe PATH] [host] [join] [down] [real]` | the internet lobby against a local mock lobby (127.0.0.1 and ::1): a host's listing through its lifecycle, a client joining through the browser and a captcha, each request over the family 4.7.0 forces; `down`: the lobby unreachable; `real`: one read of the real lobby's games list (nothing hosted) |
| `python -I tools\portmap_test.py [--exe PATH] [pcp natpmp upnp fail timeout off]` | port mapping (libplum) against a fake router on 127.0.0.1 only (the TEST-ONLY gateway override; the run is refused without it): mapping, unmapping, 4.7.0's failure and timeout messages |
| `python -I tools\sS1_test.py`, `lobby_commands_test.py`, `guide_test.py`, `notifications_test.py`, `s5_test.py`, `s6_display_test.py` | cluster S's screens on bin\ (section 39): the in-game multiplayer menu and the limits screen; lobby slash commands, admins, kicks, mute, known players (two copies); the Game Guide (skirmish, campaign pop-up, save/load); notifications, "do not show again", votes, the help overlay; the replay controller, spectators, the blind waiting room, `download`, `mod` and `ready` (a map pack and a mod downloaded by a second copy, the Ready box); the display options with the TEST-ONLY `WZ_TEST_NO_DESKTOP_EFFECTS` (nothing on the desktop changes) |
| `python -I tools\map_script_test.py` | script-generated maps (WzMap's map_script_c over chakra.dll) against 4.7.0's map_script.cpp with its pinned QuickJS (g++, ref/quickjs-wz): the maps' data byte for byte over seeds, the scripts' Math against ucrtbase; in `run_tests.py --media`. mapref and statsref link the same QuickJS |
| `python -I tools\gen_cursors.py` | writes `src/platform/cursors_win_data.bas`, the monochrome cursors of lib/sdl/cursors_sdl.cpp |
| `python -I tools\private_build.py <name> <projects> --stage DIR [--with DIR:users] [--run ARGS]` | a build of src\ plus a staged folder into build\j_<name>\bin, private type libraries, nothing of the tree touched: parallel forks' builds (section 37); src\shared's folders compiled as build.py does, the stage's files in them |
| `python -I tools\package.py [--zip] [--wz] [--no-source]` | the game to give to other people: dist\WarzoneVB-<date>\ (bin, data, Play.cmd keeping the settings beside the game, README, licences, the source zip); checks that the copy reaches the main menu |
| `python -I tools\video_display_test.py [--official]` | the videos' display against a model of sequence.cpp's: 9 video sizes and scanline modes and the intelligence screen's research film (about 12 minutes); `--official` the official build's captures too |
| `python -I tools\ui_test.py music_ingame.txt 420 --sound --debug=sound` | the Music Manager over a skirmish, the OST's last track played to its end (the next album's first follows); music_manager.txt and music_manager_check.txt (`--config-of=music_manager`): the title screen's, a mode unticked and kept |
| `powershell -NoProfile -ExecutionPolicy Bypass -File tools\linkwatch.ps1 [-hours 1]` | records the command line of each LINK.EXE that starts meanwhile (a WMI creation event) (build/linkwatch/: its length against 32,767, its .OBJ count) |
| `python -I tools\run_tests.py --media [--keep]` | cluster K's tests: zip_test, celt_test, silk_test (src/silktests), opus_test, theora_test, video_audio_test (each also alone: `python -I tools\<name>.py`) |
| `python -I tools\ui_test.py campaign2_walk.txt 3000` | campaign 2 (and campaign3_walk.txt: 3) from the selector's Start at to its end, videos waited out (`tools\make_campaign_walk.py 2`); campaign_mods.txt: Reclamation from its .wz; quickchat.txt (`--debug=console`) |
| `python -I tools\statsref_stubs.py` | statsref's own stand-ins whose VB port is now real and not skipped in the vbtest runs (`vbDisplayOff`): what to gate when statsref's outputs stop matching after a cluster ports such a function |
| `python -I tools\name_budget.py ref\warzone2100` | VB6 name budget of the C++ per area and per project |
| `python -I tools\typelib.py bin\WzGfx.dll` | a DLL's type library and class GUIDs, without the registry |
| `python -I tools\net_coupling.py ref\warzone2100` | what netplay uses from the game (INetGame's methods) and the game from netplay |
| `python -I tools\probe\run_probes.py tools\probe ref\warzone2100` | Chakra and OpenGL probes |
| `python -I tools\probe\names_gen.py <work dir> <names> <modules> ...` | VB6 name-table and module-count limits |
| `python -I tools\probe\xproj.py <work dir>` | exe and ActiveX DLL probe |

## Environment

* VB6: `C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE`
  (`/make`); resource compiler `VB98\Wizards\RC.EXE`
* Python 3.12, 64-bit (`python -I`); git 2.30
* QuakeVB (tools, conventions, mixer, window and input layers):
  `C:\Users\Owner\Desktop\githubprojects\VBQuake`
* DoomVB: `C:\Users\Owner\Desktop\websites\doom`
* Windows 10 Pro 19045; `SysWOW64` has `chakra.dll`, `opengl32.dll`,
  `gdiplus.dll`; GPU NVIDIA GeForce RTX 3080 (GL 4.6)
* Scripted file edits: Git Bash's `sed -i` drops CRLF (`build.py` puts it
  back), and bash heredocs turn `\n` and `\\` in Python strings into real
  characters: write such files with the editor, not a heredoc.
