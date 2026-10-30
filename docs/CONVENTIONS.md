# WarzoneVB porting conventions

How the C++ of Warzone 2100 4.7.0 (`ref/warzone2100`) is translated to VB6.
It builds on QuakeVB's conventions
(`C:\Users\Owner\Desktop\githubprojects\VBQuake\docs\CONVENTIONS.md`),
whose VB6 pitfalls all apply and are repeated below in short. What is new
here: C++ constructs, the seven VB6 projects, and the rules found while
setting them up. Items marked *(decide in cluster N)* are still open.

## Structure

* One `.bas` module per C++ file, same name (`lib/framework/debug.cpp` ->
  `src/framework/debug_c.bas`). The module name gets `_c` when it would
  clash with a procedure, a VB keyword or object, or a name another project
  exports (`debug` is VB's `Debug` object, `main` is `Sub Main`).
* VB-only modules (no C++ counterpart) end in `_win` when they implement a
  library's API on Win32 (`physfs_win` for PhysFS, `stdio_win` for C's
  stdio, `main_win` for SDL), and in `_vb` for support code (`printf_vb`,
  `netgame_vb`, `exemain_vb`).
* `src/shared` holds the modules every project compiles in, like a
  statically linked C runtime: the C library replacements and small
  framework modules without state (`wztime_c`). Each project gets its own
  copy, so their state (stdio's open files) never crosses a project, and
  their names cost nothing at the DLL interfaces. Keep it small: each name
  in it counts in all seven projects.
* Each module starts with a comment naming its C++ source and, for a
  partial port, what is still missing.
* A function needed before its cluster is ported gets a stand-in in the
  project's `stubs_<project>.bas`, commented with the C++ function it stands
  for and the cluster that replaces it. A stand-in keeps the real name and
  signature, so callers need no change when the real port arrives.
  Exception: a stand-in whose signature needs a type its DLL has not
  designed yet (`iV_Image`, `IMAGEFILE`, `TRACK`; `modelGet` was one until
  A4) waits in the caller's stubs module, so that the
  DLL's interface does not grow a guess that would later need
  `--break-compat`.
* Shared types, constants and globals from a header live in the module of
  the `.cpp` that defines them; a header without a `.cpp` gets its own
  module (`input.h` -> `input_h.bas`) when it holds more than enums.

### Which project

| Project | C++ | Uses |
|---|---|---|
| `WzFramework.dll` | `lib/framework`, the window-free part of `wzapp.h` (ticks, dialogs, threads), `src/wzjsonhelpers` | - |
| `WzMap.dll` | `lib/wzmaplib` | - |
| `WzSound.dll` | `lib/sound` | WzFramework |
| `WzGfx.dll` | `lib/ivis_opengl` | WzFramework |
| `WzNet.dll` | `lib/netplay` | WzFramework |
| `WzPlatform.dll` | `lib/sdl` (window, input, event loop, GL backend) | WzFramework, WzGfx |
| `Warzone.exe` | `src`, `lib/gamelib`, `lib/widget` | all |

A project only calls the projects it uses. Calls the other way go through
a callback interface (below).

## Between projects

* Each DLL has one **GlobalMultiUse API class**, `<Name>Api.cls`
  (`FrameworkApi`, `PlatformApi`, ...). Its public procedures look like
  globals to the projects that use the DLL. They forward to the DLL's
  modules with the module name in front (`debug_c.debug_init`), because
  inside the class an unqualified call would reach the class's own method.
* Public **enums and constants** that other projects need are declared in
  the API class (`Public Enum code_part`); they go into the type library.
* **Callback interfaces**: `I<Name>.cls`, PublicNotCreatable, in the lower
  DLL (`IPlatformGame` in WzPlatform, `IGLBackend` in WzGfx). The upper
  project implements it in a private class (`PlatformGame.cls` with
  `Implements IPlatformGame`) and hands an instance over at startup. Before
  the process ends, hand over `Nothing` again. `IFrameworkGame` (WzFramework,
  implemented by `FrameworkGame.cls`, registered with `frameSetGame`) carries
  the resource module's function pointers: the game registers numbers
  (`resAddFileLoad("SBODY", fn_bufferSBODYLoad, fn_dataReleaseStats)`) and
  the framework calls them back by number; `void *pData` is a Variant.
* **No underscores in the method names of a callback interface**: VB
  refuses them for `Implements` ("Bad interface for Implements: method has
  underscore in name"). Drop each underscore and capitalise the next letter:
  `war_GetWidth` -> `warGetWidth`, `version_getVersionString` ->
  `versionGetVersionString`.
* Names a project uses from a DLL count in its own name table (measured), so
  APIs stay narrow. `tools/name_budget.py` keeps the count.
* A `ParamArray` is passed on as a Variant (`Dim a As Variant: a = args`).
* `Err.Raise` in a DLL reaches the caller with number, source and
  description.
* A DLL's global that other projects use is a `Property Get`/`Let` pair in
  the API class (`selectedPlayer`, `wz_texture_compression`,
  `netGameserverPortOverride`); the fields of a global struct are
  properties named `<variable>_<field>` (`NetPlay.isHost` ->
  `NetPlay_isHost`), with the index first for an array of structs
  (`NetPlay.players[0].allocated` -> `NetPlay_players_allocated(0)`).
  A constant other projects need goes into a public enum of the API class
  (`MAX_PLAYERS` in `frame_h_players`).
* A struct passed or returned across a DLL boundary travels as its fields
  in ByRef parameters (`NETgetLobbyHostJoinOptions botProt, proxyIPs,
  hostingIPs`), or as a Long when the C++ packs it into 32 bits itself
  (`SpectatorInfo_currentNetPlayState()` returns `toUint32()`).
* Data both sides read in bulk (object pools, map tiles) belongs to the
  side that writes it; the other side sees it through a SAFEARRAY alias
  (QuakeVB's `SetView`), with the UDT declared in both projects.
* DLL interfaces grow by adding; a change VB6 calls incompatible needs
  `build.py --break-compat NAME`.
* **A DLL's public enums all go in its API class**, even one C++ nests in
  a class (`Map::LoadedFormat` is MapApi's `LoadedFormat`). A public class
  that uses the API class's enums while the API class uses one of its own
  (even as a default value) is "Circular dependencies between modules".
  For the same reason the API class calls no `Friend` member of such a
  class: it forwards to a module that does (`MapApi.Map_new` ->
  `map_c.Map_new` -> `Map.construct_empty`).
* A C++ library's own classes keep their names as public classes of its
  DLL (wzmaplib's `Map`, `MapData`, `Structure`), and its abstract classes
  become interfaces under their C++ names (`IOProvider`, `BinaryIOStream`,
  `LoggingProtocol`), which other projects implement. A virtual method
  with a body in the base class that no subclass overrides becomes a
  module procedure taking the object first (`BinaryIOStream_readULE16(s,
  v)`); a reserved name loses its underscore there too, so it gets a
  word instead (`close()` -> `closeStream`).
* An `Optional` argument the C++ defaults to a new object
  (`std::make_shared<StdIOProvider>()`) is `Optional ByVal x As Variant`,
  tested with `IsMissing`, so that `Nothing` can still mean `nullptr`.

## Names

* C++ names keep their spelling. VB is case-insensitive: when two C++ names
  differ only in case, the one used in more places keeps it and the other
  gets a trailing underscore (QuakeVB's rule).
* A leading underscore moves to the end: `_debug` -> `debug_`,
  `_syncDebug` -> `syncDebug_`.
* Namespace and class qualifiers join with underscores:
  `gfx_api::context::initialize` -> `gfx_api_context_initialize`. A method
  ported as a module procedure takes its object first:
  `psDroid->isVtol()` -> `DROID_isVtol(psDroid)`.
* Reserved and awkward names (QuakeVB's list plus `debug`): `scale`, `line`,
  `name`, `type`, `next`, `step`, `len`, `left`, `right`, `time`, `date`,
  `point`, `dir`, `str`, `print`, `open`, `close`, `get`, `put`, `me`,
  `Empty`, `string`, `width`/`height` as globals, `debug`, `erase` (VB's
  `Erase` statement; a method named so is a syntax error). Fields: `type` ->
  `typ`, `next` -> `nxt`, `string` -> `strval`; others get a trailing
  underscore (`line_`, `time_`). The same goes for methods (`erase_`,
  `empty_`, `string_`, `set_`). `check_names.py` knows the list.
* A local or parameter with the name of a global in any case hides it
  silently: rename the local (`check_names.py` finds them).

## C++ constructs

| C++ | VB |
|---|---|
| `struct` of plain data | UDT |
| `DROID`, `STRUCTURE`, `FEATURE`, `PROJECTILE` (`BASE_OBJECT`) | UDT pools; a pointer is a Long handle, 0 = NULL, that is both the typed and the base pointer; `virtual` becomes `Select Case` on the handle's pool (see "Objects") |
| `std::list<T *>` of objects | `OBJ_LIST` (`objlist_vb`): head, tail and size, the nodes in a shared pool; a node index is an iterator |
| class hierarchies outside the simulation (widgets) | class modules: `Implements` an interface, base behaviour by containment; VB's reference counting stands in for `shared_ptr` |
| lambda in an algorithm (`std::find_if`, `std::sort`) | a loop, or a named procedure |
| lambda or `std::function` stored as a callback | a dispatch number with its captured values in a UDT, or a small handler class |
| function pointer (`debug_callback_fn`) | a number and a `Select Case` (QuakeVB's rule) |
| `std::vector<T>` | dynamic array plus a count; `std::vector<std::string>` is `StringVector` (`src/shared/vector_vb.bas`: `push_back`, `erase`, `clear`, `contains`, `sort`, `unique`) |
| `continue` in a loop | `GoTo continue_` with the label before `Next`/`Loop`, or an `If` around the rest |
| `exit(n)` | `exit_ n` (`cstdlib_vb`): ends the process at once, as C does; never `Err.Raise` |
| `printf(...)`, `fprintf(stdout, ...)` | `fprintf STDOUT_FILE, ...` (`stdio_win`); `freopen` redirects stdout or stderr for every project |
| a `static inline` function of a header used by several projects | a copy in each project that needs it (`to_swap_mode`), or the header's module in `src/shared` (`string_ext_h.bas`) |
| `std::map`/`unordered_map` with Long keys | hash module *(decide in cluster A)* |
| `std::map`/`set` with string keys | `Scripting.Dictionary` (late bound, `CreateObject`, `BinaryCompare`) off the hot paths |
| `std::string`, `WzString`, `char*` text | VB String (UTF-16). Files and the network are UTF-8: convert at the boundary (`stdio_win`, `physfs_win`) |
| a `char*` that may be `nullptr` | `vbNullString` (`StrPtr(s) = 0`), as `LEVEL_DATASET.realFileName`; `printf` of it prints `(null)` as the C runtime does |
| `char buf[N]` filled by `snprintf` | a String cut to `N - 1` characters where C would truncate |
| templates | one copy per type used; over classes with the same members, one late-bound copy (`mapObj As Object` in wzmaplib's `jsonGetBaseMapObjectInfo<T>` for `Structure`, `Droid`, `Feature`); over integer types, the type as a parameter (`jsonGetListOfType<T>`: T = 32 or 16) |
| `throw`, `catch` | `Err.Raise`, `On Error GoTo`, or return codes. `catch (const std::exception& e)` around JSON code is the handler for `Err.Source = "nlohmann::json"` (`e.what()` is `Err.Description`); any other error is `catch (...)` |
| a class constructed with arguments (`WzConfig cfg(name, w)`) | a class with a factory `<Class>_new(args)` (`WzConfig_new`, `json_variant_new`) that calls its Friend `construct`/`init`; the destructor is `Class_Terminate`, which runs when the last reference goes, as at the end of the C++ scope |
| `std::stoi`, `stoul`, `stoull`, `stod`, `stof` | `std_stoi`... (`cstdlib_vb`): False and the `what()` text where C++ throws |
| `optional<T>` | the value plus a Boolean `has<Name>`; a function returning one returns the Boolean (has_value) and sets the value ByRef; an optional enum with no 0 member can use 0 for `nullopt` (`optional<MOUSE_KEY_CODE>`) |
| global with an initialiser (`bool showFPS = false; int x = 20;`) | a VB module has no initialisers: in a DLL the module sets them on its first call (`InitVolumes`); in the exe, `stubs_game_init` and the ported modules' `<module>_init` (`structure_c_init`: a null stats pointer is -1, not 0) run first in `Sub Main` |
| `std::stoi`, `std::stof` | `strtoll`/`strtod` (`cstdlib_vb`) with C++'s checks: no digits is `invalid_argument`, out of range is `out_of_range`, with MSVC's `what()` texts |
| overloads that differ by their parameters' types | the one used more keeps the name, the other gets a suffix saying what it returns or takes: `getWeaponSubClass(WEAPON_SUBCLASS)` is `getWeaponSubClass_name`, `loadTemplateCommon(WzConfig&, ...)` is `loadTemplateCommon_ini` |
| `WzString::toLower`, `trimmed`, `split` | `WzString_toLower` (A-Z only, as the classic locale), `WzString_trimmed` (" \f\n\r\t\v"), `WzString_split` (`src/shared/wzstring_c`); VB's `LCase$` and `Trim$` do something else |
| `rand()`, `srand()` | `cstdlib_vb`'s, the Microsoft runtime's generator; each project has its own state |
| `uint32_t` multiplication (`a * b` wrapping) | `MUL32(a, b)` (`math_vb`) |
| `MersenneTwister`, `gameRand` | `random_c` (a UDT and procedures that take it) |
| `int32_t`, `int`, `unsigned` | Long (unsigned compares and divides through QuakeVB's `ULT`/`UDIV` helpers) |
| `int64_t`, `uint64_t` | Double (exact to 2^53) for times, sizes and power's 32.32 fixed point; a division is `Fix(a / b)` (exact for \|a\| < 2^53); `%lld` / `PRId64` print such a Double in full. Checked case by case |
| a reference to a static vector returned by a function (`GridList const &gridStartIterate(...)`) | the elements copied into the caller's array, their number returned: `n = gridStartIterate(x, y, r, list)` |
| `float` | Single; `double` stays Double. Where the result must match C's bit for bit, each float operation is a `float_vb` call (see "Float arithmetic") |
| `sscanf` | `sscanf` (`src/shared/sscanf_vb`): the variables after the format are passed ByRef through a ParamArray; returns C's count, or -1 (EOF) |
| `std::unique_ptr<T>` to a pooled object, and chains of them (`iIMDShape::next`) | a handle; the owner frees it and what it chains to (`deleteShape`); `std::move` is the handle's copy and a 0 left behind |
| `bool` | Boolean |
| threads, mutexes, `wzAsyncExecOnMainThread` | none: the work runs on the one thread *(decide per use)* |
| `#if defined(WZ_OS_WIN)` | the Windows branch; other platforms are dropped |
| `#ifdef DEBUG` | release behaviour (DEBUG undefined), like an official build |

## Logging and asserts

`__LINE__` and `__FUNCTION__` do not exist in VB, so the call sites carry
the C++ values: **the line number of the call in the 4.7.0 source and the
function name**. Logs then read like Warzone's, and syncDebug output (which
prints function names) can be compared with the C++ build's.

| C++ | VB |
|---|---|
| `debug(LOG_WZ, "Write dir: %s", d);` (line 512 of `initialize_ConfigDir`) | `debug_ 512, LOG_WZ, "initialize_ConfigDir", "Write dir: %s", d` |
| `wz_info("...")` | `wz_info 1987, "realmain", "..."` |
| `ASSERT(x > 0, "bad %d", x);` | `If Not (x > 0) Then ASSERT_FAILURE "x > 0", "droid.cpp:123", "droidUpdate", "bad %d", x` |
| `ASSERT_OR_RETURN(false, p != nullptr, "...");` | `If Not (p <> 0) Then ASSERT_FAILURE "p != nullptr", "map.cpp:45", "mapLoad", "...": mapLoad = False: Exit Function` |

* `ASSERT_HELPER(expr, location, function, ...)` used directly with a
  location and function passed on (checkDroid, checkStructure,
  checkProjectile, checkObject for their `CHECK_*` callers) logs at the
  `__LINE__` of the `ASSERT_HELPER` call, not the line in the location:
  `ASSERT_HELPER_FAILURE 3909, "droid->asWeaps[i].lastFired <= gameTime",
  location, function_, "CHECK_DROID: ..."`. `check_debug_lines.py` checks
  that the line holds an `ASSERT_HELPER` with that expression and format.
* `debug_` checks `enabled_debug(part)` itself, like the macro. VB evaluates
  the arguments first, though, so on a hot path wrap the call in
  `If debugPartEnabled(part) Then`; the same when an argument has side effects
  (`clearPlayer`'s `getPlayerName` asserts for the slots past MAX_PLAYERS).
* The expression string in `ASSERT_FAILURE` is the C++ text as the
  preprocessor stringifies it: `ASSERT` hands its expression on to
  `ASSERT_HELPER`, so its macros are expanded (`ASSERT(disabledWhen <=
  MPFLAGS_MAX, ...)` prints `disabledWhen <= 0x3f`, and `strcasecmp` is
  MSVC's `_stricmp`), while `ASSERT_OR_RETURN` prints its expression as
  written. `check_debug_lines.py` checks both.
* A log call inside a C++ macro (`ALLOC_STATS`, `ASSERT_PLAYER_OR_RETURN`)
  carries the line of the macro's use; `check_debug_lines.py` expands the
  file's function-like macros to check it.
* The function name is the unqualified one, as GCC's `__FUNCTION__` gives
  it: `beginGroup`, `WzConfig` for the constructor, `~WzConfig` for the
  destructor (Microsoft's prefixes the class). A lambda logs under the
  function that contains it.
* `syncDebug(...)` is `syncDebug "<function>", fmt, args` (WzNet's
  `NetApi.syncDebug` formats and logs it as `_syncDebug`);
  `_syncDebugIntList` is `syncDebugIntList_ "<function>", fmt, list,
  count`, and `syncDebugStructure(p, ch)` is `syncDebugStructure_
  "<function>", p, ch` (likewise `syncDebugFeature_`). The function names
  are part of the sync log's CRC, so they must be the C++ ones.
* wzmaplib's `debug(logger, level, ...)` expands to `_printLog(logger,
  __LINE__, level, __FUNCTION__, ...)`: in VB `printLog_ pCustomLogger, 68,
  LogLevel_Info, "loadMapData_Internal", "Loading: %s", path`, in the
  C++ argument order. `check_debug_lines.py` checks these too.

## JSON (nlohmann::json)

* A `json` / `nlohmann::json` / `ordered_json` value is an `nlohmann_json`
  object (WzFramework, `json_hpp`). Make one with `json_parse(text)` or
  `json_parse_bytes(data, first, size)` (`json::parse`), `json_object()`,
  `json_array()`, `json_value(v)` (`json j = v;`), each with `ordered:=True`
  for an `ordered_json`.
* A C++ json is a value, the VB one a reference. Copy where the C++ copies
  and either side changes afterwards: `json a = b;` is `Set a = b.clone()`;
  `obj[key] = v;` is `obj.set_item key, v` (copies), `arr.push_back(v)` is
  `arr.push_back v` (copies). For a temporary or a `std::move`, the `_ref`
  forms take the object as it is: `obj.set_item_ref key, json_value(5)`.
  `*p = v;` through a pointer or reference is `p.assign v`.
* `json&` and pointers into a value (`auto& x = obj["k"]`) are references
  to the same object: `Set x = obj.item("k")`.
* Overloads: `at(key)` / `at_idx(i)`, `item(key)` / `item_idx(i)`
  (`operator[]`, inserting like the non-const one), `get<T>()` is
  `get_int`, `get_uint`, `get_int64`, `get_uint64`, `get_double`,
  `get_float`, `get_bool`, `get_string`; `value(key, def)` is `value_int`,
  `value_string`, ...; `find(key)` returns the value or Nothing (`end()`);
  `erase(key)` is `erase_`, `empty()` is `empty_`, `type()` is `type_`.
* Iteration: `For i = 0 To j.size() - 1` with `j.key_at(i)` and
  `j.value_at(i)`, in the library's order (`json`: keys sorted by their
  bytes; `ordered_json`: as inserted; a primitive iterates once, null not at
  all).
* The library's exceptions are VB errors: `Err.Number = vbObjectError + id`
  (101 for parse errors, 3xx type errors, 4xx out of range),
  `Err.Description` the `what()` text, `Err.Source = "nlohmann::json"`.

## Stats (statsdef.h and the stats stores)

* Each stats type is one UDT with its base classes' fields first
  (`BASE_STATS`: `id`, `name_`, `ref`, `index`; `COMPONENT_STATS`: `pIMD`,
  `buildPower`, ...), so `psStats->buildPower` is
  `asWeaponStats(i).buildPower`. The upgrade structs are named after their
  class (`WEAPON_STATS::UPGRADE` is `WEAPON_STATS_UPGRADE`, with
  `COMPONENT_STATS::UPGRADE`'s fields first). A field a derived class
  declares again hides the base's: `RESEARCH`'s own `index` is `index`,
  `BASE_STATS::index` is `base_index`.
* A `std::vector<X_STATS>` store is the array plus `<name>_size`
  (`asWeaponStats`, `asWeaponStats_size`); the C++ default member
  initialisers are `<TYPE>_construct`, which the allocations call.
* A typed pointer into a store (`WEAPON_STATS *`, `STRUCTURE_STATS *`,
  `RESEARCH *`) is the index, **-1 = nullptr** (the stores are 0-based, as
  the C++ indexes them). Pointers to dynamically made objects are 1-based
  pool indices, 0 = nullptr: `DROID_TEMPLATE *` (template_c's pool
  `apsTemplates`), and model handles (`iIMDBaseShape *`, `iIMDShape *`:
  WzGfx's pools, read from the other projects through GfxApi's
  `iIMDBaseShape_<field>`).
* A `BASE_STATS *` or `COMPONENT_STATS *` is a handle: the stat's
  `StatType` plus its index (`STAT_WEAPON + i`; a template's is
  `STAT_TEMPLATE` plus its pool index), 0 = nullptr. For every stat but a
  template it equals the stat's `ref`. Its fields are the module Property
  procedures `BASE_STATS_<field>(h)` and `COMPONENT_STATS_<field>(h)`
  (`psStats->getBase().hitpoints` is `COMPONENT_STATS_getBase_hitpoints(h)`),
  which dispatch on `h And STAT_MASK`; `BASE_STATS_hasType` reads the `ref`.
* `apCompLists[player][type][comp]` is one array `apCompLists(player,
  type, comp)` whose last dimension is the longest list, with
  `apCompLists_count(type)` (0: nullptr).
* `getStatsName(psStats)` and the other macros of statsdef.h are written
  out (`asWeaponStats(i).name_`); translations are not ported.

## Objects (basedef.h and the pools)

* `DROID`, `STRUCTURE` and `FEATURE` are UDTs whose first fields are
  `SIMPLE_OBJECT`'s and `BASE_OBJECT`'s, in the C++ order
  (`SIMPLE_OBJECT::type` is `typ`, `time` is `time_`). Each lives in its own
  pool (`paged_entity_container_h`): `aDroid`, `aStruct`, `aFeature`. A
  pool's array starts at `type * OBJ_HANDLE_UNIT + 1` (`OBJ_HANDLE_UNIT` =
  `&H10000000`), so a handle indexes its pool directly and is at the same
  time the `DROID *` and the `BASE_OBJECT *`: `psDroid->pos.x` is
  `aDroid(psDroid).pos.x`. 0 is nullptr.
* Through a `BASE_OBJECT *` or `SIMPLE_OBJECT *` a field is read with the
  Property procedures of `basedef_h` (`SIMPLE_OBJECT_pos(psObj)`,
  `BASE_OBJECT_visible(psObj, player)`), which pick the pool from the
  handle (`h \ OBJ_HANDLE_UNIT`). `psObj->type` is `SIMPLE_OBJECT_type`,
  the field: the destructor changes it, as in C++.
* The containers (`PagedEntityContainer`, `ReuseSlots = false`) are
  `DroidContainer_alloc`, `_find`, `_erase`, `_begin`/`_next`, `_clear`: a
  slot is taken once, in order, and reused only after `clear()`, so a stale
  handle reads the dead object as C++ reads its page. The constructors are
  `FEATURE_construct`, `STRUCTURE_construct` (on a UDT: `STRUCTURE
  building(id, player)` is a local copied into the pool by
  `GlobalStructContainer_emplace`) and `DROID_construct_` (VB would take
  `DROID_construct` for `DROID_CONSTRUCT`); each one runs `SIMPLE_OBJECT`'s
  and `BASE_OBJECT`'s first. The destructors (`<TYPE>_destruct`) end with
  `~BASE_OBJECT` and `~SIMPLE_OBJECT`.
* `union FUNCTIONALITY` is a UDT inside `STRUCTURE` (`functionality`), its
  members side by side; `pFunctionality` is the structure's own handle when
  it has one, else 0 (`STRUCTURE_calloc_functionality` gives calloc's
  zeroes, -1 for a `RESEARCH *`). A `FUNCTIONALITY *` parameter is that
  handle.
* `FLAG_POSITION *` and `DROID_GROUP *` are 1-based handles into heaps with
  free lists (`aFlagPos`, `aGroup`), as calloc/free and unique_ptr.
* The object lists (`PerPlayerObjectLists`, `std::list<T *>`) are arrays of
  `OBJ_LIST` (`objlist_vb`); `mission.apsDroidLists` is
  `mission_apsDroidLists` (an array inside a UDT cannot be passed as one),
  and `&pList[p] == &apsDroidLists[p]` compares `VarPtr`.
* The exe's UDTs hide WzMap's classes of the same name (VB names ignore
  case): write `WzMap.Feature`, `WzMap.Structure`, `WzMap.Droid`.
* A local stats object that is not in a store (`findAssemblyPointPosition`'s
  `STRUCTURE_STATS sStats` with `ref = 0`) is the handle `DUMMY_STATS_REF0`:
  no StatType, its `ref` 0.
* `CHECK_STRUCTURE(p)` is `checkStructure p, "<file>:<line>", "<function>",
  max_check_object_recursion`, with the caller's location.

## The map (map.h)

* `MAPTILE *` is a tile index into `psMapTiles` (`mapTile(x, y)` returns
  it), so `psTile->height` is `psMapTiles(psTile).height`; the struct is
  `MAPTILE_` (`mapTile` keeps the name). `psMapTiles == nullptr` is
  `Not psMapTiles_allocated`.
* `psBlockMap[slot]` and `psAuxMap[player]` are 2-D Byte arrays,
  `psBlockMap(slot, tile)` and `psAuxMap(player, tile)`.
* A member pointer (`uint16_t MAPTILE::*varContinent`) is a selector
  parameter (0 = `limitedContinent`, 1 = `hoverContinent`).
* Text files read into `fileLoadBuffer` (`loadFileToBuffer`) are parsed as
  C parses them: the buffer as text up to its terminating zero, a `char *`
  being a 1-based position in it, `sscanf(Mid$(text, p), ...)`, and
  `strchr(p, '\n') + 1` a helper (`nextLine`). A scanset that names
  `\r\n` takes the real characters (`vbCr & vbLf`): `sscanf` has no
  escapes.

## printf and strings

* The printf family (`debug_`, `va`, `ssprintf`, ...) takes the C format
  string with its backslash escapes and a `ParamArray`. `vsprintf`
  (`printf_vb.bas`, QuakeVB's, checked against C) interprets `\n`, `\t`,
  `\r`, octal escapes and `%[-+ 0#][width][.prec][hlLzjt]` with
  `d i u x X o c s f e E g G p %`. A 64-bit value prints in full from a
  Decimal, or from a Double with `ll` or `j`. Only the printf family interprets
  escapes; other procedures take plain text (`vbLf` for a newline).
* VB string literals have no escapes: a quote inside is `""`, never `\"`.
* When C passes a variable as the format, write `"%s", text`.

## Float arithmetic (float_vb, glm_vb)

* VB evaluates a Single expression on the x87 with more precision than
  C's float, so a chain like `a * b + c` does not round as C does. Where the
  result must match (model loading, the simulation's float code such as
  `structureBodyBuilt`), write each float operation as a call: `fadd`,
  `fsub`, `fmul`, `fdiv`, `fsqrt` (`src/shared/float_vb`, in every
  project) do it in Double and round to Single through memory (`F32`),
  which gives C's correctly rounded float. A uint32_t converted to float is
  `F32(U32(x))`. `double`
  arithmetic is plain VB Double (the x87 at 53 bits rounds as SSE does).
* glm's vector functions (WzGfx's `glm_vb`) are written with glm's formulas: `v3dot` is
  `(x*x + y*y) + z*z`, `v3normalize` is `v * (1 / sqrt(dot(v, v)))`,
  `v3cross` is glm's (Warzone does not force glm's SIMD code).
* `(int)f` truncates: `CLng(Fix(f))` (or `ToInt` for C's out-of-range
  result).
* **Infinities and NaNs.** VB's `/` raises on a division by zero: `fdiv`
  gives C's infinity or NaN. VB's Double to Single conversion raises an
  overflow where C gives an infinity: `F32` builds infinities and NaNs
  from their bits. An operation that makes a NaN (`0 * inf`, `inf - inf`)
  and a comparison with a NaN set the x87's sticky "invalid" flag, and
  VB's next float to integer conversion, anywhere (`2 ^ n` in `BIT_`, a
  `CLng`), then fails with error 6: `F32` clears the flags (msvcrt's
  `_clearfp`) whenever its result is an infinity or a NaN, and code that
  can see a NaN tests `IsNaNF` before it compares (C: every comparison
  with a NaN is false but `!=`; VB's compiled ones are not).
* The test dumps print a float as its bits (`FloatHex`, `%08X`).

## VB6 pitfalls (from QuakeVB; each one bit DoomVB or QuakeVB)

QuakeVB's numeric helpers named here (`ToInt`, `SAR`, `SHL`, `SHRU`,
`BIT_`, `ULT`, `U32`) are in `src/shared/math_vb.bas`; QuakeVB's
`Sys_ClearFPU` is `float_vb`'s `F32` here (see "Float arithmetic").

* **Integer literals are 16-bit**: `10240 * 4` overflows; write `10240& * 4`,
  `Const W As Long`. Hex `&H8000` to `&HFFFF` are negative Integers: use
  `&H8000&`, `&HFFFF&`.
* **Precedence**: comparisons bind tighter than `And`/`Or`: write
  `If (flags And F) <> 0`. `\` binds looser than `*` and `/`: C's
  `a*b/c*d` is `(a * b \ c) * d` (`check_intdiv.py`). `Mod` binds looser
  than `*`, `/` and `\`: C's `a % b * c` is `(a Mod b) * c`
  (`check_mod.py`).
* **No short-circuit**: `If i >= 0 And a(i) = 1` evaluates `a(-1)`. Nest Ifs.
* `/` is floating point; C's integer `/` is `\` (truncates like C).
* Shifts: `x >> n` on a possibly negative value is `SAR(x, n)`; `1 << n`
  is `BIT_(n)`.
* `Not` is bitwise: for a Long used as a boolean write `x = 0`.
* C `(int)` truncates, VB `CLng` rounds to even: use `ToInt` (`Fix` with
  C's out-of-range result); C `floor` is `Int`.
* A float to Long conversion out of range raises error 6; a NaN poisons the
  next conversion (see "Float arithmetic").
* `IIf` evaluates both branches.
* `^` is a floating point power (and a conversion back to Long): `2 ^ n`
  in a helper called per byte made the sync log's CRC the slowest part of
  the game tick. `BIT_` is a table; keep `^` out of hot code.
* Never `ReDim` an array inside a `With` on one of its elements, or while an
  element is passed ByRef; never `GoTo` out of a `With` on an array element
  (`check_with.py`).
* `Not Not arr` upsets the FPU stack: keep an "allocated" flag instead.
* **64 KB of fixed-size data per module**: big arrays are dynamic.
* Module-level declarations come before the first procedure.
* Executing `On Error` clears `Err`: read `Err.Number` first.
* Run-time errors in a compiled exe open a modal dialog: hidden test runs
  must trap them.
* **No engine code inside Windows callbacks** (window procedure, Chakra
  native functions): queue the event, or trap every error there. A VB error
  cannot unwind through a callback.
* Pass scalars `ByVal`; UDTs and arrays go ByRef.
* **A class's Public String field loses vbNullString**: assigned from
  outside the class, `obj.s = vbNullString` stores "" (VB's implicit
  Property Let copies the string). Locals, UDT members, parameters, return
  values and explicit Property Let/Get pairs keep the null. A string that
  stands for a `char *` that may be nullptr is a private field behind a
  Property pair (`LEVEL_DATASET.realFileName`).
* `memset` of a struct is the assignment of a blank UDT.

## Rules found in this port

* An interface used with `Implements` cannot have underscores in its method
  names (see above).
* **The embedded manifest must be padded**: a resource whose size is not a
  multiple of 4 gets NUL bytes, and Windows refuses to start the exe
  ("Activation context generation failed ... Invalid Xml syntax" in the
  Application event log, source SideBySide; "Accessing a corrupted shared
  library" in bash). `build.py` pads it with spaces.
* VB6 stops without a message when a project's name table is full (about
  32,100 names): an empty build log. `build.py` says so.
* **Binary compatibility needs a fresh compat copy after an interface
  grows**: until the copy has the new members, VB gives them (and new public
  enums) fresh GUIDs on every build, and a project built against the last
  ones fails ("The binary compatibility DLL or EXE contains a parameter type
  or return type whose definition can not be found"). `build.py` refreshes
  the copy whenever the type library version changes.
* `check_debug_lines.py` finds a module's C++ files in its header comment
  (before `Option Explicit`): name every C++ file whose log calls the module
  carries (`main_win` names `gfx_api.h` for `to_swap_mode`).
* A log line with no C++ counterpart carries line 0 and starts its text
  with `(VB)`; `check_debug_lines.py` skips it.
* A project's module count times its folder path length must stay under
  LINK.EXE's 32,767-character command line.
* A module's `Property Get`/`Let` pair shares one name
  (`BASE_STATS_id`); `check_names.py` knows it.
* A procedure cannot be named `put` (VB's `Put #`); `vbtest_vb` uses
  `putOut`.
* A module name must differ from its project's name.
* **On a hot path, count the procedure calls.** Each VB6 call sets up an
  exception frame, and passing an array element ByRef locks the array for
  the call. In an inner loop (the A* in `astar_c`), pass the arrays as
  parameters, keep a struct's fields in parallel arrays, and write small
  helpers out in place (with the C++ function kept as the reference);
  measure with `prof_vb` (`profBegin`/`profEnd`, `profWrite`) and remove
  the calls after.
* `trig_c`'s `i64Sqrt` takes a Variant for 64-bit values; `iHypot` and other
  callers with a Double below 2^53 use its Variant-free path.
* `int` arithmetic that can overflow in C++ (and then wraps, as on x86)
  must not raise VB's overflow: `vector_h`'s `dot2i` sums in Double and
  wraps with `FromU32`; an `int64_t` result narrowed to `int32_t` is
  `FromU32(Fix(...))`.
* `CHECK_DROID(psDroid)` is `checkDroid psDroid, "src/<file>.cpp:<line>",
  "<function>", max_check_object_recursion`, as CHECK_STRUCTURE. A direct
  `ASSERT_HELPER(expr, ...)` stringifies its expression unexpanded
  (`droid->player < MAX_PLAYERS`), unlike `ASSERT`.
* A virtual class of the simulation that is not an object (the steering
  behaviours) is a number and a `Select Case` (`steering_c`'s
  `STEERING_BEHAVIOR_...`), as a function pointer is: rayCast's callbacks
  (`RAY_CALLBACK_...`), spiralSearch's match functions, the AUDIO_CALLBACKs
  (`CB_...`). A callback's `void *data` is its module's own variables.
* A `switch` with fallthroughs (moveUpdateDroid) is an If chain with flags
  for the cases that fall through, and `GoTo` for a `break` inside a case.
* gtime.h's overloads `gameTimeAdjustedAverage(int)` and
  `(int, int)` are one function with an Optional denominator; the Vector3i
  forms are `_vec`.
* A pointer computed back from another (`psWeapon - psDroid->asWeaps`, the
  weapon's slot) is passed as what it stands for: actionTargetTurret takes
  the weapon's slot. A REPAIR_FACILITY * is its structure's handle (the
  facility is `aStruct(h).functionality.repairFacility`).
* A STRUCTURE_STATS * passed where C++ converts it to BASE_STATS * is
  `structureStatsBase(index)` (`structuredef_h`).
* The overloads of actionDroid: `actionDroid` (action), `actionDroid_loc`
  (x, y), `actionDroid_obj` (psObj), `actionDroid_objLoc` (psObj, x, y); the
  DroidOrder constructors are `DroidOrder_`, `_pos`, `_posRtr`, `_stats`,
  `_stats2`, `_obj`, `_objRtr`, `_objIndex`.
* A `std::priority_queue` whose only use is its top is a linear search for
  the greatest element by its comparator (findSomeoneToRepair).
* C++ members a constructor leaves uninitialised are 0 in VB (a fresh UDT);
  a test that prints one sets it on both sides first (DROID::actionStarted).
* A C `&&`/`||` whose right side logs (a function with `CHECK_DROID`, such
  as `isVtolRearming` or `vtolHappy`) is nested Ifs even when the right side
  has no other effect: VB's `And` would log where C++ does not (the logs are
  compared).
* Iterating a `std::unordered_map` gives the library's hash order (libstdc++'s
  in statsref, Microsoft's in the official build): where only a log shows it
  (eventResearchedHandleUpgrades' grade lines), the port keeps first-appearance
  order and run_tests compares those lines sorted. Where the order reaches the
  game state, it must be emulated.
* A C++ worker thread whose work only reads and writes copies made before it
  is woken (the danger maps' dangerFloodFill) is a plain call where the
  thread is started or woken; waiting for it is nothing.
* `MAX`/`MIN` (macros.h) and `std::max` of an int and an unsigned compare and
  pick as unsigned: a negative int then wins a MAX. objGuessFutureDamage's
  `MAX(actualDamage - armour, actualDamage * minimumDamage / 100)` (with an
  unsigned minimumDamage) is a ULT test; look at both operands' types.
* A nullable out-parameter (`TARGET_ORIGIN *targetOrigin`, nullptr allowed)
  is an Optional Boolean saying whether C++ had a pointer plus an Optional
  ByRef value (aiChooseTarget).
* Projectiles are objects in the pool aProj (paged_entity_container_h's
  ProjectileContainer, reusing freed slots); a PROJECTILE handle works with
  basedef_h's SIMPLE_OBJECT_* properties. An aProj element must not be passed
  ByRef to code that can create a projectile (the pool may grow).

* A Collection's or Variant's item that is an object, read with a Let
  (`v = c.Item(k)`): VB asks the object for its default member (run-time
  error 438; a test of the key that traps errors then answers "missing").
  `Set v = c.Item(k)`, or `IsObject(c.Item(k))` to test the key.
* A fixed-size array (`Dim a(0 To 2) As Long`) passed to a procedure that
  ReDims or Erases its parameter: run-time error 10 ("This array is fixed or
  temporarily locked") when the call runs, not when it compiles.
  `tools\check_redim.py` finds them.

## Editing the sources by script

* The sources are CRLF; `build.py` normalises them before every build.
* Git Bash's `sed -i` drops the CR. Bash heredocs turn `\n` and `\\` inside
  Python strings into real characters, so write Python and VB with the
  editor (or Python reading and writing bytes), never through a heredoc.
* Python's text mode on Windows reads CRLF as `\n` and writes `\n` as CRLF:
  a script that adds `\r\n` itself leaves CR CR LF, which VB reads as blank
  lines and reports as a syntax error on an innocent line. Use `newline=""`
  (or bytes) when the script handles line ends.
* Encode first, then open the file for writing: `open(p, "w",
  encoding=...)` empties the file before an encoding error stops the write
  (that emptied `astar_c.bas` once). Write the bytes of an already encoded
  text.
* The sources are ASCII: a C++ comment's approximately-equal, square root
  or superscript 2 sign becomes `~`, `sqrt`, `^2`.

## The scripts (wzapi, qtscript, the Chakra backend)

* Modules: `wzapi_c` (wzapi.cpp; `wzapi::X` is `wzapi_X`), `qtscript_c`
  (qtscript.cpp; `scripting_engine::X`, members and statics, is
  `scripting_engine_X`, the singleton's members its module variables; the
  free functions of qtscript.h keep their names), and the Chakra backend in
  place of quickjs_backend.cpp: `jsrt_win` (the JsRT declarations and thin
  helpers), `chakra_backend_c` (quickjs_backend.cpp's free functions: the
  object conversions, callFunction, unbox/box, the builtins, the native
  dispatcher), `chakra_functions_c` (the API wrappers and
  registerFunctions), `ChakraScriptingInstance` (quickjs_scripting_instance
  with its base wzapi::scripting_instance: one class), and
  `wzapi_execution_context` (quickjs_execution_context).
* An API function's VB signature follows from wzapi.h by one rule, so its
  JS wrapper can be written from the declaration: `WZAPI_PARAMS(...)` puts
  `ByVal context As wzapi_execution_context` first; integers and enums are
  Long, float Single, strings String, object pointers handles (0 =
  nullptr); `optional<T> _name` is two parameters `has_name As Boolean,
  opt_name As T`; string_or_string_list and va_list_treat_as_strings are a
  StringVector; a returned vector is the count, with a last `result() As
  T` parameter; a returned optional is has_value, with a last ByRef
  `result`; no_return_value is a Sub; returned UDTs (scr_area,
  researchResult: `wzapi_researchResult`, research_c has the Sub
  researchResult) are the Function's type.
* `SCRIPT_ASSERT(retval, context, expr, ...)` logs, calls
  `context.throwError`, and returns `retval`. throwError does **not**
  throw: the official build's QuickJS (Bellard's, not QuickJS-ng) only
  raises an exception when a C function returns JS_EXCEPTION, and wrap_
  returns the value it got, so a failed SCRIPT_ASSERT logs and the script
  goes on. Only real throws (JSON.parse errors, VB errors, `new` on a
  native function) become JS exceptions (`JsSetException`).
* One JsRT runtime for all instances, a context per instance (QuickJS has
  a runtime per instance, but events cross instances while a script runs,
  and JsRT switches contexts only within one runtime). The runtime is made
  with the first instance and disposed with the last.
* Every VB procedure Chakra calls (`jsNativeFunction`, the one AddressOf
  entry, with the function's number as its callback state) traps every VB
  error and turns it into a JS exception: a VB error never unwinds
  through Chakra. A JsValueRef kept outside a local variable is
  `JsAddRef`'d. `isConstructCall` (bool) and `argumentCount` (unsigned
  short) arrive in 4-byte slots: they are masked.
* The x87's sticky exception flags are cleared where control passes from
  the scripts to VB (`jsClearFPU`: after every `jsCall`, on entry to
  `jsNativeFunction`): Chakra runs the scripts' floating point on the
  same thread, and a NaN a script makes would make VB's next float to
  integer conversion anywhere an overflow (error 6).
* TimerFunc (`std::function`) is the interface `ITimerFunc`
  (`ChakraTimerFunc` implements it); a timerAdditionalData is an Object;
  a uniqueTimerID (uint64_t) a Double.
* Where Chakra's JavaScript could differ from QuickJS's: numbers saved as
  JSON (QuickJS's int/float tags are invisible through JsRT: an integral
  int32 is saved as an int), natively created functions' `length` and
  `name`, symbol-keyed properties in enumerations, error texts and stacks,
  Array.sort's order of equal elements. rules.js depends on none of them;
  the AI scripts are checked when they run (cluster C's milestone).

* Legacy Chakra lacks ES2022's built-ins: the backend's built-ins script
  (`DEBUG_BUILTINS_SOURCE`, run in every context) defines what the game's
  scripts use that QuickJS has (`Object.hasOwn`). JsRT's JsGetProperty
  refuses a primitive, where QuickJS boxes it (a script reading `.id` of a
  number gets undefined): `jsGetProp` and `jsGetPropOk` box numbers,
  strings and booleans first (`jsBoxPrimitive`).

## The renderer (gfx_api, lib/ivis_opengl)

* **The context**: one per run, chosen by `gfx_api::context::initialize`
  (`gfx_api_c`), so the backends are modules: `context::get().X(...)` is
  `gfx_api_context_X(...)` (`gfx_api_context_vb`), a `Select Case` on
  `current_backend_context` calling `gl_context_X` (`gfx_api_gl_c`) or
  `null_context_X` (`gfx_api_null_c`) with the same parameters. A VB class's
  public method cannot take a module UDT, even in a private class
  (measured), hence modules; a private class that must pass a UDT uses
  `Friend` methods. The context's non-virtual members (`loadTextureFromFile`,
  ...) are `gfx_api_c`'s `gfx_api_context_<name>`. A `std::vector` parameter
  is an array and a count; `getScreenshot`'s callback and `setSwapInterval`'s
  completion handler are run by the caller after a True return (both
  backends call them at once).
* **Pipelines**: a `pipeline_state_helper<...>` instantiation is a `pso_type`
  value named as gfx_api.h's `using` (`gfx_api::DrawImagePSO::get().bind()`
  is `gfx_api_pso_get(DrawImagePSO).bind`); `pso_create_info` writes out its
  template arguments; a C++ template over PSO types takes the `pso_type`.
* **Constant buffers and vertex data**: the UDTs of `src/shared/gfx` (compiled
  into WzGfx and the exe only: `glm_h` has glm's vectors and `mat4`,
  column-major `m(col * 4 + row)`; `gfx_api_h` the constant buffers,
  `constant_buffer_type<SHADER_X>` being `constant_buffer_type_SHADER_X`). They
  travel by address and size (`.bind_constants VarPtr(cb), LenB(cb)`;
  `buffer.upload size, VarPtr(a(0))`). A glm struct passed to WzGfx is its
  address (0 = the C++ default argument); a returned one an out-address.
* **Objects and handles**: textures and buffers are the interface objects
  `gfx_api_texture`, `gfx_api_texture_array`, `gfx_api_abstract_texture`,
  `gfx_api_buffer` (no underscores in their methods: `uploadSub`); C++'s
  delete is the last reference going. An `IMAGEFILE *` is a Long handle, an
  `AtlasImageDef *` `file * &H10000 + index + 1`; a texture page is tex_c's
  index; `iV_Image` is a public creatable class (`New iV_Image`). A PIELIGHT
  is a Long (r in the low byte); `WZCOL_X` is `psPalette(WZCOL_X)`.
* **GL calls** are `gl_win`'s (generated by `tools/gen_gl.py`; never edit
  it): GL 1.0/1.1 functions are opengl32.dll's `Declare`s, later ones
  wrappers through DispCallFunc on the loaded entry points; glad's
  `GLAD_GL_*` flags are Booleans the loader sets. A generated parameter
  whose name VB refuses (`array`) gets an underscore.
* **Images and text**: PNG through GDI+ (`png_util_gdip`, libpng's API and
  messages; the gAMA, cHRM, sRGB and iCCP chunks are dropped first so that
  GDI+ corrects nothing, as libpng does not); text through GDI
  (`textdraw_gdi`: textdraw.cpp's structure, GDI's greyscale glyphs instead
  of FreeType's, Windows' shaping instead of HarfBuzz and FriBiDi).
* A STRUCTURE that C++ makes on the stack and draws (buildBlueprint's) is a
  scratch slot of the structure pool (`StructContainer_scratch(n)`: never
  alive, so no list or iteration sees it; forgotten at clear()).
* Display randomness is cstdlib_vb's `rand`, never the simulation's
  generators; effects added by the simulation do nothing in the vbtest runs
  (`vbEffectsOff`), where statsref's effect stand-ins add nothing, and nor
  do the display's entry points that renderLoop calls (`vbDisplayOff`:
  displayRenderLoop, processInput, displayWorld, ..., statsref_render.cpp's
  stand-ins).
* VB pitfalls found here: `array` cannot name a parameter, `attribute` a
  variable, nor `each`, `output`, `object`; a procedure has at most 64 KB of
  locals ("Too many local, nonstatic variables": make the big UDTs
  `Static`); VB's compile-error line numbers in a long module were the file
  line minus 2 (a syntax-error probe calibrates them).

## Widgets (lib/widget, the title screens, the in-game interface)

334 C++ classes derive from WIDGET; a VB class each (with forwarders for
every inherited method) would pass VB's limits, so:

* **One VB class, `WIDGET`** (src/game/WIDGET.cls), is every widget; `kind`
  (`K_<Class>`) says which C++ class it is. A `std::shared_ptr<X>`, a
  `weak_ptr`'s lock, an `X *` or `X &` of any widget class is a `WIDGET`
  (Nothing = nullptr). VB's reference count is shared_ptr's; a weak pointer
  is an address (`WIDGET_fromPtr`, `W_SCREEN_fromPtr`) cleared by the
  object when it dies, or a widgweak_vb serial (`WeakWidget_lock`).
* **A class's own fields** are a data class (a .cls whose header has
  `' (VB) data of widget class <Class>`), one object per level of the chain
  that has fields, read with the generated accessor `<Class>_(w).field`.
  The data class is `<Class>.cls`; a class name over 31 characters (VB's
  limit: the programmatic id "Warzone.<name>" has 39 at most) gets a
  shorter one (`OptValueChangerWrapperData` for
  OptionsValueChangerWidgetWrapper). A UDT field is Private with Friend
  Property Get/Let.
* **Declarations**, comment lines in the module of the C++ file:
  `' (VB) widget class <Class> : <Parent>` and, for each virtual a class
  introduces, `' (VB) virtual <Class>.<method>`.
* **Methods** are `<Class>_<method>(ByVal self As WIDGET, ...)` (`this` and
  `shared_from_this()` are `self`). A WIDGET virtual is called as the
  WIDGET method (`w.display x, y`); one a subclass introduces is
  `vt_<DeclaringClass>_<method>(w, ...)`; `Base::method(...)` is
  `<Base>_method self, ...`. An override is a `<Class>_<method>` with the
  declaring class's parameters.
* **Construction**: `std::make_shared<X>(args)` is `X_new(args)`
  (`WIDGET_alloc(K_X)`, then `X_construct`). `X_construct` runs the
  parent's construct, then `self.kind = K_X`, then the member initialisers
  and the body. `~X()` is `X_destruct(self)`, run from WIDGET's
  Class_Terminate (most derived first). `dynamic_cast` is
  `WIDGET_isKindOf(w, K_X)`.
* **Lambdas, std::function, function pointers** are `WzFunc` objects:
  `WzFunc_new(FN_<name>, captures...)`, `WzFunc_call(f, args...)`. The
  module that writes the lambda declares `Public Const FN_<name>` (its
  fork's id range) and `<module>_fn(f, args())` with a `Select Case f.id`
  over its lambdas' bodies (captures `f.cap(i)`, arguments `args(i)`).
* **Generated**: `python -I tools\gen_widgets.py` writes
  `widget_classes_vb.bas` (kinds, accessors, WIDGET_alloc, isKindOf, the
  vt_ dispatchers, WIDGET_destructAll, WzFunc_dispatch) and prints
  PROBLEMs. Never edit its output.
* **`Me` is never a ParamArray argument**: VB evaluates `Me`'s default
  member there (run-time error 438; a variable, a function's result or
  `ByVal v As Variant` are fine). WIDGET.cls and OptionInfo.cls pass a
  local copy (`Set self = Me`) to `WzFunc_call`. Taking and releasing such a
  copy inside Class_Terminate does not run Terminate again (measured).
* **No error leaves Class_Terminate**: VB shows its message box there (the
  game hangs on it). WIDGET's Class_Terminate logs the error instead.
* **UI tests**: `tools\ui_test.py <script>` runs a `tools\ui_tests` script
  (clicks, keys, text, the game's own screenshots, `widgets` for the widget
  tree in the log). The coordinates are the 1024x768 screen's; the comments
  of skirmish.txt list the reticule's buttons. A left click on the radar
  orders the selected units there; a right click moves the camera.
  `key` posts the key's scan code (WzPlatform reads it from lParam, as SDL
  does), the press and the release a frame apart; `embark` orders the
  selected player's units onto its transporter (the campaign test's
  launches); `--config-of=<script>` runs on another script's saved config
  folder (a save made, then loaded). `wait N` is N frames and at least
  N / 60 s: a covered or hidden window's frames are not held by the vsync.
* **Debugging a hang** (a VB error box, or a loop): `tools\exe_debug.py`
  builds the exe with CodeView info and no optimisation; then cdb (Windows
  Kits) non-invasively: `cdb -pv -p <pid> -y bin -c "~0 kn; qd"`; where
  VB's frames don't unwind, `tools\stack_scan.py` lists the return
  addresses into Warzone.exe on the stack (`dps @esp`).
  `tools\dialog_text.ps1` prints an error box's text.

## The sound (lib/sound, WzSound)

* A `TRACK *`, `AUDIO_SAMPLE *`, `AUDIO_STREAM *` is an object of that WzSound class (Nothing =
  nullptr); TRACK and AUDIO_STREAM are PublicNotCreatable, as the game holds them (the resource module's
  WAV data, audio_PlayStream's result). A `SIMPLE_OBJECT *` in lib/sound is the game's Long handle.
* What lib/sound calls in the game (aud.h, sound_GetGameTime, realTime) is `ISoundGame`, implemented by the
  exe's `SoundGame` and handed over with `SoundSetGame` (Nothing at the ends, with the other DLLs' games).
  An AUDIO_CALLBACK is the game's number for it (`CB_droidAudioTrackStopped` ...), dispatched by
  `SoundGame.audioCallback`.
* A `std::function` onFinished of a stream is an `IStreamFinished` object (the game's lambdas: an exe
  class that calls its WzFunc; cdaudio's: a WzSound class).
* OpenAL is `al_win`: the calls openal_track.cpp makes, with OpenAL's names and semantics, mixed in VB and
  polled from sound_Update (alcProcessContext). Nothing plays while no frame calls it.
* The game must not depend on the sound: the test runs are silent (`--nosound`) and must log what they
  logged with the sound's stand-ins; with the sound on, `--vbsndcapture=<file.wav>` mixes into a file
  instead of the device; `tools\wav_activity.py <file.wav>` prints its loudness per window and channel.

## The network messages (nettypes, netqueue)

* `NETQUEUE` is a Long, `queueType * &H10000 + index * &H100 + exclude`;
  `queue.index` is `NETQUEUE_index(queue)` (and `_queueType`, `_exclude`,
  `_isPair`). The queue it names is looked up when it is used (C++ keeps
  the pointer of the moment; nothing replaces a queue between the two).
* `auto w = NETbeginEncode(queue, GAME_LASSAT);` is `Dim w As MessageWriter`
  / `Set w = NETbeginEncode(queue, GAME_LASSAT)`; a reader is a
  `MessageReader` from `NETbeginDecode`. Both are WzNet classes that only
  those functions create.
* nettypes.h's overload pairs `NETuint32_t(MessageReader&, uint32_t&)` /
  `NETuint32_t(MessageWriter&, uint32_t)` are one procedure each, with the
  C++ name, taking the reader or writer `As Object` and the value ByRef (a
  reader sets it, a writer reads it). A template over both
  (`NETQueuedDroidInfo<SerdeContext>`) is then one copy. A variable passed as
  the value must be a Long (Double for 64 bits, Boolean, String), never an
  enum type: declare such UDT fields Long with the C++ type in a comment.
* `NETPosition`, `NETRotation`, `NETVector2i` take the exe's UDTs, which do
  not cross into WzNet: they are exe procedures (`nettypes_vb`) that
  serialise the fields as nettypes.cpp does.
* The message types and the queue constants are NetApi's enums
  (`MESSAGE_TYPES`, `QueueType`, `nettypes_h_consts`); WzNet's classes write
  the few they need as numbers (a class may not use the API class's enums).
* The tests drive gtime's clock through WzFramework's `wzSetTestTicks`
  (statsref's `wzGetTicks` returns the same values), so the latency, the
  game queue times and the sync log CRCs are the same on both sides.
* nettypes.h's abstract `ReplayOptionsHandler` is an interface of WzNet;
  the exe's `WZGameReplayOptionsHandler` (and the tests' handler)
  implement it. JSON crosses as WzFramework's `nlohmann_json` objects; an
  `EmbeddedMapData` is its fields (a byte array and its size).
* `NetPlay.players` is a vector: WzNet keeps MAX_CONNECTED_PLAYERS + 1
  entries and the size (`NetPlay_players_size`, `_push_back`, `_resize`),
  for the replay viewer's slot. A `PLAYER()` is value-initialised (all 0),
  as at C++'s start; NET_InitPlayer sets the lobby's values.
* A C++ `to_json(j, x)` / `from_json(j, x)` pair is `<TYPE>_to_json j, x` /
  `<TYPE>_from_json j, x` in the module of its C++ file (j ByRef: to_json
  makes it a new object).

## Tests of game state (vbtest_vb, statsref)

* A test that needs fresh objects on a map it used before loads the map
  again (`mapshutdown`, `mapload`, `mapplayers`, `mapobjects`):
  `objshutdown` frees the objects but leaves the map tiles pointing at
  them (C++ then reads freed memory; VB reads stale handles), and
  structures built earlier have levelled the ground.
* State a test depends on is set by the test on both sides (statsref has no
  config file, no lobby and no startup): the players' table, the lobby's
  options, `bMultiMessages` as the game would set it.
* Check that a test exercises what it claims (coverage counts from the
  outputs), not only that both sides agree: identical outputs of a game
  that stalled on both sides test nothing.
* When the port replaces a stand-in that statsref also has (renderLoop,
  the script triggers), the reference must follow: extract the real C++
  into statsref with its own callees as stand-ins that return what
  stubs_game's return (`tools/statsref/statsref_render.cpp`), or put the
  VB side in the state the stand-in implies (`vbtest_vb` calls
  prepareScripts with no script loaded).

## Generated lexers and parsers (flex, bison)

* A `.lpp` / `.ypp` file is ported by hand into `<name>_lexer_c` and
  `<name>_parser_c` (or the module of the `.cpp` that drives a lexer
  directly, as `levels.cpp` does). The lexer keeps flex's rules: the
  longest match wins and the earlier rule on a tie (so a keyword beats the
  identifier rule of the same length), start conditions, `yylineno`
  counting the newlines matched, `yytext` the last match, a stray
  character returned as its signed `char` value, and `yylex_destroy`
  resetting the line count. The parser reads the tokens one at a time as
  the LALR tables would, runs each rule's action before the next token is
  read, and reports bison's plain "syntax error".
* Names keep the generator's prefix: `res_lex`, `res_parse`, and `yyerror`
  is `res_error` (bison's `%name-prefix` renames it); a rule's action logs
  under `res_parse`. `check_debug_lines.py` reads `.ypp` / `.lpp` files and
  knows these names.
* There is no C++ reference for these grammars yet: flex and bison are not
  installed. Check a port against an independent reading of the real data.

## Build and test

| Command | Does |
|---|---|
| `python -I tools\build.py [--release] [--break-compat NAME] [--unregister]` | builds the six DLLs and `Warzone.exe` into `bin/` |
| `python -I tools\check_names.py` | VB name clashes per project, locals that hide globals |
| `python -I tools\check_intdiv.py` | `\` that a later `*` or `/` takes apart |
| `python -I tools\check_with.py` | `GoTo` out of a `With` on an array element |
| `python -I tools\check_mod.py` | `Mod` that a later `*`, `/` or `\` takes apart |
| `python -I tools\check_debug_lines.py` | each log call's C++ line number, format string and function |
| `python -I tools\run_tests.py` | compares the port's output with the real C++'s (`src/tests` -> `bin\WzTests.exe`; references built with g++) |
| `python -I tools\watch_test.py [gameTime]` | an AI skirmish in a window, the game's screenshot at the game time (cluster E's milestone) |
| `python -I tools\draw_test.py` | the renderer's screenshot test (cluster D's milestone) |
| `python -I tools\ai_skirmish.py [test]` | plays a headless AI skirmish to its end (cluster C's milestone) |
| `python -I tools\replay_test.py [name]` | plays an official 4.7.0 replay headless and compares the sync logs, tick by tick, with the official build's (cluster B's milestone) |
| `python -I tools\name_budget.py ref\warzone2100` | name budget per project |

Run all five checks before a build that adds ported code, and
`run_tests.py` after a build that touches what it tests. Where a
third-party C++ library or a self-contained part of Warzone can be compiled
with g++ (Strawberry Perl's), add a reference mode to the tools and a test
command, as `tools/jsonref.cpp` does for nlohmann::json and
`tools/mapref.cpp` for wzmaplib (built from the real sources, with the
CMake-generated header written by `run_tests.py`). Where C++ leaves the
order of evaluation open (function arguments with side effects), the
reference tool and the VB test both fix one.

Test runs: `bin\Warzone.exe --frames=N` quits after N frames and logs
dialogs instead of showing them; `--nodialogs` does the latter alone;
`--vbtest=<file>` runs `run_tests.py`'s commands (`vbtest_vb`) after the
search path is set up and before the window opens, then quits (all three
VB-only options, for scripted runs). The game's own code (not a DLL) is
tested this way: `tools/statsref` compiles the real C++ of the loaders,
taken from the 4.7.0 sources by line ranges with `#line` (so its logs
carry the real file names and lines; `run_tests.py`'s `STATSREF_EXTRACTS`
lists the ranges, `STATSREF_SUBSTITUTIONS` the text an extract replaces,
as imdload's GPU buffer calls, and `STATSREF_PRELUDES` the declarations
put before it), and runs the same command file. A VB-only dump procedure
(`modelDump`) and statsref's C++ one print the same text. The capture callback
(`CB_debug_callback_capture`, `debug_capture_take`) gives the log lines of
each command, as `debug.cpp` writes them to its log file. `--configdir=<dir>` keeps a test's
config and logs apart; `--debug=<part>` and `--debugfile=<file>` work as in
Warzone. The log is `<configdir>\logs\WZlog-MMDD_HHMMSS.txt`; the default
config dir is `%APPDATA%\Warzone 2100 Project\WarzoneVB\` (Warzone's own
is `...\Warzone 2100\`), so the port never touches a real installation's
settings. `tools\capture_window.ps1` runs the game and captures its window.
