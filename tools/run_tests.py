"""
run_tests.py - checks the port's code against the C++ it was ported from.

  python -I tools\\run_tests.py [--keep] [--stats | --replays | --sound | --media]

(--stats: only the statsref tests; --replays: only the official replays,
tools/replay_test.py on each replay with reference sync logs in
ref/official: Warzone.exe plays it headless and its sync log must equal the
official 4.7.0 build's, tick by tick; --sound: only the sound tests (below);
--media: only cluster K's tests (the .wz archives, Opus, Theora, the videos'
audio: media_tests). A full run ends with them.)

Needs bin/ from tools\\build.py, and g++ (Strawberry Perl's) for the
reference tool. Builds bin/WzTests.exe (src/tests, a VB-only test program
that uses the DLLs) and build/jsonref.exe (tools/jsonref.cpp, the real
nlohmann::json 3.11.3 from ref/warzone2100), runs the same commands through
both, and compares their output byte for byte.

JSON (json_hpp, nlohmann_json): every .json file of the game data parsed and
dumped four ways (dump(4), ordered_json dump(4), dump(), dump with
ensure_ascii); crafted inputs for parse errors (the exception messages,
with line and column) and edge cases; 3000 random doubles and edge values
through dump (Grisu2).

WzConfig (WzConfig.cls, json_variant.cls): every data file whose values are
all objects read group by group (value(key).toWzString(), against the same
conversions in Python) and copied through a ReadAndWrite WzConfig (its
saved file against the real library's dump(4) of the original). Known difference: invalid UTF-8 bytes that an error
message quotes are U+FFFD in the port (a VB String holds no invalid UTF-8);
error messages are compared with them normalised.

wzmaplib (WzMap.dll): build/mapref.exe (tools/mapref.cpp with the real
lib/wzmaplib sources) and WzTests load every map folder of the game data
(campaign and skirmish, twice with different map types and player
counts) and dump what they load, with every log line; each map is also
exported in the three output formats, and the files written are compared
(the logs with the output root replaced by <root>).

The sound (cluster H): bin/VorbisTests.exe (src/vorbistests, with WzSound's
stb_vorbis port src/sound/stb_vorbis_c.bas compiled in) decodes every Ogg
Vorbis file of the game data, whole and through the streaming calls
WzSound's decoder uses; the reference is ffmpeg's own Vorbis decoder (the
downloaded stb_vorbis C is not run): the same channels, rate and sample
count, and every sample within SOUND_TOLERANCE (two float decoders round
differently). Then tools/mixer_test.py's checks of the mixer, when it exists.

The test files go to build/tests (removed afterwards unless --keep or a
test failed). Exit code 0 when everything matches.
"""
import json
import os
import random
import shutil
import struct
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import build  # noqa: E402

ROOT = build.ROOT
TESTS_SRC = os.path.join(build.SRC, "tests")
VORBIS_SRC = os.path.join(build.SRC, "vorbistests")
SOUND_WORK = os.path.join(build.BUILD, "sound_tests")
SOUND_REF = os.path.join(build.BUILD, "sound_ref")         # ffmpeg's decodes, kept between runs
FFMPEG = shutil.which("ffmpeg") or os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "WinGet", "Packages",
                                                "Gyan.FFmpeg_Microsoft.Winget.Source_8wekyb3d8bbwe",
                                                "ffmpeg-8.1-full_build", "bin", "ffmpeg.exe")
# the largest difference of a sample from ffmpeg's (16-bit units): stb_vorbis and ffmpeg's decoder are both
# float decoders but round differently (measured: see sound_tests' report of the distribution)
SOUND_TOLERANCE = 2
# the largest difference of a length (samples) from the last granule position, and from ffmpeg's: the last
# frame (half the largest Vorbis block of the game's files, 2048)
SOUND_END_TOLERANCE = 1024
WORK = os.path.join(build.BUILD, "tests")
JSONREF_SRC = os.path.join(ROOT, "tools", "jsonref.cpp")
JSONREF_EXE = os.path.join(build.BUILD, "jsonref.exe")
JSON_INCLUDE = os.path.join(build.REF, "3rdparty", "json", "include")
MAPREF_SRC = os.path.join(ROOT, "tools", "mapref.cpp")
MAPREF_EXE = os.path.join(build.BUILD, "mapref.exe")
WZMAPLIB = os.path.join(build.REF, "lib", "wzmaplib")
WZMAPLIB_SOURCES = ["map.cpp", "map_io.cpp", "map_terrain_types.cpp", "map_crc.cpp", "map_internal.cpp",
                    "map_debug.cpp", "map_jsonhelpers.cpp", "map_version.cpp", "map_script.cpp"]
# map_script.cpp's QuickJS: quickjs-wz at the commit 4.7.0 pins (ref/quickjs-wz, MIT), compiled with gcc as its
# CMakeLists.txt does (tools/map_script_test.py's reference the same way); mapref and statsref link it
QJS = os.path.join(build.ROOT, "ref", "quickjs-wz")
QJS_SOURCES = ["cutils.c", "dtoa.c", "libregexp.c", "libunicode.c", "quickjs.c"]
QJS_INCLUDES = ["-I", os.path.join(QJS, "quickjs"), "-I", os.path.join(QJS, "quickjs-wz-extensions")]
OPTIONAL_INCLUDE = os.path.join(build.REF, "3rdparty", "optional-lite", "include")
# lib/wzmaplib/CMakeLists.txt's version, as CMake writes it into map_config_internal.h.in
MAP_CONFIG_H = """#pragma once
#define WZMAPLIB_VERSION_STRING "1.4.6"
#define WZMAPLIB_VERSION_MAJOR 1
#define WZMAPLIB_VERSION_MINOR 4
#define WZMAPLIB_VERSION_REV 6
"""
GXX = shutil.which("g++") or r"C:\Strawberry\c\bin\g++.exe"
GCC = os.path.join(os.path.dirname(GXX), "gcc.exe")

# The stats loaders (cluster A3): build/statsref.exe compiles the real 4.7.0
# sources: whole files, and the loaders' part of the others by line ranges
# (with #line, so __FILE__ and __LINE__ are the real ones), with
# tools/statsref's stand-ins for what lies outside them.
STATSREF_DIR = os.path.join(ROOT, "tools", "statsref")
STATSREF_EXE = os.path.join(build.BUILD, "statsref.exe")
STATSREF_GEN = os.path.join(build.BUILD, "statsref_gen")
STATSREF_EXTRACTS = [
    # all but loadImageForWeapSubclass
    ("src/stats.cpp", [(26, 1873)]),
    # the includes and globals .. loadStructureStats (the aux bits, limits, factory flags);
    # loadStructureStrengthModifiers, structureStatsShutDown; getStructureDamage .. structureCompletionProgress;
    # structureBuild .. structureRepair (with structureHasModules, structureTotalReturn); the walls (but
    # structFindWallBlueprints and structChooseWallTypeBlueprint), foundations,
    # buildStructure, buildStructureDir; defaultAssemblyPointPos, setFunctionality; STRUCTURE's constructor and
    # destructor; the packability .. removeStruct; getStructStatFromName, getStructStatsFromName;
    # findAssemblyPointPosition .. setLasSatExists; checkForResExtractors .. getModuleStat; structureBodyBuilt ..
    # structureResistance; STRUCTURE::isFactory; holdProduction, releaseProduction; getStructureBounds;
    # checkStructure; parseFavoriteStructs, packFavoriteStructs; GlobalStructContainer; _syncDebugStructure,
    # requestOpenGate, gateCurrentOpenHeight; structGetDemolishStat, calcStructureMuzzleBaseLocation;
    # validStructResistance; structSensorDroidWeapon .. STRUCTURE::isDamaged (the sensors, the rearm pads);
    # structPowerToBuildOrAddNextModule; canStructureHaveAModuleAdded, calcLineBuild; and (B9) changeProductionPlayer,
    # setCurrentStructQuantity, handleAbandonedStructures, refundFactoryBuildPower, structSetManufacture; assignFactoryCommandDroid
    # .. structureUpdate (the factories' droids, the limits, the repair facilities, aiUpdateStructure); checkSpecificStructExists;
    # countAssignedDroids; validTemplateForFactory; FlagIsFactory .. structHeightScale (the production); resetResistanceLag,
    # lasSatStructSelected, cbNewDroid
    ("src/structure.cpp", [(28, 908), (912, 1284), (1297, 1324), (1339, 1976), (2062, 4038), (4215, 4862),
                           (4904, 5172), (5177, 5514), (5647, 7151), (7230, 7273), (7276, 7339),
                           (7341, 7345)]),
    # asFeatureStats, oilResFeature, featureInitVars, loadFeatureStats, featureStatsShutDown; featureDamage .. the end
    ("src/feature.cpp", [(25, 53), (55, 617)]),
    ("src/research.cpp", [(26, 2038)]),
    # the includes and SCRIPT_ASSERT; the execution contexts' destructors; dirtyAllDroids, dirtyAllStructures, Scrcb; setUpgradeStats, getUpgradeStats,
    # GameEntityRules' getPropertyValue and setPropertyValue; getUpgradesObject, register_common, constructStatsObject
    ("src/wzapi.cpp", [(25, 97), (222, 225), (3482, 3527), (3529, 4182), (4184, 4612)]),
    # the includes; paused, video; PAUSE_STATE, maxFastForwardTicks; the droid counts; countUpdate, gameStateUpdate (not
    # static: see STATSREF_SUBSTITUTIONS); loopPieCount, loopPolyCount; renderLoop (its UI calls: statsref's stand-ins);
    # getMaxFastForwardTicks .. gameLoop; gamePaused .. droidCountsInTransporter (the pause states, the counts)
    ("src/loop.cpp", [(26, 99), (104, 105), (110, 111), (120, 129), (131, 132), (135, 139), (155, 399), (401, 620), (622, 747),
                      (831, 963)]),
    # the globals .. loadTemplateCommon; designableTemplate; DROID_TEMPLATE's constructor and getters; loadDroidTemplates ..
    # getTemplateFromMultiPlayerID
    ("src/template.cpp", [(26, 259), (262, 345), (512, 565), (567, 736), (810, 825), (882, 888)]),
    # aCurrName; checkTemplateIsVtol's declaration; intChooseSystemStats, checkStringLength, GetDefaultTemplateName;
    # intValidTemplate; checkTemplateIsVtol
    ("src/design.cpp", [(26, 71), (312, 312), (329, 329), (1249, 1388), (2791, 2947), (3966, 3970)]),
    # the includes .. getTopExperience; droidInit; DROID's constructor and destructor; copy_experience_queue,
    # add_to_experience_queue; removeDroidBase, removeDroidFX, destroyDroid, vanishDroid; _syncDebugDroid;
    # droidWithinCommanderRange; droidTemplateType; the sums (calcSum .. calcDroidPower), calcDroidSpeed,
    # reallyBuildDroid; initDroidMovement, droidSetBits; getDroidResourceName; the levels (getDroidLevel,
    # getDroidEffectiveLevel); droidIncreaseExperience, giveExperienceForSquish; droidGetName, droidSetName;
    # nextModuleToBuild; isTransporter; isCyborg, isConstructionDroid, isFlying, isRetreatingForRepair,
    # isReturningToBase; droidOnMap, droidSetPosition, checkDroid, droidSqDist; the component getters;
    # GlobalDroidContainer; and (B7) cancelBuild, droidBodyUpgrade; recycleDroid; droidCancelRepairers, droidRemove;
    # droidUpdate; droidUpdateShields .. droidUpdateRestore (the building, demolishing, restoring); droidRepairStarted ..
    # droidUpdateDroidSelfRepair (the repairs); templateSetParts; calcDroidMuzzleBaseLocation; checkDroidsBuilding,
    # checkDroidsDemolishing; setUpBuildModule, DROID::isDamaged; electronicDroid, droidUnderRepair; vtolEmpty, vtolFull;
    # isVtolRearming .. cbSensorDroid (the VTOLs' rearming, the sensors); SelectDroid .. DeSelectDroid
    ("src/droid.cpp", [(26, 110), (112, 162), (164, 176), (282, 395), (397, 484), (486, 494), (497, 532), (534, 715), (717, 830),
                       (832, 859), (862, 1048), (1051, 1077), (1079, 1088), (1091, 1404), (1424, 1577),
                       (1586, 1631), (1633, 2051), (2055, 2072), (2366, 2479), (2481, 2529), (2587, 2896),
                       (2898, 2953), (2957, 2998), (3000, 3006), (3010, 3046), (3048, 3091), (3094, 3694),
                       (3747, 3825), (3851, 3963)]),
    ("src/move.cpp", [(20, 2471)]),
    ("src/action.cpp", [(20, 2927)]),
    ("src/cmddroid.cpp", [(20, 243)]),
    # calcDirection
    ("src/geometry.cpp", [(20, 37)]),
    # (cluster E) the includes; calcDroidIllumination (droidUpdate calls it)
    ("src/lighting.cpp", [(27, 28), (37, 37), (39, 40), (343, 389)]),
    # rayCast
    ("src/raycast.cpp", [(20, 126)]),
    ("src/objmem.cpp", [(26, 1156)]),
    ("src/baseobject.cpp", [(21, 254)]),
    ("src/visibility.cpp", [(20, 1174)]),
    # offWorldKeepLists; sLandingZone; missionIsOffworld; missionLimboExpand; missionDroidUpdate; buildMissionDroid;
    # withinLandingZone; getLandingX, getLandingY; missionTimerUpdate
    ("src/mission.cpp", [(133, 133), (140, 142), (228, 235), (263, 267), (1564, 1598), (2616, 2637), (2839, 2856), (2859, 2872),
                         (2953, 2972)]),
    ("src/projectile.cpp", [(20, 1961)]),
    # CheckInScrollLimits
    ("src/display.cpp", [(1141, 1154)]),
    # all but the includes (fmt: see STATSREF_PRELUDES)
    ("src/power.cpp", [(42, 410)]),
    # whole (dumpLocalDebugSyncLog not static: see STATSREF_SUBSTITUTIONS)
    ("lib/netplay/sync_debug.cpp", [(22, 878)]),
    # the includes; recvResearch's declaration; turnOffMultiMsg; IdToDroid .. IdToPointer; getPlayerGenericName; isHumanPlayer;
    # whosResponsible .. canGiveOrdersFor; scavengerSlot, scavengerPlayer; recvSyncRequest, sendObj, sendSyncRequest;
    # getMessageHandlingAction, shouldProcessMessage, recvMessage (the beacons and the data check are statsref's stand-ins: see
    # STATSREF_SUBSTITUTIONS), HandleBadParam, SendResearch, recvResearch, sendResearchStatus, findResearchingFacilityByResearchIndex,
    # recvResearchStatus; recvDestroyFeature
    ("src/multiplay.cpp", [(27, 93), (114, 114), (430, 441), (529, 638), (645, 793), (853, 889), (899, 913), (916, 949), (951, 961),
                           (1020, 1073), (1302, 2020), (2321, 2351)]),
    # sendDataCheck (multiPlayerLoop's: the port runs it since cluster J)
    ("src/multijoin.cpp", [(626, 638)]),
    ("src/multibot.cpp", [(26, 848)]),
    # the gifts, the alliances, the artefacts (C2: multigifts_c whole)
    ("src/multigifts.cpp", [(20, 879)]),
    # the messages and proximity displays (C2: message_c), up to initMessage; findMessage's VIEWDATA overload
    ("src/message.cpp", [(26, 347), (1014, 1047)]),
    # transporterIsEmpty (the gifts call it)
    ("src/transporter.cpp", [(692, 701)]),
    ("src/multistruct.cpp", [(28, 379)]),
    # the connection status (NET_PlayerConnectionStatus .. NETcheckPlayerConnectionStatus), NETrecvGame, messageTypeToString;
    # NETsend, NETflush and NETrecvNet are statsref's (no connections: NetPlay.bComms is false)
    ("lib/netplay/netplay.cpp", [(329, 329), (345, 345), (381, 399), (636, 716), (2929, 2960), (4885, 5031)]),
    # whole (the save thread's work done when it is joined: see STATSREF_SUBSTITUTIONS)
    ("lib/netplay/netreplay.cpp", [(20, 484)]),
    # the JSON of the replay options (Sha256 .. PLAYER, DataHashToJSON) and WZGameReplayOptionsHandler
    ("src/multiint.cpp", [(150, 150), (237, 237), (1005, 1019), (7538, 8149)]),
    # playerStats, zeroStats, hostVerifiedJoinIdentities; getMultiStats; getVerifiedJoinIdentity; the score updates
    # (updateMultiStatsDamage .. updateMultiStatsResearchComplete: the port runs them since cluster J); EcKey's and
    # PLAYERSTATS's JSON .. updateMultiStatsIdentitiesInJSON
    ("src/multistat.cpp", [(53, 55), (58, 58), (63, 66), (95, 113), (690, 859), (1323, 1460)]),
    ("src/ai.cpp", [(20, 1380)]),
    # the start and derrick positions, scriptInit, getPlayerStartPosition
    ("src/qtscript.cpp", [(285, 322)]),
    # _structureIsManufacturingPending .. FactoryGetTemplate; StatIsStructure, StatIsFeature
    ("src/intdisplay.cpp", [(1442, 1532)]),
    ("src/order.cpp", [(20, 4297)]),
    # allocComponentList, freeComponentLists, allocStructLists, freeStructureLists
    ("src/mechanics.cpp", [(27, 35), (54, 169)]),
    # selectedPlayer, realSelectedPlayer; openLoadFile, loadFile2; loadFileToBufferVector; loadFile, loadFileToBuffer
    ("lib/framework/frame.cpp", [(27, 38), (47, 48), (143, 269), (312, 322)]),
    # all but iIMDShape::getTextures, getTeamColourForModel and reloadTexturesIfLoaded, modelReloadAllModelTextures,
    # modelUpdateTilesetIdx and _imd_load_level_textures (the textures)
    ("lib/ivis_opengl/imdload.cpp", [(26, 143), (178, 281), (306, 317), (334, 1213), (1276, 2000)]),
    # all but readVisibilityData and writeVisibilityData; the danger thread's work done at once (see STATSREF_SUBSTITUTIONS)
    ("src/map.cpp", [(26, 1756), (1889, 2245)]),
    # all but fpathJobQueueLength, fpathResultQueueLength, fpathSimpleRoute and fpathTest (the threads' unit test);
    # the jobs run when queued (see STATSREF_SUBSTITUTIONS)
    ("src/fpath.cpp", [(20, 632), (752, 786)]),
    ("src/astar.cpp", [(20, 793)]),
    # RemapWzMapPlayerNumber; loadWzMapStructure, loadWzMapFeature (not static: see STATSREF_SUBSTITUTIONS);
    # shouldLoadTerrainTypeOverrides, loadTerrainTypeMapOverride
    ("src/game.cpp", [(5243, 5329), (6315, 6424), (7091, 7139), (7334, 7415)]),
    # the CRC table, crcSumU16 (trig.cpp's table check), crcSumI16, crcSumVector2i (fpathRoute's sync log line);
    # wz::crc_init, wz::crc_update (the sync log's CRC)
    ("lib/framework/crc.cpp", [(20, 79), (114, 173), (449, 507), (700, 1086)]),
    # imageLoadFilenameFromInputFilename (BASIS_ENABLED, see STATSREF_PRELUDES)
    ("lib/ivis_opengl/gfx_api.cpp", [(254, 282)]),
]
# text replaced in an extract: imdload.cpp's GPU buffers are statsref's RecordingBuffers
STATSREF_SUBSTITUTIONS = {
    "lib/ivis_opengl/imdload.cpp": [("gfx_api::context::get().create_buffer_object(", "statsref_create_buffer_object(")],
    "src/game.cpp": [("static bool loadWzMapFeature(", "bool loadWzMapFeature("),
                     ("static bool loadWzMapDroidInit(", "bool loadWzMapDroidInit("),
                     ("static bool loadWzMapStructure(", "bool loadWzMapStructure(")],
    # the debug key mappings (gInputManager): off, as stubs_game's debugMappingsAllowed
    "src/structure.cpp": [("const DebugInputManager& dbgInputManager = gInputManager.debugManager();", ""),
                          ("!dbgInputManager.debugMappingsAllowed()", "!statsref_debugMappingsAllowed()")],
    "src/multigifts.cpp": [("const DebugInputManager& dbgInputManager = gInputManager.debugManager();", ""),
                           ("!dbgInputManager.debugMappingsAllowed()", "!statsref_debugMappingsAllowed()")],
    "src/multibot.cpp": [("const DebugInputManager& dbgInputManager = gInputManager.debugManager();", ""),
                         ("!dbgInputManager.debugMappingsAllowed()", "!statsref_debugMappingsAllowed()")],
    "src/multistruct.cpp": [("const DebugInputManager& dbgInputManager = gInputManager.debugManager();", ""),
                            ("!dbgInputManager.debugMappingsAllowed()", "!statsref_debugMappingsAllowed()")],
    # the debug key mappings as in structure.cpp; the beacons and the data check (the interface, the lobby) are stand-ins
    "src/multiplay.cpp": [("const DebugInputManager& dbgInputManager = gInputManager.debugManager();", ""),
                          ("!dbgInputManager.debugMappingsAllowed()", "!statsref_debugMappingsAllowed()"),
                          ("recvBeacon(queue);", "statsref_recvBeacon(queue);"),
                          ("recvDataCheck2(queue);", "statsref_recvDataCheck2(queue);"),
                          # multiPlayerLoop's static sendDataCheck2 (the network: a stand-in, as stubs_game's)
                          ("\t\t\tsendDataCheck2();", "\t\t\tstatsref_sendDataCheck2();"),
                          ("|| dbgInputManager.debugMappingsAllowed());", "|| statsref_debugMappingsAllowed());"),
                          # the known players, the command interface and the activity log: off, as in stubs_game
                          ("addKnownPlayer(NetPlay.players[player_id].name, getMultiStats(player_id).identity);", ""),
                          ("if (wz_command_interface_enabled())", "if (false)"),
                          ("ActivityManager::instance().wasKickedByPlayer(NetPlay.players[queue.index], KICK_TYPE, reason);", "")],
    "src/qtscript.cpp": [("static std::unordered_set<uint16_t> derricks;", "std::unordered_set<uint16_t> derricks;")],
    # mapbase restores what research changes (statsref_researchState: one line, so that the lines after keep their numbers)
    "src/research.cpp": [("static UBYTE bSelfRepair[MAX_PLAYERS];",
                          "static UBYTE bSelfRepair[MAX_PLAYERS]; void statsref_researchState(bool save) { "
                          "static UBYTE s[MAX_PLAYERS]; static std::vector<PlayerUpgradeCounts> c; "
                          "static std::array<RawResearchUpgradeChangeValues, MAX_PLAYERS> r; "
                          "if (save) { memcpy(s, bSelfRepair, sizeof(s)); c = playerUpgradeCounts; r = cachedPerPlayerRawUpgradeChange; } "
                          "else { memcpy(bSelfRepair, s, sizeof(s)); playerUpgradeCounts = c; cachedPerPlayerRawUpgradeChange = r; } }")],
    # gameStateUpdate for the tests; no profiling or game story
    "src/loop.cpp": [("static void gameStateUpdate()", "void gameStateUpdate()"),
                     ("WZ_PROFILE_SCOPE(gameStateUpdate);", ""),
                     ("WZ_PROFILE_SCOPE(gameLoop);", ""),
                     ("WZ_PROFILE_SCOPE(renderLoop);", ""),
                     ("bool skipDrawing = !gfx_api::context::get().shouldDraw();", "bool skipDrawing = !statsref_shouldDraw();"),
                     ("GameStoryLogger::instance().logGameFrame();", "statsref_logGameFrame();")],
    # the danger thread's work (dangerFloodFill of lastDangerPlayer on the shadow copies) done where the thread starts or is
    # woken; nothing to wait for (dangerThread stays nullptr, so mapShutdown's lines never run)
    "src/map.cpp": [("wzSemaphoreWait(dangerDoneSemaphore);", ""),
                    ("wzSemaphorePost(dangerSemaphore);", "dangerFloodFill(lastDangerPlayer);"),
                    ('dangerThread = wzThreadCreate(dangerThreadFunc, nullptr, "wzDanger");', ""),
                    ("wzThreadStart(dangerThread);", "dangerFloodFill(lastDangerPlayer);"),
                    ("dangerSemaphore = wzSemaphoreCreate(0);", ""),
                    ("dangerDoneSemaphore = wzSemaphoreCreate(0);", "")],
    # fpathRoute for the tests; the path threads' work done at once, by the thread the job is for, in queue order
    "src/fpath.cpp": [("static FPATH_RETVAL fpathRoute(MOVE_CONTROL *psMove,", "FPATH_RETVAL fpathRoute(MOVE_CONTROL *psMove,"),
                      ("wzSemaphorePost(threadInfo.semaphore);  // Increment semaphore",
                       "{ static std::shared_ptr<FPathExecuteContext> statsrefCtx[MAX_FPATH_THREADS]; "
                       "static FpathThreadInfo *statsrefInfo[MAX_FPATH_THREADS]; "
                       "if (statsrefInfo[targetThreadId] != &threadInfo) { statsrefInfo[targetThreadId] = &threadInfo; "
                       "statsrefCtx[targetThreadId] = makeFPathExecuteContext(); } "
                       "while (!threadInfo.pathJobs.empty()) { packagedPathJob statsrefJob = std::move(threadInfo.pathJobs.front()); "
                       "threadInfo.pathJobs.pop_front(); statsrefJob(statsrefCtx[targetThreadId]); } }  "
                       "// statsref: the thread's work, at once (a new thread, a new context)")],
    # the replay's save thread: statsref has no threads (wzThreadCreate gives a dummy), so the thread's work (writing the
    # queued chunks in order) is done where NETreplaySaveStop joins it; the file's bytes are the same
    "lib/netplay/netreplay.cpp": [("wzThreadJoin(saveThread);", "replaySaveThreadFunc(replaySaveHandle);"),
                                  ('#include "netreplay.h"', '#include "lib/netplay/netreplay.h"'),
                                  ('#include "netplay.h"', '#include "lib/netplay/netplay.h"')],
    # the converters the tests' replay handler calls (not inline); the activity manager's game mode is statsref's
    "src/multiint.cpp": [("inline void from_json(const nlohmann::json& j, MULTIPLAYERGAME& p) {", "void from_json(const nlohmann::json& j, MULTIPLAYERGAME& p) {"),
                         ("inline void from_json(const nlohmann::json& j, MULTIPLAYERINGAME& p) {", "void from_json(const nlohmann::json& j, MULTIPLAYERINGAME& p) {"),
                         ("inline void from_json(const nlohmann::json& j, PLAYER& p) {", "void from_json(const nlohmann::json& j, PLAYER& p) {"),
                         ("auto currentGameMode = ActivityManager::instance().getCurrentGameMode();", "auto currentGameMode = statsref_currentGameMode();")],
    # the sync log as the tests print it (synclog)
    "lib/netplay/sync_debug.cpp": [('#include "sync_debug.h"', '#include "lib/netplay/sync_debug.h"'),
                                   ('#include "nettypes.h"', '#include "lib/netplay/nettypes.h"'),
                                   ('#include "netplay.h"', '#include "lib/netplay/netplay.h"'),
                                   ("static unsigned syncDebugNext = 0;", "unsigned syncDebugNext = 0;"),
                                   ("static std::vector<uint8_t> debugSyncTmpBuf;", "std::vector<uint8_t> debugSyncTmpBuf;"),
                                   ("static size_t dumpLocalDebugSyncLog(unsigned logIndex)", "size_t dumpLocalDebugSyncLog(unsigned logIndex)")],
}
# declarations put before an extract
STATSREF_PRELUDES = {
    "lib/ivis_opengl/imdload.cpp": ['#include "lib/ivis_opengl/gfx_api.h"',
                                    "gfx_api::buffer *statsref_create_buffer_object(const gfx_api::buffer::usage &, "
                                    "const gfx_api::context::buffer_storage_hint &, const std::string &);"],
    "src/game.cpp": ['#include "lib/framework/frame.h"', '#include "lib/framework/physfs_ext.h"', '#include "lib/framework/wzconfig.h"', '#include "lib/framework/frameresource.h"',
                     '#include "src/map.h"', '#include "src/feature.h"', '#include "src/objmem.h"', '#include "src/qtscript.h"',
                     '#include "src/structure.h"', '#include "src/geometry.h"', '#include "src/multiplay.h"', '#include "lib/netplay/netplay.h"',
                     '#include "src/template.h"', '#include "src/droid.h"',
                     '#include "wzmaplib/map.h"', "#include <algorithm>", "#include <unordered_map>"],
    "src/structure.cpp": ["bool statsref_debugMappingsAllowed();"],
    "src/multigifts.cpp": ["bool statsref_debugMappingsAllowed();"],
    "src/transporter.cpp": ['#include "lib/framework/frame.h"', '#include "src/droid.h"', '#include "src/group.h"',
                            '#include "src/objmem.h"', '#include "src/transporter.h"'],
    "src/loop.cpp": ["void statsref_logGameFrame();", '#include "src/loop.h"', "bool statsref_shouldDraw();"],
    "src/map.cpp": ["static int dangerFloodFill(int player);"],
    # addDroidDeathAnimationEffect (display: droid.cpp's, outside the extract) has no header declaration
    "src/droid.cpp": ['#include "lib/framework/frame.h"', 'struct DROID;', "void addDroidDeathAnimationEffect(DROID *psDroid);"],
    "src/mission.cpp": ['#include "lib/framework/frame.h"', '#include "src/mission.h"', '#include "src/missiondef.h"', '#include "src/levels.h"',
                        '#include "src/order.h"', '#include "src/action.h"', '#include "src/droid.h"', '#include "src/objmem.h"',
                        '#include "src/qtscript.h"', '#include "src/wzapi.h"'],
    "src/multijoin.cpp": ['#include "lib/framework/frame.h"', '#include "lib/netplay/netplay.h"', '#include "src/multijoin.h"', '#include "src/data.h"'],
    "src/multiplay.cpp": ['#include "lib/netplay/nettypes.h"', "bool statsref_recvBeacon(NETQUEUE queue);",
                          "bool statsref_recvDataCheck2(NETQUEUE queue);", "bool statsref_debugMappingsAllowed();",
                          "bool statsref_sendDataCheck2();",
                          "void autoDesyncKickRoutine(std::chrono::steady_clock::time_point now);"],
    "src/multibot.cpp": ["bool statsref_debugMappingsAllowed();"],
    "src/multistruct.cpp": ["bool statsref_debugMappingsAllowed();"],
    "lib/netplay/netplay.cpp": ['#include "lib/framework/frame.h"', '#include "lib/gamelib/gtime.h"', '#include "lib/netplay/netplay.h"',
                                '#include "lib/netplay/nettypes.h"', '#include "lib/netplay/netreplay.h"', '#include "src/multiplay.h"',
                                '#include "src/multiint.h"', '#include "src/component.h"'],
    "src/multiint.cpp": ["#include <nlohmann/json.hpp>", '#include "lib/framework/frame.h"', '#include "lib/framework/wzapp.h"',
                         '#include "lib/framework/physfs_ext.h"', '#include "lib/framework/file.h"', '#include "lib/netplay/netplay.h"',
                         '#include "lib/netplay/nettypes.h"', '#include "lib/ivis_opengl/piepalette.h"', '#include "src/multiplay.h"',
                         '#include "src/multistat.h"', '#include "src/multiint.h"', '#include "src/levels.h"', '#include "src/init.h"',
                         '#include "src/version.h"', '#include "src/data.h"', '#include "src/game.h"', '#include "src/modding.h"',
                         '#include "src/activity.h"', '#include "src/clparse.h"', '#include "3rdparty/gsl_finally.h"',
                         '#include "src/campaigninfo.h"', '#include "src/map.h"', '#include "src/random.h"',
                         "ActivitySink::GameMode statsref_currentGameMode();"],
    "src/multistat.cpp": ["#include <nlohmann/json.hpp>", '#include "lib/framework/frame.h"', '#include "lib/framework/crc.h"',
                          '#include "lib/netplay/netplay.h"', '#include "src/multistat.h"', '#include "src/multiplay.h"',
                          '#include "src/multiint.h"', '#include "src/droid.h"', '#include "src/objmem.h"'],
    "src/qtscript.cpp": ['#include "lib/framework/frame.h"', '#include "src/qtscript.h"', '#include "src/map.h"',
                         "#include <unordered_set>"],
    "src/intdisplay.cpp": ['#include "lib/framework/frame.h"', '#include "src/intdisplay.h"'],
    "src/display.cpp": ['#include "lib/framework/frame.h"', '#include "src/display.h"', '#include "src/map.h"'],
    # power.cpp's includes but fmt (3rdparty/fmt is not checked out): fmt::format("{:+.0f}", v) is printf's "%+.0f"
    "src/power.cpp": ["#include <string.h>", '#include "lib/framework/frame.h"', '#include "src/objectdef.h"', '#include "src/power.h"',
                      '#include "lib/gamelib/gtime.h"', '#include "src/objmem.h"', '#include "lib/netplay/sync_debug.h"',
                      '#include "src/multiplay.h"', '#include "src/structure.h"', '#include "src/mission.h"', "#include <string>",
                      "namespace fmt { static inline std::string format(const char *, double v) "
                      "{ char buf[64]; snprintf(buf, sizeof(buf), \"%+.0f\", v); return buf; } }"],
    # as lib/ivis_opengl/CMakeLists.txt builds it when basis_universal is found (the official builds)
    "lib/ivis_opengl/gfx_api.cpp": ["#define BASIS_ENABLED", '#include "lib/ivis_opengl/gfx_api.h"', '#include "lib/framework/physfs_ext.h"'],
}
STATSREF_SOURCES = ["src/text.cpp", "lib/framework/wzconfig.cpp", "lib/framework/wzstring.cpp", "lib/framework/debug.cpp",
                    "lib/framework/physfs_ext.cpp", "lib/framework/strres.cpp", "lib/framework/treap.cpp",
                    "lib/framework/wztime.cpp", "lib/sound/audio_id.cpp", "src/gateway.cpp", "src/random.cpp"] + \
                   ["lib/wzmaplib/src/" + n for n in WZMAPLIB_SOURCES] + ["src/faction.cpp", "src/objects.cpp", "src/group.cpp",
                                                                   "lib/framework/trig.cpp", "src/wavecast.cpp",
                                                                   "src/mapgrid.cpp", "src/pointtree.cpp",
                                                                   "src/formation.cpp", "src/steering/steering.cpp",
                                                                   "src/steering/collision_avoidance_behavior.cpp",
                                                                   "lib/netplay/byteorder_funcs_wrapper.cpp",
                                                                   "lib/framework/geometry.cpp",
                                                                   "src/combat.cpp", "src/difficulty.cpp",
                                                                   "lib/gamelib/gtime.cpp", "lib/netplay/netqueue.cpp",
                                                                   "lib/netplay/nettypes.cpp", "lib/framework/memory_pool.cpp"]
STATSREF_INCLUDES = [STATSREF_DIR, ".", "src", "lib", "lib/framework", "lib/ivis_opengl", "3rdparty", "3rdparty/glm", "3rdparty/json/include",
                     "3rdparty/optional-lite/include", "lib/wzmaplib/include", "3rdparty/utfcpp/source", "3rdparty/utf8proc"]
# lib/netplay/netplay_config.h as autorevision_netplay.cmake writes it for the 4.7.0 release (the minor version, the count of
# annotated version tags, from the official build's replays)
STATSREF_NETPLAY_CONFIG_H = """#pragma once
static uint32_t NETCODE_VERSION_MAJOR = 0x4000;
static uint32_t NETCODE_VERSION_MINOR = 104;
"""
STATSREF_CONFIG_H = """#pragma once
#define PACKAGE "warzone2100"
#define HAVE_INTTYPES_H 1
#define HAVE_STDINT_H 1
"""

# Crafted JSON inputs: (name, bytes, modes)
CASES = [
    ("empty", b"", "error"),
    ("spaces", b"   \n  ", "error"),
    ("nul", b"\x00", "error"),
    ("bad_literal", b'{"a": tru}', "error"),
    ("literal_eof", b"nul", "error"),
    ("trailing_comma_obj", b'{"a": 1,}', "error"),
    ("trailing_comma_arr", b"[1,\n2,\n]", "error"),
    ("open_array", b"[1, 2", "error"),
    ("missing_colon", b'{"a" 1}', "error"),
    ("missing_comma", b'{"a": 1 "b": 2}', "error"),
    ("leading_zero", b"[01]", "error"),
    ("minus", b"-", "error"),
    ("dot", b"1.", "error"),
    ("exp", b"1e", "error"),
    ("exp_sign", b"1e+", "error"),
    ("open_string", b'"abc', "error"),
    ("bad_escape", b'"\\x"', "error"),
    ("bad_hex", b'"\\u12g4"', "error"),
    ("lone_high", b'"\\ud800"', "error"),
    ("high_then_x", b'"\\ud800\\u0041"', "error"),
    ("lone_low", b'"\\udc00"', "error"),
    ("control", b'"a\x01b"', "error"),
    ("tab_in_string", b'"a\tb"', "error"),
    ("bad_utf8", b'"\xff"', "error"),
    ("short_utf8", b'"\xc3"', "error"),
    ("overlong", b'"\xe0\x80\x80"', "error"),
    ("bom", b'\xef\xbb\xbf{"a":1}', "error dump"),
    ("bad_bom", b"\xef\xbb{}", "error"),
    ("overflow", b"[1e400]", "error"),
    ("after_value", b'{"a":1}\n\nx', "error"),
    ("nul_after", b"[1]\x00garbage", "error dump"),
    ("comment", b"// x\n{}", "error"),
    ("dup_keys", b'{"b":1,"a":2,"b":3}', "dump dumpo compact"),
    ("unicode", '{"z":"\\u0041\\u00e9\\ud83d\\ude00","\u00e9":"\u65e5\u672c","a":"\\b\\f\\n\\r\\t\\"\\\\\\/","c":"\\u0001\\u007f\\u0080"}'.encode("utf-8"),
     "dump dumpo compact ascii"),
    ("numbers", b"[0, -0, 1.0, -1.5e-7, 123456789012345678901234567890, -9223372036854775809, -9223372036854775808,"
                b" 18446744073709551615, 18446744073709551616, 9007199254740993, 1E2, 1e-5, 0.1, 100.0, 1e15, 1e16,"
                b" 123456.789e3, 5e-324, 1.7976931348623157e308, 0.30000000000000004, 2.5e-05, 12345678.9, 1e-4]",
     "dump compact"),
    ("nested", b'{"a":{"b":[[],{},[{"c":null}],true,false]},"empty":{}}', "dump dumpo compact"),
]


def invalid_utf8_as_one_mark(data):
    """An error message quotes the bytes it read ("last read: '...'"), raw in
    C++; a VB String cannot hold invalid UTF-8, so the port has U+FFFD there.
    Compare with each run of invalid bytes (or U+FFFD) as one U+FFFD."""
    text = data.decode("utf-8", errors="replace")
    out = []
    for ch in text:
        if ch == "\ufffd" and out and out[-1] == "\ufffd":
            continue
        out.append(ch)
    return "".join(out).encode("utf-8")


def wzstring(v):
    """json_variant::toWzString: (text, ok); ok None when untouched"""
    if isinstance(v, bool):
        return ("true" if v else "false"), True
    if isinstance(v, int):
        return str(v), True
    if isinstance(v, float):
        return format(v, ".15g"), True
    if isinstance(v, str):
        return v, True
    if v is None:
        return "", None
    return "", False


def utf8_sorted(keys):
    return sorted(keys, key=lambda k: k.encode("utf-8"))


def wzread_expected(path):
    with open(path, encoding="utf-8") as f:
        root = json.load(f)
    out = []
    for g in utf8_sorted(root):
        if not isinstance(root[g], dict):
            continue
        for k in utf8_sorted(root[g]):
            text, ok = wzstring(root[g][k])
            if ok is None:
                text = "<untouched>"
            elif not ok:
                text = "<not ok>"
            out.append("%s\t%s\t%s\n" % (g, k, text))
    return "".join(out).encode("utf-8")


def groups_only(path):
    try:
        with open(path, encoding="utf-8") as f:
            root = json.load(f)
    except ValueError:
        return False
    return isinstance(root, dict) and len(root) > 0 and all(isinstance(v, dict) for v in root.values())


def run(cmd, **kw):
    r = subprocess.run(cmd, capture_output=True, text=True, **kw)
    if r.returncode != 0:
        print("FAILED (%d): %s\n%s%s" % (r.returncode, " ".join(cmd), r.stdout, r.stderr))
    return r.returncode == 0


def build_tests_exe():
    res = build.write_manifest(["WzFramework", "WzMap"], "WzTests")
    if res is None:
        return False
    vbp = build.write_vbp("WzTests", "exe", TESTS_SRC, ["WzFramework", "WzMap"], False, res)
    return build.make("WzTests", TESTS_SRC, vbp)


def build_vorbis_tests():
    """bin/VorbisTests.exe: src/vorbistests and WzSound's stb_vorbis port (no DLL; src/shared is left out: its
    modules need names their projects define, and the port uses none of them)."""
    vbp = build.write_vbp("VorbisTests", "exe", VORBIS_SRC, [], False, None)
    with open(vbp, "r", encoding="latin-1") as f:
        lines = [l for l in f.read().split("\n") if "..\\shared\\" not in l]
    k = next(i for i, l in enumerate(lines) if l.startswith("Startup="))
    lines.insert(k, "Module=stb_vorbis_c; ..\\sound\\stb_vorbis_c.bas")
    with open(vbp, "w", encoding="latin-1", newline="\r\n") as f:
        f.write("\n".join(lines))
    return build.make("VorbisTests", VORBIS_SRC, vbp)


def sound_files():
    """Every Ogg Vorbis file of the game data (data/, as the game mounts it), sorted."""
    out = []
    for dirpath, dirnames, filenames in os.walk(os.path.join(ROOT, "data")):
        dirnames.sort()
        for n in sorted(filenames):
            if n.lower().endswith(".ogg"):
                out.append(os.path.join(dirpath, n))
    return sorted(out)


def sound_reference(path):
    """ffmpeg's decode of path as 16-bit interleaved samples (cached in build/sound_ref)."""
    rel = os.path.relpath(path, os.path.join(ROOT, "data")).replace(os.sep, "_")
    out = os.path.join(SOUND_REF, rel + ".pcm")
    if not os.path.exists(out) or os.path.getmtime(out) < os.path.getmtime(path):
        os.makedirs(SOUND_REF, exist_ok=True)
        r = subprocess.run([FFMPEG, "-v", "error", "-y", "-i", path, "-f", "s16le", "-acodec", "pcm_s16le", out],
                           capture_output=True)
        if r.returncode != 0:
            return None
    with open(out, "rb") as f:
        return f.read()


def ogg_last_granule(path):
    """The granule position of a file's last Ogg page, page by page: the stream's length in samples, as
    libvorbis's ov_pcm_total gives it (4.7.0's WZVorbisDecoder::totalSamples, and what ov_read decodes; ffmpeg's
    raw decode keeps the last frame's padding after it)."""
    with open(path, "rb") as f:
        b = f.read()
    pos, granule = 0, None
    while pos + 27 <= len(b) and b[pos:pos + 4] == b"OggS":
        granule = struct.unpack_from("<q", b, pos + 6)[0]
        nseg = b[pos + 26]
        pos += 27 + nseg + sum(b[pos + 27:pos + 27 + nseg])
    return granule


def sound_tests(keep):
    """VorbisTests.exe against ffmpeg on every game sound, then tools/mixer_test.py (cluster H)."""
    import array
    import time
    try:
        import numpy
    except ImportError:
        numpy = None
    if not os.path.exists(FFMPEG):
        print("sounds: ffmpeg not found (the reference decoder)")
        return False
    if not build_vorbis_tests():
        return False
    files = sound_files()
    if os.path.exists(SOUND_WORK):
        shutil.rmtree(SOUND_WORK)
    os.makedirs(SOUND_WORK)
    listfile = os.path.join(SOUND_WORK, "files.txt")
    with open(listfile, "w", encoding="utf-8") as f:
        f.write("\n".join(files) + "\n")
    t = time.time()
    r = subprocess.run([os.path.join(build.BIN, "VorbisTests.exe"), listfile, SOUND_WORK], capture_output=True)
    print("VorbisTests.exe: %.1f s, exit code %d" % (time.time() - t, r.returncode))
    results = {}
    total_ms = None
    resfile = os.path.join(SOUND_WORK, "results.txt")
    if os.path.exists(resfile):
        for line in open(resfile, encoding="latin-1").read().splitlines():
            parts = line.split("\t")
            if parts[0] == "total":
                total_ms = float(parts[1])
            elif parts[0].isdigit():
                results[int(parts[0])] = parts
    ok = 0
    maxdiffs = {}
    problems = []
    ends = []           # files whose length is not the last granule position (within SOUND_END_TOLERANCE)
    ffmpeg_ends = []    # files that ffmpeg decodes longer or shorter (the samples they share agree)
    for i, path in enumerate(files):
        rel = os.path.relpath(path, os.path.join(ROOT, "data"))
        res = results.get(i)
        if res is None or res[1] == "error":
            problems.append("%s: %s" % (rel, "no result" if res is None else " ".join(res[2:])))
            continue
        channels, rate, samples, stream_len, check = int(res[1]), int(res[2]), int(res[3]), int(res[5]), res[6]
        ref = sound_reference(path)
        if ref is None:
            problems.append("%s: ffmpeg failed" % rel)
            continue
        with open(os.path.join(SOUND_WORK, "%d.pcm" % i), "rb") as f:
            vb = f.read()
        info = ""
        if check != "same":
            info += " stream check: %s" % check
        if abs(stream_len - samples) > SOUND_END_TOLERANCE:      # (stream_length_in_samples is the granule position)
            info += " stream_length_in_samples %d, decoded %d" % (stream_len, samples)
        # the length: the last page's granule position (libvorbis's ov_pcm_total), or within the last frame of
        # it: stb_vorbis infers the last frame's length from the granule position its own way (it never
        # outputs the first packet's samples; ffmpeg's decode may stop a frame short of the granule or run on)
        granule = ogg_last_granule(path)
        ffmpeg_samples = len(ref) // 2 // channels
        if samples != granule:
            ends.append("%s %+d" % (os.path.basename(rel), samples - granule))
            if abs(samples - granule) > SOUND_END_TOLERANCE:
                info += " samples: %d, the last granule position %d" % (samples, granule)
        if ffmpeg_samples != samples:
            ffmpeg_ends.append(ffmpeg_samples - samples)
            if abs(ffmpeg_samples - samples) > SOUND_END_TOLERANCE:
                info += " samples: ffmpeg %d, port %d" % (ffmpeg_samples, samples)
        n = min(len(ref), len(vb)) // 2
        if numpy is not None:
            a = numpy.frombuffer(ref[:n * 2], dtype="<i2").astype(numpy.int32)
            b = numpy.frombuffer(vb[:n * 2], dtype="<i2").astype(numpy.int32)
            d = int(numpy.abs(a - b).max()) if n else 0
        else:
            a = array.array("h", ref[:n * 2])
            b = array.array("h", vb[:n * 2])
            d = max((abs(x - y) for x, y in zip(a, b)), default=0)
        maxdiffs[rel] = d
        if d > SOUND_TOLERANCE:
            info += " a sample differs by %d" % d
        if info:
            problems.append("%s (%d ch, %d Hz):%s" % (rel, channels, rate, info))
        else:
            ok += 1
    dist = {}
    for d in maxdiffs.values():
        dist[d] = dist.get(d, 0) + 1
    print("sounds: %d of %d decode as ffmpeg does (samples within %d; files by largest difference: %s); decode %s ms"
          % (ok, len(files), SOUND_TOLERANCE, ", ".join("%d: %d" % (k, dist[k]) for k in sorted(dist)),
             "%.0f" % total_ms if total_ms is not None else "?"))
    longer = [d for d in ffmpeg_ends if d > 0]
    shorter = [-d for d in ffmpeg_ends if d < 0]
    print("  lengths: %d at the last granule position; %d not (port - granule): %s; ffmpeg's decode is longer in %d "
          "(by %s), shorter in %d (by %s)" % (len(files) - len(ends), len(ends), ", ".join(ends) or "-",
                                              len(longer), "%d to %d" % (min(longer), max(longer)) if longer else "-",
                                              len(shorter), "%d to %d" % (min(shorter), max(shorter)) if shorter else "-"))
    for p in problems[:20]:
        print("  " + p)
    passed = ok == len(files)
    if os.path.exists(os.path.join(ROOT, "tools", "mixer_test.py")):
        import mixer_test
        passed = mixer_test.run(keep) and passed
    if passed and not keep:
        shutil.rmtree(SOUND_WORK)
    return passed


def build_quickjs():
    """QuickJS's objects in build/quickjs (rebuilt when a source is newer); None when gcc fails."""
    out = os.path.join(build.BUILD, "quickjs")
    os.makedirs(out, exist_ok=True)
    version = open(os.path.join(QJS, "quickjs", "VERSION.txt")).read().strip()
    objs = []
    for n in QJS_SOURCES:
        src = os.path.join(QJS, "quickjs", n)
        o = os.path.join(out, n[:-2] + ".o")
        if not (os.path.exists(o) and os.path.getmtime(o) >= os.path.getmtime(src)):
            # quickjs-wz's CMakeLists.txt: _GNU_SOURCE, QJS_DISABLE_ATOMICS, WIN32_LEAN_AND_MEAN, _WIN32_WINNT, CONFIG_VERSION
            if not run([GCC, "-c", "-O2", "-w", "-D_GNU_SOURCE", "-DQJS_DISABLE_ATOMICS", "-DWIN32_LEAN_AND_MEAN",
                        "-D_WIN32_WINNT=0x0601", '-DCONFIG_VERSION="%s"' % version, "-DQUICKJS_HAVE_SYS_TIME_H"]
                       + QJS_INCLUDES + ["-o", o, src]):
                return None
        objs.append(o)
    return objs


def build_mapref():
    sources = [os.path.join(WZMAPLIB, "src", n) for n in WZMAPLIB_SOURCES]
    qjs = build_quickjs()
    if qjs is None:
        return False
    if os.path.exists(MAPREF_EXE) and os.path.getmtime(MAPREF_EXE) >= max(os.path.getmtime(p) for p in sources + [MAPREF_SRC] + qjs):
        return True
    print("building mapref.exe (g++, lib/wzmaplib from ref/warzone2100, map_script.cpp over QuickJS) ...")
    config_dir = os.path.join(build.BUILD, "mapref_include")
    os.makedirs(os.path.join(config_dir, "wzmaplib_internal"), exist_ok=True)
    with open(os.path.join(config_dir, "wzmaplib_internal", "map_config_internal.h"), "w", newline="\n") as f:
        f.write(MAP_CONFIG_H)
    return run([GXX, "-std=c++17", "-O1", "-static", "-I", JSON_INCLUDE, "-I", OPTIONAL_INCLUDE, "-I", config_dir,
                "-I", os.path.join(WZMAPLIB, "include")] + QJS_INCLUDES + ["-o", MAPREF_EXE, MAPREF_SRC] + sources + qjs
               + ["-lpthread"])


def statsref_extract(rel, ranges):
    """The lines of a 4.7.0 source that statsref compiles, each range after a #line."""
    with open(os.path.join(build.REF, rel), encoding="latin-1", newline="") as f:
        lines = f.read().replace("\r\n", "\n").split("\n")
    out = ["// generated by tools/run_tests.py from ref/warzone2100/%s, lines %s"
           % (rel, ", ".join("%d-%d" % r for r in ranges))]
    out += STATSREF_PRELUDES.get(rel, [])
    for a, b in ranges:
        out.append('#line %d "%s"' % (a, rel))
        for ln in lines[a - 1:b]:
            for old, new in STATSREF_SUBSTITUTIONS.get(rel, []):
                ln = ln.replace(old, new)
            out.append(ln)
    return "\n".join(out) + "\n"


def build_statsref():
    tool_files = [os.path.join(STATSREF_DIR, n) for n in os.listdir(STATSREF_DIR) if n.endswith((".cpp", ".h"))]
    inputs = ([os.path.join(build.REF, r) for r, _ in STATSREF_EXTRACTS] + [os.path.join(build.REF, s) for s in STATSREF_SOURCES]
              + tool_files + [os.path.abspath(__file__)])
    qjs = build_quickjs()
    if qjs is None:
        return False
    if os.path.exists(STATSREF_EXE) and os.path.getmtime(STATSREF_EXE) >= max(os.path.getmtime(p) for p in inputs + qjs):
        return True
    print("building statsref.exe (g++, the stats loaders from ref/warzone2100) ...")
    os.makedirs(STATSREF_GEN, exist_ok=True)
    with open(os.path.join(STATSREF_GEN, "wz2100-generated-config.h"), "w", newline="\n") as f:
        f.write(STATSREF_CONFIG_H)
    os.makedirs(os.path.join(STATSREF_GEN, "lib", "netplay"), exist_ok=True)
    with open(os.path.join(STATSREF_GEN, "lib", "netplay", "netplay_config.h"), "w", newline="\n") as f:
        f.write(STATSREF_NETPLAY_CONFIG_H)
    extracts = []
    for rel, ranges in STATSREF_EXTRACTS:
        p = os.path.join(STATSREF_GEN, rel.replace("/", "_").replace(".cpp", "_extract.cpp"))
        with open(p, "w", encoding="latin-1", newline="\n") as f:
            f.write(statsref_extract(rel, ranges))
        extracts.append(p)
    utf8proc_o = os.path.join(STATSREF_GEN, "utf8proc.o")
    if not run([GCC, "-c", "-O1", "-DUTF8PROC_STATIC", "-o", utf8proc_o, os.path.join(build.REF, "3rdparty", "utf8proc", "utf8proc.c")]):
        return False
    # wzmaplib's generated header, as build_mapref writes it
    os.makedirs(os.path.join(STATSREF_GEN, "wzmaplib_internal"), exist_ok=True)
    with open(os.path.join(STATSREF_GEN, "wzmaplib_internal", "map_config_internal.h"), "w", newline="\n") as f:
        f.write(MAP_CONFIG_H)
    includes = ["-I", STATSREF_GEN] + [a for d in STATSREF_INCLUDES for a in ("-I", d)]
    # -DNDEBUG: a release build's assert()s (none)
    return run([GXX, "-std=c++17", "-O1", "-static", "-w", "-DNDEBUG", "-D__STDC_FORMAT_MACROS", "-D__STDC_LIMIT_MACROS", "-DUTF8PROC_STATIC", "-D_USE_MATH_DEFINES",
                "-include", "mutex"] + includes + QJS_INCLUDES + ["-o", STATSREF_EXE] + extracts + STATSREF_SOURCES
               + [os.path.join(STATSREF_DIR, "statsref.cpp"), os.path.join(STATSREF_DIR, "statsref_physfs.cpp"),
                  os.path.join(STATSREF_DIR, "statsref_render.cpp"), utf8proc_o] + qjs
               + ["-lws2_32", "-lpthread"],                            # byteorder_funcs_wrapper.cpp's htonl
               cwd=build.REF)


STATS_ORDER = [("weapons", "weapons.json"), ("propulsion", "propulsion.json"), ("body", "body.json"), ("brain", "brain.json"),
               ("ecm", "ecm.json"), ("sensor", "sensor.json"), ("repair", "repair.json"), ("construct", "construction.json"),
               ("proptypes", "propulsiontype.json"), ("propsounds", "propulsionsounds.json"), ("terrain", "terraintable.json"),
               ("weapmods", "weaponmodifier.json"), ("templates", "templates.json"), ("structures", "structure.json"),
               ("structmods", "structuremodifier.json"), ("features", "features.json"), ("research", "research.json")]


def stats_set_commands(dataset, players="1"):
    """The commands that load a data set's stats as stats.wrf orders them (a file the set lacks comes from base)."""
    data = os.path.join(build.REF, "data")
    cmds = ["reset", "releasestrings", "players\t" + players, "selected\t0", "gamemode\t1", "debug\tall",
            "strings\t%s/messages/strings/names.txt" % dataset]
    for kind, name in STATS_ORDER:
        folder = dataset if os.path.exists(os.path.join(data, dataset, "stats", name)) else "base"
        cmds.append("load\t%s\t%s/stats/%s" % (kind, folder, name))
    for sid in ("ZNULLDESIGN", "ViperLtMGWheels", "ConstructionDroidNP", "NoSuchName"):
        cmds.append("getstring\t" + sid)
    cmds.append("dump")
    return cmds


def stats_cases():
    """[(name, files {name: bytes}, commands)]: crafted stats files for the loaders' error paths
    (the files go to the write dir's statstest folder)."""
    data = os.path.join(build.REF, "data")

    def jl(rel):
        with open(os.path.join(data, rel), encoding="utf-8") as f:
            return json.load(f)

    def jd(obj):
        return json.dumps(obj, indent=4).encode("utf-8")

    w = jl("mp/stats/weapons.json")
    mg = dict(w["MG1Mk1"])
    znull = dict(w["ZNULLWEAPON"])
    mp_components = ["load\t%s\tmp/stats/%s" % (k, n) for k, n in STATS_ORDER[:8]]
    pre = ["reset", "releasestrings", "players\t1", "selected\t0", "gamemode\t1", "debug\tall"]
    cases = []

    weapons = {"ZNULLWEAPON": znull,
               "A-BadClass": dict(mg, weaponClass="FOO", droidType="BOGUS", usageClass="Cyborg Arms"),
               "B-Ranges": dict(mg, shortRange=900, longRange=800, flags=["AirOnly", "NoFriendlyFire", "teleportcapture"]),
               "C-SameRange": dict(mg, shortRange=800, longRange=800, shortHit=50, longHit=60, flags="ShootAir, AllowedOnTransporter"),
               "D-Wraps": dict(mg, numRounds=300, rotate=40000, minElevation=-10, maxElevation=70000, effectSize=70000, numAttackRuns=-2,
                               weaponSize=5, flightSpeed=0, usageClass="SuperCyborg"),
               "E-Strings": dict(mg, damage="12", firePause="abc", longRange=12.7, designable="false", penetrate="yes",
                                 periodicalDamageWeaponClass="BAD", periodicalDamageWeaponSubClass="BAD", periodicalDamageWeaponEffect="BAD"),
               "F-Models": dict(mg, model="nosuchmodel.pie", mountModel="mibnkdrl.pie", trailGfx="FXMflare.PIE", weaponWav="-1",
                                periodicalDamageWeaponClass="HEAT", periodicalDamageWeaponSubClass="CANNON", periodicalDamageWeaponEffect="FLAMER"),
               "G-Types": dict(mg, droidType="CYBORG_SUPER", weaponSubClass="LAS_SAT", movement="HOMING-INDIRECT", weaponEffect="ALL ROUNDER"),
               "Z-BadSubClass": dict(mg, weaponSubClass="SPOON")}
    cases.append(("weapons", {"weapons1.json": jd(weapons),
                              "weapons2.json": jd({"MG": mg}),
                              "weapons3.json": jd({"ZNULLWEAPON": znull, "MG": dict(mg, weaponEffect="BAD")}),
                              "weapons4.json": jd({"ZNULLWEAPON": znull, "MG": dict(mg, movement="SIDEWAYS")}),
                              "weapons5.json": jd({"ZNULLWEAPON": znull, "MG": dict(mg, weaponWav="nosuch.ogg")}),
                              "weapons6.json": jd({"ZNULLWEAPON": znull, "MG": dict(mg, explosionWav="lrgexpl.ogg")})},
                  pre + ["load\tweapons\tstatstest/weapons1.json", "dump"] +
                  sum([["reset", "load\tweapons\tstatstest/weapons%d.json" % k, "dump"] for k in range(2, 7)], [])))

    body = jl("mp/stats/body.json")
    vipers = dict(body["Body1REC"])
    bodies = {"ZNULLBODY": body["ZNULLBODY"], "Body1REC": vipers,
              "X-Extra": dict(vipers, propulsionExtraModels={"wheeled01": {"left": "prlwhl1.pie", "right": "nosuch.pie", "moving": "chblades.pie"},
                                                             "hover01": "prlhov1.pie"}, size="HEAVY", droidType="TRANSPORTER")}
    props = jl("mp/stats/propulsion.json")
    cases.append(("bodies", {"body1.json": jd(bodies),
                             "body2.json": jd({"ZNULLBODY": body["ZNULLBODY"], "B": dict(vipers, size="ENORMOUS")}),
                             "body3.json": jd({"ZNULLBODY": body["ZNULLBODY"], "B": dict(vipers, propulsionExtraModels={"NoSuchProp": {"left": "x.pie"}})}),
                             "prop1.json": jd({"ZNULLPROP": props["ZNULLPROP"], "wheeled01": props["wheeled01"], "hover01": props["hover01"],
                                               "P-Bad": dict(props["wheeled01"], type="Floating", turnSpeed=0, spinSpeed="0", acceleration=0,
                                                             hitpointPctOfBody=-5)})},
                  pre + ["load\tpropulsion\tstatstest/prop1.json", "load\tbody\tstatstest/body1.json", "dump",
                         "reset", "load\tpropulsion\tstatstest/prop1.json", "load\tbody\tstatstest/body2.json", "dump",
                         "reset", "load\tpropulsion\tstatstest/prop1.json", "load\tbody\tstatstest/body3.json", "dump"]))

    brains = jl("mp/stats/brain.json")
    brain_ids = list(brains)
    other = jl("mp/stats/sensor.json")
    ecm = jl("mp/stats/ecm.json")
    repair = jl("mp/stats/repair.json")
    cases.append(("components", {
        "brain1.json": jd({"ZNULLBRAIN": brains["ZNULLBRAIN"], brain_ids[-1]: brains[brain_ids[-1]],
                           "C-NoTurret": dict(brains[brain_ids[-1]], turret="NoSuchGun", cmdExpRange=7, ranks=None),
                           "C-Plain": {"name": "plain", "ranks": ["a", "b"], "thresholds": [1, 2, 3], "maxDroids": "4"}}),
        "sensor1.json": jd({"ZNULLSENSOR": other["ZNULLSENSOR"], "S-Bad": dict(other["ZNULLSENSOR"], location="ROOF", type="PSYCHIC"),
                            "S-Types": dict(other["ZNULLSENSOR"], location="TURRET", type="RADAR DETECTOR")}),
        "ecm1.json": jd({"ZNULLECM": ecm["ZNULLECM"], "E-Bad": dict(ecm["ZNULLECM"], location="ROOF")}),
        "repair1.json": jd({"ZNULLREPAIR": repair["ZNULLREPAIR"], "R-Zero": dict(repair["ZNULLREPAIR"], time=0, location="ROOF")}),
        "proptypes1.json": jd({"Wheeled": {"flightName": "GROUND", "multiplier": 70000}, "Tracked": {"flightName": "AIR", "multiplier": 150},
                               "Legged": {"flightName": "SPACE"}, "Hover": {}, "Lift": {"flightName": "AIR"}, "Propellor": {},
                               "Half-Tracked": {"multiplier": 120}}),
        "proptypes2.json": jd({"Wheeled": {}, "Tracked": {}, "Legged": {}, "Hover": {}, "Lift": {}, "Propellor": {}, "Rocket": {}}),
        "propsounds1.json": jd({"Wheeled": {"szStart": "-1", "szIdle": "-1", "szMoveOff": "-1", "szMove": "-1", "szHiss": "-1", "szShutDown": "-1"},
                                "Floating": {"szStart": "-1", "szIdle": "-1", "szMoveOff": "-1", "szMove": "-1", "szHiss": "-1", "szShutDown": "-1"}}),
        "weapmods1.json": jd({"AAA": {"Wheeled": 10}, "ANTI TANK": {"Wheeled": 70000, "Legged": -20, "LIGHT": 50, "HUGE": 7},
                              "FLAMER": {"Tracked": "77", "SUPER HEAVY": 1}}),
        "terrain1.json": jd({"t%d" % k: {"id": k, "speedFactor": {"wheeled": k * 10, "lift": "abc"}} for k in range(12)})},
        pre + ["load\tweapons\tmp/stats/weapons.json", "load\tbrain\tstatstest/brain1.json", "load\tsensor\tstatstest/sensor1.json",
               "load\tecm\tstatstest/ecm1.json", "load\trepair\tstatstest/repair1.json", "load\tproptypes\tstatstest/proptypes1.json",
               "load\tpropsounds\tstatstest/propsounds1.json", "load\tweapmods\tstatstest/weapmods1.json",
               "load\tterrain\tstatstest/terrain1.json", "dump",
               "reset", "load\tproptypes\tstatstest/proptypes2.json", "dump"]))

    s = jl("mp/stats/structure.json")
    wall = s["A0HardcreteMk1Wall"]
    demolish = s["A0ADemolishStructure"]
    structs = {"A0ADemolishStructure": demolish, "A0HardcreteMk1Wall": wall,
               "S-BadType": dict(wall, type="CASTLE"),
               "S-Odd": dict(wall, strength="SQUISHY", width=150, breadth=120, userLimits=[1, 5, 9], flags=["Connected", "Other"],
                             weapons=["MG1Mk1", "MG2Mk1", "MG3Mk1", "MG1Mk1"], sensorID="SensorTower1Mk1", baseModel="nosuchbase.pie",
                             structureModel=["blwallh.pie", " nosuch.pie "], combinesWithWall=True),
               "S-BadWeapon": dict(wall, weapons=["MG1Mk1", "NoSuchGun"], ecmID="NoSuchECM", sensorID="NoSuchSensor"),
               "_config_": {"baseStructDamageExpLevel": 3}}
    cases.append(("structures", {
        "structure1.json": jd(structs),
        "structure2.json": jd({"A0ADemolishStructure": demolish, "_config_": {"baseStructDamageExpLevel": "4"}}),
        "structure3.json": jd({"A0HardcreteMk1Wall": wall, "_config_": {"baseStructDamageExpLevel": 12}}),
        "structure4.json": jd({"A0HardcreteMk1Wall": wall}),
        "structmods1.json": jd({"ANTI TANK": {"SOFT": 70000, "MEDIUM": -1, "HARD": "5", "BUNKER": 40, "PAPER": 3}, "NOPE": {"SOFT": 1},
                                "FLAMER": {"SOFT": 120}})},
        pre + mp_components + ["load\tstructures\tstatstest/structure1.json", "load\tstructmods\tstatstest/structmods1.json", "dump",
                               "load\tstructures\tstatstest/structure2.json", "dump",
                               "reset"] + mp_components + ["load\tstructures\tstatstest/structure3.json", "dump",
                                                           "reset"] + mp_components + ["load\tstructures\tstatstest/structure4.json", "dump"]))

    feats = jl("base/stats/features.json")
    fid = list(feats)[0]
    cases.append(("features", {"features1.json": jd({fid: feats[fid], "F-Odd": dict(feats[fid], type="CLOUD", model="nosuch.pie", width=70000,
                                                                                 tileDraw=0, lineOfSight="0", hitpoints=-1)})},
                  pre + ["load\tfeatures\tstatstest/features1.json", "load\tfeatures\tstatstest/features1.json", "dump"]))

    r = jl("mp/stats/research.json")
    mgres = dict(r["R-Wpn-MG1Mk1"])
    res = {"R-A": dict(mgres, keyTopic=2, techCode=2, disabledWhen=100, iconID="IMAGE_SPOON", subgroupIconID="IMAGE_RES_GRPDAM",
                       imdName="nosuch.pie", imdName2="blwallh.pie", category="Chain",
                       resultComponents=["MG1Mk1", "NoSuchComp", " MG2Mk1 "], replacedComponents=["a:b:c", "X:MG1Mk1", "MG1Mk1:MG2Mk1", "MG1Mk1:Y"],
                       redComponents=["NoSuchComp", "MG1Mk1"], resultStructures=["A0HardcreteMk1Wall", "NoSuchStruct"],
                       requiredStructures=["A0HardcreteMk1Wall"], redStructures=["NoSuch"], requiredResearch=["R-Missing"]),
           "R-B": dict(mgres, category="Chain", requiredResearch=["R-A"], msgName="RES_X", techCode=1),
           "R-C": dict(mgres, category="Chain", requiredResearch=["R-B"], statID="A0HardcreteMk1Wall"),
           "R-D": dict(mgres, category="Fork", requiredResearch=[]),
           "R-E": dict(mgres, category="Fork", requiredResearch=[]),
           "R-F": dict(mgres, category="Tree", requiredResearch=[]),
           "R-G": dict(mgres, category="Tree", requiredResearch=["R-F"]),
           "R-H": dict(mgres, category="Tree", requiredResearch=["R-F"]),
           "R-I": dict(mgres, category="Tree", requiredResearch=["R-G", "R-H"]),
           "R-J": dict(mgres, category="Rev", requiredResearch=["R-K"]),
           "R-K": dict(mgres, category="Rev", requiredResearch=["R-L"]),
           "R-L": dict(mgres, category="Rev", requiredResearch=[]),
           "_config_": {"calculationMode": "improved"}}
    cases.append(("research", {
        "research1.json": jd(res),
        "research2.json": jd({"R-A": dict(mgres, requiredResearch=["R-B"]), "R-B": dict(mgres, requiredResearch=["R-C"]),
                              "R-C": dict(mgres, requiredResearch=["R-A"])}),
        "research3.json": jd({"R-A": dict(mgres, statID="NoSuchStat")}),
        "research4.json": jd({"R-A": dict(mgres, researchPoints=70000)}),
        "research5.json": jd({"R-A": mgres, "_config_": {"calculationMode": "precise"}}),
        "research6.json": jd({"R-A": mgres, "_config_": {"calculationMode": "compat"}})},
        pre + mp_components + ["load\tstructures\tmp/stats/structure.json", "load\tresearch\tstatstest/research1.json", "dump",
                               "load\tresearch\tstatstest/research6.json", "dump",
                               "load\tresearch\tstatstest/research2.json", "dump",
                               "load\tresearch\tstatstest/research3.json", "dump",
                               "load\tresearch\tstatstest/research4.json", "dump",
                               "load\tresearch\tstatstest/research5.json", "dump"]))

    t = jl("mp/stats/templates.json")
    tid = list(t)
    templates = {tid[0]: dict(t[tid[0]], available=True), tid[1]: t[tid[1]],
                 "T-Unknown": dict(t[tid[0]], body="NoSuchBody"),
                 "T-Hydra": dict(t[tid[0]], weapons=["MG1Mk1", "MG2Mk1", "MG3Mk1", "MG1Mk1"], type="WEAPON", available=True),
                 "T-Cyborg": {"name": "Heavy Gunner", "type": "CYBORG", "body": "CyborgLightBody", "propulsion": "CyborgLegs", "weapons": ["CyborgCannon"]},
                 "T-Trans": {"name": "Lift", "type": "TRANSPORTER", "body": "TransporterBody", "propulsion": "V-Tol"},
                 "T-Sensor": {"name": "Eye", "type": "SENSOR", "body": "Body1REC", "propulsion": "wheeled01", "sensor": "SensorTurret1Mk1"},
                 "T-Bad": {"name": "Odd", "type": "SPACESHIP", "body": "Body1REC", "propulsion": "wheeled01", "brain": "CommandBrain01"},
                 "ViperLtMGWheels": dict(t[tid[0]], available=True)}
    cases.append(("templates", {"templates1.json": jd(templates),
                                "names1.txt": b"/* comment */\n// line comment\nT-Sensor\t_(\"Sensor From Strings\")\nT-Plain \"Plain\" ViperLtMGWheels \"Viper\"\n",
                                "names2.txt": b"Broken \"no end\nNext \"x\"\n",
                                "names3.txt": b"Lonely\n",
                                "names4.txt": b""},
                  pre + mp_components + ["strings\tstatstest/names1.txt", "getstring\tT-Plain", "getstring\tT-Sensor",
                                         "players\t5", "selected\t2", "load\ttemplates\tstatstest/templates1.json", "dump",
                                         "strings\tstatstest/names2.txt", "strings\tstatstest/names3.txt", "strings\tstatstest/names4.txt",
                                         "getstring\tNext"]))
    return cases


MODEL_FOLDERS = ["structs", "misc", "effects", "components/prop", "components/weapons", "components/bodies", "features",
                 "misc/micnum", "misc/minum", "misc/mivnum", "misc/researchimds"]


def model_names():
    """The .pie names in modelGet's folders of data/base and data/mp, sorted."""
    names = set()
    for dataset in ("base", "mp"):
        for folder in MODEL_FOLDERS:
            d = os.path.join(build.REF, "data", dataset, *folder.split("/"))
            if os.path.isdir(d):
                names.update(n.lower() for n in os.listdir(d) if n.lower().endswith(".pie"))
    return sorted(names)


def pie(*lines):
    return ("\n".join(lines) + "\n").encode("latin-1")


TRI = ["POINTS 3", "0 0 0", "10.5 0 0", "0 -10.25 7", "POLYGONS 1", "200 3 0 1 2 0 0 1 0 0 1"]


def model_cases():
    """[(name, files {path under the write dir: bytes}, commands)]: crafted models for the loader's paths."""
    files = {
        "zztbadver.pie": b"PIE x\n",
        "zztempty.pie": b"",
        "zztnotpie.pie": pie("PIX 3", "TYPE 200", "LEVELS 1", "LEVEL 1", *TRI),
        "zztver9.pie": pie("PIE 9", "TYPE 200", "LEVELS 1", "LEVEL 1", *TRI),
        "zztnotype.pie": pie("PIE 3", "LEVELS 1", "LEVEL 1", *TRI),
        "zztnolevels.pie": pie("PIE 3", "TYPE 200", "FOO 1"),
        "zztlevelsbad.pie": pie("PIE 3", "TYPE 200", "LEVELS 2", "LEVEL 1", *TRI, "LEVEL 3", *TRI),
        "zztlevelword.pie": pie("PIE 3", "TYPE 200", "LEVELS 1", "LEVL 1", *TRI),
        "zztlevelseof.pie": pie("PIE 3", "TYPE 200", "LEVELS 2", "LEVEL 1", *TRI),
        "zztnolevel0.pie": pie("PIE 3", "TYPE 200", "LEVELS 0"),
        "zzttwolevels.pie": pie("PIE 3", "TYPE 10200", "LEVELS 2", "LEVEL 1", *TRI, "LEVEL 2", "POINTS 4", "1 1 1", "2 2 2", "3 3 3", "-4 5 -6",
                                "POLYGONS 2", "200 3 0 1 2 0.5 0.5 1 0 0 1", "200 3 1 2 3 0 0 0.25 0.75 1 1"),
        "zztpidx.pie": pie("PIE 3", "TYPE 200", "LEVELS 1", "LEVEL 1", "POINTS 3", "0 0 0", "1 0 0", "0 1 0", "POLYGONS 1",
                           "200 3 0 1 3 0 0 1 0 0 1"),
        "zztnpnts.pie": pie("PIE 3", "TYPE 200", "LEVELS 1", "LEVEL 1", "POINTS 3", "0 0 0", "1 0 0", "0 1 0", "POLYGONS 1",
                            "200 4 0 1 2 0 0 1 0 0 1"),
        "zztmaxpoly.pie": pie("PIE 3", "TYPE 200", "LEVELS 1", "LEVEL 1", "POINTS 3", "0 0 0", "1 0 0", "0 1 0", "POLYGONS 9000"),
        "zztnopoints.pie": pie("PIE 3", "TYPE 200", "LEVELS 1", "LEVEL 1", "POLYGONS 1", "200 3 0 1 2 0 0 1 0 0 1"),
        "zztpointseof.pie": pie("PIE 3", "TYPE 200", "LEVELS 1", "LEVEL 1", "POINTS 3", "0 0 0", "1 0 0"),
        "zztpointsbad.pie": pie("PIE 3", "TYPE 200", "LEVELS 1", "LEVEL 1", "POINTS 2", "0 0 0", "1 x 0", "POLYGONS 0"),
        "zztpolyeof.pie": pie("PIE 3", "TYPE 200", "LEVELS 1", "LEVEL 1", "POINTS 3", "0 0 0", "1 0 0", "0 1 0", "POLYGONS 2",
                              "200 3 0 1 2 0 0 1 0 0 1"),
        "zztpolyuv.pie": pie("PIE 3", "TYPE 200", "LEVELS 1", "LEVEL 1", "POINTS 3", "0 0 0", "1 0 0", "0 1 0", "POLYGONS 1",
                             "200 3 0 1 2 0 0 1 0 0"),
        "zztnotex.pie": pie("PIE 3", "TYPE 200", "LEVELS 1", "LEVEL 1", "POINTS 3", "0 0 0", "1 0 0", "0 1 0", "POLYGONS 2",
                            "0 3 0 1 2", "4000 3 0 1 2"),
        "zzttexanim.pie": pie("PIE 2", "TYPE 200", "TEXTURE 0 page-7.png 256 256", "LEVELS 1", "LEVEL 1", "POINTS 4", "0 0 0", "10 0 0",
                              "0 10 0", "10 10 3", "POLYGONS 3", "4200 3 0 1 2 4 100 32 32 0 0 32 0 0 32",
                              "4200 3 1 3 2 5 50 16 16 32 0 32 32 0 32", "200 3 0 1 2 0 0 256 0 0 256"),
        "zzttexbad.pie": pie("PIE 4", "TYPE 200", "TEXTURE 0 page-7.jpg", "LEVELS 1", "LEVEL 1", *TRI),
        "zzttexidx.pie": pie("PIE 4", "TYPE 200", "TEXTURE 5 page-7.png", "LEVELS 1", "LEVEL 1", *TRI),
        "zzttex4.pie": pie("PIE 4", "TYPE 10200", "INTERPOLATE 0", "TEXTURE 0 page-7.png", "TCMASK 0 page-7_tcmask.png",
                           "NORMALMAP 1 page-7-n.png", "SPECULARMAP 2 page-7-s.png", "TEXTURE 2 arizona.png", "LEVELS 2",
                           "LEVEL 1", "TYPE 2", "INTERPOLATE 1", "TEXTURE 1 urban.png", *TRI,
                           "LEVEL 2", "TEXTURE 0 other.png", *TRI),
        "zzttex3.pie": pie("PIE 3", "TYPE 200", "TEXTURE 0 page-7.png 256 256", "NORMALMAP 0 page-7-n.png", "SPECULARMAP 0 s.png",
                           "TCMASK 0 t.png", "LEVELS 1", "LEVEL 1", *TRI),
        "zztcomment.pie": b"PIE 3\nTYPE 200 # the type\nLEVELS 1\nLEVEL 1\nPOINTS 3 # three\n0 0 0\n1 0 0\n0 1 0\r\n"
                          b"POLYGONS 1\r\n200 3 0 1 2 0 0 1 0 0 1\n\n\n  \t  # last",
        "zztmaterials.pie": pie("PIE 3", "TYPE 200", "LEVELS 1", "LEVEL 1", "MATERIALS 1 1 1 1", "SHADERS 2 a b", *TRI),
        "zztmatseof.pie": pie("PIE 3", "TYPE 200", "LEVELS 1", "LEVEL 1", "MATERIALS 1 1 1 1"),
        "zztnodirective.pie": pie("PIE 3", "TYPE 200", "LEVELS 1", "LEVEL 1"),
        "zztunexpected.pie": pie("PIE 3", "TYPE 200", "LEVELS 1", "LEVEL 1", *TRI, "FOO 7", "CONNECTORS 1", "1 2 3"),
        "zztconn.pie": pie("PIE 3", "TYPE 200", "LEVELS 1", "LEVEL 1", *TRI, "CONNECTORS 3", "1 2 3", "1.5 -2.25 3.0", "-4 5 6"),
        "zztconnbad.pie": pie("PIE 3", "TYPE 200", "LEVELS 1", "LEVEL 1", *TRI, "CONNECTORS 2", "1 2 3", "a b c"),
        "zztconneof.pie": pie("PIE 3", "TYPE 200", "LEVELS 1", "LEVEL 1", *TRI, "CONNECTORS 2", "1 2 3"),
        "zztanim.pie": pie("PIE 3", "TYPE 200", "LEVELS 1", "LEVEL 1", *TRI, "ANIMOBJECT 250 3 3",
                           "0 1000 -2000 3500 45000 -90000 12345 1 1 1", "2 0 0 0 0 0 0 0.5 2 1.25", "1 7 8 9 10 11 12 1 1"),
        "zztanimtime.pie": pie("PIE 3", "TYPE 200", "LEVELS 1", "LEVEL 1", *TRI, "ANIMOBJECT 0 1 1", "0 0 0 0 0 0 0 1 1 1"),
        "zztanimbad.pie": pie("PIE 3", "TYPE 200", "LEVELS 1", "LEVEL 1", *TRI, "ANIMOBJECT 10 x", "0 0 0 0 0 0 0 1 1 1"),
        "zztanimeof.pie": pie("PIE 3", "TYPE 200", "LEVELS 1", "LEVEL 1", *TRI, "ANIMOBJECT 10 0 2", "0 0 0 0 0 0 0 1 1 1"),
        "zztshadow.pie": pie("PIE 3", "TYPE 200", "LEVELS 2", "LEVEL 1", *TRI, "SHADOWPOINTS 4", "0 0 0", "1 0 0", "0 1 0", "1 1 1",
                             "SHADOWPOLYGONS 2", "0 3 0 1 2", "0 3 1 2 3", "LEVEL 2", *TRI, "SHADOWPOINTS 2", "0 0 0", "1 1 1"),
        "zztshadowflags.pie": pie("PIE 3", "TYPE 200", "LEVELS 1", "LEVEL 1", *TRI, "SHADOWPOINTS 3", "0 0 0", "1 0 0", "0 1 0",
                                  "SHADOWPOLYGONS 1", "200 3 0 1 2 0 0 1 0 0 1"),
        "zztshadownopts.pie": pie("PIE 3", "TYPE 200", "LEVELS 1", "LEVEL 1", *TRI, "SHADOWPOLYGONS 1", "0 3 0 1 2"),
        "zztshadowzero.pie": pie("PIE 3", "TYPE 200", "LEVELS 1", "LEVEL 1", *TRI, "SHADOWPOINTS 0"),
        "zztnormals.pie": pie("PIE 4", "TYPE 10200", "LEVELS 1", "LEVEL 1", "POINTS 4", "0 0 0", "10 0 0", "0 10 0", "10 10 2",
                              "NORMALS 2", "0 0 1 0 0 1 0 0 1", "0 0.6 0.8 0 0.6 0.8 1 0 0",
                              "POLYGONS 2", "200 3 0 1 2 0 0 1 0 0 1", "200 3 1 3 2 0.5 0.5 0.5 0.5 0.5 0.5"),
        "zztnormalsbad.pie": pie("PIE 4", "TYPE 200", "LEVELS 1", "LEVEL 1", "POINTS 3", "0 0 0", "1 0 0", "0 1 0",
                                 "NORMALS 1", "0 0 1 0 0 1 0 0 1", "POLYGONS 2", "200 3 0 1 2 0 0 1 0 0 1", "200 3 0 2 1 0 0 1 0 0 1"),
        "zztnormalseof.pie": pie("PIE 4", "TYPE 200", "LEVELS 1", "LEVEL 1", "POINTS 3", "0 0 0", "1 0 0", "0 1 0", "NORMALS 1"),
        "zztweld.pie": pie("PIE 3", "TYPE 200", "LEVELS 1", "LEVEL 1", "POINTS 4", "0 0 0", "1 0 0", "0 1 0", "1 1 0", "POLYGONS 3",
                           "200 3 0 1 2 0 0 1 0 0 1", "200 3 1 3 2 1 0 1 1 0 1", "200 3 2 1 0 0 1 1 0 0 0"),
        "zztdegenerate.pie": pie("PIE 3", "TYPE 200", "LEVELS 1", "LEVEL 1", "POINTS 3", "1 1 1", "1 1 1", "2 2 2", "POLYGONS 1",
                                 "200 3 0 1 2 0 0 1 0 0 1"),
        "zztevent.pie": pie("PIE 3", "TYPE 200", "EVENT 1 zzttwolevels.pie", "EVENT 3 nosuchanim.pie", "LEVELS 1", "LEVEL 1", *TRI),
        "zzteventbad.pie": pie("PIE 3", "TYPE 200", "EVENT 1", "LEVELS 1", "LEVEL 1", *TRI),
        "zzteventeof.pie": pie("PIE 3", "TYPE 200", "EVENT 1 zztconn.pie"),
        "zztupper.pie": pie("PIE 3", "TYPE 200", "LEVELS 1", "LEVEL 1", *TRI),
    }
    paths = {}
    for i, (fn, data) in enumerate(sorted(files.items())):
        # spread over modelGet's folders (tried in order)
        paths["%s/%s" % (["structs", "effects", "features", "misc/micnum"][i % 4], fn)] = data
    paths["misc/zztdup.pie"] = pie("PIE 3", "TYPE 200", "LEVELS 1", "LEVEL 1", *TRI)
    paths["features/zztdup.pie"] = pie("PIE 2", "TYPE 200", "LEVELS 1", "LEVEL 1", *TRI)
    cmds = ["debug\tall"] + ["model\t" + fn for fn in sorted(files)] + \
           ["model\tZZTUpper.PIE", "model\tzztdup.pie", "model\tnosuch.pie", "model\tzztconn.pie",
            "reload\tzztevent.pie", "reload\tzztanim.pie", "reload\tzzttwolevels.pie", "reload\tnosuch.pie"]
    return [("models_crafted", paths, cmds)]


MAP_TILESETS = ["texpages/tertilesc1hw", "texpages/tertilesc2hw", "texpages/tertilesc3hw", "-", "texpages/tertilesc9hw"]


def map_blocks():
    """[(name, commands)]: every map folder of the game data loaded as loadGame loads it (each with a tileset in turn,
    some with the terrain type overrides or a ground type reload), queried, its features placed (with mp's stats,
    which the first block loads) and shut down."""
    data = os.path.join(build.REF, "data")
    folders = []
    for folder, _, names in os.walk(data):
        if "game.map" in names or "game.js" in names:
            folders.append(folder)
    blocks = [("map_stats", [c for c in stats_set_commands("mp", "3") if c != "dump"])]
    for i, folder in enumerate(sorted(folders)):
        rel = os.path.relpath(folder, data).replace("\\", "/")
        path = rel.split("/", 1)[1] + "/"               # under data/<dataset>/, which the search path mounts
        name = os.path.basename(folder)
        if "multiplay" in rel.split("/"):
            kind, players = "skirmish", name.split("c-")[0]
        else:
            kind, players = "campaign", "4"
        cmds = ["debug\tall", "overrides\t%d" % (i % 4 == 1),
                "mapload\t%s\t%s\t%s\t%s" % (path, MAP_TILESETS[i % len(MAP_TILESETS)], players, kind), "mapqueries"]
        if i % 7 == 0:
            cmds.append("mapreloadground")
        cmds += ["mapplayers\t%s\t%s" % (kind, players), "syncreset", "mapobjects", "synclog", "mapgrid", "mappower", "synclog",
                 "mappaths", "synclog", "mapmove", "synclog", "maporders", "synclog", "mapcombat", "synclog",
                 "mapbase", "synclog", "mapnet", "synclog"]
        if kind == "skirmish" and i % 3 == 0:
            # a game recorded as a replay on the map loaded afresh (the earlier blocks' structures levelled its ground),
            # then played back on the map loaded afresh again
            fresh = ["objshutdown", "mapshutdown", "mapload\t%s\t%s\t%s\t%s" % (path, MAP_TILESETS[i % len(MAP_TILESETS)], players, kind),
                     "mapplayers\t%s\t%s" % (kind, players), "mapobjects"]
            cmds += fresh + ["maprecord", "synclog"] + fresh + ["mapplay", "synclog"]
        cmds += ["objshutdown", "mapshutdown"]
        blocks.append(("map_%03d" % i, cmds))
    return blocks


def stats_tests(keep):
    """Warzone.exe --vbtest against statsref on the stats of both data sets and the crafted cases; True when all match."""
    import re
    if not build_statsref():
        return False
    work = os.path.join(build.BUILD, "stats_tests")
    if os.path.exists(work):
        shutil.rmtree(work)
    cfg = os.path.join(work, "config")
    for d in (os.path.join(cfg, "statstest"), os.path.join(work, "vb"), os.path.join(work, "ref")):
        os.makedirs(d)
    blocks = [("set_base", stats_set_commands("base")), ("set_mp", stats_set_commands("mp", "3"))]
    for name, files, cmds in stats_cases():
        for fn, content in files.items():
            with open(os.path.join(cfg, "statstest", fn), "wb") as f:
                f.write(content)
        blocks.append(("case_" + name, cmds))
    # the models: every one in the data (100 to a block), then the crafted ones (their folders in the write dir)
    names = model_names()
    for k in range(0, len(names), 100):
        blocks.append(("models_%02d" % (k // 100), ["debug\tall"] + ["model\t" + n for n in names[k:k + 100]]))
    blocks += map_blocks()
    blocks.append(("factions", ["factions\t1", "factions\t0", "playerfaction\t3\t1", "playerfaction\t20\t2", "factions\t1", "factions\t0",
                                "playerfaction\t3\t2", "factions\t1", "factionmodel\t1\tblhq.pie", "factionmodel\t2\tblhowmnt.pie",
                                "factionmodel\t2\tblhq.pie", "factionmodel\t0\tblhq.pie", "factionmodel\t1\tnosuch.pie",
                                "playerfaction\t3\t0", "playerfaction\t20\t0", "factions\t0"]))
    for name, files, cmds in model_cases():
        for rel, content in files.items():
            p = os.path.join(cfg, *rel.split("/"))
            os.makedirs(os.path.dirname(p), exist_ok=True)
            with open(p, "wb") as f:
                f.write(content)
        blocks.append((name, cmds))

    # the official 4.7.0 build's replays (ref/official/replays, when present): read as the game reads them
    official = os.path.join(build.REF, "..", "official", "replays")
    if os.path.isdir(official):
        for k, fn in enumerate(sorted(f for f in os.listdir(official) if f.endswith(".wzrp"))):
            for d in (os.path.join(cfg, "replay", "official"), os.path.join(work, "refwrite", "replay", "official")):
                os.makedirs(d, exist_ok=True)
                shutil.copyfile(os.path.join(official, fn), os.path.join(d, fn))
            blocks.append(("replay_official_%02d" % k, ["debug\tall", "replayinfo\treplay/official/" + fn]))

    def command_file(side, first):
        lines = [first]
        for name, cmds in blocks:
            lines.append("output\t" + os.path.join(work, side, name + ".txt"))
            lines += cmds
        path = os.path.join(work, "commands_%s.txt" % side)
        with open(path, "w", encoding="utf-8", newline="\n") as f:
            f.write("\n".join(lines) + "\n")
        return path

    sp_file = os.path.join(work, "searchpath.txt")
    vbcmd = command_file("vb", "searchpath\t" + sp_file)
    import time
    t = time.time()
    r = subprocess.run([os.path.join(build.BIN, "Warzone.exe"), "--configdir=" + cfg, "--vbtest=" + vbcmd], capture_output=True)
    print("Warzone.exe --vbtest: %.1f s, exit code %d" % (time.time() - t, r.returncode))
    if not os.path.exists(sp_file):
        print("no search path from Warzone.exe:\n" + r.stderr.decode("utf-8", "replace")[-3000:])
        return False
    with open(sp_file, encoding="utf-8") as f:
        searchpath = f.read().strip()
    refcmd = command_file("ref", "searchpath\t" + searchpath + "\nwritedir\t" + os.path.join(work, "refwrite"))
    if not run([STATSREF_EXE, refcmd]):
        return False

    stamp = re.compile(rb"([a-z0-9]+ *\|)\d\d:\d\d:\d\d: ")
    # debug.cpp's fatal error text names the log file (statsref writes none)
    logfile = re.compile(rb"\(([^()\n]*WZlog-[0-9_]+\.txt)?\)|upload the ([^ \n]*WZlog-[0-9_]+\.txt)? file")
    # std::stoi and co.'s what(): the Microsoft library's texts (the official build's, the port's) and libstdc++'s (statsref's)
    whats = [(b"invalid %s argument" % f, f) for f in (b"stoi", b"stoul", b"stoull", b"stod", b"stof")] + \
            [(b"%s argument out of range" % f, f) for f in (b"stoi", b"stoul", b"stoull", b"stod", b"stof")]

    # a %p in a log line: a heap address in C++, an object handle in VB ("Killing off feature ... (%p)",
    # "created droid for player %d, droid = %p, ...", "%s droid %d (%p) is toast")
    pointer = re.compile(rb"\((?:[0-9a-fA-F]{16}|[0-9a-fA-F]{8})\)$", re.M)
    droidPointer = re.compile(rb"droid = (?:[0-9a-fA-F]{16}|[0-9a-fA-F]{8}),")
    droidIdPointer = re.compile(rb"(droid \d+) \((?:[0-9a-fA-F]{16}|[0-9a-fA-F]{8})\)")
    # objmemDestroy's "freeing droid at %p"
    atPointer = re.compile(rb" at (?:[0-9a-fA-F]{16}|[0-9a-fA-F]{8})$", re.M)
    # the messages' view data (visibility.cpp's "Added message for oil well or artefact, pViewData=%p", message.cpp)
    viewDataPointer = re.compile(rb"pViewData=(?:[0-9a-fA-F]{16}|[0-9a-fA-F]{8})")
    # gameStateUpdate's version line (outside the sync log's CRC)
    clientVersion = re.compile(rb"My client version = [^\r\n]*")
    # the replay settings' version string, and the local time in a replay's file name
    replayVersion = re.compile(rb'"versionString":"[^"]*"')
    replayName = re.compile(rb"replay/(\w+)/\d{8}_\d{6}_")
    # a replay file's size, which the version string's length changes (its content is compared line by line)
    replaySize = re.compile(rb"^replay \d+ bytes, magic", re.M)
    # eventResearchedHandleUpgrades' "grade" lines walk a std::unordered_map: their order is the library's hash order
    # (libstdc++'s in statsref, Microsoft's in the official build); the port keeps first-appearance order. Each run
    # of them is compared sorted.
    gradeLine = re.compile(rb"\[eventResearchedHandleUpgrades:\d+\] +Player\[\d+\], \w+\[[^\]]*\] grade: \d+$")

    def sortGradeRuns(data):
        out, run = [], []
        for ln in data.split(b"\n"):
            if gradeLine.search(ln):
                run.append(ln)
                continue
            out += sorted(run)
            run = []
            out.append(ln)
        return b"\n".join(out + sorted(run))

    def normalised(data):
        data = stamp.sub(rb"\1", data)
        data = pointer.sub(b"(%p)", data)
        data = droidPointer.sub(b"droid = %p,", data)
        data = droidIdPointer.sub(rb"\1 (%p)", data)
        data = atPointer.sub(b" at %p", data)
        data = viewDataPointer.sub(b"pViewData=%p", data)
        data = clientVersion.sub(b"My client version = %s", data)
        data = replayVersion.sub(b'"versionString":"%s"', data)
        data = replayName.sub(rb"replay/\1/YYYYMMDD_HHMMSS_", data)
        data = replaySize.sub(b"replay %d bytes, magic", data)
        data = logfile.sub(lambda m: b"()" if m.group(0).startswith(b"(") else b"upload the  file", data)
        for msvc, gnu in whats:
            data = data.replace(b"error: " + msvc, b"error: " + gnu)
        return sortGradeRuns(data)

    failed = []
    for name, _ in blocks:
        outs = []
        for side in ("ref", "vb"):
            p = os.path.join(work, side, name + ".txt")
            outs.append(normalised(open(p, "rb").read()) if os.path.exists(p) else b"<missing>")
        if outs[0] != outs[1]:
            failed.append((name, outs[0], outs[1]))
    print("stats, models and maps: %d of %d outputs match" % (len(blocks) - len(failed), len(blocks)))
    for name, ref, vb in failed[:8]:
        rl, vl = ref.split(b"\n"), vb.split(b"\n")
        k = next((k for k in range(min(len(rl), len(vl))) if rl[k] != vl[k]), min(len(rl), len(vl)))
        print("\nMISMATCH %s, line %d (lines %d / %d)" % (name, k + 1, len(rl), len(vl)))
        for j in range(max(0, k - 2), min(k + 3, max(len(rl), len(vl)))):
            print("  ref: %r" % (rl[j] if j < len(rl) else b"<end>"))
            print("  vb:  %r" % (vl[j] if j < len(vl) else b"<end>"))
    if not failed and not keep:
        shutil.rmtree(work)
    return not failed


def map_commands():
    """(mode, argument, output) for every map folder of the game data"""
    data = os.path.join(build.REF, "data")
    folders = []
    for folder, _, names in os.walk(data):
        if "game.map" in names or "game.js" in names:
            folders.append(folder)
    commands = []
    for i, folder in enumerate(sorted(folders)):
        name = os.path.basename(folder)
        if "multiplay" in folder.replace("\\", "/").split("/"):
            players = name.split("c-")[0]
            primary, other = "skirmish|%s" % players, "campaign|8"
        else:
            primary, other = "campaign|4", "skirmish|7"
        commands.append(("map", "%s|%s" % (folder, primary), "map%03d.dump" % i))
        commands.append(("map", "%s|%s" % (folder, other), "map%03d.other.dump" % i))
        for fmt in range(3):
            commands.append(("export", "%s|%s|%d" % (folder, primary, fmt), "map%03d.f%d" % (i, fmt)))
    return commands, len(folders)


def export_files(root, out):
    """{file name: bytes} of an export: its log (output root replaced by <root>) and the files written"""
    files = {}
    log = os.path.join(root, out + ".log")
    files[".log"] = open(log, "rb").read().replace(root.encode("utf-8"), b"<root>") if os.path.exists(log) else b"<missing>"
    folder = os.path.join(root, out)
    for n in sorted(os.listdir(folder)):
        files[n] = open(os.path.join(folder, n), "rb").read()
    return files


def build_jsonref():
    if os.path.exists(JSONREF_EXE) and os.path.getmtime(JSONREF_EXE) >= os.path.getmtime(JSONREF_SRC):
        return True
    print("building jsonref.exe (g++, nlohmann::json from ref/warzone2100) ...")
    return run([GXX, "-std=c++17", "-O1", "-static", "-I", JSON_INCLUDE, "-o", JSONREF_EXE, JSONREF_SRC])


def doubles_file(path):
    rnd = random.Random(2100)
    values = [0.0, -0.0, 1.0, -1.0, 0.1, 0.2, 0.3, 1.0 / 3, 2.0 / 3, 100.0, 1e15, 1e16, 1e17, 123456789012345.0,
              1234567890123456.0, 9007199254740993.0, 5e-324, 2.2250738585072014e-308, 1.7976931348623157e308,
              0.0001, 0.00001, 1e-7, 1e21, 1e22, 1e23, 4.35, 0.5, 1.5, 2.675, 1e300, 1e-300, 3.14159, 2.71828]
    values += [struct.unpack("<f", struct.pack("<f", v))[0] for v in (0.1, 0.2, 1.1, 3.3, 1e-3, 123.456, 0.7)]
    values += [10.0 ** e for e in range(-20, 25)]
    for _ in range(1500):                           # random bit patterns (finite)
        bits = rnd.getrandbits(64)
        if (bits >> 52) & 0x7FF == 0x7FF:
            continue
        values.append(struct.unpack("<d", struct.pack("<Q", bits))[0])
    for _ in range(1500):                           # "everyday" numbers
        values.append(round(rnd.uniform(-1000, 1000), rnd.randint(0, 6)))
    with open(path, "w", newline="\n") as f:
        for v in values:
            f.write("%016X\n" % struct.unpack("<Q", struct.pack("<d", v))[0])
    return len(values)


MEDIA_TESTS = ("zip_test", "celt_test", "silk_test", "opus_test", "theora_test", "video_audio_test",
               "zlib_test", "crypto_test", "net_test", "http_test", "map_script_test", "zstd_test",
               "ktx2_test")
# where they are: tools/, or beside their test exe's sources (silk_test.py in src/silktests)
MEDIA_TEST_DIRS = (os.path.join(ROOT, "tools"), os.path.join(ROOT, "src", "silktests"))


def media_tests(keep):
    """Cluster K's tests, each tools/<name>.py's run(keep) when the file exists: the .wz archives read through
    WzFramework's PhysFS against Python's zipfile (zip_test), the Opus decoder (opusfile + libopus) against
    ffmpeg's libopus (opus_test), the Theora decoder against ffmpeg and the official libtheoradec
    (theora_test), the videos' Vorbis audio through the packet API against ffmpeg (video_audio_test)."""
    import importlib
    passed = True
    for name in MEDIA_TESTS:
        for folder in MEDIA_TEST_DIRS:
            if os.path.exists(os.path.join(folder, name + ".py")):
                if folder not in sys.path:
                    sys.path.append(folder)
                passed = importlib.import_module(name).run(keep) and passed
                break
    return passed


def main(argv):
    keep = "--keep" in argv
    if "--sound" in argv:
        return 0 if sound_tests(keep) else 1
    if "--media" in argv:
        return 0 if media_tests(keep) else 1
    for name in ("WzFramework", "WzMap"):
        if not os.path.exists(build.dll_path(name)):
            print("bin/%s.dll is missing: run tools\\build.py first" % name)
            return 1
    if "--stats" in argv:
        return 0 if stats_tests(keep) else 1
    if "--replays" in argv:
        return 0 if replay_tests(keep) else 1
    if not build_tests_exe() or not build_jsonref() or not build_mapref():
        return 1

    if os.path.exists(WORK):
        shutil.rmtree(WORK)
    cases_dir = os.path.join(WORK, "cases")
    for d in (cases_dir, os.path.join(WORK, "vb"), os.path.join(WORK, "ref")):
        os.makedirs(d)

    commands = []
    data = os.path.join(build.REF, "data")
    files = []
    for folder, _, names in os.walk(data):
        for n in sorted(names):
            if n.lower().endswith(".json"):
                files.append(os.path.join(folder, n))
    wz_commands = []
    for i, path in enumerate(sorted(files)):
        for mode in ("dump", "dumpo", "compact", "ascii"):
            commands.append((mode, path, "data%04d.%s" % (i, mode)))
        if groups_only(path):
            wz_commands.append(("wzread", path, "data%04d.wzread" % i))
            wz_commands.append(("wzcopy", path, "data%04d.wzcopy" % i))
    for name, content, modes in CASES:
        path = os.path.join(cases_dir, name + ".json")
        with open(path, "wb") as f:
            f.write(content)
        for mode in modes.split():
            commands.append((mode, path, "case_%s.%s" % (name, mode)))
    dpath = os.path.join(cases_dir, "doubles.txt")
    ndoubles = doubles_file(dpath)
    commands.append(("doubles", dpath, "doubles.out"))

    cmdfile = os.path.join(WORK, "commands.txt")
    with open(cmdfile, "w", newline="\n", encoding="utf-8") as f:
        for mode, path, out in commands:
            f.write("%s\t%s\t%s\n" % (mode, path, out))
    map_cmds, nmaps = map_commands()
    mapcmdfile = os.path.join(WORK, "commands_map.txt")
    with open(mapcmdfile, "w", newline="\n", encoding="utf-8") as f:
        for mode, arg, out in map_cmds:
            f.write("%s\t%s\t%s\n" % (mode, arg, out))
    for mode, arg, out in map_cmds:
        if mode == "export":                    # Map::exportMapToPath writes into an existing folder
            os.makedirs(os.path.join(WORK, "ref", out))
            os.makedirs(os.path.join(WORK, "vb", out))
    vbcmdfile = os.path.join(WORK, "commands_vb.txt")
    with open(vbcmdfile, "w", newline="\n", encoding="utf-8") as f:
        for mode, path, out in commands + wz_commands + map_cmds:
            f.write("%s\t%s\t%s\n" % (mode, path, out))

    print("%d JSON files, %d cases, %d doubles, %d WzConfig files, %d map folders: %d commands"
          % (len(files), len(CASES), ndoubles, len(wz_commands) // 2, nmaps,
             len(commands) + len(wz_commands) + len(map_cmds)))
    if not run([JSONREF_EXE, cmdfile, os.path.join(WORK, "ref")]):
        return 1
    if not run([MAPREF_EXE, mapcmdfile, os.path.join(WORK, "ref")]):
        return 1
    import time
    t = time.time()
    if not run([os.path.join(build.BIN, "WzTests.exe"), vbcmdfile, os.path.join(WORK, "vb")]):
        return 1
    # the wz modes' expected output: Python's conversions, and jsonref's dump of the original
    for mode, path, out in wz_commands:
        if mode == "wzread":
            data = wzread_expected(path)
        else:
            with open(os.path.join(WORK, "ref", out.replace(".wzcopy", ".dump")), "rb") as f:
                data = f.read()
        with open(os.path.join(WORK, "ref", out), "wb") as f:
            f.write(data)
    commands += wz_commands
    print("WzTests.exe: %.1f s" % (time.time() - t))

    failed = []
    nexport_files = 0
    for mode, path, out in map_cmds:
        if mode != "export":
            continue
        ref_files = export_files(os.path.join(WORK, "ref"), out)
        vb_files = export_files(os.path.join(WORK, "vb"), out)
        for n in sorted(set(ref_files) | set(vb_files)):
            nexport_files += 1
            ref, vb = ref_files.get(n, b"<missing>"), vb_files.get(n, b"<missing>")
            if ref != vb:
                failed.append((mode, path.split("|")[0], out + "/" + n, ref, vb))
    commands += [c for c in map_cmds if c[0] != "export"]
    for mode, path, out in commands:
        with open(os.path.join(WORK, "ref", out), "rb") as f:
            ref = f.read()
        vbpath = os.path.join(WORK, "vb", out)
        vb = open(vbpath, "rb").read() if os.path.exists(vbpath) else b"<missing>"
        if mode == "error":
            ref, vb = invalid_utf8_as_one_mark(ref), invalid_utf8_as_one_mark(vb)
        if ref != vb:
            failed.append((mode, path, out, ref, vb))
    total = len(commands) + nexport_files
    print("%d of %d outputs match" % (total - len(failed), total))
    for mode, path, out, ref, vb in failed[:12]:
        print("\nMISMATCH %s %s (%s)" % (mode, os.path.relpath(path.split("|")[0], ROOT), out))
        if mode == "doubles":
            rl, vl = ref.split(b"\n"), vb.split(b"\n")
            for k in range(min(len(rl), len(vl))):
                if rl[k] != vl[k]:
                    print("  line %d: ref %r  vb %r" % (k + 1, rl[k], vl[k]))
                    break
        else:
            k = next((k for k in range(min(len(ref), len(vb))) if ref[k] != vb[k]), min(len(ref), len(vb)))
            print("  first difference at byte %d (sizes %d / %d)" % (k, len(ref), len(vb)))
            print("  ref: %r" % ref[max(0, k - 60):k + 60])
            print("  vb:  %r" % vb[max(0, k - 60):k + 60])
    if not failed and not keep:
        shutil.rmtree(WORK)
    stats_ok = stats_tests(keep)
    replays_ok = replay_tests(keep)
    sound_ok = sound_tests(keep)
    media_ok = media_tests(keep)
    return 1 if failed or not stats_ok or not replays_ok or not sound_ok or not media_ok else 0


def replay_tests(keep):
    """tools/replay_test.py on every official replay with reference sync logs (ref/official): Warzone.exe plays
    it headless and its sync log must equal the official 4.7.0 build's, tick by tick."""
    import replay_test
    names = sorted(os.listdir(os.path.join(replay_test.OFFICIAL, "synclogs")))
    ok = all([replay_test.run(n, None, keep, 900) for n in names])
    print("official replays: %d replay(s), %s" % (len(names), "all ticks match" if ok else "MISMATCH"))
    return ok


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
