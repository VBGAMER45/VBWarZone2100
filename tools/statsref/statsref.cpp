// statsref: reference output of Warzone 2100 4.7.0's stats loaders for the
// VB port's tests (tools/run_tests.py builds it with g++ from the real
// sources, in part by line ranges, and compares its output with
// Warzone.exe --vbtest's).
//
//   statsref <command file>
//
// The command file is the one Warzone.exe --vbtest reads (src/game/vbtest_vb.bas):
// one command per line, its words separated by tabs:
//   searchpath  <folder>|<folder>|...   the PhysFS search path (statsref_physfs.cpp)
//   writedir    <folder>                the PhysFS write folder, first in the search path (statsref only:
//                                       Warzone.exe --vbtest writes to its config folder)
//   output      <file>                  the file the following commands write to
//   debug       <part>                  debug_enable_switch(part)
//   selected    <player>                selectedPlayer
//   players     <mask>                  NetPlay.players[i].allocated = bit i
//   gamemode    <mode>                  GetGameMode()
//   reset                               every stats store released, the object ids reset
//   strings     <file>                  data.cpp's dataStrResLoad
//   getstring   <id>                    strresGetString
//   load        <kind> <file>           data.cpp's loader of a stats file (kind: weapons, body, ...)
//   dump                                every store, as text
//   model       <name>                  imdload.cpp's modelGet, and the model as text
//   reload      <name>                  debugReloadDisplayModelsForBaseModel, and the model
//   playerfaction <player> <faction>    NetPlay.players[player].faction
//   factions    <ignoreNormal>          getEnabledFactions (in id order)
//   factionmodel <faction> <name>       getFactionModelName
//   overrides   <0|1>                   map.cpp's useTerrainOverrides
//   mapload     <folder> <tilesetDir or -> <players> <skirmish|campaign>
//                                       loadGame's map part (the map through wzmaplib, its
//                                       terrain types, mapLoadFromWzMapData), and the map
//   mapreloadground                     mapReloadGroundTypes, and the map
//   mapqueries                          map_Height, getTileMaxMin, map_LineIntersect,
//                                       map_Intersect, fpathBlockingTile and tileSetFire on
//                                       points from a fixed generator
//   mapshutdown                         gwShutDown, mapShutdown
//   mapplayers  <skirmish|campaign> <n> game.type, game.maxPlayers = n, no scavengers or alliances,
//                                       bMultiPlayer for a skirmish, NetPlay.players[i].position = i
//                                       for i < n (else -1)
//   mapobjects                          loadGame's object set-up (objInitialise, grpInitialise,
//                                       aiInitialise, scriptInit, visInitialise, initFactoryNumFlag,
//                                       droidInit; gameSRand(12345) first), then loadWzMapDroidInit,
//                                       loadWzMapFeature and loadWzMapStructure on the map mapload
//                                       loaded as loadGame orders them, and the objects, flags and
//                                       tile visibility
//   objshutdown                         freeAllDroids, freeAllStructs, freeAllFeatures, the function lists cleared,
//                                       freeAllFlagPositions, the destroyed list cleared,
//                                       objmemShutdown
//   syncreset                           resetSyncDebug
//   syncdumpall <0|1>                   mapmove, maporders, mapcombat and mapbase dump the whole sync log of every tick
//                                       (for finding a divergence; the tests do not use it)
//   synclog                             the sync log since the last synclog (dumpLocalDebugSyncLog), then
//                                       nextDebugSync and its CRC
//   mapgrid                             gridInitialise, gridReset, the gridStartIterate queries on points from
//                                       a fixed generator (each object found by the unseen query in turn marked
//                                       as seen), gridShutDown
//   mappower                            allocPlayerPower, then 40 ticks of updatePlayerPower with power requests,
//                                       added and used power and a derrick of each player dying; the power each tick
//   mappaths                            fpathInitialise; over 6 ticks, routes from points of a fixed generator, each
//                                       destination wanted from 4 origins (fpathRoute: the job, then its result);
//                                       fpathDroidRoute for the droids; fpathCheck; fpathShutdown
//   mapmove                             fpathInitialise, gridInitialise; every droid stopped (moveReallyStopDroid, initDroidMovement);
//                                       up to 48 droids sent to points of a fixed generator (moveDroidTo, three to
//                                       a destination; moveDroidToNoFormation, moveDroidToDirect); 80 ticks of
//                                       gridReset and moveUpdateDroid for every droid, with stops, turns and new
//                                       orders on the way; each tick's sync log CRC (the whole log of ticks 1 and
//                                       80), the droids every 20 ticks, the features and power; gridShutDown,
//                                       fpathShutdown
//   maporders                           fpathInitialise, gridInitialise, allocPlayerPower; 16 structure types buildable
//                                       for every player, each player's power set; every droid's actionStarted and
//                                       actionPoints 0 (DROID's constructor leaves them uninitialised); a third of the droids and a
//                                       quarter of the built structures damaged; secondary orders; then for up to
//                                       64 droids an order of a fixed generator by droid type (build, line build,
//                                       demolish, repair, droid repair, observe, attack, guard, move, scout,
//                                       patrol, hold, RTB, RTR, stop, recycle, chooseOrderObj, queued orders;
//                                       VTOLs circle, patrol, attack, RTB); 90 ticks of gridReset, fpathUpdate,
//                                       droidUpdate for every droid (as the game loop: each list, mutating) and
//                                       updatePlayerPower, with new orders at ticks 20 and 40; each tick's sync log
//                                       CRC (the whole log of ticks 1 and 90), the droids' orders every 30 ticks,
//                                       the structures being built or damaged, the counts, features and power
//   mapcombat                           fpathInitialise, gridInitialise, resetDamageModifiers, proj_InitSystem; up to 40
//                                       armed droids moved 3 to 6 tiles from an enemy structure of a fixed generator
//                                       (droidSetPosition) and ordered to attack it, scout to it, hold or stop; 60
//                                       ticks as the game loop orders them (visUpdateLevel, gridReset,
//                                       processVisibility, fpathUpdate, cmdDroidUpdate, per player updatePlayerPower
//                                       and droidUpdate, proj_UpdateAll, objmemUpdate); each tick's sync log CRC and
//                                       projectile count (the whole log of ticks 1 and 60), the fighters every 20
//                                       ticks, the damaged structures, the counts and features; proj_FreeAllProjectiles
//   mapbase                             gameTime 20000, game.map "mapbase"; fpathInitialise, gridInitialise, resetDamageModifiers,
//                                       proj_InitSystem, mapInit; the research
//                                       state and the stats' upgrades saved; research and production 50 times faster,
//                                       every component available, each player's power and droid limits (rules.js's
//                                       150, 10 commanders, 15 trucks) set; VTOLs to rearm
//                                       (moveToRearm), damaged droids (one in two, a fixed generator) ordered to RTR;
//                                       each idle lab given a topic (recvResearchStatus's start) and each idle factory
//                                       a template (structSetManufacture); 100 ticks of gameStateUpdate, new work at
//                                       every 25th; each tick's sync log CRC, projectile and droid counts (the whole
//                                       log of ticks 1 and 100), the working and armed structures every 25 ticks, the
//                                       counts, research done and power; the saved state restored
//   mapnet                              a skirmish host's game without a network, after mapbase: the same set-up (game.map
//                                       "mapnet"), the game queues initialised (init.cpp), gameTimeInit and gameTime 40000; each
//                                       idle lab given a topic it may start (sendResearchStatus), each idle factory a template
//                                       (structSetManufacture, ModeQueue), the first busy lab and factory of each player
//                                       held, released or cancelled through the messages, up to 24 droids ordered and some
//                                       given secondary orders through the messages (ModeQueue); 80 frames of realTimeUpdate
//                                       and gameLoop on a clock (wzGetTicks) that moves 100 ms a frame, new work every 20th,
//                                       the debug messages (destroy a droid and a structure, add a droid, finish a structure)
//                                       at the 40th; each frame's gameTime, projectile and droid counts and the header of the
//                                       last closed sync log (the whole log of frames 2 and 80), the structures every 20
//                                       frames, the counts, research done and power; the saved state restored
//   replayinfo  <file>                  a replay read as the game reads one (NETreplayLoadStart with the tests' handler, then
//                                       every message): the options as restored (saved again by WZGameReplayOptionsHandler),
//                                       the messages' count by player and type, their CRC, the first ten in hex
//   maprecord                           mapnet's game for 40 frames (new work every 10th, the debug messages at the 20th),
//                                       recorded as a replay (gameSRand(4242), then NETreplaySaveStart; NETreplaySaveStop
//                                       after): the frames as mapnet's, then the replay file as text (its settings JSON,
//                                       the messages' count by type and CRC, the end-of-game JSON)
//   mapplay                             that replay played back (NETloadReplay with StatsrefReplayOptionsHandler: the game's
//                                       restoreOptions without the level part; a replay viewer in the spectator slot, one
//                                       tick a frame): 40 frames as mapnet's, without new work
// Each command writes "> " and its line, the log lines it caused (Warzone's
// debug.cpp, through a callback), and its results.
//
// The stand-ins (as src/game/stubs_game.bas has them): getViewData gives
// nullptr; the lightmap, the terrain textures and the danger
// thread's calls do nothing; wzmaplib's runMapScript is map_script.cpp over QuickJS (as in mapref); the strres grammar (strres_lexer.lpp, strres_parser.ypp: no flex or
// bison here) is written out again below. imdload.cpp's GPU buffers are
// RecordingBuffers (the extract calls statsref_create_buffer_object), which
// keep what is uploaded, as the VB port keeps it in the shape (vbo_*).
// What the objects call in modules of later clusters (effects, scores,
// messages, sound, scripts, the interface, the targeting and combat, the
// transporters, the network messages, the rest of the factories'
// commanders, the research) does nothing, as in stubs_game.bas.
#include "lib/framework/frame.h"
#include "lib/framework/wzconfig.h"
#include "lib/framework/strres.h"
#include "lib/framework/strresly.h"
#include "lib/framework/lexer_input.h"
#include "lib/framework/physfs_ext.h"
#include "lib/framework/wzapp.h"
#include "lib/framework/wztime.h"
#include "lib/framework/trig.h"
#include "lib/ivis_opengl/imd.h"
#include "lib/ivis_opengl/ivisdef.h"
#include "lib/ivis_opengl/pielighting.h"
#include "lib/framework/frameresource.h"
#include "lib/netplay/sync_debug.h"
#include "src/map.h"
#include "src/faction.h"
#include "src/gateway.h"
#include "src/fpath.h"
#include "src/texture.h"
#include "src/init.h"
#include "lib/gamelib/gtime.h"
#include "lib/netplay/netplay.h"
#include "src/stats.h"
#include "src/structure.h"
#include "src/feature.h"
#include "src/research.h"
#include "src/template.h"
#include "src/mechanics.h"
#include "src/text.h"
#include "src/objmem.h"
#include "src/hci.h"
#include "src/main.h"
#include "src/message.h"
#include "src/design.h"
#include "src/droid.h"
#include "src/objects.h"
#include "src/group.h"
#include "src/ai.h"
#include "src/qtscript.h"
#include "src/mission.h"
#include "src/visibility.h"
#include "src/effects.h"
#include "src/scores.h"
#include "src/order.h"
#include "src/formation.h"
#include "src/multiplay.h"
#include "src/power.h"
#include "src/mapgrid.h"
#include "src/fpath.h"
#include "src/move.h"
#include "src/action.h"
#include "src/console.h"
#include "src/component.h"
#include "src/edit3d.h"
#include "src/multigifts.h"
#include "src/projectile.h"
#include "src/campaigninfo.h"
#include "src/formationdef.h"
#include "lib/framework/math_ext.h"
#include "src/wzcrashhandlingproviders.h"
#include "lib/sound/audio.h"
#include <unordered_set>
#include "wzmaplib/terrain_type.h"
#include "lib/framework/object_list_iteration.h"
#include "lib/framework/input.h"
#include "lib/widget/widget.h"
#include "src/transporter.h"
#include "src/warcam.h"
#include "src/display.h"
#include "src/combat.h"
#include "src/lighting.h"
#include "src/difficulty.h"
#include "src/cmddroid.h"
#include "src/wzapi.h"
#include "src/multistat.h"
#include "src/miscimd.h"
#include "src/version.h"
#include "src/loop.h"
#include "src/intdisplay.h"
#include "lib/netplay/nettypes.h"
#include "lib/netplay/netreplay.h"
#include "lib/framework/crc.h"
#include "lib/ivis_opengl/piemode.h"
#include "src/scores.h"
#include "src/multirecv.h"
#include "src/data.h"
#include "src/levels.h"
#include "src/frontend.h"
#include "src/component.h"
#include "lib/ivis_opengl/piepalette.h"
#include "lib/framework/file.h"
#include "src/activity.h"
#include "src/design.h"

#include <algorithm>
#include <cstdarg>
#include <cstdio>
#include <fstream>
#include <map>
#include <sstream>
#include <string>
#include <vector>

void statsref_set_searchpath(const std::string &list);   // statsref_physfs.cpp
void statsref_set_writedir(const std::string &dir);      // statsref_physfs.cpp
extern uint32_t unsynchObjID;                             // objmem.cpp
extern uint32_t synchObjID;
extern unsigned syncDebugNext;                            // sync_debug.cpp (not static: see run_tests.py)
extern std::vector<uint8_t> debugSyncTmpBuf;
size_t dumpLocalDebugSyncLog(unsigned logIndex);
FPATH_RETVAL fpathRoute(MOVE_CONTROL *psMove, unsigned id, int startX, int startY, int tX, int tY, PROPULSION_TYPE propulsionType,
                        DROID_TYPE droidType, FPATH_MOVETYPE moveType, int owner, bool acceptNearest, StructureBounds const &dstStructure);  // fpath.cpp (not static: see run_tests.py)

// ---------------------------------------------------------------------------
// stand-ins for what the extracted code calls outside the extracted ranges

NETPLAY NetPlay;
NETPLAY::NETPLAY()
{
	players.resize(MAX_CONNECTED_PLAYERS);
	playercount = 0;
	hostPlayer = 0;
	bComms = false;
	isHost = true;
	isPortMappingEnabled = false;
	isHostAlive = false;
	gamePassword[0] = '\0';
	GamePassworded = false;
}

std::list<DROID_TEMPLATE> localTemplates;

// netplay's WZFile (a PLAYER member) and frame.cpp's saveFile (WzConfig's destructor): nothing is downloaded or written
WZFile::~WZFile()
{
}

bool saveFile(const char *pFileName, const char *pFileData, UDWORD fileSize)
{
	return false;
}

static GS_GAMEMODE gameMode = GS_TITLE_SCREEN;
GS_GAMEMODE GetGameMode()
{
	return gameMode;
}

VIEWDATA *getViewData(const WzString &name)
{
	return nullptr;
}

// what map.cpp, random.cpp and gateway.cpp call outside the extracted code
void LightMap::reset(size_t width, size_t height)
{
}

LightMap& getCurrentLightmapData()
{
	static LightMap lightmap;
	return lightmap;
}

void loadTerrainTextures(MAP_TILESET mapTileset)
{
}

void resForceBaseDir(const char *pResDir)
{
}

// the danger thread's (never started here)
void wzSemaphoreWait(WZ_SEMAPHORE *semaphore)
{
}

void wzSemaphorePost(WZ_SEMAPHORE *semaphore)
{
}

int wzThreadJoin(WZ_THREAD *thread)
{
	return 0;
}

// and the path threads' (never started either: fpathRoute runs each job when it is queued; see run_tests.py)
static char statsrefSyncObject;
WZ_MUTEX *wzMutexCreate()
{
	return reinterpret_cast<WZ_MUTEX *>(&statsrefSyncObject);
}

void wzMutexDestroy(WZ_MUTEX *mutex)
{
}

void wzMutexLock(WZ_MUTEX *mutex)
{
}

void wzMutexUnlock(WZ_MUTEX *mutex)
{
}

WZ_SEMAPHORE *wzSemaphoreCreate(int startValue)
{
	return reinterpret_cast<WZ_SEMAPHORE *>(&statsrefSyncObject);
}

WZ_THREAD *wzThreadCreate(int (*threadFunc)(void *), void *data, const char *name)
{
	return reinterpret_cast<WZ_THREAD *>(&statsrefSyncObject);
}

void wzThreadStart(WZ_THREAD *thread)
{
}

void wzSemaphoreDestroy(WZ_SEMAPHORE *semaphore)
{
}

// cluster B's callees in later clusters (see the header)
MISSION mission;                                          // mission.cpp: its object lists
PerPlayerDroidLists apsLimboDroids;
bool godMode = false;                                     // display3d.cpp
MULTIPLAYERGAME game;                                     // multiplay.cpp
MULTIPLAYERINGAME ingame;                                 // multiplay.cpp
bool bMultiPlayer = false;
bool bMultiMessages = false;
void addEffect(const Vector3i *pos, EFFECT_GROUP group, EFFECT_TYPE type, bool specified, const iIMDShape *imd, int lit, unsigned effectTime, Vector3i *rot, Vector3f *velocity)
{
}
void addEffect(const Vector3i *pos, EFFECT_GROUP group, EFFECT_TYPE type, bool specified, const iIMDShape *imd, int lit)
{
}
void initPerimeterSmoke(const iIMDShape *pImd, Vector3i base)
{
}
void shakeStart(unsigned int length)
{
}
void scoreUpdateVar(DATA_INDEX var)
{
}
void intRefreshScreen()
{
}
void intCommanderGroupChanged(const DROID *psCommander)
{
}
void jsDebugMessageUpdate()
{
}
void scriptRemoveObject(const BASE_OBJECT *psObj)
{
}
bool triggerEventDestroyed(BASE_OBJECT *psVictim)
{
	return true;
}
bool audio_PlayStaticTrack(SDWORD iX, SDWORD iY, int iTrack)
{
	return true;
}
void audio_RemoveObj(SIMPLE_OBJECT const *psObj)
{
}
// movement (B6): what move.cpp and its helpers call in modules not ported yet
bool triggerEventPickup(FEATURE *psFeat, DROID *psDroid)
{
	return true;
}
bool triggerEventDroidMoved(DROID *psDroid, int oldx, int oldy)
{
	return true;
}
bool triggerEventDroidRankGained(const DROID *psDroid, int rankNum)
{
	return true;
}
bool getCamTweakOption_FastExp()
{
	return false;
}
uint32_t getCampaignNumber()
{
	return 1;     // camNumber's initial value
}
char ConsoleString[MAX_CONSOLE_TMP_STRING_LENGTH];
bool addConsoleMessage(const char *Text, CONSOLE_TEXT_JUSTIFICATION jusType, SDWORD player, bool team, UDWORD duration)
{
	return true;
}
bool audio_PlayObjDynamicTrack(SIMPLE_OBJECT *psObj, int iTrack, AUDIO_CALLBACK pUserCallback)
{
	return false;
}
void compPersonToBits(DROID *psDroid, Vector3f &velocity)
{
}
void destroyFXDroid(DROID *psDroid, unsigned impactTime, Vector3f &velocity)
{
}
bool tryingToGetLocation()
{
	return false;
}
void kill3DBuilding()
{
}
void intGroupsChanged(optional<UBYTE> selectedGroup)
{
}
void intInformInterfaceObjectRemoved(const BASE_OBJECT *psObj)
{
}
bool bInTutorial = false;
void markTileDirty(int i, int j)
{
}
int intGetResearchState()
{
	return 0;
}
void intNotifyResearchButton(int prevState)
{
}
bool triggerEventStructureUpgradeStarted(STRUCTURE *psStruct)
{
	return true;
}
bool wzapi::scriptIsObjectQueuedForRemoval(const BASE_OBJECT *psObj)
{
	return false;
}
void keybindInformResourceExtractorRemoved(const STRUCTURE *psResExtractor)
{
}
void cancelDeliveryRepos()
{
}
void audio_PlayTrack(int iTrack, const bool playIfPaused)
{
}
bool statsref_debugMappingsAllowed()
{
	return false;
}
bool ctrlShiftDown()
{
	return false;
}
bool anyBlueprintTooClose(STRUCTURE_STATS const *stats, Vector2i pos, uint16_t dir)
{
	return false;
}
bool crashHandlingProviderCaptureException(const char *type, const char *value, const std::string &description, bool captureStackTrace, bool handled, const nlohmann::json *additionalData)
{
	return false;
}
bool missionCanReEnforce()
{
	return false;
}
bool intAddTransporterLaunch(DROID *psDroid)
{
	return true;
}
char fileLoadBuffer[FILE_LOAD_BUFFER_SIZE];               // init.cpp
// actions and orders (B7): what action.cpp, order.cpp and their helpers call in modules not ported yet
// combat (B8): what combat.cpp, projectile.cpp, ai.cpp, visibility.cpp and the damage functions call in modules
// not ported yet
bool triggerEventSeen(BASE_OBJECT *psViewer, BASE_OBJECT *psSeen)
{
	return true;
}
const char *getPlayerName(uint32_t player, bool treatAsNonHost)
{
	return "";
}
bool triggerEventAttacked(BASE_OBJECT *psVictim, BASE_OBJECT *psAttacker, int lastHit)
{
	return true;
}
void intGroupDamaged(UBYTE group, uint64_t additionalDamage, bool unitKilled)
{
}
void intCommanderGroupDamaged(const DROID *psCommander, uint64_t additionalDamage, bool unitKilled)
{
}
void audio_QueueTrackMinDelayPos(SDWORD iTrack, UDWORD iMinDelay, SDWORD iX, SDWORD iY, SDWORD iZ)
{
}
bool DrawnInLastFrame(int32_t frame)
{
	return false;
}
bool clipXY(SDWORD x, SDWORD y)
{
	return false;
}
void addDroidDeathAnimationEffect(DROID *psDroid)
{
}
void effectGiveAuxVarSec(UDWORD var)
{
}
void console(const char *pFormat, ...)
{
}
bool getCamTweakOption_PS1Modifiers()
{
	return false;
}
void addMultiEffect(const Vector3i *basePos, Vector3i *scatter, EFFECT_GROUP group, EFFECT_TYPE type, bool specified, const iIMDShape *imd, unsigned int number, bool lit, unsigned int size, unsigned effectTime)
{
}
bool audio_PlayObjStaticTrack(SIMPLE_OBJECT *psObj, int iTrack)
{
	return false;
}
bool updateTransporter(DROID *psTransporter)
{
	return false;
}
bool transporterFlying(const DROID *psTransporter)
{
	return false;
}
bool checkTransporterSpace(DROID const *psTransporter, DROID const *psAssigned, bool mayFlash)
{
	return false;
}
void transporterAddDroid(DROID *psTransporter, DROID *psDroidToAdd)
{
}
void unloadTransporter(DROID *psTransporter, UDWORD x, UDWORD y)
{
}
void resetTransporter()
{
}
void intRemoveTrans(bool skipIntModeReset)
{
}
bool getDroidsToSafetyFlag()
{
	return false;
}
UDWORD missionGetReinforcementTime()
{
	return 0;
}
void missionGetTransporterExit(SDWORD iPlayer, UDWORD *iX, UDWORD *iY)
{
}
void missionMoveTransporterOffWorld(DROID *psTransporter)
{
}
void missionSetReinforcementTime(UDWORD iTime)
{
}
void moveDroidsToSafety(DROID *psTransporter)
{
}
bool missionDroidsRemaining(UDWORD player)
{
	return false;
}
DROID *getTrackingDroid()
{
	return nullptr;
}
bool getWarCamStatus()
{
	return false;
}
void camToggleStatus()
{
}
bool intBuildSelectMode()
{
	return false;
}
bool intDemolishSelectMode()
{
	return false;
}
std::shared_ptr<W_SCREEN> psWScreen;                      // hci.cpp
WIDGET *widgGetFromID(const std::shared_ptr<W_SCREEN> &psScreen, UDWORD id)
{
	return nullptr;
}
bool keyDown(KEY_CODE code)
{
	return false;                                         // no keys held in a hidden run (specialOrderKeyDown)
}
void assignSensorTarget(BASE_OBJECT *psObj)
{
}
void setSensorAssigned()
{
}
// (B9) the game state's update and the structures' update
// (B10) netplay.cpp without connections (NetPlay.bComms is false): NETsend, NETflush and NETrecvNet return at once
bool NETsend(NETQUEUE queue, NetMessage const& message)
{
	return true;
}
void NETsendProcessDelayedActions()
{
}
void NETflush()
{
}
bool NETrecvNet(NETQUEUE *queue, uint8_t *type)
{
	return false;
}
// netlog.cpp: the packet counters are not output
void NETlogPacket(uint8_t type, uint32_t size, bool received)
{
}
// crc.cpp's SessionKeys (libsodium): no session keys exist, so these are never called
std::vector<uint8_t> SessionKeys::encryptMessageForOther(void const *data, size_t dataLen)
{
	return {};
}
bool SessionKeys::decryptMessageFromOther(void const *data, size_t dataLen, std::vector<uint8_t> &outputDecrypted)
{
	return false;
}
// loop.cpp's gameLoop: the render loop (the display clusters) and what it calls
GAMECODE statsref_renderLoop()
{
	return GAMECODE_CONTINUE;
}
// piemode.cpp's frame flag, as the port runs it (the render pass, the screenshot and the perf counters:
// nothing to do without a context)
static bool renderingFrame = true; // starts off true

void pie_ScreenFrameRenderBegin()
{
	if (renderingFrame)
	{
#line 101 "lib/ivis_opengl/piemode.cpp"
		debug(LOG_WZ, "Call to pie_ScreenFrameRenderBegin when previous frame render hasn't ended yet");
#line 697 "tools/statsref/statsref.cpp"
		return;
	}
	renderingFrame = true;
}

void pie_ScreenFrameRenderEnd()
{
	if (!renderingFrame)
	{
#line 116 "lib/ivis_opengl/piemode.cpp"
		debug(LOG_WZ, "Call to pie_ScreenFrameRenderEnd without matching pie_ScreenFrameRenderBegin");
#line 709 "tools/statsref/statsref.cpp"
		return;
	}
	renderingFrame = false;
}
void stdOutGameSummary(UDWORD realTimeThrottleSeconds, bool flush_output)
{
}
// multiplay.cpp's recvMessage: the messages of the interface, the lobby and later clusters (as stubs_game.bas: popped
// undecoded)
bool statsref_recvBeacon(NETQUEUE queue)
{
	return true;
}
bool statsref_recvDataCheck2(NETQUEUE queue)
{
	return true;
}
bool receiveInGameTextMessage(NETQUEUE queue)
{
	return true;
}
bool recvDataCheck(NETQUEUE queue)
{
	return true;
}
bool recvTextMessageAI(NETQUEUE queue)
{
	return true;
}
bool recvSpecInGameTextMessage(NETQUEUE queue)
{
	return true;
}
bool recvStrategyPlanUpdate(NETQUEUE queue)
{
	return true;
}
bool recvQuickChat(NETQUEUE queue)
{
	return true;
}
void recvProcessDebugMappings(NETQUEUE queue)
{
}
bool recvPing(NETQUEUE queue)
{
	return true;
}
bool recvVote(NETQUEUE queue, bool inLobby)
{
	return true;
}
bool recvVoteRequest(NETQUEUE queue)
{
	return true;
}
bool recvHostConfig(NETQUEUE queue)
{
	return true;
}
bool recvMultiStats(NETQUEUE queue)
{
	return true;
}
void recvPlayerLeft(NETQUEUE queue)
{
}
bool NETlogEntry(const char *str, UDWORD a, UDWORD b)
{
	return false;                                         // no net log file (NETstartLogging is not called)
}
void NETplayerKicked(UDWORD index, bool quiet)
{
}
bool NETplayerHasConnection(uint32_t index)
{
	return false;
}
void kickPlayer(uint32_t player_id, const char *reason, LOBBY_ERROR_TYPES type, bool banPlayer)
{
}
void displayKickReasonPopup(const std::string &reason)
{
}
void setPlayerHasLost(bool val)
{
}
void sendInGameSystemMessage(const char *text)
{
}
void intAlliedResearchChanged()
{
}
bool triggerEventSyncRequest(int from, int req_id, int x, int y, BASE_OBJECT *psObj, BASE_OBJECT *psObj2)
{
	return true;
}
void transporterRemoveDroid(DROID *psTransport, DROID *psDroid, QUEUE_MODE mode)
{
}
bool gameInitialised = false;                             // main.cpp
char MultiCustomMapsPath[PATH_MAX];                       // main.cpp (multiint.cpp's decideWRF)
UBYTE bDisplayMultiJoiningStatus;                         // multiplay.cpp
// (B11) what the replay options call in later clusters (as stubs_game.bas has them): crc.cpp's EcKey without libsodium
// (as multistat_c's: a public key is its bytes, libsodium's point check skipped); the campaign's name and tweak options;
// the levels, the map list, the search path and the palette (WZGameReplayOptionsHandler::restoreOptions, which the tests
// do not call); the data hash; component.cpp's player colours; netplay.cpp's host chat default (only read with a network)
OpaqueKeyData::~OpaqueKeyData()
{
}
struct StatsrefKeyData : public OpaqueKeyData
{
	EcKey::Key publicKey;
};
EcKey::EcKey()
{
}
bool EcKey::empty() const
{
	return !vKey;
}
void EcKey::clear()
{
	vKey.reset();
}
EcKey::Key EcKey::toBytes(Privacy privacy) const
{
	if (empty() || privacy != Public)
	{
		return Key();
	}
	return static_cast<StatsrefKeyData *>(vKey.get())->publicKey;
}
bool EcKey::fromBytes(EcKey::Key const &key, Privacy privacy)
{
	clear();
	if (privacy != Public || key.size() != 32)    // crypto_sign_ed25519_PUBLICKEYBYTES
	{
		return false;
	}
	auto data = std::make_shared<StatsrefKeyData>();
	data->publicKey = key;
	vKey = data;
	return true;
}
static optional<std::string> statsrefCampaignName;
void setCampaignName(const std::string& campaignName)
{
	statsrefCampaignName = campaignName;
}
void clearCampaignName()
{
	statsrefCampaignName.reset();
}
std::string getCampaignName()
{
	return statsrefCampaignName.value_or("");
}
static std::unordered_map<std::string, nlohmann::json> statsrefCamTweakOptions;
void setCamTweakOptions(std::unordered_map<std::string, nlohmann::json> options)
{
	statsrefCamTweakOptions = options;
}
void clearCamTweakOptions()
{
	statsrefCamTweakOptions.clear();
}
const std::unordered_map<std::string, nlohmann::json>& getCamTweakOptions()
{
	return statsrefCamTweakOptions;
}
LEVEL_DATASET *levFindDataSet(char const *name, Sha256 const *hash)
{
	return nullptr;
}
bool levInitialise()
{
	return true;
}
void levShutDown()
{
}
bool rebuildSearchPath(searchPathMode mode, bool force)
{
	return true;
}
bool buildMapList(bool campaignOnly)
{
	return true;
}
bool CheckForMod(char const *mapFile)
{
	return false;
}
bool CheckForRandom(char const *mapFile, char const *mapDataFile0)
{
	return false;
}
bool setSpecialInMemoryMap(std::vector<uint8_t>&& mapArchiveData)
{
	return false;
}
Sha256 findHashOfFile(char const *realFileName)
{
	return Sha256();
}
void pal_Init()
{
}
uint32_t DataHash[DATA_MAXDATA] = {0};                    // data.cpp (no hash: the stats load outside a multiplayer game)
void resetDataHash()
{
	for (unsigned i = 0; i < DATA_MAXDATA; ++i)
	{
		DataHash[i] = 0;
	}
}
char aLevelName[MAX_LEVEL_NAME_SIZE + 1];                 // frontend.cpp
bool setPlayerColour(UDWORD player, UDWORD col)
{
	return true;
}
bool NETgetDefaultMPHostFreeChatPreference()
{
	return true;
}
// activity.cpp: the game mode WZGameReplayOptionsHandler::desiredBufferSize asks for (a skirmish)
ActivitySink::GameMode statsref_currentGameMode()
{
	return ActivitySink::GameMode::SKIRMISH;
}
// main.cpp, clparse.cpp: not a headless autogame (gameLoop's stdOutGameSummary)
bool headlessGameMode()
{
	return false;
}
bool autogame_enabled()
{
	return false;
}
bool updateScripts()
{
	return true;
}
wzapi::QueuedObjectRemovalsVector &wzapi::scriptQueuedObjectRemovals()
{
	static wzapi::QueuedObjectRemovalsVector v;
	return v;
}
void wzapi::processScriptQueuedObjectRemovals()
{
}
void statsref_logGameFrame()
{
}
const char *version_getVersionString()
{
	return "statsref";
}
// qtscript.cpp's trigger events: their stats part (the scripts are not run here)
bool triggerEventResearched(RESEARCH *psResearch, STRUCTURE *psStruct, int player)
{
	updateMultiStatsResearchComplete(psResearch, player);
	return true;
}
bool triggerEventStructureReady(STRUCTURE *psStruct)
{
	return true;
}
bool triggerEventDroidBuilt(DROID *psDroid, STRUCTURE *psFactory)
{
	if (psFactory != nullptr) // Ignore droids that were gifted or cheated (debug add)
	{
		updateMultiStatsBuilt(psDroid);
	}
	return true;
}
void intResearchFinished(STRUCTURE *psBuilding)
{
}
void effectSetSize(UDWORD size)
{
}
const iIMDShape *getDisplayImdFromIndex(UDWORD index)
{
	return nullptr;
}
bool triggerEvent(SCRIPT_TRIGGER_TYPE trigger, BASE_OBJECT *psObj)
{
	return true;
}
bool triggerEventDroidIdle(DROID *psDroid)
{
	return true;
}
bool triggerEventStructBuilt(STRUCTURE *psStruct, DROID *psDroid)
{
	if (psDroid != nullptr) // Ignore structures that were gifted or cheated (debug add)
	{
		updateMultiStatsBuilt(psStruct);
	}
	return true;
}
bool triggerEventStructDemolish(STRUCTURE *psStruct, DROID *psDroid)
{
	return true;
}
bool triggerEventSelected()
{
	return true;
}
void jsDebugSelected(const BASE_OBJECT *psObj)
{
}
void audio_QueueTrackMinDelay(SDWORD iTrack, UDWORD iMinDelay)
{
}
void audio_QueueTrack(SDWORD iTrack)
{
}
void audio_PlayBuildFailedOnce()
{
}
bool audio_PlayObjStaticTrackCallback(SIMPLE_OBJECT *psObj, int iTrack, AUDIO_CALLBACK pUserCallback)
{
	return false;
}
void audio_StopObjTrack(SIMPLE_OBJECT *psObj, int iTrack)
{
}
void audio_QueueTrackPos(SDWORD iTrack, SDWORD iX, SDWORD iY, SDWORD iZ)
{
}
void effectGiveAuxVar(UDWORD var)
{
}
char DROIDDOING[512];                                     // display3d.cpp

// imdload.cpp's GPU buffers: what is uploaded is kept
struct RecordingBuffer : public gfx_api::buffer
{
	std::vector<uint8_t> data;
	void upload(const size_t &size, const void *src) override
	{
		data.assign(static_cast<const uint8_t *>(src), static_cast<const uint8_t *>(src) + size);
	}
	void update(const size_t &start, const size_t &size, const void *src, const update_flag flag) override
	{
		memcpy(data.data() + start, src, size);
	}
	size_t current_buffer_size() override
	{
		return data.size();
	}
	void bind() override
	{
	}
};

gfx_api::buffer *statsref_create_buffer_object(const gfx_api::buffer::usage &, const gfx_api::context::buffer_storage_hint &, const std::string &)
{
	return new RecordingBuffer();
}

static std::string modelStr(const iIMDBaseShape *m)
{
	return m ? m->filename.toStdString() : "null";
}

// lib/sdl's wzapp functions debug.cpp calls (headless: no window, dialogs logged)
bool wzIsFullscreen()
{
	return false;
}

bool wzChangeWindowMode(WINDOW_MODE mode, bool silent)
{
	return false;
}

void wzDisplayDialog(DialogType type, const char *title, const char *message)
{
	static bool processingDialog = false;
	if (!processingDialog)
	{
		// in headless mode, do not display a messagebox (which might block the main thread)
		// but just log the details
		processingDialog = true;
#line 590 "lib/sdl/main_sdl.cpp"
		debug(LOG_INFO, "Suppressed dialog (headless):\n\tTitle: %s\n\tMessage: %s", title, message);
#line 231 "tools/statsref/statsref.cpp"
		processingDialog = false;
	}
}

// ---------------------------------------------------------------------------
// strres_lexer.lpp and strres_parser.ypp, written out again (flex's and bison's rules)

static PHYSFS_file *strresInput = nullptr;
static std::string strresBuf;
static size_t strresPos = 0;
static bool strresInit = false;
static int strresLine = 1;
static std::string strresText;
static std::string strresSval;

enum { SR_TEXT_T = 258, SR_QTEXT_T = 259 };

void strres_set_extra(YY_EXTRA_TYPE user_defined)
{
	strresInput = user_defined->input.physfsfile;
}

int strres_get_lineno()
{
	return strresInit ? strresLine : 1;
}

char *strres_get_text()
{
	return const_cast<char *>(strresText.c_str());
}

int strres_lex_destroy()
{
	strresBuf.clear();
	strresPos = 0;
	strresInit = false;
	strresText.clear();
	return 0;
}

static void strresMatch(size_t n)
{
	strresText = strresBuf.substr(strresPos, n);
	for (size_t i = 0; i < n; ++i)
	{
		if (strresBuf[strresPos + i] == '\n')
		{
			strresLine++;
		}
	}
	strresPos += n;
}

static int strres_lex()
{
	enum { INITIAL, COMMENT, QUOTE, SLCOMMENT };
	static int state = INITIAL;
	if (!strresInit)
	{
		strresBuf.clear();
		PHYSFS_sint64 len = PHYSFS_fileLength(strresInput);
		if (len > 0)
		{
			strresBuf.resize(static_cast<size_t>(len));
			PHYSFS_sint64 got = PHYSFS_readBytes(strresInput, &strresBuf[0], len);
			strresBuf.resize(got > 0 ? static_cast<size_t>(got) : 0);
		}
		strresPos = 0;
		strresLine = 1;
		state = INITIAL;
		strresInit = true;
	}
	for (;;)
	{
		if (strresPos >= strresBuf.size())
		{
			strresText.clear();
			return 0;
		}
		unsigned char c = strresBuf[strresPos];
		auto peek = [&](size_t k) -> int { return strresPos + k < strresBuf.size() ? (unsigned char)strresBuf[strresPos + k] : -1; };
		switch (state)
		{
		case INITIAL:
			if (isalpha(c) && c < 128)
			{
				size_t n = 1;
				while (strresPos + n < strresBuf.size())
				{
					unsigned char d = strresBuf[strresPos + n];
					if (!((d < 128 && isalnum(d)) || d == '-' || d == '_'))
					{
						break;
					}
					n++;
				}
				strresMatch(n);
				strresSval = strresText;
				return SR_TEXT_T;
			}
			else if (c == '"')
			{
				strresMatch(1);
				state = QUOTE;
			}
			else if (c == ' ' || c == '\t' || c == '\n' || c == '\r')
			{
				strresMatch(1);
			}
			else if (c == '/' && peek(1) == '*')
			{
				strresMatch(2);
				state = COMMENT;
			}
			else if (c == '/' && peek(1) == '/')
			{
				strresMatch(2);
				state = SLCOMMENT;
			}
			else
			{
				strresMatch(1);
				return (signed char)c;
			}
			break;
		case QUOTE:
			if (c == '"')
			{
				strresMatch(1);
				state = INITIAL;
			}
			else if (c == '\n')
			{
				strresMatch(1);
				strres_error("Unexpected end of line in string");
			}
			else
			{
				size_t n = 1;
				while (strresPos + n < strresBuf.size() && strresBuf[strresPos + n] != '"' && strresBuf[strresPos + n] != '\n')
				{
					n++;
				}
				strresMatch(n);
				strresSval = strresText;
				return SR_QTEXT_T;
			}
			break;
		case COMMENT:
			if (c == '*' && peek(1) == '/')
			{
				strresMatch(peek(2) == '\n' ? 3 : 2);
				state = INITIAL;
			}
			else
			{
				strresMatch(1);
			}
			break;
		case SLCOMMENT:
			if (c == '\n')
			{
				strresMatch(1);
				state = INITIAL;
			}
			else
			{
				size_t n = 1;
				while (strresPos + n < strresBuf.size() && strresBuf[strresPos + n] != '\n')
				{
					n++;
				}
				strresMatch(n);
			}
			break;
		}
	}
}

void strres_error(const char *msg)
{
#line 37 "lib/framework/strres_parser.ypp"
	debug(LOG_ERROR, "STRRES file parse error:\n%s at line %d\nText: '%s'", msg, strres_get_lineno(), strres_get_text());
#line 416 "tools/statsref/statsref.cpp"
}

int strres_parse(void *psStrRes)
{
	int lines = 0;
	for (;;)
	{
		int tok = strres_lex();
		if (tok <= 0)
		{
			if (lines > 0)
			{
				return 0;
			}
			strres_error("syntax error");
			return 1;
		}
		if (tok != SR_TEXT_T)
		{
			strres_error("syntax error");
			return 1;
		}
		std::string s1 = strresSval, s2;
		tok = strres_lex();
		if (tok == SR_QTEXT_T)
		{
			s2 = strresSval;
		}
		else if (tok == '_')
		{
			if (strres_lex() != '(' || strres_lex() != SR_QTEXT_T)
			{
				strres_error("syntax error");
				return 1;
			}
			s2 = strresSval;
			if (strres_lex() != ')')
			{
				strres_error("syntax error");
				return 1;
			}
		}
		else
		{
			strres_error("syntax error");
			return 1;
		}
		if (!strresStoreString((struct STR_RES *)psStrRes, s1.c_str(), s2.c_str()))
		{
			return 1;
		}
		lines++;
	}
}

// ---------------------------------------------------------------------------
// the log capture and the dump

static std::string captured;

// as debug.cpp's debug_callback_file writes a line
static void captureCallback(void **data, const char *outputBuffer, code_part part)
{
	captured += outputBuffer;
	if (!strchr(outputBuffer, '\n'))
	{
		captured += "\n";
	}
}

static std::string out;

static void put(const char *fmt, ...)
{
	va_list ap, ap2;
	va_start(ap, fmt);
	va_copy(ap2, ap);
	int n = vsnprintf(nullptr, 0, fmt, ap);
	va_end(ap);
	std::string buf(n > 0 ? n : 0, '\0');
	if (n > 0)
	{
		vsnprintf(&buf[0], n + 1, fmt, ap2);
	}
	va_end(ap2);
	out += buf;
}

static std::string statRef(const BASE_STATS *s)
{
	if (!s)
	{
		return "null";
	}
	char buf[512];
	snprintf(buf, sizeof(buf), "%s/%x", s->id.toUtf8().c_str(), s->ref);
	return buf;
}

template <typename T>
static std::string indexOf(const T *p, const std::vector<T> &v)
{
	if (!p)
	{
		return "null";
	}
	return std::to_string(p - v.data());
}

static void putCommon(const char *kind, size_t i, const COMPONENT_STATS &s)
{
	put("%s %u|%s|%s|ref=%x|index=%u|imd=%s|power=%u|points=%u|weight=%u|comp=%d|dtype=%d|design=%d|usage=%d\n",
	    kind, (unsigned)i, s.id.toUtf8().c_str(), s.name.toUtf8().c_str(), s.ref, (unsigned)s.index, modelStr(s.pIMD).c_str(),
	    s.buildPower, s.buildPoints, s.weight, (int)s.compType, (int)s.droidTypeOverride, (int)s.designable, (int)s.usageClass);
}

template <typename T, typename F>
static void putUpgrades(const T &s, F line)
{
	put(" base%s\n", line(s.base).c_str());
	for (int j = 0; j < MAX_PLAYERS; j++)
	{
		std::string u = line(s.upgrade[j]);
		if (u != line(s.base))
		{
			put(" upgrade %d%s\n", j, u.c_str());
		}
	}
}

static std::string fmt(const char *f, ...)
{
	char buf[4096];
	va_list ap;
	va_start(ap, f);
	vsnprintf(buf, sizeof(buf), f, ap);
	va_end(ap);
	return buf;
}

static void putList(const char *label, const std::vector<iIMDBaseShape *> &v)
{
	put(" %s %u", label, (unsigned)v.size());
	for (size_t k = 0; k < v.size(); k++)
	{
		if (v[k])
		{
			put(" %u:%s", (unsigned)k, modelStr(v[k]).c_str());
		}
	}
	put("\n");
}

static std::string hexf(float f)
{
	uint32_t bits;
	memcpy(&bits, &f, 4);
	return fmt("%08X", bits);
}

static std::string v3i(const Vector3i &v)
{
	return fmt("%d %d %d", v.x, v.y, v.z);
}

static std::string v3f(const Vector3f &v)
{
	return hexf(v.x) + " " + hexf(v.y) + " " + hexf(v.z);
}

static void dumpConnectors(const std::vector<Vector3i> &c)
{
	put(" connectors %u:", (unsigned)c.size());
	for (const Vector3i &v : c)
	{
		put(" %s;", v3i(v).c_str());
	}
	put("\n");
}

static void dumpPoints(const char *label, const std::vector<Vector3f> &p)
{
	put(" %s %u:", label, (unsigned)p.size());
	for (const Vector3f &v : p)
	{
		put(" %s;", v3f(v).c_str());
	}
	put("\n");
}

static void dumpPolys(const char *label, const std::vector<iIMDPoly> &p)
{
	put(" %s %u\n", label, (unsigned)p.size());
	for (size_t i = 0; i < p.size(); i++)
	{
		const iIMDPoly &q = p[i];
		put("  %u: flags=%X p=%u %u %u tc=%s %s %s %s %s %s anim=%s %s n=%s\n", (unsigned)i, q.flags, q.pindex[0], q.pindex[1], q.pindex[2],
		    hexf(q.texCoord[0].x).c_str(), hexf(q.texCoord[0].y).c_str(), hexf(q.texCoord[1].x).c_str(), hexf(q.texCoord[1].y).c_str(),
		    hexf(q.texCoord[2].x).c_str(), hexf(q.texCoord[2].y).c_str(), hexf(q.texAnim.x).c_str(), hexf(q.texAnim.y).c_str(), v3f(q.normal).c_str());
	}
}

static const RecordingBuffer *recorded(const iIMDShape &s, int vbo)
{
	return static_cast<const RecordingBuffer *>(s.buffers[vbo]);
}

static void dumpFloats(const char *label, const RecordingBuffer *b)
{
	size_t n = b ? b->data.size() / sizeof(float) : 0;
	put(" %s %u:", label, (unsigned)n);
	for (size_t i = 0; i < n; i++)
	{
		float f;
		memcpy(&f, b->data.data() + i * sizeof(float), sizeof(float));
		put(" %s", hexf(f).c_str());
	}
	put("\n");
}

static void modelDump(const iIMDBaseShape *b)
{
	put("base filename=%s path=%s min=%s max=%s sradius=%d radius=%d ocen=%s\n", b->filename.toUtf8().c_str(), b->path.toUtf8().c_str(),
	    v3i(b->min).c_str(), v3i(b->max).c_str(), b->sradius, b->radius, v3f(b->ocen).c_str());
	dumpConnectors(b->connectors);
	int k = 0;
	for (const iIMDShape *s = b->displayModel(); s != nullptr; s = s->next.get(), k++)
	{
		put("level %d name=%s modelLevel=%u flags=%X numFrames=%u animInterval=%u interpolate=%d shadowUsesAlt=%d vertexCount=%u\n", k,
		    s->modelName.toUtf8().c_str(), s->modelLevel, s->flags, (unsigned)s->numFrames, (unsigned)s->animInterval, s->interpolate,
		    (int)(s->pShadowPoints == &s->altShadowPoints), (unsigned)s->vertexCount);
		put(" bounds min=%s max=%s sradius=%d radius=%d ocen=%s\n", v3i(s->min).c_str(), v3i(s->max).c_str(), s->sradius, s->radius, v3f(s->ocen).c_str());
		dumpConnectors(s->connectors);
		put(" textures");
		for (const TilesetTextureFiles &t : s->tilesetTextureFiles)
		{
			put(" %s|%s|%s|%s", t.texfile.c_str(), t.tcmaskfile.c_str(), t.normalfile.c_str(), t.specfile.c_str());
		}
		put("\n");
		put(" objanim time=%d cycles=%d frames=%d pie=%s %s %s %s\n", s->objanimtime, s->objanimcycles, s->objanimframes,
		    modelStr(s->objanimpie[0]).c_str(), modelStr(s->objanimpie[1]).c_str(), modelStr(s->objanimpie[2]).c_str(), modelStr(s->objanimpie[3]).c_str());
		for (int i = 0; i < s->objanimframes; i++)
		{
			const ANIMFRAME &f = s->objanimdata[i];
			put("  %d: scale=%s pos=%s rot=%u %u %u\n", i, v3f(f.scale).c_str(), v3f(f.pos).c_str(), (unsigned)f.rot.direction, (unsigned)f.rot.pitch, (unsigned)f.rot.roll);
		}
		dumpPoints("points", s->points);
		dumpPolys("polys", s->polys);
		dumpPoints("altpoints", s->altShadowPoints);
		dumpPolys("altpolys", s->altShadowPolys);
		dumpFloats("vertices", recorded(*s, VBO_VERTEX));
		dumpFloats("normals", recorded(*s, VBO_NORMAL));
		dumpFloats("texcoords", recorded(*s, VBO_TEXCOORD));
		dumpFloats("tangents", recorded(*s, VBO_TANGENT));
		const RecordingBuffer *ib = recorded(*s, VBO_INDEX);
		size_t n = ib ? ib->data.size() / sizeof(uint16_t) : 0;
		put(" indices %u:", (unsigned)n);
		for (size_t i = 0; i < n; i++)
		{
			uint16_t v;
			memcpy(&v, ib->data.data() + i * sizeof(uint16_t), sizeof(uint16_t));
			put(" %u", (unsigned)v);
		}
		put("\n");
	}
}

// loadGame's map part (src/game/vbtest_vb.bas's vbtestMapLoad)
static std::shared_ptr<WzMap::Map> testMap;              // the last map mapload loaded (its objects for mapfeatures)
static bool syncDumpAll = false;                         // syncdumpall: mapmove, maporders, mapcombat dump every tick's sync log
bool loadWzMapFeature(WzMap::Map &wzMap, std::unordered_map<UDWORD, UDWORD> &fixedMapIdToGeneratedId);  // game.cpp
bool loadWzMapDroidInit(WzMap::Map &wzMap, std::unordered_map<UDWORD, UDWORD> &fixedMapIdToGeneratedId);  // game.cpp
bool loadWzMapStructure(WzMap::Map &wzMap, std::unordered_map<UDWORD, UDWORD> &fixedMapIdToGeneratedId,
                        std::array<std::unordered_map<UDWORD, UDWORD>, MAX_PLAYER_SLOTS> &moduleToBuilding);
extern std::unordered_set<uint16_t> derricks;             // qtscript.cpp

static bool statsrefMapLoad(const std::string &folder, const std::string &tileset, int players, const std::string &kind)
{
	if (tilesetDir)
	{
		free(tilesetDir);
		tilesetDir = nullptr;
	}
	if (tileset != "-")
	{
		tilesetDir = strdup(tileset.c_str());
	}
	auto map = WzMap::Map::loadFromPath(folder, kind == "campaign" ? WzMap::MapType::CAMPAIGN : WzMap::MapType::SKIRMISH, players, 0,
	                                    std::make_shared<WzMapDebugLogger>(), std::make_shared<WzMapPhysFSIO>());
	testMap = map;
	if (!map)
	{
		put("%s", captured.c_str());
		captured.clear();
		put("map=null\n");
		return false;
	}
	if (!loadTerrainTypeMap(map->mapTerrainTypes()))
	{
		put("%s", captured.c_str());
		captured.clear();
		put("ttypes failed\n");
		return false;
	}
	auto data = map->mapData();
	if (!data)
	{
		put("%s", captured.c_str());
		captured.clear();
		put("mapdata=null\n");
		return false;
	}
	return mapLoadFromWzMapData(data);
}

static void mapDump()
{
	put("map width=%d height=%d scroll=%d %d %d %d tileset=%s current=%d groundTypes=%u\n", mapWidth, mapHeight,
	    scrollMinX, scrollMinY, scrollMaxX, scrollMaxY, tilesetDir ? tilesetDir : "(null)", (int)currentMapTileset, (unsigned)getNumGroundTypes());
	for (size_t i = 0; i < getNumGroundTypes(); i++)
	{
		const GROUND_TYPE &g = getGroundType(i);
		put("ground %u: %s %s %s|%s|%s %d\n", (unsigned)i, g.textureName.c_str(), hexf(g.textureSize).c_str(), g.normalMapTextureName.c_str(),
		    g.specularMapTextureName.c_str(), g.heightMapTextureName.c_str(), (int)g.highQualityTextures);
	}
	std::string s = "ttypes:";
	for (int i = 0; i < MAX_TILE_TEXTURES; i++)
	{
		s += " " + std::to_string(terrainTypes[i]);
	}
	put("%s\n", s.c_str());
	put("gateways %u:", (unsigned)gwNumGateways());
	for (GATEWAY *gw : gwGetGateways())
	{
		put(" %u,%u,%u,%u", (unsigned)gw->x1, (unsigned)gw->y1, (unsigned)gw->x2, (unsigned)gw->y2);
	}
	put("\n");
	for (int y = 0; y < mapHeight; y++)
	{
		std::string row = std::to_string(y) + ":";
		for (int x = 0; x < mapWidth; x++)
		{
			const MAPTILE &t = psMapTiles[x + y * mapWidth];
			row += fmt(" %u/%d/%d/%u/%X/%u/%u/%X", (unsigned)t.texture, t.height, t.waterLevel, (unsigned)t.ground, (unsigned)t.tileInfoBits,
			           (unsigned)t.limitedContinent, (unsigned)t.hoverContinent, (unsigned)blockTile(x, y, 0));
		}
		put("%s\n", row.c_str());
	}
}

// Park and Miller's minimal standard generator (48271), as vbtest_vb's
// the droids after mapobjects, as vbtest_vb's droidsDump
static void droidsDump()
{
	for (int p = 0; p < MAX_PLAYERS; p++)
	{
		if (!apsDroidLists[p].empty())
		{
			put("droids p%d %u\n", p, (unsigned)apsDroidLists[p].size());
		}
		for (const DROID *d : apsDroidLists[p])
		{
			std::string s = fmt(" d id=%u name=%s type=%d player=%u pos=%d,%d,%d rot=%u,%u,%u body=%u orig=%u weight=%u speed=%u exp=%u kills=%u shield=%d resist=%d",
			                    d->id, d->aName, (int)d->droidType, (unsigned)d->player, d->pos.x, d->pos.y, d->pos.z, (unsigned)d->rot.direction,
			                    (unsigned)d->rot.pitch, (unsigned)d->rot.roll, d->body, d->originalBody, d->weight, d->baseSpeed, d->experience,
			                    d->kills, d->shieldPoints, (int)d->resistance);
			s += " bits=";
			for (int i = 0; i < DROID_MAXCOMP; i++)
			{
				s += fmt("%u/", (unsigned)d->asBits[i]);
			}
			s += fmt(" weaps=%u", d->numWeaps);
			for (int i = 0; i < MAX_WEAPONS; i++)
			{
				s += fmt(" w%d=%u/%u/%u", i, d->asWeaps[i].nStat, d->asWeaps[i].ammo, d->asWeaps[i].lastFired);
			}
			s += fmt(" blocked=%x imd=%s t=%u,%u born=%u died=%u watched=%u flags=%d sec=%x secp=%x action=%d order=%d group=%u illum=%u",
			         (unsigned)d->blockedBits, modelStr(d->sDisplay.imd).c_str(), d->time, d->prevSpacetime.time, d->born, d->died,
			         (unsigned)d->watchedTiles.size(), (int)d->flags.to_ulong(), d->secondaryOrder, d->secondaryOrderPending, (int)d->action,
			         (int)d->order.type, (unsigned)d->group, (unsigned)d->illumination);
			if (d->psGroup)
			{
				s += fmt(" grp=%d/%d/%d", d->psGroup->id, (int)d->psGroup->type, (int)d->psGroup->refCount);
			}
			else
			{
				s += " grp=null";
			}
			s += " vis=";
			for (int i = 0; i < MAX_PLAYERS; i++)
			{
				s += fmt("%02X", (unsigned)d->visible[i]);
			}
			put("%s\n", s.c_str());
		}
	}
}

// the structures, function lists, flags and tile visibility after mapobjects, as vbtest_vb's structuresDump
static void structuresDump()
{
	for (int p = 0; p < MAX_PLAYERS; p++)
	{
		if (!apsStructLists[p].empty())
		{
			put("structures p%d %u\n", p, (unsigned)apsStructLists[p].size());
		}
		for (const STRUCTURE *st : apsStructLists[p])
		{
			std::string s = fmt(" s id=%u stat=%s player=%u pos=%d,%d,%d rot=%u,%u,%u status=%d body=%u resist=%d cap=%u pts=%u imd=%s pre=%s fd=%s state=%d t=%u,%u,%u born=%u extra=%d flags=%d watched=%u weaps=%u",
			                    st->id, st->pStructureType->id.toUtf8().c_str(), (unsigned)st->player, st->pos.x, st->pos.y, st->pos.z,
			                    (unsigned)st->rot.direction, (unsigned)st->rot.pitch, (unsigned)st->rot.roll, (int)st->status, st->body,
			                    st->resistance, (unsigned)st->capacity, st->currentBuildPts, modelStr(st->sDisplay.imd).c_str(),
			                    modelStr(st->prebuiltImd).c_str(), hexf(st->foundationDepth).c_str(), (int)st->state, st->lastStateTime,
			                    st->prevTime, st->time, st->born, (int)st->hasExtraFunction, (int)st->flags.to_ulong(),
			                    (unsigned)st->watchedTiles.size(), st->numWeaps);
			for (int i = 0; i < MAX_WEAPONS; i++)
			{
				s += fmt(" w%d=%u/%u/%u/%d", i, st->asWeaps[i].nStat, st->asWeaps[i].ammo, st->asWeaps[i].lastFired, (int)st->asWeaps[i].origin);
			}
			s += " vis=";
			for (int i = 0; i < MAX_PLAYERS; i++)
			{
				s += fmt("%02X", (unsigned)st->visible[i]);
			}
			if (st->pFunctionality)
			{
				switch (st->pStructureType->type)
				{
				case REF_FACTORY:
				case REF_CYBORG_FACTORY:
				case REF_VTOL_FACTORY:
				{
					const FLAG_POSITION *fp = st->pFunctionality->factory.psAssemblyPoint;
					s += fmt(" factory order=%x", st->pFunctionality->factory.secondaryOrder);
					if (fp)
					{
						s += fmt(" ap=%u/%u/%d,%d,%d", (unsigned)fp->factoryInc, (unsigned)fp->factoryType, fp->coords.x, fp->coords.y, fp->coords.z);
					}
					break;
				}
				case REF_REPAIR_FACILITY:
				{
					const FLAG_POSITION *fp = st->pFunctionality->repairFacility.psDeliveryPoint;
					if (fp)
					{
						s += fmt(" dp=%u/%u/%d,%d,%d", (unsigned)fp->factoryInc, (unsigned)fp->factoryType, fp->coords.x, fp->coords.y, fp->coords.z);
					}
					break;
				}
				case REF_POWER_GEN:
					s += " gen=";
					for (int i = 0; i < NUM_POWER_MODULES; i++)
					{
						const STRUCTURE *e = st->pFunctionality->powerGenerator.apResExtractors[i];
						s += e ? fmt("%u,", e->id) : std::string("0,");
					}
					break;
				case REF_RESOURCE_EXTRACTOR:
				{
					const STRUCTURE *g = st->pFunctionality->resourceExtractor.psPowerGen;
					s += g ? fmt(" pg=%u", g->id) : std::string(" pg=0");
					break;
				}
				case REF_WALL:
				case REF_GATE:
					s += fmt(" wall=%u", st->pFunctionality->wall.type);
					break;
				default:
					break;
				}
			}
			put("%s\n", s.c_str());
		}
		std::string s = fmt("extractors p%d %u:", p, (unsigned)apsExtractorLists[p].size());
		for (const STRUCTURE *e : apsExtractorLists[p])
		{
			s += fmt(" %u", e->id);
		}
		if (!apsExtractorLists[p].empty())
		{
			put("%s\n", s.c_str());
		}
		s = fmt("flags p%d %u:", p, (unsigned)apsFlagPosLists[p].size());
		for (const FLAG_POSITION *fp : apsFlagPosLists[p])
		{
			s += fmt(" %u/%u/%d,%d,%d", (unsigned)fp->factoryInc, (unsigned)fp->factoryType, fp->coords.x, fp->coords.y, fp->coords.z);
		}
		if (!apsFlagPosLists[p].empty())
		{
			put("%s\n", s.c_str());
		}
	}
	std::string s = fmt("sensors %u:", (unsigned)apsSensorList[0].size());
	for (const BASE_OBJECT *o : apsSensorList[0])
	{
		s += fmt(" %u", o->id);
	}
	put("%s\n", s.c_str());
	for (int y = 0; y < mapHeight; y++)
	{
		s = fmt("%d vis:", y);
		for (int x = 0; x < mapWidth; x++)
		{
			s += fmt("%04X", (unsigned)mapTile(x, y)->tileExploredBits);
		}
		for (int x = 0; x < mapWidth; x++)
		{
			const MAPTILE *t = mapTile(x, y);
			for (int p = 0; p < MAX_PLAYERS; p++)
			{
				if (t->watchers[p] != 0 || t->sensors[p] != 0)
				{
					s += fmt(" %d:%d=%u/%u", x, p, (unsigned)t->watchers[p], (unsigned)t->sensors[p]);
				}
			}
			if (t->sensorBits != 0 || t->jammerBits != 0)
			{
				s += fmt(" %d:s%x/%x", x, (unsigned)t->sensorBits, (unsigned)t->jammerBits);
			}
		}
		put("%s\n", s.c_str());
		s = fmt("%d aux:", y);
		for (int x = 0; x < mapWidth; x++)
		{
			for (int p = 0; p < MAX_PLAYERS; p++)
			{
				if (auxTile(x, y, p) != 0)
				{
					s += fmt(" %d:%d=%x", x, p, (unsigned)auxTile(x, y, p));
				}
			}
		}
		put("%s\n", s.c_str());
	}
}

// the objects after mapobjects, as vbtest_vb's objectsDump
static void objectsDump()
{
	put("ids unsynch=%u synch=%u destroyed=%u\n", unsynchObjID, synchObjID, (unsigned)psDestroyedObj.size());
	put("features %u\n", (unsigned)apsFeatureList[0].size());
	for (const FEATURE *f : apsFeatureList[0])
	{
		put(" f id=%u stat=%s pos=%d,%d,%d rot=%u,%u,%u player=%u body=%u fd=%s born=%u died=%u imd=%s type=%d extra=%d\n",
		    f->id, f->psStats->id.toUtf8().c_str(), f->pos.x, f->pos.y, f->pos.z, (unsigned)f->rot.direction, (unsigned)f->rot.pitch,
		    (unsigned)f->rot.roll, (unsigned)f->player, f->body, hexf(f->foundationDepth).c_str(), f->born, f->died,
		    modelStr(f->sDisplay.imd).c_str(), (int)f->type, (int)f->hasExtraFunction);
	}
	std::string s = fmt("oil %u:", (unsigned)apsOilList[0].size());
	for (const FEATURE *f : apsOilList[0])
	{
		s += fmt(" %u", f->id);
	}
	put("%s\n", s.c_str());
	std::vector<uint16_t> d(derricks.begin(), derricks.end());
	std::sort(d.begin(), d.end());
	s = fmt("derricks %u:", (unsigned)d.size());
	for (uint16_t p : d)
	{
		s += fmt(" %u,%u", (unsigned)(p >> 8), (unsigned)(p & 255));
	}
	put("%s\n", s.c_str());
	s = "starts:";
	for (int i = 0; i < MAX_PLAYERS; i++)
	{
		Vector2i p = getPlayerStartPosition(i);
		s += fmt(" %d,%d", p.x, p.y);
	}
	put("%s\n", s.c_str());
	droidsDump();
	structuresDump();
	for (int y = 0; y < mapHeight; y++)
	{
		s = fmt("%d:", y);
		for (int x = 0; x < mapWidth; x++)
		{
			const MAPTILE *t = mapTile(x, y);
			if (t->psObject)
			{
				s += fmt(" %d=%u", x, t->psObject->id);
			}
		}
		s += " |";
		for (int x = 0; x < mapWidth; x++)
		{
			s += fmt("%02X", (unsigned)blockTile(x, y, 0));
		}
		put("%s\n", s.c_str());
	}
}

static int32_t testRand(uint64_t &st)
{
	st = st * 48271 % 2147483647;
	return (int32_t)st;
}

static int32_t testRandMod(uint64_t &st, int32_t n)
{
	return testRand(st) % n;
}

static void putCaptured()
{
	put("%s", captured.c_str());
	captured.clear();
}

// the ids of a grid query's objects
static std::string gridIds(GridList const &list)
{
	std::string s;
	for (BASE_OBJECT *psObj : list)
	{
		s += fmt(s.empty() ? "%u" : ",%u", psObj->id);
	}
	return s;
}

static void mapGrid()
{
	uint64_t st = 7 + mapWidth * 13 + mapHeight;
	int wx = world_coord(mapWidth);
	int wy = world_coord(mapHeight);
	// the objects in gridReset's order, to centre half the queries on
	std::vector<BASE_OBJECT *> all;
	for (int p = 0; p < MAX_PLAYERS; p++)
	{
		for (DROID *psDroid : apsDroidLists[p])
		{
			all.push_back(psDroid);
		}
		for (STRUCTURE *psStruct : apsStructLists[p])
		{
			all.push_back(psStruct);
		}
	}
	for (FEATURE *psFeat : apsFeatureList[0])
	{
		all.push_back(psFeat);
	}
	gridInitialise();
	gridReset();
	putCaptured();
	for (int i = 0; i < 80; i++)
	{
		if (i == 40)
		{
			gridReset();
		}
		int x, y;
		if (!all.empty() && testRandMod(st, 2) == 0)
		{
			BASE_OBJECT *psObj = all[testRandMod(st, (int)all.size())];
			x = psObj->pos.x + testRandMod(st, 2 * TILE_UNITS + 1) - TILE_UNITS;
			y = psObj->pos.y + testRandMod(st, 2 * TILE_UNITS + 1) - TILE_UNITS;
		}
		else
		{
			x = testRandMod(st, wx + 2 * TILE_UNITS) - TILE_UNITS;
			y = testRandMod(st, wy + 2 * TILE_UNITS) - TILE_UNITS;
		}
		int r = testRandMod(st, 12 * TILE_UNITS);
		int kind = testRandMod(st, 5);
		int player = testRandMod(st, MAX_PLAYERS);
		std::string s;
		switch (kind)
		{
		case 0: s = gridIds(gridStartIterate(x, y, r)); break;
		case 1: s = gridIds(gridStartIterateArea(x, y, x + r, y + r)); break;
		case 2: s = gridIds(gridStartIterateDroidsByPlayer(x, y, r, player)); break;
		case 3: s = gridIds(gridStartIterateRepairCandidates(x, y, r, player)); break;
		default:
		{
			GridList const &list = gridStartIterateUnseen(x, y, r, player);
			s = gridIds(list);
			// every other object found is seen: the player's later queries skip it
			for (size_t n = 0; n < list.size(); n += 2)
			{
				list[n]->seenThisTick[player] = UINT8_MAX;
			}
			break;
		}
		}
		putCaptured();
		put("grid %d: %d %d,%d,%d p%d = %s\n", i, kind, x, y, r, player, s.c_str());
	}
	gridShutDown();
}

static std::string powerState(int p)
{
	return fmt(" p%d=%d/%lld/%lld/%lld/%d/%d/%s", p, getPower(p), (long long)getPrecisePower(p), (long long)getExtractedPower(p),
	           (long long)getWastedPower(p), getQueuedPower(p), getPowerMinusQueued(p), getApproxPowerGeneratedPerSecForDisplay(p).c_str());
}

static void mapPower()
{
	uint64_t st = 11 + mapWidth * 3 + mapHeight * 5;
	allocPlayerPower();
	for (int p = 0; p < MAX_PLAYERS; p++)
	{
		if (p % 3 == 1)
		{
			setPowerModifier(p, 50 + 25 * (p % 5));
		}
		setPower(p, testRandMod(st, 3000));
	}
	setPowerMaxStorage(2, 2500);
	putCaptured();
	for (int t = 1; t <= 40; t++)
	{
		gameTime = t * GAME_TICKS_PER_UPDATE;
		for (int p = 0; p < MAX_PLAYERS; p++)
		{
			updatePlayerPower(p, t % 3 == 0 ? 3 : 1);
		}
		std::string s = fmt("power t%d:", t);
		for (int p = 0; p < MAX_PLAYERS; p++)
		{
			if (apsStructLists[p].empty())
			{
				continue;
			}
			// a request from one of the player's structures
			int k = testRandMod(st, (int)apsStructLists[p].size());
			STRUCTURE *psStruct = *std::next(apsStructLists[p].begin(), k);
			int amount = testRandMod(st, 900) - 100;
			bool ok;
			if (t % 4 == 0)
			{
				ok = requestPrecisePowerFor(psStruct, (int64_t)amount * 1000003);
			}
			else
			{
				ok = requestPowerFor(psStruct, amount);
			}
			s += fmt(" r%u,%d=%d/%d", psStruct->id, amount, (int)ok, checkPowerRequest(psStruct));
			if (t % 9 == 0)
			{
				delPowerRequest(psStruct);
			}
		}
		if (t % 7 == 0)
		{
			for (int p = 0; p < MAX_PLAYERS; p++)
			{
				addPower(p, testRandMod(st, 600) - 100);
				usePower(p, testRandMod(st, 400));
			}
		}
		if (t == 20)
		{
			// a derrick of each player dies: its power generator lets go of it
			for (int p = 0; p < MAX_PLAYERS; p++)
			{
				if (!apsExtractorLists[p].empty())
				{
					(*apsExtractorLists[p].begin())->died = gameTime;
				}
			}
		}
		for (int p = 0; p < MAX_PLAYERS; p++)
		{
			s += powerState(p);
		}
		putCaptured();
		put("%s\n", s.c_str());
	}
	gameTime = 0;
}

// a path's points
static std::string pathPoints(MOVE_CONTROL const &m)
{
	std::string s;
	for (Vector2i const &p : m.asPath)
	{
		s += fmt(" %d,%d", p.x, p.y);
	}
	return s;
}

static void mapPaths()
{
	uint64_t st = 3 + mapWidth * 17 + mapHeight * 3;
	int wx = world_coord(mapWidth);
	int wy = world_coord(mapHeight);
	fpathInitialise();
	putCaptured();
	for (int tick = 1; tick <= 6; tick++)
	{
		gameTime = tick * GAME_TICKS_PER_UPDATE;
		// each destination wanted from several origins: the contexts are reused within a tick
		for (int d = 0; d < 3; d++)
		{
			int tX = testRandMod(st, wx);
			int tY = testRandMod(st, wy);
			PROPULSION_TYPE prop = (PROPULSION_TYPE)testRandMod(st, PROPULSION_TYPE_NUM);
			int owner = testRandMod(st, MAX_PLAYERS);
			FPATH_MOVETYPE moveType = (FPATH_MOVETYPE)testRandMod(st, 3);
			for (int o = 0; o < 4; o++)
			{
				unsigned id = 1000 * tick + 10 * d + o;
				int sX = testRandMod(st, wx);
				int sY = testRandMod(st, wy);
				MOVE_CONTROL m;
				FPATH_RETVAL r1 = fpathRoute(&m, id, sX, sY, tX, tY, prop, DROID_WEAPON, moveType, owner, o % 2 == 0, getStructureBounds((BASE_OBJECT *)nullptr));
				m.Status = MOVEWAITROUTE;
				FPATH_RETVAL r2 = fpathRoute(&m, id, sX, sY, tX, tY, prop, DROID_WEAPON, moveType, owner, o % 2 == 0, getStructureBounds((BASE_OBJECT *)nullptr));
				putCaptured();
				put("path %u: %d,%d->%d,%d prop%d p%d m%d = %d/%d %d dest=%d,%d n=%d:%s\n", id, sX, sY, tX, tY, (int)prop, owner, (int)moveType, (int)r1, (int)r2,
				    (int)m.Status, m.destination.x, m.destination.y, (int)m.asPath.size(), pathPoints(m).c_str());
			}
		}
	}
	// the droids, to points and to structures
	gameTime = 7 * GAME_TICKS_PER_UPDATE;
	std::vector<STRUCTURE *> structs;
	for (int p = 0; p < MAX_PLAYERS; p++)
	{
		for (STRUCTURE *psStruct : apsStructLists[p])
		{
			structs.push_back(psStruct);
		}
	}
	int n = 0;
	for (int p = 0; p < MAX_PLAYERS && n < 30; p++)
	{
		for (DROID *psDroid : apsDroidLists[p])
		{
			if (n++ >= 30)
			{
				break;
			}
			int tX, tY;
			if (!structs.empty() && testRandMod(st, 2) == 0)
			{
				STRUCTURE *psStruct = structs[testRandMod(st, (int)structs.size())];
				tX = psStruct->pos.x;
				tY = psStruct->pos.y;
			}
			else
			{
				tX = testRandMod(st, wx);
				tY = testRandMod(st, wy);
			}
			FPATH_RETVAL r1 = fpathDroidRoute(psDroid, tX, tY, FMT_MOVE);
			psDroid->sMove.Status = MOVEWAITROUTE;
			FPATH_RETVAL r2 = fpathDroidRoute(psDroid, tX, tY, FMT_MOVE);
			putCaptured();
			put("droid %u: ->%d,%d = %d/%d %d dest=%d,%d n=%d:%s\n", psDroid->id, tX, tY, (int)r1, (int)r2, (int)psDroid->sMove.Status,
			    psDroid->sMove.destination.x, psDroid->sMove.destination.y, (int)psDroid->sMove.asPath.size(), pathPoints(psDroid->sMove).c_str());
		}
	}
	std::string s = "check:";
	for (int i = 0; i < 40; i++)
	{
		// (one call per statement: the order of a constructor's arguments is unspecified)
		int ax = testRandMod(st, wx + 2 * TILE_UNITS) - TILE_UNITS;
		int ay = testRandMod(st, wy + 2 * TILE_UNITS) - TILE_UNITS;
		int bx = testRandMod(st, wx);
		int by = testRandMod(st, wy);
		Position a(ax, ay, 0);
		Position b(bx, by, 0);
		PROPULSION_TYPE prop = (PROPULSION_TYPE)testRandMod(st, PROPULSION_TYPE_NUM);
		s += fmt(" %d", (int)fpathCheck(a, b, prop));
	}
	putCaptured();
	put("%s\n", s.c_str());
	fpathShutdown();
	putCaptured();
	gameTime = 0;
}

// a droid's movement state
static std::string moveDroidLine(DROID const *psDroid)
{
	MOVE_CONTROL const &m = psDroid->sMove;
	std::string s = fmt("droid %u: p%d %s pos(%d,%d,%d) rot(%d,%d,%d) speed%d moveDir%d path%d/%d src(%d,%d) target(%d,%d) dest(%d,%d)"
	                    " bump(%u,%d,%d,%d,%d,%d) shuffle%u tol%u vert%d height%d exp%u anim%d",
	                    psDroid->id, (int)psDroid->player, moveDescription(m.Status), psDroid->pos.x, psDroid->pos.y, psDroid->pos.z,
	                    (int)psDroid->rot.direction, (int)psDroid->rot.pitch, (int)psDroid->rot.roll, m.speed, (int)m.moveDir,
	                    m.pathIndex, (int)m.asPath.size(), m.src.x, m.src.y, m.target.x, m.target.y, m.destination.x, m.destination.y,
	                    m.bumpTime, (int)m.lastBump, (int)m.pauseTime, (int)m.bumpDir, m.bumpPos.x, m.bumpPos.y, m.shuffleStart,
	                    m.tolerance, m.iVertSpeed, psDroid->heightAboveMap, psDroid->experience, (int)psDroid->animationEvent);
	if (m.psFormation)
	{
		FORMATION const *f = m.psFormation;
		s += fmt(" frm(%d,%d,%d,%d,%u,%d,%d)", f->x, f->y, (int)f->direction, (int)f->refCount, f->iSpeed, (int)f->size, (int)f->rankDist);
	}
	return s + "\n";
}

static std::vector<DROID *> allDroids()
{
	std::vector<DROID *> v;
	for (int p = 0; p < MAX_PLAYERS; p++)
	{
		for (DROID *psDroid : apsDroidLists[p])
		{
			v.push_back(psDroid);
		}
	}
	return v;
}

static void mapMove()
{
	uint64_t st = 5 + mapWidth * 13 + mapHeight * 7;
	int wx = world_coord(mapWidth);
	int wy = world_coord(mapHeight);
	fpathInitialise();
	gridInitialise();
	gameTime = 1000;
	deltaGameTime = GAME_TICKS_PER_UPDATE;
	std::vector<DROID *> droids = allDroids();
	for (DROID *psDroid : droids)
	{
		moveReallyStopDroid(psDroid);
		initDroidMovement(psDroid);
	}
	gridReset();
	putCaptured();
	// the orders: three droids to a destination, near the first or anywhere
	int n = 0, tX = 0, tY = 0;
	for (DROID *psDroid : droids)
	{
		if (n >= 48)
		{
			break;
		}
		if (n % 3 == 0)
		{
			if (testRandMod(st, 4) == 0)
			{
				tX = testRandMod(st, wx);
				tY = testRandMod(st, wy);
			}
			else
			{
				int rx = testRandMod(st, 24 * TILE_UNITS + 1) - 12 * TILE_UNITS;
				int ry = testRandMod(st, 24 * TILE_UNITS + 1) - 12 * TILE_UNITS;
				tX = clip(psDroid->pos.x + rx, TILE_UNITS, wx - TILE_UNITS - 1);
				tY = clip(psDroid->pos.y + ry, TILE_UNITS, wy - TILE_UNITS - 1);
			}
		}
		bool ok = true;
		const char *how;
		if (psDroid->isVtol() && testRandMod(st, 2) == 0)
		{
			moveDroidToDirect(psDroid, tX, tY);
			how = "direct";
		}
		else if (n % 5 == 4)
		{
			ok = moveDroidToNoFormation(psDroid, tX, tY);
			how = "noformation";
		}
		else
		{
			ok = moveDroidTo(psDroid, tX, tY);
			how = "to";
		}
		putCaptured();
		put("order %u: %s ->%d,%d = %d %s\n", psDroid->id, how, tX, tY, (int)ok, moveDescription(psDroid->sMove.Status));
		n++;
	}
	for (int tick = 1; tick <= 80; tick++)
	{
		gameTime += GAME_TICKS_PER_UPDATE;
		gridReset();
		fpathUpdate();
		if (tick == 30 || tick == 50)
		{
			// stops, turns and new destinations on the way
			for (size_t i = 0; i < droids.size() && i < 48; i++)
			{
				DROID *psDroid = droids[i];
				if (psDroid->died)
				{
					continue;
				}
				if (i % 7 == 3)
				{
					moveStopDroid(psDroid);
					put("stop %u: %s\n", psDroid->id, moveDescription(psDroid->sMove.Status));
				}
				else if (i % 11 == 5)
				{
					int x = testRandMod(st, wx);
					int y = testRandMod(st, wy);
					moveTurnDroid(psDroid, x, y);
					put("turn %u: ->%d,%d %s\n", psDroid->id, x, y, moveDescription(psDroid->sMove.Status));
				}
				else if (i % 13 == 6)
				{
					moveReallyStopDroid(psDroid);
					put("reallystop %u\n", psDroid->id);
				}
				else if (i % 4 == 1)
				{
					int x = testRandMod(st, wx);
					int y = testRandMod(st, wy);
					bool ok = moveDroidTo(psDroid, x, y);
					putCaptured();
					put("reorder %u: ->%d,%d = %d %s\n", psDroid->id, x, y, (int)ok, moveDescription(psDroid->sMove.Status));
				}
			}
		}
		for (DROID *psDroid : allDroids())
		{
			if (psDroid->died)
			{
				continue;
			}
			syncDebugDroid(psDroid, 'M');
			moveUpdateDroid(psDroid);
			syncDebugDroid(psDroid, '>');
		}
		putCaptured();
		if (tick == 1 || tick == 80 || syncDumpAll)
		{
			size_t len = dumpLocalDebugSyncLog(syncDebugNext);
			out.append((const char *)debugSyncTmpBuf.data(), len);
		}
		unsigned crc = nextDebugSync();
		putCaptured();
		put("tick %d: crc=%04X\n", tick, crc);
		if (tick % 20 == 0)
		{
			for (size_t i = 0; i < droids.size() && i < 48; i++)
			{
				if (!droids[i]->died)
				{
					put("%s", moveDroidLine(droids[i]).c_str());
				}
			}
		}
	}
	put("features %d\n", (int)apsFeatureList[0].size());
	std::string s = "power:";
	for (int p = 0; p < MAX_PLAYERS; p++)
	{
		s += fmt(" %d", (int)getPower(p));
	}
	put("%s\n", s.c_str());
	gridShutDown();
	fpathShutdown();
	putCaptured();
	gameTime = 0;
	deltaGameTime = 0;
}

// a droid's order and action state
static std::string orderDroidLine(DROID const *psDroid)
{
	DROID_ORDER_DATA const &o = psDroid->order;
	BASE_OBJECT const *t = psDroid->psActionTarget[0];
	return fmt("droid %u: p%d t%d %s pos(%d,%d,%d) body%u/%u order%d(%d,%d)(%d,%d)dir%d obj%u stats%d rtr%d list%d/%d"
	           " action%d(%d,%d) tgt%u start%u pts%u sec%X base%u grp%d/%d wpn(%d,%d) ur%d\n",
	           psDroid->id, (int)psDroid->player, (int)psDroid->droidType, moveDescription(psDroid->sMove.Status),
	           psDroid->pos.x, psDroid->pos.y, psDroid->pos.z, psDroid->body, psDroid->originalBody,
	           (int)o.type, o.pos.x, o.pos.y, o.pos2.x, o.pos2.y, (int)o.direction, o.psObj ? o.psObj->id : 0,
	           o.psStats ? (int)(o.psStats - asStructureStats) : -1, (int)o.rtrType, psDroid->listSize, (int)psDroid->asOrderList.size(),
	           (int)psDroid->action, psDroid->actionPos.x, psDroid->actionPos.y, t ? t->id : 0, psDroid->actionStarted, psDroid->actionPoints,
	           psDroid->secondaryOrder, psDroid->psBaseStruct ? psDroid->psBaseStruct->id : 0, (int)psDroid->group, (int)psDroid->repairGroup,
	           (int)psDroid->asWeaps[0].rot.direction, (int)psDroid->asWeaps[0].rot.pitch, (int)psDroid->underRepair);
}

// a structure's building state
static std::string orderStructLine(STRUCTURE const *s)
{
	return fmt("struct %u: p%d %s status%d body%u/%u pts%d rate%d cap%d res%d pos(%d,%d,%d)\n", s->id, (int)s->player,
	           s->pStructureType->id.toUtf8().c_str(), (int)s->status, s->body, s->structureBody(), s->currentBuildPts, s->buildRate,
	           (int)s->capacity, (int)s->resistance, s->pos.x, s->pos.y, s->pos.z);
}

static STRUCTURE *randomStructure(uint64_t &st, int player)
{
	int n = (int)apsStructLists[player].size();
	if (n == 0)
	{
		return nullptr;
	}
	int k = testRandMod(st, n);
	for (STRUCTURE *s : apsStructLists[player])
	{
		if (k-- == 0)
		{
			return s;
		}
	}
	return nullptr;
}

static DROID *randomDroid(uint64_t &st, int player)
{
	int n = (int)apsDroidLists[player].size();
	if (n == 0)
	{
		return nullptr;
	}
	int k = testRandMod(st, n);
	for (DROID *d : apsDroidLists[player])
	{
		if (k-- == 0)
		{
			return d;
		}
	}
	return nullptr;
}

// a droid or structure of a player not allied to player
static BASE_OBJECT *randomEnemy(uint64_t &st, int player)
{
	int q = testRandMod(st, MAX_PLAYERS);
	for (int i = 0; i < MAX_PLAYERS; i++, q = (q + 1) % MAX_PLAYERS)
	{
		if (q == player || aiCheckAlliances(q, player) || (apsDroidLists[q].empty() && apsStructLists[q].empty()))
		{
			continue;
		}
		if (apsStructLists[q].empty())
		{
			return randomDroid(st, q);
		}
		if (!apsDroidLists[q].empty() && testRandMod(st, 2) == 0)
		{
			return randomDroid(st, q);
		}
		return randomStructure(st, q);
	}
	return nullptr;
}

// a point up to tiles away from pos, on the map
static Vector2i nearPoint(uint64_t &st, Vector3i pos, int tiles, int wx, int wy)
{
	int rx = testRandMod(st, 2 * tiles * TILE_UNITS + 1) - tiles * TILE_UNITS;
	int ry = testRandMod(st, 2 * tiles * TILE_UNITS + 1) - tiles * TILE_UNITS;
	return Vector2i(clip(pos.x + rx, TILE_UNITS, wx - TILE_UNITS - 1), clip(pos.y + ry, TILE_UNITS, wy - TILE_UNITS - 1));
}

// an order for a droid, as a player could give it (what it was); with ModeQueue, through the game messages
static std::string giveOrder(uint64_t &st, DROID *psDroid, std::vector<STRUCTURE_STATS *> const &avail, int wall, int wx, int wy, QUEUE_MODE mode = ModeImmediate)
{
	int r = testRandMod(st, 100);
	int p = psDroid->player;
	if (psDroid->isVtol())
	{
		Vector2i t = nearPoint(st, psDroid->pos, 10, wx, wy);
		if (r < 30)
		{
			orderDroidLoc(psDroid, DORDER_CIRCLE, t.x, t.y, mode);
			return fmt("circle %d,%d", t.x, t.y);
		}
		if (r < 60)
		{
			orderDroidLoc(psDroid, DORDER_PATROL, t.x, t.y, mode);
			return fmt("patrol %d,%d", t.x, t.y);
		}
		if (r < 85)
		{
			BASE_OBJECT *e = randomEnemy(st, p);
			if (e)
			{
				orderDroidObj(psDroid, DORDER_ATTACK, e, mode);
				return fmt("attack %u", e->id);
			}
		}
		orderDroid(psDroid, DORDER_RTB, mode);
		return "rtb";
	}
	switch (psDroid->droidType)
	{
	case DROID_CONSTRUCT:
	case DROID_CYBORG_CONSTRUCT:
		if (r < 40 && !avail.empty())
		{
			STRUCTURE_STATS *s = avail[testRandMod(st, (int)avail.size())];
			Vector2i t = nearPoint(st, psDroid->pos, 3, wx, wy);
			int x = world_coord(map_coord(t.x)) + TILE_UNITS / 2;
			int y = world_coord(map_coord(t.y)) + TILE_UNITS / 2;
			uint16_t dir = (uint16_t)(testRandMod(st, 4) * DEG(90));
			orderDroidStatsLocDir(psDroid, DORDER_BUILD, s, x, y, dir, mode);
			return fmt("build %s %d,%d %d", s->id.toUtf8().c_str(), x, y, (int)dir);
		}
		if (r < 50 && wall >= 0)
		{
			Vector2i t = nearPoint(st, psDroid->pos, 3, wx, wy);
			int len = testRandMod(st, 5) + 1;
			int x1 = world_coord(map_coord(t.x)) + TILE_UNITS / 2;
			int y1 = world_coord(map_coord(t.y)) + TILE_UNITS / 2;
			int x2 = x1, y2 = y1;
			if (testRandMod(st, 2) == 0)
			{
				x2 = clip(x1 + len * TILE_UNITS, TILE_UNITS, wx - TILE_UNITS - 1);
			}
			else
			{
				y2 = clip(y1 + len * TILE_UNITS, TILE_UNITS, wy - TILE_UNITS - 1);
			}
			orderDroidStatsTwoLocDir(psDroid, DORDER_LINEBUILD, &asStructureStats[wall], x1, y1, x2, y2, 0, mode);
			return fmt("linebuild %d,%d %d,%d", x1, y1, x2, y2);
		}
		if (r < 60)
		{
			STRUCTURE *s = randomStructure(st, p);
			if (s)
			{
				orderDroidObj(psDroid, DORDER_DEMOLISH, s, mode);
				return fmt("demolish %u", s->id);
			}
		}
		else if (r < 75)
		{
			STRUCTURE *s = randomStructure(st, p);
			if (s)
			{
				orderDroidObj(psDroid, DORDER_REPAIR, s, mode);
				return fmt("repair %u", s->id);
			}
		}
		break;
	case DROID_REPAIR:
	case DROID_CYBORG_REPAIR:
		if (r < 50)
		{
			DROID *d = randomDroid(st, p);
			if (d && d != psDroid)
			{
				orderDroidObj(psDroid, DORDER_DROIDREPAIR, d, mode);
				return fmt("droidrepair %u", d->id);
			}
		}
		break;
	case DROID_SENSOR:
		if (r < 50)
		{
			BASE_OBJECT *e = randomEnemy(st, p);
			if (e)
			{
				orderDroidObj(psDroid, DORDER_OBSERVE, e, mode);
				return fmt("observe %u", e->id);
			}
		}
		break;
	default:
		if (r < 25)
		{
			BASE_OBJECT *e = randomEnemy(st, p);
			if (e)
			{
				orderDroidObj(psDroid, DORDER_ATTACK, e, mode);
				return fmt("attack %u", e->id);
			}
		}
		else if (r < 35)
		{
			STRUCTURE *s = randomStructure(st, p);
			if (s)
			{
				orderDroidObj(psDroid, DORDER_GUARD, s, mode);
				return fmt("guard %u", s->id);
			}
		}
		break;
	}
	// the orders every droid takes
	int r2 = testRandMod(st, 100);
	Vector2i t = nearPoint(st, psDroid->pos, 10, wx, wy);
	if (r2 < 30)
	{
		orderDroidLoc(psDroid, DORDER_MOVE, t.x, t.y, mode);
		return fmt("move %d,%d", t.x, t.y);
	}
	if (r2 < 45)
	{
		orderDroidLoc(psDroid, DORDER_SCOUT, t.x, t.y, mode);
		return fmt("scout %d,%d", t.x, t.y);
	}
	if (r2 < 55)
	{
		orderDroidLoc(psDroid, DORDER_PATROL, t.x, t.y, mode);
		return fmt("patrol %d,%d", t.x, t.y);
	}
	if (r2 < 62)
	{
		orderDroid(psDroid, DORDER_HOLD, mode);
		return "hold";
	}
	if (r2 < 70)
	{
		orderDroid(psDroid, DORDER_RTB, mode);
		return "rtb";
	}
	if (r2 < 78)
	{
		orderDroid(psDroid, DORDER_RTR, mode);
		return "rtr";
	}
	if (r2 < 83)
	{
		orderDroid(psDroid, DORDER_STOP, mode);
		return "stop";
	}
	if (r2 < 87)
	{
		orderDroid(psDroid, DORDER_RECYCLE, mode);
		return "recycle";
	}
	if (r2 < 95)
	{
		// a click on an object (chooseOrderObj, as orderSelectedObjAdd gives it)
		BASE_OBJECT *o = nullptr;
		if (testRandMod(st, 2) == 0)
		{
			o = randomEnemy(st, p);
		}
		else
		{
			o = randomStructure(st, p);
		}
		if (o == nullptr)
		{
			return "click none";
		}
		DroidOrder ord = chooseOrderObj(psDroid, o, false);
		if (ord.type != DORDER_NONE && validOrderForObj(ord.type))
		{
			orderDroidObj(psDroid, ord.type, ord.psObj, mode);
		}
		return fmt("click %u: %s", o->id, getDroidOrderName(ord.type));
	}
	// a move, then two queued
	orderDroidLoc(psDroid, DORDER_MOVE, t.x, t.y, mode);
	Vector2i t2 = nearPoint(st, psDroid->pos, 10, wx, wy);
	DroidOrder q1(DORDER_SCOUT, t2);
	if (mode == ModeQueue)
	{
		sendDroidInfo(psDroid, q1, true);
	}
	else
	{
		orderDroidAdd(psDroid, &q1);
	}
	Vector2i t3 = nearPoint(st, psDroid->pos, 10, wx, wy);
	DroidOrder q2(DORDER_MOVE, t3);
	if (mode == ModeQueue)
	{
		sendDroidInfo(psDroid, q2, true);
	}
	else
	{
		orderDroidAdd(psDroid, &q2);
	}
	return fmt("queue %d,%d %d,%d %d,%d", t.x, t.y, t2.x, t2.y, t3.x, t3.y);
}

static void mapOrders()
{
	uint64_t st = 11 + mapWidth * 17 + mapHeight * 3;
	int wx = world_coord(mapWidth);
	int wy = world_coord(mapHeight);
	fpathInitialise();
	gridInitialise();
	gameTime = 1000;
	deltaGameTime = GAME_TICKS_PER_UPDATE;
	allocPlayerPower();
	// the structures the trucks may build, buildable by every player
	static const char *const buildable[] = {"A0LightFactory", "A0PowerGenerator", "A0ResearchFacility", "A0HardcreteMk1Wall", "WallTower01",
	                                        "A0RepairCentre3", "Sys-SensoTower02", "A0CommandCentre", "A0ResourceExtractor", "A0FacMod1",
	                                        "A0PowMod1", "A0ResearchModule1", "PillBox1", "A0CyborgFactory", "A0VTolFactory1", "A0VtolPad"};
	std::vector<STRUCTURE_STATS *> avail;
	for (const char *name : buildable)
	{
		int i = getStructStatFromName(WzString::fromUtf8(name));
		if (i >= 0)
		{
			avail.push_back(&asStructureStats[i]);
			for (int p = 0; p < MAX_PLAYERS; p++)
			{
				apStructTypeLists[p][i] = AVAILABLE;
			}
		}
	}
	int wall = getStructStatFromName(WzString::fromUtf8("A0HardcreteMk1Wall"));
	for (int p = 0; p < MAX_PLAYERS; p++)
	{
		setPower(p, 1000 + 500 * (p % 3));
	}
	// damage: a third of the droids and a quarter of the built structures
	std::vector<DROID *> droids = allDroids();
	for (size_t i = 0; i < droids.size(); i++)
	{
		// DROID's constructor leaves actionStarted and actionPoints uninitialised (a slot's earlier droid's, here); the
		// game sets them before it reads them, the test prints them: 0, as the VB port's droids start
		droids[i]->actionStarted = 0;
		droids[i]->actionPoints = 0;
		if (i % 3 == 1)
		{
			droids[i]->body = std::max<int>(1, droids[i]->body * 2 / 5);
		}
	}
	int k = 0;
	for (int p = 0; p < MAX_PLAYERS; p++)
	{
		for (STRUCTURE *s : apsStructLists[p])
		{
			if (s->status == SS_BUILT && k++ % 4 == 2)
			{
				s->body = std::max<int>(1, s->body / 2);
			}
		}
	}
	// secondary orders
	for (size_t i = 0; i < droids.size() && i < 64; i++)
	{
		DROID *psDroid = droids[i];
		if (i % 7 == 3)
		{
			secondarySetState(psDroid, DSO_HALTTYPE, i % 2 ? DSS_HALT_HOLD : DSS_HALT_PURSUE, ModeImmediate);
		}
		else if (i % 7 == 5)
		{
			secondarySetState(psDroid, DSO_ATTACK_RANGE, DSS_ARANGE_SHORT, ModeImmediate);
		}
		else if (i % 11 == 7)
		{
			secondarySetState(psDroid, DSO_REPAIR_LEVEL, DSS_REPLEV_HIGH, ModeImmediate);
		}
		else
		{
			continue;
		}
		putCaptured();
		put("secondary %u: %X %d\n", psDroid->id, psDroid->secondaryOrder, (int)psDroid->order.type);
	}
	gridReset();
	putCaptured();
	// the orders
	for (size_t i = 0; i < droids.size() && i < 64; i++)
	{
		DROID *psDroid = droids[i];
		if (psDroid->died)
		{
			continue;
		}
		std::string how = giveOrder(st, psDroid, avail, wall, wx, wy);
		putCaptured();
		put("order %u: %s\n", psDroid->id, how.c_str());
		put("%s", orderDroidLine(psDroid).c_str());
	}
	for (int tick = 1; tick <= 90; tick++)
	{
		gameTime += GAME_TICKS_PER_UPDATE;
		gridReset();
		fpathUpdate();
		if (tick == 20 || tick == 40)
		{
			// new orders on the way
			for (size_t i = 0; i < droids.size() && i < 64; i++)
			{
				DROID *psDroid = droids[i];
				if (psDroid->died || (int)(i % 3) != tick / 20 - 1)
				{
					continue;
				}
				std::string how = giveOrder(st, psDroid, avail, wall, wx, wy);
				putCaptured();
				put("reorder %u: %s\n", psDroid->id, how.c_str());
			}
		}
		for (int p = 0; p < MAX_PLAYERS; p++)
		{
			mutating_list_iterate(apsDroidLists[p], [](DROID *psDroid)
			{
				droidUpdate(psDroid);
				return IterationResult::CONTINUE_ITERATION;
			});
		}
		for (int p = 0; p < MAX_PLAYERS; p++)
		{
			updatePlayerPower(p);
		}
		putCaptured();
		if (tick == 1 || tick == 90 || syncDumpAll)
		{
			size_t len = dumpLocalDebugSyncLog(syncDebugNext);
			out.append((const char *)debugSyncTmpBuf.data(), len);
		}
		unsigned crc = nextDebugSync();
		putCaptured();
		put("tick %d: crc=%04X\n", tick, crc);
		if (tick % 30 == 0)
		{
			for (size_t i = 0; i < droids.size() && i < 64; i++)
			{
				if (!droids[i]->died)
				{
					put("%s", orderDroidLine(droids[i]).c_str());
				}
			}
		}
	}
	std::string s = "count:";
	for (int p = 0; p < MAX_PLAYERS; p++)
	{
		s += fmt(" %d/%d", (int)apsDroidLists[p].size(), (int)apsStructLists[p].size());
		for (STRUCTURE *psStruct : apsStructLists[p])
		{
			if (psStruct->status != SS_BUILT || psStruct->body != psStruct->structureBody())
			{
				put("%s", orderStructLine(psStruct).c_str());
			}
		}
	}
	put("%s\n", s.c_str());
	put("features %d\n", (int)apsFeatureList[0].size());
	s = "power:";
	for (int p = 0; p < MAX_PLAYERS; p++)
	{
		s += fmt(" %d", (int)getPower(p));
	}
	put("%s\n", s.c_str());
	gridShutDown();
	fpathShutdown();
	putCaptured();
	gameTime = 0;
	deltaGameTime = 0;
}

// a droid's combat state
static std::string combatDroidLine(DROID const *psDroid)
{
	BASE_OBJECT const *t = psDroid->psActionTarget[0];
	WEAPON const &w = psDroid->asWeaps[0];
	return fmt("droid %u: p%d t%d %s pos(%d,%d,%d) body%u/%u order%d obj%u action%d tgt%u exp%u kills%u"
	           " w(%u,%u,%u,%u,%d,%d) hit(%u,%d) vis%d,%d,%d,%d\n",
	           psDroid->id, (int)psDroid->player, (int)psDroid->droidType, moveDescription(psDroid->sMove.Status),
	           psDroid->pos.x, psDroid->pos.y, psDroid->pos.z, psDroid->body, psDroid->originalBody, (int)psDroid->order.type,
	           psDroid->order.psObj ? psDroid->order.psObj->id : 0, (int)psDroid->action, t ? t->id : 0, psDroid->experience, psDroid->kills,
	           w.nStat, w.ammo, w.lastFired, w.shotsFired, (int)w.rot.direction, (int)w.rot.pitch, psDroid->timeLastHit, (int)psDroid->lastHitWeapon,
	           (int)psDroid->visible[0], (int)psDroid->visible[1], (int)psDroid->visible[2], (int)psDroid->visible[3]);
}

// a structure of a player not allied to player
static STRUCTURE *randomEnemyStructure(uint64_t &st, int player)
{
	int q = testRandMod(st, MAX_PLAYERS);
	for (int i = 0; i < MAX_PLAYERS; i++, q = (q + 1) % MAX_PLAYERS)
	{
		if (q == player || aiCheckAlliances(q, player) || apsStructLists[q].empty())
		{
			continue;
		}
		return randomStructure(st, q);
	}
	return nullptr;
}

static void mapCombat()
{
	uint64_t st = 13 + mapWidth * 19 + mapHeight * 11;
	int wx = world_coord(mapWidth);
	int wy = world_coord(mapHeight);
	fpathInitialise();
	gridInitialise();
	gameTime = 1000;
	deltaGameTime = GAME_TICKS_PER_UPDATE;
	resetDamageModifiers();
	proj_InitSystem();
	// up to 40 armed droids moved near an enemy structure, and ordered
	std::vector<DROID *> fighters;
	for (DROID *psDroid : allDroids())
	{
		if (fighters.size() < 40 && psDroid->numWeaps > 0 && psDroid->asWeaps[0].nStat > 0 && !psDroid->isVtol() && !psDroid->isTransporter())
		{
			fighters.push_back(psDroid);
		}
	}
	for (DROID *psDroid : fighters)
	{
		STRUCTURE *psTarget = randomEnemyStructure(st, psDroid->player);
		if (psTarget == nullptr)
		{
			continue;
		}
		PROPULSION_TYPE prop = psDroid->getPropulsionStats()->propulsionType;
		int x = 0, y = 0;
		bool found = false;
		for (int tries = 0; tries < 20 && !found; tries++)
		{
			int ang = testRandMod(st, 65536);
			int r = (3 + testRandMod(st, 4)) * TILE_UNITS;
			x = psTarget->pos.x + iSinR(ang, r);
			y = psTarget->pos.y + iCosR(ang, r);
			found = x >= TILE_UNITS && y >= TILE_UNITS && x < wx - TILE_UNITS && y < wy - TILE_UNITS
			        && !fpathBlockingTile(map_coord(x), map_coord(y), prop) && !TileIsOccupied(mapTile(map_coord(x), map_coord(y)));
		}
		if (!found)
		{
			put("fighter %u: no place near %u\n", psDroid->id, psTarget->id);
			continue;
		}
		droidSetPosition(psDroid, x, y);
		const char *how;
		switch (testRandMod(st, 4))
		{
		case 0:
			orderDroidObj(psDroid, DORDER_ATTACK, psTarget, ModeImmediate);
			how = "attack";
			break;
		case 1:
			orderDroidLoc(psDroid, DORDER_SCOUT, psTarget->pos.x, psTarget->pos.y, ModeImmediate);
			how = "scout";
			break;
		case 2:
			orderDroid(psDroid, DORDER_HOLD, ModeImmediate);
			how = "hold";
			break;
		default:
			orderDroid(psDroid, DORDER_STOP, ModeImmediate);
			how = "stop";
			break;
		}
		putCaptured();
		put("fighter %u: ->%d,%d near %u %s\n", psDroid->id, x, y, psTarget->id, how);
	}
	gridReset();
	putCaptured();
	for (int tick = 1; tick <= 60; tick++)
	{
		gameTime += GAME_TICKS_PER_UPDATE;
		visUpdateLevel();
		gridReset();
		processVisibility();
		fpathUpdate();
		cmdDroidUpdate();
		for (int p = 0; p < MAX_PLAYERS; p++)
		{
			updatePlayerPower(p);
			mutating_list_iterate(apsDroidLists[p], [](DROID *psDroid)
			{
				droidUpdate(psDroid);
				return IterationResult::CONTINUE_ITERATION;
			});
		}
		proj_UpdateAll();
		objmemUpdate();
		putCaptured();
		if (tick == 1 || tick == 60 || syncDumpAll)
		{
			size_t len = dumpLocalDebugSyncLog(syncDebugNext);
			out.append((const char *)debugSyncTmpBuf.data(), len);
		}
		int nProj = 0;
		for (PROJECTILE *p = proj_GetFirst(); p != nullptr; p = proj_GetNext())
		{
			nProj++;
		}
		unsigned crc = nextDebugSync();
		putCaptured();
		put("tick %d: crc=%04X proj=%d\n", tick, crc, nProj);
		if (tick % 20 == 0)
		{
			for (DROID *psDroid : allDroids())
			{
				if (std::find(fighters.begin(), fighters.end(), psDroid) != fighters.end())
				{
					put("%s", combatDroidLine(psDroid).c_str());
				}
			}
		}
	}
	std::string s = "count:";
	for (int p = 0; p < MAX_PLAYERS; p++)
	{
		s += fmt(" %d/%d", (int)apsDroidLists[p].size(), (int)apsStructLists[p].size());
		for (STRUCTURE *psStruct : apsStructLists[p])
		{
			if (psStruct->body != psStruct->structureBody())
			{
				put("%s", orderStructLine(psStruct).c_str());
			}
		}
	}
	put("%s\n", s.c_str());
	put("features %d\n", (int)apsFeatureList[0].size());
	proj_FreeAllProjectiles();
	gridShutDown();
	fpathShutdown();
	putCaptured();
	gameTime = 0;
	deltaGameTime = 0;
}

// (mapbase) what research and the test change in the stats the maps share, restored after the test
void statsref_researchState(bool save);  // research.cpp's extract: bSelfRepair, playerUpgradeCounts, cachedPerPlayerRawUpgradeChange
static std::vector<PLAYER_RESEARCH> savedResList[MAX_PLAYERS];
static std::vector<UBYTE> savedCompLists[MAX_PLAYERS][COMP_NUMCOMPONENTS];
static std::vector<UBYTE> savedStructTypeLists[MAX_PLAYERS];
static UDWORD savedDefaultSensor[MAX_PLAYERS], savedDefaultECM[MAX_PLAYERS], savedDefaultRepair[MAX_PLAYERS];

// the length of apCompLists[player][t] (allocComponentList's)
static size_t compListSize(int t)
{
	switch (t)
	{
	case COMP_BODY: return asBodyStats.size();
	case COMP_BRAIN: return asBrainStats.size();
	case COMP_PROPULSION: return asPropulsionStats.size();
	case COMP_REPAIRUNIT: return asRepairStats.size();
	case COMP_ECM: return asECMStats.size();
	case COMP_SENSOR: return asSensorStats.size();
	case COMP_CONSTRUCT: return asConstructStats.size();
	case COMP_WEAPON: return asWeaponStats.size();
	}
	return 0;
}

template <typename T>
static void upgradeSnapshot(bool save, T *stats, size_t n, std::vector<std::decay_t<decltype(stats->upgrade[0])>> &saved)
{
	if (save)
	{
		saved.clear();
	}
	size_t k = 0;
	for (size_t i = 0; i < n; i++)
	{
		for (int p = 0; p < MAX_PLAYERS; p++)
		{
			if (save)
			{
				saved.push_back(stats[i].upgrade[p]);
			}
			else
			{
				stats[i].upgrade[p] = saved[k++];
			}
		}
	}
}

static void baseSnapshot(bool save)
{
	static std::vector<std::decay_t<decltype(asWeaponStats[0].upgrade[0])>> w;
	static std::vector<std::decay_t<decltype(asBodyStats[0].upgrade[0])>> b;
	static std::vector<std::decay_t<decltype(asBrainStats[0].upgrade[0])>> br;
	static std::vector<std::decay_t<decltype(asPropulsionStats[0].upgrade[0])>> pr;
	static std::vector<std::decay_t<decltype(asSensorStats[0].upgrade[0])>> se;
	static std::vector<std::decay_t<decltype(asECMStats[0].upgrade[0])>> ec;
	static std::vector<std::decay_t<decltype(asRepairStats[0].upgrade[0])>> re;
	static std::vector<std::decay_t<decltype(asConstructStats[0].upgrade[0])>> co;
	static std::vector<std::decay_t<decltype(asStructureStats[0].upgrade[0])>> stu;
	for (int p = 0; p < MAX_PLAYERS; p++)
	{
		if (save)
		{
			savedResList[p] = asPlayerResList[p];
			savedDefaultSensor[p] = aDefaultSensor[p];
			savedDefaultECM[p] = aDefaultECM[p];
			savedDefaultRepair[p] = aDefaultRepair[p];
		}
		else
		{
			asPlayerResList[p] = savedResList[p];
			aDefaultSensor[p] = savedDefaultSensor[p];
			aDefaultECM[p] = savedDefaultECM[p];
			aDefaultRepair[p] = savedDefaultRepair[p];
		}
		for (int t = 0; t < COMP_NUMCOMPONENTS; t++)
		{
			size_t n = apCompLists[p][t] ? compListSize(t) : 0;
			if (save)
			{
				savedCompLists[p][t].assign(apCompLists[p][t], apCompLists[p][t] + n);
			}
			else if (n)
			{
				std::copy(savedCompLists[p][t].begin(), savedCompLists[p][t].end(), apCompLists[p][t]);
			}
		}
		if (apStructTypeLists[p])
		{
			if (save)
			{
				savedStructTypeLists[p].assign(apStructTypeLists[p], apStructTypeLists[p] + numStructureStats);
			}
			else
			{
				std::copy(savedStructTypeLists[p].begin(), savedStructTypeLists[p].end(), apStructTypeLists[p]);
			}
		}
	}
	upgradeSnapshot(save, asWeaponStats.data(), asWeaponStats.size(), w);
	upgradeSnapshot(save, asBodyStats.data(), asBodyStats.size(), b);
	upgradeSnapshot(save, asBrainStats.data(), asBrainStats.size(), br);
	upgradeSnapshot(save, asPropulsionStats.data(), asPropulsionStats.size(), pr);
	upgradeSnapshot(save, asSensorStats.data(), asSensorStats.size(), se);
	upgradeSnapshot(save, asECMStats.data(), asECMStats.size(), ec);
	upgradeSnapshot(save, asRepairStats.data(), asRepairStats.size(), re);
	upgradeSnapshot(save, asConstructStats.data(), asConstructStats.size(), co);
	upgradeSnapshot(save, asStructureStats, numStructureStats, stu);
	statsref_researchState(save);
}

// a topic the player has neither started nor completed, from a random start; -1 if none
static int chooseResearch(uint64_t &st, int player)
{
	int n = (int)asResearch.size();
	if (n == 0)
	{
		return -1;
	}
	int start = testRandMod(st, n);
	for (int k = 0; k < n; k++)
	{
		int i = (start + k) % n;
		if (!IsResearchStarted(&asPlayerResList[player][i]) && !IsResearchCompleted(&asPlayerResList[player][i]))
		{
			return i;
		}
	}
	return -1;
}

// recvResearchStatus's start of a topic (the GAME_RESEARCHSTATUS message, bStart)
static void startResearch(STRUCTURE *psBuilding, int index)
{
	int player = psBuilding->player;
	PLAYER_RESEARCH *pPlayerRes = &asPlayerResList[player][index];
	ResetPendingResearchStatus(pPlayerRes);
	RESEARCH_FACILITY *psResFacilty = &psBuilding->pFunctionality->researchFacility;
	popStatusPending(*psResFacilty);
	if (psResFacilty->psSubject)
	{
		cancelResearch(psBuilding, ModeImmediate);
	}
	if (IsResearchStarted(pPlayerRes))
	{
		STRUCTURE *psOtherBuilding = findResearchingFacilityByResearchIndex(player, index);
		if (psOtherBuilding != nullptr)
		{
			cancelResearch(psOtherBuilding, ModeImmediate);
		}
	}
	psResFacilty->psSubject = &asResearch[index];
	MakeResearchStarted(pPlayerRes);
	psResFacilty->timeStartHold = 0;
}

// a template of the player's for the factory, from a random start; nullptr if none
static DROID_TEMPLATE *chooseTemplate(uint64_t &st, STRUCTURE *psFactory)
{
	std::vector<DROID_TEMPLATE *> t;
	enumerateTemplates(psFactory->player, [&t](DROID_TEMPLATE *psTempl) { t.push_back(psTempl); return true; });
	int n = (int)t.size();
	if (n == 0)
	{
		return nullptr;
	}
	int start = testRandMod(st, n);
	for (int k = 0; k < n; k++)
	{
		DROID_TEMPLATE *psTempl = t[(start + k) % n];
		if (validTemplateForFactory(psTempl, psFactory, false))
		{
			return psTempl;
		}
	}
	return nullptr;
}

// work for the idle labs and factories
static void baseGiveWork(uint64_t &st)
{
	for (int p = 0; p < MAX_PLAYERS; p++)
	{
		for (STRUCTURE *s : apsStructLists[p])
		{
			if (s->status == SS_BUILT && s->pFunctionality)
			{
				if (s->pStructureType->type == REF_RESEARCH)
				{
					if (s->pFunctionality->researchFacility.psSubject == nullptr)
					{
						int index = chooseResearch(st, p);
						if (index >= 0)
						{
							bool ok = researchAvailable(index, p, ModeImmediate);
							startResearch(s, index);
							putCaptured();
							put("lab %u: research %d avail%d\n", s->id, index, (int)ok);
						}
					}
				}
				else if (s->isFactory())
				{
					if (s->pFunctionality->factory.psSubject == nullptr)
					{
						DROID_TEMPLATE *psTempl = chooseTemplate(st, s);
						if (psTempl != nullptr)
						{
							bool ok = structSetManufacture(s, psTempl, ModeImmediate);
							putCaptured();
							put("factory %u: %s ok%d\n", s->id, psTempl->id.toUtf8().c_str(), (int)ok);
						}
					}
				}
			}
		}
	}
}

static std::string baseStructLine(STRUCTURE const *s)
{
	std::string w = fmt("struct %u: p%d %s status%d body%u res%d tgt%u", s->id, (int)s->player, s->pStructureType->id.toUtf8().c_str(),
	                    (int)s->status, s->body, (int)s->resistance, s->psTarget[0] ? s->psTarget[0]->id : 0);
	if (s->pFunctionality)
	{
		switch (s->pStructureType->type)
		{
		case REF_RESEARCH:
			if (s->pFunctionality->researchFacility.psSubject)
			{
				int subj = s->pFunctionality->researchFacility.psSubject->index;
				w += fmt(" research %d pts%u", subj, asPlayerResList[s->player][subj].currentPoints);
			}
			else
			{
				w += " research none";
			}
			break;
		case REF_FACTORY:
		case REF_CYBORG_FACTORY:
		case REF_VTOL_FACTORY:
			if (s->pFunctionality->factory.psSubject)
			{
				w += fmt(" making %s left%d started%u", s->pFunctionality->factory.psSubject->id.toUtf8().c_str(),
				         s->pFunctionality->factory.buildPointsRemaining, s->pFunctionality->factory.timeStarted);
			}
			else
			{
				w += " making none";
			}
			break;
		case REF_REPAIR_FACILITY:
			w += fmt(" repairing %u state%d", s->pFunctionality->repairFacility.psObj ? s->pFunctionality->repairFacility.psObj->id : 0,
			         (int)s->pFunctionality->repairFacility.state);
			break;
		case REF_REARM_PAD:
			w += fmt(" rearming %u", s->pFunctionality->rearmPad.psObj ? s->pFunctionality->rearmPad.psObj->id : 0);
			break;
		default:
			break;
		}
	}
	return w + "\n";
}

void gameStateUpdate();  // loop.cpp's extract (not static: see STATSREF_SUBSTITUTIONS)

static void mapBase()
{
	uint64_t st = 17 + mapWidth * 23 + mapHeight * 13;
	int limits[MAX_PLAYERS][3];
	fpathInitialise();
	gridInitialise();
	// after the earlier blocks' times (the droids' lastFired)
	gameTime = 20000;
	deltaGameTime = GAME_TICKS_PER_UPDATE;
	sstrcpy(game.map, "mapbase");
	resetDamageModifiers();
	proj_InitSystem();
	mapInit();
	baseSnapshot(true);
	// research and production 50 times faster; every component available (the templates researched)
	for (unsigned i = 0; i < numStructureStats; i++)
	{
		for (int p = 0; p < MAX_PLAYERS; p++)
		{
			auto &u = asStructureStats[i].upgrade[p];
			u.research *= 50;
			u.moduleResearch *= 50;
			u.production *= 50;
			u.moduleProduction *= 50;
		}
	}
	for (int p = 0; p < MAX_PLAYERS; p++)
	{
		for (int t = 0; t < COMP_NUMCOMPONENTS; t++)
		{
			size_t n = apCompLists[p][t] ? compListSize(t) : 0;
			for (size_t c = 0; c < n; c++)
			{
				apCompLists[p][t][c] = AVAILABLE;
			}
		}
		setPower(p, 3000 + 1000 * (p % 3));
		// rules.js's droid limits
		limits[p][0] = getMaxDroids(p);
		limits[p][1] = getMaxCommanders(p);
		limits[p][2] = getMaxConstructors(p);
		setMaxDroids(p, 150);
		setMaxCommanders(p, 10);
		setMaxConstructors(p, 15);
	}
	// damaged droids go for repair, VTOLs to rearm
	for (DROID *psDroid : allDroids())
	{
		if (psDroid->isVtol())
		{
			moveToRearm(psDroid);
			putCaptured();
			put("rearm %u: action%d\n", psDroid->id, (int)psDroid->action);
		}
		else if (psDroid->isDamaged())
		{
			if (testRandMod(st, 2) == 0)
			{
				orderDroid(psDroid, DORDER_RTR, ModeImmediate);
				putCaptured();
				put("rtr %u: order%d\n", psDroid->id, (int)psDroid->order.type);
			}
		}
	}
	baseGiveWork(st);
	countUpdate(false);
	gridReset();
	putCaptured();
	for (int tick = 1; tick <= 100; tick++)
	{
		gameTime += GAME_TICKS_PER_UPDATE;
		deltaGameTime = GAME_TICKS_PER_UPDATE;
		gameStateUpdate();
		putCaptured();
		if (tick == 1 || tick == 100 || syncDumpAll)
		{
			size_t len = dumpLocalDebugSyncLog(syncDebugNext);
			out.append((const char *)debugSyncTmpBuf.data(), len);
		}
		int nProj = 0;
		for (PROJECTILE *p = proj_GetFirst(); p != nullptr; p = proj_GetNext())
		{
			nProj++;
		}
		unsigned crc = nextDebugSync();
		putCaptured();
		put("tick %d: crc=%04X proj=%d droids=%d\n", tick, crc, nProj, (int)allDroids().size());
		if (tick % 25 == 0)
		{
			for (int p = 0; p < MAX_PLAYERS; p++)
			{
				for (STRUCTURE *s : apsStructLists[p])
				{
					if (s->pFunctionality || s->numWeaps > 0)
					{
						put("%s", baseStructLine(s).c_str());
					}
				}
			}
			baseGiveWork(st);
		}
	}
	std::string s = "count:";
	for (int p = 0; p < MAX_PLAYERS; p++)
	{
		int n = 0;
		for (size_t i = 0; i < asResearch.size(); i++)
		{
			if (IsResearchCompleted(&asPlayerResList[p][i]))
			{
				n++;
			}
		}
		s += fmt(" %d/%d/%d/%d", (int)apsDroidLists[p].size(), (int)apsStructLists[p].size(), n, (int)getPower(p));
	}
	put("%s\n", s.c_str());
	put("features %d\n", (int)apsFeatureList[0].size());
	for (int p = 0; p < MAX_PLAYERS; p++)
	{
		setMaxDroids(p, limits[p][0]);
		setMaxCommanders(p, limits[p][1]);
		setMaxConstructors(p, limits[p][2]);
	}
	baseSnapshot(false);
	proj_FreeAllProjectiles();
	gridShutDown();
	fpathShutdown();
	putCaptured();
	gameTime = 0;
	deltaGameTime = 0;
}

// mapnet: what wzGetTicks gives (it only moves in mapnet, 100 ms a frame, across the maps)
static uint32_t netTicks = 0;
uint32_t wzGetTicks()
{
	return netTicks;
}

#define STATSREF_MAX_SYNC_HISTORY 12                       // sync_debug.cpp's MAX_SYNC_HISTORY

// the header line of the sync log nextDebugSync closed last
static std::string netLastLogHeader()
{
	size_t len = dumpLocalDebugSyncLog((syncDebugNext + STATSREF_MAX_SYNC_HISTORY - 1) % STATSREF_MAX_SYNC_HISTORY);
	std::string s((const char *)debugSyncTmpBuf.data(), len);
	size_t k = s.find('\n');
	if (k != std::string::npos)
	{
		s = s.substr(0, k + 1);
	}
	return s;
}

// a topic the player may start (researchAvailable, ModeQueue, as wzapi's pursueResearch asks), from a random start; -1 if none
static int chooseAvailableResearch(uint64_t &st, int player)
{
	int n = (int)asResearch.size();
	if (n == 0)
	{
		return -1;
	}
	int start = testRandMod(st, n);
	for (int k = 0; k < n; k++)
	{
		int i = (start + k) % n;
		if (researchAvailable(i, player, ModeQueue))
		{
			return i;
		}
	}
	return -1;
}

// work for the labs and factories through the game messages, as the interface and the scripts give it: a topic for
// each idle lab (sendResearchStatus, as wzapi's pursueResearch), a template for each idle factory (structSetManufacture,
// ModeQueue); the first busy lab and factory of each player held, released or cancelled (by round)
static void netGiveWork(uint64_t &st, int round)
{
	for (int p = 0; p < MAX_PLAYERS; p++)
	{
		bool labDone = false;
		bool facDone = false;
		for (STRUCTURE *s : apsStructLists[p])
		{
			if (s->status == SS_BUILT && s->pFunctionality)
			{
				if (s->pStructureType->type == REF_RESEARCH)
				{
					if (s->pFunctionality->researchFacility.psSubject == nullptr)
					{
						int index = chooseAvailableResearch(st, p);
						if (index >= 0)
						{
							sendResearchStatus(s, index, p, true);
							putCaptured();
							put("lab %u: research %d\n", s->id, index);
						}
					}
					else if (!labDone)
					{
						labDone = true;
						switch (round % 3)
						{
						case 0:
							holdResearch(s, ModeQueue);
							break;
						case 1:
							releaseResearch(s, ModeQueue);
							break;
						default:
							cancelResearch(s, ModeQueue);
							break;
						}
						putCaptured();
						put("lab %u: %d\n", s->id, round % 3);
					}
				}
				else if (s->isFactory())
				{
					if (s->pFunctionality->factory.psSubject == nullptr)
					{
						DROID_TEMPLATE *psTempl = chooseTemplate(st, s);
						if (psTempl != nullptr)
						{
							bool ok = structSetManufacture(s, psTempl, ModeQueue);
							putCaptured();
							put("factory %u: %s ok%d\n", s->id, psTempl->id.toUtf8().c_str(), (int)ok);
						}
					}
					else if (!facDone)
					{
						facDone = true;
						switch (round % 3)
						{
						case 0:
							holdProduction(s, ModeQueue);
							break;
						case 1:
							releaseProduction(s, ModeQueue);
							break;
						default:
							cancelProduction(s, ModeQueue);
							break;
						}
						putCaptured();
						put("factory %u: %d\n", s->id, round % 3);
					}
				}
			}
		}
	}
}

// orders through the game messages for up to 24 droids (a third of them, from a random start), secondary orders for
// some others
static void netOrders(uint64_t &st, std::vector<STRUCTURE_STATS *> const &avail, int wall, int wx, int wy)
{
	std::vector<DROID *> droids = allDroids();
	int nDroids = (int)droids.size();
	if (nDroids == 0)
	{
		return;
	}
	int k = testRandMod(st, 3);
	int n = 0;
	for (int i = 0; i < nDroids; i++)
	{
		DROID *psDroid = droids[i];
		if ((i + k) % 3 == 0)
		{
			if (n < 24)
			{
				n++;
				std::string how = giveOrder(st, psDroid, avail, wall, wx, wy, ModeQueue);
				putCaptured();
				put("order %u: %s\n", psDroid->id, how.c_str());
			}
		}
		else if (i % 5 == 1)
		{
			if (i % 2 != 0)
			{
				secondarySetState(psDroid, DSO_ATTACK_RANGE, DSS_ARANGE_SHORT, ModeQueue);
			}
			else
			{
				secondarySetState(psDroid, DSO_HALTTYPE, DSS_HALT_HOLD, ModeQueue);
			}
			putCaptured();
			put("secondary %u: %X\n", psDroid->id, (unsigned)psDroid->secondaryOrderPending);
		}
	}
}

// the debug mode's messages: refused in a skirmish (debugMappingsAllowed is false), done in the campaign
static void netDebugMessages(uint64_t &st)
{
	std::vector<DROID *> droids = allDroids();
	int nDroids = (int)droids.size();
	if (nDroids > 0)
	{
		DROID *psDroid = droids[testRandMod(st, nDroids)];
		SendDestroyDroid(psDroid);
		putCaptured();
		put("debug: destroy droid %u\n", psDroid->id);
		psDroid = droids[testRandMod(st, nDroids)];
		int p = psDroid->player;
		std::vector<DROID_TEMPLATE *> t;
		enumerateTemplates(p, [&t](DROID_TEMPLATE *psTempl) { t.push_back(psTempl); return true; });
		DROID_TEMPLATE *psTempl = nullptr;
		if (!t.empty())
		{
			psTempl = t[testRandMod(st, (int)t.size())];
		}
		if (psTempl != nullptr)
		{
			SendDroid(psTempl, psDroid->pos.x, psDroid->pos.y, p, generateNewObjectId(), nullptr);
			putCaptured();
			put("debug: droid %s for p%d\n", psTempl->id.toUtf8().c_str(), p);
		}
	}
	STRUCTURE *s = randomStructure(st, testRandMod(st, MAX_PLAYERS));
	if (s != nullptr)
	{
		SendDestroyStructure(s);
		putCaptured();
		put("debug: destroy structure %u\n", s->id);
	}
	s = randomStructure(st, testRandMod(st, MAX_PLAYERS));
	if (s != nullptr)
	{
		SendBuildFinished(s);
		putCaptured();
		put("debug: build finished %u\n", s->id);
	}
}

GAMECODE gameLoop();  // loop.cpp's extract

// the replay options of maprecord and mapplay (vbtest_vb's ReplayTestHandler does the same): the game's own
// (WZGameReplayOptionsHandler) for the options saved and the player info, a built-in map's embedded data (none), and
// on restoring the game's restoreOptions without its level part (the tests load their maps by folder, not as levels)
// and without its version warning
void from_json(const nlohmann::json& j, MULTIPLAYERGAME& p);     // multiint.cpp (not inline: see run_tests.py)
void from_json(const nlohmann::json& j, MULTIPLAYERINGAME& p);
void from_json(const nlohmann::json& j, PLAYER& p);
bool loadMultiStatsFromJSON(const nlohmann::json& json);          // multistat.cpp

class StatsrefReplayOptionsHandler : public ReplayOptionsHandler
{
public:
	bool saveOptions(nlohmann::json& object) const override
	{
		return static_cast<ReplayOptionsHandler const &>(gameOptions).saveOptions(object);
	}
	bool saveMap(EmbeddedMapData& mapData) const override
	{
		mapData.dataVersion = 1;
		return true;
	}
	bool optionsUpdatePlayerInfo(nlohmann::json& object) const override
	{
		return static_cast<ReplayOptionsHandler const &>(gameOptions).optionsUpdatePlayerInfo(object);
	}
	bool restoreOptions(const nlohmann::json& object, EmbeddedMapData&& embeddedMapData, uint32_t replay_netcodeMajor, uint32_t replay_netcodeMinor) override
	{
		// random seed
		uint32_t rand_seed = object.at("randSeed").get<uint32_t>();
		gameSRand(rand_seed);
		// restore `game`
		game = object.at("game").get<MULTIPLAYERGAME>();
		// restore `ingame`
		ingame = object.at("ingame").get<MULTIPLAYERINGAME>();
		// restore NetPlay.players
		auto netPlayers = object.at("netplay.players");
		for (size_t i = 0; i < netPlayers.size(); ++i)
		{
			from_json(netPlayers.at(i), NetPlay.players[i]);
		}
		// restore NetPlay.hostPlayer
		NetPlay.hostPlayer = object.at("netplay.hostPlayer").get<uint32_t>();
		// restore `NetPlay.bComms` (?)
		NetPlay.bComms = object.at("netplay.bComms").get<bool>();
		// restore multistats
		if (!loadMultiStatsFromJSON(object.at("multistats")))
		{
			return false;
		}
		// Set various other initialization things (see NET_FIREUP)
		ingame.TimeEveryoneIsInGame = nullopt;			// reset time
		ingame.endTime = nullopt;
		// set selectedPlayer to a *new* spectator slot
		NetPlay.players.push_back(PLAYER());
		size_t replaySpectatorIndex = NetPlay.players.size() - 1;
		NET_InitPlayer(replaySpectatorIndex, false);  // re-init everything
		NetPlay.players[replaySpectatorIndex].allocated = true;
		sstrcpy(NetPlay.players[replaySpectatorIndex].name, "Replay Viewer");
		NetPlay.players[replaySpectatorIndex].isSpectator = true;
		selectedPlayer = replaySpectatorIndex;
		realSelectedPlayer = selectedPlayer;
		// Need to initialize net/game queues here before we try to add messages to them
		for (unsigned i = 0; i < MAX_GAMEQUEUE_SLOTS; ++i)
		{
			NETinitQueue(NETgameQueue(i));
			NETsetNoSendOverNetwork(NETgameQueue(i));
		}
		// simulate "NET_PLAYERRESPONDING" sent by all clients
		// ensure that all humans players are set to "joined"
		for (size_t i = 0; i < MAX_CONNECTED_PLAYERS; ++i)
		{
			if (isHumanPlayer(i) && ingame.JoiningInProgress[i])
			{
				ingame.JoiningInProgress[i] = false;
			}
		}
		NetPlay.isHostAlive = true; // lie and say the host is alive for multiplayer games
		return true;
	}
	size_t desiredBufferSize() const override
	{
		return 0;
	}
	size_t maximumEmbeddedMapBufferSize() const override
	{
		return static_cast<ReplayOptionsHandler const &>(gameOptions).maximumEmbeddedMapBufferSize();
	}
private:
	WZGameReplayOptionsHandler gameOptions;
};

static uint32_t replayU32(std::vector<uint8_t> const &data, size_t pos)
{
	return (uint32_t)data[pos] << 24 | (uint32_t)data[pos + 1] << 16 | (uint32_t)data[pos + 2] << 8 | data[pos + 3];
}

// a replay file as text: its header, its settings (the JSON as written), the embedded map's size, the messages' count
// by type with a CRC of their bytes, and the end-of-game JSON
static std::string replaySummary(std::string const &filename)
{
	PHYSFS_file *f = PHYSFS_openRead(filename.c_str());
	if (f == nullptr)
	{
		return "replay file missing\n";
	}
	std::vector<uint8_t> data((size_t)PHYSFS_fileLength(f));
	if (!data.empty())
	{
		PHYSFS_readBytes(f, data.data(), data.size());
	}
	PHYSFS_close(f);
	std::string out = fmt("replay %d bytes, magic %08X\n", (int)data.size(), replayU32(data, 0));
	size_t pos = 4;
	uint32_t n = replayU32(data, pos);
	out += "settings: " + std::string((const char *)&data[pos + 4], n) + "\n";
	pos += 4 + n;
	out += fmt("map data v%u, %u bytes\n", replayU32(data, pos), replayU32(data, pos + 4));
	pos += 8 + replayU32(data, pos + 4);
	size_t start = pos;
	int counts[256] = {};
	int count = 0;
	while (pos + 4 <= data.size())
	{
		int type = data[pos + 1];
		int length = data[pos + 2] * 256 + data[pos + 3];
		counts[type]++;
		count++;
		pos += 4 + length;
		if (type == REPLAY_ENDED)
		{
			break;
		}
	}
	uint32_t crc = wz::crc_update(wz::crc_init(), &data[start], pos - start);
	out += fmt("messages %d, crc %08X:", count, crc);
	for (int i = 0; i < 256; i++)
	{
		if (counts[i] > 0)
		{
			out += fmt(" %s=%d", messageTypeToString(i), counts[i]);
		}
	}
	out += "\n";
	n = replayU32(data, pos);
	out += "end: " + std::string((const char *)&data[pos + 4], n) + fmt(" (%u)\n", replayU32(data, pos + 4 + n));
	return out;
}

// netRun's modes: mapnet, maprecord (mapnet's game, recorded as a replay), mapplay (that replay played back)
enum NetRunMode
{
	NET_MODE_NET, NET_MODE_RECORD, NET_MODE_PLAY
};
static std::string netReplayFile;                          // maprecord's replay, for mapplay

static void netRun(NetRunMode mode)
{
	uint64_t st = 29 + mapWidth * 31 + mapHeight * 7 + 1000 * mode;
	int nFrames = mode == NET_MODE_NET ? 80 : 40;
	int workEvery = mode == NET_MODE_NET ? 20 : 10;
	int wx = world_coord(mapWidth);
	int wy = world_coord(mapHeight);
	int limits[MAX_PLAYERS][3];
	fpathInitialise();
	gridInitialise();
	sstrcpy(game.map, "mapnet");
	resetDamageModifiers();
	proj_InitSystem();
	mapInit();
	baseSnapshot(true);
	// research and production 50 times faster; every component available (the templates researched)
	for (unsigned i = 0; i < numStructureStats; i++)
	{
		for (int p = 0; p < MAX_PLAYERS; p++)
		{
			auto &u = asStructureStats[i].upgrade[p];
			u.research *= 50;
			u.moduleResearch *= 50;
			u.production *= 50;
			u.moduleProduction *= 50;
		}
	}
	for (int p = 0; p < MAX_PLAYERS; p++)
	{
		for (int t = 0; t < COMP_NUMCOMPONENTS; t++)
		{
			size_t n = apCompLists[p][t] ? compListSize(t) : 0;
			for (size_t c = 0; c < n; c++)
			{
				apCompLists[p][t][c] = AVAILABLE;
			}
		}
		setPower(p, 4000 + 1000 * (p % 2));
		// rules.js's droid limits
		limits[p][0] = getMaxDroids(p);
		limits[p][1] = getMaxCommanders(p);
		limits[p][2] = getMaxConstructors(p);
		setMaxDroids(p, 150);
		setMaxCommanders(p, 10);
		setMaxConstructors(p, 15);
	}
	// the structures the trucks may build (as maporders)
	static const char *const buildable[] = {"A0LightFactory", "A0PowerGenerator", "A0ResearchFacility", "A0HardcreteMk1Wall", "WallTower01",
	                                        "A0RepairCentre3", "Sys-SensoTower02", "A0CommandCentre", "A0ResourceExtractor", "A0FacMod1",
	                                        "A0PowMod1", "A0ResearchModule1", "PillBox1", "A0CyborgFactory", "A0VTolFactory1", "A0VtolPad"};
	std::vector<STRUCTURE_STATS *> avail;
	for (const char *name : buildable)
	{
		int i = getStructStatFromName(WzString::fromUtf8(name));
		if (i >= 0)
		{
			avail.push_back(&asStructureStats[i]);
			for (int p = 0; p < MAX_PLAYERS; p++)
			{
				apStructTypeLists[p][i] = AVAILABLE;
			}
		}
	}
	int wall = getStructStatFromName(WzString::fromUtf8("A0HardcreteMk1Wall"));
	// a skirmish host without a network: the messages go through the game queues (init.cpp's stageTwoInitialise);
	// the clock is gtime's, on a clock that moves 100 ms a frame
	bool savedMultiMessages = bMultiMessages;
	bool savedIsHost = NetPlay.isHost;
	unsigned savedSelected = selectedPlayer;
	bMultiMessages = bMultiPlayer;                        // as a game sets it: SPinit (the campaign) false, a skirmish host true
	NetPlay.isHost = true;
	realSelectedPlayer = selectedPlayer;
	// the selected player the only human (the others AIs, which the host is responsible for)
	bool savedAllocated[MAX_CONNECTED_PLAYERS];
	for (int i = 0; i < MAX_CONNECTED_PLAYERS; i++)
	{
		savedAllocated[i] = NetPlay.players[i].allocated;
		NetPlay.players[i].allocated = i == (int)selectedPlayer;
	}
	if (netTicks == 0)
	{
		netTicks = 1000;
	}
	StatsrefReplayOptionsHandler handler;
	if (mode == NET_MODE_PLAY)
	{
		// a replay viewer (game.cpp's load of a .wzrp: NETloadReplay, then the game's start: gameTimeInit), on one tick
		// a frame (main.cpp: no fast forward for replays outside headless and autogame runs)
		NetPlay.isHost = false;
		if (!NETloadReplay(netReplayFile, handler))
		{
			put("replay not loaded\n");
		}
		setMaxFastForwardTicks(0, true);
	}
	gameTimeInit();
	// after the earlier blocks' times (the droids' lastFired)
	setGameTime(40000);
	if (mode != NET_MODE_PLAY)
	{
		sendQueuedDroidInfo();
		for (int i = 0; i < MAX_CONNECTED_PLAYERS; ++i)
		{
			NETinitQueue(NETgameQueue(i));
			if (!myResponsibility(i) || !NetPlay.bComms)
			{
				NETsetNoSendOverNetwork(NETgameQueue(i));
			}
		}
	}
	if (mode == NET_MODE_RECORD)
	{
		// the lobby's options the replay saves, as loadConfig sets them without a config file (statsref has none)
		sstrcpy(game.name, "My Game");
		game.techLevel = 1;
		game.power = LEV_MED;
		game.base = CAMP_BASE;
		game.inactivityMinutes = 5;
		game.gameTimeLimitMinutes = 0;
		game.playerLeaveMode = PLAYER_LEAVE_MODE::SPLIT_WITH_TEAM;
		// the game's seed, as at a game's start (the replay restores it)
		gameSRand(4242);
		netReplayFile = NETreplaySaveStart("skirmish", handler, 0, false);
		if (netReplayFile.empty())
		{
			put("replay not started\n");
		}
	}
	countUpdate(false);
	gridReset();
	putCaptured();
	if (mode != NET_MODE_PLAY)
	{
		netGiveWork(st, 0);
		netOrders(st, avail, wall, wx, wy);
	}
	for (int frame = 1; frame <= nFrames; frame++)
	{
		netTicks += GAME_TICKS_PER_UPDATE;
		realTimeUpdate();
		gameLoop();
		putCaptured();
		if (frame == 2 || frame == nFrames || syncDumpAll)
		{
			size_t len = dumpLocalDebugSyncLog((syncDebugNext + STATSREF_MAX_SYNC_HISTORY - 1) % STATSREF_MAX_SYNC_HISTORY);
			out.append((const char *)debugSyncTmpBuf.data(), len);
		}
		int nProj = 0;
		for (PROJECTILE *p = proj_GetFirst(); p != nullptr; p = proj_GetNext())
		{
			nProj++;
		}
		put("frame %d: gameTime=%u proj=%d droids=%d %s", frame, gameTime, nProj, (int)allDroids().size(), netLastLogHeader().c_str());
		if (frame % workEvery == 0)
		{
			for (int p = 0; p < MAX_PLAYERS; p++)
			{
				for (STRUCTURE *s : apsStructLists[p])
				{
					if (s->pFunctionality)
					{
						put("%s", baseStructLine(s).c_str());
					}
				}
			}
			if (frame < nFrames && mode != NET_MODE_PLAY)
			{
				netGiveWork(st, frame / workEvery);
				netOrders(st, avail, wall, wx, wy);
				if (frame == 2 * workEvery)
				{
					netDebugMessages(st);
				}
			}
		}
	}
	if (mode == NET_MODE_RECORD)
	{
		if (!NETreplaySaveStop(handler))
		{
			put("replay not saved\n");
		}
		putCaptured();
		put("%s", replaySummary(netReplayFile).c_str());
	}
	else if (mode == NET_MODE_PLAY)
	{
		NETshutdownReplay();
		NetPlay.players.resize(MAX_CONNECTED_PLAYERS);
		selectedPlayer = savedSelected;
		setMaxFastForwardTicks(nullopt, true);
	}
	std::string s = "count:";
	for (int p = 0; p < MAX_PLAYERS; p++)
	{
		int n = 0;
		for (size_t i = 0; i < asResearch.size(); i++)
		{
			if (IsResearchCompleted(&asPlayerResList[p][i]))
			{
				n++;
			}
		}
		s += fmt(" %d/%d/%d/%d", (int)apsDroidLists[p].size(), (int)apsStructLists[p].size(), n, (int)getPower(p));
	}
	put("%s\n", s.c_str());
	put("features %d\n", (int)apsFeatureList[0].size());
	for (int p = 0; p < MAX_PLAYERS; p++)
	{
		setMaxDroids(p, limits[p][0]);
		setMaxCommanders(p, limits[p][1]);
		setMaxConstructors(p, limits[p][2]);
	}
	baseSnapshot(false);
	sendQueuedDroidInfo();
	NETdeleteQueue();
	bMultiMessages = savedMultiMessages;
	NetPlay.isHost = savedIsHost;
	realSelectedPlayer = selectedPlayer;
	for (int i = 0; i < MAX_CONNECTED_PLAYERS; i++)
	{
		NetPlay.players[i].allocated = savedAllocated[i];
	}
	proj_FreeAllProjectiles();
	gridShutDown();
	fpathShutdown();
	setGameTime(0);
	putCaptured();
}

static void mapNet()
{
	netRun(NET_MODE_NET);
}

// a replay read as the game reads one (NETreplayLoadStart with the tests' handler, then every message): the options as
// restored (saved again by the game's own handler), the messages' count by player and type, a CRC of their bytes, and
// the first ten in hex
static void replayInfo(std::string const &filename)
{
	unsigned savedSelected = selectedPlayer;
	StatsrefReplayOptionsHandler handler;
	uint32_t ver = 0;
	if (!NETreplayLoadStart(filename, handler, ver))
	{
		putCaptured();
		put("replay not loaded\n");
		return;
	}
	putCaptured();
	put("format %d, selectedPlayer %d, players %d\n", (int)ver, (int)selectedPlayer, (int)NetPlay.players.size());
	WZGameReplayOptionsHandler gameHandler;
	nlohmann::json j = nlohmann::json::object();
	static_cast<ReplayOptionsHandler const &>(gameHandler).saveOptions(j);
	put("options: %s\n", j.dump().c_str());
	static int counts[256][MAX_GAMEQUEUE_SLOTS];
	memset(counts, 0, sizeof(counts));
	uint32_t crc = wz::crc_init();
	int n = 0;
	std::unique_ptr<NetMessage> msg;
	uint8_t player = 0;
	while (NETreplayLoadNetMessage(msg, player))
	{
		auto const &data = msg->rawData();
		if (player < MAX_GAMEQUEUE_SLOTS)
		{
			counts[msg->type()][player]++;
		}
		crc = wz::crc_update(crc, data.data(), data.size());
		if (n < 10)
		{
			std::string s = fmt("message %d: p%d", n, (int)player);
			for (size_t i = 0; i < data.size(); i++)
			{
				s += fmt(" %02X", data[i]);
			}
			put("%s\n", s.c_str());
		}
		n++;
		if (msg->type() == REPLAY_ENDED)
		{
			break;
		}
	}
	NETreplayLoadStop();
	putCaptured();
	std::string s = fmt("messages %d, crc %08X:", n, crc);
	for (int p = 0; p < MAX_GAMEQUEUE_SLOTS; p++)
	{
		for (int i = 0; i < 256; i++)
		{
			if (counts[i][p] > 0)
			{
				s += fmt(" p%d %s=%d", p, messageTypeToString(i), counts[i][p]);
			}
		}
	}
	put("%s\n", s.c_str());
	NETdeleteQueue();
	NetPlay.players.resize(MAX_CONNECTED_PLAYERS);
	selectedPlayer = savedSelected;
	realSelectedPlayer = selectedPlayer;
	putCaptured();
}

static void mapQueries()
{
	uint64_t st = 1 + mapWidth * 7 + mapHeight;
	int wx = world_coord(mapWidth);
	int wy = world_coord(mapHeight);
	std::string s = "heights:";
	for (int i = 1; i <= 100; i++)
	{
		int x = testRandMod(st, wx + TILE_UNITS) - TILE_UNITS / 2;
		int y = testRandMod(st, wy + TILE_UNITS) - TILE_UNITS / 2;
		s += fmt(" %d,%d=%d", x, y, map_Height(x, y));
	}
	putCaptured();
	put("%s\n", s.c_str());
	s = "maxmin:";
	for (int i = 1; i <= 100; i++)
	{
		int x = testRandMod(st, mapWidth + 2) - 1;
		int y = testRandMod(st, mapHeight + 2) - 1;
		int mx, mn;
		getTileMaxMin(x, y, &mx, &mn);
		s += fmt(" %d,%d=%d,%d", x, y, mx, mn);
	}
	putCaptured();
	put("%s\n", s.c_str());
	s = "lines:";
	for (int i = 1; i <= 100; i++)
	{
		int sx = testRandMod(st, wx);
		int sy = testRandMod(st, wy);
		int sz = map_Height(sx, sy);
		sz += testRandMod(st, 400) - 100;
		int dx = testRandMod(st, wx);
		int dy = testRandMod(st, wy);
		int dz = testRandMod(st, 1200) - 200;
		unsigned r = map_LineIntersect(Vector3i(sx, sy, sz), Vector3i(dx, dy, dz), 65536);
		s += fmt(" %d,%d,%d>%d,%d,%d=%u", sx, sy, sz, dx, dy, dz, r);
	}
	putCaptured();
	put("%s\n", s.c_str());
	if (wx > 2 * TILE_UNITS && wy > 2 * TILE_UNITS)
	{
		s = "intersect:";
		for (int i = 1; i <= 100; i++)
		{
			int cx = testRandMod(st, wx - 2 * TILE_UNITS) + TILE_UNITS;
			int cy = testRandMod(st, wy - 2 * TILE_UNITS) + TILE_UNITS;
			int vx = testRandMod(st, 401) - 200;
			int vy = testRandMod(st, 401) - 200;
			int ix = 0, iy = 0;
			s += fmt(" %d,%d,%d,%d", cx, cy, vx, vy);
			bool ok = map_Intersect(&cx, &cy, &vx, &vy, &ix, &iy);
			s += fmt("=%d,%d,%d,%d,%d,%d,%d", (int)ok, cx, cy, vx, vy, ix, iy);
		}
		putCaptured();
		put("%s\n", s.c_str());
	}
	s = "blocking:";
	for (int i = 1; i <= 100; i++)
	{
		int x = testRandMod(st, mapWidth + 2) - 1;
		int y = testRandMod(st, mapHeight + 2) - 1;
		s += fmt(" %d,%d=%d%d%d%d", x, y, (int)fpathBlockingTile(x, y, PROPULSION_TYPE_WHEELED), (int)fpathBlockingTile(x, y, PROPULSION_TYPE_HOVER),
		         (int)fpathBlockingTile(x, y, PROPULSION_TYPE_LIFT), (int)fpathBlockingTile(x, y, PROPULSION_TYPE_PROPELLOR));
	}
	putCaptured();
	put("%s\n", s.c_str());
	s = "fire:";
	for (int i = 1; i <= 20; i++)
	{
		gameTime = testRand(st);
		int x = testRandMod(st, wx);
		int y = testRandMod(st, wy);
		int r = testRandMod(st, 5000);
		tileSetFire(x, y, r);
		MAPTILE *t = worldTile(x, y);
		s += fmt(" %d,%d,%d=%X,%u,%d", x, y, r, (unsigned)t->tileInfoBits, (unsigned)t->fireEndTime, (int)fireOnLocation(x, y));
	}
	gameTime = 0;
	putCaptured();
	put("%s\n", s.c_str());
}

static bool terrainLoaded = false;

static void dump()
{
	put("== weapons %u\n", (unsigned)asWeaponStats.size());
	for (size_t i = 0; i < asWeaponStats.size(); i++)
	{
		const WEAPON_STATS &s = asWeaponStats[i];
		putCommon("weapon", i, s);
		putUpgrades(s, [](const decltype(s.base) &u) {
			return fmt("|hp=%u|hpPct=%d|short=%u|max=%u|min=%u|hit=%u|shortHit=%u|pause=%u|rounds=%u|reload=%u|damage=%u|emp=%u|radius=%u|radDamage=%u|pd=%u|pdRadius=%u|pdTime=%u|minDamage=%u",
			           u.hitpoints, u.hitpointPct, u.shortRange, u.maxRange, u.minRange, u.hitChance, u.shortHitChance, u.firePause, (unsigned)u.numRounds,
			           u.reloadTime, u.damage, u.empRadius, u.radius, u.radiusDamage, u.periodicalDamage, u.periodicalDamageRadius, u.periodicalDamageTime, u.minimumDamage);
		});
		put(" pdClass=%d|pdSub=%d|pdEffect=%d|class=%d|sub=%d|movement=%d|effect=%d|size=%d|flight=%u|recoil=%u|ext=%d|rotate=%d|maxElev=%d|minElev=%d|effectSize=%u|runs=%d|face=%u|faceFlight=%u|air=%u|light=%d|penetrate=%d|fireOnMove=%d|flags=%lu|radiusLife=%u|explosions=%u\n",
		    (int)s.periodicalDamageWeaponClass, (int)s.periodicalDamageWeaponSubClass, (int)s.periodicalDamageWeaponEffect, (int)s.weaponClass, (int)s.weaponSubClass,
		    (int)s.movementModel, (int)s.weaponEffect, (int)s.weaponSize, s.flightSpeed, s.recoilValue, s.distanceExtensionFactor, (int)s.rotate, (int)s.maxElevation,
		    (int)s.minElevation, (unsigned)s.effectSize, (int)s.vtolAttackRuns, (unsigned)s.facePlayer, (unsigned)s.faceInFlight, (unsigned)s.surfaceToAir, (int)s.lightWorld,
		    (int)s.penetrate, (int)s.fireOnMove, (unsigned long)s.flags.to_ulong(), s.radiusLife, s.numExplosions);
		put(" mount=%s|muzzle=%s|flight=%s|hit=%s|miss=%s|water=%s|trail=%s|fire=%d|impact=%d\n",
		    modelStr(s.pMountGraphic).c_str(), modelStr(s.pMuzzleGraphic).c_str(), modelStr(s.pInFlightGraphic).c_str(), modelStr(s.pTargetHitGraphic).c_str(),
		    modelStr(s.pTargetMissGraphic).c_str(), modelStr(s.pWaterHitGraphic).c_str(), modelStr(s.pTrailGraphic).c_str(), s.iAudioFireID, s.iAudioImpactID);
	}
	put("== bodies %u\n", (unsigned)asBodyStats.size());
	for (size_t i = 0; i < asBodyStats.size(); i++)
	{
		const BODY_STATS &s = asBodyStats[i];
		putCommon("body", i, s);
		putUpgrades(s, [](const BODY_STATS::UPGRADE &u) {
			return fmt("|hp=%u|hpPct=%d|power=%u|armour=%u|thermal=%d|resistance=%d", u.hitpoints, u.hitpointPct, u.power, u.armour, u.thermal, u.resistance);
		});
		put(" size=%d|slots=%u|class=%s\n", (int)s.size, s.weaponSlots, s.bodyClass.toUtf8().c_str());
		putList("imd", s.ppIMDList);
		putList("move", s.ppMoveIMDList);
		putList("still", s.ppStillIMDList);
	}
	put("== brains %u\n", (unsigned)asBrainStats.size());
	for (size_t i = 0; i < asBrainStats.size(); i++)
	{
		const BRAIN_STATS &s = asBrainStats[i];
		putCommon("brain", i, s);
		putUpgrades(s, [](const decltype(s.base) &u) {
			std::string t = fmt("|hp=%u|hpPct=%d|maxDroids=%d|maxDroidsMult=%d|thresholds", u.hitpoints, u.hitpointPct, u.maxDroids, u.maxDroidsMult);
			for (int v : u.rankThresholds)
			{
				t += fmt(" %d", v);
			}
			return t;
		});
		put(" weapon=%s|ranks", indexOf(s.psWeaponStat, asWeaponStats).c_str());
		for (const auto &r : s.rankNames)
		{
			put(" %s", r.c_str());
		}
		put("|cmdExpRange");
		for (int v : s.cmdExpRange)
		{
			put(" %d", v);
		}
		put("\n");
	}
	put("== propulsion %u\n", (unsigned)asPropulsionStats.size());
	for (size_t i = 0; i < asPropulsionStats.size(); i++)
	{
		const PROPULSION_STATS &s = asPropulsionStats[i];
		putCommon("propulsion", i, s);
		putUpgrades(s, [](const PROPULSION_STATS::UPGRADE &u) {
			return fmt("|hp=%u|hpPct=%d|pctOfBody=%d", u.hitpoints, u.hitpointPct, u.hitpointPctOfBody);
		});
		put(" type=%d|speed=%u|turn=%u|spin=%u|spinAngle=%u|skid=%u|decel=%u|accel=%u\n", (int)s.propulsionType, s.maxSpeed, s.turnSpeed, s.spinSpeed,
		    s.spinAngle, s.skidDeceleration, s.deceleration, s.acceleration);
	}
	put("== sensors %u\n", (unsigned)asSensorStats.size());
	for (size_t i = 0; i < asSensorStats.size(); i++)
	{
		const SENSOR_STATS &s = asSensorStats[i];
		putCommon("sensor", i, s);
		putUpgrades(s, [](const decltype(s.base) &u) { return fmt("|hp=%u|hpPct=%d|range=%u", u.hitpoints, u.hitpointPct, u.range); });
		put(" mount=%s|location=%u|type=%d\n", modelStr(s.pMountGraphic).c_str(), s.location, (int)s.type);
	}
	put("== ecm %u\n", (unsigned)asECMStats.size());
	for (size_t i = 0; i < asECMStats.size(); i++)
	{
		const ECM_STATS &s = asECMStats[i];
		putCommon("ecm", i, s);
		putUpgrades(s, [](const decltype(s.base) &u) { return fmt("|hp=%u|hpPct=%d|range=%u", u.hitpoints, u.hitpointPct, u.range); });
		put(" mount=%s|location=%u\n", modelStr(s.pMountGraphic).c_str(), s.location);
	}
	put("== repair %u\n", (unsigned)asRepairStats.size());
	for (size_t i = 0; i < asRepairStats.size(); i++)
	{
		const REPAIR_STATS &s = asRepairStats[i];
		putCommon("repair", i, s);
		putUpgrades(s, [](const decltype(s.base) &u) { return fmt("|hp=%u|hpPct=%d|points=%u", u.hitpoints, u.hitpointPct, u.repairPoints); });
		put(" mount=%s|location=%u|time=%u\n", modelStr(s.pMountGraphic).c_str(), s.location, s.time);
	}
	put("== construct %u\n", (unsigned)asConstructStats.size());
	for (size_t i = 0; i < asConstructStats.size(); i++)
	{
		const CONSTRUCT_STATS &s = asConstructStats[i];
		putCommon("construct", i, s);
		putUpgrades(s, [](const decltype(s.base) &u) { return fmt("|hp=%u|hpPct=%d|points=%u", u.hitpoints, u.hitpointPct, u.constructPoints); });
		put(" mount=%s\n", modelStr(s.pMountGraphic).c_str());
	}
	put("== propulsion types %u\n", (unsigned)asPropulsionTypes.size());
	for (size_t i = 0; i < asPropulsionTypes.size(); i++)
	{
		const PROPULSION_TYPES &t = asPropulsionTypes[i];
		put("proptype %u|travel=%d|mult=%u|start=%d|idle=%d|moveOff=%d|move=%d|hiss=%d|shutDown=%d\n", (unsigned)i, (int)t.travel, (unsigned)t.powerRatioMult,
		    (int)t.startID, (int)t.idleID, (int)t.moveOffID, (int)t.moveID, (int)t.hissID, (int)t.shutDownID);
	}
	if (terrainLoaded)
	{
		put("== terrain\n");
		for (int t = 0; t < TER_MAX; t++)
		{
			put("terrain %d", t);
			for (int p = 0; p < PROPULSION_TYPE_NUM; p++)
			{
				put(" %u", getSpeedFactor(t, p));
			}
			put("\n");
		}
	}
	put("== modifiers\n");
	for (int e = 0; e < WE_NUMEFFECTS; e++)
	{
		put("effect %d|prop", e);
		for (int p = 0; p < PROPULSION_TYPE_NUM; p++)
		{
			put(" %u", (unsigned)asWeaponModifier[e][p]);
		}
		put("|body");
		for (int b = 0; b < SIZE_NUM; b++)
		{
			put(" %u", (unsigned)asWeaponModifierBody[e][b]);
		}
		put("|struct");
		for (int b = 0; b < NUM_STRUCT_STRENGTH; b++)
		{
			put(" %u", (unsigned)asStructStrengthModifier[e][b]);
		}
		put("\n");
	}
	put("== component lists\n");
	for (int t = 0; t < COMP_NUMCOMPONENTS; t++)
	{
		if (!apCompLists[0][t])
		{
			continue;
		}
		size_t n = 0;
		switch (t)
		{
		case COMP_BODY: n = asBodyStats.size(); break;
		case COMP_BRAIN: n = asBrainStats.size(); break;
		case COMP_PROPULSION: n = asPropulsionStats.size(); break;
		case COMP_REPAIRUNIT: n = asRepairStats.size(); break;
		case COMP_ECM: n = asECMStats.size(); break;
		case COMP_SENSOR: n = asSensorStats.size(); break;
		case COMP_CONSTRUCT: n = asConstructStats.size(); break;
		case COMP_WEAPON: n = asWeaponStats.size(); break;
		}
		for (int p = 0; p < MAX_PLAYERS; p++)
		{
			put("comp %d player %d:", t, p);
			for (size_t c = 0; c < n; c++)
			{
				put(" %u", (unsigned)apCompLists[p][t][c]);
			}
			put("\n");
		}
	}
	if (apStructTypeLists[0])
	{
		for (int p = 0; p < MAX_PLAYERS; p++)
		{
			put("structlist player %d:", p);
			for (unsigned c = 0; c < numStructureStats; c++)
			{
				put(" %u", (unsigned)apStructTypeLists[p][c]);
			}
			put("\n");
		}
	}
	put("== structures %u|factoryModule=%u|powerModule=%u|researchModule=%u|baseExp=%d\n", numStructureStats, factoryModuleStat, powerModuleStat,
	    researchModuleStat, getStructureDamageBaseExperienceLevel());
	for (unsigned i = 0; i < numStructureStats; i++)
	{
		const STRUCTURE_STATS &s = asStructureStats[i];
		put("structure %u|%s|%s|ref=%x|index=%u|type=%d|strength=%d|width=%u|breadth=%u|points=%u|height=%u|power=%u|imds", i, s.id.toUtf8().c_str(),
		    s.name.toUtf8().c_str(), s.ref, (unsigned)s.index, (int)s.type, (int)s.strength, s.baseWidth, s.baseBreadth, s.buildPoints, s.height, s.powerToBuild);
		for (const auto *m : s.pIMD)
		{
			put(" %s", modelStr(m).c_str());
		}
		put("|base=%s|ecm=%s|sensor=%s|weapons=%u", modelStr(s.pBaseIMD).c_str(), indexOf(s.pECM, asECMStats).c_str(), indexOf(s.pSensor, asSensorStats).c_str(),
		    s.numWeaps);
		for (int w = 0; w < MAX_WEAPONS; w++)
		{
			put(" %s", indexOf(s.psWeapStat[w], asWeaponStats).c_str());
		}
		put("|flags=%u|walls=%d|minLimit=%u|maxLimit=%u|favorite=%d\n", (unsigned)s.flags, (int)s.combinesWithWall, s.minLimit, s.maxLimit, (int)s.isFavorite);
		putUpgrades(s, [](const decltype(s.base) &u) {
			return fmt("|research=%u|moduleResearch=%u|repair=%u|power=%u|modulePower=%u|production=%u|moduleProduction=%u|rearm=%u|armour=%u|thermal=%u|hp=%u|resistance=%u|limit=%u",
			           u.research, u.moduleResearch, u.repair, u.power, u.modulePower, u.production, u.moduleProduction, u.rearm, u.armour, u.thermal, u.hitpoints,
			           u.resistance, u.limit);
		});
	}
	put("== features %u|oil=%s\n", (unsigned)asFeatureStats.size(), indexOf(oilResFeature, asFeatureStats).c_str());
	for (size_t i = 0; i < asFeatureStats.size(); i++)
	{
		const FEATURE_STATS &s = asFeatureStats[i];
		put("feature %u|%s|%s|ref=%x|index=%u|type=%d|imd=%s|width=%u|breadth=%u|tile=%d|los=%d|visible=%d|damageable=%d|body=%u|armour=%u\n", (unsigned)i,
		    s.id.toUtf8().c_str(), s.name.toUtf8().c_str(), s.ref, (unsigned)s.index, (int)s.subType, modelStr(s.psImd).c_str(), (unsigned)s.baseWidth,
		    (unsigned)s.baseBreadth, (int)s.tileDraw, (int)s.allowLOS, (int)s.visibleAtStart, (int)s.damageable, s.body, s.armourValue);
	}
	put("== research %u|mode=%d\n", (unsigned)asResearch.size(), (int)getResearchUpgradeCalcMode());
	for (size_t i = 0; i < asResearch.size(); i++)
	{
		const RESEARCH &r = asResearch[i];
		put("research %u|%s|%s|ref=%x|baseIndex=%u|index=%d|tech=%u|subGroup=%u|points=%u|power=%u|key=%u|noCheat=%u|disabled=%u|icon=%u|stat=%s|imd=%s|imd2=%s|view=%s|category=%s",
		    (unsigned)i, r.id.toUtf8().c_str(), r.name.toUtf8().c_str(), r.ref, (unsigned)static_cast<const BASE_STATS &>(r).index, r.index, (unsigned)r.techCode,
		    (unsigned)r.subGroup, (unsigned)r.researchPoints, r.researchPower, (unsigned)r.keyTopic, (unsigned)r.excludeFromCheats, (unsigned)r.disabledWhen,
		    (unsigned)r.iconID, statRef(r.psStat).c_str(), modelStr(r.pIMD).c_str(), modelStr(r.pIMD2).c_str(), r.pViewData ? "set" : "null", r.category.toUtf8().c_str());
		if (!r.category.isEmpty())
		{
			put(" %u/%u", (unsigned)r.categoryProgress, (unsigned)r.categoryMax);
		}
		put("\n pre");
		for (auto v : r.pPRList)
		{
			put(" %u", (unsigned)v);
		}
		put("|reqStruct");
		for (auto v : r.pStructList)
		{
			put(" %u", (unsigned)v);
		}
		put("|redStruct");
		for (auto v : r.pRedStructs)
		{
			put(" %u", (unsigned)v);
		}
		put("|resStruct");
		for (auto v : r.pStructureResults)
		{
			put(" %u", (unsigned)v);
		}
		put("\n redComp");
		for (auto *c : r.pRedArtefacts)
		{
			put(" %s", statRef(c).c_str());
		}
		put("|resComp");
		for (auto *c : r.componentResults)
		{
			put(" %s", statRef(c).c_str());
		}
		put("|replace");
		for (const auto &c : r.componentReplacement)
		{
			put(" %s>%s", statRef(c.pOldComponent).c_str(), statRef(c.pNewComponent).c_str());
		}
		put("\n results %s\n", r.results.dump().c_str());
	}
	for (int p = 0; p < MAX_PLAYERS; p++)
	{
		put("playerres %d %u\n", p, (unsigned)asPlayerResList[p].size());
	}
	auto putTemplate = [](const DROID_TEMPLATE &t) {
		put("template %s|%s|ref=%x|mpid=%u|type=%d|parts", t.id.toUtf8().c_str(), t.name.toUtf8().c_str(), t.ref, t.multiPlayerID, (int)t.droidType);
		for (int k = 0; k < DROID_MAXCOMP; k++)
		{
			put(" %u", (unsigned)t.asParts[k]);
		}
		put("|weapons %d", (int)t.numWeaps);
		for (int k = 0; k < MAX_WEAPONS; k++)
		{
			put(" %u", t.asWeaps[k]);
		}
		put("|prefab=%d|stored=%d|enabled=%d\n", (int)t.prefab, (int)t.stored, (int)t.enabled);
	};
	for (int p = 0; p < MAX_PLAYERS; p++)
	{
		put("== templates player %d: %u\n", p, (unsigned)templateCount(p));
		enumerateTemplates(p, [&](DROID_TEMPLATE *t) {
			putTemplate(*t);
			return true;
		});
	}
	put("== local templates %u\n", (unsigned)localTemplates.size());
	for (const auto &t : localTemplates)
	{
		putTemplate(t);
	}
	put("== object ids %u\n", unsynchObjID);
}

// ---------------------------------------------------------------------------

static std::vector<std::string> split(const std::string &s, char sep)
{
	std::vector<std::string> v;
	std::string cur;
	for (char c : s)
	{
		if (c == sep)
		{
			v.push_back(cur);
			cur.clear();
		}
		else
		{
			cur += c;
		}
	}
	v.push_back(cur);
	return v;
}

static void flush(const std::string &file)
{
	if (file.empty())
	{
		return;
	}
	std::ofstream f(file, std::ios::binary);
	f << out;
}

static bool loadKind(const std::string &kind, const std::string &file)
{
	// data.cpp's buffer loaders (without the data hash: bMultiPlayer is false)
	if (kind == "propsounds")
	{
		return loadPropulsionSounds(file.c_str());
	}
	if (kind == "templates")
	{
		return loadDroidTemplates(file.c_str());
	}
	WzConfig ini(WzString::fromUtf8(file), WzConfig::ReadOnlyAndRequired);
	if (kind == "weapons") return loadWeaponStats(ini) && allocComponentList(COMP_WEAPON, asWeaponStats.size());
	if (kind == "body") return loadBodyStats(ini) && allocComponentList(COMP_BODY, asBodyStats.size());
	if (kind == "brain") return loadBrainStats(ini) && allocComponentList(COMP_BRAIN, asBrainStats.size());
	if (kind == "propulsion") return loadPropulsionStats(ini) && allocComponentList(COMP_PROPULSION, asPropulsionStats.size());
	if (kind == "sensor") return loadSensorStats(ini) && allocComponentList(COMP_SENSOR, asSensorStats.size());
	if (kind == "ecm") return loadECMStats(ini) && allocComponentList(COMP_ECM, asECMStats.size());
	if (kind == "repair") return loadRepairStats(ini) && allocComponentList(COMP_REPAIRUNIT, asRepairStats.size());
	if (kind == "construct") return loadConstructStats(ini) && allocComponentList(COMP_CONSTRUCT, asConstructStats.size());
	if (kind == "proptypes") return loadPropulsionTypes(ini);
	if (kind == "terrain")
	{
		terrainLoaded = true;
		return loadTerrainTable(ini);
	}
	if (kind == "weapmods") return loadWeaponModifiers(ini);
	if (kind == "structures") return loadStructureStats(ini) && allocStructLists();
	if (kind == "structmods") return loadStructureStrengthModifiers(ini);
	if (kind == "features") return loadFeatureStats(ini);
	if (kind == "research")
	{
		if (!asResearch.empty())
		{
			ResearchShutDown();
		}
		return loadResearch(ini);
	}
	put("unknown kind %s\n", kind.c_str());
	return false;
}

int main(int argc, char **argv)
{
	if (argc < 2)
	{
		fprintf(stderr, "usage: statsref <command file>\n");
		return 2;
	}
	debug_init();
	trigInitialise();                                     // frameInitialise's
	debug_register_callback(captureCallback, nullptr, nullptr, nullptr);

	std::ifstream cmds(argv[1], std::ios::binary);
	std::string line, file;
	while (std::getline(cmds, line))
	{
		if (!line.empty() && line.back() == '\r')
		{
			line.pop_back();
		}
		if (line.empty())
		{
			continue;
		}
		std::vector<std::string> w = split(line, '\t');
		const std::string &cmd = w[0];
		if (cmd == "searchpath")
		{
			statsref_set_searchpath(w.size() > 1 ? w[1] : "");
			continue;
		}
		if (cmd == "output")
		{
			flush(file);
			out.clear();
			captured.clear();
			file = w[1];
			continue;
		}
		put("> %s\n", line.c_str());
		if (cmd == "debug")
		{
			debug_enable_switch(w[1].c_str());
		}
		else if (cmd == "selected")
		{
			selectedPlayer = std::stoi(w[1]);
		}
		else if (cmd == "players")
		{
			unsigned mask = std::stoul(w[1]);
			for (int i = 0; i < MAX_PLAYERS; i++)
			{
				NetPlay.players[i].allocated = (mask >> i) & 1;
			}
		}
		else if (cmd == "gamemode")
		{
			gameMode = (GS_GAMEMODE)std::stoi(w[1]);
		}
		else if (cmd == "reset")
		{
			freeComponentLists();
			statsShutDown();
			freeStructureLists();
			structureStatsShutDown();
			featureStatsShutDown();
			featureInitVars();
			ResearchShutDown();
			researchInitVars();
			droidTemplateShutDown();
			objmemInitialise();
			terrainLoaded = false;
		}
		else if (cmd == "strings")
		{
			// data.cpp's dataStrResLoad
			bool ok = true;
			if (psStringRes == nullptr)
			{
				ok = stringsInitialise();
			}
			if (ok)
			{
				ok = strresLoad(psStringRes, w[1].c_str());
			}
			put("%s", captured.c_str());
			captured.clear();
			put("result=%d\n", (int)ok);
			continue;
		}
		else if (cmd == "releasestrings")
		{
			// data.cpp's dataStrResRelease
			if (psStringRes != nullptr)
			{
				strresDestroy(psStringRes);
				psStringRes = nullptr;
			}
		}
		else if (cmd == "getstring")
		{
			const char *s = psStringRes ? strresGetString(psStringRes, w[1].c_str()) : nullptr;
			put("%s", captured.c_str());
			captured.clear();
			put("string=%s\n", s ? s : "(null)");
			continue;
		}
		else if (cmd == "load")
		{
			bool ok = loadKind(w[1], w[2]);
			put("%s", captured.c_str());
			captured.clear();
			put("result=%d\n", (int)ok);
			continue;
		}
		else if (cmd == "dump")
		{
			put("%s", captured.c_str());
			captured.clear();
			dump();
			continue;
		}
		else if (cmd == "model")
		{
			// imdload.cpp's modelGet, and the model
			iIMDBaseShape *m = modelGet(WzString::fromUtf8(w[1]));
			put("%s", captured.c_str());
			captured.clear();
			put("errors=%u\n", (unsigned)getModelLoadingErrorCount());
			if (m == nullptr)
			{
				put("model=null\n");
			}
			else
			{
				modelDump(m);
			}
			continue;
		}
		else if (cmd == "playerfaction")
		{
			NetPlay.players[std::stoi(w[1])].faction = (FactionID)std::stoi(w[2]);
		}
		else if (cmd == "factions")
		{
			// getEnabledFactions (in id order)
			auto enabled = getEnabledFactions(std::stoi(w[1]) != 0);
			std::vector<int> ids;
			for (FactionID f : enabled)
			{
				ids.push_back((int)f);
			}
			std::sort(ids.begin(), ids.end());
			std::string s = "factions:";
			for (int id : ids)
			{
				s += " " + std::to_string(id);
			}
			putCaptured();
			put("%s\n", s.c_str());
			continue;
		}
		else if (cmd == "factionmodel")
		{
			auto m = getFactionModelName((FactionID)std::stoi(w[1]), WzString::fromUtf8(w[2]));
			put("model=%s\n", m.has_value() ? m.value().toUtf8().c_str() : "nullopt");
		}
		else if (cmd == "overrides")
		{
			useTerrainOverrides = std::stoi(w[1]) != 0;
		}
		else if (cmd == "mapload")
		{
			bool ok = statsrefMapLoad(w[1], w[2], std::stoi(w[3]), w[4]);
			putCaptured();
			put("result=%d\n", (int)ok);
			if (psMapTiles)
			{
				mapDump();
			}
			continue;
		}
		else if (cmd == "mapreloadground")
		{
			bool ok = mapReloadGroundTypes();
			putCaptured();
			put("result=%d\n", (int)ok);
			if (psMapTiles)
			{
				mapDump();
			}
			continue;
		}
		else if (cmd == "mapqueries")
		{
			putCaptured();
			if (psMapTiles)
			{
				mapQueries();
			}
			else
			{
				put("no map\n");
			}
			continue;
		}
		else if (cmd == "mapshutdown")
		{
			gwShutDown();
			mapShutdown();
		}
		else if (cmd == "mapplayers")
		{
			// a skirmish or campaign game of n players
			game.type = w[1] == "campaign" ? LEVEL_TYPE::CAMPAIGN : LEVEL_TYPE::SKIRMISH;
			game.maxPlayers = std::stoi(w[2]);
			game.scavengers = NO_SCAVENGERS;
			game.alliance = NO_ALLIANCES;
			bMultiPlayer = w[1] != "campaign";
			for (int i = 0; i < MAX_PLAYERS; i++)
			{
				NetPlay.players[i].position = i < std::stoi(w[2]) ? i : -1;
			}
		}
		else if (cmd == "mapobjects")
		{
			// loadGame's object set-up, then loadWzMapFeature and loadWzMapStructure on the map mapload loaded, as loadGame orders them
			gameSRand(12345);
			objInitialise();
			grpInitialise();
			aiInitialise();
			scriptInit();
			visInitialise();
			initFactoryNumFlag();
			droidInit();
			putCaptured();
			if (!testMap || !psMapTiles)
			{
				put("no map\n");
			}
			else
			{
				std::unordered_map<UDWORD, UDWORD> fixedMapIdToGeneratedId;
				bool ok = loadWzMapDroidInit(*testMap, fixedMapIdToGeneratedId);
				putCaptured();
				put("droids result=%d\n", (int)ok);
				ok = loadWzMapFeature(*testMap, fixedMapIdToGeneratedId);
				putCaptured();
				put("features result=%d\n", (int)ok);
				initStructLimits();
				if (game.type != LEVEL_TYPE::CAMPAIGN)
				{
					freeAllFlagPositions();
				}
				std::array<std::unordered_map<UDWORD, UDWORD>, MAX_PLAYER_SLOTS> moduleToBuilding;
				ok = loadWzMapStructure(*testMap, fixedMapIdToGeneratedId, moduleToBuilding);
				if (game.type != LEVEL_TYPE::CAMPAIGN)
				{
					resetFactoryNumFlag();
				}
				putCaptured();
				put("structures result=%d\n", (int)ok);
				objectsDump();
			}
			continue;
		}
		else if (cmd == "syncreset")
		{
			resetSyncDebug();
		}
		else if (cmd == "syncdumpall")
		{
			syncDumpAll = w.size() > 1 && w[1] == "1";
		}
		else if (cmd == "synclog")
		{
			// the sync log since the last synclog, then nextDebugSync
			putCaptured();
			size_t n = dumpLocalDebugSyncLog(syncDebugNext);
			out.append((const char *)debugSyncTmpBuf.data(), n);
			unsigned crc = nextDebugSync();
			putCaptured();
			put("crc=%04X\n", crc);
			continue;
		}
		else if (cmd == "mapgrid")
		{
			putCaptured();
			if (psMapTiles)
			{
				mapGrid();
			}
			else
			{
				put("no map\n");
			}
			continue;
		}
		else if (cmd == "mappaths")
		{
			putCaptured();
			if (psMapTiles)
			{
				mapPaths();
			}
			else
			{
				put("no map\n");
			}
			continue;
		}
		else if (cmd == "mapmove")
		{
			putCaptured();
			if (psMapTiles)
			{
				mapMove();
			}
			else
			{
				put("no map\n");
			}
			continue;
		}
		else if (cmd == "maporders")
		{
			putCaptured();
			if (psMapTiles)
			{
				mapOrders();
			}
			else
			{
				put("no map\n");
			}
			continue;
		}
		else if (cmd == "mapcombat")
		{
			putCaptured();
			if (psMapTiles)
			{
				mapCombat();
			}
			else
			{
				put("no map\n");
			}
			continue;
		}
		else if (cmd == "mapbase")
		{
			putCaptured();
			if (psMapTiles)
			{
				mapBase();
			}
			else
			{
				put("no map\n");
			}
			continue;
		}
		else if (cmd == "mapnet")
		{
			putCaptured();
			if (psMapTiles)
			{
				mapNet();
			}
			else
			{
				put("no map\n");
			}
			continue;
		}
		else if (cmd == "replayinfo")
		{
			putCaptured();
			replayInfo(w[1]);
			continue;
		}
		else if (cmd == "maprecord" || cmd == "mapplay")
		{
			putCaptured();
			if (psMapTiles)
			{
				netRun(cmd == "maprecord" ? NET_MODE_RECORD : NET_MODE_PLAY);
			}
			else
			{
				put("no map\n");
			}
			continue;
		}
		else if (cmd == "writedir")
		{
			statsref_set_writedir(w[1]);
			continue;
		}
		else if (cmd == "mappower")
		{
			putCaptured();
			mapPower();
			continue;
		}
		else if (cmd == "objshutdown")
		{
			// the messages first (their proximity displays name the objects), as a level's start frees them
			freeMessages();
			freeAllDroids();
			freeAllStructs();
			freeAllFeatures();
			apsOilList[0].clear();
			apsSensorList[0].clear();
			for (int i = 0; i < MAX_PLAYERS; i++)
			{
				apsExtractorLists[i].clear();
			}
			freeAllFlagPositions();
			psDestroyedObj.clear();
			objmemShutdown();
		}
		else if (cmd == "reload")
		{
			// debugReloadDisplayModelsForBaseModel of a loaded model, and the model
			iIMDBaseShape *m = modelGet(WzString::fromUtf8(w[1]));
			bool ok = false;
			if (m != nullptr)
			{
				ok = debugReloadDisplayModelsForBaseModel(*m);
			}
			put("%s", captured.c_str());
			captured.clear();
			if (m == nullptr)
			{
				put("model=null\n");
			}
			else
			{
				put("result=%d\n", (int)ok);
				modelDump(m);
			}
			continue;
		}
		else
		{
			put("unknown command\n");
		}
		put("%s", captured.c_str());
		captured.clear();
	}
	put("%s", captured.c_str());
	flush(file);
	return 0;
}
