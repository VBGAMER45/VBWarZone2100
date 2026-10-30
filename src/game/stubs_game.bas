Attribute VB_Name = "stubs_game"
' Stand-ins for the game (src, lib/gamelib) until their clusters port them.
' Each keeps the C++ name and signature, so its callers need no change.
' The constants, globals and settings of display3d, display, keybind, radar,
' multiplay(defs), multiint, difficulty, texture, terrain, seqdisp,
' ingameop, loadsave, hci/groups and nethelpers that warzoneconfig.cpp and
' configuration.cpp need (with their modules' clusters): each setter keeps
' the C++ checks on its value, but nothing else happens (no damage
' modifiers, shadows, terrain reload or log lines).
' component.cpp's setPlayerColour, activity.cpp's ActivityManager,
' seqdisp.cpp's scanline mode (later clusters). gtime.cpp (cluster B),
' keybind.cpp's kf_ScreenDump (cluster D), game.cpp (cluster I).
' What clparse.cpp sets or calls in frontend, keybind, loadsave, wrappers,
' stdinreader, multilobbycommands and gamehistorylogger: the values are
' kept, and nothing else happens yet. lib/exceptionhandler's addDumpInfo
' (the crash handler is not ported).
Option Explicit

' (wzGetLogicalCPUCount) GetSystemInfo's SYSTEM_INFO
Private Type SYSTEM_INFO
    wProcessorArchitecture As Integer
    wReserved As Integer
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    wProcessorLevel As Integer
    wProcessorRevision As Integer
End Type
Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)

' frontend.h
Public Const MAX_LEVEL_NAME_SIZE As Long = 256

' stdinreader.h: enum class WZ_Command_Interface
Public Const WZ_Command_Interface_None As Long = 0
Public Const WZ_Command_Interface_StdIn_Interface As Long = 1
Public Const WZ_Command_Interface_Unix_Socket As Long = 2

' gamehistorylogger.h: GameStoryLogger's OutputKey, OutputNaming, OutputModes
Public Const GameStoryLogger_OutputKey_PlayerIndex As Long = 0
Public Const GameStoryLogger_OutputKey_PlayerPosition As Long = 1
Public Const GameStoryLogger_OutputNaming_Default As Long = 0
Public Const GameStoryLogger_OutputNaming_AutohosterClassic As Long = 1

Public Type GameStoryLogger_OutputModes
    cmdInterface As Boolean
    logFile As Boolean
End Type

' display3d.cpp, frontend.cpp, multiplay.cpp, loadsave.cpp
Public aLevelName As String             ' char[MAX_LEVEL_NAME_SIZE + 1]: vital! the wrf file to use.
Public bMultiMessages As Boolean        ' == bMultiPlayer unless multimessages are disabled

' stdinreader.cpp, multilobbycommands.cpp, gamehistorylogger.cpp
Private cmdInterfaceMode As Long
Private cmdInterfaceValue As String
Private lobbyAdminIdentityHashes As StringVector
Private lobbyAdminPublicKeys As StringVector
Private gameStoryOutputModes As GameStoryLogger_OutputModes
Private gameStoryOutputKey As Long
Private gameStoryOutputNaming As Long
Private gameStoryFrameLoggingInterval As Long

' display3d.h

' keybind.h

' radar.h

' multiplay.h
Public Const MAX_SPECTATOR_SLOTS As Long = 10

' multiplaydefs.h: enum class PLAYER_LEAVE_MODE : uint8_t
Public Const PLAYER_LEAVE_MODE_DESTROY_RESOURCES As Long = 0
Public Const PLAYER_LEAVE_MODE_SPLIT_WITH_TEAM As Long = 1
Public Const PLAYER_LEAVE_MODE_MAX As Long = PLAYER_LEAVE_MODE_SPLIT_WITH_TEAM
Public Const PLAYER_LEAVE_MODE_DEFAULT As Long = PLAYER_LEAVE_MODE_SPLIT_WITH_TEAM

' lib/sequence/sequence.h: SCANLINE_MODE
Public Const SCANLINES_OFF As Long = 0
Public Const SCANLINES_50 As Long = 1
Public Const SCANLINES_BLACK As Long = 2

' radar.h: RADAR_DRAW_MODE

' ai.h: enum class AIDifficulty : int8_t (the values configuration.cpp and action.cpp use), enum AllianceType
Public Const AIDifficulty_SUPEREASY As Long = 0
Public Const AIDifficulty_INSANE As Long = 4
Public Const AIDifficulty_DISABLED As Long = -1
Public Const AIDifficulty_HUMAN As Long = -2
Public Const NO_ALLIANCES As Long = 0        ' FFA
Public Const ALLIANCES_ As Long = 1          ' Players can make and break alliances during the game.
Public Const ALLIANCES_TEAMS As Long = 2     ' Alliances are set before the game.
Public Const ALLIANCES_UNSHARED As Long = 3  ' Alliances are set before the game. No shared research.

' multiplaydefs.h: enum class BLIND_MODE : uint8_t
Public Const BLIND_MODE_NONE As Long = 0
Public Const BLIND_MODE_BLIND_LOBBY As Long = 1
Public Const BLIND_MODE_BLIND_LOBBY_SIMPLE_LOBBY As Long = 2
Public Const BLIND_MODE_BLIND_GAME As Long = 3
Public Const BLIND_MODE_BLIND_GAME_SIMPLE_LOBBY As Long = 4

' miscimd.h: the MI_ enumerators used so far

' multiplay.h: ScavType, CampType, PowerSetting
Public Const NO_SCAVENGERS As Long = 0
Public Const SCAVENGERS As Long = 1
Public Const ULTIMATE_SCAVENGERS As Long = 2
Public Const CAMP_CLEAN As Long = 0
Public Const CAMP_BASE As Long = 1
Public Const CAMP_WALLS As Long = 2
Public Const LEV_LOW As Long = 0
Public Const LEV_MED As Long = 1
Public Const LEV_HI As Long = 2

' multiint.h
Public Const DEFAULT_SKIRMISH_AI_SCRIPT_NAME As String = "nexus.js"

' terrain_defs.h: enum TerrainShaderQuality

' nethelpers.h
Public Const WZ_DEFAULT_PUBLIC_IPv4_LOOKUP_SERVICE_URL As String = "https://api.ipify.org?format=json"
Public Const WZ_DEFAULT_PUBLIC_IPv4_LOOKUP_SERVICE_JSONKEY As String = "ip"
Public Const WZ_DEFAULT_PUBLIC_IPv6_LOOKUP_SERVICE_URL As String = "https://api6.ipify.org?format=json"
Public Const WZ_DEFAULT_PUBLIC_IPv6_LOOKUP_SERVICE_JSONKEY As String = "ip"

' multiplay.h: struct MULTIPLAYERGAME. map and name are char[128] (Strings
' of at most 127 characters here); modHashes is not used yet.
Public Type MULTIPLAYERGAME
    typ As Long                         ' LEVEL_TYPE: DMATCH/CAMPAIGN/SKIRMISH/TEAMPLAY etc...
    scavengers As Long                  ' uint8_t: ultimate scavengers, scavengers, no scavengers
    map As String                       ' name of multiplayer map being used.
    maxPlayers As Long                  ' uint8_t: max players to allow
    name_ As String                     ' game name   (to be used)
    hash As Sha256                      ' Hash of map file. Zero if built in.
    modHashes() As Sha256               ' Hash of mods: modHashes(0 .. modHashCount - 1)
    modHashCount As Long
    power As Long                       ' uint32_t: power level for arena game
    base As Long                        ' uint8_t: clean/base/base&defence
    alliance As Long                    ' uint8_t: no/yes/locked
    mapHasScavengers As Boolean
    isMapMod As Boolean                 ' if a map has mods
    isRandom As Boolean                 ' If a map is non-static.
    techLevel As Long                   ' uint32_t: what technology level is being used
    inactivityMinutes As Long           ' uint32_t
    gameTimeLimitMinutes As Long        ' uint32_t
    playerLeaveMode As Long             ' PLAYER_LEAVE_MODE
    blindMode As Long                   ' BLIND_MODE (NONE = 0)
End Type

' multiplay.h: enum class InGameSide : bool
Public Const InGameSide_HOST_OR_SINGLEPLAYER As Boolean = True
Public Const InGameSide_MULTIPLAYER_CLIENT As Boolean = False

' multiplay.h: struct MULTISTRUCTLIMITS
Public Type MULTISTRUCTLIMITS
    id As Long                          ' uint32_t
    limit As Long                       ' uint32_t
End Type

' multiplay.h: struct MULTIPLAYERINGAME (the fields used so far); a std::vector<MULTISTRUCTLIMITS> is an array and its count
Public Type MULTIPLAYERINGAME
    VerifiedIdentity(0 To MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS - 1) As Boolean      ' [MAX_CONNECTED_PLAYERS]: if the multistats identity has been verified.
    localJoiningInProgress As Boolean   ' used before we know our player number.
    JoiningInProgress(0 To MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS - 1) As Boolean     ' [MAX_CONNECTED_PLAYERS]
    PendingDisconnect(0 To MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS - 1) As Boolean     ' [MAX_CONNECTED_PLAYERS]: used to mark players who have disconnected after the game has "fired up" but before it actually starts (i.e. pre-game / loading phase) - UI only
    has_TimeEveryoneIsInGame As Boolean ' optional<int32_t> TimeEveryoneIsInGame: has_value()
    TimeEveryoneIsInGame As Long        ' its value
    side As Boolean                     ' InGameSide
    has_endTime As Boolean              ' optional<std::chrono::steady_clock::time_point> endTime: has_value()
    structureLimits() As MULTISTRUCTLIMITS              ' std::vector<MULTISTRUCTLIMITS>: structureLimits(0 .. structureLimitsCount - 1)
    structureLimitsCount As Long
    flags As Long                       ' uint8_t: Bitmask, shows which structures are disabled.
    lastAppliedStructureLimits() As MULTISTRUCTLIMITS   ' a bit of a hack to retain the structureLimits used when loading / starting a game
    lastAppliedStructureLimitsCount As Long
    lastPlayerDataCheck2 As Currency    ' std::chrono::steady_clock::time_point: a QueryPerformanceCounter count
    isAllPlayersDataOK As Boolean
    PingTimes(0 To MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS - 1) As Long            ' UDWORD [MAX_CONNECTED_PLAYERS]: store for pings.
    localOptionsReceived As Boolean     ' used to show if we have game options yet..
    has_joinTimes(0 To MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS - 1) As Boolean     ' std::array<optional<std::chrono::steady_clock::time_point>, MAX_CONNECTED_PLAYERS> joinTimes: has_value()
    joinTimes(0 To MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS - 1) As Currency        ' its values: QueryPerformanceCounter counts
    has_lastReadyTimes(0 To MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS - 1) As Boolean ' lastReadyTimes: has_value()
    lastReadyTimes(0 To MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS - 1) As Currency
    has_lastNotReadyTimes(0 To MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS - 1) As Boolean ' lastNotReadyTimes: has_value()
    lastNotReadyTimes(0 To MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS - 1) As Currency
    secondsNotReady(0 To MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS - 1) As Double    ' std::array<uint64_t, MAX_CONNECTED_PLAYERS>: updated when player status switches to ready
    has_playerLeftGameTime(0 To MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS - 1) As Boolean ' std::array<optional<uint32_t>, MAX_CONNECTED_PLAYERS> playerLeftGameTime: has_value()
    playerLeftGameTime(0 To MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS - 1) As Long   ' its values (uint32_t): records when the player leaves the game (as a player)
    endTime As Currency                 ' optional<std::chrono::steady_clock::time_point> endTime's value (has_endTime): a QueryPerformanceCounter count
    muteChat(0 To MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS - 1) As Boolean          ' the local client-set mute status for this player
    LagCounter(0 To MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS - 1) As Long            ' int [MAX_CONNECTED_PLAYERS]
    DesyncCounter(0 To MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS - 1) As Long         ' int [MAX_CONNECTED_PLAYERS]
    DataIntegrity(0 To MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS - 1) As Boolean      ' [MAX_CONNECTED_PLAYERS]
    hostChatPermissions(0 To MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS - 1) As Boolean ' std::array<bool, MAX_CONNECTED_PLAYERS>: the *host*-set free chat permission status for players (true if free chat is allowed, false if only Quick Chat is allowed)
    startTime As Currency               ' std::chrono::steady_clock::time_point: a QueryPerformanceCounter count
    has_lastSentPlayerDataCheck2(0 To MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS - 1) As Boolean ' optional<std::chrono::steady_clock::time_point> lastSentPlayerDataCheck2[MAX_CONNECTED_PLAYERS]: has_value()
End Type

' multiplay.cpp
Public game As MULTIPLAYERGAME          ' info to describe game.
Public ingame As MULTIPLAYERINGAME
Public bMultiPlayer As Boolean          ' true when more than 1 player.

' multiint.cpp
Public sPlayer As String                ' char[128]: player name (to be used)
Public defaultSkirmishAI As String

' display3d.cpp

' display.cpp
Private bShakingPermitted As Boolean
Private cameraAccel As Boolean
Private bInvertMouse As Boolean
Private bRightClickOrders As Boolean
Private bDrawShadows As Boolean
Private bEdgeScrollOutsideWindowBounds As Boolean

' radar.cpp

' ingameop.cpp, loadsave.cpp

' hci/groups.cpp, seqdisp.cpp, difficulty.cpp, texture.cpp, terrain.cpp
Private groupButtonEnabled As Boolean
Private bSeqSubtitles As Boolean
Private maxTextureSize As Long          ' the maximum size texture we will create
Private terrainShaderQuality As Long
Private maxTerrainMappingTextureSize As Long ' uninitialized - pick default on first run

' nethelpers.cpp
Private publicIPv4LookupService As String
Private publicIPv4LookupServiceJSONKey As String
Private publicIPv6LookupService As String
Private publicIPv6LookupServiceJSONKey As String

Private scanlineMode As Long

'--------------------------------------------------------------------------
' The C++ initialisers of the globals above (a VB module has none): called
' first thing in realmain, before anything reads them.

' --- cluster A2: what levels.cpp, data.cpp and init.cpp's systemInitialise / systemShutdown call ---

' ai.h: enum class AIDifficulty (DEFAULT)
Public Const AIDifficulty_DEFAULT As Long = 2

' game.h: enum class GameLoadDetails::GameLoadType
Public Const GameLoadType_UserSaveGame As Long = 0
Public Const GameLoadType_Level As Long = 1
Public Const GameLoadType_MapPackage As Long = 2

Public releaseObjectives As Boolean     ' mission.cpp



' netplay.h: enum LOBBY_ERROR_TYPES, SYNC_FLAG (WzNet does not export them yet)
Public Enum LOBBY_ERROR_TYPES
    ERROR_NOERROR
    ERROR_CONNECTION
    ERROR_FULL
    ERROR_INVALID
    ERROR_KICKED
    ERROR_WRONGVERSION
    ERROR_WRONGPASSWORD
    ERROR_HOSTDROPPED
    ERROR_WRONGDATA
    ERROR_UNKNOWNFILEISSUE
    ERROR_REDIRECT
End Enum

Public Const SYNC_FLAG As Long = &H10000000     ' special flag used for logging.

' transporter.h

' display3ddef.h

' effects.h

' hci.cpp: std::shared_ptr<W_SCREEN> psWScreen (the widgets: cluster G)

' cluster B: what the objects call in modules of later clusters (effects,
' scores, messages, sound, scripts, the interface) or not yet ported
' (fpath's path jobs, formations, visibility, factories' commanders). Each
' one does nothing, which is what the C++ does to the game state while
' those modules have no state of their own (no messages, sounds, scripts,
' formations or path jobs exist yet).

' display3d.h: enum TILE_ID (special tile types)





' multiplay.h: enum STRUCTURE_INFO
Public Enum STRUCTURE_INFO
    STRUCTUREINFO_MANUFACTURE
    STRUCTUREINFO_CANCELPRODUCTION
    STRUCTUREINFO_HOLDPRODUCTION
    STRUCTUREINFO_RELEASEPRODUCTION
    STRUCTUREINFO_HOLDRESEARCH
    STRUCTUREINFO_RELEASERESEARCH
End Enum

' messagedef.h
Public Enum MESSAGE_TYPE
    MSG_RESEARCH        ' Research message
    MSG_CAMPAIGN        ' Campaign message
    MSG_MISSION         ' Mission Report messages
    MSG_PROXIMITY       ' Proximity message
    MSG_TYPES
End Enum

' activity.h: enum class ActivitySink::GameMode
Public Enum ActivitySink_GameMode
    ActivitySink_GameMode_MENUS
    ActivitySink_GameMode_TUTORIAL
    ActivitySink_GameMode_CAMPAIGN
    ActivitySink_GameMode_CHALLENGE
    ActivitySink_GameMode_SKIRMISH
    ActivitySink_GameMode_HOSTING_IN_LOBBY
    ActivitySink_GameMode_JOINING_IN_PROGRESS   ' joined but waiting on game information from host
    ActivitySink_GameMode_JOINING_IN_LOBBY
    ActivitySink_GameMode_MULTIPLAYER
End Enum

' campaigninfo.cpp: the campaign's name (optional<std::string>) and the cam tweak options
' (std::unordered_map<std::string, nlohmann::json>: a json object of them)
Private campaignName As String
Private hasCampaignName As Boolean
Private camTweakOptions As nlohmann_json

' B12 (the game's start): what main.cpp, init.cpp, game.cpp, mission.cpp and the multiplayer start-up use
' from later clusters

' lib/framework/cursors.h: enum CURSOR (the one the game sets so far; the cursors: cluster D)

' atmos.h: enum WT_CLASS

' loop.h: enum LOOP_MISSION_STATE (the first; the mission loop states come with the display clusters), loop.cpp's
' loopMissionState (LMS_NORMAL)
Public Const LMS_NORMAL As Long = 0
Public loopMissionState As Long

' input/context.h: InputContext::DEBUG_MISC and InputContext::State::ACTIVE (the key bindings' contexts: cluster G;
' only gInputManager_contexts_set takes them)


' hci.h

' challenge.cpp: std::string challengeFileName (the challenges come with cluster G)

' ai.h: enum class AIDifficulty (EASY)
Public Const AIDifficulty_EASY As Long = 1
Public Const LMS_SETUPMISSION As Long = 1       ' Do the setup for the mission
Public Const LMS_SAVECONTINUE As Long = 2       ' Save the game and then continue with the next mission
Public Const LMS_NEWLEVEL As Long = 3           ' Start a new level
Public Const LMS_LOADGAME As Long = 4           ' Load a savegame
Public Const LMS_CLEAROBJECTS As Long = 5       ' Clear all the objects

' hci.h (cluster G): the reticule buttons' ids and the interface's mode

' C1 (renderLoop, multiPlayerLoop)
' hci.h: enum INT_RETVAL; INTMODE's INT_INGAMEOP

' ingameop.h: the quit buttons' ids

' display.h: enum DragState; dragBox3D.status and wallDrag.status (display.cpp: cluster E)

' loadsave.cpp: bRequestLoad

' lib/widget/widget.cpp: static bool bWidgetsActive = true; (stubs_game_init)
Private bWidgetsActive As Boolean

' C1 (the script engine): what qtscript.cpp, wzapi.cpp and the Chakra backend use from later clusters

' advvis.cpp: static bool bRevealActive = true; (stubs_game_init)
Private bRevealActive As Boolean

' campaigninfo.cpp: static uint32_t camNumber = 1; (stubs_game_init)
Private camNumber As Long


' hci/quickchat.h: struct WzQuickChatTargeting (the quick chat: cluster G); all is all_ (a VB word);
' std::unordered_set<uint32_t> specificPlayers is specificPlayers(0 .. specificPlayers_size - 1)
Public Type WzQuickChatTargeting
    all_ As Boolean
    humanTeammates As Boolean
    aiTeammates As Boolean
    specificPlayers() As Long
    specificPlayers_size As Long
End Type

' missiondef.h, mission.h (belong in mission_c)

' loop.cpp: LEVEL_TYPE nextMissionType = LEVEL_TYPE::LDS_NONE; prevMissionType = LEVEL_TYPE::LDS_NONE
' (belong in loop_c; their initialiser: loop_c_init sets both to LEVEL_TYPE_LDS_NONE)
Public nextMissionType As Long
Public prevMissionType As Long

'   -> in loop_c_init (or stubs_game_init):  nextMissionType = LEVEL_TYPE_LDS_NONE: prevMissionType = LEVEL_TYPE_LDS_NONE
' screens/guidescreen.h: enum class ShowTopicFlags : int (the guide: cluster G)
Public Const ShowTopicFlags_FirstAdd As Long = 1      ' 1 << 0: open guide only if this topic is newly-added (this playthrough)
Public Const ShowTopicFlags_NeverViewed As Long = 2   ' 1 << 1: open guide only if this topic has never been viewed by the player before

' ai.h: enum class AIDifficulty (fork P adds EASY = 1; these two complete getUsefulConstants' list)
Public Const AIDifficulty_MEDIUM As Long = 2
Public Const AIDifficulty_HARD As Long = 3

Public Sub stubs_game_init()
    bWidgetsActive = True
    bRevealActive = True
    radarOnScreen = True
    radarPermitted = True
    BlueprintTrackAnimationSpeed = 20
    bShakingPermitted = True
    cameraAccel = True
    bInvertMouse = True
    bDrawShadows = True
    bEdgeScrollOutsideWindowBounds = True   ' DEFAULT_EDGE_SCROLL_OUTSIDE_WINDOW
    rotateRadar = True
    radarRotationArrow = True
    groupButtonEnabled = True
    maxTextureSize = 2048
    terrainShaderQuality = UNINITIALIZED_PICK_DEFAULT
    publicIPv4LookupService = WZ_DEFAULT_PUBLIC_IPv4_LOOKUP_SERVICE_URL
    publicIPv4LookupServiceJSONKey = WZ_DEFAULT_PUBLIC_IPv4_LOOKUP_SERVICE_JSONKEY
    publicIPv6LookupService = WZ_DEFAULT_PUBLIC_IPv6_LOOKUP_SERVICE_URL
    publicIPv6LookupServiceJSONKey = WZ_DEFAULT_PUBLIC_IPv6_LOOKUP_SERVICE_JSONKEY
    game.blindMode = BLIND_MODE_NONE
End Sub

'--------------------------------------------------------------------------
' frontend.cpp, keybind.cpp, loadsave.cpp




'--------------------------------------------------------------------------
' stdinreader.cpp, multilobbycommands.cpp, gamehistorylogger.cpp

' C: also starts the interface's thread when the game starts
Public Sub configSetCmdInterface(ByVal mode As Long, ByVal value As String)
    cmdInterfaceMode = mode
    cmdInterfaceValue = value
End Sub

Public Function wz_command_interface_enabled() As Boolean
    wz_command_interface_enabled = (cmdInterfaceMode <> WZ_Command_Interface_None)
End Function

' C: false (and an error) for a malformed hash
Public Function addLobbyAdminIdentityHash(ByVal playerIdentityHash As String) As Boolean
    StringVector_push_back lobbyAdminIdentityHashes, playerIdentityHash
    addLobbyAdminIdentityHash = True
End Function

' C: false (and an error) for a key that does not decode
Public Function addLobbyAdminPublicKey(ByVal publicKeyB64Str As String) As Boolean
    StringVector_push_back lobbyAdminPublicKeys, publicKeyB64Str
    addLobbyAdminPublicKey = True
End Function

' GameStoryLogger::instance().setOutputModes(modes) and the other setters
Public Sub GameStoryLogger_setOutputModes(modes As GameStoryLogger_OutputModes)
    gameStoryOutputModes = modes
End Sub

Public Sub GameStoryLogger_setOutputKey(ByVal key As Long)
    gameStoryOutputKey = key
End Sub

Public Sub GameStoryLogger_setOutputNaming(ByVal naming As Long)
    gameStoryOutputNaming = naming
End Sub

Public Sub GameStoryLogger_setFrameLoggingInterval(ByVal interval As Long)
    gameStoryFrameLoggingInterval = interval
End Sub

'--------------------------------------------------------------------------
' multiint.cpp

'--------------------------------------------------------------------------
' display3d.cpp, display.cpp



















'--------------------------------------------------------------------------
' hci/groups.cpp, seqdisp.cpp, difficulty.cpp





'--------------------------------------------------------------------------
' texture.cpp, terrain.cpp







'--------------------------------------------------------------------------
' nethelpers.cpp

Public Sub setPublicIPv4LookupService(ByVal publicIPv4LookupServiceUrl As String, ByVal jsonKey As String)
    If Len(publicIPv4LookupServiceUrl) = 0 Or Len(jsonKey) = 0 Then
        Exit Sub
    End If
    publicIPv4LookupService = publicIPv4LookupServiceUrl
    publicIPv4LookupServiceJSONKey = jsonKey
End Sub

Public Function getPublicIPv4LookupServiceUrl() As String
    getPublicIPv4LookupServiceUrl = publicIPv4LookupService
End Function

Public Function getPublicIPv4LookupServiceJSONKey() As String
    getPublicIPv4LookupServiceJSONKey = publicIPv4LookupServiceJSONKey
End Function

Public Sub setPublicIPv6LookupService(ByVal publicIPv6LookupServiceUrl As String, ByVal jsonKey As String)
    If Len(publicIPv6LookupServiceUrl) = 0 Or Len(jsonKey) = 0 Then
        Exit Sub
    End If
    publicIPv6LookupService = publicIPv6LookupServiceUrl
    publicIPv6LookupServiceJSONKey = jsonKey
End Sub

Public Function getPublicIPv6LookupServiceUrl() As String
    getPublicIPv6LookupServiceUrl = publicIPv6LookupService
End Function

Public Function getPublicIPv6LookupServiceJSONKey() As String
    getPublicIPv6LookupServiceJSONKey = publicIPv6LookupServiceJSONKey
End Function

'--------------------------------------------------------------------------
' component.cpp, activity.cpp, seqdisp.cpp

' ActivityManager::instance().changedSetting(settingKey, settingValue)
Public Sub ActivityManager_changedSetting(ByVal settingKey As String, ByVal settingValue As String)
End Sub

' ActivityManager::instance().rebuiltSearchPath()
Public Sub ActivityManager_rebuiltSearchPath()
End Sub

' lib/exceptionhandler/dumpinfo.cpp: adds a line to the crash dump's info
Public Sub addDumpInfo(ByVal inbuffer As String)
End Sub

' ActivityManager::instance().beginLoadingSettings() / endLoadingSettings()
Public Sub ActivityManager_beginLoadingSettings()
End Sub

Public Sub ActivityManager_endLoadingSettings()
End Sub

Public Sub seq_setScanlineMode(ByVal mode As Long)
    scanlineMode = mode
End Sub

Public Function seq_getScanlineMode() As Long
    seq_getScanlineMode = scanlineMode
End Function

'--------------------------------------------------------------------------
' keybind.cpp, game.cpp



'--------------------------------------------------------------------------
' cluster A2: levels.cpp's callees (init.cpp's stages, mission.cpp, game.cpp,
' effects.cpp, multiplay.cpp, gtime.cpp, qtscript.cpp, activity.cpp)




Public Sub ActivityManager_loadedLevel(ByVal levelType As Long, ByVal levelName As String)
End Sub

' lib/framework/frame.cpp (SHA-256 is not ported): the zero hash
Public Sub findHashOfFile(ByVal realFileName As String, hash As Sha256)
    Sha256_setZero hash
End Sub

'--------------------------------------------------------------------------
' cluster A2: data.cpp's callees. The images, tracks and string resources
' of the DLLs (iV_Image, IMAGEFILE, TRACK) wait here until those DLLs design
' the types: no image or track loads.

' cluster A5: what map.cpp calls in lib/ivis_opengl/pielighting.cpp,
' src/texture.cpp and terrain.cpp (clusters D and E): nothing is drawn yet.












'--------------------------------------------------------------------------
' cluster A2: init.cpp's systemInitialise / systemShutdown callees in the game



Public Function notificationsInitialize() As Boolean
    notificationsInitialize = True
End Function

Public Sub notificationsShutDown()
End Sub


Public Sub urlRequestShutdown()
End Sub




Public Function multiShutdown() As Boolean
    multiShutdown = True
End Function

Public Sub shutdownKnownPlayers()
End Sub


' lib/wzmaplib/plugins/ZipIOProvider (WzMap): WzZipIOPHYSFSSourceReadProvider::make, shared_ptr<WzZipIOSourceReadProvider>
' (zip archives are not read yet: nullptr)
Public Function WzZipIOPHYSFSSourceReadProvider_make(ByVal filename As String) As Object
    Set WzZipIOPHYSFSSourceReadProvider_make = Nothing
End Function

' ---- cluster B's callees in later clusters (see the declarations)






' qtscript.h
Public Sub jsDebugMessageUpdate()
End Sub


' lib/framework/wzapp.h (WzPlatform's wzapp, cluster F): SDL_GetCPUCount(), the logical CPUs
Public Function wzGetLogicalCPUCount() As Long
    Dim si As SYSTEM_INFO
    GetSystemInfo si
    wzGetLogicalCPUCount = si.dwNumberOfProcessors
End Function












' cluster B, movement (B6): what move.cpp and its helpers call in modules
' not ported yet











' cluster B, actions and orders (B7): what action.cpp, order.cpp and their
' helpers call in modules not ported yet




















' lib/framework/input.h (no keys held in a hidden run)
Public Function specialOrderKeyDown() As Boolean
    specialOrderKeyDown = False
End Function



Public Sub jsDebugSelected(ByVal psObj As Long)
End Sub




' cluster B, combat (B8): what combat.cpp, projectile.cpp, ai.cpp, visibility.cpp and the damage functions call
' in modules not ported yet






' effects.h
Public Sub addDroidDeathAnimationEffect(ByVal psDroid As Long)
End Sub


' multiplay.h: getPlayerName(uint32_t player, bool treatAsNonHost = false) (the players' names come with the network)


' multistat.h: the players' statistics (the multiplayer statistics come with the network)
Public Sub updateMultiStatsDamage(ByVal attacker As Long, ByVal defender As Long, ByVal inflicted As Long)
End Sub

Public Sub updateMultiStatsKills(ByVal psKilled As Long, ByVal player As Long)
End Sub



' gamehistorylogger.cpp: GameStoryLogger::instance().logGameFrame()
Public Sub GameStoryLogger_logGameFrame()
End Sub



' multistat.h (the multiplayer statistics come with the network)
Public Sub incrementMultiStatsResearchPotential(ByVal player As Long)
End Sub

Public Sub incrementMultiStatsResearchPerformance(ByVal player As Long)
End Sub


'--------------------------------------------------------------------------
' B10: what multiplay.cpp's recvMessage dispatches to outside the game state
' (each later cluster ports its own), and gameLoop's render side



' netlog.h: NETlogEntry (no net log file is open: false, as C++ without one)
Public Function NETlogEntry(ByVal str_ As String, ByVal a As Long, ByVal b As Long) As Boolean
    NETlogEntry = False
End Function

' netplay.h (the connections: cluster J): no player has a connection; kicking closes none
Public Sub NETplayerKicked(ByVal index As Long, Optional ByVal quiet As Boolean = False)
End Sub

Public Function NETplayerHasConnection(ByVal index As Long) As Boolean
    NETplayerHasConnection = False
End Function

' netplay.h: NetPlay.players[i].IPtextAddress (the connections: cluster J)
Public Function NetPlay_players_IPtextAddress(ByVal i As Long) As String
    NetPlay_players_IPtextAddress = ""
End Function

' multiplay.cpp: sendInGameSystemMessage (the chat: cluster J): neither printed nor sent
Public Sub sendInGameSystemMessage(ByVal Text As String)
End Sub


Public Sub displayKickReasonPopup(ByVal reason As String)
End Sub

' activity.h: ActivityManager::instance().wasKickedByPlayer(NetPlay.players[kicker], kick_type, reason) (kicker: the index)
Public Sub ActivityManager_wasKickedByPlayer(ByVal kicker As Long, ByVal kick_type As LOBBY_ERROR_TYPES, ByVal reason As String)
End Sub

' multistat.h: addKnownPlayer(name, getMultiStats(player).identity, override) (the player statistics and identities:
' cluster J; the player's index stands for the key)
Public Sub addKnownPlayer(ByVal name_ As String, ByVal identityPlayer As Long, Optional ByVal override As Boolean = False)
End Sub

' stdinreader.cpp (the command interface is not ported): recvMessage's NET_PLAYERRESPONDING output, the
' "WZEVENT: playerResponding: ..." line (identity, verified status, name in base64, address) and
' wz_command_interface_output_room_status_json()
Public Sub wz_command_interface_output_playerResponding(ByVal player_id As Long)
End Sub

' multirecv.h, multijoin.h, multistat.h, multivote.h, multigifts.h, cheat.h, quickchat.h, teamstrategy.h and
' multiplay.cpp's static recvBeacon and recvDataCheck2: the messages recvMessage hands on that do not change the
' game state yet: chat, beacons, data checks, pings, votes, host configuration, player statistics (the network,
' cluster J); gifts and alliances, a player leaving, debug mode and debug features (with the game messages that
' change the game state, later in cluster B or J). Each leaves its message undecoded (recvMessage pops it).
Public Function receiveInGameTextMessage(ByVal queue As Long) As Boolean
    receiveInGameTextMessage = True
End Function

Public Function recvDataCheck(ByVal queue As Long) As Boolean
    recvDataCheck = True
End Function

Public Function recvDataCheck2(ByVal queue As Long) As Boolean
    recvDataCheck2 = True
End Function

Public Function recvTextMessageAI(ByVal queue As Long) As Boolean
    recvTextMessageAI = True
End Function

Public Function recvSpecInGameTextMessage(ByVal queue As Long) As Boolean
    recvSpecInGameTextMessage = True
End Function

Public Function recvStrategyPlanUpdate(ByVal queue As Long) As Boolean
    recvStrategyPlanUpdate = True
End Function

Public Function recvQuickChat(ByVal queue As Long) As Boolean
    recvQuickChat = True
End Function


Public Function recvPing(ByVal queue As Long) As Boolean
    recvPing = True
End Function

Public Function recvVote(ByVal queue As Long, Optional ByVal inLobby As Boolean = True) As Boolean
    recvVote = True
End Function

Public Function recvVoteRequest(ByVal queue As Long) As Boolean
    recvVoteRequest = True
End Function

Public Function recvHostConfig(ByVal queue As Long) As Boolean
    recvHostConfig = True
End Function

Public Function recvMultiStats(ByVal queue As Long) As Boolean
    recvMultiStats = True
End Function

Public Sub recvPlayerLeft(ByVal queue As Long)
End Sub

'--------------------------------------------------------------------------
' B11: the replays' options (WZGameReplayOptionsHandler)








' init.h: bool setSpecialInMemoryMap(std::vector<uint8_t>&& mapArchiveData): mapArchiveData(0 .. size - 1) (the in-memory
' map archive is not mounted: false)
Public Function setSpecialInMemoryMap(mapArchiveData() As Byte, ByVal size As Long) As Boolean
    setSpecialInMemoryMap = False
End Function

'--------------------------------------------------------------------------
' B12 (the game's start): what main.cpp's loops, init.cpp's stages, game.cpp's loadGame, mission.cpp's
' start and the multiplayer start-up call in later clusters





Public Function wzIsMaximized() As Boolean
    wzIsMaximized = False
End Function

' lib/sdl (WzPlatform, until cluster D adds them to PlatformApi): the cursor, mouse grab and window
Public Sub wzSetCursor(ByVal cur As Long)
End Sub

Public Sub wzApplyCursor()
End Sub

Public Sub wzGrabMouse()
End Sub

Public Sub wzReleaseMouse()
End Sub

Public Sub wzSetWindowIsResizable(ByVal resizable As Boolean)
End Sub

' C: the new window mode after toggling fullscreen
Public Function wzAltEnterToggleFullscreen() As Long
    wzAltEnterToggleFullscreen = wzGetCurrentWindowMode()
End Function



































Public Sub shutdownGameGuide()
End Sub

Public Function specStatsViewShutdown() As Boolean
    specStatsViewShutdown = True
End Function

Public Sub strategyPlansInit()
End Sub



Public Sub runNotifications()
End Sub

















Public Sub ActivityManager_preSystemShutdown()
End Sub

Public Sub ActivityManager_shutdown()
End Sub

Public Sub ActivityManager_navigateToMenu(ByVal menuName As String)
End Sub

Public Sub GameStoryLogger_logStartGame()
End Sub

Public Sub GameStoryLogger_reset()
End Sub

' C: printf-style output to the command interface (not ported)
Public Sub wz_command_interface_output(ByVal str_ As String, ParamArray args() As Variant)
End Sub

Public Sub wz_command_interface_process_queued_status_output()
End Sub



' init.cpp: bool clearSpecialInMemoryMap() (the in-memory map archive is not mounted: nothing to unmount)
Public Function clearSpecialInMemoryMap() As Boolean
    clearSpecialInMemoryMap = True
End Function

' lib/netplay (WzNet, the connections: cluster J), multijoin_helpers.cpp, multiplay.cpp's start screen
Public Sub NETadjustConnectedTimeoutForClients()
End Sub

Public Function NETinit(ByVal pt As Long) As Long
    NETinit = 0
End Function

Public Sub NETinitPortMapping()
End Sub

Public Sub joinGameFromIPOrHostnameConnectionStr(ByVal ipOrHostnameConnectionStr As String, Optional ByVal asSpectator As Boolean = False)
End Sub

Public Sub joinLobbyGame(ByVal lobbyAddress As String, ByVal lobbyGameId As String, Optional ByVal asSpectator As Boolean = False)
End Sub

Public Function multiStartScreenInit() As Boolean
    multiStartScreenInit = True
End Function

' In files forks L and P own: drop each one they ported.













' multiplay.cpp: bool setPlayerName(int player, const char *sName) (its real body; multiplay.cpp's names come with cluster J)
Public Function setPlayerName(ByVal player As Long, ByVal sName As String) As Boolean
    If Not (player < MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS And player >= 0) Then ASSERT_FAILURE "player < MAX_CONNECTED_PLAYERS && player >= 0", "src/multiplay.cpp:892", "setPlayerName", "Player index (%u) out of range", player: setPlayerName = False: Exit Function
    NetPlay_players_name(player) = Left$(sName, 64 - 1)    ' sstrcpy(NetPlay.players[player].name, sName) (char name[StringSize])
    setPlayerName = True
End Function


































' spectatorwidgets.cpp (cluster G)
Public Function specLayerInit(Optional ByVal showButton As Boolean = True) As Boolean
    specLayerInit = True
End Function

'--------------------------------------------------------------------------
' C1 (the script engine): what qtscript.cpp, wzapi.cpp and the Chakra backend call in later clusters

' C1: what qtscript.cpp calls in later clusters
' wzscriptdebug.cpp (the script debugger dialog: cluster G): bool jsDebugShutdown()
Public Function jsDebugShutdown() As Boolean
    jsDebugShutdown = True
End Function

' wzscriptdebug.cpp (cluster G)
Public Sub jsDebugUpdateLabels()
End Sub

' wzscriptdebug.cpp (cluster G): jsDebugCreate(const std::shared_ptr<scripting_engine::DebugInterface>& debugInterface,
' const jsDebugShutdownHandlerFunction& shutdownFunc, bool readOnly = false) (the handler: a JSDEBUG_SHUTDOWN_ number)
Public Sub jsDebugCreate(ByVal debugInterface As Object, ByVal shutdownFunc As Long, Optional ByVal readOnly As Boolean = False)
End Sub



' multistat.cpp (the players' statistics: cluster J): void updateMultiStatsBuilt(BASE_OBJECT *psBuilt)
Public Sub updateMultiStatsBuilt(ByVal psBuilt As Long)
End Sub

' multistat.cpp (cluster J): void updateMultiStatsResearchComplete(RESEARCH *psResearch, UDWORD player)
Public Sub updateMultiStatsResearchComplete(ByVal psResearch As Long, ByVal player As Long)
End Sub

' gamehistorylogger.cpp (the game story log: cluster J): GameStoryLogger::instance().logResearchCompleted(psResearch, psStruct, player)
Public Sub GameStoryLogger_logResearchCompleted(ByVal psResearch As Long, ByVal psStruct As Long, ByVal player As Long)
End Sub

' gamehistorylogger.cpp (cluster J): GameStoryLogger::instance().logDebugModeChange(enabled)
Public Sub GameStoryLogger_logDebugModeChange(ByVal enabled As Boolean)
End Sub















' hci/quickchat.cpp (cluster G): targeting.specificPlayers.insert(player), to_WzQuickChatMessage (WzQuickChatMessage
' as its uint32_t value; not checked against the enum yet) and sendQuickChat (the message's network part: cluster J)
Public Sub WzQuickChatTargeting_insertSpecificPlayer(t As WzQuickChatTargeting, ByVal player As Long)
    Dim i As Long
    For i = 0 To t.specificPlayers_size - 1
        If t.specificPlayers(i) = player Then Exit Sub
    Next
    ReDim Preserve t.specificPlayers(0 To t.specificPlayers_size)
    t.specificPlayers(t.specificPlayers_size) = player
    t.specificPlayers_size = t.specificPlayers_size + 1
End Sub

Public Function to_WzQuickChatMessage(ByVal value As Long, ByRef output As Long) As Boolean
    output = value
    to_WzQuickChatMessage = True
End Function

Public Sub sendQuickChat(ByVal message As Long, ByVal fromPlayer As Long, targeting As WzQuickChatTargeting)
End Sub


' screens/guidescreen.cpp (the in-game guide: cluster G): void addGuideTopic(const std::string& guideTopicId,
' ShowTopicFlags showFlags = ShowTopicFlags::None, const std::unordered_set<std::string>& excludedTopicIDs = {});
' the overload with excluded topics is addGuideTopic_excluding (the set as a StringVector)
Public Sub addGuideTopic(ByVal guideTopicId As String, Optional ByVal showFlags As Long = 0)
End Sub

Public Sub addGuideTopic_excluding(ByVal guideTopicId As String, ByVal showFlags As Long, excludedTopicIDs As StringVector)
End Sub


' gamehistorylogger.cpp: GameStoryLogger::instance().logGameOver()
Public Sub GameStoryLogger_logGameOver()
End Sub












' multiplay.cpp: bool setGameStoryLogPlayerDataValue(uint32_t playerIndex, const std::string& key_str, const std::string&
' value_str), with its checks; NetPlay.scriptSetPlayerDataStrings (the game story log's data) is not in WzNet yet (cluster J)
Public Function setGameStoryLogPlayerDataValue(ByVal playerIndex As Long, ByVal key_str As String, ByVal value_str As String) As Boolean
    If ULT(MAX_PLAYERS, playerIndex) Then
        setGameStoryLogPlayerDataValue = False
        Exit Function
    End If

    If key_str <> "usertype" Then
        ' For now, only "usertype" is expected
        setGameStoryLogPlayerDataValue = False
        Exit Function
    End If

    ' (NetPlay.scriptSetPlayerDataStrings[playerIndex][key_str] = value_str)
    setGameStoryLogPlayerDataValue = True
End Function

' quickchat.cpp (the quick chat: clusters G / J): nlohmann::json getWzQuickChatMessageValueMap() (an empty object until then)
Public Function getWzQuickChatMessageValueMap() As nlohmann_json
    Set getWzQuickChatMessageValueMap = json_object()
End Function

'--------------------------------------------------------------------------
' C1: what loop.cpp's renderLoop and multiplay.cpp's multiPlayerLoop call in later clusters

















' notifications.cpp, screen overlays (cluster G)
Public Function isDraggingInGameNotification() As Boolean
    isDraggingInGameNotification = False
End Function














' multiplay.cpp, multiint.cpp, multistat.cpp, multivote.cpp, netplay (the network: cluster J): the data checks,
' pings, the lobby's player info and identities, the kick routines, the win sequence's effects
Public Sub sendDataCheck()
End Sub

Public Sub sendDataCheck2()
End Sub

Public Sub sendPing()
End Sub

Public Sub sendMultiStatsHostVerifiedIdentities(ByVal playerIndex As Long)
End Sub

Public Sub wz_command_interface_output_room_status_json(Optional ByVal queued As Boolean = False)
End Sub

Public Sub autoLagKickRoutine(ByVal now_ As Currency)
End Sub

Public Sub autoDesyncKickRoutine(ByVal now_ As Currency)
End Sub

Public Sub processPendingKickVotes()
End Sub

Public Function multiplayerWinSequence(ByVal firstCall As Boolean) As Boolean
    multiplayerWinSequence = True
End Function

'--------------------------------------------------------------------------
' C2 (the skirmish's start, gifts, chat and beacons): what they call in later clusters




' hci/quickchat.cpp (cluster G): the chat's spam mute (InGameChatMessage::send asks only when NetPlay.bComms):
' optional<std::chrono::steady_clock::time_point> playerSpamMutedUntil(uint32_t playerIdx) is has_value()
' with the value (a QueryPerformanceCounter count) in mutedUntil; never muted yet, nothing recorded
Public Function playerSpamMutedUntil(ByVal playerIdx As Long, mutedUntil As Currency) As Boolean
    playerSpamMutedUntil = False
End Function

Public Sub recordPlayerMessageSent(ByVal playerIdx As Long)
End Sub

' notifications.cpp (cluster F): bool cancelOrDismissNotificationsWithTag(const std::string& desiredTag,
' NotificationScope scope = NotificationScope::DISPLAYED_AND_QUEUED)
Public Function cancelOrDismissNotificationsWithTag(ByVal desiredTag As String) As Boolean
End Function

' notifications.cpp (cluster F): bool cancelOrDismissNotificationIfTag(matchTagFunc, scope); (VB) the one
' matchTagFunc used so far (startMultiplayerGame): the tags that start with prefix
Public Function cancelOrDismissNotificationIfTagHasPrefix(ByVal prefix As String) As Boolean
End Function

' notifications.cpp (cluster F): bool removeNotificationPreferencesIf(matchIdentifierFunc); (VB) the one
' matchIdentifierFunc used so far (WzMultiplayerOptionsTitleUI's start): the identifiers that start with prefix
Public Function removeNotificationPreferencesIfHasPrefix(ByVal prefix As String) As Boolean
End Function



' multivote.cpp (the lobby's votes: cluster J)
Public Sub resetLobbyChangeVoteData()
End Sub

Public Sub resetAllMultiOptionPrefValues()
End Sub

' multisync.cpp (cluster J): void multiSyncResetAllChallenges()
Public Sub multiSyncResetAllChallenges()
End Sub

' multistat.cpp (the identities, EcKey: cluster J): bool generateBlindIdentity()
Public Function generateBlindIdentity() As Boolean
    generateBlindIdentity = True
End Function

' multistat.cpp (the identities: cluster J): void multiStatsSetVerifiedHostIdentityFromJoin(const EcKey::Key &identity),
' without the identity (hostVerifiedJoinIdentities[NetPlay.hostPlayer])
Public Sub multiStatsSetVerifiedHostIdentityFromJoin()
End Sub

' multistat.cpp (the known players' SQLite database, knownPlayers.db: not kept): void initKnownPlayers()
Public Sub initKnownPlayers()
End Sub

