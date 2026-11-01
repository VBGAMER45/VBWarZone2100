Attribute VB_Name = "multiint_c"
' src/multiint.cpp: the JSON converters of the game's options (Sha256,
' MULTIPLAYERGAME, MULTISTRUCTLIMITS, InGameSide, MULTIPLAYERINGAME, PLAYER),
' JoinConnectionDescription's and KickRedirectInfo's (parseKickRedirectInfo),
' DataHashToJSON and decideWRF (static in C++: WZGameReplayOptionsHandler,
' its own class, calls it); the skirmish AIs (aidata, readAIs,
' setDefaultSkirmishAI, getDefaultSkirmishAI, getAIName) and
' loadMultiScripts; the multiplayer options screen's game state: its
' statics (locked, spectatorHost, ...), the map preview's state
' (loadMapPreview), the player configuration (resetPlayerConfiguration, the
' colours, the ready status, the test / challenge .json files:
' loadMapChallengeAndPlayerSettings), updateStructureDisabledFlags,
' startMultiplayerGame and sendRoomSystemMessage. The screen itself is
' WzMultiplayerOptionsTitleUI.cls (made by WzMultiplayerOptionsTitleUI_new);
' the statics and the static functions it uses are Public.
' defaultSkirmishAI is still stubs_game's. Missing: the lobby's widgets and
' choosers, the map preview's image, the chat box (displayRoomSystemMessage
' and displayRoomNotifyMessage are stand-ins), the notifications, and the
' lobby's messages, joins, kicks and votes (frontendMultiMessages ...):
' clusters F and J.
' C++'s to_json(j, x) is <type>_to_json(j, x): j becomes a new value; its
' from_json(j, x) is <type>_from_json(j, x). A PLAYER is a NetPlay.players
' entry: PLAYER_to_json(j, i) / PLAYER_from_json(j, i) take its index;
' PLAYER_ is the struct by value (a PLAYER local: get<PLAYER>()), with
' NetPlay_players_get / NetPlay_players_set between the two. A char[N]
' filled by sstrcpy is a String cut to N - 1 characters.
' nlohmann's enum serializer writes an enum as a signed integer (to_json:
' number_integer), whatever its underlying type; get<uint8_t> and
' get<int8_t> keep the low 8 bits.
Option Explicit

Private Const StringSize As Long = 64           ' netplay.h: size of strings used.
Private Const PATH_MAX As Long = 260            ' MAX_PATH on Windows

' multiplay.h: struct KickRedirectInfo
Public Type KickRedirectInfo
    connList() As JoinConnectionDescription     ' std::vector<JoinConnectionDescription>: connList(0 .. connListCount - 1)
    connListCount As Long
    has_gamePassword As Boolean                 ' optional<std::string> gamePassword = nullopt
    gamePassword As String
    asSpectator As Boolean                      ' = false
End Type

' netplay.h: struct PLAYER, the fields the replays keep (wzFiles, autoGame and IPtextAddress are not persisted)
Public Type PLAYER_
    name_ As String                     ' char[StringSize]: Player name
    position As Long                    ' int32_t: Map starting position
    colour As Long                      ' int32_t: Which colour slot this player is using
    allocated As Boolean                ' Allocated as a human player
    heartattacktime As Long             ' uint32_t: Time cardiac arrest started
    heartbeat As Boolean                ' If connection is still alive or not
    team As Long                        ' int32_t: Which team we are on (int32_t::max for spectator team)
    ready As Boolean                    ' player ready to start?
    ai As Long                          ' int8_t: index into sorted list of AIs, zero is always default AI
    difficulty As Long                  ' AIDifficulty (int8_t): difficulty level of AI
    faction As Long                     ' FactionID (uint8_t): which faction the player has
    isSpectator As Boolean              ' whether this slot is a spectator slot
    isAdmin As Boolean                  ' whether this slot has admin privs
End Type

' multiint.h
Public Const MAX_LEN_AI_NAME As Long = 40
Public Const AI_CUSTOM As Long = 127
Public Const AI_OPEN As Long = -2
Public Const AI_CLOSED As Long = -1

' multiint.h: struct AIDATA (AIDATA_: VB would take it for the array aidata)
Public Type AIDATA_
    name_ As String                     ' char[MAX_LEN_AI_NAME]
    js As String                        ' char[MAX_LEN_AI_NAME]
    tip As String                       ' char[255 + 128]: may contain optional AI tournament data
    difficultyTips(0 To 4) As String    ' char[5][255]: optional difficulty level info
    assigned As Long                    ' How many AIs have we assigned of this type
End Type

Private aidata() As AIDATA_             ' std::vector<AIDATA>: aidata(0 .. aidata_size - 1)
Private aidata_size As Long

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

' lib/netplay/nettypes.h
Private Const MAX_CONNECTED_PLAYERS As Long = MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS

' multiint.h: struct MultiplayOptionsLocked (the GUI lockdown of a challenge or a test)
Public Type MultiplayOptionsLocked
    scavengers As Boolean
    alliances As Boolean
    teams As Boolean
    power As Boolean
    difficulty As Boolean
    ai As Boolean
    position As Boolean
    bases As Boolean
    spectators As Boolean
    name_ As Boolean
    readybeforefull As Boolean
    cheats As Boolean
End Type

' multiint.h: the widgets WzMultiplayerOptionsTitleUI's startHost deletes
Public Const MULTIOP_REFRESH As Long = 10275
Public Const MULTIOP_FILTER_TOGGLE As Long = 30277

' multiplay.h
Public Const PING_LIMIT As Long = 4000         ' If ping is bigger than this, then worry and panic, and don't even try showing the ping.

' the notifications' tags
Public Const LOBBY_DISABLED_TAG As String = "lobbyDisabled"
Public Const SLOTTYPE_TAG_PREFIX As String = "slotType"
Public Const SLOTTYPE_NOTIFICATION_ID_PREFIX As String = "slotType::request::"

' multiint.cpp's statics (the ones WzMultiplayerOptionsTitleUI.cls uses are Public); difficultyList and
' difficultyValue are difficultyList_at and difficultyValue_at
Public locked As MultiplayOptionsLocked
Public spectatorHost As Boolean
Public defaultOpenSpectatorSlots As Long         ' uint16_t
Public bInActualHostedLobby As Boolean
Public bRequestedSelfMoveToPlayers As Boolean
Public bHostRequestedMoveToPlayers(0 To MAX_CONNECTED_PLAYERS - 1) As Boolean ' std::vector<bool>(MAX_CONNECTED_PLAYERS, false)
Public currentMultiOptionsTitleUI As Long        ' std::weak_ptr<WzMultiplayerOptionsTitleUI>: its ObjPtr (0: expired)
Public hideTime As Long                         ' UDWORD
Private perfFrequency As Currency                ' (VB) QueryPerformanceFrequency: steady_clock durations in seconds

' lib/ivis_opengl/pieblitfunc.h (WzGfx's)
Private Const BACKDROP_HACK_WIDTH As Long = 512
Private Const BACKDROP_HACK_HEIGHT As Long = 512

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

' (VB) sstrcpy into a char[size]
Public Function sstrcpy_cut(ByVal src As String, ByVal size As Long) As String
    sstrcpy_cut = Left$(src, size - 1)
End Function

' (VB) get<uint8_t>(): the low 8 bits
Private Function get_uint8(ByVal j As nlohmann_json) As Long
    get_uint8 = j.get_uint() And &HFF&
End Function

' (VB) get<int8_t>(): the low 8 bits, sign extended
Private Function get_int8(ByVal j As nlohmann_json) As Long
    Dim v As Long
    v = j.get_int() And &HFF&
    If v >= 128 Then v = v - 256
    get_int8 = v
End Function

' (VB) the type_error from_json raises for a std::vector from a value that is no array
Private Sub json_throw_not_array(ByVal j As nlohmann_json)
    Err.Raise vbObjectError + 302, "nlohmann::json", "[json.exception.type_error.302] type must be array, but is " & j.type_name()
End Sub

' NetPlay.players[i] as a PLAYER_
Public Function NetPlay_players_get(ByVal i As Long) As PLAYER_
    Dim p As PLAYER_
    p.name_ = NetPlay_players_name(i)
    p.position = NetPlay_players_position(i)
    p.colour = NetPlay_players_colour(i)
    p.allocated = NetPlay_players_allocated(i)
    p.heartattacktime = NetPlay_players_heartattacktime(i)
    p.heartbeat = NetPlay_players_heartbeat(i)
    p.team = NetPlay_players_team(i)
    p.ready = NetPlay_players_ready(i)
    p.ai = NetPlay_players_ai(i)
    p.difficulty = NetPlay_players_difficulty(i)
    p.faction = NetPlay_players_faction(i)
    p.isSpectator = NetPlay_players_isSpectator(i)
    p.isAdmin = NetPlay_players_isAdmin(i)
    NetPlay_players_get = p
End Function

' NetPlay.players[i]'s fields from a PLAYER_ (the others are left as they are)
Public Sub NetPlay_players_set(ByVal i As Long, p As PLAYER_)
    NetPlay_players_name(i) = p.name_
    NetPlay_players_position(i) = p.position
    NetPlay_players_colour(i) = p.colour
    NetPlay_players_allocated(i) = p.allocated
    NetPlay_players_heartattacktime(i) = p.heartattacktime
    NetPlay_players_heartbeat(i) = p.heartbeat
    NetPlay_players_team(i) = p.team
    NetPlay_players_ready(i) = p.ready
    NetPlay_players_ai(i) = p.ai
    NetPlay_players_difficulty(i) = p.difficulty
    NetPlay_players_faction(i) = p.faction
    NetPlay_players_isSpectator(i) = p.isSpectator
    NetPlay_players_isAdmin(i) = p.isAdmin
End Sub

Public Sub Sha256_to_json(j As nlohmann_json, p As Sha256)
    Set j = json_value(Sha256_toString(p))
End Sub

Public Sub Sha256_from_json(j As nlohmann_json, p As Sha256)
    Dim str_ As String
    str_ = j.get_string()
    Sha256_fromString p, str_
End Sub

Public Sub MULTIPLAYERGAME_to_json(j As nlohmann_json, p As MULTIPLAYERGAME)
    Dim h As nlohmann_json, hashes As nlohmann_json, i As Long
    Set j = json_object()
    j.set_item "type", json_value(p.typ)
    j.set_item "scavengers", json_value_unsigned(p.scavengers)
    j.set_item "map", json_value(p.map)                ' Wrap this in WzString to handle invalid UTF-8 before adding to json object
    j.set_item "maxPlayers", json_value_unsigned(p.maxPlayers)
    j.set_item "name", json_value(p.name_)             ' Wrap this in WzString to handle invalid UTF-8 before adding to json object
    Sha256_to_json h, p.hash
    j.set_item "hash", h
    Set hashes = json_array()
    For i = 0 To p.modHashCount - 1
        Sha256_to_json h, p.modHashes(i)
        hashes.push_back h
    Next
    j.set_item "modHashes", hashes
    j.set_item "power", json_value_unsigned(p.power)
    j.set_item "base", json_value_unsigned(p.base)
    j.set_item "alliance", json_value_unsigned(p.alliance)
    j.set_item "mapHasScavengers", json_value(p.mapHasScavengers)
    j.set_item "isMapMod", json_value(p.isMapMod)
    j.set_item "isRandom", json_value(p.isRandom)
    j.set_item "techLevel", json_value_unsigned(p.techLevel)
    j.set_item "inactivityMinutes", json_value_unsigned(p.inactivityMinutes)
    j.set_item "gameTimeLimitMinutes", json_value_unsigned(p.gameTimeLimitMinutes)
    j.set_item "playerLeaveMode", json_value(p.playerLeaveMode)
    j.set_item "blindMode", json_value(p.blindMode)
End Sub

Public Sub MULTIPLAYERGAME_from_json(j As nlohmann_json, p As MULTIPLAYERGAME)
    Dim str_ As String, hashes As nlohmann_json, i As Long
    p.typ = get_uint8(j.at("type"))                     ' LEVEL_TYPE: uint8_t
    p.scavengers = get_uint8(j.at("scavengers"))
    str_ = j.at("map").get_string()
    p.map = sstrcpy_cut(str_, 128)
    p.maxPlayers = get_uint8(j.at("maxPlayers"))
    str_ = j.at("name").get_string()
    p.name_ = sstrcpy_cut(str_, 128)
    Sha256_from_json j.at("hash"), p.hash
    ' get<std::vector<Sha256>>()
    Set hashes = j.at("modHashes")
    If Not hashes.is_array() Then json_throw_not_array hashes
    p.modHashCount = hashes.size()
    If p.modHashCount > 0 Then
        ReDim p.modHashes(0 To p.modHashCount - 1)
    Else
        Erase p.modHashes
    End If
    For i = 0 To p.modHashCount - 1
        Sha256_from_json hashes.at_idx(i), p.modHashes(i)
    Next
    p.power = j.at("power").get_uint()
    p.base = get_uint8(j.at("base"))
    p.alliance = get_uint8(j.at("alliance"))
    p.mapHasScavengers = j.at("mapHasScavengers").get_bool()
    p.isMapMod = j.at("isMapMod").get_bool()
    p.isRandom = j.at("isRandom").get_bool()
    p.techLevel = j.at("techLevel").get_uint()
    If j.contains("inactivityMinutes") Then
        p.inactivityMinutes = j.at("inactivityMinutes").get_uint()
    Else
        ' default to the old (4.2.0 beta-era) value of 4 minutes
        p.inactivityMinutes = 4
    End If
    If j.contains("gameTimeLimitMinutes") Then
        p.gameTimeLimitMinutes = j.at("gameTimeLimitMinutes").get_uint()
    Else
        ' default to the old (pre-4.4.0) default value of 0 minutes (disabled)
        p.gameTimeLimitMinutes = 0
    End If
    If j.contains("playerLeaveMode") Then
        p.playerLeaveMode = get_uint8(j.at("playerLeaveMode"))    ' PLAYER_LEAVE_MODE: uint8_t
    Else
        ' default to the old (pre-4.4.0) behavior of destroy resources
        p.playerLeaveMode = 0                           ' PLAYER_LEAVE_MODE::DESTROY_RESOURCES
    End If
    If j.contains("blindMode") Then
        p.blindMode = get_uint8(j.at("blindMode"))      ' BLIND_MODE: uint8_t
    Else
        ' default to the old (pre-4.6.0) behavior (no blind mode)
        p.blindMode = BLIND_MODE_NONE
    End If
End Sub

Public Sub MULTISTRUCTLIMITS_to_json(j As nlohmann_json, p As MULTISTRUCTLIMITS)
    Set j = json_object()
    j.set_item "id", json_value_unsigned(p.id)
    j.set_item "limit", json_value_unsigned(p.limit)
End Sub

Public Sub MULTISTRUCTLIMITS_from_json(j As nlohmann_json, p As MULTISTRUCTLIMITS)
    p.id = j.at("id").get_uint()
    p.limit = j.at("limit").get_uint()
End Sub

Public Sub InGameSide_to_json(j As nlohmann_json, ByVal p As Boolean)
    Set j = json_value(p)
End Sub

Public Sub InGameSide_from_json(j As nlohmann_json, p As Boolean)
    Dim val_ As Boolean
    val_ = j.get_bool()
    p = val_
End Sub

Public Sub MULTIPLAYERINGAME_to_json(j As nlohmann_json, p As MULTIPLAYERINGAME)
    Dim s As nlohmann_json, limits As nlohmann_json, i As Long
    Set j = json_object()

'   auto joiningInProgress = nlohmann::json::array();
'   for (size_t idx = 0; idx < MAX_CONNECTED_PLAYERS; idx++)
'   {
'       joiningInProgress.push_back(p.JoiningInProgress[idx]);
'   }
'   j["JoiningInProgress"] = joiningInProgress;
    InGameSide_to_json s, p.side
    j.set_item "side", s
    Set limits = json_array()
    For i = 0 To p.lastAppliedStructureLimitsCount - 1
        MULTISTRUCTLIMITS_to_json s, p.lastAppliedStructureLimits(i)
        limits.push_back s
    Next
    j.set_item "structureLimits", limits   ' See applyLimitSet() for why we save `lastAppliedStructureLimits` as `structureLimits`
    j.set_item "flags", json_value_unsigned(p.flags)
End Sub

Public Sub MULTIPLAYERINGAME_from_json(j As nlohmann_json, p As MULTIPLAYERINGAME)
    Dim limits As nlohmann_json, i As Long
'   auto joiningInProgress = j["JoiningInProgress"];
'   ASSERT(joiningInProgress.is_array(), "joiningInProgress should be an array");
'   for (size_t idx = 0; idx < joiningInProgress.size(); idx++)
'   {
'       if (idx >= MAX_CONNECTED_PLAYERS)
'       {
'           debug(LOG_ERROR, "Unexpected size: %zu", joiningInProgress.size());
'           break;
'       }
'       p.JoiningInProgress[idx] = joiningInProgress[idx].get<bool>();
'   }
    InGameSide_from_json j.at("side"), p.side
    ' get<std::vector<MULTISTRUCTLIMITS>>()
    Set limits = j.at("structureLimits")
    If Not limits.is_array() Then json_throw_not_array limits
    p.structureLimitsCount = limits.size()
    If p.structureLimitsCount > 0 Then
        ReDim p.structureLimits(0 To p.structureLimitsCount - 1)
    Else
        Erase p.structureLimits
    End If
    For i = 0 To p.structureLimitsCount - 1
        MULTISTRUCTLIMITS_from_json limits.at_idx(i), p.structureLimits(i)
    Next
    p.flags = get_uint8(j.at("flags"))
End Sub

Public Sub PLAYER_to_json_struct(j As nlohmann_json, p As PLAYER_)
    Set j = json_object()
    j.set_item "name", json_value(p.name_)             ' Wrap this in WzString to handle invalid UTF-8 before adding to json object
    j.set_item "position", json_value(p.position)
    j.set_item "colour", json_value(p.colour)
    j.set_item "allocated", json_value(p.allocated)
    j.set_item "heartattacktime", json_value_unsigned(p.heartattacktime)
    j.set_item "heartbeat", json_value(p.heartbeat)
    j.set_item "kick", json_value(False)               ' Store fixed false value for now (PLAYER.kick removed in WZ 4.6.2)
    j.set_item "team", json_value(p.team)
    j.set_item "ready", json_value(p.ready)
    j.set_item "ai", json_value(p.ai)
    j.set_item "difficulty", json_value(p.difficulty)  ' static_cast<int8_t>
    'j["autoGame"] = p.autoGame; // MAYBE NOT?
    ' Do not persist wzFiles
    ' Do not persist IPtextAddress
    j.set_item "faction", json_value_unsigned(p.faction)   ' static_cast<uint8_t>
    j.set_item "isSpectator", json_value(p.isSpectator)
    j.set_item "isAdmin", json_value(p.isAdmin)
End Sub

Public Sub PLAYER_from_json_struct(j As nlohmann_json, p As PLAYER_)
    Dim str_ As String, difficultyInt As Long, factionUint As Long
    str_ = j.at("name").get_string()
    p.name_ = sstrcpy_cut(str_, StringSize)
    p.position = j.at("position").get_int()
    p.colour = j.at("colour").get_int()
    p.allocated = j.at("allocated").get_bool()
    p.heartattacktime = j.at("heartattacktime").get_uint()
    p.heartbeat = j.at("heartbeat").get_bool()
    p.team = j.at("team").get_int()
    p.ready = j.at("ready").get_bool()
    p.ai = get_int8(j.at("ai"))
    difficultyInt = get_int8(j.at("difficulty"))
    p.difficulty = difficultyInt                        ' TODO CHECK
    ' autoGame // MAYBE NOT?
    ' Do not persist wzFiles: (VB) no downloads are kept, so there is nothing to clear
    ' Do not persist IPtextAddress
    factionUint = get_uint8(j.at("faction"))
    p.faction = factionUint                             ' TODO CHECK
    p.isSpectator = j.at("isSpectator").get_bool()
    If j.contains("isAdmin") Then
        p.isAdmin = j.at("isAdmin").get_bool()
    Else
        ' default to the old (pre-4.5.4) default value of false
        p.isAdmin = False
    End If
End Sub

' to_json(j, NetPlay.players[i])
Public Sub PLAYER_to_json(j As nlohmann_json, ByVal i As Long)
    Dim p As PLAYER_
    p = NetPlay_players_get(i)
    PLAYER_to_json_struct j, p
End Sub

' from_json(j, NetPlay.players[i])
Public Sub PLAYER_from_json(j As nlohmann_json, ByVal i As Long)
    Dim p As PLAYER_
    p = NetPlay_players_get(i)
    PLAYER_from_json_struct j, p
    NetPlay_players_set i, p
End Sub

Public Function DataHashToJSON() As nlohmann_json
    Dim jsonArray As nlohmann_json, i As Long
    Set jsonArray = json_array()
    For i = 0 To DATA_MAXDATA - 1
        jsonArray.push_back json_value_unsigned(DataHash(i))
    Next
    Set DataHashToJSON = jsonArray
End Function

'sets sWRFILE form game.map (static in C++)
Public Sub decideWRF()
    ' try and load it from the maps directory first,
    aLevelName = Left$(MultiCustomMapsPath, MAX_LEVEL_NAME_SIZE)
    aLevelName = Left$(aLevelName & game.map, MAX_LEVEL_NAME_SIZE)
    aLevelName = Left$(aLevelName & ".wrf", MAX_LEVEL_NAME_SIZE)
    debug_ 1012, LOG_WZ, "decideWRF", "decideWRF: %s", aLevelName
    'if the file exists in the downloaded maps dir then use that one instead.
    ' FIXME: Try to incorporate this into physfs setup somehow for sane paths
    If PHYSFS_exists(aLevelName) = 0 Then
        aLevelName = Left$(game.map, MAX_LEVEL_NAME_SIZE)  ' doesn't exist, must be a predefined one.
    End If
End Sub

Public Sub setDefaultSkirmishAI(ByVal name_ As String)
    defaultSkirmishAI = name_
End Sub

Public Function getDefaultSkirmishAI(Optional ByVal displayNameOnly As Boolean = False) As String
    Dim fallbackAI As String, fallbackIdx As Long, name_ As String, i As Long
    fallbackAI = DEFAULT_SKIRMISH_AI_SCRIPT_NAME
    fallbackIdx = 0
    name_ = ""

    For i = 0 To aidata_size - 1
        If aidata(i).js = fallbackAI Then
            fallbackIdx = i
        End If
        If aidata(i).js = defaultSkirmishAI Then
            If displayNameOnly Then name_ = aidata(i).name_ Else name_ = aidata(i).js
        End If
    Next

    ' Something weird happened, config messed with, or mod AI missing?
    If Len(name_) = 0 And aidata_size > 0 Then
        setDefaultSkirmishAI aidata(fallbackIdx).js
        If displayNameOnly Then name_ = aidata(fallbackIdx).name_ Else name_ = aidata(fallbackIdx).js
    End If

    getDefaultSkirmishAI = name_
End Function


' const std::vector<AIDATA>& getAIData(): the vector's size and an element's js field (aidata is Private)
Public Function getAIData_size() As Long
    getAIData_size = aidata_size
End Function

Public Function getAIData_js(ByVal i As Long) As String
    getAIData_js = aidata(i).js
End Function

' (VB) getAIData()[i].name, .tip, .difficultyTips[d], .assigned
Public Function getAIData_name(ByVal i As Long) As String
    getAIData_name = aidata(i).name_
End Function

Public Function getAIData_tip(ByVal i As Long) As String
    getAIData_tip = aidata(i).tip
End Function

Public Function getAIData_difficultyTip(ByVal i As Long, ByVal d As Long) As String
    getAIData_difficultyTip = aidata(i).difficultyTips(d)
End Function

Public Function getLockedOptions() As MultiplayOptionsLocked
    getLockedOptions = locked
End Function

' (displayRoomSystemMessage, displayRoomNotifyMessage: multiint_lobby_c's, with the chat box)

' static std::array<const char *, 5> difficultyList = { N_("Super Easy"), N_("Easy"), N_("Medium"), N_("Hard"), N_("Insane") };
Public Function difficultyList_at(ByVal j As Long) As String
    Select Case j
    Case 0: difficultyList_at = "Super Easy"
    Case 1: difficultyList_at = "Easy"
    Case 2: difficultyList_at = "Medium"
    Case 3: difficultyList_at = "Hard"
    Case 4: difficultyList_at = "Insane"
    End Select
End Function

' static const AIDifficulty difficultyValue[] = { SUPEREASY, EASY, MEDIUM, HARD, INSANE };
Public Function difficultyValue_at(ByVal j As Long) As Long
    Select Case j
    Case 0: difficultyValue_at = AIDifficulty_SUPEREASY
    Case 1: difficultyValue_at = AIDifficulty_EASY
    Case 2: difficultyValue_at = AIDifficulty_MEDIUM
    Case 3: difficultyValue_at = AIDifficulty_HARD
    Case 4: difficultyValue_at = AIDifficulty_INSANE
    End Select
End Function

Public Function getAIName(ByVal player As Long) As String
    If NetPlay_players_ai(player) >= 0 And NetPlay_players_ai(player) <> AI_CUSTOM Then
        If Not (NetPlay_players_ai(player) < aidata_size) Then ASSERT_FAILURE "NetPlay.players[player].ai < aidata.size()", "src/multiint.cpp:491", "getAIName", "Invalid AI (index: %i, num AIs: %zu)", NetPlay_players_ai(player), aidata_size: getAIName = "Commander": Exit Function
        getAIName = aidata(NetPlay_players_ai(player)).name_
    Else
        getAIName = "Commander"
    End If
End Function

Public Sub loadMultiScripts()
    Dim defaultRules As Boolean, aFileName As String, aPathName As String, psLevel As LEVEL_DATASET
    Dim ininame As String, path As String, loadExtra As Boolean, it As Long, ini As WzConfig, i As Long
    Dim bIsLocalPlayer As Boolean, bShouldSkipThisPlayer As Boolean, ai As Long
    defaultRules = True
    Set psLevel = levFindDataSet_hash(game.map, game.hash, True)
    If psLevel Is Nothing Then ASSERT_FAILURE "psLevel", "src/multiint.cpp:506", "loadMultiScripts", "No level found for %s", game.map: Exit Sub
    If Not (psLevel.game_ >= 0 And psLevel.game_ < LEVEL_MAXFILES) Then ASSERT_FAILURE "psLevel->game >= 0 && psLevel->game < LEVEL_MAXFILES", "src/multiint.cpp:507", "loadMultiScripts", "Invalid psLevel->game: %i - may be a corrupt level load (%s; hash: %s)", psLevel.game_, game.map, Sha256_toString(game.hash): Exit Sub
    aFileName = sstrcpy_cut(psLevel.apDataFiles(psLevel.game_), 256)
    aFileName = Left$(aFileName, Len(aFileName) - 4)   ' aFileName[strlen(aFileName) - 4] = '\0'
    aPathName = aFileName
    aFileName = sstrcpy_cut(aFileName & ".json", 256)
    aPathName = sstrcpy_cut(aPathName & "/", 256)
    loadExtra = False

    If Len(challengeFileName) > 0 Then
        ininame = challengeFileName
        path = "challenges/"
        loadExtra = True
    End If

    If getHostLaunch() = HostLaunch_Skirmish Then
        ininame = "tests/" & wz_skirmish_test()
        path = "tests/"
        loadExtra = True
    End If

    If getHostLaunch() = HostLaunch_Autohost Then
        ininame = "autohost/" & wz_skirmish_test()
        path = "autohost/"
        loadExtra = True
    End If

    ' Reset assigned counter
    For it = 0 To aidata_size - 1
        aidata(it).assigned = 0
    Next

    ' Load map scripts
    If loadExtra Then
        If PHYSFS_exists(ininame) <> 0 Then
            Set ini = WzConfig_new(ininame, WzConfig_ReadOnly)
            debug_ 548, LOG_SAVE, "loadMultiScripts", "Loading map scripts"
            If ini.beginGroup("scripts") Then
                If ini.contains("extra") Then
                    loadGlobalScript path & ini.value("extra").toWzString()
                End If
                If ini.contains("rules") Then
                    loadGlobalScript path & ini.value("rules").toWzString()
                    defaultRules = False
                End If
            End If
            ini.endGroup
        End If
    End If

    ' Load multiplayer rules
    resForceBaseDir "messages/"
    resLoadFile "SMSG", "multiplay.txt"
    If defaultRules Then
        debug_ 569, LOG_SAVE, "loadMultiScripts", "Loading default rules"
        loadGlobalScript "multiplay/script/rules/init.js"
    End If

    ' Load AI players for skirmish games
    resForceBaseDir "multiplay/skirmish/"
    If bMultiPlayer And game.typ = LEVEL_TYPE_SKIRMISH Then
        For i = 0 To game.maxPlayers - 1
            ' Skip human players. Do not skip local player for autogames
            bIsLocalPlayer = i = selectedPlayer
            bShouldSkipThisPlayer = Not bIsLocalPlayer Or Not autogame_enabled()
            If NetPlay_players_allocated(i) And bShouldSkipThisPlayer Then
                GoTo continue_
            End If

            ' Make sure local player has an AI in autogames
            If bIsLocalPlayer And autogame_enabled() Then  ' (/*NetPlay.players[i].ai < 0 &&*/)
                'NetPlay.players[i].ai = 0;
                ' levLoadData handles making sure the local player has an AI in autogames
                ' TODO: clean this mess up, and pick one place to handle initializing the AI in autogames
                GoTo continue_
            End If

            If NetPlay_players_ai(i) >= 0 Then
                If myResponsibility(i) Then
                    ai = NetPlay_players_ai(i)
                    If Len(aidata(ai).js) > 0 Then
                        debug_ 600, LOG_SAVE, "loadMultiScripts", "Loading javascript AI for player %d", i
                        If loadPlayerScript("multiplay/skirmish/" & aidata(ai).js, i, NetPlay_players_difficulty(i)) Is Nothing Then
                            debug_ 603, LOG_ERROR, "loadMultiScripts", "Failed to load AI!: %s", aidata(ai).js
                        End If
                    End If
                End If
            End If
continue_:
        Next
    End If

    ' Load scavengers
    If game.scavengers <> NO_SCAVENGERS Then
        If myResponsibility(scavengerPlayer()) Then
            debug_ 613, LOG_SAVE, "loadMultiScripts", "Loading scavenger AI for player %d", scavengerPlayer()
            loadPlayerScript "multiplay/script/scavengers/init.js", scavengerPlayer(), AIDifficulty_EASY
        End If
    End If

    ' Reset resource path, otherwise things break down the line
    resForceBaseDir ""
End Sub

' multiplaydefs.h: static inline bool isBlindSimpleLobby(BLIND_MODE mode)
Public Function isBlindSimpleLobby(ByVal mode As Long) As Boolean
    isBlindSimpleLobby = mode = BLIND_MODE_BLIND_LOBBY_SIMPLE_LOBBY Or mode = BLIND_MODE_BLIND_GAME_SIMPLE_LOBBY
End Function

Private Sub loadEmptyMapPreview(Optional ByVal withQuestionMarks As Boolean = True)
    Const ex As Long = 100, ey As Long = 100
    Const bx As Long = 5, by As Long = 8
    Dim bckImage As iV_Image, imageData As Long, n As Long, sx As Long, sy As Long, col(0 To 2) As Byte
    Dim x As Long, y As Long, mark(0 To 4) As Long, i As Long, playerpos_(0 To MAX_PLAYERS - 1) As Vector2i
    ' No map is available to preview, so improvise.
    Set bckImage = New iV_Image
    If Not bckImage.allocate(BACKDROP_HACK_WIDTH, BACKDROP_HACK_HEIGHT, 3, True) Then
        debug_ 656, LOG_ERROR, "loadEmptyMapPreview", "Failed to allocate memory for map preview"
        Exit Sub
    End If
    imageData = bckImage.bmp_w()
    ' memset(imageData, 0, BACKDROP_HACK_WIDTH * BACKDROP_HACK_HEIGHT * 3); //dunno about background color (allocate's zeroMemory)
    If withQuestionMarks Then
        mark(0) = 2: mark(1) = 1: mark(2) = &HB1: mark(3) = 9: mark(4) = 6   ' "\2\1\261\11\6"
        For n = 0 To 124
            sx = rand() Mod (ex - bx)
            sy = rand() Mod (ey - by)
            col(0) = rand() Mod 256
            col(1) = rand() Mod 256
            col(2) = rand() Mod 256
            For y = 0 To by - 1
                For x = 0 To bx - 1
                    If ((mark(x) \ (2 ^ y)) And 1) = 1 Then ' ?
                        CopyMemory ByVal imageData + 3 * (sx + x + BACKDROP_HACK_WIDTH * (sy + y)), col(0), 3
                    End If
                Next
            Next
        Next
    End If

    ' Slight hack to init array with a special value used to determine how many players on map
    For i = 0 To MAX_PLAYERS - 1
        playerpos_(i).x = &H77777777
        playerpos_(i).y = &H77777777
    Next

    screen_enableMapPreview ex, ey, VarPtr(playerpos_(0))

    screen_Upload bckImage
End Sub

' (VB) the map preview's game state: game.mapHasScavengers, builtInMap, useTerrainOverrides, the search path,
' the map's load (gameLoadDetails.getMap(rand()), which takes a rand()) and hideTime. The preview itself
' (generate2DMapPreview, the backdrop image, the HQ positions, screen_enableMapPreview) is the front end's
' display, cluster F.
Public Sub loadMapPreview(ByVal hideInterface As Boolean)
    Dim aFileName As String, psLevel As LEVEL_DATASET, gameLoadDetails As GameLoadDetails, data As Map
    Dim h As Sha256, realFileHash As Sha256, realDir As String, bckImage As iV_Image, i As Long, md As MapData
    Dim playerpos_(0 To MAX_PLAYERS - 1) As Vector2i

    ' absurd hack, since there is a problem with updating this crap piece of info, we're setting it to
    ' true by default for now, like it used to be
    ' TODO: Is this still needed at all?
    game.mapHasScavengers = True ' this is really the wrong place for it, but this is where it has to be

    If isBlindSimpleLobby(game.blindMode) And Not NetPlay_isHost Then
        loadEmptyMapPreview False
        Exit Sub
    End If

    ' load the terrain types
    h = game.hash
    Set psLevel = levFindDataSet_hash(game.map, h, True)
    If psLevel Is Nothing Then
        ' (NET_getDownloadingWzFiles(): the downloads come with cluster J, there are none)
        debug_ 764, LOG_INFO, "loadMapPreview", "Could not find level dataset ""%s"" %s. We %s waiting for a download.", game.map, Sha256_toString(game.hash), "aren't"
        loadEmptyMapPreview
        Exit Sub
    End If
    If psLevel.game_ < 0 Or psLevel.game_ >= LEVEL_MAXFILES Then
        debug_ 770, LOG_ERROR, "loadMapPreview", "apDataFiles index (%i) is out of bounds for: ""%s"" %s.", psLevel.game_, game.map, Sha256_toString(game.hash)
        loadEmptyMapPreview
        Exit Sub
    End If
    If StrPtr(psLevel.realFileName) = 0 Then
        builtInMap = True
        useTerrainOverrides = shouldLoadTerrainTypeOverrides(psLevel.pName)
        debug_ 778, LOG_WZ, "loadMapPreview", "Loading map preview: ""%s"" builtin t%d override %d", psLevel.pName, psLevel.dataDir_, -CLng(useTerrainOverrides)
    Else
        builtInMap = False
        useTerrainOverrides = False
        realFileHash = psLevel.realFileHash
        debug_ 784, LOG_WZ, "loadMapPreview", "Loading map preview: ""%s"" in (%s)""%s""  %s t%d", psLevel.pName, PHYSFS_getRealDir(psLevel.realFileName), psLevel.realFileName, Sha256_toString(realFileHash), psLevel.dataDir_
    End If
    rebuildSearchPath psLevel.dataDir_, False
    If Len(psLevel.apDataFiles(psLevel.game_)) = 0 Then
        If StrPtr(psLevel.realFileName) <> 0 Then realDir = psLevel.realFileName Else realDir = "null"
        debug_ 789, LOG_ERROR, "loadMapPreview", "No path for level ""%s""? (%s)", psLevel.pName, realDir
        loadEmptyMapPreview
        Exit Sub
    End If
    aFileName = psLevel.apDataFiles(psLevel.game_)

    If StrPtr(psLevel.realFileName) <> 0 Then
        Set gameLoadDetails = GameLoadDetails_makeMapPackageLoad(psLevel.realFileName)
    Else
        Set gameLoadDetails = GameLoadDetails_makeLevelFileLoad(aFileName)
    End If

    ' load the map data
    Set data = gameLoadDetails.getMap(rand())
    If data Is Nothing Then
        debug_ 801, LOG_ERROR, "loadMapPreview", "Failed to load map from path: %s", aFileName
        loadEmptyMapPreview
        Exit Sub
    End If

    ' (VB) WzMap has no generate2DMapPreview (lib/wzmaplib's map_preview.cpp is not ported): the preview's
    ' colour scheme (guessMapTilesetType), the generated image and the players' HQ positions are left out,
    ' the backdrop is the blank BACKDROP_HACK_WIDTH x BACKDROP_HACK_HEIGHT image the C++ copies the preview
    ' into (no rand() calls, as on the C++'s success path), with the playerpos markers of no players
    Set md = data.mapData()
    If md Is Nothing Then
        ' Failed to generate map preview
        debug_ 840, LOG_ERROR, "loadMapPreview", "Failed to generate map preview for: %s", psLevel.pName
        loadEmptyMapPreview
        Exit Sub
    End If
    Set bckImage = New iV_Image
    If Not bckImage.allocate(BACKDROP_HACK_WIDTH, BACKDROP_HACK_HEIGHT, 3, True) Then
        debug_ 849, LOG_FATAL, "loadMapPreview", "Failed to allocate memory for map preview"
        Exit Sub    ' (abort();	// should be a fatal error ?)
    End If
    For i = 0 To MAX_PLAYERS - 1
        playerpos_(i).x = &H77777777
        playerpos_(i).y = &H77777777
    Next

    screen_enableMapPreview md.width, md.height, VarPtr(playerpos_(0))

    screen_Upload bckImage

    If hideInterface Then
        hideTime = gameTime
    End If
End Sub

Public Sub readAIs()
    Dim basepath As String, sSearchPath As String, files() As String, nFiles As Long, f As Long, path As String
    Dim aiconf As WzConfig, ai As AIDATA_, blank As AIDATA_, difficultyKeys As Variant, i As Long, k As Long
    Dim wins As Long, losses As Long, draws As Long, total As Long, statistics As String
    Dim win_percentage As Single, loss_percentage As Single, draw_percentage As Single
    sSearchPath = "multiplay/skirmish/"

    Erase aidata
    aidata_size = 0

    basepath = sstrcpy_cut(sSearchPath, PATH_MAX)
    WZ_PHYSFS_enumerateFiles basepath, files, nFiles
    For f = 0 To nFiles - 1
        ' See if this filename contains the extension we're looking for
        If InStr(1, files(f), ".json", vbBinaryCompare) = 0 Then
            GoTo continue_ ' continue;
        End If
        path = sstrcpy_cut(basepath & files(f), PATH_MAX)
        Set aiconf = WzConfig_new(path, WzConfig_ReadOnly)
        ai = blank
        aiconf.beginGroup "AI"

        If aiconf.contains("name") Then
            ai.name_ = sstrcpy_cut(aiconf.value("name", json_variant_new("")).toWzString(), MAX_LEN_AI_NAME)
        Else
            ai.name_ = sstrcpy_cut("MISSING AI NAME", MAX_LEN_AI_NAME)
        End If

        ai.js = sstrcpy_cut(aiconf.value("js", json_variant_new("")).toWzString(), MAX_LEN_AI_NAME)

        difficultyKeys = Array("supereasy_dummy_tip", "easy_tip", "medium_tip", "hard_tip", "insane_tip")
        For i = 0 To UBound(difficultyKeys)
            If aiconf.contains(difficultyKeys(i)) Then
                ai.difficultyTips(i) = sstrcpy_cut(aiconf.value(difficultyKeys(i), json_variant_new("")).toWzString(), 255)
            Else
                ' note that the empty string "" must never be translated
                ai.difficultyTips(i) = ""
            End If
        Next

        If aiconf.contains("tip") Then
            ai.tip = sstrcpy_cut(aiconf.value("tip", json_variant_new("")).toWzString(), 255 + 128)
        Else
            ai.tip = sstrcpy_cut("MISSING AI DESCRIPTION", 255 + 128)
        End If

        wins = aiconf.value("wins", json_variant_new(0)).toInt()
        losses = aiconf.value("losses", json_variant_new(0)).toInt()
        draws = aiconf.value("draws", json_variant_new(0)).toInt()
        total = wins + losses + draws
        If total <> 0 Then
            win_percentage = fmul(fdiv(CSng(wins), CSng(total)), 100!)
            loss_percentage = fmul(fdiv(CSng(losses), CSng(total)), 100!)
            draw_percentage = fmul(fdiv(CSng(draws), CSng(total)), 100!)
            ai.tip = sstrcpy_cut(ai.tip & vbLf, 255 + 128)
            statistics = sstrcpy_cut(va("AI tournament: %3.1f%% wins, %3.1f%% losses, %3.1f%% draws", CDbl(win_percentage), CDbl(loss_percentage), CDbl(draw_percentage)), 127)
            ai.tip = sstrcpy_cut(ai.tip & statistics, 255 + 128)
        End If

        If ai.js = defaultSkirmishAI Then
            ' aidata.insert(aidata.begin(), ai)
            ReDim Preserve aidata(0 To aidata_size)
            For k = aidata_size To 1 Step -1
                aidata(k) = aidata(k - 1)
            Next
            aidata(0) = ai
            aidata_size = aidata_size + 1
        Else
            ReDim Preserve aidata(0 To aidata_size)
            aidata(aidata_size) = ai
            aidata_size = aidata_size + 1
        End If
        aiconf.endGroup
continue_:
    Next
End Sub

Public Sub updateStructureDisabledFlags()
    Dim flags As Long, i As Long, stat As Long, disabled As Boolean
    ' The host works out the flags.
    If ingame.side <> InGameSide_HOST_OR_SINGLEPLAYER Then
        Exit Sub
    End If

    flags = ingame.flags And MPFLAGS_FORCELIMITS

    ' (static_assert(MPFLAGS_FORCELIMITS == (1 << (limitIcons.size() - 1))): 0x20 == 1 << 5)
    For i = 0 To limitIcons_size() - 2  ' skip last item, MPFLAGS_FORCELIMITS
        stat = getStructStatFromName(limitIcons(i).stat)
        disabled = False
        If stat >= 0 Then disabled = asStructureStats(stat).upgrade(0).limit = 0
        If disabled Then
            flags = flags Or BIT_(i)
        End If
    Next

    ingame.flags = flags
End Sub

Public Function isPlayerHostOrAdmin(ByVal playerIdx As Long) As Boolean
    If Not ULT(playerIdx, MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "playerIdx < MAX_CONNECTED_PLAYERS", "src/multiint.cpp:1289", "isPlayerHostOrAdmin", "Invalid player idx: %u", playerIdx: isPlayerHostOrAdmin = False: Exit Function
    isPlayerHostOrAdmin = (playerIdx = NetPlay_hostPlayer) Or NetPlay_players_isAdmin(playerIdx)
End Function

Public Function shouldInformOfAdminAction(ByVal targetPlayerIdx As Long, ByVal responsibleIdx As Long) As Boolean
    If Not ULT(targetPlayerIdx, MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "targetPlayerIdx < MAX_CONNECTED_PLAYERS", "src/multiint.cpp:1295", "shouldInformOfAdminAction", "Invalid targetPlayerIdx: %u", targetPlayerIdx: shouldInformOfAdminAction = False: Exit Function
    If Not ULT(responsibleIdx, MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "responsibleIdx < MAX_CONNECTED_PLAYERS", "src/multiint.cpp:1296", "shouldInformOfAdminAction", "Invalid responsibleIdx: %u", responsibleIdx: shouldInformOfAdminAction = False: Exit Function

    If responsibleIdx = targetPlayerIdx Then
        shouldInformOfAdminAction = False ' do not inform about self-action
        Exit Function
    End If
    If Not isPlayerHostOrAdmin(responsibleIdx) Then
        shouldInformOfAdminAction = False
        Exit Function
    End If
    If Not NetPlay_players_allocated(targetPlayerIdx) Then
        shouldInformOfAdminAction = False ' do not inform if target isn't a human player
        Exit Function
    End If

    shouldInformOfAdminAction = True
End Function

Private Function getPlayerTeam(ByVal i As Long) As Long
    If alliancesSetTeamsBeforeGame(game.alliance) Then getPlayerTeam = NetPlay_players_team(i) Else getPlayerTeam = i
End Function

' Checks if all players are on the same team. If so, return that team; if not, return -1;
' if there are no players, return team MAX_PLAYERS.
Public Function allPlayersOnSameTeam(ByVal except As Long) As Long
    Dim minTeam As Long, maxTeam As Long, i As Long, team As Long, c As Boolean
    minTeam = MAX_PLAYERS
    maxTeam = 0
    For i = 0 To game.maxPlayers - 1
        c = i <> except And (Not NetPlay_players_isSpectator(i))
        If c Then c = NetPlay_players_allocated(i) Or NetPlay_players_ai(i) >= 0
        If c Then
            team = getPlayerTeam(i)
            If team < minTeam Then minTeam = team
            If team > maxTeam Then maxTeam = team
        End If
    Next
    If minTeam = MAX_PLAYERS Or minTeam = maxTeam Then
        allPlayersOnSameTeam = minTeam  ' Players all on same team.
        Exit Function
    End If
    allPlayersOnSameTeam = -1  ' Players not all on same team.
End Function

Public Function sendReadyRequest(ByVal player As Long, ByVal bReady As Boolean) As Boolean
    Dim changedValue As Boolean, w As MessageWriter, p As Long, b As Boolean
    If NetPlay_isHost Then          ' do or request the change.
        changedValue = changeReadyStatus(player, bReady)
        If changedValue And wz_command_interface_enabled() Then
            ' (the "WZEVENT: readyStatus=%d: ..." line: the player's identity and IP address, cluster J)
            wz_command_interface_output_room_status_json
        End If
        sendReadyRequest = changedValue
        Exit Function
    Else
        Set w = NETbeginEncode(NETnetQueue(NetPlay_hostPlayer), NET_READY_REQUEST)
        p = player
        NETuint8_t w, p
        b = bReady
        NETbool w, b
        NETend w
    End If
    sendReadyRequest = True
End Function

Public Function changeReadyStatusImpl(ByVal player As Long, ByVal bReady As Boolean, ByVal skipBroadcastUpdate As Boolean) As Boolean
    Dim changedValue As Boolean, now_ As Currency
    changedValue = NetPlay_players_ready(player) <> bReady
    NetPlay_players_ready(player) = bReady
    If changedValue Then
        If bReady Then
            QueryPerformanceCounter now_
            If ingame.has_lastNotReadyTimes(player) Then
                ' accumulate time spent since the player last went "not ready"
                If perfFrequency = 0 Then QueryPerformanceFrequency perfFrequency
                ingame.secondsNotReady(player) = ingame.secondsNotReady(player) + Fix(CDbl(now_ - ingame.lastNotReadyTimes(player)) / CDbl(perfFrequency))
                ingame.has_lastNotReadyTimes(player) = False
            End If
            ingame.has_lastReadyTimes(player) = True
            ingame.lastReadyTimes(player) = now_
        Else
            If multiplayPlayersShouldCheckReady() Then
                ingame.has_lastNotReadyTimes(player) = True
                QueryPerformanceCounter ingame.lastNotReadyTimes(player)
            End If
        End If
    End If
    If Not skipBroadcastUpdate Then
        NETBroadcastPlayerInfo player
    End If
    netPlayersUpdated = True
    ' Player is fast! Clicked the "Ready" button before we had a chance to ping him/her
    ' change PingTime to some value less than PING_LIMIT, so that multiplayPlayersReady
    ' doesnt block
    If ingame.PingTimes(player) = PING_LIMIT Then ingame.PingTimes(player) = 1
    changeReadyStatusImpl = changedValue
End Function

Public Function changeReadyStatus(ByVal player As Long, ByVal bReady As Boolean) As Boolean
    changeReadyStatus = changeReadyStatusImpl(player, bReady, False)
End Function

Private Sub informIfAdminChangedOtherColor(ByVal targetPlayerIdx As Long, ByVal responsibleIdx As Long)
    If Not shouldInformOfAdminAction(targetPlayerIdx, responsibleIdx) Then
        Exit Sub
    End If

    sendQuickChat_data WzQuickChatMessage_INTERNAL_ADMIN_ACTION_NOTICE, realSelectedPlayer, WzQuickChatTargeting_targetAll(), WzQuickChatDataContexts_INTERNAL_ADMIN_ACTION_NOTICE_constructMessageData(WzQuickChatDataContexts_INTERNAL_ADMIN_ACTION_NOTICE_Context_Color, responsibleIdx, targetPlayerIdx)

    ' (the "Admin %s (%s) changed color of player ..." line with the admin's identity: the identities, cluster J)
End Sub

Public Function changeColour(ByVal player As Long, ByVal col As Long, ByVal responsibleIdx As Long) As Boolean
    Dim i As Long
    If col < 0 Or col >= MAX_PLAYERS_IN_GUI Then
        changeColour = False
        Exit Function
    End If

    If Not ULT(player, MAX_PLAYERS) Then
        changeColour = False
        Exit Function
    End If

    If getPlayerColour(player) = col Then
        changeColour = False  ' Nothing to do.
        Exit Function
    End If

    For i = 0 To MAX_PLAYERS - 1
        If getPlayerColour(i) = col Then
            If Not isPlayerHostOrAdmin(responsibleIdx) And NetPlay_players_allocated(i) Then
                changeColour = False  ' May not swap.
                Exit Function
            End If

            debug_ 2829, LOG_NET, "changeColour", "Swapping colours between players %d(%d) and %d(%d)", _
                  player, getPlayerColour(player), i, getPlayerColour(i)
            setPlayerColour i, getPlayerColour(player)
            setPlayerColour player, col
            NETBroadcastTwoPlayerInfo player, i
            informIfAdminChangedOtherColor player, responsibleIdx
            netPlayersUpdated = True
            changeColour = True
            Exit Function
        End If
    Next
    debug_ 2839, LOG_ERROR, "changeColour", "Failed to swap colours for player %d, colour %d", player, col
    If player < game.maxPlayers And col < MAX_PLAYERS Then
        ' Colours were corrupted. Attempt to fix.
        debug_ 2843, LOG_NET, "changeColour", "corrupted colours: player (%u) new colour (%u) old colour (%d)", player, col, NetPlay_players_colour(player)
        setPlayerColour player, col
        NETBroadcastPlayerInfo player
        informIfAdminChangedOtherColor player, responsibleIdx
        netPlayersUpdated = True
        changeColour = True
        Exit Function
    End If
    changeColour = False
End Function

' Notify all players of host launching the game
Private Sub SendFireUp()
    Dim randomSeed As Long, w As MessageWriter
    randomSeed = rand()  ' Pick a random random seed for the synchronised random number generator.

    debug_ 3808, LOG_INFO, "SendFireUp", "Sending NET_FIREUP"

    Set w = NETbeginEncode(NETbroadcastQueue(), NET_FIREUP)
    NETuint32_t w, randomSeed
    NETend w
    printSearchPath
    gameSRand randomSeed  ' Set the seed for the synchronised random number generator. The clients will use the same seed.
End Sub

' ////////////////////////////////////////////////////////////////////////////
' (static in C++; WzMultiplayerOptionsTitleUI.cls's startHost calls it)
Public Sub updateInActualHostedLobby(ByVal value As Boolean)
    Dim pChatBox As WIDGET
    bInActualHostedLobby = value
    Set pChatBox = widgGetFromID(psWScreen, MULTIOP_CHATBOX)
    If WIDGET_isKindOf(pChatBox, K_ChatBoxWidget) Then
        ChatBoxWidget_handleUpdateInHostedLobby pChatBox
    End If
End Sub

' mode_output: a BLIND_MODE
Public Function blindModeFromStr(ByVal str_ As String, mode_output As Long) As Boolean
    ' (const std::unordered_map<WzString, BLIND_MODE> mappings)
    Select Case str_
    Case "none"
        mode_output = BLIND_MODE_NONE
    Case "blind_lobby"
        mode_output = BLIND_MODE_BLIND_LOBBY
    Case "blind_lobby_simple_lobby"
        mode_output = BLIND_MODE_BLIND_LOBBY_SIMPLE_LOBBY
    Case "blind_game"
        mode_output = BLIND_MODE_BLIND_GAME
    Case "blind_game_simple_lobby"
        mode_output = BLIND_MODE_BLIND_GAME_SIMPLE_LOBBY
    Case Else
        mode_output = BLIND_MODE_NONE
        blindModeFromStr = False
        Exit Function
    End Select
    blindModeFromStr = True
End Function

Private Function loadMapChallengeSettings(ByVal ini As WzConfig) As Boolean
    Dim bIsAutoHostOrAutoGame As Boolean, mapName As String, mapHash As Sha256, mapData As LEVEL_DATASET
    Dim log_level As Long, configuredMaxPlayers As Long, openSpectatorSlots_uint As Long, blindMode_str As String
    Dim gamePass_str As String

    ini.beginGroup "locked" ' GUI lockdown
    locked.power = ini.value("power", json_variant_new(challengeActive)).toBool()
    locked.alliances = ini.value("alliances", json_variant_new(challengeActive)).toBool()
    locked.teams = ini.value("teams", json_variant_new(challengeActive)).toBool()
    locked.difficulty = ini.value("difficulty", json_variant_new(challengeActive)).toBool()
    locked.ai = ini.value("ai", json_variant_new(challengeActive)).toBool()
    locked.scavengers = ini.value("scavengers", json_variant_new(challengeActive)).toBool()
    locked.position = ini.value("position", json_variant_new(challengeActive)).toBool()
    locked.bases = ini.value("bases", json_variant_new(challengeActive)).toBool()
    locked.spectators = ini.value("spectators", json_variant_new(challengeActive)).toBool()
    locked.name_ = ini.value("name", json_variant_new(False)).toBool()
    locked.readybeforefull = ini.value("readybeforefull", json_variant_new(getHostLaunch() = HostLaunch_Autohost)).toBool()
    locked.cheats = ini.value("cheats", json_variant_new(False)).toBool()
    ini.endGroup

    bIsAutoHostOrAutoGame = getHostLaunch() = HostLaunch_Skirmish Or getHostLaunch() = HostLaunch_Autohost
    If challengeActive Or bIsAutoHostOrAutoGame Then
        ini.beginGroup "challenge"
        mapName = Left$(ini.value("map", json_variant_new(game.map)).toWzString(), 255)    ' char mapName[256]
        mapHash = levGetMapNameHash(mapName)

        Set mapData = levFindDataSet_hash(mapName, mapHash, True)
        If mapData Is Nothing Then
            If bIsAutoHostOrAutoGame Then log_level = LOG_ERROR Else log_level = LOG_FATAL
            debug_ 4880, log_level, "loadMapChallengeSettings", "Map %s not found!", mapName
            If wz_command_interface_enabled() Then
                ' (the "WZEVENT: mapNotFound: %s %s" line, with the name in base64: the command interface, cluster J)
            End If
            If bIsAutoHostOrAutoGame And headlessGameMode() Then
                wzQuit 1
            End If
            loadMapChallengeSettings = False
            Exit Function
        End If
        game.map = Left$(mapName, 127)
        game.hash = mapHash
        game.maxPlayers = mapData.players

        configuredMaxPlayers = ini.value("maxPlayers", json_variant_new(game.maxPlayers)).toUInt() And &HFF&   ' uint8_t
        If getHostLaunch() = HostLaunch_Autohost Then
            ' always use the autohost config - if it specifies an invalid number of players, this is a bug in the config
            ' however, maxPlayers must still be limited to MAX_PLAYERS
            If Not (configuredMaxPlayers < MAX_PLAYERS) Then ASSERT_FAILURE "configuredMaxPlayers < 11", "src/multiint.cpp:4901", "loadMapChallengeSettings", "Configured maxPlayers (%u) exceeds MAX_PLAYERS (%d)", configuredMaxPlayers, MAX_PLAYERS
            If configuredMaxPlayers > MAX_PLAYERS Then configuredMaxPlayers = MAX_PLAYERS
            If configuredMaxPlayers < 1 Then game.maxPlayers = 1 Else game.maxPlayers = configuredMaxPlayers
        Else
            If configuredMaxPlayers < 1 Then configuredMaxPlayers = 1
            If configuredMaxPlayers < game.maxPlayers Then game.maxPlayers = configuredMaxPlayers
        End If
        game.scavengers = ini.value("scavengers", json_variant_new(game.scavengers)).toInt() And &HFF&   ' uint8_t
        game.alliance = ini.value("alliances", json_variant_new(ALLIANCES_TEAMS)).toInt() And &HFF&
        game.power = ini.value("powerLevel", json_variant_new(game.power)).toInt()
        game.base = (ini.value("bases", json_variant_new(game.base + 1)).toInt() - 1) And &HFF&     ' count from 1 like the humans do
        game.name_ = Left$(ini.value("name").toWzString(), 127)
        game.techLevel = ini.value("techLevel", json_variant_new(game.techLevel)).toInt()

        ' Allow making the host a spectator (for MP games)
        spectatorHost = ini.value("spectatorHost", json_variant_new(False)).toBool()

        ' Allow configuring of open spectator slots (for MP games)
        openSpectatorSlots_uint = ini.value("openSpectatorSlots", json_variant_new(0)).toUInt()
        If ULT(openSpectatorSlots_uint, MAX_SPECTATOR_SLOTS) Then defaultOpenSpectatorSlots = openSpectatorSlots_uint Else defaultOpenSpectatorSlots = MAX_SPECTATOR_SLOTS

        ' Allow setting "blind mode"
        blindMode_str = ini.value("blindMode", json_variant_new("none")).toWzString()
        If Not blindModeFromStr(blindMode_str, game.blindMode) Then
            debug_ 4927, LOG_ERROR, "loadMapChallengeSettings", "Invalid blindMode (%s) specified in config .json - ignoring", blindMode_str
        End If

        ' Allow setting gamePassword
        gamePass_str = ini.value("gamePassword", json_variant_new("")).toWzString()
        If Len(gamePass_str) > 0 Then
            NETsetGamePassword gamePass_str
        End If

        ' DEPRECATED: This seems to have been odd workaround for not having the locked group handled.
        '             Keeping it around in case mods use it.
        locked.position = Not ini.value("allowPositionChange", json_variant_new(Not locked.position)).toBool()
        ini.endGroup
    Else
        ini.beginGroup "defaults"
        game.scavengers = ini.value("scavengers", json_variant_new(game.scavengers)).toInt() And &HFF&
        game.base = ini.value("bases", json_variant_new(game.base)).toInt() And &HFF&
        game.alliance = ini.value("alliances", json_variant_new(game.alliance)).toInt() And &HFF&
        game.power = ini.value("powerLevel", json_variant_new(game.power)).toInt()
        ini.endGroup
    End If

    loadMapChallengeSettings = True
End Function

Private Sub resolveAIForPlayer(ByVal player As Long, ByVal aiValue As String)
    Dim filename As String, ai As Long
    If aiValue = "null" Then
        Exit Sub
    End If

    ' strip given path down to filename (WzPathInfo::fromPlatformIndependentPath(...).fileName())
    filename = Mid$(aiValue, InStrRev(aiValue, "/") + 1)

    ' look up AI value in vector of known skirmish AIs
    For ai = 0 To aidata_size - 1
        If filename = aidata(ai).js Or aiValue = aidata(ai).js Then
            NetPlay_players_ai(player) = ai
            Exit Sub
        End If
    Next

    ' did not find from known skirmish AIs, assume custom AI
    NetPlay_players_ai(player) = AI_CUSTOM
End Sub

Private Sub loadMapPlayerSettings(ByVal ini As WzConfig)
    Dim i As Long, val_ As String, j As Long, f_id As Long, value As String, havePosition As Long, old As Long, pos As Long
    Dim position As Long

    For i = 0 To MAX_PLAYERS - 1
        ini.beginGroup "player_" & CStr(i)
        If ini.contains("team") Then
            NetPlay_players_team(i) = ini.value("team").toInt()
        ElseIf challengeActive Then ' team is a required key for challenges
            NetPlay_players_ai(i) = AI_CLOSED
        End If

        ' Load pre-configured AIs
        If ini.contains("ai") Then
            val_ = ini.value("ai").toWzString()
            resolveAIForPlayer i, val_
        End If

        ' Try finding a name field, if not found use AI names for AI players if in SP skirmish
        If ini.contains("name") Then
            setPlayerName i, ini.value("name").toWzString()
        ElseIf Not NetPlay_bComms And i <> selectedPlayer Then
            setPlayerName i, getAIName(i)
        End If

        NetPlay_players_position(i) = MAX_PLAYERS  ' Invalid value, fix later.
        If ini.contains("position") Then
            position = ini.value("position").toInt()
            If position < 0 Then position = 0
            If position > MAX_PLAYERS Then position = MAX_PLAYERS
            NetPlay_players_position(i) = position
        End If
        If ini.contains("difficulty") Then
            ' If difficulty is set, but we have no AI, use default
            If Not ini.contains("ai") Then
                NetPlay_players_ai(i) = 0
            End If

            value = ini.value("difficulty", json_variant_new("Medium")).toWzString()
            For j = 0 To 4  ' difficultyList.size()
                If LCase$(difficultyList_at(j)) = LCase$(value) Then   ' strcasecmp
                    NetPlay_players_difficulty(i) = difficultyValue_at(j)
                End If
            Next
        End If
        If ini.contains("faction") Then
            value = ini.value("faction", json_variant_new("Normal")).toWzString()
            For f_id = 0 To NUM_FACTIONS - 1
                If FACTION_name(getFactionByID(f_id)) = value Then
                    NetPlay_players_faction(i) = f_id
                End If
            Next
        End If
        If ini.contains("spectator") And Not challengeActive Then
            NetPlay_players_isSpectator(i) = ini.value("spectator", json_variant_new(False)).toBool()
            If NetPlay_players_isSpectator(i) Then
                If Not (Not ini.contains("ai")) Then ASSERT_FAILURE "!ini.contains(""ai"")", "src/multiint.cpp:5053", "loadMapPlayerSettings", "Player %d is set to be a spectator, but also has the ""ai"" key specified - this is unsupported", i
                NetPlay_players_ai(i) = AI_OPEN
            End If
        End If

        If ini.contains("closed") Then
            If ini.value("closed").toBool() Then
                NetPlay_players_ai(i) = AI_CLOSED
            End If
        End If
        ini.endGroup
    Next

    ' Fix duplicate or unset player positions.
    havePosition = 0                ' PlayerMask
    For i = 0 To MAX_PLAYERS - 1
        If NetPlay_players_position(i) < MAX_PLAYERS Then
            old = havePosition
            havePosition = havePosition Or BIT_(NetPlay_players_position(i))
            If havePosition = old Then
                ASSERT_FAILURE "false", "src/multiint.cpp:5078", "loadMapPlayerSettings", "Duplicate position %d at index: %d", NetPlay_players_position(i), i
                NetPlay_players_position(i) = MAX_PLAYERS
            End If
        End If
    Next
    pos = 0
    For i = 0 To MAX_PLAYERS - 1
        If NetPlay_players_position(i) >= MAX_PLAYERS Then
            Do While (havePosition And BIT_(pos)) <> 0
                pos = pos + 1
            Loop
            NetPlay_players_position(i) = pos
            pos = pos + 1
        End If
    Next
End Sub

Private Function playersPerTeam() As Long
    Dim numTeams As Long
    For numTeams = game.maxPlayers - 1 To 2 Step -1
        If game.maxPlayers Mod numTeams = 0 Then
            playersPerTeam = numTeams
            Exit Function
        End If
    Next
    playersPerTeam = 1
End Function

Private Sub swapPlayerColours(ByVal player1 As Long, ByVal player2 As Long)
    Dim player1Colour As Long
    If Not (ULT(player1, MAX_PLAYERS) And ULT(player2, MAX_PLAYERS)) Then ASSERT_FAILURE "player1 < MAX_PLAYERS && player2 < MAX_PLAYERS", "src/multiint.cpp:5112", "swapPlayerColours", "player1 (%u) & player2 (%u) must be < MAX_PLAYERS", player1, player2: Exit Sub
    player1Colour = getPlayerColour(player1)
    setPlayerColour player1, getPlayerColour(player2)
    setPlayerColour player2, player1Colour
End Sub

' Resets all player difficulties, positions, teams and colors etc.
' (static in C++; WzMultiplayerOptionsTitleUI.cls's start calls it)
Public Sub resetPlayerConfiguration(Optional ByVal bShouldResetLocal As Boolean = False)
    Dim selectedPlayerPosition As Long, playerIndex As Long, t As Long
    If bShouldResetLocal Then selectedPlayerPosition = 0 Else selectedPlayerPosition = NetPlay_players_position(selectedPlayer)
    For playerIndex = 0 To MAX_PLAYERS_IN_GUI - 1
        setPlayerColour playerIndex, playerIndex
        swapPlayerColours playerIndex, rand() Mod (playerIndex + 1)
        NetPlay_players_position(playerIndex) = playerIndex

        If Not bShouldResetLocal And playerIndex = selectedPlayer Then
            GoTo continue_
        End If

        NetPlay_players_team(playerIndex) = playerIndex \ playersPerTeam()

        If NetPlay_bComms Then
            NetPlay_players_difficulty(playerIndex) = AIDifficulty_DISABLED
            NetPlay_players_ai(playerIndex) = AI_OPEN
            clearPlayerName playerIndex
            If playerIndex = PLAYER_FEATURE Then
                NetPlay_players_ai(playerIndex) = AI_CLOSED
            End If
        Else
            NetPlay_players_difficulty(playerIndex) = AIDifficulty_DEFAULT
            NetPlay_players_ai(playerIndex) = 0

            ' ensure all players have a name in One Player Skirmish games
            setPlayerName playerIndex, getAIName(playerIndex)
        End If
continue_:
    Next

    For playerIndex = MAX_PLAYERS_IN_GUI To MAX_CONNECTED_PLAYERS - 1
        clearPlayerName playerIndex
    Next

    If Not bShouldResetLocal And selectedPlayerPosition < game.maxPlayers And selectedPlayer <> selectedPlayerPosition Then
        ' std::swap(NetPlay.players[selectedPlayer].position, NetPlay.players[selectedPlayerPosition].position)
        t = NetPlay_players_position(selectedPlayer)
        NetPlay_players_position(selectedPlayer) = NetPlay_players_position(selectedPlayerPosition)
        NetPlay_players_position(selectedPlayerPosition) = t
    End If

    setPlayerName selectedPlayer, sPlayer

    If selectedPlayer < MAX_PLAYERS Then
        For playerIndex = 0 To MAX_PLAYERS - 1
            If getPlayerColour(playerIndex) = war_getMPcolour() Then
                swapPlayerColours selectedPlayer, playerIndex
                Exit For
            End If
        Next
    End If
End Sub

' Loads challenge and player configurations from level/autohost/test .json-files.
' (static in C++; WzMultiplayerOptionsTitleUI.cls's start calls it)
Public Function loadMapChallengeAndPlayerSettings(Optional ByVal forceLoadPlayers As Boolean = False) As Boolean
    Dim aFileName As String, psLevel As LEVEL_DATASET, h As Sha256, ininame As String, warnIfMissing As Boolean
    Dim bIsOnline As Boolean, ini As WzConfig, retVal As Boolean, n As Long

    h = game.hash
    Set psLevel = levFindDataSet_hash(game.map, h, True)

    If psLevel Is Nothing Then ASSERT_FAILURE "psLevel", "src/multiint.cpp:5189", "loadMapChallengeAndPlayerSettings", "No level found for %s", game.map: loadMapChallengeAndPlayerSettings = False: Exit Function
    If Not (psLevel.game_ >= 0 And psLevel.game_ < LEVEL_MAXFILES) Then ASSERT_FAILURE "psLevel->game >= 0 && psLevel->game < LEVEL_MAXFILES", "src/multiint.cpp:5190", "loadMapChallengeAndPlayerSettings", "Invalid psLevel->game: %i - may be a corrupt level load (%s; hash: %s)", psLevel.game_, game.map, Sha256_toString(game.hash): loadMapChallengeAndPlayerSettings = False: Exit Function
    aFileName = Left$(psLevel.apDataFiles(psLevel.game_), 255)    ' char aFileName[256]
    n = Len(aFileName)
    If n < 4 Then n = 4
    aFileName = Left$(aFileName, n - 4)
    aFileName = Left$(aFileName & ".json", 255)

    If challengeActive Then ininame = sRequestResult Else ininame = aFileName
    warnIfMissing = False
    If getHostLaunch() = HostLaunch_Skirmish Then
        ininame = "tests/" & wz_skirmish_test()
        warnIfMissing = True
    End If
    If getHostLaunch() = HostLaunch_Autohost Then
        ininame = "autohost/" & wz_skirmish_test()
        warnIfMissing = True
    End If

    bIsOnline = NetPlay_bComms And NetPlay_isHost
    If PHYSFS_exists(ininame) = 0 Then
        If warnIfMissing Then
            debug_ 5213, LOG_ERROR, "loadMapChallengeAndPlayerSettings", "Missing specified file: %s", ininame
        End If

        ' Just reset the players if config is not found and host is not started yet
        If Not bIsOnline Then
            resetPlayerConfiguration
        End If

        loadMapChallengeAndPlayerSettings = Not warnIfMissing ' not treated as fatal failure unless warnIfMissing
        Exit Function
    End If
    Set ini = WzConfig_new(ininame, WzConfig_ReadOnly)

    retVal = loadMapChallengeSettings(ini)

    ' Do not load player settings if we are already hosting an online match
    If Not bIsOnline Or forceLoadPlayers Then
        loadMapPlayerSettings ini
    End If

    loadMapChallengeAndPlayerSettings = retVal
End Function

Public Sub displayLobbyDisabledNotification()
    Dim notification As WZ_Notification, infoLink As String
    If Not hasNotificationsWithTag(LOBBY_DISABLED_TAG) Then
        notification.duration = 0
        notification.contentTitle = "Multiplayer Lobby Support Unavailable"

        notification.contentText = "Your client cannot connect to the mutiplayer lobby."
        notification.contentText = notification.contentText & vbLf & vbLf
        notification.contentText = notification.contentText & "Please click the button below for more information on how to fix it."

        infoLink = NET_getLobbyDisabledInfoLinkURL()
        notification.action_title = "More Information"  ' WZ_Notification_Action(_("More Information"), [infoLink](const WZ_Notification&) {...})
        Set notification.action_onAction = WzFunc_new(FN_lobbyDisabled_moreInformation, infoLink)
        notification.tag = LOBBY_DISABLED_TAG

        addNotification notification, 0                 ' WZ_Notification_Trigger::Immediate()
    End If
End Sub

' Process click events on the multiplayer/skirmish options screen
' 'id' is id of the button that was pressed
Public Sub startMultiplayerGame()
    Dim i As Long
    If Not NetPlay_isHost Then ASSERT_FAILURE "false", "src/multiint.cpp:5544", "startMultiplayerGame", "Host only routine detected for client!": Exit Sub

    wz_command_interface_output "WZEVENT: startMultiplayerGame" & vbLf
    debug_ 5547, LOG_INFO, "startMultiplayerGame", "startMultiplayerGame"

    cancelOrDismissNotificationsWithTag LOBBY_DISABLED_TAG
    cancelOrDismissNotificationIfTagHasPrefix SLOTTYPE_TAG_PREFIX  ' (the lambda: tag.rfind(SLOTTYPE_TAG_PREFIX, 0) == 0)

    decideWRF                                       ' set up swrf & game.map
    bMultiPlayer = True
    bMultiMessages = True
    NETsetPlayerConnectionStatus CONNECTIONSTATUS_NORMAL, NET_ALL_PLAYERS  ' reset disconnect conditions
    initLoadingScreen True
    If NetPlay_isHost Then
        ' This sets the limits to whatever the defaults are for the limiter screen
        ' If host sets limits, then we do not need to do the following routine.
        If Not bLimiterLoaded Then
            debug_ 5565, LOG_NET, "startMultiplayerGame", "limiter was NOT activated, setting defaults"

            If Not resLoad("wrf/limiter_data.wrf", 503) Then
                debug_ 5569, LOG_INFO, "startMultiplayerGame", "Unable to load limiter_data."
            End If
        Else
            debug_ 5574, LOG_NET, "startMultiplayerGame", "limiter was activated"
        End If

        resetDataHash    ' need to reset it, since host's data has changed.
        createLimitSet
        debug_ 5579, LOG_NET, "startMultiplayerGame", "sending our options to all clients"
        sendOptions
        NEThaltJoining True                         ' stop new players entering.
        ingame.has_TimeEveryoneIsInGame = False
        ingame.has_endTime = False
        ingame.isAllPlayersDataOK = False
        For i = 0 To MAX_CONNECTED_PLAYERS - 1      ' memset(&ingame.DataIntegrity, 0x0, sizeof(ingame.DataIntegrity)): clear all player's array
            ingame.DataIntegrity(i) = False
        Next
        SendFireUp                                  ' bcast a fireup message
    End If

    debug_ 5589, LOG_NET, "startMultiplayerGame", "title mode STARTGAME is set--Starting Game!"
    changeTitleMode STARTGAME

    If NetPlay_isHost Then
        sendRoomSystemMessage "Host is Starting Game"
    End If
End Sub

' (static in C++; WzMultiplayerOptionsTitleUI.cls calls it)
Public Sub printHostHelpMessagesToConsole()
    Dim buf As String
    If challengeActive Then
        buf = "Hit the ready box to begin your challenge!"
        displayRoomNotifyMessage buf
    ElseIf Not NetPlay_isHost Then
        buf = "Press the start hosting button to begin hosting a game."
        displayRoomNotifyMessage buf
    End If
    If NetPlay_bComms Then
        If NetPlay_isPortMappingEnabled Then
            If Not ipv4MappingRequest_exists() Then
                Exit Sub
            End If
            Select Case PortMappingManager_get_status_ipv4MappingRequest()
            ' rely on callbacks / output in NETaddRedirects() for everything except pending
            Case PortMappingDiscoveryStatus_IN_PROGRESS
                buf = "Port mapping creation is in progress..."
                displayRoomNotifyMessage buf
            Case Else
            End Select
        Else
            buf = va("Port mapping disabled by user. Autoconfig of port %d will not happen.", NETgetGameserverPort())
            displayRoomNotifyMessage buf
        End If
    End If
End Sub

Public Sub printBlindModeHelpMessagesToConsole()
    Dim blindLobbyMessage As String, blindLobbyCodenameMessage As String
    If game.blindMode = BLIND_MODE_NONE Then
        Exit Sub
    End If

    If game.blindMode >= BLIND_MODE_BLIND_GAME Then
        blindLobbyMessage = "BLIND GAME NOTICE:"
        blindLobbyMessage = blindLobbyMessage & vbLf & "- "
        blindLobbyMessage = blindLobbyMessage & "Player names will not be visible to other players (until the game is over)"
    Else
        blindLobbyMessage = "BLIND LOBBY NOTICE:"
        blindLobbyMessage = blindLobbyMessage & vbLf & "- "
        blindLobbyMessage = blindLobbyMessage & "Player names will not be visible to other players (until the game has started)"
    End If
    If NetPlay_hostPlayer < MAX_PLAYERS Then
        blindLobbyMessage = blindLobbyMessage & vbLf & "- "
        blindLobbyMessage = blindLobbyMessage & "IMPORTANT: The host is a player, and can see and configure other players (!)"
    Else
        blindLobbyMessage = blindLobbyMessage & vbLf & "- "
        blindLobbyMessage = blindLobbyMessage & "The host is a spectator, and can see player identities and configure players / teams"
    End If
    displayRoomNotifyMessage blindLobbyMessage
    If selectedPlayer < MAX_PLAYERS Then
        blindLobbyCodenameMessage = "- "
        blindLobbyCodenameMessage = blindLobbyCodenameMessage & va("You have been assigned the codename: ""%s""", getPlayerGenericName(selectedPlayer))
        displayRoomNotifyMessage blindLobbyCodenameMessage
    End If
End Sub

Private Function multiplayHasOpenPlayerSlot() As Boolean
    Dim j As Long
    For j = 0 To MAX_PLAYERS - 1
        If NetPlay_players_ai(j) = AI_OPEN And Not NetPlay_players_allocated(j) And NetPlay_players_position(j) < game.maxPlayers Then
            multiplayHasOpenPlayerSlot = True
            Exit Function
        End If
    Next
    multiplayHasOpenPlayerSlot = False
End Function

Public Function multiplayPlayersCanCheckReady() As Boolean
    If isBlindSimpleLobby(game.blindMode) Then
        ' if in a blind simple lobby, players can always check ready
        multiplayPlayersCanCheckReady = True
        Exit Function
    End If

    If locked.readybeforefull Then
        If multiplayHasOpenPlayerSlot() Then
            multiplayPlayersCanCheckReady = False
            Exit Function
        End If
    End If

    multiplayPlayersCanCheckReady = (allPlayersOnSameTeam(-1) = -1)
End Function

Public Function multiplayPlayersShouldCheckReady() As Boolean
    If Not multiplayPlayersCanCheckReady() Then
        multiplayPlayersShouldCheckReady = False
        Exit Function
    End If
    multiplayPlayersShouldCheckReady = isBlindSimpleLobby(game.blindMode) Or Not multiplayHasOpenPlayerSlot()
End Function

Public Sub sendRoomSystemMessage(ByVal text As String)
    Dim message As NetworkTextMessage
    If Not NetPlay_isHost Then ASSERT_FAILURE "false", "src/multiint.cpp:7435", "sendRoomSystemMessage", "Host only routine detected for client!": Exit Sub
    message = NetworkTextMessage_new(SYSTEM_MESSAGE, text)
    displayRoomSystemMessage text
    NetworkTextMessage_enqueue message, NETbroadcastQueue()
End Sub

' WzMultiplayerOptionsTitleUI::WzMultiplayerOptionsTitleUI(std::shared_ptr<WzTitleUI> parent),
' std::make_shared<WzMultiplayerOptionsTitleUI>(parent)
Public Function WzMultiplayerOptionsTitleUI_new(ByVal parent As WzTitleUI) As WzTitleUI
    Dim ui As WzMultiplayerOptionsTitleUI
    Set ui = New WzMultiplayerOptionsTitleUI
    ui.init parent
    Set WzMultiplayerOptionsTitleUI_new = ui
End Function

Public Function formatGameName(ByVal name_ As String) As String
    Dim withoutTechlevel As String
    withoutTechlevel = mapNameWithoutTechlevel(name_)
    formatGameName = withoutTechlevel & " (T" & CStr(game.techLevel) & " " & CStr(game.maxPlayers) & "P)"
End Function

' multiint.cpp: void to_json(nlohmann::json& j, const JoinConnectionDescription::JoinConnectionType& v)
Public Function JoinConnectionType_to_json(ByVal v As JoinConnectionType) As nlohmann_json
    Set JoinConnectionType_to_json = json_value(JoinConnectionDescription_connectiontype_to_string(v))
End Function

' multiint.cpp: void to_json(nlohmann::json& j, const JoinConnectionDescription& v)
Public Function JoinConnectionDescription_to_json(v As JoinConnectionDescription) As nlohmann_json
    Dim j As nlohmann_json
    Set j = json_object()
    j.set_item "h", json_value(v.host)
    j.set_item "p", json_value_unsigned(v.port)
    j.set_item "t", JoinConnectionType_to_json(v.type_)
    Set JoinConnectionDescription_to_json = j
End Function

' multiint.cpp: void from_json(const nlohmann::json& j, JoinConnectionDescription::JoinConnectionType& v)
Public Function JoinConnectionType_from_json(ByVal j As nlohmann_json) As JoinConnectionType
    Dim str_ As String, result As Long
    str_ = j.get_string()
    If JoinConnectionDescription_connectiontype_from_string(str_, result) Then
        JoinConnectionType_from_json = result
    Else
        Err.Raise vbObjectError + 302, "nlohmann::json", "[json.exception.type_error.302] JoinConnectionType value is unknown: """ & str_ & """"
    End If
End Function

' multiint.cpp: void from_json(const nlohmann::json& j, JoinConnectionDescription& v)
Public Sub JoinConnectionDescription_from_json(ByVal j As nlohmann_json, v As JoinConnectionDescription)
    v.host = j.at("h").get_string()
    v.port = j.at("p").get_int()                ' get<int32_t>()
    v.type_ = JoinConnectionType_from_json(j.at("t"))
End Sub

' multiint.cpp: void from_json(const nlohmann::json& j, KickRedirectInfo& v)
Public Sub KickRedirectInfo_from_json(ByVal j As nlohmann_json, v As KickRedirectInfo)
    Dim c As nlohmann_json, i As Long
    ' v.connList = j.at("c").get<std::vector<JoinConnectionDescription>>()
    Set c = j.at("c")
    If Not c.is_array() Then json_throw_not_array c
    v.connListCount = c.size()
    If v.connListCount > 0 Then ReDim v.connList(0 To v.connListCount - 1) Else Erase v.connList
    For i = 0 To v.connListCount - 1
        JoinConnectionDescription_from_json c.at_idx(i), v.connList(i)
    Next
    v.has_gamePassword = False
    v.gamePassword = ""
    If j.contains("p") Then
        v.gamePassword = j.at("p").get_string()
        v.has_gamePassword = True
    End If
    v.asSpectator = False
    If j.contains("s") Then
        v.asSpectator = j.at("s").get_uint() <> 0
    End If
End Sub

' optional<KickRedirectInfo> parseKickRedirectInfo(const std::string& redirectString, const std::string& currHostAddress)
' (has_value: the result; the value in out)
Public Function parseKickRedirectInfo(ByVal redirectString As String, ByVal currHostAddress As String, out As KickRedirectInfo) As Boolean
    Dim redirectInfo As KickRedirectInfo, obj As nlohmann_json, i As Long, n As Long
    ' Parse the redirect string (should be a valid json object containing an array of connection descriptions and other connection options)
    On Error GoTo invalid
    Set obj = json_parse(redirectString)
    KickRedirectInfo_from_json obj, redirectInfo
    On Error GoTo 0

    ' Process & verify the connection descriptions
    For i = 0 To redirectInfo.connListCount - 1
        If Len(redirectInfo.connList(i).host) = 0 Or redirectInfo.connList(i).host = "=" Then
            ' replace with the same host's address
            redirectInfo.connList(i).host = currHostAddress
        End If
    Next
    ' erase(remove_if(..., desc.host != currHostAddress || desc.port <= 1024)): the others keep their order
    n = 0
    For i = 0 To redirectInfo.connListCount - 1
        If Not (redirectInfo.connList(i).host <> currHostAddress Or Not ULT(1024, redirectInfo.connList(i).port)) Then
            redirectInfo.connList(n) = redirectInfo.connList(i)
            n = n + 1
        End If
    Next
    redirectInfo.connListCount = n
    If n > 0 Then ReDim Preserve redirectInfo.connList(0 To n - 1) Else Erase redirectInfo.connList

    out = redirectInfo
    parseKickRedirectInfo = True
    Exit Function
invalid:
    debug_ 5895, LOG_ERROR, "parseKickRedirectInfo", "Invalid redirect string - could not be processed: %s", redirectString
    parseKickRedirectInfo = False
End Function
