VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WZGameReplayOptionsHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/multiint.cpp, multiplay.h: class WZGameReplayOptionsHandler, the
' game's ReplayOptionsHandler (WzNet's interface): what a replay saves of
' the game's options and how it restores them. The JSON converters are
' multiint_c's and multistat_c's. EmbeddedMapData is its fields; the
' translated texts (_()) are the English ones.
Option Explicit

Implements ReplayOptionsHandler

Private Const MAX_REPLAY_EMBEDDED_MAP_SIZE_LIMIT As Double = 150# * 1024#  ' PHYSFS_sint64
Private Const SIZE_MAX_D As Double = 18446744073709551615#                  ' std::numeric_limits<size_t>::max()
Private Const MAX_CONNECTED_PLAYERS As Long = MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS

Private Function ReplayOptionsHandler_saveOptions(ByVal object_ As nlohmann_json) As Boolean
    Dim j As nlohmann_json, players As nlohmann_json, multiStatsJson As nlohmann_json, i As Long

    ' random seed
    object_.set_item "randSeed", json_value_unsigned(gameRand_GetSeed())

    ' Save versionString
    object_.set_item "versionString", json_value(version_getVersionString())

    ' Save DataHash here, so we can provide a warning on load (after all data / the game is loaded) in the future
    object_.set_item "dataHash", DataHashToJSON()

    ' Save `game`
    MULTIPLAYERGAME_to_json j, game
    object_.set_item "game", j

    ' Save `ingame` (? all of it? some of it?)
    MULTIPLAYERINGAME_to_json j, ingame
    object_.set_item "ingame", j

    ' Save `selectedPlayer` (even though we don't actually use it on restore - for now)
    object_.set_item "selectedPlayer", json_value_unsigned(selectedPlayer)

    ' Save campaignName
    object_.set_item "campaignName", json_value(getCampaignName())

    ' Save tweakOptions
    object_.set_item "tweakOptions", getCamTweakOptions()

    ' ? Do not need to save alliances (?)

    ' Save `NetPlay.players` -- possibly recreate NetPlay.playerReferences on load?
    Set players = json_array()
    For i = 0 To NetPlay_players_size() - 1
        PLAYER_to_json j, i
        players.push_back j
    Next
    object_.set_item "netplay.players", players

    ' Save `NetPlay.hostPlayer`
    object_.set_item "netplay.hostPlayer", json_value_unsigned(NetPlay_hostPlayer)

    ' Save `NetPlay.bComms`
    object_.set_item "netplay.bComms", json_value(NetPlay_bComms)

'   // Save `NetPlay.isHost` (but don't load it)
'   object["netplay.isHost"] = NetPlay.isHost;

    ' multistats
    Set multiStatsJson = json_array()
    If Not saveMultiStatsToJSON(multiStatsJson) Then
        debug_ 7810, LOG_ERROR, "saveOptions", "Failed to save multistats info to JSON"
        ReplayOptionsHandler_saveOptions = False
        Exit Function
    End If
    object_.set_item "multistats", multiStatsJson

    ReplayOptionsHandler_saveOptions = True
End Function

Private Function ReplayOptionsHandler_optionsUpdatePlayerInfo(ByVal object_ As nlohmann_json) As Boolean
    Dim netPlayers As nlohmann_json, i As Long, p As PLAYER_, blank As PLAYER_, j As nlohmann_json, multistats As nlohmann_json

    ' update player names
    Set netPlayers = object_.at("netplay.players")
    If Not netPlayers.is_array() Then
        debug_ 7824, LOG_ERROR, "optionsUpdatePlayerInfo", "Invalid netplay.players (not an array)"
        ReplayOptionsHandler_optionsUpdatePlayerInfo = False
        Exit Function
    End If
    If netPlayers.size() > NetPlay_players_size() Then
        debug_ 7829, LOG_ERROR, "optionsUpdatePlayerInfo", "Unsupported netplay.players size (%zu)", netPlayers.size()
        ReplayOptionsHandler_optionsUpdatePlayerInfo = False
        Exit Function
    End If
    For i = 0 To netPlayers.size() - 1
        p = blank                               ' get<PLAYER>(): a value-initialised PLAYER, filled
        PLAYER_from_json_struct netPlayers.at_idx(i), p
        p.name_ = Left$(NetPlay_players_name(i), 63)   ' sstrcpy(p.name, NetPlay.players[i].name): char[StringSize]
        PLAYER_to_json_struct j, p
        netPlayers.set_item_idx i, j
    Next

    ' update identities stored in multistats
    If game.blindMode <> BLIND_MODE_NONE Then
        Set multistats = object_.at("multistats")
        updateMultiStatsIdentitiesInJSON multistats, True
    End If

    ReplayOptionsHandler_optionsUpdatePlayerInfo = True
End Function

Private Function ReplayOptionsHandler_maximumEmbeddedMapBufferSize() As Double
    ReplayOptionsHandler_maximumEmbeddedMapBufferSize = MAX_REPLAY_EMBEDDED_MAP_SIZE_LIMIT
End Function

' EmbeddedMapData& embeddedMapData: dataVersion, mapBinaryData(0 .. mapBinaryDataSize - 1)
Private Function ReplayOptionsHandler_saveMap(dataVersion As Long, mapBinaryData() As Byte, mapBinaryDataSize As Long) As Boolean
    Dim mapData As LEVEL_DATASET, mapFileHandle As Long, filesize As Double, readBytes As Long

    If Not (mapBinaryDataSize = 0) Then ASSERT_FAILURE "embeddedMapData.mapBinaryData.empty()", "src/multiint.cpp:7858", "saveMap", "Non empty output buffer?": ReplayOptionsHandler_saveMap = False: Exit Function

    ' Provide "version" for this data, just in case it needs to change in the future
    dataVersion = 1

    If game.isMapMod Then
        ' do not store map-mods in the replay as they are (a) deprecated and (b) may be huge
        ReplayOptionsHandler_saveMap = True
        Exit Function
    End If

    Set mapData = levFindDataSet_hash(game.map, game.hash, True)
    If Not (Not mapData Is Nothing) Then ASSERT_FAILURE "mapData != nullptr", "src/multiint.cpp:7870", "saveMap", "Unable to find map??": ReplayOptionsHandler_saveMap = False: Exit Function

    If StrPtr(mapData.realFileName) = 0 Then
        ' built-in maps don't have realFileNames for the map archives, nor do we want to save them
        ReplayOptionsHandler_saveMap = True
        Exit Function
    End If

    mapFileHandle = PHYSFS_openRead(mapData.realFileName)
    If mapFileHandle = 0 Then
        debug_ 7881, LOG_INFO, "saveMap", "Unable to open for reading: %s; error: %s", mapData.realFileName, WZ_PHYSFS_getLastError()
        ReplayOptionsHandler_saveMap = False
        Exit Function
    End If
    ' (VB) gsl::finally(PHYSFS_close): the handle is closed before each return below
    filesize = PHYSFS_fileLength(mapFileHandle)
    If Not (filesize > 0) Then
        ASSERT_FAILURE "filesize > 0", "src/multiint.cpp:7888", "saveMap", "Invalid filesize"
        PHYSFS_close mapFileHandle
        ReplayOptionsHandler_saveMap = False
        Exit Function
    End If
    If filesize > MAX_REPLAY_EMBEDDED_MAP_SIZE_LIMIT Then
        ' currently, an expected failure
        PHYSFS_close mapFileHandle
        ReplayOptionsHandler_saveMap = True
        Exit Function
    End If
    ' resize the mapBinaryData to be able to contain the map data
    ReDim mapBinaryData(0 To CLng(filesize) - 1)
    mapBinaryDataSize = CLng(filesize)
    readBytes = PHYSFS_readBytes(mapFileHandle, mapBinaryData, 0, CLng(filesize))
    If Not (readBytes = filesize) Then
        ASSERT_FAILURE "readBytes == filesize", "src/multiint.cpp:7897", "saveMap", "Read failed"
        PHYSFS_close mapFileHandle
        ReplayOptionsHandler_saveMap = False
        Exit Function
    End If
    PHYSFS_close mapFileHandle

    ReplayOptionsHandler_saveMap = True
End Function

Private Function ReplayOptionsHandler_desiredBufferSize() As Double
    Dim currentGameMode As Long
    currentGameMode = ActivityManager_getCurrentGameMode()
    Select Case currentGameMode
    Case ActivitySink_GameMode_MENUS
        ' should not happen
    Case ActivitySink_GameMode_CAMPAIGN
        ' replays not currently supported
    Case ActivitySink_GameMode_CHALLENGE, ActivitySink_GameMode_SKIRMISH
        ReplayOptionsHandler_desiredBufferSize = 0 ' use default
        Exit Function
    Case ActivitySink_GameMode_MULTIPLAYER
        ' big games need a big buffer
        ReplayOptionsHandler_desiredBufferSize = SIZE_MAX_D
        Exit Function
    Case Else
        debug_ 7921, LOG_INFO, "desiredBufferSize", "Unhandled case: %u", currentGameMode
    End Select
    ReplayOptionsHandler_desiredBufferSize = 0
End Function

' EmbeddedMapData&& embeddedMapData: dataVersion, mapBinaryData(0 .. mapBinaryDataSize - 1)
Private Function ReplayOptionsHandler_restoreOptions(ByVal object_ As nlohmann_json, ByVal dataVersion As Long, mapBinaryData() As Byte, ByVal mapBinaryDataSize As Long, ByVal replayNetcodeMajor As Long, ByVal replayNetcodeMinor As Long) As Boolean
    Dim rand_seed As Long, replayVersionStr As String, expectedVersionStr As String, nonMatchingVersionString As Boolean
    Dim nonMatchingNetcodeVersion As Boolean, mismatchVersionDescription As String, dialogKind As Long
    Dim restoredGame As MULTIPLAYERGAME, restoredIngame As MULTIPLAYERINGAME, restoredCampaignName As String
    Dim tweakOptionsJSON As nlohmann_json, tweakOptionsJSONMap As nlohmann_json, k As Long
    Dim netPlayers As nlohmann_json, i As Long, mapData As LEVEL_DATASET, h As Long, filename As String, fileHash As Sha256
    Dim replaySpectatorIndex As Long

    ' random seed
    rand_seed = object_.at("randSeed").get_uint()
    gameSRand rand_seed

    ' compare version_string and throw a pop-up warning if not equal
    replayVersionStr = object_.at("versionString").get_string()
    expectedVersionStr = version_getVersionString()
    nonMatchingVersionString = replayVersionStr <> expectedVersionStr
    nonMatchingNetcodeVersion = Not NETisCorrectVersion(replayNetcodeMajor, replayNetcodeMinor)
    If nonMatchingVersionString Or nonMatchingNetcodeVersion Then
        If nonMatchingVersionString Then
            debug_ 7941, LOG_INFO, "restoreOptions", "Version string mismatch: (replay: %s) - (current: %s)", replayVersionStr, expectedVersionStr
        End If
        mismatchVersionDescription = "The version of Warzone 2100 used to save this replay file does not match the currently-running version."
        mismatchVersionDescription = mismatchVersionDescription & vbLf & vbLf
        mismatchVersionDescription = mismatchVersionDescription & va("Replay File Saved With: ""%s""", replayVersionStr)
        mismatchVersionDescription = mismatchVersionDescription & vbLf
        If nonMatchingVersionString Then
            mismatchVersionDescription = mismatchVersionDescription & va("Current Warzone 2100 Version: ""%s""", expectedVersionStr)
        Else
            ' only a netcode mismatch - weird, but display the netcode version info
            mismatchVersionDescription = mismatchVersionDescription & va("Replay File NetcodeVer: (0x%x, 0x%x)", replayNetcodeMajor, replayNetcodeMinor)
            mismatchVersionDescription = mismatchVersionDescription & vbLf
            mismatchVersionDescription = mismatchVersionDescription & va("Current Warzone 2100 NetcodeVer: (0x%x, 0x%x)", NETGetMajorVersion(), NETGetMinorVersion())
        End If
        mismatchVersionDescription = mismatchVersionDescription & vbLf & vbLf
        mismatchVersionDescription = mismatchVersionDescription & "Replays should usually be played back with the same version used to save the replay."
        mismatchVersionDescription = mismatchVersionDescription & vbLf
        mismatchVersionDescription = mismatchVersionDescription & "The replay may not playback successfully, or there may be differences in the simulation."
        If nonMatchingVersionString Then dialogKind = Dialog_Warning Else dialogKind = Dialog_Error
        wzDisplayDialog dialogKind, "Replay Version Mismatch", mismatchVersionDescription
    End If

    ' restore `game`
    MULTIPLAYERGAME_from_json object_.at("game"), restoredGame
    game = restoredGame

    ' restore `ingame`
    MULTIPLAYERINGAME_from_json object_.at("ingame"), restoredIngame
    ingame = restoredIngame

    ' restore `campaignName`
    If object_.contains("campaignName") Then
        restoredCampaignName = object_.at("campaignName").get_string()
        If Len(restoredCampaignName) > 0 Then
            setCampaignName restoredCampaignName
        Else
            clearCampaignName
        End If
    Else
        clearCampaignName
    End If

    ' restore `tweakOptions`
    If object_.contains("tweakOptions") Then
        Set tweakOptionsJSON = object_.at("tweakOptions").clone()
        ' (VB) std::unordered_map<std::string, nlohmann::json> tweakOptionsJSONMap: a json object of the items' keys and values
        Set tweakOptionsJSONMap = json_object()
        If tweakOptionsJSON.is_object() Then
            For k = 0 To tweakOptionsJSON.size() - 1
                tweakOptionsJSONMap.set_item tweakOptionsJSON.key_at(k), tweakOptionsJSON.value_at(k)
            Next
        ElseIf tweakOptionsJSON.is_array() Then
            For k = 0 To tweakOptionsJSON.size() - 1
                tweakOptionsJSONMap.set_item CStr(k), tweakOptionsJSON.at_idx(k)   ' an array's items: the indexes as keys
            Next
        ElseIf Not tweakOptionsJSON.is_null() Then
            tweakOptionsJSONMap.set_item "", tweakOptionsJSON                      ' a primitive's one item: an empty key
        End If
        setCamTweakOptions tweakOptionsJSONMap
    Else
        clearCamTweakOptions
    End If

    ' ? Do not need to restore alliances (?)

    ' restore NetPlay.players
    Set netPlayers = object_.at("netplay.players").clone()
    If Not netPlayers.is_array() Then
        debug_ 8011, LOG_ERROR, "restoreOptions", "Invalid netplay.players (not an array)"
        ReplayOptionsHandler_restoreOptions = False
        Exit Function
    End If
    If netPlayers.size() > NetPlay_players_size() Then
        debug_ 8016, LOG_ERROR, "restoreOptions", "Unsupported netplay.players size (%zu)", netPlayers.size()
        ReplayOptionsHandler_restoreOptions = False
        Exit Function
    End If
    For i = 0 To netPlayers.size() - 1
        PLAYER_from_json netPlayers.at_idx(i), i
    Next

    ' restore NetPlay.hostPlayer
    NetPlay_hostPlayer = object_.at("netplay.hostPlayer").get_uint()

    ' restore `NetPlay.bComms` (?)
    NetPlay_bComms = object_.at("netplay.bComms").get_bool()

    ' restore multistats
    If Not loadMultiStatsFromJSON(object_.at("multistats")) Then
        debug_ 8033, LOG_ERROR, "restoreOptions", "Failed to load multistats info from JSON"
        ReplayOptionsHandler_restoreOptions = False
        Exit Function
    End If

    ' Now verify the info - especially that we can load the level!

    ' clear out the old level list.
    levShutDown
    levInitialise
    rebuildSearchPath mod_multiplay, True       ' MUST rebuild search path for the new maps we just got!
    pal_Init                                    'Update palettes.
    If Not buildMapList() Then
        debug_ 8046, LOG_ERROR, "restoreOptions", "Failed to build map list"
        ReplayOptionsHandler_restoreOptions = False
        Exit Function
    End If

    If dataVersion = 1 Then
        If mapBinaryDataSize <> 0 Then
            If mapBinaryDataSize <= MAX_REPLAY_EMBEDDED_MAP_SIZE_LIMIT Then
                setSpecialInMemoryMap mapBinaryData, mapBinaryDataSize
            Else
                debug_ 8060, LOG_ERROR, "restoreOptions", "Embedded map data size (%zu) exceeds maximum supported size", mapBinaryDataSize
            End If
        End If
    End If

    Set mapData = levFindDataSet_hash(game.map, game.hash, True)
    ' See if we have the map or not
    If mapData Is Nothing Then
        ' Failed to load map
        debug_ 8070, LOG_POPUP, "restoreOptions", "Missing map used for replay: ""%s"" (hash: %s)", game.map, Sha256_toString(game.hash)
        ReplayOptionsHandler_restoreOptions = False
        Exit Function
    End If
    ' Must restore `useTerrainOverrides` (this matters for re-loading the map!) - see loadMapPreview() in multiint.cpp
    builtInMap = (StrPtr(mapData.realFileName) = 0)
    useTerrainOverrides = False
    If builtInMap Then useTerrainOverrides = shouldLoadTerrainTypeOverrides(mapData.pName)

    For h = 0 To game.modHashCount - 1
        ' TODO: Actually check the loaded mods??

        filename = Left$("mods/downloads/" & Sha256_toString(game.modHashes(h)), 255)  ' char[256]

        If PHYSFS_exists(filename) = 0 Then
            debug_ 8086, LOG_ERROR, "restoreOptions", "Downloaded mod does not exist: %s", filename
            ReplayOptionsHandler_restoreOptions = False
            Exit Function
        Else
            findHashOfFile filename, fileHash
            If Not Sha256_equals(fileHash, game.modHashes(h)) Then
                debug_ 8091, LOG_ERROR, "restoreOptions", "Downloaded mod is incomplete or corrupt: %s", filename
                ReplayOptionsHandler_restoreOptions = False
                Exit Function
            Else
                ' Have the file!
                debug_ 8097, LOG_INFO, "restoreOptions", "Downloaded mod found: %s", filename
            End If
        End If
    Next

    If Not mapData Is Nothing Then
        If CheckForMod(mapData.realFileName) Then
            game.isMapMod = True
        End If
    End If
    game.isRandom = False
    If Not mapData Is Nothing Then game.isRandom = CheckForRandom(mapData.realFileName, mapData.apDataFiles(0))

    ' Set various other initialization things (see NET_FIREUP)
    ingame.has_TimeEveryoneIsInGame = False     ' reset time
    ingame.has_endTime = False
    resetDataHash
    decideWRF

    ' set selectedPlayer to a *new* spectator slot
    If NetPlay_players_size() <> MAX_CONNECTED_PLAYERS Then
        ' NetPlay.players.size() should always start at MAX_CONNECTED_PLAYERS before this is called
        debug_ 8117, LOG_ERROR, "restoreOptions", "Unexpected NetPlay.players.size(): %zu", NetPlay_players_size()
        ReplayOptionsHandler_restoreOptions = False
        Exit Function
    End If
    NetPlay_players_push_back
    replaySpectatorIndex = NetPlay_players_size() - 1
    NET_InitPlayer replaySpectatorIndex, False  ' re-init everything
    NetPlay_players_allocated(replaySpectatorIndex) = True
    NetPlay_players_name(replaySpectatorIndex) = "Replay Viewer"  'setPlayerName() asserts cause this index is now equal to MAX_CONNECTED_PLAYERS
    NetPlay_players_isSpectator(replaySpectatorIndex) = True

    selectedPlayer = replaySpectatorIndex
    realSelectedPlayer = selectedPlayer

    ' Need to initialize net/game queues here before we try to add messages to them
    For i = 0 To MAX_GAMEQUEUE_SLOTS - 1
        NETinitQueue NETgameQueue(i)
        NETsetNoSendOverNetwork NETgameQueue(i)
    Next

    ' simulate "NET_PLAYERRESPONDING" sent by all clients
    ' ensure that all humans players are set to "joined"
    For i = 0 To MAX_CONNECTED_PLAYERS - 1
        If isHumanPlayer(i) Then
            If ingame.JoiningInProgress(i) Then
                ingame.JoiningInProgress(i) = False
            End If
        End If
    Next
    NetPlay_isHostAlive = True                  ' lie and say the host is alive for multiplayer games

    ReplayOptionsHandler_restoreOptions = True
End Function
