Attribute VB_Name = "game_main_c"
' src/game.cpp (its lines 1-5242; game_c has the start of a game and loadGame): the main save file.
' The SAVE_GAME versions (SAVE_GAME_V7 ... SAVE_GAME_V38 in one Type, SAVE_GAME, as the C++'s
' typedef SAVE_GAME_V38 SAVE_GAME; a field keeps the version that added it in its comment) with
' their JSON serialisers and deserialisers and the binary deserialisers of the (de)serialisation API
' (a version 39 .gam without its gam.json: deserializeSaveGameData), saveGame, writeGameFile,
' writeMainFile, writeMapFile, writeGameInfo (game.cpp:6678), saveJSONToFile, loadMainFile,
' loadMainFileFinal, sanityUpdate, savegameWithoutExtension, gameScreenSizeDidChange and
' gameDisplayScaleFactorDidChange. game_c's gameLoadV reads its saveGameData through
' deserializeSaveGameData_json / deserializeSaveGameData.
' Not ported: the binary serialisers (serializeSaveGameV7Data ... serializeSaveGameData, the
' serialize*(PHYSFS_file *) helpers): 4.7.0 writes the .gam as gam.json only and never calls them
' (game.cpp's pragma: "Ignore unused functions for now (until full cleanup of old binary savegame
' .gam writing)"). A char[N] field is a String (sstrcpy: N - 1 characters at most); uint8_t and
' uint16_t fields are Longs. The object files (droid.json, struct.json, ...) are game_objects_c's.
Option Explicit

' gamedef.h: the version numbers used in save games (the ones this module tests)
Private Const VERSION_34 As Long = 34   ' saves AI names for multiplayer (.gam file)
Private Const VERSION_39 As Long = 39   ' lots of changes, breaking everything

Private Const CURRENT_VERSION_NUM As Long = VERSION_39

Private Const currentGameVersion As Long = 410  ' static unsigned int currentGameVersion = 410;

Private Const SAVEKEY_ONMISSION As Long = &H100

' general save definitions
Private Const MAX_LEVEL_SIZE As Long = 20

Private Const PATH_MAX As Long = 260            ' MAX_PATH on Windows

Private Const StringSize As Long = 64           ' netplay.h: size of strings used.
Private Const modlist_string_size As Long = 255 ' netplay.h: For a concatenated list of mods

' general save definitions
Public Type SAVE_POWER
    currentPower As Long                ' uint32_t
    extractedPower As Long              ' uint32_t: used for hacks
End Type

'* struct used to store the data for retreating.
Public Type RUN_DATA
    sPos As Vector2i                    ' = Vector2i(0, 0): position to where units should flee to.
    forceLevel As Long                  ' uint8_t = 0: number of units below which others might flee.
    healthLevel As Long                 ' uint8_t = 0: health percentage value below which it might flee. This value is used for groups only.
    leadership As Long                  ' uint8_t = 0: basic value that will be used on calculations of the flee probability.
End Type

' a save game's PLAYER (what serializePlayer / deserializePlayer keep of NetPlay.players[i])
Public Type SAVE_GAME_PLAYER
    position As Long                    ' uint32_t
    name_ As String                     ' char name[StringSize]
    ai As Long                          ' int8_t
    difficulty As Long                  ' AIDifficulty
    allocated As Boolean
    colour As Long                      ' uint32_t
    team As Long                        ' uint32_t
End Type

' a save game's NETPLAY (serializeNetPlay / deserializeNetPlay)
Public Type SAVE_GAME_NETPLAY
    players(0 To MAX_PLAYERS - 1) As SAVE_GAME_PLAYER
    bComms As Boolean
    playercount As Long                 ' uint32_t
    hostPlayer As Long                  ' uint32_t
    isHost As Boolean
End Type

' SAVE_GAME (SAVE_GAME_V38): its versions' fields in their order
Public Type SAVE_GAME
    ' SAVE_GAME_V7
    gameTime As Long                    ' uint32_t
    GameType As Long                    ' uint32_t: Type of game , one of the GTYPE_... enums.
    ScrollMinX As Long                  ' int32_t: Scroll Limits
    ScrollMinY As Long
    ScrollMaxX As Long                  ' uint32_t
    ScrollMaxY As Long
    levelName As String                 ' char levelName[MAX_LEVEL_SIZE]: name of the level to load up when mid game
    ' SAVE_GAME_V10
    power(0 To MAX_PLAYERS - 1) As SAVE_POWER
    ' SAVE_GAME_V11
    currentPlayerPos As iView
    ' SAVE_GAME_V12
    missionTime As Long                 ' uint32_t
    saveKey As Long                     ' uint32_t
    ' SAVE_GAME_V14
    missionOffTime As Long              ' int32_t
    missionETA As Long                  ' int32_t
    missionHomeLZ_X As Long             ' uint16_t
    missionHomeLZ_Y As Long             ' uint16_t
    missionPlayerX As Long              ' int32_t
    missionPlayerY As Long              ' int32_t
    iTranspEntryTileX(0 To MAX_PLAYERS - 1) As Long  ' uint16_t
    iTranspEntryTileY(0 To MAX_PLAYERS - 1) As Long
    iTranspExitTileX(0 To MAX_PLAYERS - 1) As Long
    iTranspExitTileY(0 To MAX_PLAYERS - 1) As Long
    aDefaultSensor(0 To MAX_PLAYERS - 1) As Long     ' uint32_t
    aDefaultECM(0 To MAX_PLAYERS - 1) As Long
    aDefaultRepair(0 To MAX_PLAYERS - 1) As Long
    ' SAVE_GAME_V15 (uint8_t aDroidExperience[MAX_PLAYERS][MAX_RECYCLED_DROIDS]: read into the
    ' experience queue, not kept)
    offWorldKeepLists As Long           ' int32_t: was BOOL (which was a int)
    RubbleTile As Long                  ' uint32_t
    WaterTile As Long                   ' uint32_t
    fogColour As Long                   ' uint32_t
    fogState As Long                    ' uint32_t
    ' SAVE_GAME_V16
    sLandingZone(0 To MAX_NOGO_AREAS - 1) As LANDING_ZONE
    ' SAVE_GAME_V17
    objId As Long                       ' uint32_t
    ' SAVE_GAME_V18
    buildDate As String                 ' char buildDate[MAX_STR_LENGTH]
    oldestVersion As Long               ' uint32_t
    validityKey As Long                 ' uint32_t
    ' SAVE_GAME_V19
    alliances(0 To MAX_PLAYERS - 1, 0 To MAX_PLAYERS - 1) As Long    ' uint8_t
    playerColour(0 To MAX_PLAYERS - 1) As Long                      ' uint8_t
    radarZoom As Long                   ' uint8_t
    ' SAVE_GAME_V20
    bDroidsToSafetyFlag As Long         ' uint8_t
    asVTOLReturnPos(0 To MAX_PLAYERS - 1) As Vector2i
    ' SAVE_GAME_V22
    asRunData(0 To MAX_PLAYERS - 1) As RUN_DATA
    ' SAVE_GAME_V24
    reinforceTime As Long               ' uint32_t
    bPlayCountDown As Long              ' uint8_t
    bPlayerHasWon As Long               ' uint8_t
    bPlayerHasLost As Long              ' uint8_t
    dummy3 As Long                      ' uint8_t
    ' SAVE_GAME_V27 (uint16_t awDroidExperience[MAX_PLAYERS][MAX_RECYCLED_DROIDS]: read into the
    ' experience queue, not kept)
    ' SAVE_GAME_V29
    missionScrollMinX As Long           ' uint16_t
    missionScrollMinY As Long
    missionScrollMaxX As Long
    missionScrollMaxY As Long
    ' SAVE_GAME_V30
    scrGameLevel As Long                ' int32_t
    bExtraVictoryFlag As Long           ' uint8_t
    bExtraFailFlag As Long              ' uint8_t
    bTrackTransporter As Long           ' uint8_t
    ' SAVE_GAME_V31: extra code for the patch - saves out whether cheated with the mission timer
    missionCheatTime As Long            ' int32_t
    ' SAVE_GAME_V33: alexl. skirmish saves
    sGame As MULTIPLAYERGAME
    sNetPlay As SAVE_GAME_NETPLAY
    savePlayer As Long                  ' uint32_t
    sPName As String                    ' char sPName[32]
    multiPlayer As Long                 ' int32_t: was BOOL (int) ** see warning about conversion
    sPlayerIndex(0 To MAX_PLAYERS - 1) As Long       ' uint32_t
    ' SAVE_GAME_V34: Now holds AI names for multiplayer
    sPlayerName(0 To MAX_PLAYERS - 1) As String      ' char sPlayerName[MAX_PLAYERS][StringSize]
    ' SAVE_GAME_V38: Store loaded mods in savegame
    modList As String                   ' char modList[modlist_string_size]
End Type

' (VB) sstrcpy into a char[size]: at most size - 1 characters
Private Function sstrcpyN(ByVal src As String, ByVal size As Long) As String
    sstrcpyN = Left$(src, size - 1)
End Function

' (VB) get<int8_t>() of a JSON number: its low byte, signed
Private Function json_int8(ByVal j As nlohmann_json) As Long
    Dim v As Long
    v = j.get_int() And &HFF&
    If v > 127 Then v = v - 256
    json_int8 = v
End Function

' (VB) physfs_ext.h's PHYSFS_readUBE8, PHYSFS_readSBE8 and PHYSFS_readUBE16 (WzFramework has the 32-bit ones):
' false when the bytes are not there
Private Function readUBE8(ByVal fileHandle As Long, v As Long) As Boolean
    Dim b(0 To 0) As Byte
    If WZ_PHYSFS_readBytes(fileHandle, b, 0, 1) <> 1 Then Exit Function
    v = b(0)
    readUBE8 = True
End Function

Private Function readSBE8(ByVal fileHandle As Long, v As Long) As Boolean
    Dim b(0 To 0) As Byte
    If WZ_PHYSFS_readBytes(fileHandle, b, 0, 1) <> 1 Then Exit Function
    v = b(0)
    If v > 127 Then v = v - 256
    readSBE8 = True
End Function

Private Function readUBE16(ByVal fileHandle As Long, v As Long) As Boolean
    Dim b(0 To 1) As Byte
    If WZ_PHYSFS_readBytes(fileHandle, b, 0, 2) <> 2 Then Exit Function
    v = b(0) * 256& + b(1)
    readUBE16 = True
End Function

Private Function readUBE32(ByVal fileHandle As Long, v As Long) As Boolean
    readUBE32 = PHYSFS_readUBE32(fileHandle, v) <> 0
End Function

Private Function readSBE32(ByVal fileHandle As Long, v As Long) As Boolean
    readSBE32 = PHYSFS_readSBE32(fileHandle, v) <> 0
End Function

' (VB) WZ_PHYSFS_readBytes(fileHandle, char[n], n) == n: the text up to its terminating zero in s
Private Function readChars(ByVal fileHandle As Long, ByVal n As Long, s As String) As Boolean
    Dim b() As Byte, i As Long
    ReDim b(0 To n - 1)
    If WZ_PHYSFS_readBytes(fileHandle, b, 0, n) <> n Then Exit Function
    s = ""
    For i = 0 To n - 1
        If b(i) = 0 Then Exit For
        s = s & Chr$(b(i))
    Next
    readChars = True
End Function

Public Function saveJSONToFile(ByVal obj As nlohmann_json, ByVal pFileName As String) As Boolean
    Dim jsonString As String, bytes() As Byte, n As Long
    On Error GoTo catch_
    jsonString = obj.dump(4)
    On Error GoTo 0
    debug_ 141, LOG_SAVE, "saveJSONToFile", "%s %s", "Saving", pFileName
    n = utf8_encode(jsonString, bytes)
    saveJSONToFile = saveFile(pFileName, bytes, n)
    Exit Function
catch_:
    ASSERT_FAILURE "false", "src/game.cpp:138", "saveJSONToFile", "Failed to save JSON to %s with error: %s", pFileName, Err.Description
    saveJSONToFile = False
End Function

Public Sub gameScreenSizeDidChange(ByVal oldWidth As Long, ByVal oldHeight As Long, ByVal newWidth As Long, ByVal newHeight As Long)
    If GetGameMode() = GS_NORMAL And Not gamePaused() Then ' if in match / game and not paused (i.e. no in-game menus open, etc)
        ' update mouse trap status if window properties changed
        If shouldTrapCursor() Then
            wzGrabMouse
        Else
            wzReleaseMouse
        End If
    End If
    intScreenSizeDidChange oldWidth, oldHeight, newWidth, newHeight
    loadSaveScreenSizeDidChange oldWidth, oldHeight, newWidth, newHeight
    challengesScreenSizeDidChange oldWidth, oldHeight, newWidth, newHeight
    multiMenuScreenSizeDidChange oldWidth, oldHeight, newWidth, newHeight
    display3dScreenSizeDidChange oldWidth, oldHeight, newWidth, newHeight
    consoleScreenDidChangeSize oldWidth, oldHeight, newWidth, newHeight
    frontendScreenSizeDidChange oldWidth, oldHeight, newWidth, newHeight
    widgOverlaysScreenSizeDidChange oldWidth, oldHeight, newWidth, newHeight ' must be last!
End Sub

Public Sub gameDisplayScaleFactorDidChange(ByVal newDisplayScaleFactor As Single)
    Dim horizGameToRendererScaleFactor As Long, vertGameToRendererScaleFactor As Long
    ' The text subsystem requires the game -> renderer scale factor, which potentially differs from
    ' the display scale factor.
    wzGetGameToRendererScaleFactorInt horizGameToRendererScaleFactor, vertGameToRendererScaleFactor
    iV_TextUpdateScaleFactor horizGameToRendererScaleFactor, vertGameToRendererScaleFactor
End Sub

'--------------------------------------------------------------------------
' The SAVE_GAME versions' JSON serialisers and deserialisers, and their binary deserialisers

Private Sub serializeSavePowerData_json(ByVal o As nlohmann_json, serializePower As SAVE_POWER)
    o.set_item "currentPower", json_value_unsigned(serializePower.currentPower)
    o.set_item "extractedPower", json_value_unsigned(serializePower.extractedPower)
End Sub

Private Sub deserializeSavePowerData_json(ByVal o As nlohmann_json, serializePower As SAVE_POWER)
    serializePower.currentPower = o.at("currentPower").get_uint()
    serializePower.extractedPower = o.at("extractedPower").get_uint()
End Sub

Private Function deserializeSavePowerData(ByVal fileHandle As Long, serializePower As SAVE_POWER) As Boolean
    If Not readUBE32(fileHandle, serializePower.currentPower) Then Exit Function
    deserializeSavePowerData = readUBE32(fileHandle, serializePower.extractedPower)
End Function

Private Sub serializeVector3i_json(ByVal o As nlohmann_json, serializeVector As Vector3i)
    o.set_item "x", json_value(serializeVector.x)
    o.set_item "y", json_value(serializeVector.y)
    o.set_item "z", json_value(serializeVector.z)
End Sub

Private Sub deserializeVector3i_json(ByVal o As nlohmann_json, serializeVector As Vector3i)
    serializeVector.x = o.at("x").get_int()
    serializeVector.y = o.at("y").get_int()
    serializeVector.z = o.at("z").get_int()
End Sub

Private Function deserializeVector3i(ByVal fileHandle As Long, serializeVector As Vector3i) As Boolean
    Dim x As Long, y As Long, z As Long

    If Not readSBE32(fileHandle, x) Then Exit Function
    If Not readSBE32(fileHandle, y) Then Exit Function
    If Not readSBE32(fileHandle, z) Then Exit Function

    serializeVector.x = x
    serializeVector.y = y
    serializeVector.z = z

    deserializeVector3i = True
End Function

Private Sub serializeVector2i_json(ByVal o As nlohmann_json, serializeVector As Vector2i)
    o.set_item "x", json_value(serializeVector.x)
    o.set_item "y", json_value(serializeVector.y)
End Sub

Private Sub deserializeVector2i_json(ByVal o As nlohmann_json, serializeVector As Vector2i)
    serializeVector.x = o.at("x").get_int()
    serializeVector.y = o.at("y").get_int()
End Sub

Private Function deserializeVector2i(ByVal fileHandle As Long, serializeVector As Vector2i) As Boolean
    Dim x As Long, y As Long

    If Not readSBE32(fileHandle, x) Then Exit Function
    If Not readSBE32(fileHandle, y) Then Exit Function

    serializeVector.x = x
    serializeVector.y = y

    deserializeVector2i = True
End Function

Private Sub serializeiViewData_json(ByVal o As nlohmann_json, serializeView As iView)
    Dim viewP As nlohmann_json, viewR As nlohmann_json
    Set viewP = json_object()
    serializeVector3i_json viewP, serializeView.p
    o.set_item "viewDataP", viewP
    Set viewR = json_object()
    serializeVector3i_json viewR, serializeView.r
    o.set_item "viewDataR", viewR
End Sub

Private Sub deserializeiViewData_json(ByVal o As nlohmann_json, serializeView As iView)
    deserializeVector3i_json o.at("viewDataP"), serializeView.p
    deserializeVector3i_json o.at("viewDataR"), serializeView.r
End Sub

Private Function deserializeiViewData(ByVal fileHandle As Long, serializeView As iView) As Boolean
    If Not deserializeVector3i(fileHandle, serializeView.p) Then Exit Function
    deserializeiViewData = deserializeVector3i(fileHandle, serializeView.r)
End Function

Private Sub serializeRunData_json(ByVal o As nlohmann_json, serializeRun As RUN_DATA)
    serializeVector2i_json o, serializeRun.sPos
    o.set_item "forceLevel", json_value_unsigned(serializeRun.forceLevel)
    o.set_item "healthLevel", json_value_unsigned(serializeRun.healthLevel)
    o.set_item "leadership", json_value_unsigned(serializeRun.leadership)
End Sub

Private Sub deserializeRunData_json(ByVal o As nlohmann_json, serializeRun As RUN_DATA)
    deserializeVector2i_json o, serializeRun.sPos
    serializeRun.forceLevel = o.at("forceLevel").get_uint() And &HFF&
    serializeRun.healthLevel = o.at("healthLevel").get_uint() And &HFF&
    serializeRun.leadership = o.at("leadership").get_uint() And &HFF&
End Sub

Private Function deserializeRunData(ByVal fileHandle As Long, serializeRun As RUN_DATA) As Boolean
    If Not deserializeVector2i(fileHandle, serializeRun.sPos) Then Exit Function
    If Not readUBE8(fileHandle, serializeRun.forceLevel) Then Exit Function
    If Not readUBE8(fileHandle, serializeRun.healthLevel) Then Exit Function
    deserializeRunData = readUBE8(fileHandle, serializeRun.leadership)
End Function

Private Sub serializeLandingZoneData_json(ByVal o As nlohmann_json, serializeLandZone As LANDING_ZONE)
    o.set_item "x1", json_value_unsigned(serializeLandZone.x1)
    o.set_item "y1", json_value_unsigned(serializeLandZone.y1)
    o.set_item "x2", json_value_unsigned(serializeLandZone.x2)
    o.set_item "y2", json_value_unsigned(serializeLandZone.y2)
End Sub

Private Sub deserializeLandingZoneData_json(ByVal o As nlohmann_json, serializeLandZone As LANDING_ZONE)
    serializeLandZone.x1 = o.at("x1").get_uint() And &HFF&
    serializeLandZone.y1 = o.at("y1").get_uint() And &HFF&
    serializeLandZone.x2 = o.at("x2").get_uint() And &HFF&
    serializeLandZone.y2 = o.at("y2").get_uint() And &HFF&
End Sub

Private Function deserializeLandingZoneData(ByVal fileHandle As Long, serializeLandZone As LANDING_ZONE) As Boolean
    Dim v As Long
    If Not readUBE8(fileHandle, v) Then Exit Function
    serializeLandZone.x1 = v
    If Not readUBE8(fileHandle, v) Then Exit Function
    serializeLandZone.y1 = v
    If Not readUBE8(fileHandle, v) Then Exit Function
    serializeLandZone.x2 = v
    If Not readUBE8(fileHandle, v) Then Exit Function
    serializeLandZone.y2 = v
    deserializeLandingZoneData = True
End Function

Private Sub serializeMultiplayerGame_json(ByVal o As nlohmann_json, serializeMulti As MULTIPLAYERGAME)
    o.set_item "multiType", json_value_unsigned(serializeMulti.typ)
    o.set_item "multiMapName", json_value(serializeMulti.map)
    o.set_item "multiMaxPlayers", json_value_unsigned(serializeMulti.maxPlayers)
    o.set_item "multiGameName", json_value(serializeMulti.name_)
    o.set_item "multiPower", json_value_unsigned(serializeMulti.power)
    o.set_item "multiBase", json_value_unsigned(serializeMulti.base)
    o.set_item "multiAlliance", json_value_unsigned(serializeMulti.alliance)
    o.set_item "multiHashBytes", json_value(32&) ' serializeMulti->hash.Bytes
    o.set_item "multiHash", json_value(Sha256_toString(serializeMulti.hash))
    ' skip more dummy

End Sub

Private Sub deserializeMultiplayerGame_json(ByVal o As nlohmann_json, serializeMulti As MULTIPLAYERGAME)
    Dim sha256_ As Sha256             ' Sha256 sha256;
    serializeMulti.typ = o.at("multiType").get_uint() And &HFF&
    serializeMulti.map = sstrcpyN(o.at("multiMapName").get_string(), 128)
    serializeMulti.maxPlayers = o.at("multiMaxPlayers").get_uint() And &HFF&
    serializeMulti.name_ = sstrcpyN(o.at("multiGameName").get_string(), 128)
    serializeMulti.power = o.at("multiPower").get_uint()
    serializeMulti.base = o.at("multiBase").get_uint() And &HFF&
    serializeMulti.alliance = o.at("multiAlliance").get_uint() And &HFF&
    Sha256_fromString sha256_, o.at("multiHash").get_string()
    serializeMulti.hash = sha256_
End Sub

Private Function deserializeMultiplayerGame(ByVal fileHandle As Long, serializeMulti As MULTIPLAYERGAME) As Boolean
    Dim boolFog As Long, dummy8 As Long, dummy16 As Long, dummy8c As String, hashSize As Long, b() As Byte, k As Long

    Sha256_setZero serializeMulti.hash

    If Not readUBE8(fileHandle, serializeMulti.typ) Then Exit Function
    If Not readChars(fileHandle, 128, serializeMulti.map) Then Exit Function
    If Not readChars(fileHandle, 8, dummy8c) Then Exit Function
    If Not readUBE8(fileHandle, serializeMulti.maxPlayers) Then Exit Function
    If Not readChars(fileHandle, 128, serializeMulti.name_) Then Exit Function
    If Not readSBE32(fileHandle, boolFog) Then Exit Function
    If Not readUBE32(fileHandle, serializeMulti.power) Then Exit Function
    If Not readUBE8(fileHandle, serializeMulti.base) Then Exit Function
    If Not readUBE8(fileHandle, serializeMulti.alliance) Then Exit Function
    If Not readUBE8(fileHandle, hashSize) Then Exit Function
    If hashSize = 32 Then               ' serializeMulti->hash.Bytes
        ReDim b(0 To 31)
        If WZ_PHYSFS_readBytes(fileHandle, b, 0, 32) = 0 Then Exit Function
        For k = 0 To 31
            serializeMulti.hash.bytes(k) = b(k)
        Next
    End If
    If Not readUBE16(fileHandle, dummy16) Then Exit Function    ' dummy, was bytesPerSec
    If Not readUBE8(fileHandle, dummy8) Then Exit Function      ' dummy, was packetsPerSec
    If Not readUBE8(fileHandle, dummy8) Then Exit Function      ' reused for challenge, was encryptKey
    challengeActive = dummy8 <> 0       ' hack

    For k = 0 To MAX_PLAYERS - 1
        ' dummy, was `skDiff` for each player
        If Not readUBE8(fileHandle, dummy8) Then Exit Function
    Next

    deserializeMultiplayerGame = True
End Function

Private Sub serializePlayer_json(ByVal o As nlohmann_json, serializePlayer As SAVE_GAME_PLAYER, ByVal player As Long)
    o.set_item "position", json_value_unsigned(serializePlayer.position)
    o.set_item "name", json_value(serializePlayer.name_)
    o.set_item "aiName", json_value(getAIName(player))
    o.set_item "difficulty", json_value(serializePlayer.difficulty)   ' static_cast<int8_t>
    o.set_item "allocated", json_value_unsigned(IIf(serializePlayer.allocated, 1&, 0&))
    o.set_item "colour", json_value_unsigned(serializePlayer.colour)
    o.set_item "team", json_value_unsigned(serializePlayer.team)
End Sub

Private Sub deserializePlayer_json(ByVal o As nlohmann_json, serializePlayer As SAVE_GAME_PLAYER, ByVal player As Long)
    Dim aiName As String
    aiName = "THEREISNOAI"
    If Not (o.is_object()) Then ASSERT_FAILURE "o.is_object()", "src/game.cpp:616", "deserializePlayer_json", "unexpected type, wanted object"
    serializePlayer.name_ = sstrcpyN(o.at("name").get_string(), StringSize)
    aiName = sstrcpyN(o.at("aiName").get_string(), MAX_LEN_AI_NAME)
    serializePlayer.difficulty = json_int8(o.at("difficulty"))
    serializePlayer.allocated = (o.at("allocated").get_uint() And &HFF&) <> 0
    If player < game.maxPlayers Then
        serializePlayer.ai = matchAIbyName(aiName)
        If Not (serializePlayer.ai <> AI_NOT_FOUND) Then ASSERT_FAILURE "serializePlayer->ai != -99", "src/game.cpp:624", "deserializePlayer_json", "AI ""%s"" not found -- script loading will fail (player %d / %d)", _
                aiName, player, game.maxPlayers
    End If
    serializePlayer.position = o.at("position").get_uint()
    serializePlayer.colour = o.at("colour").get_uint()
    serializePlayer.team = o.at("team").get_uint()
End Sub

Private Function deserializePlayer(ByVal fileHandle As Long, serializePlayer As SAVE_GAME_PLAYER, ByVal player As Long) As Boolean
    Dim aiName As String, position As Long, colour As Long, team As Long, retval As Boolean, allocated As Long

    aiName = "THEREISNOAI"
    retval = False
    retval = readUBE32(fileHandle, position)
    If retval Then retval = readChars(fileHandle, StringSize, serializePlayer.name_)
    If retval Then retval = readChars(fileHandle, MAX_LEN_AI_NAME, aiName)
    If retval Then retval = readSBE8(fileHandle, serializePlayer.difficulty)
    If retval Then retval = readUBE8(fileHandle, allocated)
    If retval Then retval = readUBE32(fileHandle, colour)
    If retval Then retval = readUBE32(fileHandle, team)

    serializePlayer.allocated = allocated <> 0
    If player < game.maxPlayers Then
        serializePlayer.ai = matchAIbyName(aiName)
        If Not (serializePlayer.ai <> AI_NOT_FOUND) Then ASSERT_FAILURE "serializePlayer->ai != -99", "src/game.cpp:650", "deserializePlayer", "AI ""%s"" not found -- script loading will fail (player %d / %d)", _
               aiName, player, game.maxPlayers
    End If
    serializePlayer.position = position
    serializePlayer.colour = colour
    serializePlayer.team = team
    deserializePlayer = retval
End Function

Private Sub serializeNetPlay_json(ByVal o As nlohmann_json, serializeNetPlay As SAVE_GAME_NETPLAY)
    Dim arr As nlohmann_json, tmp As nlohmann_json, i As Long
    Set arr = json_array()
    For i = 0 To MAX_PLAYERS - 1
        Set tmp = json_object()
        serializePlayer_json tmp, serializeNetPlay.players(i), i
        arr.push_back tmp
    Next
    o.set_item "netbComms", json_value(serializeNetPlay.bComms)
    o.set_item "netPlayerCount", json_value_unsigned(serializeNetPlay.playercount)
    o.set_item "netHostPlayer", json_value_unsigned(serializeNetPlay.hostPlayer)
    o.set_item "netSelectedPlayer", json_value_unsigned(selectedPlayer)
    o.set_item "netScavengers", json_value_unsigned(game.scavengers)
    ' skip dummy
    o.set_item "netPlayers", arr
End Sub

Private Sub deserializeNetPlay_json(ByVal o As nlohmann_json, serializeNetPlay As SAVE_GAME_NETPLAY)
    Dim players As nlohmann_json, i As Long
    Set players = o.at("netPlayers")
    If Not (players.is_array()) Then ASSERT_FAILURE "players.is_array()", "src/game.cpp:697", "deserializeNetPlay_json", "unexpected type, wanted array": Exit Sub
    For i = 0 To MAX_PLAYERS - 1
        deserializePlayer_json players.at_idx(i), serializeNetPlay.players(i), i
    Next
    serializeNetPlay.isHost = True ' only host can load
    serializeNetPlay.playercount = o.at("netPlayerCount").get_uint()
    serializeNetPlay.bComms = o.at("netbComms").get_bool()
    selectedPlayer = o.at("netSelectedPlayer").get_uint()
    game.scavengers = o.at("netScavengers").get_uint() And &HFF&
End Sub

Private Function deserializeNetPlay(ByVal fileHandle As Long, serializeNetPlay As SAVE_GAME_NETPLAY) As Boolean
    Dim i As Long, retv As Boolean, dummy As Long, bComms As Long, scavs As Long, sel As Long

    For i = 0 To MAX_PLAYERS - 1
        If Not deserializePlayer(fileHandle, serializeNetPlay.players(i), i) Then Exit Function
    Next

    bComms = IIf(serializeNetPlay.bComms, 1&, 0&)
    scavs = game.scavengers

    serializeNetPlay.isHost = True  ' only host can load
    sel = selectedPlayer
    retv = readUBE32(fileHandle, bComms)
    If retv Then retv = readUBE32(fileHandle, serializeNetPlay.playercount)
    If retv Then retv = readUBE32(fileHandle, serializeNetPlay.hostPlayer)
    If retv Then
        retv = readUBE32(fileHandle, sel)
        selectedPlayer = sel
    End If
    If retv Then retv = readUBE32(fileHandle, scavs)
    If retv Then retv = readUBE32(fileHandle, dummy)
    If retv Then retv = readUBE32(fileHandle, dummy)
    serializeNetPlay.bComms = bComms <> 0
    game.scavengers = scavs And &HFF&
    deserializeNetPlay = retv
End Function

' SAVE_GAME_V7
Private Sub serializeSaveGameV7Data_json(ByVal o As nlohmann_json, serializeGame As SAVE_GAME)
    o.set_item "gameTime", json_value_unsigned(serializeGame.gameTime)
    o.set_item "GameType", json_value_unsigned(serializeGame.GameType)
    o.set_item "ScrollMinX", json_value(serializeGame.ScrollMinX)
    o.set_item "ScrollMinY", json_value(serializeGame.ScrollMinY)
    o.set_item "ScrollMaxX", json_value_unsigned(serializeGame.ScrollMaxX)
    o.set_item "ScrollMaxY", json_value_unsigned(serializeGame.ScrollMaxY)
    o.set_item "levelName", json_value(serializeGame.levelName)
End Sub

Private Sub deserializeSaveGameV7Data_json(ByVal o As nlohmann_json, serializeGame As SAVE_GAME)
    serializeGame.gameTime = o.at("gameTime").get_uint()
    serializeGame.GameType = o.at("GameType").get_uint()
    serializeGame.ScrollMinX = o.at("ScrollMinX").get_int()
    serializeGame.ScrollMinY = o.at("ScrollMinY").get_int()
    serializeGame.ScrollMaxX = o.at("ScrollMaxX").get_uint()
    serializeGame.ScrollMaxY = o.at("ScrollMaxY").get_uint()
    serializeGame.levelName = sstrcpyN(o.at("levelName").get_string(), MAX_LEVEL_SIZE)
End Sub

Private Function deserializeSaveGameV7Data(ByVal fileHandle As Long, serializeGame As SAVE_GAME) As Boolean
    If Not readUBE32(fileHandle, serializeGame.gameTime) Then Exit Function
    If Not readUBE32(fileHandle, serializeGame.GameType) Then Exit Function
    If Not readSBE32(fileHandle, serializeGame.ScrollMinX) Then Exit Function
    If Not readSBE32(fileHandle, serializeGame.ScrollMinY) Then Exit Function
    If Not readUBE32(fileHandle, serializeGame.ScrollMaxX) Then Exit Function
    If Not readUBE32(fileHandle, serializeGame.ScrollMaxY) Then Exit Function
    deserializeSaveGameV7Data = readChars(fileHandle, MAX_LEVEL_SIZE, serializeGame.levelName)
End Function

' SAVE_GAME_V10 : public SAVE_GAME_V7
Private Sub serializeSaveGameV10Data_json(ByVal o As nlohmann_json, serializeGame As SAVE_GAME)
    Dim arr As nlohmann_json, tmp As nlohmann_json, i As Long
    serializeSaveGameV7Data_json o, serializeGame
    Set arr = json_array()
    For i = 0 To MAX_PLAYERS - 1
        Set tmp = json_object()
        serializeSavePowerData_json tmp, serializeGame.power(i)
        arr.push_back tmp
    Next
    o.set_item "power", arr
End Sub

Private Sub deserializeSaveGameV10Data_json(ByVal o As nlohmann_json, serializeGame As SAVE_GAME)
    Dim power As nlohmann_json, i As Long
    deserializeSaveGameV7Data_json o, serializeGame
    Set power = o.at("power")
    If Not (power.is_array()) Then ASSERT_FAILURE "power.is_array()", "src/game.cpp:830", "deserializeSaveGameV10Data_json", "unexpected type"
    For i = 0 To MAX_PLAYERS - 1
        deserializeSavePowerData_json power.at_idx(i), serializeGame.power(i)
    Next
End Sub

Private Function deserializeSaveGameV10Data(ByVal fileHandle As Long, serializeGame As SAVE_GAME) As Boolean
    Dim i As Long

    If Not deserializeSaveGameV7Data(fileHandle, serializeGame) Then Exit Function

    For i = 0 To MAX_PLAYERS - 1
        If Not deserializeSavePowerData(fileHandle, serializeGame.power(i)) Then Exit Function
    Next

    deserializeSaveGameV10Data = True
End Function

' SAVE_GAME_V11 : public SAVE_GAME_V10
Private Sub serializeSaveGameV11Data_json(ByVal o As nlohmann_json, serializeGame As SAVE_GAME)
    serializeSaveGameV10Data_json o, serializeGame
    serializeiViewData_json o, serializeGame.currentPlayerPos

End Sub

Private Sub deserializeSaveGameV11Data_json(ByVal o As nlohmann_json, serializeGame As SAVE_GAME)
    deserializeSaveGameV10Data_json o, serializeGame
    deserializeiViewData_json o, serializeGame.currentPlayerPos
End Sub

Private Function deserializeSaveGameV11Data(ByVal fileHandle As Long, serializeGame As SAVE_GAME) As Boolean
    If Not deserializeSaveGameV10Data(fileHandle, serializeGame) Then Exit Function
    deserializeSaveGameV11Data = deserializeiViewData(fileHandle, serializeGame.currentPlayerPos)
End Function

' SAVE_GAME_V12 : public SAVE_GAME_V11
Private Sub serializeSaveGameV12Data_json(ByVal o As nlohmann_json, serializeGame As SAVE_GAME)
    serializeSaveGameV11Data_json o, serializeGame
    o.set_item "missionTime", json_value_unsigned(serializeGame.missionTime)
    o.set_item "saveKey", json_value_unsigned(serializeGame.saveKey)
End Sub

Private Sub deserializeSaveGameV12Data_json(ByVal o As nlohmann_json, serializeGame As SAVE_GAME)
    deserializeSaveGameV11Data_json o, serializeGame
    serializeGame.missionTime = o.at("missionTime").get_uint()
    serializeGame.saveKey = o.at("saveKey").get_uint()
End Sub

Private Function deserializeSaveGameV12Data(ByVal fileHandle As Long, serializeGame As SAVE_GAME) As Boolean
    If Not deserializeSaveGameV11Data(fileHandle, serializeGame) Then Exit Function
    If Not readUBE32(fileHandle, serializeGame.missionTime) Then Exit Function
    deserializeSaveGameV12Data = readUBE32(fileHandle, serializeGame.saveKey)
End Function

' SAVE_GAME_V14 : public SAVE_GAME_V12
Private Sub serializeSaveGameV14Data_json(ByVal o As nlohmann_json, serializeGame As SAVE_GAME)
    Dim arr As nlohmann_json, tmp As nlohmann_json, i As Long
    serializeSaveGameV12Data_json o, serializeGame
    o.set_item "missionOffTime", json_value(serializeGame.missionOffTime)
    o.set_item "missionETA", json_value(serializeGame.missionETA)
    o.set_item "missionHomeLZ_X", json_value_unsigned(serializeGame.missionHomeLZ_X)
    o.set_item "missionHomeLZ_Y", json_value_unsigned(serializeGame.missionHomeLZ_Y)
    o.set_item "missionPlayerX", json_value(serializeGame.missionPlayerX)
    o.set_item "missionPlayerY", json_value(serializeGame.missionPlayerY)
    Set arr = json_array()
    For i = 0 To MAX_PLAYERS - 1
        Set tmp = json_object()
        tmp.set_item "iTranspEntryTileX", json_value_unsigned(serializeGame.iTranspEntryTileX(i))
        tmp.set_item "iTranspEntryTileY", json_value_unsigned(serializeGame.iTranspEntryTileY(i))
        tmp.set_item "iTranspExitTileX", json_value_unsigned(serializeGame.iTranspExitTileX(i))
        tmp.set_item "iTranspExitTileY", json_value_unsigned(serializeGame.iTranspExitTileY(i))
        tmp.set_item "aDefaultSensor", json_value_unsigned(serializeGame.aDefaultSensor(i))
        tmp.set_item "aDefaultECM", json_value_unsigned(serializeGame.aDefaultECM(i))
        tmp.set_item "aDefaultRepair", json_value_unsigned(serializeGame.aDefaultRepair(i))
        arr.push_back tmp
    Next
    o.set_item "data", arr
End Sub

Private Sub deserializeSaveGameV14Data_json(ByVal o As nlohmann_json, serializeGame As SAVE_GAME)
    Dim arr As nlohmann_json, i As Long
    deserializeSaveGameV12Data_json o, serializeGame
    serializeGame.missionOffTime = o.at("missionOffTime").get_int()
    serializeGame.missionETA = o.at("missionETA").get_int()
    serializeGame.missionHomeLZ_X = o.at("missionHomeLZ_X").get_uint() And &HFFFF&
    serializeGame.missionHomeLZ_Y = o.at("missionHomeLZ_Y").get_uint() And &HFFFF&
    serializeGame.missionPlayerX = o.at("missionPlayerX").get_int()
    serializeGame.missionPlayerY = o.at("missionPlayerY").get_int()
    Set arr = o.at("data")
    If Not (arr.is_array()) Then ASSERT_FAILURE "arr.is_array()", "src/game.cpp:1039", "deserializeSaveGameV14Data_json", "unexpected type, wanted array": Exit Sub
    For i = 0 To MAX_PLAYERS - 1
        serializeGame.iTranspEntryTileX(i) = o.at("data").at_idx(i).at("iTranspEntryTileX").get_uint() And &HFFFF&
        serializeGame.iTranspEntryTileY(i) = o.at("data").at_idx(i).at("iTranspEntryTileY").get_uint() And &HFFFF&
        serializeGame.iTranspExitTileX(i) = o.at("data").at_idx(i).at("iTranspExitTileX").get_uint() And &HFFFF&
        serializeGame.iTranspExitTileY(i) = o.at("data").at_idx(i).at("iTranspExitTileY").get_uint() And &HFFFF&
        serializeGame.aDefaultSensor(i) = o.at("data").at_idx(i).at("aDefaultSensor").get_uint()
        serializeGame.aDefaultECM(i) = o.at("data").at_idx(i).at("aDefaultECM").get_uint()
        serializeGame.aDefaultRepair(i) = o.at("data").at_idx(i).at("aDefaultRepair").get_uint()
    Next
End Sub

Private Function deserializeSaveGameV14Data(ByVal fileHandle As Long, serializeGame As SAVE_GAME) As Boolean
    Dim i As Long

    If Not deserializeSaveGameV12Data(fileHandle, serializeGame) Then Exit Function
    If Not readSBE32(fileHandle, serializeGame.missionOffTime) Then Exit Function
    If Not readSBE32(fileHandle, serializeGame.missionETA) Then Exit Function
    If Not readUBE16(fileHandle, serializeGame.missionHomeLZ_X) Then Exit Function
    If Not readUBE16(fileHandle, serializeGame.missionHomeLZ_Y) Then Exit Function
    If Not readSBE32(fileHandle, serializeGame.missionPlayerX) Then Exit Function
    If Not readSBE32(fileHandle, serializeGame.missionPlayerY) Then Exit Function

    For i = 0 To MAX_PLAYERS - 1
        If Not readUBE16(fileHandle, serializeGame.iTranspEntryTileX(i)) Then Exit Function
    Next

    For i = 0 To MAX_PLAYERS - 1
        If Not readUBE16(fileHandle, serializeGame.iTranspEntryTileY(i)) Then Exit Function
    Next

    For i = 0 To MAX_PLAYERS - 1
        If Not readUBE16(fileHandle, serializeGame.iTranspExitTileX(i)) Then Exit Function
    Next

    For i = 0 To MAX_PLAYERS - 1
        If Not readUBE16(fileHandle, serializeGame.iTranspExitTileY(i)) Then Exit Function
    Next

    For i = 0 To MAX_PLAYERS - 1
        If Not readUBE32(fileHandle, serializeGame.aDefaultSensor(i)) Then Exit Function
    Next

    For i = 0 To MAX_PLAYERS - 1
        If Not readUBE32(fileHandle, serializeGame.aDefaultECM(i)) Then Exit Function
    Next

    For i = 0 To MAX_PLAYERS - 1
        If Not readUBE32(fileHandle, serializeGame.aDefaultRepair(i)) Then Exit Function
    Next

    deserializeSaveGameV14Data = True
End Function

' SAVE_GAME_V15 : public SAVE_GAME_V14
Private Sub serializeSaveGameV15Data_json(ByVal o As nlohmann_json, serializeGame As SAVE_GAME)
    serializeSaveGameV14Data_json o, serializeGame
    o.set_item "offWorldKeepLists", json_value(serializeGame.offWorldKeepLists)
    o.set_item "RubbleTile", json_value_unsigned(serializeGame.RubbleTile)
    o.set_item "WaterTile", json_value_unsigned(serializeGame.WaterTile)

End Sub

Private Sub deserializeSaveGameV15Data_json(ByVal o As nlohmann_json, serializeGame As SAVE_GAME)
    deserializeSaveGameV14Data_json o, serializeGame
    serializeGame.offWorldKeepLists = o.at("offWorldKeepLists").get_int()
    serializeGame.RubbleTile = o.at("RubbleTile").get_uint()
    serializeGame.WaterTile = o.at("WaterTile").get_uint()
    serializeGame.fogColour = 0
    serializeGame.fogState = 0
End Sub

Private Function deserializeSaveGameV15Data(ByVal fileHandle As Long, serializeGame As SAVE_GAME) As Boolean
    Dim i As Long, j As Long, boolOffWorldKeepLists As Long, tmp As Long

    If Not deserializeSaveGameV14Data(fileHandle, serializeGame) Then Exit Function
    If Not readSBE32(fileHandle, boolOffWorldKeepLists) Then Exit Function

    serializeGame.offWorldKeepLists = boolOffWorldKeepLists

    For i = 0 To MAX_PLAYERS - 1
        For j = 0 To MAX_RECYCLED_DROIDS - 1
            If Not readUBE8(fileHandle, tmp) Then Exit Function
            If tmp > 0 Then
                add_to_experience_queue i, tmp * 65536
            End If
        Next
    Next

    If Not readUBE32(fileHandle, serializeGame.RubbleTile) Then Exit Function
    If Not readUBE32(fileHandle, serializeGame.WaterTile) Then Exit Function
    If Not readUBE32(fileHandle, serializeGame.fogColour) Then Exit Function
    deserializeSaveGameV15Data = readUBE32(fileHandle, serializeGame.fogState)
End Function

' SAVE_GAME_V16 : public SAVE_GAME_V15
Private Sub serializeSaveGameV16Data_json(ByVal o As nlohmann_json, serializeGame As SAVE_GAME)
    Dim arr As nlohmann_json, tmp As nlohmann_json, i As Long
    serializeSaveGameV15Data_json o, serializeGame
    Set arr = json_array()
    For i = 0 To MAX_NOGO_AREAS - 1
        Set tmp = json_object()
        serializeLandingZoneData_json tmp, serializeGame.sLandingZone(i)
        arr.push_back tmp
    Next
    o.set_item "landingZones", arr
End Sub

Private Sub deserializeSaveGameV16Data_json(ByVal o As nlohmann_json, serializeGame As SAVE_GAME)
    Dim i As Long
    deserializeSaveGameV15Data_json o, serializeGame
    For i = 0 To MAX_NOGO_AREAS - 1
        deserializeLandingZoneData_json o.at("landingZones").at_idx(i), serializeGame.sLandingZone(i)
    Next
End Sub

Private Function deserializeSaveGameV16Data(ByVal fileHandle As Long, serializeGame As SAVE_GAME) As Boolean
    Dim i As Long

    If Not deserializeSaveGameV15Data(fileHandle, serializeGame) Then Exit Function

    For i = 0 To MAX_NOGO_AREAS - 1
        If Not deserializeLandingZoneData(fileHandle, serializeGame.sLandingZone(i)) Then Exit Function
    Next

    deserializeSaveGameV16Data = True
End Function

' SAVE_GAME_V17 : public SAVE_GAME_V16
Private Sub serializeSaveGameV17Data_json(ByVal o As nlohmann_json, serializeGame As SAVE_GAME)
    serializeSaveGameV16Data_json o, serializeGame
    o.set_item "objId", json_value_unsigned(serializeGame.objId)
End Sub

Private Sub deserializeSaveGameV17Data_json(ByVal o As nlohmann_json, serializeGame As SAVE_GAME)
    deserializeSaveGameV16Data_json o, serializeGame
    serializeGame.objId = o.at("objId").get_uint()
End Sub

Private Function deserializeSaveGameV17Data(ByVal fileHandle As Long, serializeGame As SAVE_GAME) As Boolean
    If Not deserializeSaveGameV16Data(fileHandle, serializeGame) Then Exit Function
    deserializeSaveGameV17Data = readUBE32(fileHandle, serializeGame.objId)
End Function

' SAVE_GAME_V18 : public SAVE_GAME_V17
Private Sub serializeSaveGameV18Data_json(ByVal o As nlohmann_json, serializeGame As SAVE_GAME)
    serializeSaveGameV17Data_json o, serializeGame
    o.set_item "buildDate", json_value(serializeGame.buildDate)
    o.set_item "oldestVersion", json_value_unsigned(serializeGame.oldestVersion)
    o.set_item "validityKey", json_value_unsigned(serializeGame.validityKey)
End Sub

Private Sub deserializeSaveGameV18Data_json(ByVal o As nlohmann_json, serializeGame As SAVE_GAME)
    deserializeSaveGameV17Data_json o, serializeGame
    serializeGame.buildDate = sstrcpyN(o.at("buildDate").get_string(), MAX_STR_LENGTH)
    serializeGame.oldestVersion = o.at("oldestVersion").get_uint()
    serializeGame.validityKey = o.at("validityKey").get_uint()
End Sub

Private Function deserializeSaveGameV18Data(ByVal fileHandle As Long, serializeGame As SAVE_GAME) As Boolean
    If Not deserializeSaveGameV17Data(fileHandle, serializeGame) Then Exit Function
    If Not readChars(fileHandle, MAX_STR_LENGTH, serializeGame.buildDate) Then Exit Function
    If Not readUBE32(fileHandle, serializeGame.oldestVersion) Then Exit Function
    deserializeSaveGameV18Data = readUBE32(fileHandle, serializeGame.validityKey)
End Function

' SAVE_GAME_V19 : public SAVE_GAME_V18
Private Sub serializeSaveGameV19Data_json(ByVal o As nlohmann_json, serializeGame As SAVE_GAME)
    Dim alliancesArray As nlohmann_json, colours As nlohmann_json, i As Long, j As Long
    serializeSaveGameV18Data_json o, serializeGame
    Set alliancesArray = json_array()
    For i = 0 To MAX_PLAYERS - 1
        For j = 0 To MAX_PLAYERS - 1
            alliancesArray.push_back json_value_unsigned(serializeGame.alliances(i, j))
        Next
    Next
    o.set_item "alliances", alliancesArray

    Set colours = json_array()
    For i = 0 To MAX_PLAYERS - 1
        colours.push_back json_value_unsigned(serializeGame.playerColour(i))
    Next
    o.set_item "colours", colours
    o.set_item "radarZoom", json_value_unsigned(serializeGame.radarZoom)
End Sub

Private Sub deserializeSaveGameV19Data_json(ByVal o As nlohmann_json, serializeGame As SAVE_GAME)
    Dim i As Long, j As Long
    deserializeSaveGameV18Data_json o, serializeGame
    For i = 0 To MAX_PLAYERS - 1
        For j = 0 To MAX_PLAYERS - 1
            serializeGame.alliances(i, j) = o.at("alliances").at_idx(i * MAX_PLAYERS + j).get_uint() And &HFF&
        Next
        serializeGame.playerColour(i) = o.at("colours").at_idx(i).get_uint() And &HFF&
    Next
    serializeGame.radarZoom = o.at("radarZoom").get_uint() And &HFF&
End Sub

Private Function deserializeSaveGameV19Data(ByVal fileHandle As Long, serializeGame As SAVE_GAME) As Boolean
    Dim i As Long, j As Long

    If Not deserializeSaveGameV18Data(fileHandle, serializeGame) Then Exit Function

    For i = 0 To MAX_PLAYERS - 1
        For j = 0 To MAX_PLAYERS - 1
            If Not readUBE8(fileHandle, serializeGame.alliances(i, j)) Then Exit Function
        Next
    Next

    For i = 0 To MAX_PLAYERS - 1
        If Not readUBE8(fileHandle, serializeGame.playerColour(i)) Then Exit Function
    Next

    deserializeSaveGameV19Data = readUBE8(fileHandle, serializeGame.radarZoom)
End Function

' SAVE_GAME_V20 : public SAVE_GAME_V19
Private Sub serializeSaveGameV20Data_json(ByVal o As nlohmann_json, serializeGame As SAVE_GAME)
    Dim arr As nlohmann_json, tmp As nlohmann_json, i As Long
    serializeSaveGameV19Data_json o, serializeGame
    o.set_item "bDroidsToSafetyFlag", json_value_unsigned(serializeGame.bDroidsToSafetyFlag)
    Set arr = json_array()
    For i = 0 To MAX_PLAYERS - 1
        Set tmp = json_object()
        serializeVector2i_json tmp, serializeGame.asVTOLReturnPos(i)
        arr.push_back tmp
    Next
    o.set_item "VTOLReturnPos", arr
End Sub

Private Sub deserializeSaveGameV20Data_json(ByVal o As nlohmann_json, serializeGame As SAVE_GAME)
    Dim i As Long
    deserializeSaveGameV19Data_json o, serializeGame
    serializeGame.bDroidsToSafetyFlag = o.at("bDroidsToSafetyFlag").get_uint() And &HFF&

    For i = 0 To MAX_PLAYERS - 1
        deserializeVector2i_json o.at("VTOLReturnPos").at_idx(i), serializeGame.asVTOLReturnPos(i)
    Next
End Sub

Private Function deserializeSaveGameV20Data(ByVal fileHandle As Long, serializeGame As SAVE_GAME) As Boolean
    Dim i As Long

    If Not deserializeSaveGameV19Data(fileHandle, serializeGame) Then Exit Function
    If Not readUBE8(fileHandle, serializeGame.bDroidsToSafetyFlag) Then Exit Function

    For i = 0 To MAX_PLAYERS - 1
        If Not deserializeVector2i(fileHandle, serializeGame.asVTOLReturnPos(i)) Then Exit Function
    Next

    deserializeSaveGameV20Data = True
End Function

' SAVE_GAME_V22 : public SAVE_GAME_V20
Private Sub serializeSaveGameV22Data_json(ByVal o As nlohmann_json, serializeGame As SAVE_GAME)
    Dim arr As nlohmann_json, tmp As nlohmann_json, i As Long
    serializeSaveGameV20Data_json o, serializeGame
    Set arr = json_array()
    For i = 0 To MAX_PLAYERS - 1
        Set tmp = json_object()
        serializeRunData_json tmp, serializeGame.asRunData(i)
        arr.push_back tmp
    Next
    o.set_item "runData", arr
End Sub

Private Sub deserializeSaveGameV22Data_json(ByVal o As nlohmann_json, serializeGame As SAVE_GAME)
    Dim runData As nlohmann_json, i As Long
    deserializeSaveGameV20Data_json o, serializeGame
    If Not (o.at("runData").is_array()) Then ASSERT_FAILURE "o.at(""runData"").is_array()", "src/game.cpp:1561", "deserializeSaveGameV22Data_json", "unexpected type, wanted array": Exit Sub
    For i = 0 To MAX_PLAYERS - 1
        Set runData = o.at("runData")
        If Not (runData.is_array()) Then ASSERT_FAILURE "runData.is_array()", "src/game.cpp:1565", "deserializeSaveGameV22Data_json", "unexpected type, wanted array": Exit Sub
        deserializeRunData_json runData.at_idx(i), serializeGame.asRunData(i)

    Next
End Sub

Private Function deserializeSaveGameV22Data(ByVal fileHandle As Long, serializeGame As SAVE_GAME) As Boolean
    Dim i As Long

    If Not deserializeSaveGameV20Data(fileHandle, serializeGame) Then Exit Function

    For i = 0 To MAX_PLAYERS - 1
        If Not deserializeRunData(fileHandle, serializeGame.asRunData(i)) Then Exit Function
    Next

    deserializeSaveGameV22Data = True
End Function

' SAVE_GAME_V24 : public SAVE_GAME_V22
Private Sub serializeSaveGameV24Data_json(ByVal o As nlohmann_json, serializeGame As SAVE_GAME)
    serializeSaveGameV22Data_json o, serializeGame
    o.set_item "reinforceTime", json_value_unsigned(serializeGame.reinforceTime)
    o.set_item "bPlayCountDown", json_value_unsigned(serializeGame.bPlayCountDown)
    o.set_item "bPlayerHasWon", json_value_unsigned(serializeGame.bPlayerHasWon)
    o.set_item "bPlayerHasLost", json_value_unsigned(serializeGame.bPlayerHasLost)
End Sub

Private Sub deserializeSaveGameV24Data_json(ByVal o As nlohmann_json, serializeGame As SAVE_GAME)
    deserializeSaveGameV22Data_json o, serializeGame
    serializeGame.reinforceTime = o.at("reinforceTime").get_uint()
    serializeGame.bPlayCountDown = o.at("bPlayCountDown").get_uint() And &HFF&
    serializeGame.bPlayerHasWon = o.at("bPlayerHasWon").get_uint() And &HFF&
    serializeGame.bPlayerHasLost = o.at("bPlayerHasLost").get_uint() And &HFF&
End Sub

Private Function deserializeSaveGameV24Data(ByVal fileHandle As Long, serializeGame As SAVE_GAME) As Boolean
    If Not deserializeSaveGameV22Data(fileHandle, serializeGame) Then Exit Function
    If Not readUBE32(fileHandle, serializeGame.reinforceTime) Then Exit Function
    If Not readUBE8(fileHandle, serializeGame.bPlayCountDown) Then Exit Function
    If Not readUBE8(fileHandle, serializeGame.bPlayerHasWon) Then Exit Function
    If Not readUBE8(fileHandle, serializeGame.bPlayerHasLost) Then Exit Function
    deserializeSaveGameV24Data = readUBE8(fileHandle, serializeGame.dummy3)
End Function

' SAVE_GAME_V27 : public SAVE_GAME_V24
Private Sub serializeSaveGameV27Data_json(ByVal o As nlohmann_json, serializeGame As SAVE_GAME)
    serializeSaveGameV24Data_json o, serializeGame
End Sub

Private Sub deserializeSaveGameV27Data_json(ByVal o As nlohmann_json, serializeGame As SAVE_GAME)
    deserializeSaveGameV24Data_json o, serializeGame
End Sub

Private Function deserializeSaveGameV27Data(ByVal fileHandle As Long, serializeGame As SAVE_GAME) As Boolean
    Dim i As Long, j As Long, tmp As Long

    If Not deserializeSaveGameV24Data(fileHandle, serializeGame) Then Exit Function

    For i = 0 To MAX_PLAYERS - 1
        For j = 0 To MAX_RECYCLED_DROIDS - 1
            If Not readUBE16(fileHandle, tmp) Then Exit Function
            If tmp > 0 Then
                add_to_experience_queue i, tmp * 65536
            End If
        Next
    Next

    deserializeSaveGameV27Data = True
End Function

' SAVE_GAME_V29 : public SAVE_GAME_V27
Private Sub serializeSaveGameV29Data_json(ByVal o As nlohmann_json, serializeGame As SAVE_GAME)
    serializeSaveGameV27Data_json o, serializeGame
    o.set_item "missionScrollMinX", json_value_unsigned(serializeGame.missionScrollMinX)
    o.set_item "missionScrollMinY", json_value_unsigned(serializeGame.missionScrollMinY)
    o.set_item "missionScrollMaxX", json_value_unsigned(serializeGame.missionScrollMaxX)
    o.set_item "missionScrollMaxY", json_value_unsigned(serializeGame.missionScrollMaxY)
End Sub

Private Sub deserializeSaveGameV29Data_json(ByVal o As nlohmann_json, serializeGame As SAVE_GAME)
    deserializeSaveGameV27Data_json o, serializeGame
    serializeGame.missionScrollMinX = o.at("missionScrollMinX").get_uint() And &HFFFF&
    serializeGame.missionScrollMinY = o.at("missionScrollMinY").get_uint() And &HFFFF&
    serializeGame.missionScrollMaxX = o.at("missionScrollMaxX").get_uint() And &HFFFF&
    serializeGame.missionScrollMaxY = o.at("missionScrollMaxY").get_uint() And &HFFFF&
End Sub

Private Function deserializeSaveGameV29Data(ByVal fileHandle As Long, serializeGame As SAVE_GAME) As Boolean
    If Not deserializeSaveGameV27Data(fileHandle, serializeGame) Then Exit Function
    If Not readUBE16(fileHandle, serializeGame.missionScrollMinX) Then Exit Function
    If Not readUBE16(fileHandle, serializeGame.missionScrollMinY) Then Exit Function
    If Not readUBE16(fileHandle, serializeGame.missionScrollMaxX) Then Exit Function
    deserializeSaveGameV29Data = readUBE16(fileHandle, serializeGame.missionScrollMaxY)
End Function

' SAVE_GAME_V30 : public SAVE_GAME_V29
Private Sub serializeSaveGameV30Data_json(ByVal o As nlohmann_json, serializeGame As SAVE_GAME)
    serializeSaveGameV29Data_json o, serializeGame
    o.set_item "scrGameLevel", json_value(serializeGame.scrGameLevel)
    o.set_item "bExtraVictoryFlag", json_value_unsigned(serializeGame.bExtraVictoryFlag)
    o.set_item "bExtraFailFlag", json_value_unsigned(serializeGame.bExtraFailFlag)
    o.set_item "bTrackTransporter", json_value_unsigned(serializeGame.bTrackTransporter)
End Sub

Private Sub deserializeSaveGameV30Data_json(ByVal o As nlohmann_json, serializeGame As SAVE_GAME)
    deserializeSaveGameV29Data_json o, serializeGame
    serializeGame.scrGameLevel = o.at("scrGameLevel").get_int()
    serializeGame.bExtraVictoryFlag = o.at("bExtraVictoryFlag").get_uint() And &HFF&
    serializeGame.bExtraFailFlag = o.at("bExtraFailFlag").get_uint() And &HFF&
    serializeGame.bTrackTransporter = o.at("bTrackTransporter").get_uint() And &HFF&
End Sub

Private Function deserializeSaveGameV30Data(ByVal fileHandle As Long, serializeGame As SAVE_GAME) As Boolean
    If Not deserializeSaveGameV29Data(fileHandle, serializeGame) Then Exit Function
    If Not readSBE32(fileHandle, serializeGame.scrGameLevel) Then Exit Function
    If Not readUBE8(fileHandle, serializeGame.bExtraVictoryFlag) Then Exit Function
    If Not readUBE8(fileHandle, serializeGame.bExtraFailFlag) Then Exit Function
    deserializeSaveGameV30Data = readUBE8(fileHandle, serializeGame.bTrackTransporter)
End Function

' SAVE_GAME_V31 : public SAVE_GAME_V30 (extra code for the patch - saves out whether cheated with the mission timer)
Private Sub serializeSaveGameV31Data_json(ByVal o As nlohmann_json, serializeGame As SAVE_GAME)
    serializeSaveGameV30Data_json o, serializeGame
    o.set_item "missionCheatTime", json_value(serializeGame.missionCheatTime)
End Sub

Private Sub deserializeSaveGameV31Data_json(ByVal o As nlohmann_json, serializeGame As SAVE_GAME)
    deserializeSaveGameV30Data_json o, serializeGame
    serializeGame.missionCheatTime = o.at("missionCheatTime").get_int()
End Sub

Private Function deserializeSaveGameV31Data(ByVal fileHandle As Long, serializeGame As SAVE_GAME) As Boolean
    If Not deserializeSaveGameV30Data(fileHandle, serializeGame) Then Exit Function
    deserializeSaveGameV31Data = readSBE32(fileHandle, serializeGame.missionCheatTime)
End Function

' SAVE_GAME_V33 : public SAVE_GAME_V31 (alexl. skirmish saves)
Private Sub serializeSaveGameV33Data_json(ByVal o As nlohmann_json, serializeGame As SAVE_GAME)
    Dim arr As nlohmann_json, i As Long
    serializeSaveGameV31Data_json o, serializeGame
    serializeMultiplayerGame_json o, serializeGame.sGame
    serializeNetPlay_json o, serializeGame.sNetPlay
    o.set_item "sPname", json_value(serializeGame.sPName)
    o.set_item "multiPlayer", json_value(serializeGame.multiPlayer)
    o.set_item "savePlayer", json_value_unsigned(serializeGame.savePlayer)
    Set arr = json_array()
    For i = 0 To MAX_PLAYERS - 1
        arr.push_back json_value_unsigned(serializeGame.sPlayerIndex(i))
    Next
    o.set_item "sPlayerIndices", arr
End Sub

Private Sub deserializeSaveGameV33Data_json(ByVal o As nlohmann_json, serializeGame As SAVE_GAME)
    Dim playerIndices As nlohmann_json, i As Long
    deserializeSaveGameV31Data_json o, serializeGame
    deserializeMultiplayerGame_json o, serializeGame.sGame
    serializeGame.savePlayer = o.at("savePlayer").get_uint()
    deserializeNetPlay_json o, serializeGame.sNetPlay
    serializeGame.sPName = sstrcpyN(o.at("sPname").get_string(), 32)
    serializeGame.multiPlayer = o.at("multiPlayer").get_int()
    Set playerIndices = o.at("sPlayerIndices")
    If Not (playerIndices.is_array()) Then ASSERT_FAILURE "playerIndices.is_array()", "src/game.cpp:1864", "deserializeSaveGameV33Data_json", "unexpected type, wanted array": Exit Sub
    For i = 0 To MAX_PLAYERS - 1
        serializeGame.sPlayerIndex(i) = playerIndices.at_idx(i).get_uint()
    Next
End Sub

Private Function deserializeSaveGameV33Data(ByVal fileHandle As Long, serializeGame As SAVE_GAME) As Boolean
    Dim i As Long, boolMultiPlayer As Long

    If Not deserializeSaveGameV31Data(fileHandle, serializeGame) Then Exit Function
    If Not deserializeMultiplayerGame(fileHandle, serializeGame.sGame) Then Exit Function
    If Not deserializeNetPlay(fileHandle, serializeGame.sNetPlay) Then Exit Function
    If Not readUBE32(fileHandle, serializeGame.savePlayer) Then Exit Function
    If Not readChars(fileHandle, 32, serializeGame.sPName) Then Exit Function
    If Not readSBE32(fileHandle, boolMultiPlayer) Then Exit Function

    serializeGame.multiPlayer = boolMultiPlayer

    For i = 0 To MAX_PLAYERS - 1
        If Not readUBE32(fileHandle, serializeGame.sPlayerIndex(i)) Then Exit Function
    Next

    deserializeSaveGameV33Data = True
End Function

' SAVE_GAME_V34 : public SAVE_GAME_V33 (Now holds AI names for multiplayer)
Private Sub serializeSaveGameV34Data_json(ByVal o As nlohmann_json, serializeGame As SAVE_GAME)
    Dim arr As nlohmann_json, i As Long
    serializeSaveGameV33Data_json o, serializeGame
    Set arr = json_array()
    For i = 0 To MAX_PLAYERS - 1
        arr.push_back json_value(serializeGame.sPlayerName(i))
    Next
    o.set_item "playerNames", arr
End Sub

Private Sub deserializeSaveGameV34Data_json(ByVal o As nlohmann_json, serializeGame As SAVE_GAME)
    Dim i As Long
    deserializeSaveGameV33Data_json o, serializeGame
    For i = 0 To MAX_PLAYERS - 1
        serializeGame.sPlayerName(i) = sstrcpyN(o.at("playerNames").at_idx(i).get_string(), StringSize)
    Next
End Sub

Private Function deserializeSaveGameV34Data(ByVal fileHandle As Long, serializeGame As SAVE_GAME) As Boolean
    Dim i As Long
    If Not deserializeSaveGameV33Data(fileHandle, serializeGame) Then Exit Function

    For i = 0 To MAX_PLAYERS - 1
        If Not readChars(fileHandle, StringSize, serializeGame.sPlayerName(i)) Then Exit Function
    Next

    deserializeSaveGameV34Data = True
End Function

' SAVE_GAME_V35 : public SAVE_GAME_V34 (First version to utilize (de)serialization API and first to be big-endian
' (instead of little-endian))
Private Sub serializeSaveGameV35Data_json(ByVal o As nlohmann_json, serializeGame As SAVE_GAME)
    serializeSaveGameV34Data_json o, serializeGame
End Sub

Private Sub deserializeSaveGameV35Data_json(ByVal o As nlohmann_json, serializeGame As SAVE_GAME)
    deserializeSaveGameV34Data_json o, serializeGame
End Sub

Private Function deserializeSaveGameV35Data(ByVal fileHandle As Long, serializeGame As SAVE_GAME) As Boolean
    deserializeSaveGameV35Data = deserializeSaveGameV34Data(fileHandle, serializeGame)
End Function

' SAVE_GAME_V38 : public SAVE_GAME_V35 (Store loaded mods in savegame)
Private Sub serializeSaveGameV38Data_json(ByVal o As nlohmann_json, serializeGame As SAVE_GAME)
    serializeSaveGameV35Data_json o, serializeGame
    o.set_item "modList", json_value(serializeGame.modList)
End Sub

Private Sub deserializeSaveGameV38Data_json(ByVal o As nlohmann_json, serializeGame As SAVE_GAME)
    deserializeSaveGameV35Data_json o, serializeGame
    serializeGame.modList = sstrcpyN(o.at("modList").get_string(), modlist_string_size)
End Sub

Private Function deserializeSaveGameV38Data(ByVal fileHandle As Long, serializeGame As SAVE_GAME) As Boolean
    If Not deserializeSaveGameV35Data(fileHandle, serializeGame) Then Exit Function

    If Not readChars(fileHandle, modlist_string_size, serializeGame.modList) Then Exit Function

    deserializeSaveGameV38Data = True
End Function

' Current save game version (typedef SAVE_GAME_V38 SAVE_GAME)
Private Sub serializeSaveGameData_json(ByVal o As nlohmann_json, ByVal saveinfo As nlohmann_json, ByVal saveName As String, serializeGame As SAVE_GAME)
    Dim tagVersion() As Long, tagQualifier As String, tagHasValue As Boolean, latest As nlohmann_json
    Dim epoch As Double, ourtime As String, curtime As Double, timeinfo As tm
    serializeSaveGameV38Data_json o, serializeGame
    ' not sure whether its 38, 39 or 40... different .cpp files are using different numbers
    o.set_item "version", json_value(VERSION_39)

    ' This file lists saved games, and their build info
    ' one per savegame directory
    tagHasValue = version_extractVersionNumberFromTag(version_getLatestTag(), tagVersion, tagQualifier)
    If Not (tagHasValue) Then ASSERT_FAILURE "tagResult.has_value()", "src/game.cpp:2044", "serializeSaveGameData_json", "No extractable latest tag?? - Please try re-downloading the latest official source bundle"
    If Not tagHasValue Then             ' tagResult.value_or(TagVer())
        tagVersion(0) = 0
        tagVersion(1) = 0
        tagVersion(2) = 0
    End If
    Set latest = json_array()
    latest.push_back json_value_unsigned(tagVersion(0))
    latest.push_back json_value_unsigned(tagVersion(1))
    latest.push_back json_value_unsigned(tagVersion(2))
    saveinfo.set_item "latestTagArray", latest
    epoch = time_()                     ' std::chrono::system_clock::now().time_since_epoch(), in seconds
    saveinfo.set_item "epoch", json_value(CDec(Int(epoch)))
    If saveinfo.contains(saveName) Then
        saveinfo.erase_ saveName
    End If
    curtime = time_()
    timeinfo = getLocalTime(curtime)
    ' "YYYY-MM-DD HH:MM:SS" (%F %T)
    ourtime = strftime("%Y-%m-%d %H:%M:%S", timeinfo)
    saveinfo.set_item "time", json_value(ourtime)
End Sub

Public Function deserializeSaveGameData_json(ByVal o As nlohmann_json, serializeGame As SAVE_GAME) As Boolean
    On Error GoTo catch_
    deserializeSaveGameV38Data_json o, serializeGame
    deserializeSaveGameData_json = True
    Exit Function
catch_:
    debug_ 2072, LOG_ERROR, "deserializeSaveGameData_json", "%s", Err.Description
    deserializeSaveGameData_json = False

End Function

Public Function deserializeSaveGameData(ByVal fileHandle As Long, serializeGame As SAVE_GAME) As Boolean
    deserializeSaveGameData = deserializeSaveGameV38Data(fileHandle, serializeGame)
End Function

'--------------------------------------------------------------------------

Private Sub sanityUpdate()
    Dim player As Long, it As Long, psDroid As Long
    For player = 0 To game.maxPlayers - 1
        it = apsDroidLists(player).head
        Do While it <> 0
            psDroid = lnObj(it)
            orderCheckList psDroid
            actionSanity psDroid
            it = lnNext(it)
        Loop
    Next
End Sub

' -----------------------------------------------------------------------------------------
' (VB) CurrentFileName[fileExtension] = '\0'; strcat(CurrentFileName, name): the file name's first
' fileExtension characters followed by name
Private Function saveGameFile(ByVal CurrentFileName As String, ByVal fileExtension As Long, ByVal name_ As String) As String
    saveGameFile = Left$(CurrentFileName, fileExtension) & name_
End Function

Public Function saveGame(ByVal aFileName As String, ByVal saveType As Long, Optional ByVal isAutoSave As Boolean = False) As Boolean
    Dim fileExtension As Long, CurrentFileName As String, it As Long, psDroid As Long

    ' executeFnAndProcessScriptQueuedRemovals([]() { triggerEvent(TRIGGER_GAME_SAVING); })
    triggerEvent TRIGGER_GAME_SAVING
    wzapi_processScriptQueuedObjectRemovals

    If Not (Len(aFileName) > 4) Then ASSERT_FAILURE "aFileName && strlen(aFileName) > 4", "src/game.cpp:3549", "saveGame", "Bad savegame filename": saveGame = False: Exit Function
    CurrentFileName = Left$(aFileName, PATH_MAX - 1)    ' sstrcpy(CurrentFileName, aFileName)
    debug_ 3551, LOG_WZ, "saveGame", "saveGame: %s", CurrentFileName

    fileExtension = Len(CurrentFileName) - 3
    gameTimeStop
    sanityUpdate

    ' Write the data to the file
    If Not writeGameFile(CurrentFileName, saveType) Then
        debug_ 3560, LOG_ERROR, "saveGame", "writeGameFile(""%s"") failed", CurrentFileName
        GoTo error_
    End If

    'remove the file extension
    CurrentFileName = Left$(CurrentFileName, Len(CurrentFileName) - 4)

    'create dir will fail if directory already exists but don't care!
    PHYSFS_mkdir CurrentFileName

    writeMainFile CurrentFileName & "/main.json", saveType

    'save the map file
    CurrentFileName = CurrentFileName & "/game.map"
    ' Write the data to the file
    If Not writeMapFile(CurrentFileName) Then
        debug_ 3577, LOG_ERROR, "saveGame", "saveGame: writeMapFile(""%s"") failed", CurrentFileName
        GoTo error_
    End If

    ' Save some game info
    CurrentFileName = saveGameFile(CurrentFileName, fileExtension, "gameinfo.json")
    writeGameInfo CurrentFileName

    ' Save labels
    CurrentFileName = saveGameFile(CurrentFileName, fileExtension, "labels.json")
    writeLabels CurrentFileName

    'create the droids filename
    CurrentFileName = saveGameFile(CurrentFileName, fileExtension, "droid.json")
    'Write the current droid lists to the file
    If Not writeDroidFile(CurrentFileName, apsDroidLists) Then
        debug_ 3597, LOG_ERROR, "saveGame", "writeDroidFile(""%s"") failed", CurrentFileName
        GoTo error_
    End If

    'create the structures filename
    CurrentFileName = saveGameFile(CurrentFileName, fileExtension, "struct.json")
    'Write the data to the file
    If Not writeStructFile(CurrentFileName) Then
        debug_ 3607, LOG_ERROR, "saveGame", "saveGame: writeStructFile(""%s"") failed", CurrentFileName
        GoTo error_
    End If

    'create the templates filename
    CurrentFileName = saveGameFile(CurrentFileName, fileExtension, "templates.json")
    'Write the data to the file
    If Not writeTemplateFile(CurrentFileName) Then
        debug_ 3617, LOG_ERROR, "saveGame", "saveGame: writeTemplateFile(""%s"") failed", CurrentFileName
        GoTo error_
    End If

    'create the features filename
    CurrentFileName = saveGameFile(CurrentFileName, fileExtension, "feature.json")
    'Write the data to the file
    If Not writeFeatureFile(CurrentFileName) Then
        debug_ 3627, LOG_ERROR, "saveGame", "saveGame: writeFeatureFile(""%s"") failed", CurrentFileName
        GoTo error_
    End If

    'create the terrain types filename
    CurrentFileName = saveGameFile(CurrentFileName, fileExtension, "ttypes.ttp")
    'Write the data to the file
    If Not writeTerrainTypeMapFile(CurrentFileName) Then
        debug_ 3637, LOG_ERROR, "saveGame", "saveGame: writeTerrainTypeMapFile(""%s"") failed", CurrentFileName
        GoTo error_
    End If

    'create the strucutLimits filename
    CurrentFileName = saveGameFile(CurrentFileName, fileExtension, "limits.json")
    'Write the data to the file
    If Not writeStructLimitsFile(CurrentFileName) Then
        debug_ 3647, LOG_ERROR, "saveGame", "saveGame: writeStructLimitsFile(""%s"") failed", CurrentFileName
        GoTo error_
    End If

    'create the component lists filename
    CurrentFileName = saveGameFile(CurrentFileName, fileExtension, "complist.json")
    'Write the data to the file
    If Not writeCompListFile(CurrentFileName) Then
        debug_ 3657, LOG_ERROR, "saveGame", "saveGame: writeCompListFile(""%s"") failed", CurrentFileName
        GoTo error_
    End If
    'create the structure type lists filename
    CurrentFileName = saveGameFile(CurrentFileName, fileExtension, "strtype.json")
    'Write the data to the file
    If Not writeStructTypeListFile(CurrentFileName) Then
        debug_ 3666, LOG_ERROR, "saveGame", "saveGame: writeStructTypeListFile(""%s"") failed", CurrentFileName
        GoTo error_
    End If

    'create the research filename
    CurrentFileName = saveGameFile(CurrentFileName, fileExtension, "resstate.json")
    'Write the data to the file
    If Not writeResearchFile(CurrentFileName) Then
        debug_ 3676, LOG_ERROR, "saveGame", "saveGame: writeResearchFile(""%s"") failed", CurrentFileName
        GoTo error_
    End If

    'create the message filename
    CurrentFileName = saveGameFile(CurrentFileName, fileExtension, "messtate.json")
    'Write the data to the file
    If Not writeMessageFile(CurrentFileName) Then
        debug_ 3686, LOG_ERROR, "saveGame", "saveGame: writeMessageFile(""%s"") failed", CurrentFileName
        GoTo error_
    End If

    CurrentFileName = saveGameFile(CurrentFileName, fileExtension, "visstate.bjo")
    'Write the data to the file
    If Not writeVisibilityData(CurrentFileName) Then
        debug_ 3695, LOG_ERROR, "saveGame", "saveGame: writeVisibilityData(""%s"") failed", CurrentFileName
        GoTo error_
    End If

    CurrentFileName = saveGameFile(CurrentFileName, fileExtension, "fxstate.json")
    'Write the data to the file
    If Not writeFXData(CurrentFileName) Then
        debug_ 3704, LOG_ERROR, "saveGame", "saveGame: writeFXData(""%s"") failed", CurrentFileName
        GoTo error_
    End If

    'added at V15 save
    CurrentFileName = saveGameFile(CurrentFileName, fileExtension, "score.json")
    'Write the data to the file
    If Not writeScoreData(CurrentFileName) Then
        debug_ 3714, LOG_ERROR, "saveGame", "saveGame: writeScoreData(""%s"") failed", CurrentFileName
        GoTo error_
    End If

    CurrentFileName = saveGameFile(CurrentFileName, fileExtension, "firesupport.json")
    'Write the data to the file
    If Not writeFiresupportDesignators(CurrentFileName) Then
        debug_ 3723, LOG_ERROR, "saveGame", "saveGame: writeFiresupportDesignators(""%s"") failed", CurrentFileName
        GoTo error_
    End If

    ' save the script state if necessary
    If saveType = GTYPE_SAVE_MIDMISSION Then
        CurrentFileName = saveGameFile(CurrentFileName, fileExtension, "scriptstate.es")
        'Write the data to the file
        If Not writeScriptState(CurrentFileName) Then
            debug_ 3735, LOG_ERROR, "saveGame", "saveGame: writeScriptState(""%s"") failed", CurrentFileName
            GoTo error_
        End If
    End If

    ' save the loaded guide topics
    CurrentFileName = saveGameFile(CurrentFileName, fileExtension, "guidetopics.json")
    writeSaveGuideTopics CurrentFileName

    'create the droids filename
    CurrentFileName = saveGameFile(CurrentFileName, fileExtension, "mdroid.json")
    'Write the swapped droid lists to the file
    If Not writeDroidFile(CurrentFileName, mission_apsDroidLists) Then
        debug_ 3751, LOG_ERROR, "saveGame", "writeDroidFile(""%s"") failed", CurrentFileName
        GoTo error_
    End If

    'create the limbo filename
    'clear the list
    If ULT(saveGameVersionNow(), 25) Then   ' saveGameVersion < VERSION_25
        If Not (ULT(selectedPlayer, MAX_PLAYERS)) Then ASSERT_FAILURE "selectedPlayer < apsLimboDroids.size()", "src/game.cpp:3759", "saveGame", "selectedPlayer is out of bounds: %u", selectedPlayer
        If ULT(selectedPlayer, MAX_PLAYERS) Then
            it = apsLimboDroids(selectedPlayer).head
            Do While it <> 0
                psDroid = lnObj(it)
                'limbo list invalidate XY
                aDroid(psDroid).pos.x = INVALID_XY
                aDroid(psDroid).pos.y = INVALID_XY
                'this is mainly for VTOLs
                aDroid(psDroid).psBaseStruct = 0    ' setSaveDroidBase(psDroid, nullptr) (droid.h)
                orderDroid psDroid, DORDER_STOP, ModeImmediate
                it = lnNext(it)
            Loop
        End If
    End If

    CurrentFileName = saveGameFile(CurrentFileName, fileExtension, "limbo.json")
    'Write the swapped droid lists to the file
    If Not writeDroidFile(CurrentFileName, apsLimboDroids) Then
        debug_ 3779, LOG_ERROR, "saveGame", "saveGame: writeDroidFile(""%s"") failed", CurrentFileName
        GoTo error_
    End If

    If saveGameOnMissionNow() Then
        'mission save swap the mission pointers and save the changes
        swapMissionPointers
        'now save the map and droids

        'save the map file
        CurrentFileName = saveGameFile(CurrentFileName, fileExtension, "mission.map")
        ' Write the data to the file
        If Not writeMapFile(CurrentFileName) Then
            debug_ 3795, LOG_ERROR, "saveGame", "saveGame: writeMapFile(""%s"") failed", CurrentFileName
            GoTo error_
        End If

        'save the map file
        CurrentFileName = saveGameFile(CurrentFileName, fileExtension, "misvis.bjo")
        ' Write the data to the file
        If Not writeVisibilityData(CurrentFileName) Then
            debug_ 3805, LOG_ERROR, "saveGame", "saveGame: writeVisibilityData(""%s"") failed", CurrentFileName
            GoTo error_
        End If

        'create the structures filename
        CurrentFileName = saveGameFile(CurrentFileName, fileExtension, "mstruct.json")
        'Write the data to the file
        If Not writeStructFile(CurrentFileName) Then
            debug_ 3815, LOG_ERROR, "saveGame", "saveGame: writeStructFile(""%s"") failed", CurrentFileName
            GoTo error_
        End If

        'create the features filename
        CurrentFileName = saveGameFile(CurrentFileName, fileExtension, "mfeature.json")
        'Write the data to the file
        If Not writeFeatureFile(CurrentFileName) Then
            debug_ 3825, LOG_ERROR, "saveGame", "saveGame: writeFeatureFile(""%s"") failed", CurrentFileName
            GoTo error_
        End If

        'mission save swap back so we can restart the game
        swapMissionPointers
    End If

    ' strip the last filename
    CurrentFileName = Left$(CurrentFileName, fileExtension - 1)

    ' Start the game clock
    triggerEvent TRIGGER_GAME_SAVED
    gameTimeStart
    saveGame = True
    Exit Function

error_:
    ' Start the game clock
    gameTimeStart

    saveGame = False
End Function

' -----------------------------------------------------------------------------------------
Private Function writeMapFile(ByVal fileName As String) As Boolean
    Dim mapData As MapData, status As Boolean, mapIO As WzMapPhysFSIO, debugLoggerInstance As WzMapDebugLogger
    If Not (Len(fileName) <> 0) Then ASSERT_FAILURE "fileName != nullptr", "src/game.cpp:3855", "writeMapFile", "filename is null": writeMapFile = False: Exit Function

    ' Get the save data
    Set mapData = Map_new().mapData()   ' WzMap::MapData mapData; (VB: MapData is not creatable outside WzMap: an empty Map's)
    status = mapSaveToWzMapData(mapData)
    If Not status Then
        writeMapFile = False
        Exit Function
    End If

    ' Write out the map data
    Set mapIO = New WzMapPhysFSIO
    Set debugLoggerInstance = New WzMapDebugLogger
    status = writeMapData(mapData, fileName, mapIO, OutputFormat_VER3, debugLoggerInstance)  ' WzMap::LatestOutputFormat

    writeMapFile = status
End Function

' -----------------------------------------------------------------------------------------
' Load main game data from JSON. Only implement stuff here that we actually use instead of
' the binary blobbery.
Public Function loadMainFile(ByVal fileName As String) As Boolean
    Dim save As WzConfig, saveCampaignName As String, saveScavValue As Long, index As Long, FactionValue As Long
    Dim rootObj As nlohmann_json, arr As nlohmann_json, idx As Long, n As Long
    Set save = WzConfig_new(fileName, WzConfig_ReadOnly)

    If save.contains("gameType") Then
        game.typ = save.value("gameType").toInt()
    End If
    If save.contains("campaignName") Then
        saveCampaignName = save.value("campaignName").toWzString()
        If Len(saveCampaignName) <> 0 Then
            setCampaignName saveCampaignName
        Else
            clearCampaignName
        End If
    End If
    If save.contains("scavengers") Then
        saveScavValue = save.value("scavengers").toUInt()
        If Not ULT(ULTIMATE_SCAVENGERS, saveScavValue) Then
            game.scavengers = saveScavValue And &HFF&
        Else
            debug_ 4673, LOG_ERROR, "loadMainFile", "Invalid scavengers value: %u", saveScavValue
        End If
    End If
    If save.contains("maxPlayers") Then
        game.maxPlayers = save.value("maxPlayers").toUInt() And &HFF&
    End If
    If save.contains("mapHasScavengers") Then
        game.mapHasScavengers = save.value("mapHasScavengers").toBool()
    End If
    If save.contains("playerBuiltHQ") Then
        playerBuiltHQ = save.value("playerBuiltHQ").toBool()
    End If
    If save.contains("challengeFileName") Then
        challengeFileName = save.string_("challengeFileName")
    End If
    If save.contains("challengeActive") Then
        challengeActive = save.value("challengeActive").toBool()
    End If
    If save.contains("builtInMap") Then
        builtInMap = save.value("builtInMap").toBool()
    End If
    If save.contains("useTerrainOverrides") Then
        useTerrainOverrides = save.value("useTerrainOverrides").toBool()
    End If
    If save.contains("inactivityMinutes") Then
        game.inactivityMinutes = save.value("inactivityMinutes").toUInt()
    End If
    If save.contains("gameTimeLimitMinutes") Then
        game.gameTimeLimitMinutes = save.value("gameTimeLimitMinutes").toUInt()
    End If
    If save.contains("playerLeaveMode") Then
        game.playerLeaveMode = save.value("playerLeaveMode").toInt()
    End If
    If save.contains("blindMode") Then
        game.blindMode = save.value("blindMode").toInt()
    End If
    If save.contains("multiplayer") Then
        bMultiPlayer = save.value("multiplayer").toBool()
    End If
    If save.contains("tweakOptions") Then
        ' std::unordered_map<std::string, nlohmann::json> tweakOptionsJSONMap of the object's items
        setCamTweakOptions save.value("tweakOptions").jsonValue()
    End If

    save.beginArray "players"
    Do While save.remainingArrayItems() > 0
        index = save.value("index").toInt()
        If Not (index >= 0 And index < MAX_PLAYERS) Then
            debug_ 4741, LOG_ERROR, "loadMainFile", "Invalid player index: %d", index
            save.nextArrayItem
        Else
            FactionValue = save.value("faction", json_variant_new(CLng(FACTION_NORMAL))).toUInt()
            NetPlay_players_faction(index) = FactionValue
            save.nextArrayItem
        End If
    Loop
    save.endArray

    Set rootObj = save.currentJsonValue()
    If rootObj.contains("scriptSetPlayerDataStrings") Then
        Set arr = rootObj.at("scriptSetPlayerDataStrings")
        If arr.is_array() Then
            n = arr.size()
            If NetPlay_scriptSetPlayerDataStrings_size() < n Then n = NetPlay_scriptSetPlayerDataStrings_size()
            For idx = 0 To n - 1
                NetPlay_scriptSetPlayerDataStrings_assign idx, arr.at_idx(idx)
            Next
        End If
    End If

    loadMainFile = True
End Function

Public Function loadMainFileFinal(ByVal fileName As String) As Boolean
    Dim save As WzConfig, index As Long, value As nlohmann_json, k As Long
    Set save = WzConfig_new(fileName, WzConfig_ReadOnly)

    If save.contains("techLevel") Then
        game.techLevel = save.value("techLevel").toInt()
    End If

    save.beginArray "players"
    Do While save.remainingArrayItems() > 0
        index = save.value("index").toInt()
        If Not (index >= 0 And index < MAX_PLAYERS) Then
            debug_ 4780, LOG_ERROR, "loadMainFileFinal", "Invalid player index: %d", index
            save.nextArrayItem
        Else
            Set value = save.value("recycled_droids").jsonValue()
            For k = 0 To value.size() - 1
                add_to_experience_queue index, json_variant_new(value.value_at(k)).toInt()
            Next
            setMultiPlayRecentScore index, save.value("recentScore", json_variant_new(0&)).toUInt()
            setMultiPlayUnitsKilled index, save.value("recentKills", json_variant_new(0&)).toUInt()
            setMultiPlayRecentDroidsKilled index, save.value("recentDroidsKilled", json_variant_new(0&)).toUInt()
            setMultiPlayRecentDroidsLost index, save.value("recentDroidsLost", json_variant_new(0&)).toUInt()
            setMultiPlayRecentDroidsBuilt index, save.value("recentDroidsBuilt", json_variant_new(0&)).toUInt()
            setMultiPlayRecentStructuresKilled index, save.value("recentStructuresKilled", json_variant_new(0&)).toUInt()
            setMultiPlayRecentStructuresLost index, save.value("recentStructuresLost", json_variant_new(0&)).toUInt()
            setMultiPlayRecentStructuresBuilt index, save.value("recentStructuresBuilt", json_variant_new(0&)).toUInt()
            setMultiPlayRecentPowerLost index, save.value("recentPowerLost", json_variant_new(0&)).toUint64()
            setMultiPlayRecentDroidPowerLost index, save.value("recentDroidPowerLost", json_variant_new(0&)).toUint64()
            setMultiPlayRecentStructurePowerLost index, save.value("recentStructurePowerLost", json_variant_new(0&)).toUint64()
            setMultiPlayRecentPowerWon index, save.value("recentPowerWon", json_variant_new(0&)).toUint64()
            setMultiPlayRecentResearchComplete index, save.value("recentResearchComplete", json_variant_new(0&)).toUInt()
            setMultiPlayRecentResearchPotential index, save.value("recentResearchPotential", json_variant_new(0&)).toUint64()
            setMultiPlayRecentResearchPerformance index, save.value("recentResearchPerformance", json_variant_new(0&)).toUint64()
            save.nextArrayItem
        End If
    Loop
    save.endArray

    loadMainFileFinal = True
End Function

' (VB) MAX(unsynchObjID * 2, (synchObjID + 3) / 4) in uint32_t: the save's objId
Private Function saveObjId() As Double
    Dim a As Double, b As Double
    a = U32(unsynchObjID) * 2
    If a >= 4294967296# Then a = a - 4294967296#
    b = U32(synchObjID) + 3
    If b >= 4294967296# Then b = b - 4294967296#
    b = Int(b / 4)
    If a > b Then saveObjId = a Else saveObjId = b
End Function

' -----------------------------------------------------------------------------------------
' Save main game data to JSON. We save more here than we need to, and duplicate some of the
' binary blobbery, for future usage.
Private Function writeMainFile(ByVal fileName As String, ByVal saveType As Long) As Boolean
    Dim save As WzConfig, saveKey As Long, i As Long, j As Long, n As Long, k As Long
    Dim experience() As Long, recycled_droids As nlohmann_json, allies_ As nlohmann_json
    Dim currPlayerPos As iView, psLandingZone As LANDING_ZONE, st As PLAYERSTATS_

    If Not (saveType = GTYPE_SAVE_START Or saveType = GTYPE_SAVE_MIDMISSION) Then ASSERT_FAILURE "saveType == GTYPE_SAVE_START || saveType == GTYPE_SAVE_MIDMISSION", "src/game.cpp:4816", "writeMainFile", "invalid save type"

    Set save = WzConfig_new(fileName, WzConfig_ReadAndWrite)

    saveKey = getCampaignNumber()
    If missionIsOffworld() Then
        saveKey = saveKey Or SAVEKEY_ONMISSION
        setSaveGameOnMission True
    Else
        setSaveGameOnMission False
    End If

    ' Put the save game data into the buffer
    save.setValue "version", json_value_unsigned(currentGameVersion) ' game version save was made on
    save.setValue "versionFile", json_value(2&)    ' version of this file. Bump for significant changes
    save.setValue "saveKey", json_value_unsigned(saveKey)
    save.setValue "campaignName", json_value(getCampaignName())
    save.setValue "gameTime", json_value_unsigned(gameTime)
    save.setValue "missionTime", json_value_unsigned(mission_startTime)
    save.setVector2i "scrollMin", scrollMinX, scrollMinY
    save.setVector2i "scrollMax", scrollMaxX, scrollMaxY
    save.setValue "saveType", json_value(saveType)
    If Not (Len(aLevelName) < MAX_LEVEL_SIZE) Then ASSERT_FAILURE "strlen(aLevelName) < MAX_LEVEL_SIZE", "src/game.cpp:4841", "writeMainFile", "Unable to save level name - too long (max %d) - %s", MAX_LEVEL_SIZE, aLevelName: writeMainFile = False: Exit Function
    save.setValue "levelName", json_value(aLevelName)
    save.setValue "radarPermitted", json_value(radarPermitted)
    save.setValue "allowDesign", json_value(allowDesign)
    save.setValue "missionOffTime", json_value(mission_time)
    save.setValue "missionETA", json_value(mission_ETA)
    save.setValue "missionCheatTime", json_value_unsigned(CLng(mission_cheatTime))
    save.setVector2i "missionHomeLZ", mission_homeLZ_X, mission_homeLZ_Y
    save.setVector2i "missionPlayerPos", mission_playerX, mission_playerY
    save.setVector2i "missionScrollMin", mission_scrollMinX, mission_scrollMinY
    save.setVector2i "missionScrollMax", mission_scrollMaxX, mission_scrollMaxY
    save.setValue "offWorldKeepLists", json_value(offWorldKeepLists)
    save.setValue "rubbleTile", json_value_unsigned(getRubbleTileNum())
    save.setValue "waterTile", json_value_unsigned(getWaterTileNum())
    save.setValue "objId", json_value_unsigned(saveObjId())
    save.setValue "radarZoom", json_value_unsigned(GetRadarZoom())
    save.setValue "droidsToSafetyFlag", json_value(getDroidsToSafetyFlag())
    save.setValue "reinforceTime", json_value_unsigned(missionGetReinforcementTime())
    save.setValue "playCountDown", json_value(getPlayCountDown())

    save.beginArray "players"
    For i = 0 To MAX_PLAYERS - 1
        save.setValue "index", json_value(i)
        save.setValue "power", json_value_unsigned(getPower(i))
        save.setVector2i "iTranspEntryTile", mission_iTranspEntryTileX(i), mission_iTranspEntryTileY(i)
        save.setVector2i "iTranspExitTile", mission_iTranspExitTileX(i), mission_iTranspExitTileY(i)
        save.setValue "aDefaultSensor", json_value_unsigned(aDefaultSensor(i))
        save.setValue "aDefaultECM", json_value_unsigned(aDefaultECM(i))
        save.setValue "aDefaultRepair", json_value_unsigned(aDefaultRepair(i))

        ' std::priority_queue<int> experience = copy_experience_queue(i): its pops, top first
        n = copy_experience_queue(i, experience)
        Set recycled_droids = json_array()
        For k = 0 To n - 1
            recycled_droids.push_back json_value(experience(k))
        Next
        save.set_ "recycled_droids", recycled_droids

        Set allies_ = json_array()
        For j = 0 To MAX_PLAYERS - 1
            allies_.push_back json_value_unsigned(alliances(i, j))
        Next
        save.set_ "alliances", allies_
        save.setValue "difficulty", json_value(NetPlay_players_difficulty(i))
        save.setValue "position", json_value(NetPlay_players_position(i))
        save.setValue "colour", json_value_unsigned(getPlayerColour(i))
        save.setValue "allocated", json_value(NetPlay_players_allocated(i))
        save.setValue "faction", json_value_unsigned(NetPlay_players_faction(i))
        save.setValue "team", json_value(NetPlay_players_team(i))
        save.setValue "ai", json_value(NetPlay_players_ai(i))
        save.setValue "autoGame", json_value(NetPlay_players_autoGame(i))
        save.setValue "ip", json_value(NetPlay_players_IPtextAddress(i))
        save.setValue "name", json_value(getPlayerName(i))
        save.setValue "nameAI", json_value(getAIName(i))
        save.setValue "recentScore", json_value_unsigned(getMultiPlayRecentScore(i))
        save.setValue "recentKills", json_value_unsigned(getMultiPlayUnitsKilled(i))
        st = getMultiStats(i)
        save.setValue "recentDroidsKilled", json_value_unsigned(st.recentDroidsKilled)
        save.setValue "recentDroidsLost", json_value_unsigned(st.recentDroidsLost)
        save.setValue "recentDroidsBuilt", json_value_unsigned(st.recentDroidsBuilt)
        save.setValue "recentStructuresKilled", json_value_unsigned(st.recentStructuresKilled)
        save.setValue "recentStructuresLost", json_value_unsigned(st.recentStructuresLost)
        save.setValue "recentStructuresBuilt", json_value_unsigned(st.recentStructuresBuilt)
        save.setValue "recentPowerLost", json_value_unsigned(st.recentPowerLost)
        save.setValue "recentDroidPowerLost", json_value_unsigned(st.recentDroidPowerLost)
        save.setValue "recentStructurePowerLost", json_value_unsigned(st.recentStructurePowerLost)
        save.setValue "recentPowerWon", json_value_unsigned(st.recentPowerWon)
        save.setValue "recentResearchComplete", json_value_unsigned(st.recentResearchComplete)
        save.setValue "recentResearchPotential", json_value_unsigned(st.recentResearchPotential)
        save.setValue "recentResearchPerformance", json_value_unsigned(st.recentResearchPerformance)

        save.nextArrayItem
    Next
    save.endArray

    disp3d_getView currPlayerPos
    save.setVector3i "camera_position", currPlayerPos.p.x, currPlayerPos.p.y, currPlayerPos.p.z
    save.setVector3i "camera_rotation", currPlayerPos.r.x, currPlayerPos.r.y, currPlayerPos.r.z

    save.beginArray "landing_zones"
    For i = 0 To MAX_NOGO_AREAS - 1
        psLandingZone = getLandingZone(i)
        save.setVector2i "start", psLandingZone.x1, psLandingZone.y1
        save.setVector2i "end", psLandingZone.x2, psLandingZone.y2
        save.nextArrayItem
    Next
    save.endArray

    save.setValue "playerHasWon", json_value(testPlayerHasWon())
    save.setValue "playerHasLost", json_value(testPlayerHasLost())
    save.setValue "gameType", json_value(game.typ)
    save.setValue "scavengers", json_value_unsigned(game.scavengers)
    save.setValue "mapName", json_value(game.map)
    save.setValue "maxPlayers", json_value_unsigned(game.maxPlayers)
    save.setValue "gameName", json_value(game.name_)
    save.setValue "powerSetting", json_value_unsigned(game.power)
    save.setValue "baseSetting", json_value_unsigned(game.base)
    save.setValue "allianceSetting", json_value_unsigned(game.alliance)
    save.setValue "mapHasScavengers", json_value(game.mapHasScavengers)
    save.setValue "mapMod", json_value(game.isMapMod)
    save.setValue "mapRandom", json_value(game.isRandom)
    save.setValue "techLevel", json_value_unsigned(game.techLevel)
    save.setValue "gameHash", json_value(Sha256_toString(game.hash))
    save.setValue "selectedPlayer", json_value_unsigned(selectedPlayer)
    save.setValue "multiplayer", json_value(bMultiPlayer)
    save.setValue "playerCount", json_value_unsigned(NetPlay_playercount())
    save.setValue "hostPlayer", json_value_unsigned(NetPlay_hostPlayer)
    save.setValue "bComms", json_value(NetPlay_bComms)
    save.setValue "modList", json_value(getModList())
    save.setValue "playerBuiltHQ", json_value(playerBuiltHQ)
    save.setValue "challengeFileName", json_value(challengeFileName)
    save.setValue "challengeActive", json_value(challengeActive)
    save.setValue "builtInMap", json_value(builtInMap)
    save.setValue "useTerrainOverrides", json_value(useTerrainOverrides)
    save.setValue "inactivityMinutes", json_value_unsigned(game.inactivityMinutes)
    save.setValue "gameTimeLimitMinutes", json_value_unsigned(game.gameTimeLimitMinutes)
    save.setValue "playerLeaveMode", json_value(game.playerLeaveMode)
    save.setValue "blindMode", json_value(game.blindMode)
    save.setValue "tweakOptions", getCamTweakOptions()

    save.beginArray "scriptSetPlayerDataStrings"
    For i = 0 To NetPlay_scriptSetPlayerDataStrings_size() - 1
        ' (for (const auto& kv : NetPlay.scriptSetPlayerDataStrings[i]): the strings are not kept before cluster J)
        save.nextArrayItem
    Next
    save.endArray

    writeMainFile = True
End Function

'-- Write gam.json
'--
Private Function writeGameFile(ByVal fileName As String, ByVal saveType As Long) As Boolean
    Dim saveGame_ As SAVE_GAME, i As Long, j As Long, psLandingZone As LANDING_ZONE, len_ As Long, lastSep As Long
    Dim pathToCommonSaveDir As String, gameName As String, pathToThisSaveDir As String, saveInfoJsonFilename As String
    Dim jsonFileName As String, gamJson As nlohmann_json, saveInfoJsonOpt As nlohmann_json, hasSaveInfoJson As Boolean
    Dim saveInfoJson As nlohmann_json

    ' (GAME_SAVEHEADER fileHeader: 'g' 'a' 'm' 'e', CURRENT_VERSION_NUM: only its version is used)
    debug_ 4994, LOG_SAVE, "writeGameFile", "fileversion is %u, (%s) ", CURRENT_VERSION_NUM, fileName

    If Not (saveType = GTYPE_SAVE_START Or saveType = GTYPE_SAVE_MIDMISSION) Then ASSERT_FAILURE "saveType == GTYPE_SAVE_START || saveType == GTYPE_SAVE_MIDMISSION", "src/game.cpp:4996", "writeGameFile", "invalid save type"
    saveGame_.saveKey = getCampaignNumber()
    If missionIsOffworld() Then
        saveGame_.saveKey = saveGame_.saveKey Or SAVEKEY_ONMISSION
        setSaveGameOnMission True
    Else
        setSaveGameOnMission False
    End If


    ' Put the save game data into the buffer
    saveGame_.gameTime = gameTime
    saveGame_.missionTime = mission_startTime

    'put in the scroll data
    saveGame_.ScrollMinX = scrollMinX
    saveGame_.ScrollMinY = scrollMinY
    saveGame_.ScrollMaxX = scrollMaxX
    saveGame_.ScrollMaxY = scrollMaxY

    saveGame_.GameType = saveType

    'save the current level so we can load up the STARTING point of the mission
    If Not (Len(aLevelName) < MAX_LEVEL_SIZE) Then ASSERT_FAILURE "strlen(aLevelName) < MAX_LEVEL_SIZE", "src/game.cpp:5022", "writeGameFile", "Unable to save level name - too long (max %d) - %s", _
                     MAX_LEVEL_SIZE, aLevelName: writeGameFile = False: Exit Function
    saveGame_.levelName = sstrcpyN(aLevelName, MAX_LEVEL_SIZE)

    'save out the players power
    For i = 0 To MAX_PLAYERS - 1
        saveGame_.power(i).currentPower = getPower(i)
    Next
    saveGame_.power(0).extractedPower = IIf(radarPermitted, 1&, 0&)  ' hideous hack, don't want to break savegames now...
    saveGame_.power(1).extractedPower = IIf(allowDesign, 1&, 0&)     ' hideous hack, don't want to break savegames now...

    'camera position
    disp3d_getView saveGame_.currentPlayerPos

    'mission data
    saveGame_.missionOffTime = mission_time
    saveGame_.missionETA = mission_ETA
    saveGame_.missionCheatTime = CLng(mission_cheatTime)
    saveGame_.missionHomeLZ_X = mission_homeLZ_X
    saveGame_.missionHomeLZ_Y = mission_homeLZ_Y
    saveGame_.missionPlayerX = mission_playerX
    saveGame_.missionPlayerY = mission_playerY
    saveGame_.missionScrollMinX = mission_scrollMinX And &HFFFF&
    saveGame_.missionScrollMinY = mission_scrollMinY And &HFFFF&
    saveGame_.missionScrollMaxX = mission_scrollMaxX And &HFFFF&
    saveGame_.missionScrollMaxY = mission_scrollMaxY And &HFFFF&

    saveGame_.offWorldKeepLists = IIf(offWorldKeepLists, 1&, 0&)
    saveGame_.RubbleTile = getRubbleTileNum()
    saveGame_.WaterTile = getWaterTileNum()

    For i = 0 To MAX_PLAYERS - 1
        saveGame_.iTranspEntryTileX(i) = mission_iTranspEntryTileX(i)
        saveGame_.iTranspEntryTileY(i) = mission_iTranspEntryTileY(i)
        saveGame_.iTranspExitTileX(i) = mission_iTranspExitTileX(i)
        saveGame_.iTranspExitTileY(i) = mission_iTranspExitTileY(i)
        saveGame_.aDefaultSensor(i) = aDefaultSensor(i)
        saveGame_.aDefaultECM(i) = aDefaultECM(i)
        saveGame_.aDefaultRepair(i) = aDefaultRepair(i)
    Next

    For i = 0 To MAX_NOGO_AREAS - 1
        psLandingZone = getLandingZone(i)
        saveGame_.sLandingZone(i).x1 = psLandingZone.x1 ' in case struct changes
        saveGame_.sLandingZone(i).x2 = psLandingZone.x2
        saveGame_.sLandingZone(i).y1 = psLandingZone.y1
        saveGame_.sLandingZone(i).y2 = psLandingZone.y2
    Next

    'version 17
    saveGame_.objId = FromU32(saveObjId())

    'version 18
    saveGame_.buildDate = ""            ' memset(saveGame.buildDate, 0, sizeof(saveGame.buildDate));
    saveGame_.oldestVersion = 0
    saveGame_.validityKey = 0

    'version 19
    For i = 0 To MAX_PLAYERS - 1
        For j = 0 To MAX_PLAYERS - 1
            saveGame_.alliances(i, j) = alliances(i, j) And &HFF&
        Next
    Next
    For i = 0 To MAX_PLAYERS - 1
        saveGame_.playerColour(i) = getPlayerColour(i) And &HFF&
    Next
    saveGame_.radarZoom = GetRadarZoom() And &HFF&

    'version 20
    saveGame_.bDroidsToSafetyFlag = IIf(getDroidsToSafetyFlag(), 1&, 0&)

    'version 24
    saveGame_.reinforceTime = missionGetReinforcementTime()
    saveGame_.bPlayCountDown = IIf(getPlayCountDown(), 1&, 0&)
    saveGame_.bPlayerHasWon = IIf(testPlayerHasWon(), 1&, 0&)
    saveGame_.bPlayerHasLost = IIf(testPlayerHasLost(), 1&, 0&)

    'version 30
    saveGame_.scrGameLevel = 0
    saveGame_.bExtraFailFlag = 0
    saveGame_.bExtraVictoryFlag = 0
    saveGame_.bTrackTransporter = 0

    ' version 33
    saveGame_.sGame = game
    saveGame_.savePlayer = selectedPlayer
    saveGame_.multiPlayer = IIf(bMultiPlayer, 1&, 0&)
    ' saveGame.sNetPlay = NetPlay (the fields serializeNetPlay_json keeps)
    For i = 0 To MAX_PLAYERS - 1
        saveGame_.sNetPlay.players(i).position = NetPlay_players_position(i)
        saveGame_.sNetPlay.players(i).name_ = NetPlay_players_name(i)
        saveGame_.sNetPlay.players(i).ai = NetPlay_players_ai(i)
        saveGame_.sNetPlay.players(i).difficulty = NetPlay_players_difficulty(i)
        saveGame_.sNetPlay.players(i).allocated = NetPlay_players_allocated(i)
        saveGame_.sNetPlay.players(i).colour = NetPlay_players_colour(i)
        saveGame_.sNetPlay.players(i).team = NetPlay_players_team(i)
    Next
    saveGame_.sNetPlay.bComms = NetPlay_bComms
    saveGame_.sNetPlay.playercount = NetPlay_playercount()
    saveGame_.sNetPlay.hostPlayer = NetPlay_hostPlayer
    saveGame_.sNetPlay.isHost = NetPlay_isHost
    saveGame_.sPName = sstrcpyN(getPlayerName(selectedPlayer), 32)
    For i = 0 To MAX_PLAYERS - 1
        saveGame_.sPlayerIndex(i) = i
    Next

    'version 34
    For i = 0 To MAX_PLAYERS - 1
        ' (VB) !challengeActive && NetPlay.players[i].ai >= 0 && !NetPlay.players[i].allocated, in order
        If Not challengeActive And NetPlay_players_ai(i) >= 0 And Not NetPlay_players_allocated(i) Then
            saveGame_.sPlayerName(i) = sstrcpyN(getAIName(i), StringSize)
        Else
            saveGame_.sPlayerName(i) = sstrcpyN(getPlayerName(i), StringSize)
        End If
    Next

    'version 38
    saveGame_.modList = sstrcpyN(getModList(), modlist_string_size)
    ' Attempt to see if we have a corrupted game structure in campaigns.
    If saveGame_.sGame.typ = LEVEL_TYPE_CAMPAIGN Then
        ' player 0 is always a human in campaign games
        For i = 1 To MAX_PLAYERS - 1
            If saveGame_.sNetPlay.players(i).difficulty = AIDifficulty_HUMAN Then
                ASSERT_FAILURE "!""savegame corruption!""", "src/game.cpp:5138", "writeGameFile", "savegame corruption!"
                debug_ 5139, LOG_ERROR, "writeGameFile", "Savegame corruption detected, trying to salvage.  Please Report this issue @ wz2100.net"
                debug_ 5140, LOG_ERROR, "writeGameFile", "players[i].difficulty was %d, level %s / %s, ", saveGame_.sNetPlay.players(i).difficulty, saveGame_.levelName, saveGame_.sGame.map
                saveGame_.sNetPlay.players(i).difficulty = AIDifficulty_DISABLED
            End If
        Next
    End If
    len_ = Len(fileName)
    If Not (Right$(fileName, 4) = ".gam") Then ASSERT_FAILURE "strcmp(fileName + (len - 4), "".gam"") == 0", "src/game.cpp:5147", "writeGameFile", "hmm... not .gam?"

    lastSep = InStrRev(fileName, "/") - 1          ' fileNameStr.rfind("/"): -1 (npos + 1 = 0) when there is none
    ' TODO: convert argument "const char* filename" to a manageable struct with path/filename/extension
    '        and remove this mess...
    pathToCommonSaveDir = Left$(fileName, lastSep + 1)
    gameName = Mid$(fileName, lastSep + 2, len_ - Len(pathToCommonSaveDir) - 4)
    pathToThisSaveDir = pathToCommonSaveDir & gameName & "/"
    If PHYSFS_exists(pathToThisSaveDir) = 0 Then
        PHYSFS_mkdir pathToThisSaveDir
    End If
    saveInfoJsonFilename = pathToThisSaveDir & "save-info.json"
    jsonFileName = pathToThisSaveDir & "gam.json"
    Set gamJson = json_object()

    hasSaveInfoJson = False             ' nonstd::optional<nlohmann::json> saveInfoJsonOpt = nullopt;
    If PHYSFS_exists(saveInfoJsonFilename) <> 0 Then
        hasSaveInfoJson = parseJsonFile(saveInfoJsonFilename, saveInfoJsonOpt)
        If Not (hasSaveInfoJson And saveInfoJsonIsObject(hasSaveInfoJson, saveInfoJsonOpt)) Then ASSERT_FAILURE "saveInfoJsonOpt.has_value() && saveInfoJsonOpt.value().is_object()", "src/game.cpp:5167", "writeGameFile", "save-info.json looks broken, wanted an object"
    End If
    If Not saveInfoJsonIsObject(hasSaveInfoJson, saveInfoJsonOpt) Then
        Set saveInfoJsonOpt = json_object()
        hasSaveInfoJson = True
    End If
    Set saveInfoJson = saveInfoJsonOpt
    ' new .json format
    serializeSaveGameData_json gamJson, saveInfoJson, gameName, saveGame_
    If Not saveJSONToFile(gamJson, jsonFileName) Then
        debug_ 5178, LOG_ERROR, "writeGameFile", "Failed to save: %s", jsonFileName
        writeGameFile = False
        Exit Function
    End If
    If Not saveJSONToFile(saveInfoJson, saveInfoJsonFilename) Then
        debug_ 5183, LOG_ERROR, "writeGameFile", "Failed to save: %s", saveInfoJsonFilename
        writeGameFile = False
        Exit Function
    End If
    debug_ 5186, LOG_SAVEGAME, "writeGameFile", "saved game into %s", jsonFileName

    ' Return our success status with writing out the file!
    writeGameFile = True
End Function

' (VB) saveInfoJsonOpt.has_value() && saveInfoJsonOpt.value().is_object() (no short-circuit in VB)
Private Function saveInfoJsonIsObject(ByVal hasValue As Boolean, ByVal v As nlohmann_json) As Boolean
    If Not hasValue Then Exit Function
    If v Is Nothing Then Exit Function
    saveInfoJsonIsObject = v.is_object()
End Function

' (VB) ctime(&currentTime): "Www Mmm dd hh:mm:ss yyyy\n" in local time
Private Function ctime_() As String
    Dim t As tm, days As Variant, months As Variant
    days = Array("Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat")
    months = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    localtime_s t, time_()
    ' "%.3s %.3s%3d %.2d:%.2d:%.2d %d\n"
    ctime_ = days(t.tm_wday) & " " & months(t.tm_mon) & Right$("   " & CStr(t.tm_mday), 3) & " " & _
             Format$(t.tm_hour, "00") & ":" & Format$(t.tm_min, "00") & ":" & Format$(t.tm_sec, "00") & " " & _
             CStr(t.tm_year + 1900) & vbLf
End Function

' (game.cpp:6678)
Public Function writeGameInfo(ByVal pFileName As String) As Boolean
    Dim ini As WzConfig, ourtime As String, time_s As String, backendInfo As Object, keys As Variant, n As Long, k As Long
    Dim j As Long, t As Variant
    Dim modList As String
    ' const DebugInputManager& dbgInputManager = gInputManager.debugManager(); (debugMappingsAllowed below)

    Set ini = WzConfig_new(pFileName, WzConfig_ReadAndWrite)
    time_s = ctime_()                   ' std::string time(ctime(&currentTime));

    ini.beginGroup "GameProperties"
    ini.setValue "current_time", json_value(time_s)
    ourtime = getAsciiTime(graphicsTime)
    ini.setValue "graphics_time", json_value(ourtime)
    ourtime = getAsciiTime(gameTime)
    ini.setValue "game_time", json_value(ourtime)
    ourtime = getAsciiTime(gameTime - missionData.missionStarted)
    ini.setValue "playing_time", json_value(ourtime)
    ini.setValue "version", json_value(version_getVersionString())
    ini.setValue "full_version", json_value(version_getFormattedVersionString())
    ini.setValue "cheated", json_value(Cheated)
    ini.setValue "debug", json_value(debugMappingsAllowed())
    ini.setValue "level/map", json_value(getLevelName())
    modList = getModList()
    ini.setValue "mods", json_value(IIf(Len(modList) <> 0, modList, "None"))
    ' for (auto& kv : gfx_api::context::get().getBackendGameInfo()): (VB) a Scripting.Dictionary; a std::map's
    ' keys come sorted
    Set backendInfo = gfx_api_context_getBackendGameInfo()
    If Not backendInfo Is Nothing Then
        n = backendInfo.Count
        If n > 0 Then
            keys = backendInfo.Keys
            For k = 1 To n - 1
                t = keys(k)
                j = k - 1
                Do While j >= 0
                    If StrComp(keys(j), t, vbBinaryCompare) <= 0 Then Exit Do
                    keys(j + 1) = keys(j)
                    j = j - 1
                Loop
                keys(j + 1) = t
            Next
            For k = 0 To n - 1
                ini.setValue CStr(keys(k)), json_value(CStr(backendInfo.Item(keys(k))))
            Next
        End If
    End If
    ini.endGroup
    writeGameInfo = True
End Function

' -----------------------------------------------------------------------------------------
Public Function savegameWithoutExtension(ByVal name_ As String) As String
    Dim displaySavegameName As String, n As Long   ' static char displaySavegameName[256]
    displaySavegameName = Left$(name_, 255)        ' sstrcpy
    n = Len(displaySavegameName)
    If n < 4 Then n = 4
    displaySavegameName = Left$(displaySavegameName, n - 4) 'axe the .gam

    savegameWithoutExtension = displaySavegameName
End Function
