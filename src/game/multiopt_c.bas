Attribute VB_Name = "multiopt_c"
' src/multiopt.cpp, multiopt.h: the game options' exchange (sendOptions,
' recvOptions with the map and mod requests, sendHostConfig, recvHostConfig
' and the NETlockedOptions writer and reader), hostCampaign, sendLeavingMsg,
' multiShutdown, the start screen (multiStartScreenInit), the game's start-up
' (gameInit, playerResponding, multiGameInit), multiGameShutdown and
' printStructureLimitsInfo. ingame.startTime and the other
' std::chrono::steady_clock::time_points are QueryPerformanceCounter counts.
' The ActivityManager's updateMultiplayGameData (the activity sink) is not
' kept. The downloads' list (NET_getDownloadingWzFiles) is WzNet's.
Option Explicit

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' lib/netplay/nettypes.h
Private Const MAX_CONNECTED_PLAYERS As Long = MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS

Private Const MAX_STRUCTURE_LIMITS As Long = 4096 ' Set a high (but explicit) maximum for the number of structure limits supported

' (VB) Sha256::Bytes
Private Const Sha256_Bytes As Long = 32

' recvOptions' [](const std::string& line) {...} for printStructureLimitsInfo, multiStartScreenInit's on-close lambda
Public Const FN_MOPT_LIMITSINFO_PRINTLINE As Long = 17003
Public Const FN_MOPT_STARTSCREEN_ONCLOSE As Long = 17004

' recvOptions' enum class FileRequestResult
Private Const FileRequestResult_StartingDownload As Long = 0
Private Const FileRequestResult_DownloadInProgress As Long = 1
Private Const FileRequestResult_FileExists As Long = 2
Private Const FileRequestResult_FailedToOpenFileForWriting As Long = 3

' std::vector<MULTISTRUCTLIMITS> oldStructureLimits: oldStructureLimits(0 .. oldStructureLimitsCount - 1)
Private oldStructureLimits() As MULTISTRUCTLIMITS
Private oldStructureLimitsCount As Long

' (VB) NETbin(w, hash.bytes, hash.Bytes)
Private Sub NETbin_Sha256(ByVal w As MessageWriter, h As Sha256)
    Dim b() As Byte, n As Long
    ReDim b(0 To Sha256_Bytes - 1)
    For n = 0 To Sha256_Bytes - 1
        b(n) = h.bytes(n)
    Next
    NETbin w, b, Sha256_Bytes
End Sub

' send complete game info set!
Public Sub sendOptions()
    Dim w As MessageWriter, v As Long, modHashesSize As Long, i As Long, j As Long, numStructureLimits As Long, k As Long
    Dim modHashes() As Sha256
    If Not NetPlay_isHost Then ASSERT_FAILURE "false", "src/multiopt.cpp:76", "sendOptions", "Host only routine detected for client!": Exit Sub
    If Not (GetGameMode() <> GS_NORMAL) Then ASSERT_FAILURE "GetGameMode() != GS_NORMAL", "src/multiopt.cpp:77", "sendOptions", "sendOptions shouldn't be called after the game has started": Exit Sub

    game.modHashCount = getModHashList(modHashes)
    If game.modHashCount > 0 Then
        ReDim game.modHashes(0 To game.modHashCount - 1)
        For k = 0 To game.modHashCount - 1
            game.modHashes(k) = modHashes(k)
        Next
    Else
        Erase game.modHashes
    End If

    Set w = NETbeginEncode(NETbroadcastQueue(), NET_OPTIONS)
    ' First send information about the game
    v = game.typ And &HFF&
    NETuint8_t w, v
    NETstring w, game.map, 128
    NETbin_Sha256 w, game.hash
    modHashesSize = game.modHashCount
    NETuint32_t w, modHashesSize
    For k = 0 To game.modHashCount - 1
        NETbin_Sha256 w, game.modHashes(k)
    Next
    NETuint8_t w, game.maxPlayers
    NETstring w, game.name_, 128
    NETuint32_t w, game.power
    NETuint8_t w, game.base
    NETuint8_t w, game.alliance
    NETuint8_t w, game.scavengers
    NETbool w, game.isMapMod
    NETuint32_t w, game.techLevel
    If game.inactivityMinutes > 0 And ULT(game.inactivityMinutes, MIN_MPINACTIVITY_MINUTES) Then
        debug_ 102, LOG_ERROR, "sendOptions", "Invalid inactivityMinutes value specified: %u; resetting to: %u", game.inactivityMinutes, MIN_MPINACTIVITY_MINUTES
        game.inactivityMinutes = MIN_MPINACTIVITY_MINUTES
    End If
    NETuint32_t w, game.inactivityMinutes
    If game.gameTimeLimitMinutes > 0 And ULT(game.gameTimeLimitMinutes, MIN_MPGAMETIMELIMIT_MINUTES) Then
        debug_ 108, LOG_ERROR, "sendOptions", "Invalid gameTimeLimitMinutes value specified: %u; resetting to: %u", game.gameTimeLimitMinutes, MIN_MPGAMETIMELIMIT_MINUTES
        game.gameTimeLimitMinutes = MIN_MPGAMETIMELIMIT_MINUTES
    End If
    NETuint32_t w, game.gameTimeLimitMinutes
    v = game.playerLeaveMode And &HFF&
    NETuint8_t w, v

    For i = 0 To MAX_PLAYERS - 1
        v = NetPlay_players_difficulty(i)
        NETint8_t w, v
    Next

    ' Send the list of who is still joining
    For i = 0 To MAX_CONNECTED_PLAYERS - 1
        NETbool w, ingame.JoiningInProgress(i)
    Next

    ' Same goes for the alliances
    For i = 0 To MAX_PLAYERS - 1
        For j = 0 To MAX_PLAYERS - 1
            NETuint8_t w, alliances(i, j)
        Next
    Next

    ' Send the number of structure limits to expect
    numStructureLimits = ingame.structureLimitsCount
    If numStructureLimits > MAX_STRUCTURE_LIMITS Then
        debug_ 138, LOG_ERROR, "sendOptions", "Number of structure limits (%u) exceeds maximum supported - truncating", numStructureLimits
        numStructureLimits = MAX_STRUCTURE_LIMITS
    End If
    NETuint32_t w, numStructureLimits
    debug_ 142, LOG_NET, "sendOptions", "(Host) Structure limits to process on client is %zu", ingame.structureLimitsCount
    ' Send the structures changed
    For k = 0 To ingame.structureLimitsCount - 1
        NETuint32_t w, ingame.structureLimits(k).id
        NETuint32_t w, ingame.structureLimits(k).limit
    Next
    updateStructureDisabledFlags
    NETuint8_t w, ingame.flags

    NETend w

    ' also send a NET_HOST_CONFIG msg here
    sendHostConfig

    ' (ActivityManager::instance().updateMultiplayGameData(game, ingame, NETGameIsLocked()): the activity sink, not kept)
End Sub

' ////////////////////////////////////////////////////////////////////////////
' Host Campaign.
Public Function hostCampaign(ByVal SessionName As String, ByVal hostPlayerName As String, ByVal spectatorHost As Boolean, ByVal skipResetAIs As Boolean, ByVal desiredOpenSpectatorSlots As Long) As Boolean
    Dim playerStats As PLAYERSTATS_, i As Long, identityBytes() As Byte, n As Long
    debug_ 459, LOG_WZ, "hostCampaign", "Hosting campaign: '%s', player: '%s'", SessionName, hostPlayerName

    freeMessages

    loadMultiStats hostPlayerName, playerStats

    ' (VB) NEThostGame's playerIdentity (playerStats.identity) is left out: NetApi's signature, as the exe's stand-in
    ' had it (WzNet's GAMESTRUCT's host.publicIdentity, for the lobby server)
    If Not NEThostGame(SessionName, hostPlayerName, spectatorHost, game.typ, game.blindMode, game.maxPlayers, desiredOpenSpectatorSlots, game.alliance, game.techLevel, game.power, game.base) Then
        hostCampaign = False
        Exit Function
    End If

    ' Skip resetting AIs if we are doing autohost
    If NetPlay_bComms And Not skipResetAIs Then
        For i = 0 To MAX_CONNECTED_PLAYERS - 1
            If Not (i = selectedPlayer And i < MAX_PLAYERS) Then
                NetPlay_players_difficulty(i) = AIDifficulty_DISABLED
            End If
        Next
    End If

    NetPlay_players_ready(selectedPlayer) = False
    setPlayerName selectedPlayer, hostPlayerName

    ingame.localJoiningInProgress = True
    ingame.JoiningInProgress(selectedPlayer) = True
    ingame.PendingDisconnect(selectedPlayer) = False
    ingame.has_joinTimes(selectedPlayer) = True
    QueryPerformanceCounter ingame.joinTimes(selectedPlayer)
    ingame.has_lastReadyTimes(selectedPlayer) = False
    ingame.has_lastNotReadyTimes(selectedPlayer) = False
    ingame.secondsNotReady(selectedPlayer) = 0
    ingame.has_playerLeftGameTime(selectedPlayer) = False
    bMultiPlayer = True
    bMultiMessages = True ' enable messages

    setMultiStats selectedPlayer, playerStats, False
    setMultiStats selectedPlayer, playerStats, True

    n = EcKey_toBytes(playerStats.identity, EcKey_Public, identityBytes)
    multiStatsSetVerifiedHostIdentityFromJoin identityBytes, n

    ' (ActivityManager::instance().updateMultiplayGameData(game, ingame, NETGameIsLocked()): the activity sink, not kept)
    hostCampaign = True
End Function

' ////////////////////////////////////////////////////////////////////////////
Private Function gameInit() As Boolean
    Dim player As Long, i As Long, playerCount As Long, index As Long

    ' Various sanity checks (for *true* multiplayer)
    For player = 0 To MAX_CONNECTED_PLAYERS - 1
        If player < MAX_PLAYERS Then
            If NetPlay_players_allocated(player) Then
                If Not (NetPlay_players_difficulty(player) = AIDifficulty_HUMAN) Then ASSERT_FAILURE "NetPlay.players[player].difficulty == AIDifficulty::HUMAN", "src/multiopt.cpp:553", "gameInit", "Found an allocated (human) player (%u) with mis-matched difficulty (%d)", player, NetPlay_players_difficulty(player)

            End If
        End If

        If bMultiPlayer And NetPlay_bComms Then
            If Not NetPlay_players_allocated(player) Then
                If Not (NetPlay_players_difficulty(player) <> AIDifficulty_HUMAN) Then ASSERT_FAILURE "NetPlay.players[player].difficulty != AIDifficulty::HUMAN", "src/multiopt.cpp:562", "gameInit", "Found a non-human slot (%u) with mis-matched (human) difficulty (%d)", player, NetPlay_players_difficulty(player)
                If Not (NetPlay_players_ai(player) < 0 Or NetPlay_players_difficulty(player) <> AIDifficulty_DISABLED) Then ASSERT_FAILURE "NetPlay.players[player].ai < 0 || NetPlay.players[player].difficulty != AIDifficulty::DISABLED", "src/multiopt.cpp:563", "gameInit", "Slot (%u) has AI, but disabled difficulty?", player
            End If
        End If
    Next

    For player = 1 To MAX_CONNECTED_PLAYERS - 1
        ' we want to remove disabled AI & all the other players that don't belong
        If (NetPlay_players_difficulty(player) = AIDifficulty_DISABLED Or player >= game.maxPlayers) _
            And player <> scavengerPlayer() _
            And Not (NetPlay_players_isSpectator(player) And NetPlay_players_allocated(player)) _
            And Not (player < game.maxPlayers And NetPlay_players_allocated(player)) Then
            clearPlayer player, True            ' do this quietly
            debug_ 577, LOG_NET, "gameInit", "removing disabled AI (%d) from map.", player
        End If
    Next

    For i = 0 To NetPlay_players_size() - 1
        If NetPlay_players_isSpectator(i) Then
            ' player is starting as a spectator
            makePlayerSpectator i, True, True
            If i = selectedPlayer Then
                setPlayerHasLost True ' set this flag to true so we don't accumulate loss statistics
            End If
        End If
    Next

    playerCount = 0
    For index = 0 To game.maxPlayers - 1
        If NetPlay_players_ai(index) >= 0 Or NetPlay_players_allocated(index) Then playerCount = playerCount + 1
    Next
    debug_ 599, LOG_NET, "gameInit", "Player count: %u", playerCount

    gameInit = True
End Function

' ////////////////////////////////////////////////////////////////////////////
' say hi to everyone else....
Public Sub playerResponding()
    Dim w As MessageWriter, player As Long
    QueryPerformanceCounter ingame.startTime
    ingame.localJoiningInProgress = False ' No longer joining.
    If ULT(selectedPlayer, MAX_CONNECTED_PLAYERS) Then
        ingame.JoiningInProgress(selectedPlayer) = False
    End If

    ' Tell the world we're here
    Set w = NETbeginEncode(NETbroadcastQueue(), NET_PLAYERRESPONDING)
    player = selectedPlayer
    NETuint32_t w, player
    NETend w
End Sub

'//////////////////////////////
' at the end of every game.
Public Function multiGameShutdown() As Boolean
    Dim st As PLAYERSTATS_, time_ As Long
    debug_ 658, LOG_NET, "multiGameShutdown", "%s is shutting down.", getPlayerName(selectedPlayer)

    shutdownGameStartScreen     ' make sure the start screen overlay is closed (in case the game shuts down before it fully starts)

    sendLeavingMsg              ' say goodbye

    If ULT(selectedPlayer, MAX_CONNECTED_PLAYERS) Then
        st = getMultiStats(selectedPlayer)  ' save stats

        saveMultiStats getPlayerName(selectedPlayer), getPlayerName(selectedPlayer), st
    End If

    ' if we terminate the socket too quickly, then, it is possible not to get the leave message
    time_ = wzGetTicks()
    Do While U32(wzGetTicks() - time_) < 1000
        Sleep 0                 ' wzYieldCurrentThread(): std::this_thread::yield()  ' TODO Make a wzDelay() function?
    Loop
    ' close game
    NETclose
    NETremRedirects

    Erase ingame.structureLimits    ' ingame.structureLimits.clear()
    ingame.structureLimitsCount = 0
    ingame.flags = 0

    ingame.localJoiningInProgress = False ' Clean up
    ingame.localOptionsReceived = False
    ingame.side = InGameSide_HOST_OR_SINGLEPLAYER
    ingame.has_TimeEveryoneIsInGame = False     ' nullopt
    ingame.startTime = 0                        ' std::chrono::steady_clock::time_point()
    ingame.has_endTime = False                  ' nullopt
    ingame.lastLagCheck = 0
    ingame.lastDesyncCheck = 0
    ingame.lastPlayerDataCheck2 = 0
    NetPlay_isHost = False
    bMultiPlayer = False        ' Back to single player mode
    bMultiMessages = False
    selectedPlayer = 0          ' Back to use player 0 (single player friendly)

    NET_InitPlayers

    resetKickVoteData
    resetAllMultiOptionPrefValues

    multiGameShutdown = True
End Function

Public Function multiGameInit() As Boolean
    Dim player As Long

    For player = 0 To MAX_CONNECTED_PLAYERS - 1
        openchannels(player) = True                             'open comms to this player.
    Next

    gameInit

    multiGameInit = True
End Function

Private Sub NETlockedOptions(ByVal w As MessageWriter, lockedOpts As MultiplayOptionsLocked)
    NETbool w, lockedOpts.scavengers
    NETbool w, lockedOpts.alliances
    NETbool w, lockedOpts.teams
    NETbool w, lockedOpts.power
    NETbool w, lockedOpts.difficulty
    NETbool w, lockedOpts.ai
    NETbool w, lockedOpts.position
    NETbool w, lockedOpts.bases
    NETbool w, lockedOpts.spectators
    NETbool w, lockedOpts.name_
    NETbool w, lockedOpts.readybeforefull
    NETbool w, lockedOpts.cheats
End Sub

Public Sub sendHostConfig()
    Dim w As MessageWriter, i As Long, lockedOpts As MultiplayOptionsLocked
    If Not NetPlay_isHost Then ASSERT_FAILURE "false", "src/multiopt.cpp:857", "sendHostConfig", "Host only routine detected for client!": Exit Sub

    Set w = NETbeginEncode(NETbroadcastQueue(), NET_HOST_CONFIG)

    ' Send the list of host-set player chat permissions
    For i = 0 To MAX_CONNECTED_PLAYERS - 1
        NETbool w, ingame.hostChatPermissions(i)
    Next

    ' Send the "locked" options configuration
    lockedOpts = getLockedOptions()
    NETlockedOptions w, lockedOpts

    NETend w
End Sub

' (VB) NETbin(r, hash.bytes, hash.Bytes)
Private Sub NETbin_Sha256_r(ByVal r As MessageReader, h As Sha256)
    Dim b() As Byte, n As Long
    ReDim b(0 To Sha256_Bytes - 1)
    NETbin r, b, Sha256_Bytes
    For n = 0 To Sha256_Bytes - 1
        h.bytes(n) = b(n)
    Next
End Sub

' (VB) oldStructureLimits != ingame.structureLimits (the sizes are equal)
Private Function structureLimitsDiffer() As Boolean
    Dim i As Long
    For i = 0 To oldStructureLimitsCount - 1
        If oldStructureLimits(i).id <> ingame.structureLimits(i).id Or oldStructureLimits(i).limit <> ingame.structureLimits(i).limit Then
            structureLimitsDiffer = True
            Exit Function
        End If
    Next
    structureLimitsDiffer = False
End Function

' (VB) game.modHashes != priorGameInfo.modHashes
Private Function modHashesDiffer(a As MULTIPLAYERGAME, b As MULTIPLAYERGAME) As Boolean
    Dim i As Long
    If a.modHashCount <> b.modHashCount Then
        modHashesDiffer = True
        Exit Function
    End If
    For i = 0 To a.modHashCount - 1
        If Not Sha256_equals(a.modHashes(i), b.modHashes(i)) Then
            modHashesDiffer = True
            Exit Function
        End If
    Next
    modHashesDiffer = False
End Function

' recvOptions' auto requestFile = [](Sha256 &hash, char const *filename) -> FileRequestResult
Private Function requestFile(hash As Sha256, ByVal filename As String) As Long
    Dim fileHash As Sha256, pFileHandle As Long, w As MessageWriter
    If NET_getDownloadingWzFiles_hasHash(hash.bytes) Then
        debug_ 333, LOG_INFO, "recvOptions", "Already requested file, continue waiting."
        requestFile = FileRequestResult_DownloadInProgress  ' Downloading the file already
        Exit Function
    End If

    If PHYSFS_exists(filename) = 0 Then
        debug_ 339, LOG_INFO, "recvOptions", "Creating new file %s", filename
    Else
        findHashOfFile filename, fileHash
        If Not Sha256_equals(fileHash, hash) Then
            debug_ 343, LOG_INFO, "recvOptions", "Overwriting old incomplete or corrupt file %s", filename
        Else
            pal_Init ' Palette could be modded. // Why is this here - isn't there a better place for it?
            requestFile = FileRequestResult_FileExists  ' Have the file already.
            Exit Function
        End If
    End If

    pFileHandle = PHYSFS_openWrite(filename)
    If pFileHandle = 0 Then
        debug_ 354, LOG_ERROR, "recvOptions", "Failed to open %s for writing: %s", filename, WZ_PHYSFS_getLastError()
        requestFile = FileRequestResult_FailedToOpenFileForWriting
        Exit Function
    End If

    NET_addDownloadingWZFile pFileHandle, filename, hash.bytes   ' (VB) a Sha256 crosses into WzNet as its bytes

    ' Request the map/mod from the host
    Set w = NETbeginEncode(NETnetQueue(NetPlay_hostPlayer), NET_FILE_REQUESTED)
    NETbin_Sha256 w, hash
    NETend w

    requestFile = FileRequestResult_StartingDownload  ' Starting download now.
End Function

' ////////////////////////////////////////////////////////////////////////////
' options for a game. (usually recvd in frontend)
' returns: false if the options should be considered invalid and the client should disconnect
Public Function recvOptions(ByVal queue As Long) As Boolean
    Dim i As Long, j As Long, priorGameInfo As MULTIPLAYERGAME, r As MessageReader, v As Long, modHashesSize As Long
    Dim tempPlayerLeaveModeValue As Long, numStructureLimits As Long, structurelimitsUpdated As Boolean, tmpConsoleMsgStr As String
    Dim bRebuildMapList As Boolean, mapData As LEVEL_DATASET, mapName As String, filename As String, requestResult As Long
    Dim str_ As String, n As Long
    If NETQUEUE_index(queue) <> NetPlay_hostPlayer Then ASSERT_FAILURE "queue.index == NetPlay.hostPlayer", "src/multiopt.cpp:167", "recvOptions", "NET_OPTIONS received from unexpected player: %u - ignoring", NETQUEUE_index(queue): recvOptions = True: Exit Function
    If Not (GetGameMode() <> GS_NORMAL) Then ASSERT_FAILURE "GetGameMode() != GS_NORMAL", "src/multiopt.cpp:168", "recvOptions", "NET_OPTIONS received after the game has started??": recvOptions = False: Exit Function

    ' store prior map / mod info
    priorGameInfo = game

    debug_ 175, LOG_NET, "recvOptions", "Receiving options from host"
    Set r = NETbeginDecode(queue, NET_OPTIONS)

    ' Get general information about the game
    v = game.typ
    NETuint8_t r, v
    game.typ = v
    NETstring r, game.map, 128
    NETbin_Sha256_r r, game.hash
    modHashesSize = 0
    NETuint32_t r, modHashesSize
    If Not ULT(modHashesSize, 1000000) Then ASSERT_FAILURE "modHashesSize < 1000000", "src/multiopt.cpp:184", "recvOptions", "Way too many mods %u", modHashesSize: recvOptions = False: Exit Function
    game.modHashCount = modHashesSize
    If modHashesSize > 0 Then
        ReDim game.modHashes(0 To modHashesSize - 1)
    Else
        Erase game.modHashes
    End If
    For n = 0 To game.modHashCount - 1
        NETbin_Sha256_r r, game.modHashes(n)
    Next
    NETuint8_t r, game.maxPlayers
    NETstring r, game.name_, 128
    NETuint32_t r, game.power
    NETuint8_t r, game.base
    NETuint8_t r, game.alliance
    NETuint8_t r, game.scavengers
    NETbool r, game.isMapMod
    NETuint32_t r, game.techLevel
    NETuint32_t r, game.inactivityMinutes
    If game.inactivityMinutes > 0 And ULT(game.inactivityMinutes, MIN_MPINACTIVITY_MINUTES) Then
        debug_ 201, LOG_ERROR, "recvOptions", "Invalid inactivityMinutes value specified: %u", game.inactivityMinutes
        recvOptions = False
        Exit Function
    End If
    NETuint32_t r, game.gameTimeLimitMinutes
    If game.gameTimeLimitMinutes > 0 And ULT(game.gameTimeLimitMinutes, MIN_MPGAMETIMELIMIT_MINUTES) Then
        debug_ 207, LOG_ERROR, "recvOptions", "Invalid gameTimeLimitMinutes value specified: %u", game.gameTimeLimitMinutes
        recvOptions = False
        Exit Function
    End If
    tempPlayerLeaveModeValue = 0
    NETuint8_t r, tempPlayerLeaveModeValue
    If tempPlayerLeaveModeValue > PLAYER_LEAVE_MODE_MAX Then
        debug_ 214, LOG_ERROR, "recvOptions", "Invalid playerLeaveMode value specified: %u", tempPlayerLeaveModeValue
        recvOptions = False
        Exit Function
    End If
    game.playerLeaveMode = tempPlayerLeaveModeValue

    For i = 0 To MAX_PLAYERS - 1
        v = NetPlay_players_difficulty(i)
        NETint8_t r, v
        NetPlay_players_difficulty(i) = v
    Next

    ' Send the list of who is still joining
    For i = 0 To MAX_CONNECTED_PLAYERS - 1
        NETbool r, ingame.JoiningInProgress(i)
    Next

    ' Alliances
    For i = 0 To MAX_PLAYERS - 1
        For j = 0 To MAX_PLAYERS - 1
            NETuint8_t r, alliances(i, j)
        Next
    Next

    netPlayersUpdated = True

    ' Make a copy of old structure limits to see what got changed
    oldStructureLimitsCount = ingame.structureLimitsCount
    If oldStructureLimitsCount > 0 Then
        ReDim oldStructureLimits(0 To oldStructureLimitsCount - 1)
        For n = 0 To oldStructureLimitsCount - 1
            oldStructureLimits(n) = ingame.structureLimits(n)
        Next
    Else
        Erase oldStructureLimits
    End If

    ' Free any structure limits we may have in-place
    Erase ingame.structureLimits
    ingame.structureLimitsCount = 0

    ' Get the number of structure limits to expect
    numStructureLimits = 0
    NETuint32_t r, numStructureLimits
    debug_ 252, LOG_NET, "recvOptions", "Host is sending us %u structure limits", numStructureLimits
    If ULT(MAX_STRUCTURE_LIMITS, numStructureLimits) Then
        debug_ 255, LOG_POPUP, "recvOptions", "Number of structure limits (%u) exceeds maximum supported. Incompatible host.", numStructureLimits
        NETend r
        recvOptions = False
        Exit Function
    End If
    ' If there were any changes allocate memory for them
    If numStructureLimits <> 0 Then
        ReDim ingame.structureLimits(0 To numStructureLimits - 1)
        ingame.structureLimitsCount = numStructureLimits
        ' If host have limits changed everyone should load limits too
        ' Do not load limits if mods present because mods are not loaded yet
        If modHashesSize = 0 And Not bLimiterLoaded Then
            initLoadingScreen True
            If Not resLoad("wrf/limiter_data.wrf", 503) Then
                debug_ 270, LOG_INFO, "recvOptions", "Unable to load limiter_data during recvOptions!"
            Else
                bLimiterLoaded = True
                closeLoadingScreen
            End If
        End If
    End If

    For i = 0 To numStructureLimits - 1
        NETuint32_t r, ingame.structureLimits(i).id
        NETuint32_t r, ingame.structureLimits(i).limit
    Next
    NETuint8_t r, ingame.flags

    NETend r

    ' Do not print limits information if we don't have them loaded
    If bLimiterLoaded Then
        ' Check if those vectors are different
        structurelimitsUpdated = (oldStructureLimitsCount <> ingame.structureLimitsCount)
        If Not structurelimitsUpdated Then structurelimitsUpdated = structureLimitsDiffer()

        ' Notify if structure limits were changed
        If structurelimitsUpdated Then
            printStructureLimitsInfo ingame.structureLimits, ingame.structureLimitsCount, WzFunc_new(FN_MOPT_LIMITSINFO_PRINTLINE)
        End If
    Else
        If numStructureLimits <> 0 Then
            tmpConsoleMsgStr = va("Host initialized %u limits, unable to show them due to mods", numStructureLimits)
            addConsoleMessage tmpConsoleMsgStr, DEFAULT_JUSTIFY, SYSTEM_MESSAGE
        End If
    End If

    bRebuildMapList = game.map <> priorGameInfo.map
    If Not bRebuildMapList Then bRebuildMapList = Not Sha256_equals(game.hash, priorGameInfo.hash)
    If Not bRebuildMapList Then bRebuildMapList = modHashesDiffer(game, priorGameInfo)
    If bRebuildMapList Then
        debug_ 316, LOG_INFO, "recvOptions", "Rebuilding map list"
        ' clear out the old level list.
        levShutDown
        levInitialise
        rebuildSearchPath mod_multiplay, True    ' MUST rebuild search path for the new maps we just got!
        buildMapList
    End If

    Set mapData = levFindDataSet_hash(game.map, game.hash, True)
    ' See if we have the map or not
    If mapData Is Nothing Then
        mapName = Left$(game.map, 255)
        removeWildcards mapName

        If Len(mapName) >= 3 Then
            If Mid$(mapName, Len(mapName) - 2, 1) = "-" And Mid$(mapName, Len(mapName) - 1, 1) = "T" And ULT(AscW(Right$(mapName, 1)) - AscW("1"), 3) Then
                mapName = Left$(mapName, Len(mapName) - 3)  ' Cut off "-T1", "-T2" or "-T3".
            End If
        End If
        filename = Left$(va("maps/%uc-%s-%s.wz", game.maxPlayers, mapName, Sha256_toString(game.hash)), 255)  ' Wonder whether game.maxPlayers is initialised already?

        requestResult = requestFile(game.hash, filename)
        Select Case requestResult
        Case FileRequestResult_StartingDownload
            debug_ 387, LOG_INFO, "recvOptions", "Map was not found, requesting map %s from host, type %d", game.map, IIf(game.isMapMod, 1, 0)
            addConsoleMessage "MAP REQUESTED!", DEFAULT_JUSTIFY, SYSTEM_MESSAGE
        Case FileRequestResult_DownloadInProgress
            ' do nothing - just wait
        Case FileRequestResult_FileExists
            debug_ 394, LOG_FATAL, "recvOptions", "Can't load map %s, even though we downloaded %s", game.map, filename
            recvOptions = False
            Exit Function
        Case FileRequestResult_FailedToOpenFileForWriting
            ' TODO: How best to handle? Ideally, message + back out of lobby?
            debug_ 398, LOG_FATAL, "recvOptions", "Failed to open file for writing - unable to download file: %s", filename
            recvOptions = False
            Exit Function
        End Select
    End If

    For n = 0 To game.modHashCount - 1
        filename = Left$(va("mods/downloads/%s", Sha256_toString(game.modHashes(n))), 255)

        requestResult = requestFile(game.modHashes(n), filename)
        Select Case requestResult
        Case FileRequestResult_StartingDownload
            debug_ 412, LOG_INFO, "recvOptions", "Mod was not found, requesting mod %s from host", Sha256_toString(game.modHashes(n))
            addConsoleMessage "MOD REQUESTED!", DEFAULT_JUSTIFY, SYSTEM_MESSAGE
        Case FileRequestResult_DownloadInProgress
            ' do nothing - just wait
        Case FileRequestResult_FileExists
            ' Mod already exists / downloaded
        Case FileRequestResult_FailedToOpenFileForWriting
            ' TODO: How best to handle? Ideally, message + back out of lobby?
            debug_ 423, LOG_FATAL, "recvOptions", "Failed to open file for writing - unable to download file: %s", filename
            recvOptions = False
            Exit Function
        End Select
    Next

    If Not NetPlay_isHost And Not NET_getDownloadingWzFiles_empty() Then
        ' spectators should automatically become not-ready when files remain to be downloaded
        handleAutoReadyRequest
    End If

    If Not mapData Is Nothing Then
        If CheckForMod(mapData.realFileName) Then
            If game.isMapMod Then
                str_ = "Warning, this is a map-mod, it could alter normal gameplay."
            Else
                str_ = "Warning, HOST has altered the game code, and can't be trusted!"
            End If
            addConsoleMessage str_, DEFAULT_JUSTIFY, NOTIFY_MESSAGE
            game.isMapMod = True
        End If
    End If
    If Not mapData Is Nothing Then
        game.isRandom = CheckForRandom(mapData.realFileName, mapData.apDataFiles(0))
    Else
        game.isRandom = False
    End If

    If Not mapData Is Nothing Then
        loadMapPreview False
    End If

    ' (ActivityManager::instance().updateMultiplayGameData(game, ingame, NETGameIsLocked()): the activity sink, not kept)

    recvOptions = True
End Function


' ////////////////////////////////////////////////////////////////////////////
' Tell the host we are leaving the game 'nicely', (we wanted to) and not
' because we have some kind of error. (dropped or disconnected)
Public Function sendLeavingMsg() As Boolean
    Dim w As MessageWriter, host As Boolean, id As Long
    debug_ 511, LOG_NET, "sendLeavingMsg", "We are leaving 'nicely'"
    Set w = NETbeginEncode(NETnetQueue(NetPlay_hostPlayer), NET_PLAYER_LEAVING)
    host = NetPlay_isHost
    id = selectedPlayer

    NETuint32_t w, id
    NETbool w, host
    NETend w
    NETflush

    sendLeavingMsg = True
End Function

' ////////////////////////////////////////////////////////////////////////////
' called in Init.c to shutdown the whole netgame gubbins.
Public Function multiShutdown() As Boolean
    debug_ 530, LOG_MAIN, "multiShutdown", "shutting down networking"
    NETshutdown

    debug_ 533, LOG_MAIN, "multiShutdown", "free game data (structure limits)"
    Erase ingame.structureLimits    ' ingame.structureLimits.clear()
    ingame.structureLimitsCount = 0

    shutdownGameStartScreen

    multiShutdown = True
End Function

Public Function multiStartScreenInit() As Boolean
    bDisplayMultiJoiningStatus = 1 ' always display this
    gameTimeStop
    intHideInterface True
    intHideInGameOptionsButton
    createGameStartScreen WzFunc_new(FN_MOPT_STARTSCREEN_ONCLOSE)

    multiStartScreenInit = True
End Function


Private Sub informOnHostChatPermissionChanges(priorHostChatPermissions() As Boolean)
    Dim numSlotsWithChatPermissionsEnabled As Long, numSlotsWithChatPermissionsDisabled As Long
    Dim numPlayersWithChatPermissionsFreshlyEnabled As Long, numPlayersWithChatPermissionsFreshlyDisabled As Long
    Dim i As Long, msg As String
    Dim playerIdxFreechatEnabled() As Long, nEnabled As Long, playerIdxFreechatDisabled() As Long, nDisabled As Long
    numSlotsWithChatPermissionsEnabled = 0
    numSlotsWithChatPermissionsDisabled = 0
    numPlayersWithChatPermissionsFreshlyEnabled = 0
    numPlayersWithChatPermissionsFreshlyDisabled = 0
    For i = 0 To MAX_CONNECTED_PLAYERS - 1
        If priorHostChatPermissions(i) <> ingame.hostChatPermissions(i) Then
            If isHumanPlayer(i) Then
                If ingame.hostChatPermissions(i) Then
                    numPlayersWithChatPermissionsFreshlyEnabled = numPlayersWithChatPermissionsFreshlyEnabled + 1
                Else
                    numPlayersWithChatPermissionsFreshlyDisabled = numPlayersWithChatPermissionsFreshlyDisabled + 1
                End If
            End If
        End If

        If ingame.hostChatPermissions(i) Then
            numSlotsWithChatPermissionsEnabled = numSlotsWithChatPermissionsEnabled + 1
        Else
            numSlotsWithChatPermissionsDisabled = numSlotsWithChatPermissionsDisabled + 1
        End If
    Next

    If numPlayersWithChatPermissionsFreshlyEnabled = 0 And numPlayersWithChatPermissionsFreshlyDisabled = 0 Then
        ' no changes detected
        Exit Sub
    End If

    If numSlotsWithChatPermissionsEnabled > 0 And numSlotsWithChatPermissionsDisabled = 0 _
        And numPlayersWithChatPermissionsFreshlyEnabled > 1 Then
        ' Enabled for everyone (and more than one person changed)
        addConsoleMessage "The host has enabled free chat for everyone.", DEFAULT_JUSTIFY, SYSTEM_MESSAGE
        Exit Sub
    End If

    If numSlotsWithChatPermissionsEnabled = 0 And numSlotsWithChatPermissionsDisabled > 0 _
        And numPlayersWithChatPermissionsFreshlyDisabled > 1 Then
        ' Disabled for everyone
        addConsoleMessage "The host has muted free chat for everyone.", DEFAULT_JUSTIFY, SYSTEM_MESSAGE
        Exit Sub
    End If

    ' Otherwise, output changes for each changed player
    ReDim playerIdxFreechatEnabled(0 To MAX_CONNECTED_PLAYERS - 1)
    ReDim playerIdxFreechatDisabled(0 To MAX_CONNECTED_PLAYERS - 1)
    For i = 0 To MAX_CONNECTED_PLAYERS - 1
        If priorHostChatPermissions(i) <> ingame.hostChatPermissions(i) Then
            If Not isHumanPlayer(i) Then
                ' skip notices for non-human slots
                GoTo continue_
            End If
            If i = selectedPlayer Then
                If GetGameMode() <> GS_NORMAL Then
                    ' skip notices for self when in lobby - those are handled elsewhere
                    GoTo continue_
                End If
            End If
            If ingame.hostChatPermissions(i) Then
                ' Host enabled free chat for player
                playerIdxFreechatEnabled(nEnabled) = i
                nEnabled = nEnabled + 1
            Else
                ' Host disabled free chat for player
                playerIdxFreechatDisabled(nDisabled) = i
                nDisabled = nDisabled + 1
            End If
        End If
continue_:
    Next

    If nEnabled <> 0 Then
        If nEnabled = 1 Then
            msg = va("The host has enabled free chat for player: %s", getPlayerName(playerIdxFreechatEnabled(0)))
            addConsoleMessage msg, DEFAULT_JUSTIFY, SYSTEM_MESSAGE
        Else
            msg = va("The host has enabled free chat for %d players", nEnabled)
            addConsoleMessage msg, DEFAULT_JUSTIFY, SYSTEM_MESSAGE
        End If
    End If

    If nDisabled <> 0 Then
        If nDisabled = 1 Then
            msg = va("The host has muted free chat for player: %s", getPlayerName(playerIdxFreechatDisabled(0)))
            addConsoleMessage msg, DEFAULT_JUSTIFY, SYSTEM_MESSAGE
        Else
            msg = va("The host has muted free chat for %d players", nDisabled)
            addConsoleMessage msg, DEFAULT_JUSTIFY, SYSTEM_MESSAGE
        End If
    End If
End Sub

' static void NETlockedOptions(MessageReader &r, MultiplayOptionsLocked& lockedOpts)
Private Sub NETlockedOptions_r(ByVal r As MessageReader, lockedOpts As MultiplayOptionsLocked)
    NETbool r, lockedOpts.scavengers
    NETbool r, lockedOpts.alliances
    NETbool r, lockedOpts.teams
    NETbool r, lockedOpts.power
    NETbool r, lockedOpts.difficulty
    NETbool r, lockedOpts.ai
    NETbool r, lockedOpts.position
    NETbool r, lockedOpts.bases
    NETbool r, lockedOpts.spectators
    NETbool r, lockedOpts.name_
    NETbool r, lockedOpts.readybeforefull
    NETbool r, lockedOpts.cheats
End Sub

' ////////////////////////////////////////////////////////////////////////////
' host config (options that don't impact gameplay but might impact things like chat). (recvd in both frontend and in-game)
' returns: false if the host config details should be considered invalid and the client should disconnect
Public Function recvHostConfig(ByVal queue As Long) As Boolean
    Dim priorHostChatPermissions(0 To MAX_CONNECTED_PLAYERS - 1) As Boolean, i As Long, r As MessageReader
    Dim newLockedOpts As MultiplayOptionsLocked
    If NETQUEUE_index(queue) <> NetPlay_hostPlayer Then ASSERT_FAILURE "queue.index == NetPlay.hostPlayer", "src/multiopt.cpp:878", "recvHostConfig", "NET_HOST_CONFIG received from unexpected player: %u - ignoring", NETQUEUE_index(queue): recvHostConfig = True: Exit Function

    For i = 0 To MAX_CONNECTED_PLAYERS - 1
        priorHostChatPermissions(i) = ingame.hostChatPermissions(i)
    Next

    debug_ 882, LOG_NET, "recvHostConfig", "Receiving host_config from host"
    Set r = NETbeginDecode(queue, NET_HOST_CONFIG)

    ' Host-set player chat permissions
    For i = 0 To MAX_CONNECTED_PLAYERS - 1
        NETbool r, ingame.hostChatPermissions(i)
    Next

    ' "Locked" options configuration
    NETlockedOptions_r r, newLockedOpts
    updateLockedOptionsFromHost newLockedOpts

    NETend r

    informOnHostChatPermissionChanges priorHostChatPermissions
    multiLobbyHandleHostOptionsChanges priorHostChatPermissions

    recvHostConfig = True
End Function

' void printStructureLimitsInfo(std::vector<MULTISTRUCTLIMITS>& structureLimits, const PrintStructureLimitsLineFunc& printLineFunc)
' (structureLimits(0 .. structureLimitsCount - 1); printLineFunc: a WzFunc taking the line)
Public Sub printStructureLimitsInfo(structureLimits() As MULTISTRUCTLIMITS, ByVal structureLimitsCount As Long, ByVal printLineFunc As WzFunc)
    Dim nondefaultlimitsize As Long, i As Long, changedNum As Long, structname As String, tmpConsoleMsgStr As String
    Dim id As Long
    If Not bLimiterLoaded Then
        ' Can't print full limits information if we don't have them loaded
        If structureLimitsCount <> 0 Then
            WzFunc_call printLineFunc, va("Host initialized %u limits, unable to show them due to mods", structureLimitsCount)
        Else
            WzFunc_call printLineFunc, "Limits set to default."
        End If
        Exit Sub
    End If

    ' need to calculate nondefaultlimitsize
    nondefaultlimitsize = 0
    For i = 0 To structureLimitsCount - 1
        id = structureLimits(i).id
        If ULT(id, numStructureStats) Then
            ' (VB) ASSERT(asStructureStats != nullptr, ...): asStructureStats is an array
            If asStructureStats(id).base.limit <> structureLimits(i).limit Then
                nondefaultlimitsize = nondefaultlimitsize + 1
            End If
        End If
    Next

    If nondefaultlimitsize = 0 Then
        WzFunc_call printLineFunc, "Limits set to default."
        Exit Sub
    End If

    WzFunc_call printLineFunc, va("Changed structure limits [%u]:", nondefaultlimitsize)
    changedNum = 1
    For i = 0 To structureLimitsCount - 1
        id = structureLimits(i).id
        If ULT(id, numStructureStats) Then
            If asStructureStats(id).base.limit <> structureLimits(i).limit Then
                structname = asStructureStats(id).name_
                If asStructureStats(id).base.limit <> LOTS_OF Then
                    tmpConsoleMsgStr = va("[%d] Limit [%s]: %u (default: %u)", changedNum, structname, structureLimits(i).limit, asStructureStats(id).base.limit)
                Else
                    tmpConsoleMsgStr = va("[%d] Limit [%s]: %u (default: no limit)", changedNum, structname, structureLimits(i).limit)
                End If
                WzFunc_call printLineFunc, tmpConsoleMsgStr
                changedNum = changedNum + 1
            End If
        Else
            tmpConsoleMsgStr = va("[%u] Limit that is bigger than numStructureStats (%u): %u", i, structureLimits(i).id, structureLimits(i).limit)
            WzFunc_call printLineFunc, tmpConsoleMsgStr
        End If
    Next
End Sub

' the module's lambdas (WzFunc ids FN_MOPT_*)
Public Function multiopt_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim line_ As String
    Select Case f.id
    Case FN_MOPT_LIMITSINFO_PRINTLINE           ' [](const std::string& line)
        line_ = args(0)
        addConsoleMessage line_, DEFAULT_JUSTIFY, SYSTEM_MESSAGE

    Case FN_MOPT_STARTSCREEN_ONCLOSE            ' createGameStartScreen([] {...})
        ' on game start overlay screen close...
        bDisplayMultiJoiningStatus = 0
        intShowInterface
        ' restart gameTime
        gameTimeStart
    End Select
End Function
