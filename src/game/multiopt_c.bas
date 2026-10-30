Attribute VB_Name = "multiopt_c"
' src/multiopt.cpp, multiopt.h: so far the host's options (sendOptions,
' sendHostConfig and its NETlockedOptions writer: NET_OPTIONS and
' NET_HOST_CONFIG go nowhere before cluster J, but sendOptions' game.modHashes
' and its clamps of the options are kept), hostCampaign and the game's
' start-up (gameInit, playerResponding, multiGameInit). ingame.startTime and
' the other std::chrono::steady_clock::time_points are QueryPerformanceCounter
' counts. Not ported yet: recvOptions, recvHostConfig and the options'
' exchange (cluster J), the ActivityManager's updateMultiplayGameData (the
' activity sink, not kept), multiShutdown, multiStartScreenInit and the start
' screen (clusters F and J). multiGameShutdown ends every game.
Option Explicit

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' lib/netplay/nettypes.h
Private Const MAX_CONNECTED_PLAYERS As Long = MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS

Private Const MAX_STRUCTURE_LIMITS As Long = 4096 ' Set a high (but explicit) maximum for the number of structure limits supported

' (VB) Sha256::Bytes
Private Const Sha256_Bytes As Long = 32

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
' (VB) NEThostGame's playerIdentity (playerStats.identity) is left out: the identities, cluster J
Public Function hostCampaign(ByVal SessionName As String, ByVal hostPlayerName As String, ByVal spectatorHost As Boolean, ByVal skipResetAIs As Boolean, ByVal desiredOpenSpectatorSlots As Long) As Boolean
    Dim playerStats As PLAYERSTATS_, i As Long
    debug_ 459, LOG_WZ, "hostCampaign", "Hosting campaign: '%s', player: '%s'", SessionName, hostPlayerName

    freeMessages

    loadMultiStats hostPlayerName, playerStats

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

    multiStatsSetVerifiedHostIdentityFromJoin  ' (playerStats.identity.toBytes(EcKey::Public): the identities, cluster J)

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

' ////////////////////////////////////////////////////////////////////////////
'called when the game finally gets fired up.
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
    ' (ingame.lastLagCheck and lastDesyncCheck: the network's checks, cluster J)
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
