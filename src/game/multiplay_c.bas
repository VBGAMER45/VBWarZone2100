Attribute VB_Name = "multiplay_c"
' src/multiplay.cpp, multiplay.h: so far turnOffMultiMsg, the IdTo* lookups
' (_IdToStruct is IdToStruct_), isHumanPlayer, whosResponsible,
' myResponsibility, responsibleFor, canGiveOrdersFor, scavengerSlot,
' scavengerPlayer, getPlayerGenericName, the sync requests, the game
' messages' dispatch (getMessageHandlingAction, shouldProcessMessage,
' recvMessage, HandleBadParam), the research messages (SendResearch,
' recvResearch, sendResearchStatus, recvResearchStatus),
' findResearchingFacilityByResearchIndex (the list overload is _lists),
' recvDestroyFeature, NetworkTextMessage (its constructor, enqueue,
' receive and decode), printInGameTextMessage, the chat messages
' (sendInGameSystemMessage, receiveInGameTextMessage, recvTextMessageAI,
' the spectators' messages), the beacons (sendBeacon, CreateBeaconViewData,
' findBeaconMsg, addBeaconBlip, sendBeaconToPlayer, recvBeacon),
' resetReadyStatus, findPlayerIndexByPosition, makePlayerSpectator,
' the auto-kick routines (lag, desync, not ready), the win sequence,
' setPlayerName, setPlayerMuted, the data check 2, the map and mod
' transfer (sendMap, recvMapFileRequested, recvMapFileData),
' multiPlayerLoop and openchannels (multiplay_c_init gives its
' initialiser). A NETQUEUE is WzNet's Long (NETQUEUE_index, ...). Its
' globals (game, ingame, bMultiPlayer, bMultiMessages, ...) are still
' stubs_game's.
Option Explicit

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

' lib/netplay/nettypes.h
Private Const MAX_CONNECTED_PLAYERS As Long = MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS

' multiplay.h
Public Const ANYPLAYER As Long = 99
Public Const MAX_KICK_REASON As Long = 1024     ' max array size for the reason your kicking someone

Public Const MPFLAGS_NO_TANKS As Long = &H1      ' Flag for tanks disabled
Public Const MPFLAGS_NO_CYBORGS As Long = &H2    ' Flag for cyborgs disabled
Public Const MPFLAGS_NO_VTOLS As Long = &H4      ' Flag for VTOLs disabled
Public Const MPFLAGS_NO_UPLINK As Long = &H8     ' Flag for Satellite Uplink disabled
Public Const MPFLAGS_NO_LASSAT As Long = &H10    ' Flag for Laser Satellite Command Post disabled
Public Const MPFLAGS_FORCELIMITS As Long = &H20  ' Flag to force structure limits

Public Enum HandleMessageAction
    HandleMessageAction_Process_Message
    HandleMessageAction_Silently_Ignore
    HandleMessageAction_Disallow_And_Kick_Sender
End Enum

Public bDisplayMultiJoiningStatus As Long       ' UBYTE

Public openchannels(0 To MAX_CONNECTED_PLAYERS - 1) As Boolean ' outgoing message channels (i.e. who you are willing to send messages to)
' multiplay.h: struct NetworkTextMessage (Sender can be a player index, SYSTEM_MESSAGE or NOTIFY_MESSAGE)
Public Type NetworkTextMessage
    sender As Long                      ' int32_t
    text As String                      ' char text[MAX_CONSOLE_STRING_LENGTH]
    teamSpecific As Boolean             ' false
End Type

Public beaconReceiveMsg(0 To MAX_PLAYERS - 1) As String    ' char[MAX_PLAYERS][MAX_CONSOLE_STRING_LENGTH]: beacon msg for each player

'--------------------------------------------------------------------------
' Cluster J (fork JL): the network's game side of multiplay.cpp: the auto-kick routines (lag, desync, not ready), the
' win sequence, setPlayerName, setPlayerMuted, the data check 2, the chat messages' decoding and receiving, the map and
' mod transfer (the files through NetPlay.players[i].wzFiles, WzNet's). std::chrono::steady_clock::time_points are
' QueryPerformanceCounter counts (perfFreqJ the frequency).

Private Declare Function QueryPerformanceCounter_J Lib "kernel32" Alias "QueryPerformanceCounter" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency_J Lib "kernel32" Alias "QueryPerformanceFrequency" (lpFrequency As Currency) As Long
Private perfFreqJ As Currency
Private Const MAX_CONNECTED_PLAYERS_J As Long = MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS
Private Const MAX_NET_TRANSFERRABLE_FILE_SIZE As Long = &H8000000    ' netplay.h

Private Const DATACHECK2_INTERVAL_MS As Long = 10000
Private Const LAG_INITIAL_LOAD_GRACEPERIOD As Long = 60
Private Const LAG_CHECK_INTERVAL As Long = 1000
Private Const DESYNC_CHECK_INTERVAL As Long = 1000
Private Const NOTREADY_CHECK_INTERVAL As Long = 1000

' multiplayerWinSequence's statics
Private winSeqPos As Vector3i          ' static Position pos = Position(0, 0, 0);
Private winSeqLast As Long             ' static UDWORD last = 0;

' (VB) the C++ initialisers of this module's globals (Sub Main calls it)
Public Sub multiplay_c_init()
    openchannels(0) = True                      ' {true}
End Sub

' temporarily disable multiplayer mode.
Public Sub turnOffMultiMsg(ByVal bDoit As Boolean)
    If Not bMultiPlayer Then
        Exit Sub
    End If

    bMultiMessages = Not bDoit
End Sub

' ////////////////////////////////////////////////////////////////////////////
' quikie functions.

' to get droids ...
Public Function IdToDroid(ByVal id As Long, ByVal player As Long) As Long
    Dim i As Long, d As Long
    If player = ANYPLAYER Then
        For i = 0 To MAX_PLAYERS - 1
            d = getBaseObjFromId_list(apsDroidLists(i), id)
            If d <> 0 Then
                IdToDroid = d
                Exit Function
            End If
        Next
    ElseIf ULT(player, MAX_PLAYERS) Then
        d = getBaseObjFromId_list(apsDroidLists(player), id)
        If d <> 0 Then
            IdToDroid = d
            Exit Function
        End If
    End If
    IdToDroid = 0
End Function

' find off-world droids
Public Function IdToMissionDroid(ByVal id As Long, ByVal player As Long) As Long
    Dim i As Long, d As Long
    If player = ANYPLAYER Then
        For i = 0 To MAX_PLAYERS - 1
            d = getBaseObjFromId_list(mission_apsDroidLists(i), id)
            If d <> 0 Then
                IdToMissionDroid = d
                Exit Function
            End If
        Next
    ElseIf ULT(player, MAX_PLAYERS) Then
        d = getBaseObjFromId_list(mission_apsDroidLists(player), id)
        If d <> 0 Then
            IdToMissionDroid = d
            Exit Function
        End If
    End If
    IdToMissionDroid = 0
End Function

' static STRUCTURE* _IdToStruct(UDWORD id, UDWORD beginPlayer, UDWORD endPlayer)
Private Function IdToStruct_(ByVal id As Long, ByVal beginPlayer As Long, ByVal endPlayer As Long) As Long
    Dim i As Long, s As Long
    i = beginPlayer
    Do While ULT(i, endPlayer)
        s = getBaseObjFromId_list(apsStructLists(i), id)
        If s <> 0 Then
            IdToStruct_ = s
            Exit Function
        End If
        s = getBaseObjFromId_list(mission_apsStructLists(i), id)
        If s <> 0 Then
            IdToStruct_ = s
            Exit Function
        End If
        i = i + 1
    Loop
    IdToStruct_ = 0
End Function

' ////////////////////////////////////////////////////////////////////////////
' find a structure
Public Function IdToStruct(ByVal id As Long, ByVal player As Long) As Long
    Dim beginPlayer As Long, endPlayer As Long, out As Long
    beginPlayer = 0
    endPlayer = MAX_PLAYERS
    If player <> ANYPLAYER Then
        beginPlayer = player
        endPlayer = MIN_(player + 1, MAX_PLAYERS)
    End If
    out = 0
    out = IdToStruct_(id, beginPlayer, endPlayer)
    If out <> 0 Then
        IdToStruct = out
        Exit Function
    End If
    IdToStruct = 0
End Function

' ////////////////////////////////////////////////////////////////////////////
' find a feature
Public Function IdToFeature(ByVal id As Long, ByVal player As Long) As Long
    ' (void)player: unused, all features go into player 0
    IdToFeature = getBaseObjFromId_list(apsFeatureList(0), id)
End Function

' ////////////////////////////////////////////////////////////////////////////

Public Function IdToTemplate(ByVal tempId As Long, ByVal player As Long) As Long
    Dim i As Long, psTempl As Long
    ' Check if we know which player this is from, in that case, assume it is a player template
    ' FIXME: nuke the ANYPLAYER hack
    If player <> ANYPLAYER And ULT(player, MAX_PLAYERS) Then
        IdToTemplate = findPlayerTemplateById(player, tempId)
        Exit Function
    End If

    ' It could be a AI template...or that of another player
    For i = 0 To MAX_PLAYERS - 1
        psTempl = findPlayerTemplateById(i, tempId)
        If psTempl <> 0 Then
            IdToTemplate = psTempl
            Exit Function
        End If
    Next

    ' no error, since it is possible that we don't have this template defined yet.
    IdToTemplate = 0
End Function

' ///////////////////////////////////////////////////////////////////////////////
'  Returns a pointer to base object, given an id and optionally a player.
Public Function IdToPointer(ByVal id As Long, ByVal player As Long) As Long
    Dim pD As Long, pS As Long, pF As Long
    ' droids.

    pD = IdToDroid(id, player)
    If pD <> 0 Then
        IdToPointer = pD
        Exit Function
    End If

    ' structures
    pS = IdToStruct(id, player)
    If pS <> 0 Then
        IdToPointer = pS
        Exit Function
    End If

    ' features
    pF = IdToFeature(id, player)
    If pF <> 0 Then
        IdToPointer = pF
        Exit Function
    End If

    IdToPointer = 0
End Function

'to tell if the player is a computer or not.
Public Function isHumanPlayer(ByVal player As Long) As Boolean
    If player >= MAX_CONNECTED_PLAYERS Or player < 0 Then
        isHumanPlayer = False   ' obvious, really
        Exit Function
    End If
    isHumanPlayer = NetPlay_players_allocated(player)
End Function

' Clear player name data after game quit.
Public Sub clearPlayerName(ByVal player As Long)
    If Not (player >= 0 And player < MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "player < MAX_CONNECTED_PLAYERS", "src/multiplay.cpp:911", "clearPlayerName", "Player index (%u) out of range", player: Exit Sub
    NetPlay_players_name(player) = ""
End Sub

' returns player responsible for 'player'
Public Function whosResponsible(ByVal player As Long) As Long
    If isHumanPlayer(player) Then
        whosResponsible = player            ' Responsible for him or her self
    ElseIf player = realSelectedPlayer Then
        whosResponsible = player            ' We are responsibly for ourselves
    Else
        whosResponsible = NetPlay_hostPlayer   ' host responsible for all AIs
    End If
End Function

'returns true if selected player is responsible for 'player'
Public Function myResponsibility(ByVal player As Long) As Boolean
    myResponsibility = (whosResponsible(player) = selectedPlayer Or whosResponsible(player) = realSelectedPlayer)
End Function

'returns true if 'player' is responsible for 'playerinquestion'
Public Function responsibleFor(ByVal player As Long, ByVal playerinquestion As Long) As Boolean
    responsibleFor = whosResponsible(playerinquestion) = player
End Function

Public Function canGiveOrdersFor(ByVal player As Long, ByVal playerInQuestion As Long) As Boolean
    If playerInQuestion >= 0 And playerInQuestion < MAX_PLAYERS Then
        canGiveOrdersFor = player = playerInQuestion Or responsibleFor(player, playerInQuestion) Or debugMappingsAllowed()
    End If
End Function

Public Function scavengerSlot() As Long
    ' Scavengers used to always be in position 7, when scavengers were only supported in less than 8 player maps.
    ' Scavengers should be in position N in N-player maps, where N >= 8.
    scavengerSlot = MAX_(game.maxPlayers, 7)
End Function

Public Function scavengerPlayer() As Long
    If game.scavengers <> NO_SCAVENGERS Then scavengerPlayer = scavengerSlot() Else scavengerPlayer = -1
End Function

' ////////////////////////////////////////////////////////////////////////////
' probably temporary. Places the camera on the players 1st droid or struct.
Public Function cameraToHome(ByVal player As Long, ByVal scroll As Boolean, ByVal fromSave As Boolean) As Vector3i
    Dim x As Long, y As Long, psBuilding As Long, it As Long, res As Vector3i
    psBuilding = 0

    If ULT(player, MAX_PLAYERS) Then
        ' std::find_if(..., building->pStructureType->type == REF_HQ)
        it = apsStructLists(player).head
        Do While it <> 0
            If asStructureStats(aStruct(lnObj(it)).pStructureType).typ = REF_HQ Then
                psBuilding = lnObj(it)
                Exit Do
            End If
            it = lnNext(it)
        Loop
    End If

    If psBuilding <> 0 Then
        x = map_coord(aStruct(psBuilding).pos.x)
        y = map_coord(aStruct(psBuilding).pos.y)
    ElseIf ULT(player, MAX_PLAYERS) And apsDroidLists(player).size <> 0 Then ' or first droid
        x = map_coord(aDroid(lnObj(apsDroidLists(player).head)).pos.x)
        y = map_coord(aDroid(lnObj(apsDroidLists(player).head)).pos.y)
    ElseIf ULT(player, MAX_PLAYERS) And apsStructLists(player).size <> 0 Then ' center on first struct
        x = map_coord(aStruct(lnObj(apsStructLists(player).head)).pos.x)
        y = map_coord(aStruct(lnObj(apsStructLists(player).head)).pos.y)
    Else                                ' or map center.
        x = mapWidth \ 2
        y = mapHeight \ 2
    End If

    If scroll Then
        requestRadarTrack world_coord(x), world_coord(y)
    ElseIf Not fromSave Then ' This will override the saved camera position in skirmish games if not checked
        setViewPos x, y, True
    End If

    res.x = world_coord(x)
    res.y = map_TileHeight(x, y)
    res.z = world_coord(y)
    cameraToHome = res
End Function

Private Sub recvSyncRequest(ByVal queue As Long)
    Dim req_id As Long, x As Long, y As Long, obj_id As Long, obj_id2 As Long, player_id As Long, player_id2 As Long
    Dim psObj As Long, psObj2 As Long, r As MessageReader

    Set r = NETbeginDecode(queue, GAME_SYNC_REQUEST)
    NETint32_t r, req_id
    NETint32_t r, x
    NETint32_t r, y
    NETint32_t r, obj_id
    NETint32_t r, player_id
    NETint32_t r, obj_id2
    NETint32_t r, player_id2
    NETend r

    syncDebug "recvSyncRequest", "sync request received from%d req_id%d x%u y%u %obj1 %obj2", NETQUEUE_index(queue), req_id, x, y, obj_id, obj_id2
    If obj_id <> 0 Then
        psObj = IdToPointer(obj_id, player_id)
    End If
    If obj_id2 <> 0 Then
        psObj2 = IdToPointer(obj_id2, player_id2)
    End If
    triggerEventSyncRequest NETQUEUE_index(queue), req_id, x, y, psObj, psObj2
End Sub

Private Sub sendObj(ByVal w As MessageWriter, ByVal psObj As Long)
    Dim obj_id As Long, player As Long, dummy As Long
    If psObj <> 0 Then
        obj_id = SIMPLE_OBJECT_id(psObj)
        player = SIMPLE_OBJECT_player(psObj)
        NETint32_t w, obj_id
        NETint32_t w, player
    Else
        dummy = 0
        NETint32_t w, dummy
        NETint32_t w, dummy
    End If
End Sub

Public Sub sendSyncRequest(ByVal req_id As Long, ByVal x As Long, ByVal y As Long, ByVal psObj As Long, ByVal psObj2 As Long)
    Dim w As MessageWriter
    Set w = NETbeginEncode(NETgameQueue(realSelectedPlayer), GAME_SYNC_REQUEST)
    NETint32_t w, req_id
    NETint32_t w, x
    NETint32_t w, y
    sendObj w, psObj
    sendObj w, psObj2
    NETend w
End Sub


Public Function isBlindPlayerInfoState() As Boolean
    Select Case game.blindMode
    Case BLIND_MODE_NONE
        isBlindPlayerInfoState = False
    Case BLIND_MODE_BLIND_LOBBY, BLIND_MODE_BLIND_LOBBY_SIMPLE_LOBBY
        ' blind when game hasn't fully started yet
        isBlindPlayerInfoState = Not ingame.has_TimeEveryoneIsInGame
    Case BLIND_MODE_BLIND_GAME, BLIND_MODE_BLIND_GAME_SIMPLE_LOBBY
        ' blind when game hasn't ended yet
        isBlindPlayerInfoState = Not ingame.has_endTime
    Case Else
        isBlindPlayerInfoState = False ' silence warning
    End Select
End Function

' (VB) the name of a player (the AI players' with their colour in a game; "Commander" for one without a
' name: the campaign's): C++ returns the AIs' from a static array, so that tool tips can keep it
Public Function getPlayerName(ByVal player As Long, Optional ByVal treatAsNonHost As Boolean = False) As String
    If vbDisplayOff Then getPlayerName = "": Exit Function  ' (VB) the vbtest runs: statsref's stand-in
    Dim aiPlayer As Boolean
    If ULT(player, NetPlay_players_size()) Then aiPlayer = (NetPlay_players_ai(player) >= 0) And Not NetPlay_players_allocated(player)

    If aiPlayer And GetGameMode() = GS_NORMAL And Not challengeActive Then
        If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/multiplay.cpp:822", "getPlayerName", "invalid player: %d", player: getPlayerName = "": Exit Function
        ' Add colour to player name.
        getPlayerName = getPlayerColourName(player) & "-" & NetPlay_players_name(player)
        Exit Function
    End If

    If Not ULT(player, NetPlay_players_size()) Then
        ' for campaign and tutorials
        getPlayerName = "Commander"
        Exit Function
    End If
    If Len(NetPlay_players_name(player)) = 0 Then
        getPlayerName = "Commander"
        Exit Function
    End If

    If isBlindPlayerInfoState() Then
        If (Not NetPlay_isHost Or NetPlay_hostPlayer < MAX_PLAYER_SLOTS Or treatAsNonHost) And Not NETisReplay() Then
            ' Get stable "generic" names (unless it's a spectator host)
            If player <> NetPlay_hostPlayer Or NetPlay_hostPlayer < MAX_PLAYER_SLOTS Then
                getPlayerName = getPlayerGenericName(player)
                Exit Function
            End If
        End If
    End If

    getPlayerName = NetPlay_players_name(player)
End Function

' return a "generic" player name that is fixed based on the player idx (useful for blind mode games)
Public Function getPlayerGenericName(ByVal player As Long) As String
    Dim genericNames As Variant
    ' genericNames are *not* localized - we want the same display across all systems (just like player-set names)
    genericNames = Array("Alpha", "Beta", "Gamma", "Delta", "Epsilon", "Zeta", "Omega", "Theta", "Iota", "Kappa", "Lambda", "Omicron", "Pi", "Rho", "Sigma", "Tau")
    If Not (ULT(player, 16)) Then ASSERT_FAILURE "player < genericNames.size()", "src/multiplay.cpp:876", "getPlayerGenericName", "player number (%d) exceeds maximum (%zu)", player, 16

    If Not ULT(player, 16) Then
        If player < MAX_PLAYERS Then getPlayerGenericName = "Player" Else getPlayerGenericName = "Spectator"
        Exit Function
    End If

    If player >= MAX_PLAYER_SLOTS Then
        getPlayerGenericName = "Spectator"
        Exit Function
    End If

    getPlayerGenericName = genericNames(player)
End Function

Public Function getMessageHandlingAction(ByVal queue As Long, ByVal typ As Long) As HandleMessageAction
    Dim senderIsSpectator As Boolean, senderIsAdmin As Boolean
    If NETQUEUE_index(queue) = NetPlay_hostPlayer Then
        ' host gets access to all messages
        getMessageHandlingAction = HandleMessageAction_Process_Message
        Exit Function
    End If

    senderIsSpectator = NetPlay_players_isSpectator(NETQUEUE_index(queue))
    senderIsAdmin = NetPlay_players_isAdmin(NETQUEUE_index(queue))

    If typ > NET_MIN_TYPE And typ < NET_MAX_TYPE Then
        Select Case typ
        Case NET_OPTIONS, NET_PLAYER_INFO, NET_PLAYER_JOINED, NET_FILE_PAYLOAD, NET_VOTE_REQUEST, NET_HOST_CONFIG
            ' only the host is allowed to send these messages
            If NETQUEUE_index(queue) <> NetPlay_hostPlayer Then
                getMessageHandlingAction = HandleMessageAction_Disallow_And_Kick_Sender
                Exit Function
            End If
        Case NET_KICK, NET_TEAMREQUEST, NET_FACTIONREQUEST, NET_POSITIONREQUEST
            ' NET_TEAMREQUEST: spectators should not be allowed to request a team / non-spectator slot status
            If senderIsSpectator And Not senderIsAdmin Then
                getMessageHandlingAction = HandleMessageAction_Disallow_And_Kick_Sender
                Exit Function
            End If
        Case NET_AITEXTMSG, NET_BEACONMSG
            If senderIsSpectator Then
                getMessageHandlingAction = HandleMessageAction_Disallow_And_Kick_Sender
                Exit Function
            End If
        Case NET_TEXTMSG
            ' Normal chat messages are available to spectators in the game room / lobby chat, but *not* in-game
            If senderIsSpectator And GetGameMode() = GS_NORMAL Then
                If gameInitialised And bDisplayMultiJoiningStatus = 0 Then
                    ' If the game is actually initialized and everyone has joined the game, treat this as a kickable offense
                    getMessageHandlingAction = HandleMessageAction_Disallow_And_Kick_Sender
                    Exit Function
                Else
                    ' Otherwise just silently ignore it
                    getMessageHandlingAction = HandleMessageAction_Silently_Ignore
                    Exit Function
                End If
            End If
        Case NET_SPECTEXTMSG
            If Not senderIsSpectator Then
                getMessageHandlingAction = HandleMessageAction_Silently_Ignore
                Exit Function
            End If
        Case NET_COLOURREQUEST
            ' for now, *must* be allowed
            getMessageHandlingAction = HandleMessageAction_Process_Message
            Exit Function
        Case NET_DATA_CHECK2
            If senderIsSpectator Then
                getMessageHandlingAction = HandleMessageAction_Silently_Ignore
                Exit Function
            End If
        Case Else
            ' certain messages are always allowed, no matter who it is
            getMessageHandlingAction = HandleMessageAction_Process_Message
            Exit Function
        End Select
    End If

    If typ > GAME_MIN_TYPE And typ < GAME_MAX_TYPE Then
        Select Case typ
        Case GAME_GAME_TIME, GAME_PLAYER_LEFT
            ' always allowed
            getMessageHandlingAction = HandleMessageAction_Process_Message
            Exit Function
        Case GAME_SYNC_REQUEST
            If senderIsSpectator Then
                getMessageHandlingAction = HandleMessageAction_Silently_Ignore
                Exit Function
            End If
        Case GAME_DEBUG_MODE, GAME_DEBUG_ADD_DROID, GAME_DEBUG_ADD_STRUCTURE, GAME_DEBUG_ADD_FEATURE, GAME_DEBUG_REMOVE_DROID, _
             GAME_DEBUG_REMOVE_STRUCTURE, GAME_DEBUG_REMOVE_FEATURE, GAME_DEBUG_FINISH_RESEARCH
            If senderIsSpectator Then
                getMessageHandlingAction = HandleMessageAction_Disallow_And_Kick_Sender
                Exit Function
            End If
        Case Else
            If senderIsSpectator Then
                getMessageHandlingAction = HandleMessageAction_Disallow_And_Kick_Sender
                Exit Function
            End If
        End Select
    End If

    If typ = REPLAY_ENDED Then
        getMessageHandlingAction = HandleMessageAction_Silently_Ignore
        Exit Function
    End If

    getMessageHandlingAction = HandleMessageAction_Process_Message
End Function

Public Function shouldProcessMessage(ByVal queue As Long, ByVal typ As Long) As Boolean
    Dim action As HandleMessageAction, buf As String, senderPlayerIdx As Long
    action = getMessageHandlingAction(queue, typ)
    Select Case action
    Case HandleMessageAction_Process_Message
        shouldProcessMessage = True
        Exit Function
    Case HandleMessageAction_Silently_Ignore
        NETpop queue    ' remove message from queue
        shouldProcessMessage = False
        Exit Function
    Case HandleMessageAction_Disallow_And_Kick_Sender
        NETpop queue    ' remove message from queue
        If NetPlay_isHost Then
            ' kick sender for sending unauthorized message
            senderPlayerIdx = NETQUEUE_index(queue)
            debug_ 1444, LOG_INFO, "shouldProcessMessage", "Auto kicking player %s, invalid command received: %s", getPlayerName(senderPlayerIdx), messageTypeToString(typ)
            buf = Left$(va("Auto kicking player %s, invalid command received: %u", getPlayerName(senderPlayerIdx, True), typ), 254)    ' char buf[255]
            sendInGameSystemMessage buf
            kickPlayer NETQUEUE_index(queue), "Unauthorized network command", ERROR_INVALID, False
        End If
        shouldProcessMessage = False
        Exit Function
    End Select
    shouldProcessMessage = False    ' silence warnings
End Function

' ////////////////////////////////////////////////////////////////////////////
' ////////////////////////////////////////////////////////////////////////////
' Recv Messages. Get a message and dispatch to relevant function.
Public Function recvMessage() As Boolean
    Dim queue As Long, typ As Long
    Dim processedMessage1 As Boolean, processedMessage2 As Boolean
    Dim r As MessageReader, player_id As Long, reason As String, KICK_TYPE As Long, buf As String

    Do  ' for all incoming messages.
        If Not NETrecvNet(queue, typ) Then
            If Not NETrecvGame(queue, typ) Then Exit Do
        End If
        processedMessage1 = False
        processedMessage2 = False

        If NETQUEUE_queueType(queue) = QUEUE_GAME Then
            syncDebug "recvMessage", "Processing player %d, message %s", NETQUEUE_index(queue), messageTypeToString(typ)
        End If

        If Not shouldProcessMessage(queue, typ) Then
            GoTo continue_
        End If

        ' messages only in game.
        If Not ingame.localJoiningInProgress Then
            processedMessage1 = True
            Select Case typ
            Case GAME_DROIDINFO                 'droid update info
                recvDroidInfo queue
            Case NET_TEXTMSG                    ' simple text message
                receiveInGameTextMessage queue
            Case NET_DATA_CHECK
                recvDataCheck queue
            Case NET_DATA_CHECK2
                recvDataCheck2 queue
            Case NET_AITEXTMSG                  'multiplayer AI text message
                recvTextMessageAI queue
            Case NET_SPECTEXTMSG
                recvSpecInGameTextMessage queue ' multiplayer spectator text message
            Case NET_BEACONMSG                  'beacon (blip) message
                recvBeacon queue
            Case NET_TEAM_STRATEGY
                recvStrategyPlanUpdate queue
            Case NET_QUICK_CHAT_MSG
                recvQuickChat queue
            Case GAME_SYNC_REQUEST
                recvSyncRequest queue
            Case GAME_DROIDDISEMBARK
                recvDroidDisEmbark queue        'droid has disembarked from a Transporter
            Case GAME_GIFT                      ' an alliance gift from one player to another.
                recvGift queue
            Case GAME_LASSAT
                recvLasSat queue
            Case GAME_DEBUG_MODE
                recvProcessDebugMappings queue
            Case GAME_DEBUG_ADD_DROID
                recvDroid queue
            Case GAME_DEBUG_ADD_STRUCTURE
                recvBuildFinished queue
            Case GAME_DEBUG_ADD_FEATURE
                recvMultiPlayerFeature queue
            Case GAME_DEBUG_REMOVE_DROID
                recvDestroyDroid queue
            Case GAME_DEBUG_REMOVE_STRUCTURE
                recvDestroyStructure queue
            Case GAME_DEBUG_REMOVE_FEATURE
                recvDestroyFeature queue
            Case GAME_DEBUG_FINISH_RESEARCH
                recvResearch queue
            Case GAME_SYNC_OPT_CHANGE
                recvSyncOptChange queue
            Case REPLAY_ENDED
                If Not NETisReplay() Then
                    ' ignore
                Else
                    addConsoleMessage "REPLAY HAS ENDED", CENTRE_JUSTIFY, SYSTEM_MESSAGE, False, MAX_CONSOLE_MESSAGE_DURATION
                    addConsoleMessage "(Press ESC to quit.)", CENTRE_JUSTIFY, SYSTEM_MESSAGE, False, MAX_CONSOLE_MESSAGE_DURATION
                End If
            Case Else
                processedMessage1 = False
            End Select
        End If

        ' messages usable all the time
        processedMessage2 = True
        Select Case typ
        Case NET_PING                           ' diagnostic ping msg.
            recvPing queue
        Case NET_PLAYER_DROPPED                 ' remote player got disconnected
            Set r = NETbeginDecode(queue, NET_PLAYER_DROPPED)
            NETuint32_t r, player_id
            NETend r

            If Not ULT(player_id, MAX_CONNECTED_PLAYERS) Then
                debug_ 1584, LOG_INFO, "recvMessage", "** player %u has dropped - huh?", player_id
            ElseIf NETQUEUE_index(queue) <> NetPlay_hostPlayer Then  ' only host should be sending this message
                HandleBadParam "NET_PLAYER_DROPPED given incorrect params.", player_id, NETQUEUE_index(queue)
            Else
                debug_ 1594, LOG_INFO, "recvMessage", "** player %u has dropped!", player_id

                If NetPlay_players_allocated(player_id) And ingame.JoiningInProgress(player_id) Then
                    ' only set ingame.JoiningInProgress[player_id] to false
                    ' when the game starts, it will handle the GAME_PLAYER_LEFT message in their queue properly
                    ingame.JoiningInProgress(player_id) = False
                    ingame.PendingDisconnect(player_id) = True  ' used as a UI indicator that a disconnect will be processed in the future
                End If
                NETsetPlayerConnectionStatus CONNECTIONSTATUS_PLAYER_DROPPED, player_id
            End If
        Case NET_PLAYERRESPONDING               ' remote player is now playing
            Set r = NETbeginDecode(queue, NET_PLAYERRESPONDING)
            ' the player that has just responded
            NETuint32_t r, player_id
            NETend r
            If Not ULT(player_id, MAX_CONNECTED_PLAYERS) Then
                debug_ 1616, LOG_ERROR, "recvMessage", "Bad NET_PLAYERRESPONDING received, ID is %d", player_id
            ElseIf whosResponsible(player_id) <> NETQUEUE_index(queue) And NETQUEUE_index(queue) <> NetPlay_hostPlayer Then
                HandleBadParam "NET_PLAYERRESPONDING given incorrect params.", player_id, NETQUEUE_index(queue)
            Else
                ' This player is now with us!
                If ingame.JoiningInProgress(player_id) Then
                    If game.blindMode = BLIND_MODE_NONE Then
                        addKnownPlayer NetPlay_players_name(player_id), player_id
                    End If
                    ingame.JoiningInProgress(player_id) = False

                    If wz_command_interface_enabled() Then
                        wz_command_interface_output_playerResponding player_id
                    End If
                End If
            End If
        Case GAME_ALLIANCE
            recvAlliance queue, True
        Case NET_VOTE
            If NetPlay_isHost Then
                recvVote queue, False
            End If
        Case NET_VOTE_REQUEST
            If Not NetPlay_isHost Then
                If Not NetPlay_players_isSpectator(selectedPlayer) Then
                    recvVoteRequest queue
                End If
            End If
        Case NET_KICK   ' in-game kick message
            KICK_TYPE = ERROR_NOERROR

            Set r = NETbeginDecode(queue, NET_KICK)
            NETuint32_t r, player_id
            NETstring r, reason, MAX_KICK_REASON
            NETenum r, KICK_TYPE
            NETend r

            If player_id = NetPlay_hostPlayer Then
                buf = Left$(va("Player %d (%s : %s) tried to kick %u", NETQUEUE_index(queue), getPlayerName(NETQUEUE_index(queue), True), NetPlay_players_IPtextAddress(NETQUEUE_index(queue)), player_id), 249)    ' char buf[250]
                NETlogEntry buf, SYNC_FLAG, 0
                debug_ 1683, LOG_ERROR, "recvMessage", "%s", buf
                If NetPlay_isHost Then
                    NETplayerKicked NETQUEUE_index(queue)
                End If
            ElseIf selectedPlayer = player_id Then  ' we've been told to leave.
                debug_ 1692, LOG_INFO, "recvMessage", "You were kicked because %s", reason
                displayKickReasonPopup reason
                setPlayerHasLost True
                ActivityManager_wasKickedByPlayer NETQUEUE_index(queue), KICK_TYPE, reason
            Else
                debug_ 1699, LOG_NET, "recvMessage", "Player %d was kicked: %s", player_id, reason
                NETplayerKicked player_id
            End If
        Case NET_HOST_CONFIG
            If Not recvHostConfig(queue) Then
                ' supplied NET_HOST_CONFIG is not valid
                debug_ 1709, LOG_INFO, "recvMessage", "Bad NET_HOST_CONFIG received"
            End If
        Case GAME_RESEARCHSTATUS
            recvResearchStatus queue
        Case GAME_STRUCTUREINFO
            recvStructureInfo queue
        Case NET_PLAYER_STATS
            recvMultiStats queue
        Case GAME_PLAYER_LEFT
            recvPlayerLeft queue
        Case Else
            processedMessage2 = False
        End Select

        If processedMessage1 And processedMessage2 Then
            debug_ 1733, LOG_ERROR, "recvMessage", "Processed %s message twice!", messageTypeToString(typ)
        End If
        If Not processedMessage1 And Not processedMessage2 Then
            debug_ 1737, LOG_ERROR, "recvMessage", "Didn't handle %s message!", messageTypeToString(typ)
        End If

        NETpop queue
continue_:
    Loop

    recvMessage = True
End Function

Public Sub HandleBadParam(ByVal msg As String, ByVal from As Long, ByVal actual As Long)
    Dim buf As String, KICK_TYPE As LOBBY_ERROR_TYPES

    KICK_TYPE = ERROR_INVALID
    buf = Left$(va("!!>Msg: %s, Actual: %d, Bad: %d", msg, actual, from), 254)    ' char buf[255]
    NETlogEntry buf, SYNC_FLAG, actual
    If NetPlay_isHost Then
        If NETplayerHasConnection(actual) Then
            buf = Left$(va("Auto kicking player %s, invalid command received.", getPlayerName(actual, True)), 254)
            sendInGameSystemMessage buf
        End If
        kickPlayer actual, buf, KICK_TYPE, False
    End If
End Sub

' ////////////////////////////////////////////////////////////////////////////
' Research Stuff. Nat games only send the result of research procedures.
Public Function SendResearch(ByVal player As Long, ByVal index As Long, ByVal trigger As Boolean) As Boolean
    Dim w As MessageWriter
    ' Send the player that is researching the topic and the topic itself
    Set w = NETbeginEncode(NETgameQueue(realSelectedPlayer), GAME_DEBUG_FINISH_RESEARCH)
    NETuint8_t w, player
    NETuint32_t w, index
    NETend w

    SendResearch = True
End Function

' recv a research topic that is now complete.
Private Function recvResearch(ByVal queue As Long) As Boolean
    Dim player As Long, index As Long, i As Long, r As MessageReader

    Set r = NETbeginDecode(queue, GAME_DEBUG_FINISH_RESEARCH)
    NETuint8_t r, player
    NETuint32_t r, index
    NETend r

    If Not debugMappingsAllowed() And bMultiPlayer Then
        debug_ 1793, LOG_WARNING, "recvResearch", "Failed to finish research for player %u.", NetPlay_players_position(NETQUEUE_index(queue))
        recvResearch = False
        Exit Function
    End If

    syncDebug "recvResearch", "player%d, index%u", player, index

    If player >= MAX_PLAYERS Or Not ULT(index, asResearch_size) Then
        debug_ 1801, LOG_ERROR, "recvResearch", "Bad GAME_DEBUG_FINISH_RESEARCH received, player is %d, index is %u", player, index
        recvResearch = False
        Exit Function
    End If

    ' pPlayerRes = &asPlayerResList[player][index]
    syncDebug "recvResearch", "research status = %d", asPlayerResList(player, index).ResearchStatus And RESBITS

    If Not IsResearchCompleted(asPlayerResList(player, index)) Then
        researchResult index, player, False, 0, True
    End If

    ' Update allies research accordingly
    If game.typ = LEVEL_TYPE_SKIRMISH Then
        For i = 0 To MAX_PLAYERS - 1
            If alliances(i, player) = ALLIANCE_FORMED Then
                ' pPlayerRes = &asPlayerResList[i][index]

                If Not IsResearchCompleted(asPlayerResList(i, index)) Then
                    ' Do the research for that player
                    researchResult index, i, False, 0, True
                End If
            End If
        Next
    End If

    recvResearch = True
End Function

' ////////////////////////////////////////////////////////////////////////////
' ////////////////////////////////////////////////////////////////////////////
' New research stuff, so you can see what others are up to!
' inform others that I'm researching this.
Public Function sendResearchStatus(ByVal psBuilding As Long, ByVal index As Long, ByVal player As Long, ByVal bStart As Boolean) As Boolean
    Dim w As MessageWriter, buildingID As Long, zero As Long

    If Not myResponsibility(player) Or ULT(gameTime, 5) Then
        sendResearchStatus = True
        Exit Function
    End If

    Set w = NETbeginEncode(NETgameQueue(realSelectedPlayer), GAME_RESEARCHSTATUS)
    NETuint8_t w, player
    NETbool w, bStart

    ' If we know the building researching it then send its ID
    If psBuilding <> 0 Then
        buildingID = aStruct(psBuilding).id
        NETuint32_t w, buildingID
    Else
        zero = 0
        NETuint32_t w, zero
    End If

    ' Finally the topic in question
    NETuint32_t w, index
    NETend w

    ' Tell UI to remove from the list of available research.
    MakeResearchStartedPending asPlayerResList(player, index)

    sendResearchStatus = True
End Function

Public Function findResearchingFacilityByResearchIndex_lists(pList() As OBJ_LIST, ByVal player As Long, ByVal index As Long) As Long
    Dim it As Long, psBuilding As Long
    ' Go through the structs to find the one doing this topic
    it = pList(player).head
    Do While it <> 0
        psBuilding = lnObj(it)
        If asStructureStats(aStruct(psBuilding).pStructureType).typ = REF_RESEARCH Then
            If aStruct(psBuilding).functionality.researchFacility.psSubject <> -1 Then
                If asResearch(aStruct(psBuilding).functionality.researchFacility.psSubject).ref - STAT_RESEARCH = index Then
                    findResearchingFacilityByResearchIndex_lists = psBuilding
                    Exit Function
                End If
            End If
        End If
        it = lnNext(it)
    Loop
    findResearchingFacilityByResearchIndex_lists = 0  ' Not found.
End Function

Public Function findResearchingFacilityByResearchIndex(ByVal player As Long, ByVal index As Long) As Long
    findResearchingFacilityByResearchIndex = findResearchingFacilityByResearchIndex_lists(apsStructLists, player, index)
End Function

' structure.h's popStatusPending, for a RESEARCH_FACILITY
Private Sub popStatusPending_research(f As RESEARCH_FACILITY)
    If f.pendingCount = 0 Then
        f.pendingCount = f.pendingCount + 1
    End If
    f.pendingCount = f.pendingCount - 1
    If f.pendingCount = 0 Then
        ' Subject is now synchronised, remove pending.
        f.psSubjectPending = -1
        f.statusPending = FACTORY_NOTHING_PENDING
    End If
End Sub

Public Function recvResearchStatus(ByVal queue As Long) As Boolean
    Dim psBuilding As Long, pResearch As Long, player As Long, bStart As Boolean, index As Long, structRef As Long
    Dim r As MessageReader, prevResearchState As Long, psOtherBuilding As Long, ok As Boolean, typeName As String

    bStart = False

    Set r = NETbeginDecode(queue, GAME_RESEARCHSTATUS)
    NETuint8_t r, player
    NETbool r, bStart
    NETuint32_t r, structRef
    NETuint32_t r, index
    NETend r

    syncDebug "recvResearchStatus", "player%d, bStart%d, structRef%u, index%u", player, bStart, structRef, index

    If player >= MAX_PLAYERS Or Not ULT(index, asResearch_size) Then
        debug_ 1912, LOG_ERROR, "recvResearchStatus", "Bad GAME_RESEARCHSTATUS received, player is %d, index is %u", player, index
        recvResearchStatus = False
        Exit Function
    End If
    If Not canGiveOrdersFor(NETQUEUE_index(queue), player) Then
        debug_ 1917, LOG_WARNING, "recvResearchStatus", "Droid order for wrong player."
        syncDebug "recvResearchStatus", "Wrong player."
        recvResearchStatus = False
        Exit Function
    End If

    prevResearchState = 0
    If ULT(selectedPlayer, MAX_PLAYERS) Then
        If aiCheckAlliances(selectedPlayer, player) Then
            prevResearchState = intGetResearchState()
        End If
    End If

    ' pPlayerRes = &asPlayerResList[player][index]

    ' psBuilding may be null if finishing
    If bStart Then                      ' Starting research
        ResetPendingResearchStatus asPlayerResList(player, index)  ' Reset pending state, even if research state is not changed due to the structure being destroyed.

        psBuilding = IdToStruct(structRef, player)

        ' Set that facility to research
        If psBuilding <> 0 Then
            If aStruct(psBuilding).pFunctionality <> 0 Then
                ok = aStruct(psBuilding).pStructureType <> -1
                If ok Then ok = asStructureStats(aStruct(psBuilding).pStructureType).typ = REF_RESEARCH
                If Not ok Then
                    If aStruct(psBuilding).pStructureType <> -1 Then typeName = asStructureStats(aStruct(psBuilding).pStructureType).id Else typeName = ""
                    debug_ 1942, LOG_INFO, "recvResearchStatus", "Structure is not a research facility: ""%s"".", typeName
                    recvResearchStatus = False
                    Exit Function
                End If

                ' psResFacilty = (RESEARCH_FACILITY *) psBuilding->pFunctionality

                popStatusPending_research aStruct(psBuilding).functionality.researchFacility  ' Research is no longer pending, as it's actually starting now.

                If aStruct(psBuilding).functionality.researchFacility.psSubject <> -1 Then
                    cancelResearch psBuilding, ModeImmediate
                End If

                If IsResearchStarted(asPlayerResList(player, index)) Then
                    psOtherBuilding = findResearchingFacilityByResearchIndex(player, index)
                    If Not (psOtherBuilding <> 0) Then ASSERT_FAILURE "psOtherBuilding != nullptr", "src/multiplay.cpp:1958", "recvResearchStatus", "Something researched but no facility."
                    If psOtherBuilding <> 0 Then
                        cancelResearch psOtherBuilding, ModeImmediate
                    End If
                End If

                If Not researchAvailable(index, player, ModeImmediate) And bMultiPlayer Then
                    debug_ 1967, LOG_ERROR, "recvResearchStatus", "Player %d researching impossible topic ""%s"".", player, asResearch(index).name_
                    recvResearchStatus = False
                    Exit Function
                End If

                ' Set the subject up
                pResearch = index           ' &asResearch[index]
                aStruct(psBuilding).functionality.researchFacility.psSubject = pResearch

                ' Start the research
                MakeResearchStarted asPlayerResList(player, index)
                aStruct(psBuilding).functionality.researchFacility.timeStartHold = 0
            End If
        End If
    ' Finished/cancelled research
    Else
        ' If they completed the research, we're done
        If IsResearchCompleted(asPlayerResList(player, index)) Then
            recvResearchStatus = True
            Exit Function
        End If

        ' If they did not say what facility it was, look it up orselves
        If structRef = 0 Then
            psBuilding = findResearchingFacilityByResearchIndex(player, index)
        Else
            psBuilding = IdToStruct(structRef, player)
        End If

        ' Stop the facility doing any research
        If psBuilding <> 0 Then
            ok = aStruct(psBuilding).pStructureType <> -1
            If ok Then ok = asStructureStats(aStruct(psBuilding).pStructureType).typ = REF_RESEARCH
            If Not ok Then
                If aStruct(psBuilding).pStructureType <> -1 Then typeName = asStructureStats(aStruct(psBuilding).pStructureType).id Else typeName = ""
                debug_ 2004, LOG_INFO, "recvResearchStatus", "Structure is not a research facility: ""%s"".", typeName
                recvResearchStatus = False
                Exit Function
            End If

            cancelResearch psBuilding, ModeImmediate
            popStatusPending_research aStruct(psBuilding).functionality.researchFacility  ' Research cancellation is no longer pending, as it's actually cancelling now.
        End If
    End If

    If ULT(selectedPlayer, MAX_PLAYERS) Then
        If aiCheckAlliances(selectedPlayer, player) Then
            intAlliedResearchChanged
            intNotifyResearchButton prevResearchState
        End If
    End If

    recvResearchStatus = True
End Function

' ////////////////////////////////////////////////////////////////////////////
' ////////////////////////////////////////////////////////////////////////////
' Features

' process a destroy feature msg.
Public Function recvDestroyFeature(ByVal queue As Long) As Boolean
    Dim pF As Long, id As Long, r As MessageReader

    Set r = NETbeginDecode(queue, GAME_DEBUG_REMOVE_FEATURE)
    NETuint32_t r, id
    NETend r

    If Not debugMappingsAllowed() And bMultiPlayer Then
        debug_ 2333, LOG_WARNING, "recvDestroyFeature", "Failed to remove feature for player %u.", NetPlay_players_position(NETQUEUE_index(queue))
        recvDestroyFeature = False
        Exit Function
    End If

    pF = IdToFeature(id, ANYPLAYER)
    If pF = 0 Then
        debug_ 2340, LOG_FEATURE, "recvDestroyFeature", "feature id %d not found (probably already destroyed)", id
        recvDestroyFeature = False
        Exit Function
    End If

    debug_ 2344, LOG_FEATURE, "recvDestroyFeature", "p%d feature id %d destroyed (%s)", aFeature(pF).player, aFeature(pF).id, asFeatureStats(aFeature(pF).psStats).name_
    ' Remove the feature locally
    turnOffMultiMsg True
    destroyFeature pF, gameTime - deltaGameTime + 1     ' deltaGameTime is actually 0 here, since we're between updates. However, the value of gameTime - deltaGameTime + 1 will not change when we start the next tick.
    turnOffMultiMsg False

    recvDestroyFeature = True
End Function

Public Function NetworkTextMessage_new(ByVal messageSender As Long, ByVal messageText As String) As NetworkTextMessage
    NetworkTextMessage_new.sender = messageSender
    NetworkTextMessage_new.text = Left$(messageText, MAX_CONSOLE_STRING_LENGTH - 1)   ' sstrcpy
End Function

Public Sub NetworkTextMessage_enqueue(m As NetworkTextMessage, ByVal queue As Long)
    Dim w As MessageWriter
    Set w = NETbeginEncode(queue, NET_TEXTMSG)
    NETint32_t w, m.sender
    NETbool w, m.teamSpecific
    NETstring w, m.text, MAX_CONSOLE_STRING_LENGTH
    NETend w
End Sub

Public Sub printInGameTextMessage(message As NetworkTextMessage)
    Dim formatted As String
    Select Case message.sender
    Case SYSTEM_MESSAGE, NOTIFY_MESSAGE
        addConsoleMessage message.text, DEFAULT_JUSTIFY, message.sender, message.teamSpecific

    Case Else
        formatted = Left$("[" & formatLocalDateTime("%H:%M") & "] " & message.text, MAX_CONSOLE_STRING_LENGTH - 1)
        addConsoleMessage formatted, DEFAULT_JUSTIFY, message.sender, message.teamSpecific
    End Select
End Sub

'
' At this time, we do NOT support messages for beacons
'
Public Function sendBeacon(ByVal locX As Long, ByVal locY As Long, ByVal forPlayer As Long, ByVal sender As Long, ByVal pStr As String) As Boolean
    Dim sendPlayer As Long, w As MessageWriter
    'debug(LOG_WZ, "sendBeacon: '%s'",pStr);

    'find machine that is hosting this human or AI
    sendPlayer = whosResponsible(forPlayer)

    If sendPlayer >= MAX_PLAYERS And sendPlayer <> NetPlay_hostPlayer Then
        debug_ 2152, LOG_ERROR, "sendBeacon", "sendBeacon() - whosResponsible() failed."
        sendBeacon = False
        Exit Function
    End If

    ' I assume this is correct, looks like it sends it to ONLY that person, and the routine
    ' kf_AddHelpBlip() iterates for each player it needs.
    Set w = NETbeginEncode(NETnetQueue(sendPlayer), NET_BEACONMSG)    ' send to the player who is hosting 'to' player (might be himself if human and not AI)
    NETint32_t w, sender                                ' save the actual sender

    ' save the actual player that is to get this msg on the source machine (source can host many AIs)
    NETint32_t w, forPlayer                             ' save the actual receiver (might not be the same as the one we are actually sending to, in case of AIs)
    NETint32_t w, locX                                  ' save location
    NETint32_t w, locY

    NETstring w, pStr, MAX_CONSOLE_STRING_LENGTH        ' copy message in.
    NETend w

    sendBeacon = True
End Function

'prepare viewdata for help blip
Public Function CreateBeaconViewData(ByVal sender As Long, ByVal LocX As Long, ByVal LocY As Long) As Long
    Dim height As Long, psViewData As Long, audioID As Long, name_ As String

    'allocate message space
    psViewData = viewDataNew()

    'store name
    name_ = "Beacon " & CStr(sender)
    aViewData(psViewData).name = name_

    'store text message, hardcoded for now
    StringVector_push_back aViewData(psViewData).textMsg, getPlayerName(sender)

    'store message type
    aViewData(psViewData).typ = VIEW_BEACON

    'allocate memory for blip location etc
    aViewData(psViewData).hasProxData = True

    'store audio
    audioID = NO_SOUND
    aViewData(psViewData).prox_audioID = audioID

    'store blip location
    aViewData(psViewData).prox_x = LocX
    aViewData(psViewData).prox_y = LocY

    'check the z value is at least the height of the terrain
    height = map_Height(LocX, LocY)

    aViewData(psViewData).prox_z = height

    'store prox message type
    aViewData(psViewData).prox_proxType = PROX_ENEMY  'PROX_ENEMY for now

    'remember who sent this msg, so we could remove this one, when same player sends a new help-blip msg
    aViewData(psViewData).prox_sender = sender

    'remember when the message was created so can remove it after some time
    aViewData(psViewData).prox_timeAdded = gameTime
    debug_ 2577, LOG_MSG, "CreateBeaconViewData", "Added message"

    CreateBeaconViewData = psViewData
End Function

' Looks through the players list of messages to find VIEW_BEACON (one per player!) pointer
Public Function findBeaconMsg(ByVal player As Long, ByVal sender As Long) As Long
    Dim it As Long, psCurr As Long
    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/multiplay.cpp:2585", "findBeaconMsg", "Unsupported player: %u", player: findBeaconMsg = 0: Exit Function

    it = apsMessages(player).head
    Do While it <> 0
        psCurr = lnObj(it)
        'look for VIEW_BEACON, should only be 1 per player
        If aMessage(psCurr).dataType = MSG_DATA_BEACON Then
            If aViewData(aMessage(psCurr).pViewData).typ = VIEW_BEACON Then
                debug_ 2594, LOG_WZ, "findBeaconMsg", "findBeaconMsg: %d ALREADY HAS A MESSAGE STORED", player
                If aViewData(aMessage(psCurr).pViewData).prox_sender = sender Then
                    debug_ 2597, LOG_WZ, "findBeaconMsg", "findBeaconMsg: %d ALREADY HAS A MESSAGE STORED from %d", player, sender
                    findBeaconMsg = psCurr
                    Exit Function
                End If
            End If
        End If
        it = lnNext(it)
    Loop

    'not found the message so return NULL
    findBeaconMsg = 0
End Function

' Add a beacon (blip)
Public Function addBeaconBlip(ByVal locX As Long, ByVal locY As Long, ByVal forPlayer As Long, ByVal sender As Long, ByVal textMsg As String) As Boolean
    Dim psMessage As Long, pTempData As Long

    If forPlayer >= MAX_PLAYERS Then
        debug_ 2615, LOG_ERROR, "addBeaconBlip", "addBeaconBlip: player number is too high"
        addBeaconBlip = False
        Exit Function
    End If

    'find the message if was already added previously
    psMessage = findBeaconMsg(forPlayer, sender)
    If psMessage <> 0 Then
        'remove it
        removeMessage psMessage, forPlayer
    End If

    'create new message
    psMessage = addBeaconMessage(MSG_PROXIMITY, False, forPlayer)
    If psMessage <> 0 Then
        pTempData = CreateBeaconViewData(sender, locX, locY)
        If Not (pTempData <> 0) Then ASSERT_FAILURE "pTempData != nullptr", "src/multiplay.cpp:2632", "addBeaconBlip", "Empty help data for radar beacon": addBeaconBlip = False: Exit Function
        aMessage(psMessage).pViewData = pTempData
        debug_ 2634, LOG_MSG, "addBeaconBlip", "blip added, pViewData=%p", aMessage(psMessage).pViewData
        jsDebugMessageUpdate
    Else
        debug_ 2639, LOG_WARNING, "addBeaconBlip", "call failed"
    End If

    'Received a blip message from a player callback
    'store and call later
    '-------------------------------------------------
    'call beacon callback only if not adding for ourselves
    If forPlayer <> sender Then
        triggerEventBeacon sender, forPlayer, textMsg, locX, locY

        If selectedPlayer = forPlayer Then
            ' show console message
            CONPRINTF "Beacon received from %s!", getPlayerName(sender)
            ' play audio
            audio_QueueTrackPos ID_SOUND_BEACON, locX, locY, 0
        End If
    End If

    addBeaconBlip = True
End Function

Public Function sendBeaconToPlayer(ByVal locX As Long, ByVal locY As Long, ByVal forPlayer As Long, ByVal sender As Long, ByVal beaconMsg As String) As Boolean
    Dim retval As Boolean, local_ As Boolean
    ' (||: myResponsibility only when the sender is not the destination)
    local_ = (sender = forPlayer)
    If Not local_ Then local_ = myResponsibility(forPlayer)
    If local_ Then  'if destination player is on this machine
        debug_ 2667, LOG_WZ, "sendBeaconToPlayer", "sending beacon to player %d (local player) from %d", forPlayer, sender
        retval = addBeaconBlip(locX, locY, forPlayer, sender, beaconMsg)
    Else
        debug_ 2672, LOG_WZ, "sendBeaconToPlayer", "sending beacon to player %d (remote player) from %d", forPlayer, sender
        retval = sendBeacon(locX, locY, forPlayer, sender, beaconMsg)
    End If
    jsDebugMessageUpdate
    sendBeaconToPlayer = retval
End Function

Private Function recvBeacon(ByVal queue As Long) As Boolean
    Dim sender As Long, receiver As Long, locX As Long, locY As Long, msg As String, r As MessageReader

    Set r = NETbeginDecode(queue, NET_BEACONMSG)
    NETint32_t r, sender            ' the actual sender
    NETint32_t r, receiver          ' the actual receiver (might not be the same as the one we are actually sending to, in case of AIs)
    NETint32_t r, locX
    NETint32_t r, locY
    NETstring r, msg, MAX_CONSOLE_STRING_LENGTH    ' Receive the actual message
    NETend r

    If Not canGiveOrdersFor(NETQUEUE_index(queue), sender) Then
        debug_ 2694, LOG_WARNING, "recvBeacon", "Beacon (by %d) for wrong player (%d).", NETQUEUE_index(queue), sender
        syncDebug "recvBeacon", "Wrong player."
        recvBeacon = False
        Exit Function
    End If

    debug_ 2699, LOG_WZ, "recvBeacon", "Received beacon for player: %d, from: %d", receiver, sender

    msg = Left$(msg & getPlayerName(sender), MAX_CONSOLE_STRING_LENGTH - 1)   ' name (sstrcat)
    beaconReceiveMsg(sender) = msg                                              ' sstrcpy

    recvBeacon = addBeaconBlip(locX, locY, receiver, sender, beaconReceiveMsg(sender))
End Function

' Reset ready status for all players
Public Sub resetReadyStatus(ByVal bSendOptions As Boolean, Optional ByVal ignoreReadyReset As Boolean = False)
    Dim i As Long
    ' notify all clients if needed
    If bSendOptions Then
        sendOptions
    End If
    netPlayersUpdated = True

    'Really reset ready status
    If NetPlay_isHost And Not ignoreReadyReset Then
        wz_command_interface_output "WZEVENT: readyStatus=RESET" & vbLf

        For i = 0 To MAX_CONNECTED_PLAYERS - 1
            'Ignore for autohost launch option.
            If selectedPlayer = i And getHostLaunch() = HostLaunch_Autohost Then
                GoTo continue_
            End If

            If isHumanPlayer(i) And ingame.JoiningInProgress(i) Then
                changeReadyStatus i, False
            End If
continue_:
        Next

        wz_command_interface_output_room_status_json True
    End If
End Sub

Public Function findPlayerIndexByPosition(ByVal position As Long) As Long
    Dim playerIndex As Long
    For playerIndex = 0 To game.maxPlayers - 1
        If NetPlay_players_position(playerIndex) = position Then
            findPlayerIndexByPosition = playerIndex
            Exit Function
        End If
    Next

    findPlayerIndexByPosition = -1
End Function

Public Function makePlayerSpectator(ByVal playerIndex As Long, ByVal removeAllStructs As Boolean, ByVal quietly As Boolean) As Boolean
    Dim hqStructs() As Long, hqStructs_size As Long, k As Long, it As Long, nxt As Long, psStruct As Long, typ As Long
    Dim savedIntMode As Long, reticuleButtonsToDisable As Variant, lowUISpectatorMode As Boolean, d1 As Long, d2 As Long

    ' Remove objects quietly if the player is starting off as a spectator
    quietly = quietly Or NetPlay_players_isSpectator(playerIndex)

    turnOffMultiMsg True

    If ULT(playerIndex, MAX_PLAYERS) Then
        setPower playerIndex, 0

        ' Destroy HQ
        it = apsStructLists(playerIndex).head
        Do While it <> 0
            psStruct = lnObj(it)
            If REF_HQ = asStructureStats(aStruct(psStruct).pStructureType).typ Then
                ReDim Preserve hqStructs(0 To hqStructs_size)
                hqStructs(hqStructs_size) = psStruct
                hqStructs_size = hqStructs_size + 1
            End If
            it = lnNext(it)
        Loop
        For k = 0 To hqStructs_size - 1
            psStruct = hqStructs(k)
            If quietly Then
                removeStruct psStruct, True
            Else            ' show effects
                destroyStruct psStruct, gameTime
            End If
        Next

        ' Destroy all droids
        debug_ 2843, LOG_DEATH, "makePlayerSpectator", "killing off all droids for player %d", playerIndex
        ' mutating_list_iterate
        it = apsDroidLists(playerIndex).head
        Do While it <> 0
            nxt = lnNext(it)
            If quietly Then         ' don't show effects
                killDroid lnObj(it)
            Else                    ' show effects
                destroyDroid lnObj(it), gameTime
            End If
            it = nxt
        Loop

        ' Destroy structs
        debug_ 2858, LOG_DEATH, "makePlayerSpectator", "killing off structures for player %d", playerIndex
        ' mutating_list_iterate
        it = apsStructLists(playerIndex).head
        Do While it <> 0
            nxt = lnNext(it)
            psStruct = lnObj(it)
            typ = asStructureStats(aStruct(psStruct).pStructureType).typ
            If removeAllStructs _
                Or typ = REF_POWER_GEN _
                Or typ = REF_RESEARCH _
                Or typ = REF_COMMAND_CONTROL _
                Or STRUCTURE_isFactory(psStruct) Then
                ' FIXME: look why destroyStruct() doesn't put back the feature like removeStruct() does
                If quietly Or typ = REF_RESOURCE_EXTRACTOR Then      ' don't show effects
                    removeStruct psStruct, True
                Else            ' show effects
                    destroyStruct psStruct, gameTime
                End If
            End If
            it = nxt
        Loop
    End If

    If Not quietly Then
        debug_ 2883, LOG_INFO, "makePlayerSpectator", "player: %u (gameTime: %u)", playerIndex, gameTime
    End If
    If Not NETisReplay() Or playerIndex <> realSelectedPlayer Then
        syncDebug "makePlayerSpectator", "player%u", playerIndex
    End If
    NetPlay_players_isSpectator(playerIndex) = True ' must come before enableGodMode

    If playerIndex = selectedPlayer Then
        ' reset the widget screen to just the reticule (close all panels)
        savedIntMode = intMode
        intResetScreen False, True
        intMode = savedIntMode ' restore intMode from before intResetScreen call (or it may not be possible to click "continue game" on the mission results screen)

        ' disable various reticule buttons
        reticuleButtonsToDisable = Array(IDRET_MANUFACTURE, IDRET_RESEARCH, IDRET_BUILD, IDRET_DESIGN, IDRET_COMMAND)
        For k = 0 To UBound(reticuleButtonsToDisable)
            If intCheckReticuleButEnabled(reticuleButtonsToDisable(k)) Then
                setReticuleStats reticuleButtonsToDisable(k), "", "", ""
            End If
        Next

        ' hide the power bar
        forceHidePowerBar True

        If Not headlessGameMode() Then
            ' enable "god mode" for map + object visibility (+ minimap)
            enableGodMode
        End If

        ' add spectator mode message
        lowUISpectatorMode = streamer_spectator_mode() Or NETisReplay()
        If Not lowUISpectatorMode Then d1 = MAX_CONSOLE_MESSAGE_DURATION Else d1 = 15
        If Not lowUISpectatorMode Then d2 = 30 Else d2 = 15
        addConsoleMessage "Spectator Mode", CENTRE_JUSTIFY, SYSTEM_MESSAGE, False, d1
        addConsoleMessage "You are a spectator. Enjoy watching the game!", CENTRE_JUSTIFY, SYSTEM_MESSAGE, False, d2

        specLayerInit Not streamer_spectator_mode()
    End If

    turnOffMultiMsg False

    makePlayerSpectator = True
End Function

' ////////////////////////////////////////////////////////////////////////////
' ////////////////////////////////////////////////////////////////////////////
' MultiPlayer main game loop code.
Public Function multiPlayerLoop() As Boolean
    Dim i As Long, joinCount As Long, idx As Long, index As Long, msg As String, now_ As Currency

    joinCount = 0
    For i = 0 To MAX_CONNECTED_PLAYERS - 1
        If isHumanPlayer(i) And ingame.JoiningInProgress(i) Then
            joinCount = joinCount + 1
        End If
    Next
    If joinCount <> 0 Then
        ' deselect anything selected.
        selDroidDeselect selectedPlayer
    Else        'everyone is in the game now!
        If Not ingame.has_TimeEveryoneIsInGame Then
            ingame.has_TimeEveryoneIsInGame = True
            ingame.TimeEveryoneIsInGame = gameTime
            debug_ 556, LOG_NET, "multiPlayerLoop", "I have entered the game @ %u", gameTime
            If Not NetPlay_isHost Then
                debug_ 559, LOG_NET, "multiPlayerLoop", "=== Sending hash to host ==="
                sendDataCheck
            End If
            QueryPerformanceCounter ingame.lastPlayerDataCheck2
            wz_command_interface_output "WZEVENT: allPlayersJoined" & vbLf
            wz_command_interface_output_room_status_json

            ' If in blind *lobby* mode, send data on who the players are
            If game.blindMode <> BLIND_MODE_NONE And game.blindMode < BLIND_MODE_BLIND_GAME Then
                If NetPlay_isHost Then
                    debug_ 571, LOG_INFO, "multiPlayerLoop", "Revealing actual player names and identities to all players"

                    ' Send updated player info (which will include real player names) to all players
                    NETSendAllPlayerInfoTo NET_ALL_PLAYERS

                    ' Send the verified player identity from initial join for each player
                    For idx = 0 To MAX_CONNECTED_PLAYERS - 1
                        sendMultiStatsHostVerifiedIdentities idx
                    Next
                End If
            End If
        End If
        If NetPlay_bComms Then
            sendPing
        End If
        If NetPlay_isHost And NetPlay_bComms Then
            sendDataCheck2
        End If
        ' Only have to do this on a true MP game
        If NetPlay_isHost And Not ingame.isAllPlayersDataOK And NetPlay_bComms Then
            ' (unsigned: UDWORD gameTime - the int32_t value)
            If ULT(GAME_TICKS_PER_SEC * 60, FromU32(U32(gameTime) - U32(ingame.TimeEveryoneIsInGame))) Then
                ' we waited 60 secs to make sure people didn't bypass the data integrity checks
                For index = 0 To MAX_CONNECTED_PLAYERS - 1
                    If ingame.DataIntegrity(index) = False And isHumanPlayer(index) And index <> NetPlay_hostPlayer Then
                        If Not ingame.PendingDisconnect(index) Then
                            msg = Left$(va("Kicking player %s, because they tried to bypass data integrity check!", getPlayerName(index)), 255)
                            sendInGameSystemMessage msg
                            addConsoleMessage msg, LEFT_JUSTIFY, NOTIFY_MESSAGE
                            NETlogEntry msg, SYNC_FLAG, index
                        End If

                        ' (#ifndef DEBUG)
                        kickPlayer index, "Invalid data!", ERROR_INVALID, False
                        debug_ 616, LOG_WARNING, "multiPlayerLoop", "Kicking Player %s (%u), they tried to bypass data integrity check!", getPlayerName(index), index
                    End If
                Next
                ingame.isAllPlayersDataOK = True
            End If
        End If
    End If

    If NetPlay_isHost Then
        QueryPerformanceCounter now_
        autoLagKickRoutine now_
        autoDesyncKickRoutine now_
        processPendingKickVotes
    End If

    ' if player has won then process the win effects...
    If testPlayerHasWon() Then
        multiplayerWinSequence False
    End If
    multiPlayerLoop = True
End Function

' (VB) std::chrono::duration_cast<std::chrono::milliseconds>(a - b).count()
Private Function steadyMsBetween(ByVal a As Currency, ByVal b As Currency) As Double
    If perfFreqJ = 0 Then QueryPerformanceFrequency_J perfFreqJ
    steadyMsBetween = Fix(CDbl(a - b) * 1000# / CDbl(perfFreqJ))
End Function

' (VB) std::chrono::duration_cast<std::chrono::seconds>(a - b).count()
Private Function steadySecondsBetween(ByVal a As Currency, ByVal b As Currency) As Double
    If perfFreqJ = 0 Then QueryPerformanceFrequency_J perfFreqJ
    steadySecondsBetween = Fix(CDbl(a - b) / CDbl(perfFreqJ))
End Function

' (VB) base64Encode(identity.toBytes(EcKey::Public))
Private Function publicKeyB64(identity As EcKey) As String
    Dim b() As Byte, n As Long
    n = EcKey_toBytes(identity, EcKey_Public, b)
    publicKeyB64 = base64Encode(b, n)
End Function

' ////////////////////////////////////////////////////////////////////////////
' Auto Bad Connection Kick Handling

Public Sub autoLagKickRoutine(ByVal now_ As Currency)
    Dim LagAutoKickSeconds As Long, LagKickAggressiveness As Long, isLobby As Boolean, isInitialLoad As Boolean
    Dim numPlayersLoaded As Long, totalNumPlayers As Long, playerCheckLimit As Long, i As Long, isLagging As Boolean
    Dim isWaitingForInitialLoad As Boolean, waitingForLoadTime As Double, loadGracePeriod As Long, LagSecondsCount As Long
    Dim msg As String, identity As EcKey
    If Not bMultiPlayer Or Not NetPlay_bComms Or Not NetPlay_isHost Then
        Exit Sub
    End If

    LagAutoKickSeconds = war_getAutoLagKickSeconds()
    If LagAutoKickSeconds <= 0 Then
        Exit Sub
    End If
    LagKickAggressiveness = war_getAutoLagKickAggressiveness()
    If LagKickAggressiveness < 1 Then LagKickAggressiveness = 1

    If steadyMsBetween(now_, ingame.lastLagCheck) < LAG_CHECK_INTERVAL Then
        Exit Sub
    End If

    isLobby = ingame.localJoiningInProgress
    isInitialLoad = Not isLobby And Not ingame.has_TimeEveryoneIsInGame
    numPlayersLoaded = 0
    totalNumPlayers = 0

    ingame.lastLagCheck = now_
    If isLobby Or isInitialLoad Then playerCheckLimit = MAX_CONNECTED_PLAYERS_J Else playerCheckLimit = MAX_PLAYERS
    For i = 0 To playerCheckLimit - 1
        If Not isHumanPlayer(i) Or Not NetPlay_players_heartbeat(i) Then
            GoTo continue_
        End If
        If i = NetPlay_hostPlayer Then
            GoTo continue_
        End If
        If i > MAX_PLAYERS And Not gtimeShouldWaitForPlayer(i) Then
            GoTo continue_
        End If

        If i < MAX_PLAYERS Then
            totalNumPlayers = totalNumPlayers + 1
            If Not ingame.JoiningInProgress(i) Then
                numPlayersLoaded = numPlayersLoaded + 1
            End If
        End If

        isLagging = Not ULT(ingame.PingTimes(i), PING_LIMIT)
        isWaitingForInitialLoad = isInitialLoad And ingame.JoiningInProgress(i)
        If isWaitingForInitialLoad Then
            waitingForLoadTime = steadySecondsBetween(now_, ingame.startTime)
            loadGracePeriod = LAG_INITIAL_LOAD_GRACEPERIOD
            If i > MAX_PLAYERS Then
                ' special handling for spectator slots:
                ' if all actual players are loaded
                ' - reduce the grace period for spectators to load
                ' - reduce the applicable auto lag kick time
                If totalNumPlayers > 0 And numPlayersLoaded = totalNumPlayers And i <> NetPlay_hostPlayer Then
                    loadGracePeriod = 0
                    If LagAutoKickSeconds > 10 Then LagAutoKickSeconds = 10 ' (fine to set this here because any i after this will all be spectators)
                End If
            End If
            If waitingForLoadTime < loadGracePeriod Then
                ' within grace period for initial load (some machines may take longer to load into the match)
                isLagging = False
                isWaitingForInitialLoad = False
            Else
                ' exceeded the grace period for initial load, and still waiting on this player to join
                ' treat them as lagging below
            End If
        End If

        If Not isLagging And Not isWaitingForInitialLoad Then
            If ingame.LagCounter(i) > 0 Then
                ingame.LagCounter(i) = ingame.LagCounter(i) - 1
            End If
            GoTo continue_
        End If

        If ingame.PendingDisconnect(i) Then
            ' player already technically left, but we're still in the "pre-game" phase so the GAME_PLAYER_LEFT hasn't been processed yet
            GoTo continue_
        End If

        ingame.LagCounter(i) = ingame.LagCounter(i) + LagKickAggressiveness
        LagSecondsCount = ingame.LagCounter(i) \ LagKickAggressiveness
        If LagSecondsCount >= LagAutoKickSeconds Then
            msg = va("Auto-kicking player %u (""%s"") because of ping issues. (Timeout: %u seconds)", i, getPlayerName(i, True), LagAutoKickSeconds)
            debug_ 224, LOG_INFO, "autoLagKickRoutine", "%s", msg
            sendInGameSystemMessage msg
            If wz_command_interface_enabled() Then
                identity = getOutputPlayerIdentity(i)
                wz_command_interface_output "WZEVENT: lag-kick: %u %s %s" & vbLf, i, NetPlay_players_IPtextAddress(i), publicKeyB64(identity)
            End If
            kickPlayer i, "Your connection was too laggy.", ERROR_CONNECTION, False
            ingame.LagCounter(i) = 0
        ElseIf LagSecondsCount = (LagAutoKickSeconds - 3) Then
            msg = va("Auto-kicking player %u (""%s"") in %u seconds. (lag)", i, getPlayerName(i, True), (LagAutoKickSeconds - LagSecondsCount))
            debug_ 236, LOG_INFO, "autoLagKickRoutine", "%s", msg
            sendInGameSystemMessage msg
        ElseIf LagSecondsCount Mod 15 = 0 Then ' every 15 seconds
            msg = va("Auto-kicking player %u (""%s"") in %u seconds. (lag)", i, getPlayerName(i, True), (LagAutoKickSeconds - LagSecondsCount))
            debug_ 241, LOG_INFO, "autoLagKickRoutine", "%s", msg
            sendInGameSystemMessage msg
        End If
continue_:
    Next
End Sub

Public Sub autoDesyncKickRoutine(ByVal now_ As Currency)
    Dim DesyncAutoKickSeconds As Long, playerCheckLimit As Long, i As Long, isDesynced As Boolean, msg As String
    Dim identity As EcKey
    If Not bMultiPlayer Or Not NetPlay_bComms Or Not NetPlay_isHost Then
        Exit Sub
    End If

    DesyncAutoKickSeconds = war_getAutoDesyncKickSeconds()
    If DesyncAutoKickSeconds <= 0 Then
        Exit Sub
    End If

    If steadyMsBetween(now_, ingame.lastDesyncCheck) < DESYNC_CHECK_INTERVAL Then
        Exit Sub
    End If

    If ingame.has_endTime Then
        ' game ended - skip desync check / kick
        Exit Sub
    End If

    ingame.lastDesyncCheck = now_
    playerCheckLimit = MAX_PLAYERS
    For i = 0 To playerCheckLimit - 1
        If Not isHumanPlayer(i) Or Not NetPlay_players_heartbeat(i) Then
            GoTo continue_
        End If
        If i = NetPlay_hostPlayer Then
            GoTo continue_
        End If
        If i > MAX_PLAYERS And Not gtimeShouldWaitForPlayer(i) Then
            GoTo continue_
        End If

        isDesynced = NETcheckPlayerConnectionStatus(CONNECTIONSTATUS_DESYNC, i)
        If Not isDesynced Then
            ingame.DesyncCounter(i) = 0
            GoTo continue_
        End If

        If ingame.PendingDisconnect(i) Then
            ' player already technically left, but we're still in the "pre-game" phase so the GAME_PLAYER_LEFT hasn't been processed yet
            GoTo continue_
        End If

        ingame.DesyncCounter(i) = ingame.DesyncCounter(i) + 1

        If ingame.DesyncCounter(i) >= DesyncAutoKickSeconds Then
            msg = va("Auto-kicking player %u (""%s"") because of desync. (Timeout: %u seconds)", i, getPlayerName(i, True), DesyncAutoKickSeconds)
            debug_ 308, LOG_INFO, "autoDesyncKickRoutine", "%s", msg
            sendInGameSystemMessage msg
            If wz_command_interface_enabled() Then
                identity = getOutputPlayerIdentity(i)
                wz_command_interface_output "WZEVENT: desync-kick: %u %s %s" & vbLf, i, NetPlay_players_IPtextAddress(i), publicKeyB64(identity)
            End If
            kickPlayer i, "Your game simulation deviated too far from the host - desync.", ERROR_CONNECTION, False
            ingame.DesyncCounter(i) = 0
        ElseIf ingame.DesyncCounter(i) >= (DesyncAutoKickSeconds - 3) Then
            msg = va("Auto-kicking player %u (""%s"") in %u seconds. (desync)", i, getPlayerName(i, True), (DesyncAutoKickSeconds - ingame.DesyncCounter(i)))
            debug_ 320, LOG_INFO, "autoDesyncKickRoutine", "%s", msg
            sendInGameSystemMessage msg
        ElseIf ingame.DesyncCounter(i) Mod 2 = 0 Then ' every 2 seconds
            msg = va("Auto-kicking player %u (""%s"") in %u seconds. (desync)", i, getPlayerName(i, True), (DesyncAutoKickSeconds - ingame.DesyncCounter(i)))
            debug_ 325, LOG_INFO, "autoDesyncKickRoutine", "%s", msg
            sendInGameSystemMessage msg
        End If
continue_:
    Next
End Sub

' ////////////////////////////////////////////////////////////////////////////

' uint64_t calculateSecondsNotReadyForPlayer(size_t i, std::chrono::steady_clock::time_point now) (a Double)
Public Function calculateSecondsNotReadyForPlayer(ByVal i As Long, ByVal now_ As Currency) As Double
    Dim totalSecondsNotReady As Double
    If i >= NetPlay_players_size() Then calculateSecondsNotReadyForPlayer = 0: Exit Function
    totalSecondsNotReady = ingame.secondsNotReady(i)
    If ingame.has_lastNotReadyTimes(i) Then
        ' accumulate time since last not ready
        totalSecondsNotReady = totalSecondsNotReady + steadySecondsBetween(now_, ingame.lastNotReadyTimes(i))
    End If
    calculateSecondsNotReadyForPlayer = totalSecondsNotReady
End Function

Public Sub autoLobbyNotReadyKickRoutine(ByVal now_ As Currency)
    Dim isLobby As Boolean, NotReadyAutoKickSeconds As Long, queuedToKickIdx(0 To MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS - 1) As Boolean
    Dim i As Long, totalSecondsNotReady As Double, msg As String, identity As EcKey
    If Not bMultiPlayer Or Not NetPlay_bComms Or Not NetPlay_isHost Then
        Exit Sub
    End If

    isLobby = ingame.localJoiningInProgress
    If Not isLobby Then
        Exit Sub
    End If

    If Not multiplayPlayersShouldCheckReady() Then
        ' If players can't check ready, skip
        Exit Sub
    End If

    NotReadyAutoKickSeconds = war_getAutoNotReadyKickSeconds()
    If NotReadyAutoKickSeconds <= 0 Then
        Exit Sub
    End If

    If steadyMsBetween(now_, ingame.lastNotReadyCheck) < NOTREADY_CHECK_INTERVAL Then
        Exit Sub
    End If

    ingame.lastNotReadyCheck = now_

    For i = 0 To MAX_CONNECTED_PLAYERS_J - 1
        If Not isHumanPlayer(i) Then
            GoTo continue_
        End If
        If i = NetPlay_hostPlayer Then
            GoTo continue_
        End If
        If i > MAX_PLAYERS And Not gtimeShouldWaitForPlayer(i) Then
            GoTo continue_
        End If
        If ingame.PendingDisconnect(i) Then
            ' player already technically left, but we're still in the "pre-game" phase so the GAME_PLAYER_LEFT hasn't been processed yet
            GoTo continue_
        End If

        totalSecondsNotReady = calculateSecondsNotReadyForPlayer(i, now_)
        If totalSecondsNotReady >= NotReadyAutoKickSeconds Then
            ' record that player should be kicked, but wait to kick until after processing all slots
            queuedToKickIdx(i) = True
        ElseIf Not NetPlay_players_ready(i) And totalSecondsNotReady >= (NotReadyAutoKickSeconds - 8) Then
            sendQuickChat_data WzQuickChatMessage_INTERNAL_LOCALIZED_LOBBY_NOTICE, realSelectedPlayer, WzQuickChatTargeting_targetAll(), _
                WzQuickChatDataContexts_INTERNAL_LOCALIZED_LOBBY_NOTICE_constructMessageData(WzQuickChatDataContexts_INTERNAL_LOCALIZED_LOBBY_NOTICE_Context_NotReadyKickWarning, i, FromU32(NotReadyAutoKickSeconds - totalSecondsNotReady))
        End If
continue_:
    Next

    ' actually process pending kicks
    For i = 0 To MAX_CONNECTED_PLAYERS_J - 1
        If queuedToKickIdx(i) Then
            msg = va("Auto-kicking player %u (""%s"") because they aren't ready. (Timeout: %u seconds)", i, getPlayerName(i), NotReadyAutoKickSeconds)
            debug_ 418, LOG_INFO, "autoLobbyNotReadyKickRoutine", "%s", msg
            sendQuickChat_data WzQuickChatMessage_INTERNAL_LOCALIZED_LOBBY_NOTICE, realSelectedPlayer, WzQuickChatTargeting_targetAll(), _
                WzQuickChatDataContexts_INTERNAL_LOCALIZED_LOBBY_NOTICE_constructMessageData(WzQuickChatDataContexts_INTERNAL_LOCALIZED_LOBBY_NOTICE_Context_NotReadyKicked, i, NotReadyAutoKickSeconds)
            If wz_command_interface_enabled() Then
                identity = getOutputPlayerIdentity(i)
                wz_command_interface_output "WZEVENT: notready-kick: %u %s %s" & vbLf, i, NetPlay_players_IPtextAddress(i), publicKeyB64(identity)
            End If
            kickPlayer i, "You have been removed from the room." & vbLf & "You have spent too much time without checking Ready." & vbLf & vbLf & _
                "In the future, please check Ready and leave it checked, to avoid delaying games for other players.", ERROR_CONNECTION, False
        End If
    Next
End Sub

' ////////////////////////////////////////////////////////////////////////////
' throw a party when you win!
Public Function multiplayerWinSequence(ByVal firstCall As Boolean) As Boolean
    Dim rotAmount As Single, it As Long, psStruct As Long, pos2 As Vector3i
    If selectedPlayer >= MAX_PLAYERS Then
        multiplayerWinSequence = False
        Exit Function
    End If

    If firstCall Then
        winSeqPos = cameraToHome(selectedPlayer, True, False) ' pan the camera to home if not already doing so
        winSeqLast = 0

        ' stop all research
        CancelAllResearch selectedPlayer

        ' stop all manufacture.
        it = apsStructLists(selectedPlayer).head
        Do While it <> 0
            psStruct = lnObj(it)
            If psStruct <> 0 And STRUCTURE_isFactory(psStruct) Then
                If aStruct(psStruct).functionality.factory.psSubject <> 0 Then 'check if active
                    cancelProduction psStruct, ModeQueue
                End If
            End If
            it = lnNext(it)
        Loop
    End If

    ' rotate world
    If MissionResUp And Not getWarCamStatus() Then
        rotAmount = graphicsTimeAdjustedIncrement(MAP_SPIN_RATE / 12)
        playerPos.r.y = CLng(Fix(CSng(playerPos.r.y) + rotAmount))
    End If

    If ULT(gameTime, winSeqLast) Then
        winSeqLast = 0
    End If
    If ULT(gameTime - winSeqLast, 500) Then         ' only  if not done recently.
        multiplayerWinSequence = True
        Exit Function
    End If
    winSeqLast = gameTime

    If rand() Mod 3 = 0 Then
        pos2 = winSeqPos
        pos2.x = pos2.x + (rand() Mod world_coord(8)) - world_coord(4)
        pos2.z = pos2.z + (rand() Mod world_coord(8)) - world_coord(4)

        If pos2.x < 0 Then
            pos2.x = 128
        End If

        If ULT(world_coord(mapWidth), pos2.x) Then
            pos2.x = world_coord(mapWidth)
        End If

        If pos2.z < 0 Then
            pos2.z = 128
        End If

        If ULT(world_coord(mapHeight), pos2.z) Then
            pos2.z = world_coord(mapHeight)
        End If

        addEffect pos2, EFFECT_FIREWORK, FIREWORK_TYPE_LAUNCHER, False, 0, 0 ' throw up some fire works.
    End If

    ' show the score..

    multiplayerWinSequence = True
End Function

Public Function setPlayerName(ByVal player As Long, ByVal sName As String) As Boolean
    If Not (player < MAX_CONNECTED_PLAYERS_J And player >= 0) Then ASSERT_FAILURE "player < MAX_CONNECTED_PLAYERS && player >= 0", "src/multiplay.cpp:892", "setPlayerName", "Player index (%u) out of range", player: setPlayerName = False: Exit Function
    NetPlay_players_name(player) = Left$(sName, 64 - 1)    ' sstrcpy(NetPlay.players[player].name, sName) (char name[StringSize])
    setPlayerName = True
End Function

Public Sub setPlayerMuted(ByVal playerIdx As Long, ByVal muted As Boolean)
    Dim trueId As TrueIdentity
    If Not ULT(playerIdx, MAX_CONNECTED_PLAYERS_J) Then ASSERT_FAILURE "playerIdx < MAX_CONNECTED_PLAYERS", "src/multiplay.cpp:2024", "setPlayerMuted", "Invalid playerIdx: %u", playerIdx: Exit Sub
    If muted = ingame.muteChat(playerIdx) Then
        ' no change
        Exit Sub
    End If
    ingame.muteChat(playerIdx) = muted
    If isHumanPlayer(playerIdx) Then
        trueId = getTruePlayerIdentity(playerIdx)
        If Not EcKey_empty(trueId.identity) And (NetPlay_isHost Or Not isBlindPlayerInfoState()) Then
            storePlayerMuteOption NetPlay_players_name(playerIdx), trueId.identity, muted
        End If
    End If
End Sub

Public Function isPlayerMuted(ByVal sender As Long) As Boolean
    If Not ULT(sender, MAX_CONNECTED_PLAYERS_J) Then ASSERT_FAILURE "sender < MAX_CONNECTED_PLAYERS", "src/multiplay.cpp:2044", "isPlayerMuted", "Invalid sender: %u", sender: Exit Function
    isPlayerMuted = ingame.muteChat(sender) Or Not ingame.hostChatPermissions(sender)
End Function

Public Sub printConsoleNameChange(ByVal oldName As String, ByVal newName As String)
    If game.blindMode <> BLIND_MODE_NONE Then
        Exit Sub
    End If
    ' ssprintf(msg, "%s \u2192 %s", oldName, newName) into char msg[MAX_CONSOLE_STRING_LENGTH] (the names fit)
    displayRoomSystemMessage oldName & " " & ChrW$(&H2192) & " " & newName
End Sub

Public Function getPlayerColourName(ByVal player As Long) As String
    ' static const char *playerColors[] (N_(): no translations here)
    ' STATIC_ASSERT(MAX_PLAYERS <= ARRAY_SIZE(playerColors));
    If Not ULT(player, 16) Then ASSERT_FAILURE "player < ARRAY_SIZE(playerColors)", "src/multiplay.cpp:2730", "getPlayerColourName", "player number (%d) exceeds maximum (%lu)", player, 16

    If Not ULT(player, 16) Then
        getPlayerColourName = ""
        Exit Function
    End If

    Select Case getPlayerColour(player)
    Case 0: getPlayerColourName = "Green"
    Case 1: getPlayerColourName = "Orange"
    Case 2: getPlayerColourName = "Grey"
    Case 3: getPlayerColourName = "Black"
    Case 4: getPlayerColourName = "Red"
    Case 5: getPlayerColourName = "Blue"
    Case 6: getPlayerColourName = "Pink"
    Case 7: getPlayerColourName = "Cyan"
    Case 8: getPlayerColourName = "Yellow"
    Case 9: getPlayerColourName = "Purple"
    Case 10: getPlayerColourName = "White"
    Case 11: getPlayerColourName = "Bright blue"
    Case 12: getPlayerColourName = "Neon green"
    Case 13: getPlayerColourName = "Infrared"
    Case 14: getPlayerColourName = "Ultraviolet"
    Case 15: getPlayerColourName = "Brown"
    End Select
End Function

Public Function shouldSkipReadyResetOnPlayerJoinLeaveEvent() As Boolean
    ' If a player joins or leaves, do not reset existing player "ready" status if:
    ' 1. It's a blind simple lobby (i.e. waiting room)
    ' 2. min_autostart_player_count is set (ex. via the --startplayers= command line option) *and* locked.readybeforefull is false
    shouldSkipReadyResetOnPlayerJoinLeaveEvent = isBlindSimpleLobby(game.blindMode) Or ((min_autostart_player_count() > 0) And Not getLockedOptions().readybeforefull)
End Function

' (VB) static inline std::chrono::seconds maxDataCheck2WaitSeconds() (seconds)
Private Function maxDataCheck2WaitSeconds() As Long
    maxDataCheck2WaitSeconds = war_getAutoLagKickSeconds() + 3
    If maxDataCheck2WaitSeconds < 60 Then maxDataCheck2WaitSeconds = 60
End Function

Public Function sendDataCheck2() As Boolean
    Dim now_ As Currency, maxWaitSeconds As Long, player As Long, lim As Long, msg As String, w As MessageWriter
    Dim zOrders() As Long, nZ As Long, layerZ() As Long, layerCount() As Long, nLayers As Long, i As Long, j As Long
    Dim layersSize As Long, zOrder As Long, cnt As Long, aiIndex As Long, bValue As Boolean, t As Long
    If NetPlay_isHost Then
        QueryPerformanceCounter_J now_
        If steadyMsBetween(now_, ingame.lastPlayerDataCheck2) < DATACHECK2_INTERVAL_MS Then
            sendDataCheck2 = True
            Exit Function
        End If

        ' Send a request to all active players
        maxWaitSeconds = maxDataCheck2WaitSeconds()
        lim = game.maxPlayers
        If lim > MAX_PLAYERS Then lim = MAX_PLAYERS
        For player = 0 To lim - 1
            If player = NetPlay_hostPlayer Or Not isHumanPlayer(player) Or NetPlay_players_isSpectator(player) Or Not NetPlay_players_heartbeat(player) Then
                GoTo continue_
            End If

            ' Check when the last unanswered request was sent
            If ingame.has_lastSentPlayerDataCheck2(player) Then
                If steadySecondsBetween(now_, ingame.lastSentPlayerDataCheck2(player)) >= maxWaitSeconds Then
                    ' If it's after the allowed time, kick the player
                    If Not ingame.PendingDisconnect(player) Then
                        msg = va("%s (%u) has an incompatible mod, and has been kicked.", getPlayerName(player), player)
                        sendInGameSystemMessage msg
                        addConsoleMessage msg, LEFT_JUSTIFY, NOTIFY_MESSAGE
                    End If
                    kickPlayer player, "Your data doesn't match the host's!", ERROR_WRONGDATA, False
                    debug_ 1111, LOG_INFO, "sendDataCheck2", "%s (%u) did not respond with a NET_DATA_CHECK2 within the required timeframe (%s seconds), and has been kicked", getPlayerName(player), player, CStr(maxWaitSeconds)
                    ingame.has_lastSentPlayerDataCheck2(player) = False
                    GoTo continue_
                End If
            End If

            Set w = NETbeginEncode(NETnetQueue(player), NET_DATA_CHECK2)
            NETuint32_t w, NetPlay_hostPlayer
            NETend w

            If Not ingame.has_lastSentPlayerDataCheck2(player) Then
                ingame.lastSentPlayerDataCheck2(player) = now_
                ingame.has_lastSentPlayerDataCheck2(player) = True
            End If
continue_:
        Next
        ingame.lastPlayerDataCheck2 = now_
        sendDataCheck2 = True
        Exit Function
    End If

    ' For a player, respond to the host
    Set w = NETbeginEncode(NETnetQueue(NetPlay_hostPlayer), NET_DATA_CHECK2)     ' only need to send to HOST
    NETuint32_t w, selectedPlayer
    NETuint32_t w, realSelectedPlayer
    ' std::unordered_map<uint16_t, uint32_t> layers (widgForEachOverlayScreen counts each zOrder). (VB) the map's
    ' iteration order (MSVC's buckets) is not kept: ascending zOrder (the host only looks the counts up)
    nZ = widgGetOverlayScreenZOrders(zOrders)
    If nZ > 0 Then
        ReDim layerZ(0 To nZ - 1): ReDim layerCount(0 To nZ - 1)
    End If
    For i = 0 To nZ - 1
        For j = 0 To nLayers - 1
            If layerZ(j) = zOrders(i) Then Exit For
        Next
        If j = nLayers Then
            layerZ(nLayers) = zOrders(i)
            nLayers = nLayers + 1
        End If
        layerCount(j) = layerCount(j) + 1
    Next
    For i = 0 To nLayers - 2                    ' ascending zOrder
        For j = i + 1 To nLayers - 1
            If layerZ(j) < layerZ(i) Then
                t = layerZ(i): layerZ(i) = layerZ(j): layerZ(j) = t
                t = layerCount(i): layerCount(i) = layerCount(j): layerCount(j) = t
            End If
        Next
    Next
    layersSize = nLayers
    NETuint32_t w, layersSize
    For i = 0 To nLayers - 1
        zOrder = layerZ(i)
        NETuint16_t w, zOrder
        cnt = layerCount(i)
        NETuint32_t w, cnt
    Next
    For i = 0 To DATA_MAXDATA - 1
        NETuint32_t w, DataHash(i)
    Next
    aiIndex = NetPlay_players_ai(realSelectedPlayer)
    NETint8_t w, aiIndex
    bValue = godMode
    NETbool w, bValue
    NETend w
    sendDataCheck2 = True
End Function

Public Function recvDataCheck2(ByVal queue As Long) As Boolean
    Dim player As Long, recvSelectedPlayer As Long, recvRealSelectedPlayer As Long, tempBuffer(0 To DATA_MAXDATA - 1) As Long
    Dim aiIndex As Long, recvGM As Boolean, r As MessageReader, layersSize As Long, zOrder As Long, layerCount As Long, i As Long
    Dim layerZ() As Long, layerN() As Long, nLayers As Long, k As Long, hasWrongData As Boolean, zCheck As Long, foundIt As Long
    Dim msg As String, cmpEqual As Boolean
    player = NETQUEUE_index(queue)
    recvSelectedPlayer = 0
    recvRealSelectedPlayer = 0
    aiIndex = 0

    If Not NetPlay_isHost Then ' the host can send NET_DATA_CHECK2 messages to clients to request a check
        If Not (NetPlay_hostPlayer = NETQUEUE_index(queue)) Then ASSERT_FAILURE "NetPlay.hostPlayer == queue.index", "src/multiplay.cpp:1169", "recvDataCheck2", "Non-host player (%u) is sending NET_DATA_CHECK2 to us??", NETQUEUE_index(queue): recvDataCheck2 = False: Exit Function
        Set r = NETbeginDecode(queue, NET_DATA_CHECK2)
        NETuint32_t r, recvSelectedPlayer
        NETend r
        If Not (NetPlay_hostPlayer = recvSelectedPlayer) Then ASSERT_FAILURE "NetPlay.hostPlayer == recvSelectedPlayer", "src/multiplay.cpp:1173", "recvDataCheck2", "Non-host player (selectedPlayer: %u) is sending NET_DATA_CHECK2 to us??", recvSelectedPlayer: recvDataCheck2 = False: Exit Function
        sendDataCheck2
        recvDataCheck2 = True
        Exit Function
    End If

    Set r = NETbeginDecode(queue, NET_DATA_CHECK2)
    NETuint32_t r, recvSelectedPlayer
    NETuint32_t r, recvRealSelectedPlayer
    layersSize = 0
    zOrder = 0
    layerCount = 0
    NETuint32_t r, layersSize
    i = 0
    Do While ULT(i, layersSize)
        NETuint16_t r, zOrder
        NETuint32_t r, layerCount
        ' layers[zOrder] = layerCount;
        For k = 0 To nLayers - 1
            If layerZ(k) = zOrder Then Exit For
        Next
        If k = nLayers Then
            ReDim Preserve layerZ(0 To nLayers): ReDim Preserve layerN(0 To nLayers)
            layerZ(nLayers) = zOrder
            nLayers = nLayers + 1
        End If
        layerN(k) = layerCount
        i = i + 1
    Loop
    For i = 0 To DATA_MAXDATA - 1
        NETuint32_t r, tempBuffer(i)
    Next
    NETint8_t r, aiIndex
    NETbool r, recvGM
    NETend r

    If Not ULT(player, MAX_CONNECTED_PLAYERS_J) Then ' invalid player number.
        debug_ 1201, LOG_ERROR, "recvDataCheck2", "invalid player number (%u) detected.", player
        recvDataCheck2 = False
        Exit Function
    End If

    If whosResponsible(player) <> NETQUEUE_index(queue) Then
        HandleBadParam "NET_DATA_CHECK2 given incorrect params.", player, NETQUEUE_index(queue)
        recvDataCheck2 = False
        Exit Function
    End If

    If Not ULT(recvRealSelectedPlayer, NetPlay_players_size()) Then
        HandleBadParam "NET_DATA_CHECK2 given invalid param.", recvRealSelectedPlayer, NETQUEUE_index(queue)
        recvDataCheck2 = False
        Exit Function
    End If

    If Not isHumanPlayer(player) Or Not NetPlay_players_heartbeat(player) Then
        ' Ignoring
        recvDataCheck2 = False
        Exit Function
    End If

    debug_ 1223, LOG_NET, "recvDataCheck2", "** Received NET_DATA_CHECK2 from player %u", player
    ingame.has_lastSentPlayerDataCheck2(player) = False

    hasWrongData = False

    If Not NetPlay_players_isSpectator(player) And (recvSelectedPlayer <> player Or recvRealSelectedPlayer <> player) Then
        debug_ 1230, LOG_INFO, "recvDataCheck2", "%s (%u) has a corrupted player index. (selectedPlayer: %u, realSelectedPlayer: %u)", getPlayerName(player), player, recvSelectedPlayer, recvRealSelectedPlayer
        hasWrongData = True
    End If

    If ULT(1024, layersSize) Then
        debug_ 1236, LOG_INFO, "recvDataCheck2", "%s (%u) has a very high layersSize - something is probably wrong. (layersSize: %u)", getPlayerName(player), player, layersSize
        hasWrongData = True
    End If
    If Not NetPlay_players_isSpectator(player) Then
        ' for (uint16_t zCheck = 65530; zOrder < std::numeric_limits<uint16_t>::max() - 2; ++zOrder): (the C++ steps zOrder)
        zCheck = 65530
        Do While zOrder < 65535 - 2
            foundIt = layerIndex(layerZ, nLayers, zCheck)
            If foundIt >= 0 Then
                debug_ 1247, LOG_INFO, "recvDataCheck2", "%s (%u) has an unexpected display layer. (layer: %u, count: %u)", getPlayerName(player), player, zCheck, layerN(foundIt)
            End If
            zOrder = (zOrder + 1) And &HFFFF&
        Loop
        zCheck = 65535 - 2
        If layerIndex(layerZ, nLayers, zCheck) >= 0 And Not DebugInputManager_debugMappingsAllowed() Then
            debug_ 1253, LOG_INFO, "recvDataCheck2", "%s (%u) has an unexpected display layer (script debugger).", getPlayerName(player), player
            hasWrongData = True
        End If
        zCheck = 65535
        foundIt = layerIndex(layerZ, nLayers, zCheck)
        If foundIt >= 0 Then
            If ULT(1, layerN(foundIt)) Then
                debug_ 1260, LOG_INFO, "recvDataCheck2", "%s (%u) has an unexpected number of notification layers. (count: %u)", getPlayerName(player), player, layerN(foundIt)
            End If
        End If
    End If

    cmpEqual = True
    For i = 0 To DATA_MAXDATA - 1
        If DataHash(i) <> tempBuffer(i) Then cmpEqual = False: Exit For
    Next
    If Not cmpEqual Then
        i = 0
        Do While DataHash(i) = tempBuffer(i)
            i = i + 1
        Loop
        debug_ 1271, LOG_INFO, "recvDataCheck2", "%s (%u) has an incompatible mod. ([%d] got %x, expected %x)", getPlayerName(player), player, i, tempBuffer(i), DataHash(i)
        hasWrongData = True
    End If

    If aiIndex <> NetPlay_players_ai(player) Then
        debug_ 1277, LOG_INFO, "recvDataCheck2", "%s (%u) has a corrupted player state value. (ai: %i; should be: %i)", getPlayerName(player), player, aiIndex, NetPlay_players_ai(player)
        hasWrongData = True
    End If

    If Not NetPlay_players_isSpectator(player) And recvGM Then
        If recvGM Then msg = "true" Else msg = "false"
        debug_ 1283, LOG_INFO, "recvDataCheck2", "%s (%u) has a corrupted global state value. (godMode: %s)", getPlayerName(player), player, msg
        hasWrongData = True
    End If

    If hasWrongData Then
        If Not NetPlay_isHost Then
            ASSERT_FAILURE "false", "src/multiplay.cpp:1289", "recvDataCheck2", "Host only routine detected for client!"
            recvDataCheck2 = False
            Exit Function
        End If
        msg = va("%s (%u) has an incompatible mod, and has been kicked.", getPlayerName(player), player)
        sendInGameSystemMessage msg
        addConsoleMessage msg, LEFT_JUSTIFY, NOTIFY_MESSAGE

        kickPlayer player, "Your data doesn't match the host's!", ERROR_WRONGDATA, False
        recvDataCheck2 = False
        Exit Function
    End If

    recvDataCheck2 = True
End Function

' (VB) layers.find(zCheck): the index, -1 for end()
Private Function layerIndex(layerZ() As Long, ByVal nLayers As Long, ByVal zCheck As Long) As Long
    Dim k As Long
    For k = 0 To nLayers - 1
        If layerZ(k) = zCheck Then layerIndex = k: Exit Function
    Next
    layerIndex = -1
End Function

' bool NetworkTextMessage::receive(NETQUEUE queue)
Public Function NetworkTextMessage_receive(m As NetworkTextMessage, ByVal queue As Long) As Boolean
    Dim r As MessageReader
    Set r = NETbeginDecode(queue, NET_TEXTMSG)
    NetworkTextMessage_receive = NetworkTextMessage_decode(m, r, NETQUEUE_index(queue))
End Function

' bool NetworkTextMessage::decode(MessageReader& r, uint8_t senderIdx)
Public Function NetworkTextMessage_decode(m As NetworkTextMessage, ByVal r As MessageReader, ByVal senderIdx As Long) As Boolean
    m.text = ""                                 ' memset(text, 0x0, sizeof(text));

    If Not (NET_TEXTMSG = MessageReader_msgType(r)) Then ASSERT_FAILURE "NET_TEXTMSG == r.msgType", "src/multiplay.cpp:2079", "decode", "Unexpected message type: %u", MessageReader_msgType(r): NetworkTextMessage_decode = False: Exit Function

    NETint32_t r, m.sender
    NETbool r, m.teamSpecific
    NETstring r, m.text, MAX_CONSOLE_STRING_LENGTH
    NETend r

    If whosResponsible(m.sender) <> senderIdx Then
        m.sender = senderIdx  ' Fix corrupted sender.
    End If

    If m.sender >= MAX_CONNECTED_PLAYERS_J Then
        NetworkTextMessage_decode = False
        Exit Function
    End If
    If m.sender >= 0 Then
        If Not NetPlay_players_allocated(m.sender) And NetPlay_players_ai(m.sender) = AI_OPEN Then
            NetworkTextMessage_decode = False
            Exit Function
        End If
    End If

    NetworkTextMessage_decode = True
End Function

Public Sub sendInGameSystemMessage(ByVal text As String)
    Dim message As NetworkTextMessage
    message = NetworkTextMessage_new(SYSTEM_MESSAGE, text)
    printInGameTextMessage message
    If NetPlay_isHost Or Not NetPlay_players_isSpectator(selectedPlayer) Or GetGameMode() <> GS_NORMAL Then
        ' host + players can broadcast these at any time
        ' spectators can only broadcast in-game system messages before the game has started (i.e. in lobby)
        NetworkTextMessage_enqueue message, NETbroadcastQueue()
    End If
End Sub

'
' Read a message from the queue, and write it to the console.
'
' This message can be:
' - In game chat message
' - In game system message (player got kicked, player used cheat, etc.)
'
Public Function receiveInGameTextMessage(ByVal queue As Long) As Boolean
    Dim message As NetworkTextMessage, mutedUntil As Currency
    If Not NetworkTextMessage_receive(message, queue) Then
        receiveInGameTextMessage = False
        Exit Function
    End If

    If message.sender >= 0 Then
        If isPlayerMuted(message.sender) Then
            receiveInGameTextMessage = False
            Exit Function
        End If
    End If

    If message.sender >= 0 Then
        If playerSpamMutedUntil(message.sender, mutedUntil) Then
            ' discard messages from sender while spam-muted
            receiveInGameTextMessage = False
            Exit Function
        End If
        recordPlayerMessageSent message.sender
    End If

    printInGameTextMessage message
    cmdInterfaceLogChatMsg message, "WZCHATGAM"

    ' make some noise!
    If GetGameMode() <> GS_NORMAL Then
        audio_PlayTrack FE_AUDIO_MESSAGEEND
    ElseIf Not ingame.localJoiningInProgress Then
        audio_PlayTrack ID_SOUND_MESSAGEEND
    End If

    receiveInGameTextMessage = True
End Function

'AI multiplayer message - received message for AI (for hosted scripts)
Public Function recvTextMessageAI(ByVal queue As Long) As Boolean
    Dim sender As Long, receiver As Long, newmsg As String, r As MessageReader
    Set r = NETbeginDecode(queue, NET_AITEXTMSG)
    NETuint32_t r, sender           'in-game player index ('normal' one)
    NETuint32_t r, receiver         'in-game player index
    NETstring r, newmsg, MAX_CONSOLE_STRING_LENGTH
    NETend r

    If whosResponsible(sender) <> NETQUEUE_index(queue) Then
        sender = NETQUEUE_index(queue)  ' Fix corrupted sender.
    End If

    If isPlayerMuted(sender) Then
        recvTextMessageAI = False
        Exit Function
    End If

    triggerEventChat sender, receiver, newmsg

    recvTextMessageAI = True
End Function

' optional<NetworkTextMessage> decodeSpecInGameTextMessage(MessageReader& r, uint8_t senderIdx): False for nullopt
Public Function decodeSpecInGameTextMessage(ByVal r As MessageReader, ByVal senderIdx As Long, result As NetworkTextMessage) As Boolean
    Dim sender As Long, newmsg As String
    If Not (NET_SPECTEXTMSG = MessageReader_msgType(r)) Then ASSERT_FAILURE "NET_SPECTEXTMSG == r.msgType", "src/multiplay.cpp:2248", "decodeSpecInGameTextMessage", "Unexpected message type: %u", MessageReader_msgType(r): decodeSpecInGameTextMessage = False: Exit Function

    NETuint32_t r, sender           'in-game player index ('normal' one)
    NETstring r, newmsg, MAX_CONSOLE_STRING_LENGTH
    NETend r

    If whosResponsible(sender) <> senderIdx Then
        sender = senderIdx  ' Fix corrupted sender.
    End If

    If Not ULT(sender, MAX_CONNECTED_PLAYERS_J) Then
        decodeSpecInGameTextMessage = False
        Exit Function
    End If
    If Not NetPlay_players_allocated(sender) And NetPlay_players_ai(sender) = AI_OPEN Then
        decodeSpecInGameTextMessage = False
        Exit Function
    End If

    If Not NetPlay_players_isSpectator(selectedPlayer) Then
        decodeSpecInGameTextMessage = False ' ignore
        Exit Function
    End If

    result = NetworkTextMessage_new(sender, newmsg)
    decodeSpecInGameTextMessage = True
End Function

Public Function recvSpecInGameTextMessage(ByVal queue As Long) As Boolean
    Dim r As MessageReader, message As NetworkTextMessage, mutedUntil As Currency
    Set r = NETbeginDecode(queue, NET_SPECTEXTMSG)
    If Not decodeSpecInGameTextMessage(r, NETQUEUE_index(queue), message) Then
        recvSpecInGameTextMessage = False
        Exit Function
    End If

    If isPlayerMuted(message.sender) Then
        recvSpecInGameTextMessage = False
        Exit Function
    End If

    If message.sender >= 0 Then
        If playerSpamMutedUntil(message.sender, mutedUntil) Then
            ' discard messages from sender while spam-muted
            recvSpecInGameTextMessage = False
            Exit Function
        End If
        recordPlayerMessageSent message.sender
    End If

    printInGameTextMessage message
    cmdInterfaceLogChatMsg message, "WZCHATSPC"

    ' make some noise!
    If GetGameMode() <> GS_NORMAL Then
        audio_PlayTrack FE_AUDIO_MESSAGEEND
    ElseIf Not ingame.localJoiningInProgress Then
        audio_PlayTrack ID_SOUND_MESSAGEEND
    End If

    recvSpecInGameTextMessage = True
End Function

' ////////////////////////////////////////////////////////////////////////////
' Network File packet processor.
Public Function recvMapFileRequested(ByVal queue As Long) As Boolean
    Dim player As Long, hash As Sha256, r As MessageReader, b() As Byte, i As Long, filename As String, mapData As LEVEL_DATASET
    Dim mods() As WzMods_LoadedMod
    Dim pFileHandle As Long, fileSize_64 As Double, fileSize_u32 As Long, w As MessageWriter
    If Not NetPlay_isHost Then ASSERT_FAILURE "NetPlay.isHost", "src/multiplay.cpp:2357", "recvMapFileRequested", "Host only routine detected for client!": recvMapFileRequested = False: Exit Function
    player = NETQUEUE_index(queue)
    Sha256_setZero hash

    Set r = NETbeginDecode(queue, NET_FILE_REQUESTED)
    ReDim b(0 To 31)
    NETbin r, b, 32
    For i = 0 To 31: hash.bytes(i) = b(i): Next
    NETend r

    ' (ASSERT_OR_RETURN(false, files != nullptr, ...): WzNet's wzFiles is always made)
    If NetPlay_players_wzFiles_hasHash(player, hash.bytes) Then
        recvMapFileRequested = True  ' Already sending this file, do nothing.
        Exit Function
    End If

    netPlayersUpdated = True  ' Show download icon on player.

    If Sha256_equals(hash, game.hash) Then
        addConsoleMessage "Map was requested: SENDING MAP!", DEFAULT_JUSTIFY, SYSTEM_MESSAGE

        Set mapData = levFindDataSet_hash(game.map, game.hash, True)
        If mapData Is Nothing Then ASSERT_FAILURE "mapData", "src/multiplay.cpp:2382", "recvMapFileRequested", "levFindDataSet failed for game.map: %s", game.map: recvMapFileRequested = False: Exit Function
        If StrPtr(mapData.realFileName) = 0 Then ASSERT_FAILURE "mapData->realFileName != nullptr", "src/multiplay.cpp:2383", "recvMapFileRequested", "levFindDataSet found game.map: %s; but realFileName is empty - requesting a built-in map??", game.map: recvMapFileRequested = False: Exit Function
        filename = mapData.realFileName
        If Len(filename) = 0 Then
            debug_ 2387, LOG_INFO, "recvMapFileRequested", "Unknown map requested by %u.", player
            recvMapFileRequested = False
            Exit Function
        End If
        debug_ 2390, LOG_INFO, "recvMapFileRequested", "Map was requested. Looking for %s", filename
    Else
        filename = getModFilename(hash)
        If Len(filename) = 0 Then
            debug_ 2397, LOG_INFO, "recvMapFileRequested", "Unknown file requested by %u. (Hash: %s, loadedMods: %zu)", player, Sha256_toString(hash), getLoadedMods(mods)
            recvMapFileRequested = False
            Exit Function
        End If

        addConsoleMessage "Mod was requested: SENDING MOD!", DEFAULT_JUSTIFY, SYSTEM_MESSAGE
    End If

    ' Checking to see if file is available...
    pFileHandle = PHYSFS_openRead(filename)
    If pFileHandle = 0 Then
        debug_ 2408, LOG_ERROR, "recvMapFileRequested", "Failed to open %s for reading: %s", filename, WZ_PHYSFS_getLastError()
        debug_ 2409, LOG_FATAL, "recvMapFileRequested", "You have a map (%s) that can't be located." & vbLf & vbLf & "Make sure it is in the correct directory and or format! (No map packs!)", filename
        ' NOTE: if we get here, then the game is basically over, The host can't send the file for whatever reason...
        ' Which also means, that we can't continue.
        debug_ 2412, LOG_NET, "recvMapFileRequested", "***Host has a file issue, and is being forced to quit!***"
        Set w = NETbeginEncode(NETbroadcastQueue(), NET_HOST_DROPPED)
        NETend w
        exit_ 3                 ' abort()
    End If

    fileSize_64 = PHYSFS_fileLength(pFileHandle)
    If Not (fileSize_64 <= 4294967295#) Then ASSERT_FAILURE "fileSize_64 <= 0xFFFFFFFF", "src/multiplay.cpp:2419", "recvMapFileRequested", "File is too big!": recvMapFileRequested = False: Exit Function
    If Not (fileSize_64 >= 0) Then ASSERT_FAILURE "fileSize_64 >= 0", "src/multiplay.cpp:2420", "recvMapFileRequested", "Filesize < 0; can't be determined": recvMapFileRequested = False: Exit Function
    fileSize_u32 = FromU32(fileSize_64)
    If ULT(MAX_NET_TRANSFERRABLE_FILE_SIZE, fileSize_u32) Then ASSERT_FAILURE "fileSize_u32 <= MAX_NET_TRANSFERRABLE_FILE_SIZE", "src/multiplay.cpp:2422", "recvMapFileRequested", "Filesize is too large; (size: %u)", fileSize_u32: recvMapFileRequested = False: Exit Function

    ' Schedule file to be sent.
    debug_ 2425, LOG_INFO, "recvMapFileRequested", "File is valid, sending [directory: %s] %s to client %u", PHYSFS_getRealDir(filename), filename, player
    NetPlay_players_wzFiles_emplace_back player, pFileHandle, filename, hash.bytes, fileSize_u32

    recvMapFileRequested = True
End Function

' Continue sending maps and mods.
Public Sub sendMap()
    Const maxMicroSecondsPerSendMapCall As Double = (4 * 1000)
    Dim totalFilesToSend As Double, maxMicroSecondsPerFile As Double, i As Long, f As Long, done As Long
    Dim file_startTime As Currency, now_ As Currency, file_currentDuration As Double
    ' maximum "budget" in time per call to sendMap
    ' (at 60fps, total frame budget is ~16ms - allocate 4ms max for each call to sendMap)

    ' calculate the time budget per file
    totalFilesToSend = 0
    For i = 0 To MAX_CONNECTED_PLAYERS_J - 1
        totalFilesToSend = totalFilesToSend + NetPlay_players_wzFiles_size(i)
    Next
    If totalFilesToSend < 1 Then
        maxMicroSecondsPerFile = Fix(maxMicroSecondsPerSendMapCall / 1)
    Else
        maxMicroSecondsPerFile = Fix(maxMicroSecondsPerSendMapCall / totalFilesToSend)
    End If

    If perfFreqJ = 0 Then QueryPerformanceFrequency_J perfFreqJ
    For i = 0 To MAX_CONNECTED_PLAYERS_J - 1
        If NetPlay_players_wzFiles_size(i) = 0 Then GoTo continue_   ' (pFiles == nullptr: an empty list does nothing either)
        For f = 0 To NetPlay_players_wzFiles_size(i) - 1
            done = 0
            QueryPerformanceCounter_J file_startTime
            Do
                done = NETsendFile_idx(i, f)
                QueryPerformanceCounter_J now_
                file_currentDuration = Fix(CDbl(now_ - file_startTime) * 1000000# / CDbl(perfFreqJ))
            Loop While done < 100 And (file_currentDuration < maxMicroSecondsPerFile)
            If done = 100 Then
                netPlayersUpdated = True  ' Remove download icon from player.
                addConsoleMessage "FILE SENT!", DEFAULT_JUSTIFY, SYSTEM_MESSAGE
                debug_ 2472, LOG_INFO, "sendMap", "=== File has been sent to player %d ===", i
            End If
        Next
        NetPlay_players_wzFiles_eraseClosed i
continue_:
    Next
End Sub

' Another player is broadcasting a map, recv a chunk. Returns false if not yet done.
Public Function recvMapFileData(ByVal queue As Long) As Boolean
    Dim targeting As WzQuickChatTargeting, mapData As LEVEL_DATASET, buf As String
    NETrecvFile queue
    If NET_getDownloadingWzFiles_empty() Then
        netPlayersUpdated = True  ' Remove download icon from ourselves.
        addConsoleMessage "MAP DOWNLOADED!", DEFAULT_JUSTIFY, SYSTEM_MESSAGE
        targeting.all_ = True
        sendQuickChat WzQuickChatMessage_INTERNAL_LOBBY_NOTICE_MAP_DOWNLOADED, selectedPlayer, targeting
        debug_ 2492, LOG_INFO, "recvMapFileData", "=== File has been received. ==="

        ' clear out the old level list.
        levShutDown
        levInitialise
        rebuildSearchPath mod_multiplay, True  ' MUST rebuild search path for the new maps we just got!
        pal_Init 'Update palettes.
        If Not buildMapList() Then
            recvMapFileData = False
            Exit Function
        End If

        Set mapData = levFindDataSet_hash(game.map, game.hash, True)
        If Not mapData Is Nothing Then
            If CheckForMod(mapData.realFileName) Then
                If game.isMapMod Then
                    buf = "Warning, this is a map-mod, it could alter normal gameplay."
                Else
                    buf = "Warning, HOST has altered the game code, and can't be trusted!"
                End If
                addConsoleMessage buf, DEFAULT_JUSTIFY, NOTIFY_MESSAGE
                game.isMapMod = True
            End If
        End If
        If Not mapData Is Nothing Then
            If CheckForRandom(mapData.realFileName, mapData.apDataFiles(0)) Then
                game.isRandom = True
            End If
        End If

        loadMapPreview False
        recvMapFileData = True
        Exit Function
    End If

    recvMapFileData = False
End Function
