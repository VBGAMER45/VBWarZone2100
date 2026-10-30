Attribute VB_Name = "multiplay_c"
' src/multiplay.cpp, multiplay.h: so far turnOffMultiMsg, the IdTo* lookups
' (_IdToStruct is IdToStruct_), isHumanPlayer, whosResponsible,
' myResponsibility, responsibleFor, canGiveOrdersFor, scavengerSlot,
' scavengerPlayer, getPlayerGenericName, the sync requests, the game
' messages' dispatch (getMessageHandlingAction, shouldProcessMessage,
' recvMessage, HandleBadParam), the research messages (SendResearch,
' recvResearch, sendResearchStatus, recvResearchStatus),
' findResearchingFacilityByResearchIndex (the list overload is _lists),
' recvDestroyFeature, NetworkTextMessage (its constructor and enqueue),
' printInGameTextMessage, the beacons (sendBeacon, CreateBeaconViewData,
' findBeaconMsg, addBeaconBlip, sendBeaconToPlayer, recvBeacon),
' resetReadyStatus, findPlayerIndexByPosition, makePlayerSpectator,
' multiPlayerLoop and openchannels (multiplay_c_init gives its
' initialiser). A NETQUEUE is WzNet's Long (NETQUEUE_index, ...).
' What recvMessage hands on to the chat, lobby, votes, gifts, ... is
' stubs_game's for now. Its globals (game, ingame, bMultiPlayer,
' bMultiMessages, ...) are still stubs_game's; the multiplayer code comes
' with clusters B and J.
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
