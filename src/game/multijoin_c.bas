Attribute VB_Name = "multijoin_c"
' src/multijoin.cpp, multijoin.h: the players leaving and joining (clearPlayer,
' destroyPlayerResources, splitResourcesAmongTeam, handlePlayerLeftInGame,
' recvPlayerLeft, MultiPlayerLeave, MultiPlayerJoin), the data checks
' (sendDataCheck, recvDataCheck), setupNewPlayer and ShowLobbyStatusMessage.
' NetPlay.playerReferences (lib/netplay's PlayerReference objects) and
' NetPlay.players[i].wzFiles are WzNet's: NetPlay_playerReferences_reset,
' NetPlay_players_wzFiles_clear.
Option Explicit

' lib/netplay/nettypes.h
Private Const MAX_CONNECTED_PLAYERS As Long = MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

' multilobbycommands.h
Private Const LOBBY_COMMAND_PREFIX As String = "/"

' splitResourcesAmongTeam's struct PlayerItemsReceived
Private Type PlayerItemsReceived
    player As Long                      ' uint32_t
    itemsRecv As Long                   ' size_t
End Type

'//////////////////////////////////////////////////////////////////////////////
'
' when a remote player leaves an arena game do this!
'
' @param player -- the one we need to clear
' @param quietly -- true means without any visible effects
'
Public Sub clearPlayer(ByVal player As Long, ByVal quietly As Boolean)
    Dim q As String
    If Not (ULT(player, MAX_CONNECTED_PLAYERS)) Then ASSERT_FAILURE "player < MAX_CONNECTED_PLAYERS", "src/multijoin.cpp:94", "clearPlayer", "Invalid player: %u", player: Exit Sub

    ' ASSERT(player < NetPlay.playerReferences.size(), ...); NetPlay.playerReferences[player]->disconnect(); NetPlay.playerReferences[player] = std::make_shared<PlayerReference>(player);
    NetPlay_playerReferences_reset player

    If quietly Then q = "true" Else q = "false"
    ' (VB) the macro evaluates getPlayerName only with LOG_NET on (it asserts for the empty slots past MAX_PLAYERS)
    If debugPartEnabled(LOG_NET) Then debug_ 100, LOG_NET, "clearPlayer", "R.I.P. %s (%u). quietly is %s", getPlayerName(player), player, q

    ingame.LagCounter(player) = 0
    ingame.DesyncCounter(player) = 0
    ingame.JoiningInProgress(player) = False   ' if they never joined, reset the flag
    ingame.DataIntegrity(player) = False
    ingame.hostChatPermissions(player) = False
    ingame.has_lastSentPlayerDataCheck2(player) = False

    If Not ULT(player, MAX_PLAYERS) Then
        Exit Sub ' no more to do
    End If

    destroyPlayerResources player, quietly
End Sub

Public Sub destroyPlayerResources(ByVal player As Long, ByVal quietly As Boolean)
    Dim i As Long, it As Long, nxt As Long, psStruct As Long

    If Not ULT(player, MAX_PLAYERS) Then
        Exit Sub ' no more to do
    End If

    For i = 0 To MAX_PLAYERS - 1                ' remove alliances
        ' Never remove a player's self-alliance, as the player can be selected and units added via the debug menu
        ' even after they have left, and this would lead to them firing on each other.
        If i <> player Then
            alliances(player, i) = ALLIANCE_BROKEN
            alliances(i, player) = ALLIANCE_BROKEN
            alliancebits(i) = alliancebits(i) And Not BIT_(player)
            alliancebits(player) = alliancebits(player) And Not BIT_(i)
        End If
    Next

    debug_ 139, LOG_DEATH, "destroyPlayerResources", "killing off all droids for player %d", player
    ' delete all droids
    ' mutating_list_iterate
    it = apsDroidLists(player).head
    Do While it <> 0
        nxt = lnNext(it)
        If quietly Then             ' don't show effects
            killDroid lnObj(it)
        Else                        ' show effects
            destroyDroid lnObj(it), gameTime
        End If
        it = nxt
    Loop

    debug_ 154, LOG_DEATH, "destroyPlayerResources", "killing off all structures for player %d", player
    ' delete all structs
    ' mutating_list_iterate
    it = apsStructLists(player).head
    Do While it <> 0
        nxt = lnNext(it)
        psStruct = lnObj(it)
        ' FIXME: look why destroyStruct() doesn't put back the feature like removeStruct() does
        If quietly Or asStructureStats(aStruct(psStruct).pStructureType).typ = REF_RESOURCE_EXTRACTOR Then       ' don't show effects
            removeStruct psStruct, True
        Else                        ' show effects
            destroyStruct psStruct, gameTime
        End If
        it = nxt
    Loop

    Exit Sub
End Sub


'--------------------------------------------------------------------------
' Cluster J (fork JL): the players leaving and joining during a game (splitResourcesAmongTeam, handlePlayerLeftInGame,
' recvPlayerLeft, MultiPlayerLeave, MultiPlayerJoin), the data checks, setupNewPlayer, ShowLobbyStatusMessage.
' NetPlay.playerReferences and NetPlay.players[i].wzFiles are WzNet's (NetApi). splitResourcesAmongTeam's lambdas
' (the structure tests) are a kind number: 0 resource extractor, 1 power generator, 2 research, 3 command control,
' 4 factory.

' (VB) the distributeMatchingStructs / destroyMatchingStructs cmp lambdas
Private Function structMatchesKind(ByVal psStruct As Long, ByVal kind As Long) As Boolean
    Select Case kind
    Case 0: structMatchesKind = asStructureStats(aStruct(psStruct).pStructureType).typ = REF_RESOURCE_EXTRACTOR
    Case 1: structMatchesKind = asStructureStats(aStruct(psStruct).pStructureType).typ = REF_POWER_GEN
    Case 2: structMatchesKind = asStructureStats(aStruct(psStruct).pStructureType).typ = REF_RESEARCH
    Case 3: structMatchesKind = asStructureStats(aStruct(psStruct).pStructureType).typ = REF_COMMAND_CONTROL
    Case 4: structMatchesKind = STRUCTURE_isFactory(psStruct)
    End Select
End Function

Private Function destroyMatchingStructs(ByVal player As Long, ByVal cmpKind As Long, ByVal quietly As Boolean) As Boolean
    Dim destroyedAnyStructs As Boolean, it As Long, nxt As Long, psStruct As Long
    destroyedAnyStructs = False
    ' mutating_list_iterate
    it = apsStructLists(player).head
    Do While it <> 0
        nxt = lnNext(it)
        psStruct = lnObj(it)
        If structMatchesKind(psStruct, cmpKind) Then
            ' FIXME: look why destroyStruct() doesn't put back the feature like removeStruct() does
            If quietly Or asStructureStats(aStruct(psStruct).pStructureType).typ = REF_RESOURCE_EXTRACTOR Then       ' don't show effects
                removeStruct psStruct, True
            Else            ' show effects
                destroyStruct psStruct, gameTime
            End If
            destroyedAnyStructs = True
        End If
        it = nxt
    Loop
    destroyMatchingStructs = destroyedAnyStructs
End Function

' auto incrRecvItem = [&](size_t idx) { items[idx].itemsRecv += 1; std::stable_sort(... by itemsRecv ...); }
Private Sub incrRecvItem(items() As PlayerItemsReceived, ByVal count As Long, ByVal idx As Long)
    Dim i As Long, j As Long, t As PlayerItemsReceived
    items(idx).itemsRecv = items(idx).itemsRecv + 1
    ' stable insertion sort by itemsRecv (ascending)
    For i = 1 To count - 1
        t = items(i)
        j = i - 1
        Do While j >= 0
            If items(j).itemsRecv <= t.itemsRecv Then Exit Do
            items(j + 1) = items(j)
            j = j - 1
        Loop
        items(j + 1) = t
    Next
End Sub

' auto distributeMatchingStructs = [&](std::function<bool (STRUCTURE *)> cmp)
Private Sub distributeMatchingStructs(ByVal player As Long, possibleTargets() As Long, ByVal nTargets As Long, ByVal cmpKind As Long)
    Dim structsGiftedPerTarget() As PlayerItemsReceived, i As Long, it As Long, nxt As Long, psStruct As Long
    ReDim structsGiftedPerTarget(0 To nTargets - 1)
    For i = 0 To nTargets - 1
        structsGiftedPerTarget(i).player = possibleTargets(i)
        structsGiftedPerTarget(i).itemsRecv = 0
    Next
    ' mutating_list_iterate
    it = apsStructLists(player).head
    Do While it <> 0
        nxt = lnNext(it)
        psStruct = lnObj(it)
        If psStruct <> 0 Then
            If structMatchesKind(psStruct, cmpKind) Then
                giftSingleStructure psStruct, structsGiftedPerTarget(0).player, False
                incrRecvItem structsGiftedPerTarget, nTargets, 0
            End If
        End If
        it = nxt
    Loop
End Sub

Public Function splitResourcesAmongTeam(ByVal player As Long) As Boolean
    Dim team As Long, possibleTargets() As Long, nTargets As Long, i As Long, powerPerTarget As Long
    Dim additionalDroidsLimitPerTarget As Long, additionalCommandersLimitPerTarget As Long, additionalConstructorsLimitPerTarget As Long
    Dim droidsGiftedPerTarget() As PlayerItemsReceived, it As Long, nxt As Long, d As Long, transferredDroid As Boolean
    Dim noPos As Vector2i
    team = NetPlay_players_team(player)

    ' Build a list of team members who are still around
    ReDim possibleTargets(0 To MAX_PLAYERS - 1)
    For i = 0 To MAX_PLAYERS - 1
        If i <> player And i <> scavengerSlot() And NetPlay_players_team(i) = team And aiCheckAlliances(i, player) _
            And isHumanPlayer(i) And Not NetPlay_players_isSpectator(i) Then
            ' ...not scavenger player, belonging to the same team, the alliance hasn't been broken, is a human, not spectator
            ' (NetPlay.players[i].difficulty != AIDifficulty::DISABLED is not checked: the host may set difficulty ==
            ' DISABLED for slots before clients do, leading to sync issues)
            possibleTargets(nTargets) = i
            nTargets = nTargets + 1
        End If
    Next

    If nTargets = 0 Then
        ' no valid targets for resources...
        splitResourcesAmongTeam = False
        Exit Function
    End If

    ' Distribute power evenly
    powerPerTarget = getPower(player) \ nTargets
    For i = 0 To nTargets - 1
        addPower possibleTargets(i), powerPerTarget
    Next
    setPower player, 0

    ' Distribute the player's additional unit limits
    additionalDroidsLimitPerTarget = getMaxDroids(player) \ nTargets
    additionalCommandersLimitPerTarget = getMaxCommanders(player) \ nTargets
    additionalConstructorsLimitPerTarget = getMaxConstructors(player) \ nTargets
    For i = 0 To nTargets - 1
        setMaxDroids possibleTargets(i), getMaxDroids(possibleTargets(i)) + additionalDroidsLimitPerTarget
        setMaxCommanders possibleTargets(i), getMaxCommanders(possibleTargets(i)) + additionalCommandersLimitPerTarget
        setMaxConstructors possibleTargets(i), getMaxConstructors(possibleTargets(i)) + additionalConstructorsLimitPerTarget
    Next

    ' Distribute droids between targets as evenly as possible
    ReDim droidsGiftedPerTarget(0 To nTargets - 1)
    For i = 0 To nTargets - 1
        droidsGiftedPerTarget(i).player = possibleTargets(i)
        droidsGiftedPerTarget(i).itemsRecv = 0
    Next
    ' mutating_list_iterate
    it = apsDroidLists(player).head
    Do While it <> 0
        nxt = lnNext(it)
        d = lnObj(it)
        transferredDroid = False
        If Not isDead(d) Then
            For i = 0 To nTargets - 1
                If giftSingleDroid(d, droidsGiftedPerTarget(i).player, False, noPos) <> 0 Then
                    transferredDroid = True
                    incrRecvItem droidsGiftedPerTarget, nTargets, i
                    Exit For
                End If
                ' if we can't gift this droid to this player, try again with the next player in priority queue
            Next
        End If
        If Not transferredDroid Then
            destroyDroid d, gameTime
        End If
        it = nxt
    Loop

    ' Distribute key structures
    distributeMatchingStructs player, possibleTargets, nTargets, 0
    distributeMatchingStructs player, possibleTargets, nTargets, 1
    If alliancesSharedResearch(game.alliance) Then
        distributeMatchingStructs player, possibleTargets, nTargets, 2
    Else
        ' destroy the research centers in unshared research mode
        destroyMatchingStructs player, 2, False
    End If
    distributeMatchingStructs player, possibleTargets, nTargets, 3
    distributeMatchingStructs player, possibleTargets, nTargets, 4

    splitResourcesAmongTeam = True
End Function

Public Sub handlePlayerLeftInGame(ByVal player As Long)
    Dim leftWhilePlayer As Boolean
    If Not ULT(player, MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "player < MAX_CONNECTED_PLAYERS", "src/multijoin.cpp:328", "handlePlayerLeftInGame", "Invalid player: %u", player: Exit Sub

    leftWhilePlayer = Not NetPlay_players_isSpectator(player)

    ' ASSERT(player < NetPlay.playerReferences.size(), ...); NetPlay.playerReferences[player]->disconnect(); NetPlay.playerReferences[player] = std::make_shared<PlayerReference>(player);
    NetPlay_playerReferences_reset player

    ' (VB) the macro evaluates getPlayerName only with LOG_NET on (it asserts for the empty slots past MAX_PLAYERS)
    If debugPartEnabled(LOG_NET) Then debug_ 336, LOG_NET, "handlePlayerLeftInGame", "R.I.P. %s (%u).", getPlayerName(player), player

    ingame.LagCounter(player) = 0
    ingame.DesyncCounter(player) = 0
    ingame.JoiningInProgress(player) = False   ' if they never joined, reset the flag
    ingame.PendingDisconnect(player) = False
    ingame.DataIntegrity(player) = False
    ingame.has_lastSentPlayerDataCheck2(player) = False
    If leftWhilePlayer Then
        ingame.playerLeftGameTime(player) = gameTime
        ingame.has_playerLeftGameTime(player) = True
    End If

    If player >= MAX_PLAYERS Then
        Exit Sub ' no more to do
    End If

    Select Case game.playerLeaveMode
    Case PLAYER_LEAVE_MODE_DESTROY_RESOURCES
        destroyPlayerResources player, False
    Case PLAYER_LEAVE_MODE_SPLIT_WITH_TEAM
        If Not splitResourcesAmongTeam(player) Then
            ' no valid targets to split resources among
            ' instead, destroy the player
            destroyPlayerResources player, False
        End If
    End Select
End Sub

' Reset visibility, so a new player can't see the old stuff!!
Private Sub resetMultiVisibility(ByVal player As Long)
    Dim owned As Long, it As Long
    If player >= MAX_PLAYERS Then
        Exit Sub
    End If

    For owned = 0 To MAX_PLAYERS - 1       ' for each player
        If owned <> player Then             ' done reset own stuff..
            'droids
            it = apsDroidLists(owned).head
            Do While it <> 0
                BASE_OBJECT_visible(lnObj(it), player) = 0
                it = lnNext(it)
            Loop

            'structures
            it = apsStructLists(owned).head
            Do While it <> 0
                BASE_OBJECT_visible(lnObj(it), player) = 0
                it = lnNext(it)
            Loop
        End If
    Next
End Sub

Private Sub sendPlayerLeft(ByVal playerIndex As Long)
    Dim forcedPlayerIndex As Long, w As MessageWriter, q As Long
    If Not NetPlay_isHost Then ASSERT_FAILURE "NetPlay.isHost", "src/multijoin.cpp:405", "sendPlayerLeft", "Only host should call this.": Exit Sub

    forcedPlayerIndex = whosResponsible(playerIndex)
    If forcedPlayerIndex <> selectedPlayer Then q = NETgameQueueForced(forcedPlayerIndex) Else q = NETgameQueue(forcedPlayerIndex)
    Set w = NETbeginEncode(q, GAME_PLAYER_LEFT)
    NETuint32_t w, playerIndex
    NETend w
End Sub

Private Sub addConsolePlayerLeftMessage(ByVal playerIndex As Long)
    Dim msg As String
    If Not NetPlay_isHost And isBlindSimpleLobby(game.blindMode) And (GetGameMode() <> GS_NORMAL) Then
        Exit Sub
    End If
    If selectedPlayer <> playerIndex Then
        msg = va("%s has Left the Game", getPlayerName(playerIndex))
        addConsoleMessage msg, DEFAULT_JUSTIFY, SYSTEM_MESSAGE
    End If
End Sub

Private Sub addConsolePlayerJoinMessage(ByVal playerIndex As Long)
    Dim msg As String
    If Not NetPlay_isHost And isBlindSimpleLobby(game.blindMode) And (GetGameMode() <> GS_NORMAL) Then
        Exit Sub
    End If
    If selectedPlayer <> playerIndex Then
        msg = va("%s joined the Game", getPlayerName(playerIndex))
        If (game.blindMode <> BLIND_MODE_NONE) And NetPlay_isHost And (NetPlay_hostPlayer >= MAX_PLAYER_SLOTS) Then
            msg = msg & " "
            msg = msg & va("(codename: %s)", getPlayerGenericName(playerIndex))
        End If
        addConsoleMessage msg, DEFAULT_JUSTIFY, SYSTEM_MESSAGE
    End If
End Sub

Public Sub recvPlayerLeft(ByVal queue As Long)
    Dim playerIndex As Long, r As MessageReader
    playerIndex = 0
    Set r = NETbeginDecode(queue, GAME_PLAYER_LEFT)
    NETuint32_t r, playerIndex
    NETend r

    addConsolePlayerLeftMessage playerIndex

    If whosResponsible(playerIndex) <> NETQUEUE_index(queue) Then
        Exit Sub
    End If

    turnOffMultiMsg True
    handlePlayerLeftInGame playerIndex
    turnOffMultiMsg False
    If Not ingame.has_TimeEveryoneIsInGame Then ' If game hasn't actually started
        clearPlayerMultiStats playerIndex ' local only
    End If
    NetPlay_players_allocated(playerIndex) = False

    NETsetPlayerConnectionStatus CONNECTIONSTATUS_PLAYER_DROPPED, playerIndex
    cancelOrDismissKickVote playerIndex

    debug_ 471, LOG_INFO, "recvPlayerLeft", "** player %u has dropped, in-game! (gameTime: %u)", playerIndex, gameTime

    ' fire script callback to reassign skirmish players.
    If GetGameMode() = GS_NORMAL Then
        triggerEventPlayerLeft playerIndex
    End If

    ActivityManager_updateMultiplayGameData
    wz_command_interface_output_room_status_json
End Sub

' ////////////////////////////////////////////////////////////////////////////
' A remote player has left the game
Public Function MultiPlayerLeave(ByVal playerIndex As Long) As Boolean
    Dim identity As EcKey, keyBytes() As Byte, n As Long, playerPublicKeyB64 As String, playerIdentityHash As String
    Dim playerVerifiedStatus As String, playerName As String, nameBytes() As Byte, playerNameB64 As String, w As MessageWriter
    Dim buf As String
    If Not ULT(playerIndex, MAX_CONNECTED_PLAYERS) Then
        ASSERT_FAILURE "false", "src/multijoin.cpp:490", "MultiPlayerLeave", "Bad player number"
        MultiPlayerLeave = False
        Exit Function
    End If

    If NetPlay_isHost Then
        multiClearHostRequestMoveToPlayer playerIndex
        multiSyncResetPlayerChallenge playerIndex
        resetMultiOptionPrefValues playerIndex
    End If

    NETlogEntry "Player leaving game", SYNC_FLAG, playerIndex
    ' (VB) the macro evaluates getPlayerName only with LOG_NET on
    If debugPartEnabled(LOG_NET) Then debug_ 502, LOG_NET, "MultiPlayerLeave", "** Player %u [%s], has left the game at game time %u.", playerIndex, getPlayerName(playerIndex), gameTime

    ingame.muteChat(playerIndex) = False

    If wz_command_interface_enabled() And NetPlay_players_allocated(playerIndex) Then
        ' WZEVENT: playerLeft: <playerIdx> <gameTime> <b64pubkey> <hash> <V|?> <b64name> <ip>
        identity = getOutputPlayerIdentity(playerIndex)
        n = EcKey_toBytes(identity, EcKey_Public, keyBytes)
        playerPublicKeyB64 = base64Encode(keyBytes, n)
        playerIdentityHash = EcKey_publicHashString(identity)
        If ingame.VerifiedIdentity(playerIndex) Then playerVerifiedStatus = "V" Else playerVerifiedStatus = "?"
        playerName = getPlayerName(playerIndex)
        n = utf8_encode(playerName, nameBytes)
        playerNameB64 = base64Encode(nameBytes, n)
        wz_command_interface_output "WZEVENT: playerLeft: %u %u %s %s %s %s %s" & vbLf, playerIndex, gameTime, playerPublicKeyB64, playerIdentityHash, playerVerifiedStatus, playerNameB64, NetPlay_players_IPtextAddress(playerIndex)
    End If

    If ingame.localJoiningInProgress Then
        addConsolePlayerLeftMessage playerIndex
        clearPlayer playerIndex, False
        clearPlayerMultiStats playerIndex ' local only
        NetPlay_players_difficulty(playerIndex) = AIDifficulty_DISABLED
    ElseIf NetPlay_isHost Then  ' If hosting, and game has started (not in pre-game lobby screen, that is).
        sendPlayerLeft playerIndex
        If bDisplayMultiJoiningStatus <> 0 Then ' if still waiting for players to load *or* waiting for game to start...
            Set w = NETbeginEncode(NETbroadcastQueue(), NET_PLAYER_DROPPED)
            NETuint32_t w, playerIndex
            NETend w
            ' only set ingame.JoiningInProgress[player_id] to false
            ' when the game starts, it will handle the GAME_PLAYER_LEFT message in their queue properly
            ingame.JoiningInProgress(playerIndex) = False
            ingame.PendingDisconnect(playerIndex) = True ' used as a UI indicator that a disconnect will be processed in the future
        End If
    End If

    If NetPlay_players_fileSendInProgress(playerIndex) Then
        buf = va("File transfer has been aborted for %d.", playerIndex)
        addConsoleMessage buf, DEFAULT_JUSTIFY, SYSTEM_MESSAGE
        debug_ 547, LOG_INFO, "MultiPlayerLeave", "=== File has been aborted for %d ===", playerIndex
        NetPlay_players_wzFiles_clear playerIndex
    End If

    If Not widgGetFromID(psWScreen, IDRET_FORM) Is Nothing Then
        If playerIndex < MAX_PLAYERS Then ' only play audio when *player* slots drop (ignore spectator slots)
            audio_QueueTrack ID_CLAN_EXIT
        End If
    End If

    netPlayersUpdated = True
    MultiPlayerLeave = True
End Function

' ////////////////////////////////////////////////////////////////////////////
' A Remote Player has joined the game.
' optional<EcKey::Key> verifiedJoinIdentity: hasVerifiedJoinIdentity and the public key's bytes
' verifiedJoinIdentity(0 .. identityLen - 1)
Public Function MultiPlayerJoin(ByVal playerIndex As Long, ByVal hasVerifiedJoinIdentity As Boolean, verifiedJoinIdentity() As Byte, _
        ByVal identityLen As Long) As Boolean
    Dim i As Long
    If Not widgGetFromID(psWScreen, IDRET_FORM) Is Nothing Then  ' if ingame.
        audio_QueueTrack ID_CLAN_ENTER
    End If

    If Not widgGetFromID(psWScreen, MULTIOP_PLAYERS) Is Nothing Then ' if in multimenu.
        If Not multiRequestUp And (NetPlay_isHost Or ingame.localJoiningInProgress) Then
            netPlayersUpdated = True   ' update the player box.
        End If
    End If

    playerSpamMuteReset playerIndex

    If NetPlay_isHost Then     ' host responsible for welcoming this player.
        ' if we've already received a request from this player don't reallocate.
        If ingame.JoiningInProgress(playerIndex) Then
            MultiPlayerJoin = True
            Exit Function
        End If
        If Not (NetPlay_playercount <= MAX_PLAYERS) Then ASSERT_FAILURE "NetPlay.playercount <= 11", "src/multijoin.cpp:589", "MultiPlayerJoin", "Too many players!"
        If Not (GetGameMode() <> GS_NORMAL) Then ASSERT_FAILURE "GetGameMode() != GS_NORMAL", "src/multijoin.cpp:590", "MultiPlayerJoin", "A player joined after the game started??"

        ' setup data for this player, then broadcast it to the other players.
        setupNewPlayer playerIndex                         ' setup all the guff for that player.
        If hasVerifiedJoinIdentity Then
            multiStatsSetVerifiedIdentityFromJoin playerIndex, verifiedJoinIdentity, identityLen
        End If
        sendOptions
        ' if skirmish and game full, then kick...
        If NetPlay_playercount > game.maxPlayers Then
            kickPlayer playerIndex, "The game is already full.", ERROR_FULL, False
        End If
        ' send everyone's stats to the new guy
        For i = 0 To MAX_CONNECTED_PLAYERS - 1
            If NetPlay_players_allocated(i) Then
                sendMultiStats i
            End If
        Next
        If lobby_slashcommands_enabled() Then
            ' Inform the new player that this lobby has slash commands enabled.
            sendRoomSystemMessageToSingleReceiver "Lobby slash commands enabled. Type " & LOBBY_COMMAND_PREFIX & "help to see details.", playerIndex, True
        End If
    End If
    addConsolePlayerJoinMessage playerIndex
    MultiPlayerJoin = True
End Function

Public Function sendDataCheck() As Boolean
    Dim i As Long, w As MessageWriter
    Set w = NETbeginEncode(NETnetQueue(NetPlay_hostPlayer), NET_DATA_CHECK)       ' only need to send to HOST
    For i = 0 To DATA_MAXDATA - 1
        NETuint32_t w, DataHash(i)
    Next
    NETend w
    debug_ 636, LOG_NET, "sendDataCheck", "sent hash to host"
    sendDataCheck = True
End Function

Public Function recvDataCheck(ByVal queue As Long) As Boolean
    Dim i As Long, player As Long, tempBuffer(0 To DATA_MAXDATA - 1) As Long, r As MessageReader, msg As String
    Dim same As Boolean
    player = NETQUEUE_index(queue)

    If Not NetPlay_isHost Then              ' only host should act
        ASSERT_FAILURE "false", "src/multijoin.cpp:648", "recvDataCheck", "Host only routine detected for client!"
        recvDataCheck = False
        Exit Function
    End If

    Set r = NETbeginDecode(queue, NET_DATA_CHECK)
    For i = 0 To DATA_MAXDATA - 1
        NETuint32_t r, tempBuffer(i)
    Next
    NETend r

    If Not ULT(player, MAX_CONNECTED_PLAYERS) Then ' invalid player number.
        debug_ 661, LOG_ERROR, "recvDataCheck", "invalid player number (%u) detected.", player
        recvDataCheck = False
        Exit Function
    End If

    If whosResponsible(player) <> NETQUEUE_index(queue) Then
        HandleBadParam "NET_DATA_CHECK given incorrect params.", player, NETQUEUE_index(queue)
        recvDataCheck = False
        Exit Function
    End If

    debug_ 671, LOG_NET, "recvDataCheck", "** Received NET_DATA_CHECK from player %u", player

    If NetPlay_isHost Then
        same = True
        For i = 0 To DATA_MAXDATA - 1
            If DataHash(i) <> tempBuffer(i) Then same = False: Exit For
        Next
        If Not same Then
            i = 0
            Do While DataHash(i) = tempBuffer(i)
                i = i + 1
            Loop
            msg = Left$(va("%s (%u) has an incompatible mod, and has been kicked.", getPlayerName(player), player), 255)
            sendInGameSystemMessage msg
            addConsoleMessage msg, LEFT_JUSTIFY, NOTIFY_MESSAGE

            kickPlayer player, "Your data doesn't match the host's!", ERROR_WRONGDATA, False
            debug_ 688, LOG_ERROR, "recvDataCheck", "%s (%u) has an incompatible mod. ([%d] got %x, expected %x)", getPlayerName(player), player, i, tempBuffer(i), DataHash(i)

            recvDataCheck = False
            Exit Function
        Else
            If debugPartEnabled(LOG_NET) Then debug_ 694, LOG_NET, "recvDataCheck", "DataCheck message received and verified for player %s (slot=%u)", getPlayerName(player), player
            ingame.DataIntegrity(player) = True
        End If
    End If
    recvDataCheck = True
End Function

' ////////////////////////////////////////////////////////////////////////////
' Setup Stuff for a new player.
Public Sub setupNewPlayer(ByVal player As Long)
    Dim buf As String, st As PLAYERSTATS_
    If Not NetPlay_isHost Then
        ASSERT_FAILURE "false", "src/multijoin.cpp:704", "setupNewPlayer", "Host only routine detected for client!"
        Exit Sub
    End If
    If Not ULT(player, MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "player < MAX_CONNECTED_PLAYERS", "src/multijoin.cpp:705", "setupNewPlayer", "Invalid player: %u", player: Exit Sub

    ingame.PingTimes(player) = 0                   ' Reset ping time
    ingame.LagCounter(player) = 0
    ingame.DesyncCounter(player) = 0
    ingame.VerifiedIdentity(player) = False
    ingame.JoiningInProgress(player) = True         ' Note that player is now joining
    QueryPerformanceCounter ingame.joinTimes(player)
    ingame.has_joinTimes(player) = True
    ingame.PendingDisconnect(player) = False
    ingame.DataIntegrity(player) = False
    If NetPlay_bComms Then ingame.hostChatPermissions(player) = NETgetDefaultMPHostFreeChatPreference() Else ingame.hostChatPermissions(player) = True
    ingame.has_lastSentPlayerDataCheck2(player) = False
    ingame.muteChat(player) = False
    ingame.has_lastReadyTimes(player) = False
    If multiplayPlayersShouldCheckReady() Then
        ingame.lastNotReadyTimes(player) = ingame.joinTimes(player)
        ingame.has_lastNotReadyTimes(player) = True
    Else
        ingame.has_lastNotReadyTimes(player) = False
    End If
    ingame.secondsNotReady(player) = 0
    ingame.has_playerLeftGameTime(player) = False

    multiSyncResetPlayerChallenge player
    resetMultiVisibility player                       ' set visibility flags.

    st = getMultiStats(player)
    setMultiStats player, st, True  ' get the players score

    If selectedPlayer <> player Then
        buf = Left$(va("%s is joining the game", getPlayerName(player)), 254)
        addConsoleMessage buf, DEFAULT_JUSTIFY, SYSTEM_MESSAGE
    End If
End Sub

' While not the perfect place for this, it has to do when a HOST joins (hosts) game
' unfortunately, we don't get the message until after the setup is done.
' msgs: msgs(0 .. count - 1)
Public Sub ShowLobbyStatusMessage(msgs() As String, ByVal count As Long)
    Dim buf As String, i As Long
    ' when HOST joins the game, show server MOTD message first
    addConsoleMessage "Server message:", DEFAULT_JUSTIFY, NOTIFY_MESSAGE
    If count > 0 Then
        For i = 0 To count - 1
            addConsoleMessage msgs(i), DEFAULT_JUSTIFY, NOTIFY_MESSAGE
        Next
    Else
        buf = "Null message"
        addConsoleMessage buf, DEFAULT_JUSTIFY, NOTIFY_MESSAGE
    End If
    audio_PlayTrack FE_AUDIO_MESSAGEEND
End Sub
