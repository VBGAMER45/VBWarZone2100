Attribute VB_Name = "netplay_sys_c"
' lib/netplay/netplay.cpp (3 of 4): the system messages (NETprocessSystemMessage, with the filtering while a
' player's index is being swapped and the chat log of the command interface), the player index swaps
' (swapPlayerIndexes, the spectator slots: NETcanOpenNewSpectatorSlot, NETopenNewSpectatorSlot,
' NETmovePlayerToSpectatorOnlySlot, NETmoveSpectatorToPlayerSlot) and class PlayerManagementRecord (the IPs
' and identities the host moved to the spectators: two std::unordered_set<std::string>, string arrays here).
Option Explicit

Private Const INDEX_CHANGE_ACK_TIMEOUT As Long = 5 * GAME_TICKS_PER_SEC

' (netplay.h's enum class SpectatorToPlayerMoveResult: NetApi's)

' class PlayerManagementRecord's two sets
Private identitiesMovedToSpectatorsByHost() As String
Private identitiesMovedN As Long
Private ipsMovedToSpectatorsByHost() As String
Private ipsMovedN As Long

'--------------------------------------------------------------------------
' class PlayerManagementRecord

Public Sub playerManagementRecord_clear()
    identitiesMovedN = 0
    ipsMovedN = 0
End Sub

' movedPlayerToSpectators(const std::string& ipAddress, const EcKey::Key& publicIdentity, bool byHost)
Public Sub playerManagementRecord_movedPlayerToSpectators(ByVal ipAddress As String, publicIdentity() As Byte, ByVal publicIdentitySize As Long, ByVal byHost As Boolean)
    If Not byHost Then Exit Sub
    setInsert ipsMovedToSpectatorsByHost, ipsMovedN, ipAddress
    setInsert identitiesMovedToSpectatorsByHost, identitiesMovedN, base64Encode(publicIdentity, publicIdentitySize)
End Sub

' movedSpectatorToPlayers(const std::string& ipAddress, const EcKey::Key& publicIdentity, bool byHost)
Public Sub playerManagementRecord_movedSpectatorToPlayers(ByVal ipAddress As String, publicIdentity() As Byte, ByVal publicIdentitySize As Long, ByVal byHost As Boolean)
    If Not byHost Then Exit Sub
    setErase ipsMovedToSpectatorsByHost, ipsMovedN, ipAddress
    setErase identitiesMovedToSpectatorsByHost, identitiesMovedN, base64Encode(publicIdentity, publicIdentitySize)
End Sub

Public Function playerManagementRecord_hostMovedPlayerToSpectators_ip(ByVal ipAddress As String) As Boolean
    playerManagementRecord_hostMovedPlayerToSpectators_ip = setFind(ipsMovedToSpectatorsByHost, ipsMovedN, ipAddress) >= 0
End Function

Public Function playerManagementRecord_hostMovedPlayerToSpectators_key(publicIdentity() As Byte, ByVal publicIdentitySize As Long) As Boolean
    playerManagementRecord_hostMovedPlayerToSpectators_key = setFind(identitiesMovedToSpectatorsByHost, identitiesMovedN, base64Encode(publicIdentity, publicIdentitySize)) >= 0
End Function

Private Function setFind(s() As String, ByVal n As Long, ByVal v As String) As Long
    Dim k As Long
    For k = 0 To n - 1
        If s(k) = v Then
            setFind = k
            Exit Function
        End If
    Next
    setFind = -1
End Function

Private Sub setInsert(s() As String, n As Long, ByVal v As String)
    If setFind(s, n, v) >= 0 Then Exit Sub
    If n = 0 Then
        ReDim s(0 To 15)
    ElseIf n > UBound(s) Then
        ReDim Preserve s(0 To 2 * n - 1)
    End If
    s(n) = v
    n = n + 1
End Sub

Private Sub setErase(s() As String, n As Long, ByVal v As String)
    Dim k As Long, j As Long
    k = setFind(s, n, v)
    If k < 0 Then Exit Sub
    For j = k To n - 2
        s(j) = s(j + 1)
    Next
    n = n - 1
End Sub

'--------------------------------------------------------------------------
' Player index swapping

' Warning: This should be used with care, and generally should not be called directly.
' Use higher-level functions like NETmovePlayerToSpectatorOnlySlot & NETrequestSpectatorToPlay instead.
Private Function swapPlayerIndexes(ByVal playerIndexA As Long, ByVal playerIndexB As Long) As Boolean
    Dim w As MessageWriter, playersData(0 To 1) As PLAYER_, playerIndexes(0 To 1) As Long, k As Long, playerIndex As Long
    Dim tmpSock As IClientConnection, t As Long, tb As Boolean
    If Not netGame.ingameLocalJoiningInProgress() Then ASSERT_FAILURE "ingame.localJoiningInProgress", "lib/netplay/netplay.cpp:1827", "swapPlayerIndexes", "Only supported in lobby, before game starts": swapPlayerIndexes = False: Exit Function
    If Not ULT(playerIndexA, MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "playerIndexA < MAX_CONNECTED_PLAYERS", "lib/netplay/netplay.cpp:1828", "swapPlayerIndexes", "playerIndexA out of bound: %u", playerIndexA: swapPlayerIndexes = False: Exit Function
    If Not ULT(playerIndexB, MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "playerIndexB < MAX_CONNECTED_PLAYERS", "lib/netplay/netplay.cpp:1829", "swapPlayerIndexes", "playerIndexB out of bound: %u", playerIndexB: swapPlayerIndexes = False: Exit Function
    If Not (playerIndexA <> NetPlay.hostPlayer And playerIndexB <> NetPlay.hostPlayer) Then ASSERT_FAILURE _
        "playerIndexA != NetPlay.hostPlayer && playerIndexA != NetPlay.hostPlayer && playerIndexB != NetPlay.hostPlayer && playerIndexB != NetPlay.hostPlayer", _
        "lib/netplay/netplay.cpp:1830", "swapPlayerIndexes", "Can't swap host player index: (index A: %u, index B: %u)", playerIndexA, playerIndexB: swapPlayerIndexes = False: Exit Function

    ' Send the NET_SWAPPING_PLAYER_INDEX message *first*
    Set w = NETbeginEncode(NETbroadcastQueue(), NET_PLAYER_SWAP_INDEX)
    NETuint32_t_w w, playerIndexA
    NETuint32_t_w w, playerIndexB
    NETend_w w

    ' Then swap the networking stuff for these slots
    Set tmpSock = connected_bsocket(playerIndexA)
    Set connected_bsocket(playerIndexA) = connected_bsocket(playerIndexB)
    Set connected_bsocket(playerIndexB) = tmpSock
    ' should be no need to call SocketSet_AddSocket, since should already be in the socket_set
    NETswapQueues NETnetQueue(playerIndexA), NETnetQueue(playerIndexB)

    ' Backup the old NETPLAY PLAYERS data
    playersData(0) = NetPlay.players(playerIndexA)
    playersData(1) = NetPlay.players(playerIndexB)

    ' Instead of calling clearPlayer() which has all kinds of unintended effects,
    ' hmmm...
    playerIndexes(0) = playerIndexA
    playerIndexes(1) = playerIndexB
    For k = 0 To 1
        playerIndex = playerIndexes(k)
        ' From MultiPlayerLeave()
        '  From clearPlayer() - this is needed to disconnect PlayerReferences so, for example, old chat messages are still associated with the proper player
        NetPlay.playerReferences(playerIndex).disconnect
        Set NetPlay.playerReferences(playerIndex) = PlayerReference_new(playerIndex)
    Next

    ' Swap the player multistats / identity info
    netGame.swapPlayerMultiStatsLocal playerIndexA, playerIndexB

    ' Swap the NetPlay PLAYER entries
    NetPlay.players(playerIndexB) = playersData(0)
    NetPlay.players(playerIndexA) = playersData(1)
    ' Just like changePosition(), we should *preserve* the team and "position" from the original index
    t = NetPlay.players(playerIndexA).position: NetPlay.players(playerIndexA).position = NetPlay.players(playerIndexB).position: NetPlay.players(playerIndexB).position = t
    t = NetPlay.players(playerIndexA).team: NetPlay.players(playerIndexA).team = NetPlay.players(playerIndexB).team: NetPlay.players(playerIndexB).team = t
    ' And also swap spectator flag!
    tb = NetPlay.players(playerIndexA).isSpectator: NetPlay.players(playerIndexA).isSpectator = NetPlay.players(playerIndexB).isSpectator: NetPlay.players(playerIndexB).isSpectator = tb
    t = NetPlay.players(playerIndexA).colour: NetPlay.players(playerIndexA).colour = NetPlay.players(playerIndexB).colour: NetPlay.players(playerIndexB).colour = t ' And the color!
    ' (Essentially all of the above are "slot-associated" more than "player-associated" properties, and so should stay with the slot)

    ' Swap certain ingame player-associated entries (PingTimes, LagCounter, DesyncCounter, VerifiedIdentity,
    ' JoiningInProgress, PendingDisconnect, DataIntegrity, hostChatPermissions, joinTimes, lastReadyTimes,
    ' lastNotReadyTimes, secondsNotReady, playerLeftGameTime, lastSentPlayerDataCheck2, muteChat)
    netGame.ingameSwapPlayerEntries playerIndexA, playerIndexB
    netGame.playerSpamMuteNotifyIndexSwap playerIndexA, playerIndexB
    netGame.multiSyncPlayerSwap playerIndexA, playerIndexB

    ' Ensure we filter messages appropriately waiting for the client ack *at each new index*
    If NetPlay.players(playerIndexA).allocated Then
        has_NET_waitingForIndexChangeAckSince(playerIndexA) = True
        NET_waitingForIndexChangeAckSince(playerIndexA) = netGame.realTime()
    End If
    If NetPlay.players(playerIndexB).allocated Then
        has_NET_waitingForIndexChangeAckSince(playerIndexB) = True
        NET_waitingForIndexChangeAckSince(playerIndexB) = netGame.realTime()
    End If

    ' Fix up any AI players - if they get moved to a spectator slot, get rid of the AI
    For k = 0 To 1
        playerIndex = playerIndexes(k)
        If Not NetPlay.players(playerIndex).allocated And NetPlay.players(playerIndex).ai >= 0 _
           And (NetPlay.players(playerIndex).isSpectator Or playerIndex > MAX_PLAYERS) Then
            ' remove the AI
            If Not (playerIndex <> PLAYER_FEATURE) Then ASSERT_FAILURE "playerIndex != (11 + 1)", "lib/netplay/netplay.cpp:1916", "swapPlayerIndexes", "Not expecting to see PLAYER_FEATURE here!"
            NetPlay.players(playerIndex).difficulty = AIDifficulty_DISABLED
            NetPlay.players(playerIndex).ai = AI_OPEN
            netGame.clearPlayerName playerIndex
        End If
        If Not NetPlay.players(playerIndex).allocated And NetPlay.players(playerIndex).ai < 0 Then
            If Not (NetPlay.players(playerIndex).difficulty = AIDifficulty_DISABLED) Then ASSERT_FAILURE "NetPlay.players[playerIndex].difficulty == AIDifficulty::DISABLED", _
                "lib/netplay/netplay.cpp:1918", "swapPlayerIndexes", "Unexpected difficulty (%d) for player (%d)", NetPlay.players(playerIndex).difficulty, playerIndex
        End If

        If NetPlay.players(playerIndex).allocated And playerIndex < MAX_PLAYERS Then
            If NetPlay.players(playerIndex).difficulty <> AIDifficulty_HUMAN Then
                debug_ 1931, LOG_INFO, "swapPlayerIndexes", "Fixing up human difficulty for player: %d", playerIndex
                NetPlay.players(playerIndex).difficulty = AIDifficulty_HUMAN
            End If
        End If
    Next

    If Not NetPlay.players(playerIndexA).allocated Or Not NetPlay.players(playerIndexB).allocated Then
        ' Recalculate NetPlay.playercount (etc), as it may have changed
        NETfixPlayerCount
    End If

    swapPlayerIndexes = True
End Function

Private Function NET_isSpectatorOnlySlot(ByVal playerIdx As Long) As Boolean
    If Not ULT(playerIdx, NetPlay.playersSize) Then ASSERT_FAILURE "playerIdx < NetPlay.players.size()", "lib/netplay/netplay.cpp:1948", "_NET_isSpectatorOnlySlot", "Invalid playerIdx: %u", playerIdx: NET_isSpectatorOnlySlot = False: Exit Function
    NET_isSpectatorOnlySlot = playerIdx >= MAX_PLAYERS Or NetPlay.players(playerIdx).position >= netGame.gameMaxPlayers()
End Function

' optional<uint32_t>: the result is has_value(), the slot in result
Private Function NET_findSpectatorSlotToOpen(result As Long) As Boolean
    Dim i As Long
    For i = MAX_PLAYER_SLOTS To MAX_CONNECTED_PLAYERS - 1
        If Not NET_isSpectatorOnlySlot(i) Then
            GoTo continue_
        End If
        If NetPlay.players(i).allocated Or NetPlay.players(i).isSpectator Then
            GoTo continue_
        End If
        If netGame.gameMapHasScavengers() And NetPlay.players(i).position = netGame.scavengerSlot() Then
            GoTo continue_ ' skip it
        End If
        If i = PLAYER_FEATURE Then
            GoTo continue_ ' skip it
        End If

        result = i
        NET_findSpectatorSlotToOpen = True
        Exit Function
continue_:
    Next
    NET_findSpectatorSlotToOpen = False
End Function

Public Function NETcanOpenNewSpectatorSlot() As Boolean
    Dim i As Long
    NETcanOpenNewSpectatorSlot = NET_findSpectatorSlotToOpen(i)
End Function

' _NET_openNewSpectatorSlot_internal(bool broadcastUpdate): optional<uint32_t> (the result is has_value(), the slot in result)
Public Function NET_openNewSpectatorSlot_internal(ByVal broadcastUpdate As Boolean, result As Long) As Boolean
    Dim i As Long
    If Not NetPlay.isHost Then ASSERT_FAILURE "false", "lib/netplay/netplay.cpp:1985", "_NET_openNewSpectatorSlot_internal", "Host only routine detected for client!": NET_openNewSpectatorSlot_internal = False: Exit Function
    If Not NET_findSpectatorSlotToOpen(i) Then
        NET_openNewSpectatorSlot_internal = False
        Exit Function
    End If

    NetPlay.players(i).ai = AI_OPEN
    NetPlay.players(i).isSpectator = True
    ' common code
    NetPlay.players(i).difficulty = AIDifficulty_DISABLED ' disable AI for this slot
    If broadcastUpdate Then
        NETBroadcastPlayerInfo i
        netPlayersUpdated = True
    End If
    result = i
    NET_openNewSpectatorSlot_internal = True
End Function

Public Function NETopenNewSpectatorSlot() As Boolean
    Dim i As Long
    If Not NetPlay.isHost Then ASSERT_FAILURE "false", "lib/netplay/netplay.cpp:2007", "NETopenNewSpectatorSlot", "Host only routine detected for client!": NETopenNewSpectatorSlot = False: Exit Function
    NETopenNewSpectatorSlot = NET_openNewSpectatorSlot_internal(True, i)
End Function

Public Function NETmovePlayerToSpectatorOnlySlot(ByVal playerIdx As Long, Optional ByVal hostOverride As Boolean = False) As Boolean
    Dim availableSpectatorIndex As Long, hasAvailable As Boolean, playerPublicKeyIdentity() As Byte, keyN As Long
    Dim newSpecIdx As Long, outputKey() As Byte, outN As Long, outputIdentity As EcKey, playerPublicKeyB64 As String
    Dim playerIdentityHash As String, playerVerifiedStatus As String, playerName As String, playerNameB64 As String, s As String
    If Not NetPlay.isHost Then ASSERT_FAILURE "false", "lib/netplay/netplay.cpp:2013", "NETmovePlayerToSpectatorOnlySlot", "Host only routine detected for client!": NETmovePlayerToSpectatorOnlySlot = False: Exit Function
    If Not ULT(playerIdx, MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "playerIdx < MAX_CONNECTED_PLAYERS", "lib/netplay/netplay.cpp:2014", "NETmovePlayerToSpectatorOnlySlot", "playerIdx out of bounds: %u", playerIdx: NETmovePlayerToSpectatorOnlySlot = False: Exit Function
    ' Verify it's a human player
    If Not (netGame.isHumanPlayer(playerIdx) And Not NetPlay.players(playerIdx).isSpectator) Then ASSERT_FAILURE "isHumanPlayer(playerIdx) && !NetPlay.players[playerIdx].isSpectator", _
        "lib/netplay/netplay.cpp:2016", "NETmovePlayerToSpectatorOnlySlot", "playerIdx is not a currently-connected human player: %u", playerIdx: NETmovePlayerToSpectatorOnlySlot = False: Exit Function

    ' Try to grab a new spectator-only slot index
    hasAvailable = NET_FindOpenSlotForPlayer(availableSpectatorIndex, False, True, True)
    If Not hasAvailable And hostOverride Then
        ' Attempt to open a new spectator slot just for this player
        hasAvailable = NET_openNewSpectatorSlot_internal(False, availableSpectatorIndex)
    End If
    If Not hasAvailable Then
        debug_ 2027, LOG_ERROR, "NETmovePlayerToSpectatorOnlySlot", "No available spectator slots to move player %u to", playerIdx
        NETmovePlayerToSpectatorOnlySlot = False
        Exit Function
    End If

    ' Backup the player's identity for later recording
    keyN = netGame.getTruePlayerIdentity(playerIdx, playerPublicKeyIdentity)

    ' Swap the player indexes
    If Not swapPlayerIndexes(playerIdx, availableSpectatorIndex) Then
        debug_ 2037, LOG_ERROR, "NETmovePlayerToSpectatorOnlySlot", "Failed to swap player indexes: %u, %u", playerIdx, availableSpectatorIndex
        NETmovePlayerToSpectatorOnlySlot = False
        Exit Function
    End If
    If Not NetPlay.players(availableSpectatorIndex).isSpectator Then ASSERT_FAILURE "NetPlay.players[availableSpectatorIndex.value()].isSpectator", "lib/netplay/netplay.cpp:2040", "NETmovePlayerToSpectatorOnlySlot", "New slot doesn't have spectator set??"

    playerManagementRecord_movedPlayerToSpectators NetPlay.players(availableSpectatorIndex).IPtextAddress, playerPublicKeyIdentity, keyN, hostOverride

    If netGame.wzCommandInterfaceEnabled() Then
        newSpecIdx = availableSpectatorIndex
        outN = netGame.getOutputPlayerIdentity(newSpecIdx, outputKey)
        playerPublicKeyB64 = base64Encode(outputKey, outN)
        If outN > 0 Then EcKey_fromBytes outputIdentity, outputKey, outN, EcKey_Public
        playerIdentityHash = EcKey_publicHashString(outputIdentity)
        If netGame.ingameVerifiedIdentity(newSpecIdx) Then playerVerifiedStatus = "V" Else playerVerifiedStatus = "?"
        playerName = netGame.getPlayerName(newSpecIdx)
        playerNameB64 = base64EncodeString(playerName)
        If hostOverride Then s = "host" Else s = "user"
        netGame.wzCommandInterfaceOutput va("WZEVENT: movedPlayerToSpec: %u -> %u %s %s %s %s %s %s" & vbLf, playerIdx, newSpecIdx, playerPublicKeyB64, playerIdentityHash, _
            playerVerifiedStatus, playerNameB64, NetPlay.players(newSpecIdx).IPtextAddress, s)
    End If

    ' Broadcast the swapped player info
    NETBroadcastTwoPlayerInfo playerIdx, availableSpectatorIndex

    NETmovePlayerToSpectatorOnlySlot = True
End Function

Private Function wasAlreadyMovedToSpectatorsByHost(ByVal playerIdx As Long) As Boolean
    Dim key() As Byte, n As Long
    If playerManagementRecord_hostMovedPlayerToSpectators_ip(NetPlay.players(playerIdx).IPtextAddress) Then
        wasAlreadyMovedToSpectatorsByHost = True
        Exit Function
    End If
    n = netGame.getTruePlayerIdentity(playerIdx, key)
    wasAlreadyMovedToSpectatorsByHost = playerManagementRecord_hostMovedPlayerToSpectators_key(key, n)
End Function

' optional<uint32_t> newPlayerIdx: hasNewPlayerIdx
Public Function NETmoveSpectatorToPlayerSlot(ByVal playerIdx As Long, ByVal hasNewPlayerIdx As Boolean, ByVal newPlayerIdx As Long, Optional ByVal hostOverride As Boolean = False) As Long
    Dim previousPlayersShouldCheckReadyValue As Boolean, spectatorPublicKeyIdentity() As Byte, keyN As Long, newIdxValueOr0 As Long
    Dim identityKey() As Byte, idN As Long, identity As EcKey, playerPublicKeyB64 As String, playerIdentityHash As String
    Dim playerVerifiedStatus As String, playerName As String, playerNameB64 As String
    If Not NetPlay.isHost Then ASSERT_FAILURE "false", "lib/netplay/netplay.cpp:2070", "NETmoveSpectatorToPlayerSlot", "Host only routine detected for client!": NETmoveSpectatorToPlayerSlot = SpectatorToPlayerMoveResult_FAILED: Exit Function
    If Not ULT(playerIdx, MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "playerIdx < MAX_CONNECTED_PLAYERS", "lib/netplay/netplay.cpp:2071", "NETmoveSpectatorToPlayerSlot", "playerIdx out of bounds: %u", playerIdx: NETmoveSpectatorToPlayerSlot = SpectatorToPlayerMoveResult_FAILED: Exit Function
    If hasNewPlayerIdx Then newIdxValueOr0 = newPlayerIdx Else newIdxValueOr0 = 0
    If Not ULT(newIdxValueOr0, MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "newPlayerIdx.value_or(0) < MAX_CONNECTED_PLAYERS", "lib/netplay/netplay.cpp:2072", "NETmoveSpectatorToPlayerSlot", "newPlayerIdx out of bounds: %u", newIdxValueOr0: NETmoveSpectatorToPlayerSlot = SpectatorToPlayerMoveResult_FAILED: Exit Function
    If hasNewPlayerIdx Then
        If newPlayerIdx = NetPlay.hostPlayer Then ASSERT_FAILURE "!newPlayerIdx.has_value() || newPlayerIdx.value() != NetPlay.hostPlayer", "lib/netplay/netplay.cpp:2073", "NETmoveSpectatorToPlayerSlot", "newPlayerIdx cannot be host player: %u", newIdxValueOr0: NETmoveSpectatorToPlayerSlot = SpectatorToPlayerMoveResult_FAILED: Exit Function
    End If
    ' Verify it's a human player
    If Not (netGame.isHumanPlayer(playerIdx) And NetPlay.players(playerIdx).isSpectator) Then ASSERT_FAILURE "isHumanPlayer(playerIdx) && NetPlay.players[playerIdx].isSpectator", _
        "lib/netplay/netplay.cpp:2075", "NETmoveSpectatorToPlayerSlot", "playerIdx is not a currently-connected spectator: %u", playerIdx: NETmoveSpectatorToPlayerSlot = SpectatorToPlayerMoveResult_FAILED: Exit Function

    previousPlayersShouldCheckReadyValue = netGame.multiplayPlayersShouldCheckReady()

    ' Check if the host has moved this player to a spectator before, if so deny (unless the *host* is the one triggering the move)
    If Not hostOverride And wasAlreadyMovedToSpectatorsByHost(playerIdx) Then
        NETmoveSpectatorToPlayerSlot = SpectatorToPlayerMoveResult_FAILED
        Exit Function
    End If

    If Not hasNewPlayerIdx Then
        ' Attempt to find an open slot
        hasNewPlayerIdx = NET_FindOpenSlotForPlayer(newPlayerIdx, False, True, False)
        If Not hasNewPlayerIdx Then
            NETmoveSpectatorToPlayerSlot = SpectatorToPlayerMoveResult_NEEDS_SLOT_SELECTION
            Exit Function
        End If
    End If

    ' Backup the spectator's identity for later recording
    keyN = netGame.getTruePlayerIdentity(playerIdx, spectatorPublicKeyIdentity)

    ' Swap the player indexes
    If Not swapPlayerIndexes(playerIdx, newPlayerIdx) Then
        debug_ 2101, LOG_ERROR, "NETmoveSpectatorToPlayerSlot", "Failed to swap player indexes: %u, %u", playerIdx, newPlayerIdx
        NETmoveSpectatorToPlayerSlot = SpectatorToPlayerMoveResult_FAILED
        Exit Function
    End If
    If NetPlay.players(newPlayerIdx).isSpectator Then ASSERT_FAILURE "!NetPlay.players[newPlayerIdx.value()].isSpectator", "lib/netplay/netplay.cpp:2104", "NETmoveSpectatorToPlayerSlot", "New slot should not be a spectator??"

    playerManagementRecord_movedSpectatorToPlayers NetPlay.players(newPlayerIdx).IPtextAddress, spectatorPublicKeyIdentity, keyN, hostOverride

    If netGame.wzCommandInterfaceEnabled() Then
        idN = netGame.getOutputPlayerIdentity(newPlayerIdx, identityKey)
        playerPublicKeyB64 = base64Encode(identityKey, idN)
        If idN > 0 Then EcKey_fromBytes identity, identityKey, idN, EcKey_Public
        playerIdentityHash = EcKey_publicHashString(identity)
        If netGame.ingameVerifiedIdentity(newPlayerIdx) Then playerVerifiedStatus = "V" Else playerVerifiedStatus = "?"
        playerName = netGame.getPlayerName(newPlayerIdx)
        playerNameB64 = base64EncodeString(playerName)
        netGame.wzCommandInterfaceOutput va("WZEVENT: movedSpecToPlayer: %u -> %u %s %s %s %s %s" & vbLf, playerIdx, newPlayerIdx, playerPublicKeyB64, playerIdentityHash, _
            playerVerifiedStatus, playerNameB64, NetPlay.players(newPlayerIdx).IPtextAddress)
    End If

    ' Broadcast the swapped player info
    NETBroadcastTwoPlayerInfo playerIdx, newPlayerIdx

    netGame.handlePossiblePlayersShouldCheckReadyChange previousPlayersShouldCheckReadyValue

    NETmoveSpectatorToPlayerSlot = SpectatorToPlayerMoveResult_SUCCESS
End Function

Private Function NETFilterMessageWhileSwappingPlayer(ByVal sender As Long, ByVal typ As Long) As Boolean
    Dim msg As String
    If Not NetPlay.isHost Then
        NETFilterMessageWhileSwappingPlayer = False ' only host filters these messages
        Exit Function
    End If

    If sender = NetPlay.hostPlayer Then
        NETFilterMessageWhileSwappingPlayer = False ' never filter host messages
        Exit Function
    End If

    If sender >= MAX_CONNECTED_PLAYERS Then
        NETFilterMessageWhileSwappingPlayer = False ' no filtering - not waiting for player to acknowledge index change
        Exit Function
    End If
    If Not has_NET_waitingForIndexChangeAckSince(sender) Then
        NETFilterMessageWhileSwappingPlayer = False ' no filtering - not waiting for player to acknowledge index change
        Exit Function
    End If

    If ULT(INDEX_CHANGE_ACK_TIMEOUT, netGame.realTime() - NET_waitingForIndexChangeAckSince(sender)) Then
        ' this client did not acknowledge the player index change before the timeout - kick them
        If NETplayerHasConnection(sender) Or NetPlay.players(sender).allocated Then
            msg = sstrcpyN(va("Auto-kicking player %u, did not ack player index change within required timeframe.", sender), 256)
            netGame.sendInGameSystemMessage msg
            debug_ 2153, LOG_INFO, "NETFilterMessageWhileSwappingPlayer", "Client (player: %u) failed to ack player index swap (ignoring message type: %u)", sender, typ
            netGame.kickPlayer sender, "Client failed to ack player index swap", ERROR_INVALID, False
        End If
        NETFilterMessageWhileSwappingPlayer = True ' filter original message, of course
        Exit Function
    End If

    ' Filter certain net messages if player index swap is in progress (until player acknowledges swap)

    Select Case typ
    Case NET_TEXTMSG, NET_QUICK_CHAT_MSG
        ' Just send a message to the player that this text message was undelivered and to try again - it's easier and this should be quite rare
        netGame.sendQuickChat sender            ' WzQuickChatMessage::INTERNAL_MSG_DELIVERY_FAILURE_TRY_AGAIN from selectedPlayer to sender
        NETFilterMessageWhileSwappingPlayer = True ' filter / ignore
        Exit Function
    Case NET_VOTE
        ' POSSIBLE TODO: Send the player a new VOTE_REQUEST so they can send a new vote? (after they ack the swap)
        NETFilterMessageWhileSwappingPlayer = True ' filter / ignore
        Exit Function
    Case NET_PLAYERNAME_CHANGEREQUEST       ' non-host human player is changing their name.
        ' We can ignore if we force the player to re-send their name when they change slots
        NETFilterMessageWhileSwappingPlayer = True ' filter / ignore
        Exit Function

    ' Client messages to ignore while waiting for a swap player ACK
    ' (because they may contain the old player index, and also the necessary messages will be resent after the ack)
    Case NET_PING, NET_PLAYER_STATS, NET_COLOURREQUEST, NET_FACTIONREQUEST, NET_TEAMREQUEST, NET_READY_REQUEST, NET_POSITIONREQUEST, NET_DATA_CHECK, NET_DATA_CHECK2
        NETFilterMessageWhileSwappingPlayer = True ' filter / ignore
        Exit Function

    ' one slot / index change at a time...
    Case NET_PLAYER_SLOTTYPE_REQUEST
        NETFilterMessageWhileSwappingPlayer = True ' filter / ignore
        Exit Function

    ' client messages to be processed normally
    Case NET_FILE_REQUESTED, NET_FILE_CANCELLED
        NETFilterMessageWhileSwappingPlayer = False ' process normally (do *not* filter)
        Exit Function

    ' host-only messages
    Case NET_OPTIONS, NET_KICK, NET_FIREUP, NET_PLAYER_INFO, NET_PLAYER_JOINED, NET_PLAYER_LEAVING, NET_PLAYER_DROPPED, NET_GAME_FLAGS, _
         NET_HOST_DROPPED, NET_FILE_PAYLOAD, NET_VOTE_REQUEST, NET_PLAYER_SWAP_INDEX, NET_HOST_CONFIG
        ASSERT_FAILURE "false", "lib/netplay/netplay.cpp:2215", "NETFilterMessageWhileSwappingPlayer", "Received unexpected host-only message (%u) from sender: %u", typ, sender

    ' only possible with initial join
    Case NET_JOIN, NET_ACCEPTED, NET_REJECTED
        ASSERT_FAILURE "false", "lib/netplay/netplay.cpp:2222", "NETFilterMessageWhileSwappingPlayer", "Received unexpected initial-join message (%u) from sender: %u", typ, sender

    ' should only be possible once game has started
    Case NET_PLAYERRESPONDING, NET_AITEXTMSG, NET_BEACONMSG, NET_SHARE_GAME_QUEUE, NET_DEBUG_SYNC, NET_SPECTEXTMSG, NET_TEAM_STRATEGY

    ' just process normally
    Case NET_SEND_TO_PLAYER
        NETFilterMessageWhileSwappingPlayer = False ' process normally (do *not* filter)
        Exit Function

    Case NET_SECURED_NET_MESSAGE        ' This just wraps another message, so permit this wrapper message
        NETFilterMessageWhileSwappingPlayer = False ' process normally (do *not* filter)
        Exit Function

    Case NET_PLAYER_SWAP_INDEX_ACK
        ' **MUST** be permitted
        NETFilterMessageWhileSwappingPlayer = False ' process normally (do *not* filter)
        Exit Function

    Case Else
        ' just process normally
    End Select

    NETFilterMessageWhileSwappingPlayer = False
End Function

' (the message's sender and text: NetworkTextMessage's; senderPublicKeyB64 is optional: hasSenderPublicKeyB64)
Private Sub cmdInterfaceLogChatMsgInternal(ByVal messageSender As Long, ByVal messageText As String, ByVal log_prefix As String, _
        Optional ByVal hasSenderPublicKeyB64 As Boolean = False, Optional ByVal senderPublicKeyB64_ As String = "")
    Dim key() As Byte, n As Long, senderPublicKeyB64 As String, senderVerifiedStatus As String, sendername As String
    Dim sendername64 As String, messagetext64 As String
    If Not NetPlay.isHost Or Not netGame.wzCommandInterfaceEnabled() Or (NetPlay.hostPlayer < MAX_PLAYER_SLOTS) Then
        Exit Sub
    End If

    If messageSender < 0 Then
        ' for now, skip system messages
        Exit Sub
    End If

    If Not (messageSender < MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "message.sender < MAX_CONNECTED_PLAYERS", "lib/netplay/netplay.cpp:2267", "cmdInterfaceLogChatMsgInternal", "Invalid message.sender (%d)", messageSender: Exit Sub

    n = netGame.getOutputPlayerIdentity(messageSender, key)
    If hasSenderPublicKeyB64 Then senderPublicKeyB64 = senderPublicKeyB64_ Else senderPublicKeyB64 = base64Encode(key, n)
    If netGame.ingameVerifiedIdentity(messageSender) Then senderVerifiedStatus = "V" Else senderVerifiedStatus = "?"
    sendername = netGame.getPlayerName(messageSender)
    sendername64 = base64EncodeString(sendername)
    messagetext64 = base64EncodeString(messageText)
    netGame.wzCommandInterfaceOutput va("%s: %i %s %s %s %s %s" & vbLf, log_prefix, messageSender, senderPublicKeyB64, sendername64, messagetext64, senderVerifiedStatus, NetPlay.players(messageSender).IPtextAddress)
End Sub

' receiver may be a player index or NET_ALL_PLAYERS
Private Sub LogChatMsg(ByVal msgType As Long, ByVal message As NetMessage, ByVal senderIdx As Long, ByVal receiver As Long)
    Dim sender As Long, text As String
    If Not netGame.wzCommandInterfaceEnabled() Then
        Exit Sub
    End If

    Select Case msgType
    Case NET_TEXTMSG
        If netGame.NetworkTextMessageDecode(message, senderIdx, sender, text) Then
            ' Process & log
            cmdInterfaceLogChatMsgInternal sender, text, "WZCHAT"
        End If
    Case NET_SPECTEXTMSG
        If netGame.decodeSpecInGameTextMessage(message, senderIdx, sender, text) Then
            ' Process & log
            cmdInterfaceLogChatMsgInternal sender, text, "WZCHAT"
        End If
    End Select
End Sub

'/////////////////////////////////////////////////////////////////////////
' Check if a message is a system message
Public Function NETprocessSystemMessage(ByVal playerQueue As Long, typ As Long) As Boolean
    Dim sender As Long, receiver As Long, message As NetMessage, r As MessageReader, w As MessageWriter, msgType As Long
    Dim msg As String, rawLen As Long, player As Long, num As Long, n As Long, isSentByCorrectClient As Boolean
    Dim oldName As String, newName As String, indexLen As Long, index As Long, colour As Long, position As Long, team As Long
    Dim ai As Long, difficulty As Long, faction As Long, isSpectator As Boolean, isAdmin As Boolean, error_ As Boolean
    Dim wasAllocated As Boolean, oldName2 As String, tmpDiscard As Boolean, newFactionId As Long, s As String
    Dim i As Long, count As Long, oldPlayerIndex As Long, newPlayerIndex As Long, b As Boolean, v As Long, pqIndex As Long
    Dim rejectKick As Boolean

    pqIndex = NETQUEUE_index(playerQueue)

    If typ = NET_SECURED_NET_MESSAGE Then
        If Not NETdecryptSecuredNetMessage(playerQueue, typ) Then
            b = True
        ElseIf typ = NET_SECURED_NET_MESSAGE Then   ' nested secured messages are not supported
            b = True
        Else
            b = False
        End If
        If b Then
            ' remove the invalid secured message, and return true to say we "handled it"
            debug_ 2321, LOG_INFO, "NETprocessSystemMessage", "Ignoring invalid secured message allegedly from player: %u", pqIndex
            NETpop playerQueue
            NETprocessSystemMessage = True
            Exit Function
        End If
    End If

    If NETFilterMessageWhileSwappingPlayer(pqIndex, typ) Then
        NETpop playerQueue
        NETprocessSystemMessage = True
        Exit Function
    End If

    Select Case typ
    Case NET_SEND_TO_PLAYER
        Set message = Nothing
        Set r = NETbeginDecode(playerQueue, NET_SEND_TO_PLAYER)
        NETuint8_t_r r, sender
        NETuint8_t_r r, receiver
        NETnetMessage_r r, message  ' Must delete message later.
        If Not NETend_r(r) Then
            debug_ 2347, LOG_ERROR, "NETprocessSystemMessage", "Incomplete NET_SEND_TO_PLAYER."
            GoTo break_
        End If
        If sender >= MAX_CONNECTED_PLAYERS Or (receiver >= MAX_CONNECTED_PLAYERS And receiver <> NET_ALL_PLAYERS) Then
            debug_ 2352, LOG_ERROR, "NETprocessSystemMessage", "Bad NET_SEND_TO_PLAYER."
            GoTo break_
        End If
        If (receiver = selectedPlayer Or receiver = NET_ALL_PLAYERS) And pqIndex = NetPlay.hostPlayer Then
            ' Message was sent to us via the host.
            If sender <> selectedPlayer Then  ' Make sure host didn't send us our own broadcast messages, which shouldn't happen anyway.
                NETlogPacket message.type_(), message.rawDataSize(), True
                NETinsertMessageFromNet NETnetQueue(sender), message
            End If
        ElseIf NetPlay.isHost And sender = pqIndex Then
            msgType = message.type_()
            rejectKick = False
            If (msgType = NET_OPTIONS Or msgType = NET_FIREUP Or msgType = NET_KICK Or msgType = NET_PLAYER_LEAVING Or msgType = NET_PLAYER_DROPPED _
                Or msgType = NET_REJECTED Or msgType = NET_GAME_FLAGS Or msgType = NET_PLAYER_JOINED Or msgType = NET_PLAYER_INFO _
                Or msgType = NET_FILE_PAYLOAD Or msgType = NET_PLAYER_SWAP_INDEX Or msgType = NET_HOST_CONFIG) And sender <> NetPlay.hostPlayer Then rejectKick = True
            If (msgType = NET_HOST_DROPPED Or msgType = NET_FILE_REQUESTED Or msgType = NET_READY_REQUEST Or msgType = NET_TEAMREQUEST _
                Or msgType = NET_COLOURREQUEST Or msgType = NET_POSITIONREQUEST Or msgType = NET_FACTIONREQUEST Or msgType = NET_FILE_CANCELLED _
                Or msgType = NET_DATA_CHECK Or msgType = NET_JOIN Or msgType = NET_PLAYERNAME_CHANGEREQUEST Or msgType = NET_PLAYER_SWAP_INDEX_ACK) _
                And receiver <> NetPlay.hostPlayer Then rejectKick = True
            If (msgType = NET_PLAYER_SLOTTYPE_REQUEST Or msgType = NET_DATA_CHECK2) And (sender <> NetPlay.hostPlayer And receiver <> NetPlay.hostPlayer) Then rejectKick = True
            If rejectKick Then
                msg = sstrcpyN(va("Auto-kicking player %u, lacked the required access level for command(%d).", sender, message.type_()), 256)
                NETlogEntry msg, SYNC_FLAG, sender

                If NETplayerHasConnection(sender) Then
                    netGame.sendRoomSystemMessage msg
                    debug_ 2404, LOG_ERROR, "NETprocessSystemMessage", "%s", msg
                    netGame.kickPlayer sender, "Invalid command attempted", ERROR_INVALID, True
                End If
                GoTo break_
            End If

            ' Certain messages should be filtered while we're waiting for the ack of a player index switch
            If NETFilterMessageWhileSwappingPlayer(sender, msgType) Then
                GoTo break_
            End If

            If (msgType = NET_SHARE_GAME_QUEUE) And (receiver <> NET_ALL_PLAYERS) Then
                debug_ 2418, LOG_NET, "NETprocessSystemMessage", "Ignoring non-broadcast message type (%d) from Player %d", message.type_(), pqIndex
                GoTo break_
            End If

            ' Certain messages should be filtered due to hostChatPermissions
            If msgType = NET_TEXTMSG Or msgType = NET_SPECTEXTMSG Or msgType = NET_AITEXTMSG Then
                If Not netGame.ingameHostChatPermissions(sender) Then
                    ' Only allow messages direct to host in this case! (Carve-out to allow /hostmsg commands...)
                    If receiver <> NetPlay.hostPlayer Then
                        GoTo break_
                    End If
                End If

                LogChatMsg msgType, message, sender, receiver
            End If

            rawLen = message.rawDataSize()

            ' We are the host, and player is asking us to send the message to receiver.
            Set w = NETbeginEncode(NETnetQueue(receiver, sender), NET_SEND_TO_PLAYER)
            NETuint8_t_w w, sender
            NETuint8_t_w w, receiver
            NETnetMessage_w w, message
            NETend_w w

            If receiver = NET_ALL_PLAYERS Then
                NETlogPacket msgType, rawLen, True
                NETinsertMessageFromNet NETnetQueue(sender), message  ' Message is also for the host.
                ' Not sure if flushing here can make a difference, maybe it can:
                'NETflush();  // Send the message to everyone as fast as possible.
            End If
        Else
            If NetPlay.isHost Then
                If NETFilterMessageWhileSwappingPlayer(pqIndex, message.type_()) Then
                    debug_ 2458, LOG_NET, "NETprocessSystemMessage", "Ignoring message type (%d) from Player %d while swapping player index", message.type_(), pqIndex
                    GoTo break_
                End If
            End If

            debug_ 2462, LOG_INFO, "NETprocessSystemMessage", "Report this: Player %d sent us message type (%d) addressed to %d from %d. We are %d.", pqIndex, message.type_(), receiver, sender, selectedPlayer
        End If

    Case NET_SHARE_GAME_QUEUE
        If netGame.ingameLocalJoiningInProgress() Then
            ' NET_SHARE_GAME_QUEUE should only be processed after the game has started - the game queues probably aren't yet created!
            debug_ 2472, LOG_ERROR, "NETprocessSystemMessage", "Ignoring NET_SHARE_GAME_QUEUE message from %u (only permitted in-game).", pqIndex
            GoTo break_
        End If

        player = 0
        num = 0
        Set message = Nothing

        ' Encoded in NETprocessSystemMessage in nettypes.cpp.
        Set r = NETbeginDecode(playerQueue, NET_SHARE_GAME_QUEUE)
        NETuint8_t_r r, player
        NETuint32_t_r r, num
        isSentByCorrectClient = netGame.responsibleFor(pqIndex, player)
        isSentByCorrectClient = isSentByCorrectClient Or (pqIndex = NetPlay.hostPlayer And pqIndex <> selectedPlayer)  ' Let host spoof other people's NET_SHARE_GAME_QUEUE messages, but not our own. This allows the host to spoof a GAME_PLAYER_LEFT message (but spoofing any message when the player is still there will fail with desynch).
        If Not isSentByCorrectClient Or player >= MAX_CONNECTED_PLAYERS Then
            NETend_r r
            GoTo break_
        End If
        For n = 0 To num - 1
            NETnetMessage_r r, message

            NETlogPacket message.type_(), message.rawDataSize(), True
            NETinsertMessageFromNet NETgameQueue(player), message

            Set message = Nothing
        Next
        If Not NETend_r(r) Then
            debug_ 2503, LOG_ERROR, "NETprocessSystemMessage", "Bad NET_SHARE_GAME_QUEUE message."
            GoTo break_
        End If

    Case NET_PLAYERNAME_CHANGEREQUEST
        If Not netGame.ingameLocalJoiningInProgress() Then
            ' player name change only permitted in the lobby
            debug_ 2513, LOG_ERROR, "NETprocessSystemMessage", "Ignoring NET_PLAYERNAME_CHANGEREQUEST message (only permitted in the lobby)."
            GoTo break_
        End If
        If Not NetPlay.isHost Then
            ' Only the host should receive and process these messages
            debug_ 2519, LOG_ERROR, "NETprocessSystemMessage", "Ignoring NET_PLAYERNAME_CHANGEREQUEST message (should only be sent to the host)."
            GoTo break_
        End If

        player = 0
        oldName = ""
        newName = ""

        ' Encoded in NETchangePlayerName() in netplay.cpp.
        Set r = NETbeginDecode(playerQueue, NET_PLAYERNAME_CHANGEREQUEST)
        NETuint8_t_r r, player
        NETwzstring_r r, newName
        NETend_r r

        ' Bail out if the given ID number is out of range
        b = player >= MAX_CONNECTED_PLAYERS
        If Not b Then
            If pqIndex <> NetPlay.hostPlayer Then b = (pqIndex <> player Or Not NetPlay.players(player).allocated)
        End If
        If b Then
            debug_ 2536, LOG_ERROR, "NETprocessSystemMessage", "NET_PLAYERNAME_CHANGEREQUEST from %u: Player ID (%u) out of range (max %u)", pqIndex, player, MAX_CONNECTED_PLAYERS
            GoTo break_
        End If

        If NetPlay.isHost And (Not netGame.ingameHostChatPermissions(player) Or netGame.getLockedOptionsName()) Then
            ' Name changes are denied when the host has muted a player (or name changes are locked)
            ' Inform the player that their name is still what it was (resets their local display)
            If NetPlay.players(player).allocated Then
                NETSendPlayerInfoTo player, player
            End If
            GoTo break_
        End If

        oldName = NetPlay.players(player).name
        netGame.setPlayerName player, newName

        If NetPlay.players(player).allocated And oldName <> NetPlay.players(player).name Then
            netGame.printConsoleNameChange oldName, NetPlay.players(player).name
            ' Send the updated data to all other clients as well.
            NETBroadcastPlayerInfo player ' ultimately triggers updateMultiplayGameData inside NETSendNPlayerInfoTo
            NETfixDuplicatePlayerNames
            netPlayersUpdated = True
        End If

    Case NET_PLAYER_STATS
        netGame.recvMultiStats playerQueue
        netPlayersUpdated = True

    Case NET_PLAYER_INFO
        indexLen = 0
        index = MAX_CONNECTED_PLAYERS
        colour = 0
        position = 0
        team = 0
        ai = 0
        difficulty = 0
        faction = FACTION_NORMAL
        isSpectator = False
        isAdmin = False
        error_ = False

        Set r = NETbeginDecode(playerQueue, NET_PLAYER_INFO)
        NETuint32_t_r r, indexLen
        If Not ULT(MAX_CONNECTED_PLAYERS, indexLen) Then b = (pqIndex <> NetPlay.hostPlayer) Else b = True
        If b Then
            debug_ 2589, LOG_ERROR, "NETprocessSystemMessage", "MSG_PLAYER_INFO: Bad number of players updated: %u", indexLen
            NETend_r r
            GoTo break_
        End If

        For n = 0 To indexLen - 1
            wasAllocated = False
            oldName2 = ""

            ' Retrieve the player's ID
            NETuint32_t_r r, index

            ' Bail out if the given ID number is out of range
            b = Not ULT(index, MAX_CONNECTED_PLAYERS)
            If Not b Then
                If pqIndex <> NetPlay.hostPlayer Then b = (pqIndex <> index Or Not NetPlay.players(index).allocated)
            End If
            If b Then
                debug_ 2605, LOG_ERROR, "NETprocessSystemMessage", "MSG_PLAYER_INFO from %u: Player ID (%u) out of range (max %u)", pqIndex, index, MAX_CONNECTED_PLAYERS
                error_ = True
                Exit For
            End If

            ' Retrieve the rest of the data
            wasAllocated = NetPlay.players(index).allocated
            NETbool_r r, NetPlay.players(index).allocated
            NETbool_r r, NetPlay.players(index).heartbeat
            tmpDiscard = False
            NETbool_r r, tmpDiscard ' to maintain message format, discard old "kick" variable // FUTURE TODO: Remove
            oldName2 = NetPlay.players(index).name
            NETstring_r r, NetPlay.players(index).name, StringSize
            NETuint32_t_r r, NetPlay.players(index).heartattacktime
            NETint32_t_r r, colour
            NETint32_t_r r, position
            NETint32_t_r r, team
            NETbool_r r, NetPlay.players(index).ready
            NETint8_t_r r, ai
            NETint8_t_r r, difficulty
            NETuint8_t_r r, faction
            NETbool_r r, isSpectator
            NETbool_r r, isAdmin

            If Not uintToFactionID(faction, newFactionId) Then
                debug_ 2633, LOG_ERROR, "NETprocessSystemMessage", "MSG_PLAYER_INFO from %u: Faction ID (%u) out of range (max %u)", pqIndex, faction, MAX_FACTION_ID
                error_ = True
                Exit For
            End If

            ' Don't let anyone except the host change these, otherwise it will end up inconsistent at some point, and the game gets really messed up.
            If pqIndex = NetPlay.hostPlayer Then
                netGame.setPlayerColour index, colour
                NetPlay.players(index).position = position
                NetPlay.players(index).team = team
                NetPlay.players(index).ai = ai
                NetPlay.players(index).difficulty = difficulty
                NetPlay.players(index).faction = newFactionId
                NetPlay.players(index).isSpectator = isSpectator
                NetPlay.players(index).isAdmin = isAdmin
            End If

            If n = 0 Then s = "Receiving MSG_PLAYER_INFO" Else s = "                      and"
            If NetPlay.players(index).allocated Then
                debug_ 2651, LOG_NET, "NETprocessSystemMessage", "%s for player %u (%s)", s, index, "human"
            Else
                debug_ 2651, LOG_NET, "NETprocessSystemMessage", "%s for player %u (%s)", s, index, "AI"
            End If

            If wasAllocated And NetPlay.players(index).allocated And oldName2 <> NetPlay.players(index).name Then
                netGame.printConsoleNameChange oldName2, NetPlay.players(index).name
            End If
        Next
        NETend_r r
        ' If we're the game host make sure to send the updated
        ' data to all other clients as well.
        If NetPlay.isHost And Not error_ Then
            NETBroadcastPlayerInfo index ' ultimately triggers updateMultiplayGameData inside NETSendNPlayerInfoTo
            NETfixDuplicatePlayerNames
        ElseIf Not error_ Then
            If index = selectedPlayer Then
                netGame.handleAutoReadyRequest
            End If
            ' (ActivityManager::instance().updateMultiplayGameData(game, ingame, NETGameIsLocked()): the activity sinks, not ported)
        End If
        netPlayersUpdated = True

    Case NET_PLAYER_JOINED
        Set r = NETbeginDecode(playerQueue, NET_PLAYER_JOINED)
        NETuint8_t_r r, index
        NETend_r r

        debug_ 2685, LOG_NET, "NETprocessSystemMessage", "Receiving NET_PLAYER_JOINED for player %u using socket %p", index, ObjPtr(bsocket)

        Dim noKey() As Byte
        netGame.MultiPlayerJoin index, False, noKey, 0
        netPlayersUpdated = True

    ' This message type is when player is leaving 'nicely', and socket is still valid.
    Case NET_PLAYER_LEAVING
        Set r = NETbeginDecode(playerQueue, NET_PLAYER_LEAVING)
        NETuint32_t_r r, index
        NETend_r r

        If pqIndex <> NetPlay.hostPlayer And index <> pqIndex Then
            debug_ 2703, LOG_ERROR, "NETprocessSystemMessage", "Player %d left, but accidentally set player %d as leaving.", pqIndex, index
            index = pqIndex
        End If

        If Not ULT(index, MAX_CONNECTED_PLAYERS) Then
            debug_ 2714, LOG_NET, "NETprocessSystemMessage", "Receiving NET_PLAYER_LEAVING for player %u (no socket?)", index
        ElseIf Not connected_bsocket(index) Is Nothing Then
            debug_ 2709, LOG_NET, "NETprocessSystemMessage", "Receiving NET_PLAYER_LEAVING for player %u on socket %p", index, ObjPtr(connected_bsocket(index))
        Else
            ' dropped from join screen most likely
            debug_ 2714, LOG_NET, "NETprocessSystemMessage", "Receiving NET_PLAYER_LEAVING for player %u (no socket?)", index
        End If

        If NetPlay.isHost Then
            debug_ 2719, LOG_NET, "NETprocessSystemMessage", "Broadcast leaving message to everyone else"
            Set w = NETbeginEncode(NETbroadcastQueue(), NET_PLAYER_LEAVING)
            NETuint32_t_w w, index
            NETend_w w
        End If

        debug_ 2725, LOG_INFO, "NETprocessSystemMessage", "Player %u has left the game.", index
        NETplayerLeaving index      ' need to close socket for the player that left.
        NETsetPlayerConnectionStatus CONNECTIONSTATUS_PLAYER_LEAVING, index

    Case NET_GAME_FLAGS
        debug_ 2732, LOG_NET, "NETprocessSystemMessage", "Receiving game flags"

        Set r = NETbeginDecode(playerQueue, NET_GAME_FLAGS)

        If pqIndex <> NetPlay.hostPlayer Then
            debug_ 2738, LOG_ERROR, "NETprocessSystemMessage", "NET_GAME_FLAGS sent by wrong player: %u", pqIndex
            NETend_r r
            GoTo break_
        End If

        ' static unsigned int max_flags = ARRAY_SIZE(NetGameFlags);
        ' Retrieve the amount of game flags that we should receive
        NETuint8_t_r r, count

        ' Make sure that we won't get buffer overflows by checking that we
        ' have enough space to store the given amount of game flags.
        If count > 4 Then
            debug_ 2753, LOG_NET, "NETprocessSystemMessage", "NET_GAME_FLAGS: More game flags sent (%u) than our buffer can hold (%u)", count, 4
            count = 4
        End If

        ' Retrieve all game flags
        For i = 0 To count - 1
            NETint32_t_r r, NetGameFlags(i)
        Next
        NETend_r r

        If NetPlay.isHost Then
            NETsendGameFlags
        End If

    Case NET_DEBUG_SYNC
        recvDebugSync playerQueue

    Case NET_PLAYER_SWAP_INDEX_ACK
        If Not NetPlay.isHost Then ASSERT_FAILURE "false", "lib/netplay/netplay.cpp:2778", "NETprocessSystemMessage", "Host only routine detected for client!": GoTo break_

        Set r = NETbeginDecode(playerQueue, NET_PLAYER_SWAP_INDEX_ACK)
        NETuint32_t_r r, oldPlayerIndex
        NETuint32_t_r r, newPlayerIndex
        NETend_r r

        isSentByCorrectClient = netGame.responsibleFor(pqIndex, newPlayerIndex)
        If Not isSentByCorrectClient Then
            debug_ 2791, LOG_ERROR, "NETprocessSystemMessage", "NET_PLAYER_SWAP_INDEX_ACK sent by wrong player: %u", pqIndex
            GoTo break_
        End If

        b = Not ULT(newPlayerIndex, MAX_CONNECTED_PLAYERS)
        If Not b Then b = Not has_NET_waitingForIndexChangeAckSince(newPlayerIndex)
        If b Then
            debug_ 2797, LOG_ERROR, "NETprocessSystemMessage", "NET_PLAYER_SWAP_INDEX_ACK sent despite us not waiting for it (for player: %u)", newPlayerIndex
            GoTo break_
        End If

        has_NET_waitingForIndexChangeAckSince(newPlayerIndex) = False

    Case Else
        NETprocessSystemMessage = False
        Exit Function
    End Select

break_:
    NETpop playerQueue
    NETprocessSystemMessage = True
End Function

' FactionID uintToFactionID(uint8_t value): optional (the result is has_value(), the id in factionId)
Private Function uintToFactionID(ByVal value As Long, factionId As Long) As Boolean
    If value > MAX_FACTION_ID Then
        uintToFactionID = False
        Exit Function
    End If
    factionId = value
    uintToFactionID = True
End Function
