Attribute VB_Name = "netplay_conn_c"
' lib/netplay/netplay.cpp (2 of 4): the connections. Reading (NET_fillBuffer, NETrecvNet, NETrecvGame), sending
' (NETsend, NETflush and the delayed disconnects), the players leaving and dropping, NETclose, the host's side
' of joining (NETacceptIncomingConnections, NETallowJoining with the temporary sockets' states, the session
' bans, NEThostGame, NEThaltJoining) and the client's (NETpromoteJoinAttemptToEstablishedConnectionToHost).
' The connections are the IClientConnection / IListenSocket / IConnectionPollGroup objects of the connection
' layer; a net::result's error is an ec Long (net_errorValue, net_errorCategory, net_errorMessage; ec == std::errc::x
' is net_errorEqualsErrc).
' The pending writes go out from PendingWritesManager_pollAll (the C++'s writer thread), run here at NETrecvNet,
' NETrecvGame and NETflush.
' tmp_pendingIPs (std::unordered_map<std::string, size_t>) and tmp_badIPs (an lru11::Cache of 512 + 64
' entries) are arrays searched by their key.
Option Explicit

Private Const SOCKET_ERROR As Long = -1
Private Const MAX_TMP_BAD_IPS As Long = 512         ' lru11::Cache<std::string, TmpBadIPInfo> tmp_badIPs(512, 64)
Private Const TMP_BAD_IPS_ELASTICITY As Long = 64

Private netSendPendingDisconnectPlayerIndexes(0 To MAX_CONNECTED_PLAYERS - 1) As Boolean   ' static std::set<uint32_t>
Private netSendPendingDisconnectCount As Long

' std::unordered_map<std::string, size_t> tmp_pendingIPs
Private pendingIPKey() As String
Private pendingIPCount() As Long
Private pendingIPN As Long

' lru11::Cache<std::string, TmpBadIPInfo> tmp_badIPs: struct TmpBadIPInfo { expiry; attempts; } and the recency
Private badIPKey() As String
Private badIPExpiry() As Double                     ' std::chrono::steady_clock::time_point: steadyNowMs
Private badIPAttempts() As Double                   ' size_t
Private badIPUse() As Double                        ' (VB) the cache's recency order: a counter
Private badIPN As Long
Private badIPUseCounter As Double

Private recvBuffer() As Byte                        ' uint8_t buffer[NET_BUFFER_SIZE] (NETrecvNet's)

' (VB) a std::vector<uint8_t>'s size: the whole array, 0 when unallocated
Public Function byteVecSize(b() As Byte) As Long
    On Error GoTo none
    byteVecSize = UBound(b) - LBound(b) + 1
    Exit Function
none:
    byteVecSize = 0
End Function

' (VB) a big-endian uint32_t from b(ofs .. ofs + 3): memcpy, then wz_ntohl
Private Function be32(b() As Byte, ByVal ofs As Long) As Long
    be32 = SHL(b(ofs), 24) Or SHL(b(ofs + 1), 16) Or SHL(b(ofs + 2), 8) Or b(ofs + 3)
End Function

' (VB) wz_htonl(value) stored at b(ofs .. ofs + 3)
Private Sub putBe32(b() As Byte, ByVal ofs As Long, ByVal value As Long)
    b(ofs) = SHRU(value, 24) And &HFF&
    b(ofs + 1) = SHRU(value, 16) And &HFF&
    b(ofs + 2) = SHRU(value, 8) And &HFF&
    b(ofs + 3) = value And &HFF&
End Sub

' *********** Socket with buffer that read NETMSGs ******************

' pSocket: the variable holding the socket (connected_bsocket(i) or bsocket), set to Nothing when closed
Private Function NET_fillBuffer(pSocket As IClientConnection, ByVal pSocketSet As IConnectionPollGroup, bufstart() As Byte, ByVal bufsize As Long) As Long
    Dim socket As IClientConnection, rawBytes As Long, ec As Long, size As Long, readErrMsg As String
    Set socket = pSocket

    If Not socket.readReady() Then
        NET_fillBuffer = 0
        Exit Function
    End If

    size = socket.readNoInt(bufstart, 0, bufsize, rawBytes, ec)

    If ec = 0 Then
        nStats.rawBytes.received = nStats.rawBytes.received + U32(rawBytes)
        nStats.uncompressedBytes.received = nStats.uncompressedBytes.received + U32(size)
        nStats.packets.received = nStats.packets.received + 1

        NET_fillBuffer = size
        Exit Function
    Else
        If net_errorEqualsErrc(ec, errc_timed_out) Or net_errorEqualsErrc(ec, errc_connection_reset) Then
            debug_ 604, LOG_NET, "NET_fillBuffer", "Connection closed from the other side"
            NETlogEntry "Connection closed from the other side..", SYNC_FLAG, selectedPlayer
        Else
            readErrMsg = net_errorMessage(ec)
            debug_ 610, LOG_NET, "NET_fillBuffer", "%s socket %p is now invalid", readErrMsg, ObjPtr(socket)
        End If

        ' an error occurred, or the remote host has closed the connection.
        If Not pSocketSet Is Nothing Then
            pSocketSet.remove socket
        End If
        If bsocket Is socket Then
            debug_ 620, LOG_NET, "NET_fillBuffer", "Host connection was lost!"
            NETlogEntry "Host connection was lost!", SYNC_FLAG, selectedPlayer
            'Game is pretty much over --should just end everything when HOST dies.
            NETclientHandleHostDisconnected
            Set bsocket = Nothing
            netGame.setIngameLocalJoiningInProgress False   ' ingame.localJoiningInProgress = false
            NETclose
            NET_fillBuffer = 0
            Exit Function
        End If
        socket.close_
        Set pSocket = Nothing
    End If

    NET_fillBuffer = 0
End Function

' Connection dropped. Handle it gracefully. indexes: std::set<uint32_t> (indexes(i) for i in it)
Private Sub NETplayerClientsDisconnect(indexes() As Boolean)
    Dim index As Long, any_ As Boolean, w As MessageWriter
    Dim playersActuallyDisconnected(0 To MAX_CONNECTED_PLAYERS - 1) As Boolean
    If Not NetPlay.isHost Then
        ASSERT_FAILURE "false", "lib/netplay/netplay.cpp:1005", "NETplayerClientsDisconnect", "Host only routine detected for client!"
        Exit Sub
    End If

    For index = 0 To MAX_CONNECTED_PLAYERS - 1
        If indexes(index) Then any_ = True
    Next
    If Not any_ Then
        Exit Sub
    End If

    ' 1.) Do a pass just closing the sockets (in case multiple players involved)
    For index = 0 To MAX_CONNECTED_PLAYERS - 1
        If indexes(index) Then
            If Not connected_bsocket(index) Is Nothing Then
                debug_ 1026, LOG_NET, "NETplayerClientsDisconnect", "Player (%u) has left unexpectedly, closing socket %p", index, ObjPtr(connected_bsocket(index))

                NETplayerCloseSocket index, False
                playersActuallyDisconnected(index) = True
            Else
                debug_ 1033, LOG_ERROR, "NETplayerClientsDisconnect", "Player (%u) has left unexpectedly - but socket already closed?", index
            End If
        End If
    Next

    ' 2.) Notify other players about which players were disconnected
    For index = 0 To MAX_CONNECTED_PLAYERS - 1
        If playersActuallyDisconnected(index) Then
            NETplayerLeaving index, True

            NETlogEntry "Player has left unexpectedly.", SYNC_FLAG, index
            ' Announce to the world. This was really icky, because we may have been calling the send
            ' function recursively. We really ought to have had a send queue, and now we finally do...
            If netGame.ingameLocalJoiningInProgress() Then  ' Only if game hasn't actually started yet.
                Set w = NETbeginEncode(NETbroadcastQueue(), NET_PLAYER_DROPPED)
                NETuint32_t_w w, index
                NETend_w w
            End If
        End If
    Next
End Sub

Private Sub NETplayerCloseSocket(ByVal index As Long, ByVal quietSocketClose As Boolean)
    If Not ULT(index, MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "index < MAX_CONNECTED_PLAYERS", "lib/netplay/netplay.cpp:1061", "NETplayerCloseSocket", "Invalid index: %u", index: Exit Sub
    If Not connected_bsocket(index) Is Nothing Then
        debug_ 1064, LOG_NET, "NETplayerCloseSocket", "Player (%u) has left, closing socket %p", index, ObjPtr(connected_bsocket(index))
        NETlogEntry "Player has left nicely.", SYNC_FLAG, index

        ' Although we can get a error result from DelSocket, it don't really matter here.
        server_socket_set.remove connected_bsocket(index)
        connected_bsocket(index).close_
        Set connected_bsocket(index) = Nothing
    Else
        debug_ 1074, LOG_NET, "NETplayerCloseSocket", "Player (%u) has left nicely, socket already closed?", index
    End If

    NetPlay.players(index).heartbeat = False ' mark client as dead
End Sub

' When a player leaves nicely (ie, we got a NET_PLAYER_LEAVING message), we clean up the socket that we used.
Public Sub NETplayerLeaving(ByVal index As Long, Optional ByVal quietSocketClose As Boolean = False)
    Dim previousPlayersShouldCheckReadyValue As Boolean, wasSpectator As Boolean
    If Not ULT(index, MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "index < MAX_CONNECTED_PLAYERS", "lib/netplay/netplay.cpp:1087", "NETplayerLeaving", "Invalid index: %u", index: Exit Sub
    previousPlayersShouldCheckReadyValue = netGame.multiplayPlayersShouldCheckReady()
    NETplayerCloseSocket index, quietSocketClose
    sync_counter.left_ = (sync_counter.left_ + 1) And &HFFFF&
    wasSpectator = NetPlay.players(index).isSpectator
    netGame.MultiPlayerLeave index      ' more cleanup
    If netGame.ingameLocalJoiningInProgress() Then  ' Only if game hasn't actually started yet.
        NET_DestroyPlayer index         ' sets index player's array to false
        If Not wasSpectator Then
            netGame.resetReadyStatus False, netGame.shouldSkipReadyResetOnPlayerJoinLeaveEvent()     ' reset ready status for all players
        End If
        netGame.handlePossiblePlayersShouldCheckReadyChange previousPlayersShouldCheckReadyValue

        netGame.wzCommandInterfaceOutputRoomStatusJson True
    End If
End Sub

' When a player's connection is broken we broadcast the NET_PLAYER_DROPPED message.
Private Sub NETplayerDropped(ByVal index As Long)
    Dim id As Long, previousPlayersShouldCheckReadyValue As Boolean, wasSpectator As Boolean, w As MessageWriter
    If Not ULT(index, NetPlay.playersSize) Then ASSERT_FAILURE "index < NetPlay.players.size()", "lib/netplay/netplay.cpp:1113", "NETplayerDropped", "Invalid index: %u", index: Exit Sub
    id = index

    If Not NetPlay.isHost Then
        ASSERT_FAILURE "false", "lib/netplay/netplay.cpp:1118", "NETplayerDropped", "Host only routine detected for client!"
        Exit Sub
    End If

    previousPlayersShouldCheckReadyValue = netGame.multiplayPlayersShouldCheckReady()

    sync_counter.drops = (sync_counter.drops + 1) And &HFFFF&
    wasSpectator = NetPlay.players(index).isSpectator
    netGame.MultiPlayerLeave id         ' more cleanup
    If netGame.ingameLocalJoiningInProgress() Then  ' Only if game hasn't actually started yet.
        ' Send message type specifically for dropped / disconnects
        Set w = NETbeginEncode(NETbroadcastQueue(), NET_PLAYER_DROPPED)
        NETuint32_t_w w, id
        NETend_w w
        debug_ 1133, LOG_INFO, "NETplayerDropped", "sending NET_PLAYER_DROPPED for player %d", id
        NET_DestroyPlayer id            ' just clears array
        If Not wasSpectator Then
            netGame.resetReadyStatus False, netGame.shouldSkipReadyResetOnPlayerJoinLeaveEvent()     ' reset ready status for all players
        End If
        netGame.handlePossiblePlayersShouldCheckReadyChange previousPlayersShouldCheckReadyValue

        netGame.wzCommandInterfaceOutputRoomStatusJson True
    End If

    NETsetPlayerConnectionStatus CONNECTIONSTATUS_PLAYER_DROPPED, id
End Sub

' Cleanup for when a player is kicked.
Public Sub NETplayerKicked(ByVal index As Long, Optional ByVal quiet As Boolean = False)
    Dim lvl As Long
    If Not ULT(index, MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "index < MAX_CONNECTED_PLAYERS", "lib/netplay/netplay.cpp:1153", "NETplayerKicked", "NETplayerKicked invalid player_id: (%u)", index: Exit Sub

    ' kicking a player counts as "leaving nicely", since "nicely" in this case
    ' simply means "there wasn't a connection error."
    If Not quiet Then lvl = LOG_INFO Else lvl = LOG_NET
    debug_ 1157, lvl, "NETplayerKicked", "Player %u was kicked.", index
    sync_counter.kicks = (sync_counter.kicks + 1) And &HFFFF&
    NETlogEntry "Player was kicked.", SYNC_FLAG, index
    NETplayerLeaving index      ' need to close socket for the player that left.
    NETsetPlayerConnectionStatus CONNECTIONSTATUS_PLAYER_LEAVING, index
End Sub

' ////////////////////////////////////////////////////////////////////////
'close the open game..
Public Function NETclose() As Long
    Dim i As Long
    NEThaltJoining False

    debug_ 1449, LOG_NET, "NETclose", "Terminating sockets."

    NetPlay.isHost = False
    server_not_there = False
    allow_joining = False

    knownExternalIPv4Info.host = ""     ' knownExternalIPv4Info = KnownExternalConnInfo()
    knownExternalIPv4Info.port = 0

    For i = 0 To MAX_CONNECTED_PLAYERS - 1
        If Not connected_bsocket(i) Is Nothing Then
            debug_ 1461, LOG_NET, "NETclose", "Closing connected_bsocket[%u], %p", i, ObjPtr(connected_bsocket(i))
            connected_bsocket(i).close_
            Set connected_bsocket(i) = Nothing
        End If
        NET_DestroyPlayer i, True
    Next

    If Not tmp_socket_set Is Nothing Then
        debug_ 1470, LOG_NET, "NETclose", "Freeing tmp_socket_set %p", ObjPtr(tmp_socket_set)
        Set tmp_socket_set = Nothing    ' delete tmp_socket_set
    End If

    For i = 0 To MAX_TMP_SOCKETS - 1
        If Not tmp_socket(i) Is Nothing Then
            ' FIXME: need SocketSet_DelSocket() as well, socket_set or tmp_socket_set?
            debug_ 1480, LOG_NET, "NETclose", "Closing tmp_socket[%d] %p", i, ObjPtr(tmp_socket(i))
            tmp_socket(i).close_
            Set tmp_socket(i) = Nothing
        End If
    Next

    If Not client_socket_set Is Nothing Then
        If Not bsocket Is Nothing Then
            client_socket_set.remove bsocket
        End If
        debug_ 1492, LOG_NET, "NETclose", "Freeing socket_set %p", ObjPtr(client_socket_set)
        Set client_socket_set = Nothing
    ElseIf Not server_socket_set Is Nothing Then
        debug_ 1498, LOG_NET, "NETclose", "Freeing socket_set %p", ObjPtr(server_socket_set)
        Set server_socket_set = Nothing
    End If
    If Not server_listen_socket Is Nothing Then
        debug_ 1504, LOG_NET, "NETclose", "Closing server_listen_socket %p", ObjPtr(server_listen_socket)
        server_listen_socket.close_     ' delete server_listen_socket
        Set server_listen_socket = Nothing
    End If
    If Not bsocket Is Nothing Then
        debug_ 1510, LOG_NET, "NETclose", "Closing bsocket %p", ObjPtr(bsocket)
        bsocket.close_
        Set bsocket = Nothing
    End If

    playerManagementRecord_clear

    NETclose = 0
End Function

Public Sub NETsendProcessDelayedActions()
    Dim pendingDisconnectPlayers(0 To MAX_CONNECTED_PLAYERS - 1) As Boolean, player As Long
    If netSendPendingDisconnectCount = 0 Then
        Exit Sub
    End If

    ' "consume" the netSendPendingDisconnectPlayerIndexes up-front, and do *not* reference it again (as NETplayerClientsDisconnect may lead to calls that mutate it)
    For player = 0 To MAX_CONNECTED_PLAYERS - 1
        pendingDisconnectPlayers(player) = netSendPendingDisconnectPlayerIndexes(player)
        netSendPendingDisconnectPlayerIndexes(player) = False
    Next
    netSendPendingDisconnectCount = 0

    For player = 0 To MAX_CONNECTED_PLAYERS - 1
        If pendingDisconnectPlayers(player) Then
            NETlogEntry "client disconnect?", SYNC_FLAG, player
        End If
    Next

    NETplayerClientsDisconnect pendingDisconnectPlayers
End Sub

' ////////////////////////////////////////////////////////////////////////
' Send a message to a player, option to guarantee message
Public Function NETsend(ByVal queue As Long, ByVal message As NetMessage) As Boolean
    Dim player As Long, isTmpQueue As Boolean, firstPlayer As Long, lastPlayer As Long
    Dim rawData() As Byte, rawLen As Long, compressedRawLen As Long, ec As Long, res As Long, n As Long
    Dim msgType As Long, sender As Long, w As MessageWriter, writeErrMsg As String, sock As IClientConnection
    player = NETQUEUE_index(queue)

    If Not NetPlay.bComms Then
        NETsend = True
        Exit Function
    End If

    ' IClientConnection** sockets = connected_bsocket; (or tmp_socket for QUEUE_TMP)
    isTmpQueue = False
    Select Case NETQUEUE_queueType(queue)
    Case QUEUE_BROADCAST
        If Not (player = NET_ALL_PLAYERS) Then ASSERT_FAILURE "player == NET_ALL_PLAYERS", "lib/netplay/netplay.cpp:1596", "NETsend", "Wrong queue index.": NETsend = False: Exit Function
    Case QUEUE_NET
        If Not (player < MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "player < MAX_CONNECTED_PLAYERS", "lib/netplay/netplay.cpp:1599", "NETsend", "Wrong queue index.": NETsend = False: Exit Function
    Case QUEUE_TMP
        isTmpQueue = True

        If Not (player < MAX_TMP_SOCKETS And NetPlay.isHost) Then ASSERT_FAILURE "player < MAX_TMP_SOCKETS && NetPlay.isHost", "lib/netplay/netplay.cpp:1605", "NETsend", "Wrong queue index.": NETsend = False: Exit Function
    Case Else
        ASSERT_FAILURE "false", "lib/netplay/netplay.cpp:1608", "NETsend", "Wrong queue type.": NETsend = False: Exit Function
    End Select

    If NetPlay.isHost Then
        If player = NET_ALL_PLAYERS Then firstPlayer = 0 Else firstPlayer = player
        If player = NET_ALL_PLAYERS Then lastPlayer = MAX_CONNECTED_PLAYERS - 1 Else lastPlayer = player
        For player = firstPlayer To lastPlayer
            ' We are the host, send directly to player.
            If isTmpQueue Then Set sock = tmp_socket(player) Else Set sock = connected_bsocket(player)
            If Not sock Is Nothing And player <> NETQUEUE_exclude(queue) Then
                rawData = message.rawData()
                rawLen = message.rawDataSize()
                If rawLen = 0 Then
                    debug_ 1623, LOG_FATAL, "NETsend", "Failed to allocate raw data (message type: %u, player: %d)", message.type_(), player
                    exit_ 1                     ' abort()
                End If
                msgType = message.type_()
                compressedRawLen = 0
                n = sock.writeAll(rawData, 0, rawLen, compressedRawLen, ec)
                If ec = 0 Then res = n Else res = SOCKET_ERROR

                If res = rawLen Then
                    nStats.rawBytes.sent = nStats.rawBytes.sent + U32(compressedRawLen)
                    nStats.uncompressedBytes.sent = nStats.uncompressedBytes.sent + rawLen
                    nStats.packets.sent = nStats.packets.sent + 1
                ElseIf res = SOCKET_ERROR Then
                    writeErrMsg = net_errorMessage(ec)
                    ' Write error, most likely client disconnect.
                    debug_ 1642, LOG_ERROR, "NETsend", "Failed to send message (type: %u, rawLen: %zu, compressedRawLen: %zu) to %u: %s", _
                        msgType, rawLen, compressedRawLen, player, writeErrMsg
                    If Not isTmpQueue Then
                        If Not netSendPendingDisconnectPlayerIndexes(player) Then
                            netSendPendingDisconnectPlayerIndexes(player) = True
                            netSendPendingDisconnectCount = netSendPendingDisconnectCount + 1
                        End If
                    End If
                End If
            End If
        Next
        NETsend = True
        Exit Function
    ElseIf player = NetPlay.hostPlayer Then
        ' We are a client, send directly to player, who happens to be the host.
        If Not bsocket Is Nothing And NetPlay.isHostAlive Then
            rawData = message.rawData()
            rawLen = message.rawDataSize()
            compressedRawLen = 0
            n = bsocket.writeAll(rawData, 0, rawLen, compressedRawLen, ec)
            If ec = 0 Then res = n Else res = SOCKET_ERROR

            If res = rawLen Then
                nStats.rawBytes.sent = nStats.rawBytes.sent + U32(compressedRawLen)
                nStats.uncompressedBytes.sent = nStats.uncompressedBytes.sent + rawLen
                nStats.packets.sent = nStats.packets.sent + 1
            ElseIf res = SOCKET_ERROR Then
                writeErrMsg = net_errorMessage(ec)
                ' Write error, most likely host disconnect.
                debug_ 1674, LOG_ERROR, "NETsend", "Failed to send message: %s", writeErrMsg
                debug_ 1675, LOG_ERROR, "NETsend", "Host connection was broken, socket %p.", ObjPtr(bsocket)
                NETlogEntry "write error--client disconnect.", SYNC_FLAG, NetPlay.hostPlayer
                NETclientHandleHostDisconnected
            End If

            NETsend = (res = rawLen)
            Exit Function
        End If
    Else
        ' We are a client and can't send the data directly, ask the host to send the data to the player.
        sender = selectedPlayer
        Set w = NETbeginEncode(NETnetQueue(NetPlay.hostPlayer), NET_SEND_TO_PLAYER)
        NETuint8_t_w w, sender
        NETuint8_t_w w, player
        NETnetMessage_w w, message
        NETend_w w
        NETsend = True
        Exit Function
    End If

    NETsend = False
End Function

' (VB) TmpSocketInfo::reset()
Private Sub TmpSocketInfo_reset(ByVal i As Long)
    Dim n As Long
    tmp_connectState(i).ip = ""
    tmp_connectState(i).connectTime = 0
    For n = 0 To 9
        tmp_connectState(i).buffer(n) = 0
    Next
    tmp_connectState(i).usedBuffer = 0
    Erase tmp_connectState(i).connectChallenge
    tmp_connectState(i).connectChallengeSize = 0
    tmp_connectState(i).connectState = TmpConnectState_None
    ReceivedJoinInfo_reset tmp_connectState(i).receivedJoinInfo
    tmp_connectState(i).uniqueJoinID = ""
    tmp_connectState(i).has_asyncJoinApprovalResult = False
    tmp_connectState(i).has_asyncJoinApprovalExplicitPlayerIdx = False
    tmp_connectState(i).asyncJoinRejectCode = ERROR_NOERROR
    tmp_connectState(i).asyncJoinRejectCustomMessage = ""
End Sub

' (VB) TmpSocketInfo::ReceivedJoinInfo::reset()
Private Sub ReceivedJoinInfo_reset(info As ReceivedJoinInfo)
    Dim noKeys As SessionKeys
    info.name = ""
    info.playerType = 0
    EcKey_clear info.identity
    info.has_connectionAuthSessionKeys = False  ' connectionAuthSessionKeys.reset()
    info.connectionAuthSessionKeys = noKeys
    Erase info.challengeForHost
    info.challengeForHostSize = 0
End Sub

Private Sub NETcloseTempSocket(ByVal i As Long)
    Dim joinerPublicKeyB64 As String, joinerName As String, joinerNameB64 As String, rIP As String
    Dim key() As Byte, n As Long, s As String

    If tmp_connectState(i).connectState = TmpConnectState_PendingAsyncApproval Or tmp_connectState(i).connectState = TmpConnectState_ProcessJoin Then
        ' If there's any chance we may have issued an async join approval request, ensure we issue a joinFailed event
        ' WZEVENT: joinFailed: <b64pubkey> <b64name> [spec|play] <reason>
        n = EcKey_toBytes(tmp_connectState(i).receivedJoinInfo.identity, EcKey_Public, key)
        joinerPublicKeyB64 = base64Encode(key, n)
        joinerName = tmp_connectState(i).receivedJoinInfo.name
        joinerNameB64 = base64EncodeString(joinerName)
        If tmp_connectState(i).receivedJoinInfo.playerType = NET_JOIN_SPECTATOR Then s = "spec" Else s = "play"
        netGame.wzCommandInterfaceOutput va("WZEVENT: joinFailed: %s %s %s %s" & vbLf, joinerPublicKeyB64, joinerNameB64, s, "full")
    End If

    rIP = tmp_socket(i).textAddress()
    tmp_socket_set.remove tmp_socket(i)
    tmp_socket(i).close_
    Set tmp_socket(i) = Nothing
    TmpSocketInfo_reset i
    n = pendingIPs_find(rIP)
    If n >= 0 Then
        If pendingIPCount(n) > 1 Then
            pendingIPCount(n) = pendingIPCount(n) - 1
        Else
            pendingIPs_erase n
        End If
    Else
        debug_ 1734, LOG_INFO, "NETcloseTempSocket", "Did not find in pending IPs list: (IP: %s, Identity: %s, Name: %s)", rIP, joinerPublicKeyB64, joinerName
    End If
End Sub

' (VB) base64Encode(std::vector<unsigned char>(s.begin(), s.end())): the string's UTF-8 bytes
Public Function base64EncodeString(ByVal s As String) As String
    Dim b() As Byte, n As Long
    n = utf8_encode(s, b)
    base64EncodeString = base64Encode(b, n)
End Function

Public Sub NETclientHandleHostDisconnected()
    If Not (Not NetPlay.isHost) Then ASSERT_FAILURE "!NetPlay.isHost", "lib/netplay/netplay.cpp:1740", "NETclientHandleHostDisconnected", "Should only be called by clients!": Exit Sub
    ' NOTE: Do *not* close the socket or connection group here, as this is called from both send and receiving functions
    ' And if this was detected on a send, we still might have buffered data to read!
    If Not ULT(NetPlay.hostPlayer, NetPlay.playersSize) Then ASSERT_FAILURE "NetPlay.hostPlayer < NetPlay.players.size()", "lib/netplay/netplay.cpp:1743", "NETclientHandleHostDisconnected", "Invalid NetPlay.hostPlayer: %u", NetPlay.hostPlayer
    NetPlay.players(NetPlay.hostPlayer).heartbeat = False   ' mark host as dead
    'Game is pretty much over --should just end everything when HOST dies.
    NetPlay.isHostAlive = False
End Sub

Public Sub NETflush()
    Dim compressedRawLen As Long, player As Long, ec As Long
    Dim invalidPlayerIndices(0 To MAX_CONNECTED_PLAYERS - 1) As Boolean
    If Not NetPlay.bComms Then
        Exit Sub
    End If

    NETflushGameQueues

    compressedRawLen = 0
    If NetPlay.isHost Then
        ' Preliminary check to see if any player sockets are still valid.
        For player = 0 To MAX_CONNECTED_PLAYERS - 1
            If Not connected_bsocket(player) Is Nothing Then
                If Not connected_bsocket(player).isValid() Then
                    invalidPlayerIndices(player) = True
                End If
            End If
        Next
        ' Gracefully handle disconnected players.
        NETplayerClientsDisconnect invalidPlayerIndices

        For player = 0 To MAX_CONNECTED_PLAYERS - 1
            ' We are the host, send directly to player.
            If Not invalidPlayerIndices(player) And Not connected_bsocket(player) Is Nothing Then
                ec = connected_bsocket(player).flush(compressedRawLen)
                If ec <> 0 Then
                    invalidPlayerIndices(player) = True
                    GoTo continue1
                End If
                nStats.rawBytes.sent = nStats.rawBytes.sent + U32(compressedRawLen)
            End If
continue1:
        Next

        ' Once again, if any connected client sockets had problems during flush(), consider them disconnected and handle appropriately.
        NETplayerClientsDisconnect invalidPlayerIndices

        For player = 0 To MAX_TMP_SOCKETS - 1
            ' We are the host, send directly to player.
            If Not tmp_socket(player) Is Nothing Then
                If Not tmp_socket(player).isValid() Then
                    ec = 1
                Else
                    ec = tmp_socket(player).flush(compressedRawLen)
                End If
                If ec <> 0 Then
                    debug_ 1797, LOG_NET, "NETflush", "Failed to flush temporary socket for player slot %d", player
                    NETcloseTempSocket player
                    GoTo continue2
                End If
                nStats.rawBytes.sent = nStats.rawBytes.sent + U32(compressedRawLen)
            End If
continue2:
        Next
    Else
        If Not bsocket Is Nothing And NetPlay.isHostAlive Then
            If Not bsocket.isValid() Then
                ec = 1
            Else
                ec = bsocket.flush(compressedRawLen)
            End If
            If ec <> 0 Then
                debug_ 1811, LOG_NET, "NETflush", "Failed to flush socket for host. Host probably disconnected."
                NETlogEntry "flush error--client disconnect.", SYNC_FLAG, NetPlay.hostPlayer
                NETclientHandleHostDisconnected
                Exit Sub
            End If
            nStats.rawBytes.sent = nStats.rawBytes.sent + U32(compressedRawLen)
        End If
    End If
    PendingWritesManager_pollAll                ' (VB) the writer thread's pass
End Sub

' Checks to see if a human player is still with us.
' @note: resuscitation isn't possible with current code, so once we lose
' the socket, then we have no way to connect with them again. Future
' item to enhance.
Private Sub NETcheckPlayers()
    Dim i As Long
    If Not NetPlay.isHost Then
        ASSERT_FAILURE "false", "lib/netplay/netplay.cpp:2824", "NETcheckPlayers", "Host only routine detected for client or not hosting yet!"
        Exit Sub
    End If

    For i = 0 To MAX_CONNECTED_PLAYERS - 1
        If NetPlay.players(i).allocated = False Then
            GoTo continue_    ' not allocated means that it most likely it is a AI player
        End If
        If Not NetPlay.players(i).heartbeat And NetPlay.players(i).heartattacktime = 0 Then    ' looks like they are dead
            NetPlay.players(i).heartattacktime = netGame.realTime()     ' mark when this occurred
        End If
continue_:
    Next
End Sub

' ////////////////////////////////////////////////////////////////////////
' Receive a message over the current connection. We return true if there
' is a message for the higher level code to process, and false otherwise.
' We should not block here.
Public Function NETrecvNet(queue As Long, typ As Long) As Boolean
    Dim pollGroup As IConnectionPollGroup, current As Long, dataLen As Long, ec As Long, readable As Long
    If Not NetPlay.bComms Then
        NETrecvNet = False
        Exit Function
    End If

    PendingWritesManager_pollAll                ' (VB) the writer thread's pass

    If Not activeConnProvider Is Nothing Then
        activeConnProvider.processConnectionStateChanges
    End If

    If NetPlay.isHost Then
        NETfixPlayerCount
        NETacceptIncomingConnections
        NETallowJoining
        If lobbyHostingHandler_exists() Then
            lobbyHostingHandler_run
        End If
        NETcheckPlayers         ' make sure players are still alive & well
    End If

    If NetPlay.isHost Then Set pollGroup = server_socket_set Else Set pollGroup = client_socket_set
    If pollGroup Is Nothing Then GoTo checkMessages
    readable = pollGroup.checkConnectionsReadable(NET_READ_TIMEOUT, ec)
    If ec <> 0 Then readable = 0                ' .value_or(0)
    If readable <= 0 Then GoTo checkMessages

    If byteVecSize(recvBuffer) = 0 Then ReDim recvBuffer(0 To NET_BUFFER_SIZE - 1)
    For current = 0 To MAX_CONNECTED_PLAYERS - 1
        If Not NetPlay.isHost And current <> NetPlay.hostPlayer Then
            GoTo continue_  ' Don't have a socket open to this player.
        End If

        If NetPlay.isHost Then
            If connected_bsocket(current) Is Nothing Then GoTo continue_
            dataLen = NET_fillBuffer(connected_bsocket(current), pollGroup, recvBuffer, NET_BUFFER_SIZE)
        Else
            If bsocket Is Nothing Then GoTo continue_
            dataLen = NET_fillBuffer(bsocket, pollGroup, recvBuffer, NET_BUFFER_SIZE)
        End If
        If dataLen > 0 Then
            ' we received some data, add to buffer
            NETinsertRawData NETnetQueue(current), recvBuffer, dataLen
        ElseIf socketIsGone(current) Then
            ' If there is a error in NET_fillBuffer() then socket is already invalid.
            ' This means that the player dropped / disconnected for whatever reason.
            debug_ 2901, LOG_INFO, "NETrecvNet", "Player, (player %u) seems to have dropped/disconnected.", current

            If NetPlay.isHost Then
                ' Send message type specifically for dropped / disconnects
                NETplayerDropped current
                Set connected_bsocket(current) = Nothing    ' clear their socket
            End If
        End If
continue_:
    Next

checkMessages:
    For current = 0 To MAX_CONNECTED_PLAYERS - 1
        queue = NETnetQueue(current)
        Do While NETisMessageReady(queue)
            typ = NETgetMessage(queue).type_()
            If Not NETprocessSystemMessage(queue, typ) Then
                NETrecvNet = True  ' We couldn't process the message, let the caller deal with it..
                Exit Function
            End If
        Loop
    Next

    NETrecvNet = False
End Function

' (VB) *pSocket == nullptr after NET_fillBuffer
Private Function socketIsGone(ByVal current As Long) As Boolean
    If NetPlay.isHost Then socketIsGone = connected_bsocket(current) Is Nothing Else socketIsGone = bsocket Is Nothing
End Function

Public Function NETrecvGame(queue As Long, typ As Long) As Boolean
    Dim current As Long
    PendingWritesManager_pollAll                ' (VB) the writer thread's pass
    For current = 0 To MAX_GAMEQUEUE_SLOTS - 1
        queue = NETgameQueue(current)
        If NETQUEUE_queueIsNull(queue) Then
            GoTo continue_
        End If
        Do While Not netGame.checkPlayerGameTime(current)  ' Check for any messages that are scheduled to be read now.
            If Not NETisMessageReady(queue) Then
                NETrecvGame = False             ' Still waiting for messages from this player, and all players should process messages in the same order. Will have to freeze the game while waiting.
                Exit Function
            End If

            NETreplaySaveNetMessage NETgetMessage(queue), NETQUEUE_index(queue)
            typ = NETgetMessage(queue).type_()

            If typ = GAME_GAME_TIME Then
                netGame.recvPlayerGameTime queue
                NETpop queue
                GoTo nextMessage
            End If

            NETrecvGame = True                  ' Have a message ready to read now.
            Exit Function
nextMessage:
        Loop
continue_:
    Next

    NETrecvGame = False                         ' No messages sceduled to be read yet. Game can continue.
End Function

' Now that we know whether a game is listed in the lobby, send a signal to ActivityManager
' (ActivityManager::instance().hostGame(...) with the listening interfaces: the activity sinks, not ported)
Public Sub handleLobbyHostingStatusIsKnown(ByVal hasLobbyGameId As Boolean, ByVal lobbyGameId As String)
End Sub

Public Sub NETupdateLobbyGameListing()
    If Not lobbyHostingHandler_exists() Then
        Exit Sub
    End If

    ' updateGameListing(getMultiStats(realSelectedPlayer).identity, gameDetails, hostJoinOptions)
    lobbyHostingHandler_updateGameListing realSelectedPlayer
End Sub

' ////////////////////////////////////////////////////////////////////////
'  Check player "slots" & update player count if needed.
Public Sub NETfixPlayerCount()
    Dim maxPlayers As Long, playercount As Long, index As Long, latestSpecInfo As Long
    If Not NetPlay.isHost Then ASSERT_FAILURE "false", "lib/netplay/netplay.cpp:3369", "NETfixPlayerCount", "Host only routine detected for client!": Exit Sub
    If Not allow_joining Then
        Exit Sub
    End If

    maxPlayers = netGame.gameMaxPlayers()
    playercount = 0
    For index = 0 To netGame.gameMaxPlayers() - 1
        If NetPlay.players(index).ai = AI_CLOSED Then
            maxPlayers = maxPlayers - 1
        ElseIf NetPlay.players(index).isSpectator Then
            maxPlayers = maxPlayers - 1
        ElseIf NetPlay.players(index).ai <> AI_OPEN Or NetPlay.players(index).allocated Then
            playercount = playercount + 1
        End If
    Next

    latestSpecInfo = SpectatorInfo_currentNetPlayState()

    If NetPlay.playercount <> playercount Or gameDetails.players_max <> maxPlayers Or gameDetails.spectators_current <> SpectatorInfo_spectatorsJoined(latestSpecInfo) _
       Or gameDetails.spectators_max <> SpectatorInfo_totalSpectatorSlots(latestSpecInfo) Then
        debug_ 3397, LOG_NET, "NETfixPlayerCount", "Updating player count from %u/%u to %d/%d", NetPlay.playercount, gameDetails.players_max, playercount, maxPlayers

        NetPlay.playercount = playercount
        gameDetails.players_current = playercount
        gameDetails.players_max = maxPlayers
        gameDetails.spectators_current = SpectatorInfo_spectatorsJoined(latestSpecInfo)
        gameDetails.spectators_max = SpectatorInfo_totalSpectatorSlots(latestSpecInfo)

        NETupdateLobbyGameListing
    End If
End Sub

Public Sub NETsetLobbyConfigFlagsFields(ByVal alliancesType As Long, ByVal techLevel As Long, ByVal powerLevel As Long, ByVal basesLevel As Long)
    gameDetails.alliancesMode = alliancesType And &HFF&

    NETupdateLobbyGameListing
End Sub

' ingameLimitFlags: ingame.flags (uint8_t)
Public Sub NETsetLobbyLimitFlags(ByVal ingameLimitFlags As Long)
    Dim priorLimits As Long
    priorLimits = gameDetails.limits

    gameDetails.limits = 0

    If (ingameLimitFlags And &H1&) <> 0 Then    ' MPFLAGS_NO_TANKS
        gameDetails.limits = gameDetails.limits Or &H1&     ' StructureLimits::NO_TANKS
    End If
    If (ingameLimitFlags And &H2&) <> 0 Then    ' MPFLAGS_NO_CYBORGS
        gameDetails.limits = gameDetails.limits Or &H2&     ' StructureLimits::NO_BORGS
    End If
    If (ingameLimitFlags And &H4&) <> 0 Then    ' MPFLAGS_NO_VTOLS
        gameDetails.limits = gameDetails.limits Or &H4&     ' StructureLimits::NO_VTOLS
    End If

    If priorLimits <> gameDetails.limits Then
        NETupdateLobbyGameListing
    End If
End Sub

' expirySeconds: std::chrono::seconds expiry (1800 by default)
Private Sub NETaddSessionBanBadIP(ByVal badIP As String, Optional ByVal expirySeconds As Long = 1800)
    Dim now_ As Double, idx As Long
    If isLoopbackIP(badIP) Then
        Exit Sub
    End If
    now_ = steadyNowMs()
    idx = badIPs_tryGet(badIP)
    If idx >= 0 Then
        If now_ < badIPExpiry(idx) Then
            If badIPAttempts(idx) < 1.84467440737096E+19 Then
                badIPAttempts(idx) = badIPAttempts(idx) + 1
            End If
        Else
            badIPAttempts(idx) = 1
        End If

        If now_ + expirySeconds * 1000# > badIPExpiry(idx) Then badIPExpiry(idx) = now_ + expirySeconds * 1000#

        If badIPAttempts(idx) = 100 Then
            ' add to permanent ban list
            addIPToBanList badIP, "BAD_USER"
        End If
    Else
        badIPs_insert badIP, now_ + expirySeconds * 1000#, 1
    End If
End Sub

Private Function quickRejectConnection(ByVal ip As String) As Boolean
    Dim n As Long, idx As Long
    If isLoopbackIP(ip) Then
        quickRejectConnection = False
        Exit Function
    End If
    n = pendingIPs_find(ip)
    If n >= 0 Then
        If pendingIPCount(n) > 1 Then
            quickRejectConnection = True
            Exit Function
        End If
    End If
    idx = badIPs_tryGet(ip)
    If idx >= 0 Then
        If steadyNowMs() < badIPExpiry(idx) Then
            NETaddSessionBanBadIP ip
            quickRejectConnection = True
            Exit Function
        Else
            ' clean up expired bad entry
            badIPs_remove idx
        End If
    End If

    quickRejectConnection = False
End Function

Private Sub NEThostPromoteTempSocketToPermanentPlayerConnection(ByVal tempSocketIdx As Long, ByVal index As Long)
    Dim rIP As String, n As Long
    rIP = tmp_socket(tempSocketIdx).textAddress()

    debug_ 3514, LOG_NET, "NEThostPromoteTempSocketToPermanentPlayerConnection", "freeing temp socket %p (%d), creating permanent socket.", ObjPtr(tmp_socket(tempSocketIdx)), 3514
    tmp_socket_set.remove tmp_socket(tempSocketIdx)
    Set connected_bsocket(index) = tmp_socket(tempSocketIdx)
    Set tmp_socket(tempSocketIdx) = Nothing
    has_NET_waitingForIndexChangeAckSince(index) = False
    server_socket_set.add connected_bsocket(index)
    NETmoveQueue NETnetTmpQueue(tempSocketIdx), NETnetQueue(index)

    ' Copy player's IP address
    NetPlay.players(index).IPtextAddress = sstrcpyN(rIP, 40)

    ' Decrement pending IP counter
    n = pendingIPs_find(rIP)
    If n >= 0 Then
        If pendingIPCount(n) > 1 Then
            pendingIPCount(n) = pendingIPCount(n) - 1
        Else
            pendingIPs_erase n
        End If
    End If
End Sub

Private Sub NETrejectTempSocketClient(ByVal i As Long, ByVal rejectedReason As Long, Optional ByVal sessionBan As Boolean = False)
    Dim w As MessageWriter
    Set w = NETbeginEncode(NETnetTmpQueue(i), NET_REJECTED)
    NETuint8_t_w w, rejectedReason
    NETstring_w w, "", 0
    NETend_w w
    NETflush

    If sessionBan Then
        NETaddSessionBanBadIP tmp_connectState(i).ip
    End If
    NETcloseTempSocket i
    sync_counter.cantjoin = (sync_counter.cantjoin + 1) And &HFFFF&
End Sub

Private Sub NETtempSocketBeginHandleJoinApproval(ByVal i As Long)
    Dim currentConnectState As Long, joinerPublicKeyB64 As String, joinerIdentityHash As String, joinerName As String
    Dim joinerNameB64 As String, key() As Byte, n As Long, s As String
    currentConnectState = tmp_connectState(i).connectState
    If Not ((currentConnectState = TmpConnectState_PendingJoinRequest) Or (currentConnectState = TmpConnectState_PendingLobbyJoinCheckResult)) Then ASSERT_FAILURE _
        "(currentConnectState == TmpSocketInfo::TmpConnectState::PendingJoinRequest) || (currentConnectState == TmpSocketInfo::TmpConnectState::PendingLobbyJoinCheckResult)", _
        "lib/netplay/netplay.cpp:3559", "NETtempSocketBeginHandleJoinApproval", "Unexpected connectState: %d", currentConnectState: Exit Sub

    If bAsyncJoinApprovalEnabled Then
        tmp_connectState(i).connectState = TmpConnectState_PendingAsyncApproval

        ' use the connectChallenge as a unique, random join ID
        tmp_connectState(i).uniqueJoinID = base64Encode(tmp_connectState(i).connectChallenge, tmp_connectState(i).connectChallengeSize)

        n = EcKey_toBytes(tmp_connectState(i).receivedJoinInfo.identity, EcKey_Public, key)
        joinerPublicKeyB64 = base64Encode(key, n)
        joinerIdentityHash = EcKey_publicHashString(tmp_connectState(i).receivedJoinInfo.identity)
        joinerName = tmp_connectState(i).receivedJoinInfo.name
        joinerNameB64 = base64EncodeString(joinerName)
        If tmp_connectState(i).receivedJoinInfo.playerType = NET_JOIN_SPECTATOR Then s = "spec" Else s = "play"
        netGame.wzCommandInterfaceOutput va("WZEVENT: join approval needed: %s %s %s %s %s %s" & vbLf, tmp_connectState(i).uniqueJoinID, tmp_connectState(i).ip, joinerIdentityHash, joinerPublicKeyB64, joinerNameB64, s)
    Else
        ' if async join approval is not enabled, go directly to pending success
        tmp_connectState(i).connectState = TmpConnectState_ProcessJoin
    End If
    tmp_connectState(i).connectTime = steadyNowMs() ' reset connect time
End Sub

' (the lobby's join check result: result has_value() is hasResult; pass; the error's code and message)
Public Sub NETtempSocketHandleLobbyJoinCheckResult(ByVal i As Long, ByVal uniqueJoinID As String, ByVal hasResult As Boolean, ByVal pass As Boolean, _
        ByVal errCode As String, ByVal errMessage As String)
    If tmp_connectState(i).uniqueJoinID <> uniqueJoinID Then
        ' received response for old join attempt - discard
        Exit Sub
    End If

    If Not (tmp_connectState(i).connectState = TmpConnectState_PendingLobbyJoinCheckResult) Then ASSERT_FAILURE "(tmp_connectState[i].connectState == TmpSocketInfo::TmpConnectState::PendingLobbyJoinCheckResult)", _
        "lib/netplay/netplay.cpp:3592", "NETtempSocketHandleLobbyJoinCheckResult", "Unexpected connectState: %d", tmp_connectState(i).connectState: Exit Sub
    If tmp_socket(i) Is Nothing Then ASSERT_FAILURE "tmp_socket[i] != nullptr", "lib/netplay/netplay.cpp:3593", "NETtempSocketHandleLobbyJoinCheckResult", "No socket?"

    If Not hasResult Then
        ' there was an error checking the join request - reject with general message
        debug_ 3598, LOG_INFO, "NETtempSocketHandleLobbyJoinCheckResult", "Error checking join request %s: %s", errCode, errMessage
        NETrejectTempSocketClient i, ERROR_HOSTDROPPED, False
        Exit Sub
    End If

    If pass Then
        ' proceed to handling join approval
        NETtempSocketBeginHandleJoinApproval i
    Else
        ' lobby join check failed
        NETaddSessionBanBadIP tmp_connectState(i).ip, 4
        NETrejectTempSocketClient i, ERROR_WRONGDATA, False
    End If
End Sub

' netlobby::ConnectionMethod lobbyConnectionMethodFrom(ConnectionProviderType type): TCP_DIRECT (0) only here
Private Function lobbyConnectionMethodFrom(ByVal typ As Long) As Long
    lobbyConnectionMethodFrom = 0               ' netlobby::ConnectionMethod::TCP_DIRECT
End Function

' static std::vector<netlobby::ConnectionInfo> buildHostingConnectionDescriptions() (the vector as a JSON array of
' netlobby.cpp's to_json(ConnectionInfo), for the exe's netlobby: {"type":"tcp:ip4","host":"","port":2100})
Private Function buildHostingConnectionDescriptions() As String
    Dim result As nlohmann_json, conn As nlohmann_json, connectionType As Long, supportedProtocols As Long, port As Long
    Set result = json_array(True)

    If Not server_listen_socket Is Nothing Then
        connectionType = lobbyConnectionMethodFrom(server_listen_socket.type_())
        supportedProtocols = server_listen_socket.supportedIpVersions()
        If (supportedProtocols And 1) <> 0 Then             ' IListenSocket::IPVersions::IPV4
            Set conn = json_object(True)
            conn.set_item_ref "type", json_value(lobbyConnectionMethodName(connectionType) & ":ip4", True)   ' conn.type.method, IPVersion::IPv4
            conn.set_item_ref "host", json_value(knownExternalIPv4Info.host, True)   ' may be empty, which is fine
            If knownExternalIPv4Info.port > 0 Then port = knownExternalIPv4Info.port Else port = NETgetGameserverPort()
            conn.set_item_ref "port", json_value_unsigned(port And &HFFFF&, True)

            result.push_back_ref conn
        End If
        If (supportedProtocols And 2) <> 0 Then             ' IListenSocket::IPVersions::IPV6
            If Not netGame.warGetLobbyDisableIPv6() Then
                Set conn = json_object(True)
                conn.set_item_ref "type", json_value(lobbyConnectionMethodName(connectionType) & ":ip6", True)
                conn.set_item_ref "host", json_value("", True)
                conn.set_item_ref "port", json_value_unsigned(NETgetGameserverPort() And &HFFFF&, True)
                result.push_back_ref conn
            Else
                debug_ 3703, LOG_INFO, "buildHostingConnectionDescriptions", "Disabling listing IPv6 address due to ""Enable IPv6"" lobby setting"
            End If
        End If
    End If

    buildHostingConnectionDescriptions = result.dump(-1, " ", False, error_handler_t_replace)
End Function

' (VB) netlobby.cpp's to_string(ConnectionMethod)
Private Function lobbyConnectionMethodName(ByVal m As Long) As String
    If m = 1 Then lobbyConnectionMethodName = "gns" Else lobbyConnectionMethodName = "tcp"
End Function

Private Function allowNonLobbyMediatedJoin(ByVal ipAddress As String) As Boolean
    Dim ipClass As Long
    If lobbyDirectJoinOption = LobbyHostDirectJoinOption_All Then
        allowNonLobbyMediatedJoin = True
        Exit Function
    End If
    If lobbyDirectJoinOption = LobbyHostDirectJoinOption_None Then
        allowNonLobbyMediatedJoin = False
        Exit Function
    End If

    ipClass = GetIPClassification(ipAddress)
    If ipClass < 0 Then
        debug_ 3644, LOG_ERROR, "allowNonLobbyMediatedJoin", "Failed to get ip classification for: %s", ipAddress
        allowNonLobbyMediatedJoin = True
        Exit Function
    End If

    Select Case lobbyDirectJoinOption
    Case LobbyHostDirectJoinOption_None
        allowNonLobbyMediatedJoin = False   ' handled above, but included here for compile-time check of complete enum handling
    Case LobbyHostDirectJoinOption_Local
        Select Case ipClass
        Case IPClassification_PrivateOrULA, IPClassification_LinkLocal, IPClassification_Loopback
            allowNonLobbyMediatedJoin = True
        Case Else
            allowNonLobbyMediatedJoin = False
        End Select
    Case LobbyHostDirectJoinOption_All
        allowNonLobbyMediatedJoin = True    ' handled above, but included here for compile-time check of complete enum handling
    Case Else
        allowNonLobbyMediatedJoin = False   ' silence compiler warning
    End Select
End Function

' optional<size_t> firstAvailableTempSocketIdx(): the result is has_value(), the index in idx
Private Function firstAvailableTempSocketIdx(idx As Long) As Boolean
    Dim i As Long
    ' Find the first empty socket slot
    For i = 0 To MAX_TMP_SOCKETS - 1
        If tmp_socket(i) Is Nothing Then
            idx = i
            firstAvailableTempSocketIdx = True
            Exit Function
        End If
    Next
    firstAvailableTempSocketIdx = False
End Function

' ////////////////////////////////////////////////////////////////////////
' Host a game with a given name and player name. & 4 user game flags
Private Sub NETallowJoining()
    Dim i As Long, result As Long, connectFailed As Boolean, major As Long, minor As Long
    Dim ec As Long, n As Long, rawDummy As Long, readable As Long, w As MessageWriter, r As MessageReader, r2 As MessageReader
    Dim serverPublicKey() As Byte, rejected As Long, buffer() As Byte, size As Long, hostKey() As Byte
    Dim jsonStringLen As Long, jsonBytes() As Byte
    Dim name_ As String, ModList As String, GamePassword As String, playerType As Long, pkey() As Byte, identity As EcKey
    Dim encryptedChallengeResponse() As Byte, joinToken As String, rejectMsg As String, buf As String
    Dim connectPerm As Long, hasConnectPermissions As Boolean, decryptedMessageRawData() As Byte, decryptedLen As Long
    Dim tmpMessageBuilder As NetMessageBuilder, challengeResponse() As Byte, connectionDescriptionSerializedBytes() As Byte
    Dim challengeForHost() As Byte, uniqueJoinIDCopy As String, currentSteadTime As Double, timeout As Double
    Dim previousPlayersShouldCheckReadyValue As Boolean, tmp As Long, hasTmp As Boolean, index As Long
    Dim joinRequestInfo As ReceivedJoinInfo, encryptedHostChallengeResponse() As Byte, encLen As Long, hostChallengeResponse() As Byte
    Dim sigLen As Long, localKey As EcKey, pPlayerType As String, blindModeVal As Long, j As Long, rIP As String, tmpQueue As Long
    Dim maxrejectlen As Long, b() As Byte, s As String, joinerPublicKeyB64 As String, joinerIdentityHash As String, joinerNameB64 As String
    Dim lobbyCheck As Long, jsonString As String, lobbyResponse As String, lobbyResponseError As Long, lobbyResponseLen As Long
    Dim lobbyResponseBytes() As Byte, responseLength As Long, lobbyRespBuffer() As Byte

    If Not NetPlay.isHost Then ASSERT_FAILURE "false", "lib/netplay/netplay.cpp:3728", "NETallowJoining", "Host only routine detected for client!": Exit Sub

    connectFailed = True

    If allow_joining = False Then
        Exit Sub
    End If

    If Not lobbyHostingHandler_exists() Then
        ' First time attempting lobby connection for this listing - create lobbyHostingHandler
        NETcreateLobbyHostingHandler

        ' Attempt to create game listing
        If lobby_disabled Then
            lobbyCheck = 0
        Else
            ' createGameListing(getMultiStats(realSelectedPlayer).identity, gameDetails, hostJoinOptions, buildHostingConnectionDescriptions())
            If lobbyHostingHandler_createGameListing(realSelectedPlayer, buildHostingConnectionDescriptions()) Then lobbyCheck = 1 Else lobbyCheck = 0
        End If
        If lobbyCheck = 0 Then
            handleLobbyHostingStatusIsKnown False, ""

            If lobby_disabled Then
                debug_ 3752, LOG_INFO, "NETallowJoining", "Multiplayer lobby support unavailable. Please update your client."
                ShowLobbyStatusMessage1 "Multiplayer lobby support unavailable. Please update your client."

                netGame.wzCommandInterfaceOutput "WZEVENT: lobbyerror: Client support disabled / unavailable" & vbLf
            Else
                debug_ 3759, LOG_INFO, "NETallowJoining", "Failed to create game listing"
                ShowLobbyStatusMessage1 va("%s" & vbLf & "%s", "Game not listed in the lobby:", "Failed to create game listing. Check WZ logs.")

                netGame.wzCommandInterfaceOutput "WZEVENT: lobbyerror: Failed to create game listing" & vbLf
            End If
        End If
    End If

    If tmp_socket_set Is Nothing Then ASSERT_FAILURE "tmp_socket_set != nullptr", "lib/netplay/netplay.cpp:3767", "NETallowJoining", "Null tmp_socket_set"
    readable = tmp_socket_set.checkConnectionsReadable(NET_READ_TIMEOUT, ec)
    If ec <> 0 Then readable = 0
    If readable > 0 Then
        For i = 0 To MAX_TMP_SOCKETS - 1
            If tmp_socket(i) Is Nothing Then
                GoTo continue1
            End If

            If Not tmp_socket(i).readReady() Then
                GoTo continue1
            End If

            If tmp_connectState(i).connectState = TmpConnectState_PendingInitialConnect Then
                ReDim b(0 To 7)
                n = tmp_socket(i).readNoInt(b, 0, 8 - tmp_connectState(i).usedBuffer, rawDummy, ec)
                If ec = 0 Then
                    For j = 0 To n - 1
                        tmp_connectState(i).buffer(tmp_connectState(i).usedBuffer + j) = b(j)
                    Next
                    tmp_connectState(i).usedBuffer = tmp_connectState(i).usedBuffer + n
                End If

                ' A 2.3.7 client sends a "list" command first, just drop the connection.
                If tmp_connectState(i).usedBuffer >= 5 And bufferIsList(i) Then
                    debug_ 3795, LOG_INFO, "NETallowJoining", "An old client tried to connect, closing the socket."
                    NETlogEntry "Dropping old client.", SYNC_FLAG, i
                    NETlogEntry "Invalid (old)game version", SYNC_FLAG, i
                    NETaddSessionBanBadIP tmp_connectState(i).ip
                    connectFailed = True
                ElseIf tmp_connectState(i).usedBuffer >= 8 Then
                    ' New clients send NETCODE_VERSION_MAJOR and NETCODE_VERSION_MINOR
                    ' Check these numbers with our own.

                    major = be32(tmp_connectState(i).buffer, 0)
                    minor = be32(tmp_connectState(i).buffer, 4)

                    If major = 0 And minor = 0 Then
                        ' special case for lobby server "alive" check
                        tmp_connectState(i).connectState = TmpConnectState_PendingLobbyConnectCheckRequest
                        tmp_connectState(i).connectTime = steadyNowMs()     ' reset connect time

                        ' Continue to wait (until timeout) for the full initial lobby connect check request message
                        GoTo continue1
                    ElseIf NETisCorrectVersion(major, minor) Then
                        result = ERROR_NOERROR
                        ReDim b(0 To 3)
                        putBe32 b, 0, result
                        For j = 0 To 3
                            tmp_connectState(i).buffer(j) = b(j)
                        Next
                        tmp_socket(i).writeAll b, 0, 4, rawDummy, ec
                        If ec <> 0 Then
                            debug_ 3828, LOG_NET, "NETallowJoining", "writeAll to tmpSocket[%u] failed with error?: %d", i, net_errorValue(ec)
                        End If
                        tmp_socket(i).enableCompression

                        ' Connection is successful.
                        connectFailed = False

                        ' Give client a challenge to solve before connecting
                        tmp_connectState(i).connectChallengeSize = NETgetJoinConnectionNETPINGChallengeFromHostSize()
                        ReDim tmp_connectState(i).connectChallenge(0 To tmp_connectState(i).connectChallengeSize - 1)
                        genSecRandomBytes tmp_connectState(i).connectChallenge, 0, tmp_connectState(i).connectChallengeSize
                        Set w = NETbeginEncode(NETnetTmpQueue(i), NET_PING)
                        NETbytes_w w, tmp_connectState(i).connectChallenge
                        ' Send the server's public identity
                        ' - As long as host is a spectator host, or this is a regular (non-blind) game, this is always the host's actual public identity
                        ' - If host is a player-host *and* this is a blind game, it's the blind identity
                        n = getLocalSharedIdentityKey(localKey)
                        n = EcKey_toBytes(localKey, EcKey_Public, serverPublicKey)
                        If n = 0 Then Erase serverPublicKey
                        NETbytes_w w, serverPublicKey
                        NETend_w w
                    Else
                        debug_ 3849, LOG_INFO, "NETallowJoining", "Received an invalid version ""%u.%u"".", major, minor
                        result = ERROR_WRONGVERSION
                        ReDim b(0 To 3)
                        putBe32 b, 0, result
                        For j = 0 To 3
                            tmp_connectState(i).buffer(j) = b(j)
                        Next
                        tmp_socket(i).writeAll b, 0, 4, rawDummy, ec
                        If ec <> 0 Then
                            debug_ 3855, LOG_NET, "NETallowJoining", "writeAll to tmpSocket[%u] failed with error?: %d", i, net_errorValue(ec)
                        End If
                        NETlogEntry "Invalid game version", SYNC_FLAG, i
                        NETaddSessionBanBadIP tmp_connectState(i).ip
                        connectFailed = True
                    End If
                    If (Not connectFailed) And (Not NET_HasAnyOpenSlots()) Then
                        ' early player count test, in case they happen to get in before updates.
                        ' Tell the player that we are completely full.
                        rejected = ERROR_FULL
                        Set w = NETbeginEncode(NETnetTmpQueue(i), NET_REJECTED)
                        NETuint8_t_w w, rejected
                        NETend_w w
                        NETflush
                        connectFailed = True
                    End If
                Else
                    ' Continue to wait (until timeout) for the full initial connect message
                    GoTo continue1
                End If

                ' Remove a failed connection.
                If connectFailed Then
                    debug_ 3882, LOG_NET, "NETallowJoining", "freeing temp socket %p (%d)", ObjPtr(tmp_socket(i)), 3871
                    NETcloseTempSocket i
                Else
                    ' on initial connect success...
                    tmp_connectState(i).connectState = TmpConnectState_PendingJoinRequest
                    tmp_connectState(i).connectTime = steadyNowMs()     ' reset connect time
                End If
                GoTo continue1
            ElseIf tmp_connectState(i).connectState = TmpConnectState_PendingJoinRequest Or tmp_connectState(i).connectState = TmpConnectState_PendingLobbyConnectCheckRequest Then
                ReDim buffer(0 To MaxMsgSize - 1)
                size = tmp_socket(i).readNoInt(buffer, 0, MaxMsgSize, rawDummy, ec)
                rejected = 0

                If ec <> 0 Then
                    ' disconnect or programmer error
                    If net_errorEqualsErrc(ec, errc_timed_out) Or net_errorEqualsErrc(ec, errc_connection_reset) Then
                        debug_ 3905, LOG_NET, "NETallowJoining", "Client socket disconnected."
                    Else
                        debug_ 3910, LOG_NET, "NETallowJoining", "Client socket encountered error: %s", net_errorMessage(ec)
                    End If
                    NETlogEntry "Client socket disconnected (allowJoining)", SYNC_FLAG, i
                    debug_ 3913, LOG_NET, "NETallowJoining", "freeing temp socket %p (%d)", ObjPtr(tmp_socket(i)), 3899
                    NETcloseTempSocket i
                    GoTo continue1
                End If
                NETinsertRawData NETnetTmpQueue(i), buffer, size

                If Not NETisMessageReady(NETnetTmpQueue(i)) Then
                    ' need to wait for a full and complete join message
                    ' sanity check
                    If NETincompleteMessageDataBuffered(NETnetTmpQueue(i)) > (MaxMsgSize * 8) Then  ' something definitely big enough to encompass the expected message(s) at this point
                        ' client is sending data that doesn't appear to be a properly formatted message - cut it off
                        NETpop NETnetTmpQueue(i)
                        NETrejectTempSocketClient i, ERROR_WRONGDATA, True
                    End If
                    GoTo continue1
                End If

                If tmp_connectState(i).connectState = TmpConnectState_PendingLobbyConnectCheckRequest And NETgetMessage(NETnetTmpQueue(i)).type_() = 0 Then  ' special lobby connection check request
                    jsonStringLen = 0
                    Set r = NETbeginDecode(NETnetTmpQueue(i), 0)
                    NETuint16_t_r r, jsonStringLen
                    If jsonStringLen > 0 Then ReDim jsonBytes(0 To jsonStringLen - 1)
                    NETbin_r r, jsonBytes, jsonStringLen
                    NETend_r r
                    NETpop NETnetTmpQueue(i)

                    If lobbyHostingHandler_exists() Then
                        If jsonStringLen > 0 Then jsonString = utf8_decode(jsonBytes, 0, jsonStringLen) Else jsonString = ""
                        If lobbyHostingHandler_respondToLobbyConnectionCheckRequest(jsonString, gameDetails.host_name, realSelectedPlayer, lobbyResponse, lobbyResponseError) Then
                            ' output response to lobby connection check
                            lobbyResponseLen = utf8_encode(lobbyResponse, lobbyResponseBytes)
                            If Not (lobbyResponseLen <= 65535) Then ASSERT_FAILURE "response.value().size() <= 65535", "lib/netplay/netplay.cpp:3954", "NETallowJoining", "response length is too long! (%zu)", lobbyResponseLen
                            If lobbyResponseLen < 65535 Then responseLength = lobbyResponseLen Else responseLength = 65535

                            ReDim lobbyRespBuffer(0 To 4 + 4 + 4 + responseLength - 1)

                            ' Copy response prefix chars ("WZLR")
                            lobbyRespBuffer(0) = Asc("W")
                            lobbyRespBuffer(1) = Asc("Z")
                            lobbyRespBuffer(2) = Asc("L")
                            lobbyRespBuffer(3) = Asc("R")

                            ' Copy version of response
                            putBe32 lobbyRespBuffer, 4, 2       ' const uint32_t response_version = 2

                            ' Copy length of response (as uint32)
                            putBe32 lobbyRespBuffer, 8, responseLength

                            ' Copy response
                            For j = 0 To responseLength - 1
                                lobbyRespBuffer(12 + j) = lobbyResponseBytes(j)
                            Next

                            If Not (Not tmp_socket(i).isCompressed()) Then ASSERT_FAILURE "!tmp_socket[i]->isCompressed()", "lib/netplay/netplay.cpp:3981", "NETallowJoining", "Socket should not be compressed!"
                            tmp_socket(i).writeAll lobbyRespBuffer, 0, 12 + responseLength, rawDummy, ec
                            If ec <> 0 Then
                                debug_ 3985, LOG_NET, "NETallowJoining", "writeAll to tmpSocket[%u] failed with error?: %d", i, net_errorValue(ec)
                            End If
                        Else
                            ' failed to process request
                            Select Case lobbyResponseError
                            Case 0                      ' LobbyConnectionCheckRequestError::GameNotListed
                                ' ignore
                            Case 1                      ' LobbyConnectionCheckRequestError::InvalidRequest
                                NETaddSessionBanBadIP tmp_connectState(i).ip
                            End Select
                        End If
                    End If

                    debug_ 4003, LOG_NET, "NETallowJoining", "freeing temp socket %p (%d)", ObjPtr(tmp_socket(i)), 3984
                    NETcloseTempSocket i

                ElseIf tmp_connectState(i).connectState = TmpConnectState_PendingJoinRequest And NETgetMessage(NETnetTmpQueue(i)).type_() = NET_JOIN Then
                    name_ = ""
                    ModList = ""
                    GamePassword = ""
                    playerType = 0
                    Erase pkey
                    EcKey_clear identity
                    Erase encryptedChallengeResponse
                    joinToken = ""

                    Set r = NETbeginDecode(NETnetTmpQueue(i), NET_JOIN)
                    NETstring_r r, name_, StringSize
                    NETstring_r r, ModList, modlist_string_size
                    NETstring_r r, GamePassword, password_string_size
                    NETuint8_t_r r, playerType
                    NETbytes_r r, pkey
                    NETbytes_r r, encryptedChallengeResponse
                    NETstring_std_r r, joinToken
                    NETend_r r
                    NETpop NETnetTmpQueue(i)

                    ' Determine if it's a valid public identity
                    If Not EcKey_fromBytes(identity, pkey, byteVecSize(pkey), EcKey_Public) Then
                        ' Invalid public identity provided - just reject
                        rejectMsg = va("**Rejecting player(%s), invalid player identity.", tmp_connectState(i).ip)
                        debug_ 4037, LOG_INFO, "NETallowJoining", "%s", rejectMsg
                        debug_ 4038, LOG_NET, "NETallowJoining", "freeing temp socket %p, invalid player identity", ObjPtr(tmp_socket(i))

                        NETrejectTempSocketClient i, ERROR_WRONGDATA, True
                        GoTo continue1
                    End If

                    ' Do an initial pass of blacklist checks (even though we haven't verified the identity yet)
                    hasConnectPermissions = netPermissionsCheck_ConnectKey(identity, connectPerm)
                    If (hasConnectPermissions And connectPerm = ConnectPermissions_Blocked) Or (Not hasConnectPermissions And onBanList(tmp_connectState(i).ip)) Then
                        buf = sstrcpyN(va("** A player that you have kicked tried to rejoin the game, and was rejected. IP: %s", tmp_connectState(i).ip), 256)
                        debug_ 4037, LOG_INFO, "NETallowJoining", "%s", buf
                        NETlogEntry buf, SYNC_FLAG, i

                        NETrejectTempSocketClient i, ERROR_KICKED, False
                        GoTo continue1
                    End If

                    ' Save join info in the tmp_connectState
                    tmp_connectState(i).receivedJoinInfo.name = sstrcpyN(name_, 64)
                    tmp_connectState(i).receivedJoinInfo.playerType = playerType
                    tmp_connectState(i).receivedJoinInfo.identity = identity

                    ' Construct the auth session keys
                    n = getLocalSharedIdentityKey(localKey)
                    If Not SessionKeys_init(tmp_connectState(i).receivedJoinInfo.connectionAuthSessionKeys, localKey, 0, identity, 1) Then
                        rejectMsg = va("**Rejecting player(%s), failed to establish session keys, error: %s", tmp_connectState(i).ip, SessionKeys_lastError())
                        debug_ 4051, LOG_INFO, "NETallowJoining", "%s", rejectMsg
                        debug_ 4072, LOG_NET, "NETallowJoining", "freeing temp socket %p, couldn't establish session keys", ObjPtr(tmp_socket(i))
                        NETrejectTempSocketClient i, ERROR_WRONGDATA, False
                        GoTo continue1
                    End If
                    tmp_connectState(i).receivedJoinInfo.has_connectionAuthSessionKeys = True

                    ' Decrypt the encryptedChallengeResponse
                    If Not SessionKeys_decryptMessageFromOther(tmp_connectState(i).receivedJoinInfo.connectionAuthSessionKeys, encryptedChallengeResponse, 0, _
                            byteVecSize(encryptedChallengeResponse), decryptedMessageRawData, decryptedLen) Then
                        rejectMsg = va("**Rejecting player(%s), failed to decrypt player auth data", tmp_connectState(i).ip)
                        debug_ 4071, LOG_INFO, "NETallowJoining", "%s", rejectMsg
                        debug_ 4082, LOG_NET, "NETallowJoining", "freeing temp socket %p, couldn't verify player identity", ObjPtr(tmp_socket(i))
                        NETrejectTempSocketClient i, ERROR_WRONGDATA, True
                        GoTo continue1
                    End If
                    Set tmpMessageBuilder = NetMessageBuilder_new(NET_JOIN, decryptedLen) ' dummy message for parsing
                    If decryptedLen > 0 Then tmpMessageBuilder.append_bytes decryptedMessageRawData, 0, decryptedLen
                    NETinsertMessageFromNet NETnetTmpQueue(i), tmpMessageBuilder.build() ' insert virtual message into temp queue for parsing

                    ' Parse the decrypted response
                    Erase challengeResponse
                    Erase connectionDescriptionSerializedBytes
                    ReDim challengeForHost(0 To NETgetJoinConnectionNETPINGChallengeFromClientSize() - 1)
                    Set r2 = NETbeginDecode(NETnetTmpQueue(i), NET_JOIN)
                    NETbytes_r r2, challengeResponse
                    NETbytes_r r2, connectionDescriptionSerializedBytes
                    NETbytes_r r2, challengeForHost
                    NETend_r r2
                    NETpop NETnetTmpQueue(i)

                    ' Verify signature that player is joining with - reject on failure
                    If Not EcKey_verify(identity, challengeResponse, byteVecSize(challengeResponse), tmp_connectState(i).connectChallenge, 0, tmp_connectState(i).connectChallengeSize) Then
                        rejectMsg = va("**Rejecting player(%s), failed to verify player identity.", tmp_connectState(i).ip)
                        debug_ 4108, LOG_INFO, "NETallowJoining", "%s", rejectMsg
                        debug_ 4108, LOG_NET, "NETallowJoining", "freeing temp socket %p, couldn't verify player identity", ObjPtr(tmp_socket(i))
                        NETrejectTempSocketClient i, ERROR_WRONGDATA, True
                        GoTo continue1
                    End If

                    ' Verify that the challengeForHost is expected length
                    If byteVecSize(challengeForHost) <> 0 And byteVecSize(challengeForHost) <> NETgetJoinConnectionNETPINGChallengeFromClientSize() Then
                        rejectMsg = va("**Rejecting player(%s), invalid host challenge length.", tmp_connectState(i).ip)
                        debug_ 4118, LOG_INFO, "NETallowJoining", "%s", rejectMsg
                        debug_ 4118, LOG_NET, "NETallowJoining", "freeing temp socket %p, invalid host challenge length", ObjPtr(tmp_socket(i))
                        NETrejectTempSocketClient i, ERROR_WRONGDATA, True
                        GoTo continue1
                    End If
                    tmp_connectState(i).receivedJoinInfo.challengeForHost = challengeForHost
                    tmp_connectState(i).receivedJoinInfo.challengeForHostSize = byteVecSize(challengeForHost)
                    Erase challengeForHost

                    ' Additional rejection checks
                    If tmp_connectState(i).receivedJoinInfo.playerType <> NET_JOIN_SPECTATOR And Not bAsyncJoinApprovalEnabled And playerManagementRecord_hostMovedPlayerToSpectators_ip(tmp_connectState(i).ip) Then
                        ' The host previously relegated a player from this IP address to Spectators (this game), and it seems they are trying to rejoin as a Player - deny this
                        buf = sstrcpyN(va("** A player that you moved to Spectators tried to rejoin the game (as a Player), and was rejected. IP: %s", tmp_connectState(i).ip), 256)
                        debug_ 4132, LOG_INFO, "NETallowJoining", "%s", buf
                        NETlogEntry buf, SYNC_FLAG, i

                        ' Player has been kicked before, kick again.
                        rejected = ERROR_KICKED
                    ElseIf NetPlay.GamePassworded And NetPlay.gamePassword <> GamePassword Then
                        ' Wrong password. Reject.
                        rejected = ERROR_WRONGPASSWORD
                    ElseIf Not ULT(NetPlay.playercount, gameDetails.players_max) Then
                        ' Game full. Reject.
                        rejected = ERROR_FULL
                    End If

                    If rejected <> 0 Then
                        buf = sstrcpyN(va("**Rejecting player(%s), reason (%u).", tmp_connectState(i).ip, rejected), 256)
                        debug_ 4153, LOG_INFO, "NETallowJoining", "%s", buf
                        NETlogEntry buf, SYNC_FLAG, i
                        Set w = NETbeginEncode(NETnetTmpQueue(i), NET_REJECTED)
                        NETuint8_t_w w, rejected
                        NETstring_w w, "", 0
                        NETend_w w
                        NETflush

                        NETcloseTempSocket i
                        sync_counter.cantjoin = (sync_counter.cantjoin + 1) And &HFFFF&
                        GoTo continue1
                    End If

                    ' on passing all built-in checks for connect...

                    If joinToken = "" Then
                        If Not lobbyHostingHandler_exists() Then
                            NETtempSocketBeginHandleJoinApproval i
                        ElseIf Not lobbyHostingHandler_isConnectedToLobby() Then
                            NETtempSocketBeginHandleJoinApproval i
                        ElseIf allowNonLobbyMediatedJoin(tmp_connectState(i).ip) Then
                            NETtempSocketBeginHandleJoinApproval i
                        Else
                            NETaddSessionBanBadIP tmp_connectState(i).ip, 4
                            NETrejectTempSocketClient i, ERROR_WRONGDATA, False
                        End If
                    Else
                        ' Check joinToken
                        If lobbyHostingHandler_exists() And lobbyHostingHandler_isConnectedToLobby() Then
                            ' use the connectChallenge as a unique, random join ID
                            tmp_connectState(i).uniqueJoinID = base64Encode(tmp_connectState(i).connectChallenge, tmp_connectState(i).connectChallengeSize)
                            uniqueJoinIDCopy = tmp_connectState(i).uniqueJoinID

                            ' netlobby::JoinRequest joinRequest (its fields: joinToken, ip, playerName, playerPublicIdentity,
                            ' connectionMethod, asSpectator)
                            n = EcKey_toBytes(tmp_connectState(i).receivedJoinInfo.identity, EcKey_Public, b)
                            If lobbyHostingHandler_checkJoinRequest(joinToken, tmp_connectState(i).ip, name_, base64Encode(b, n), _
                                    lobbyConnectionMethodFrom(server_listen_socket.type_()), (tmp_connectState(i).receivedJoinInfo.playerType = NET_JOIN_SPECTATOR), _
                                    i, uniqueJoinIDCopy) Then
                                ' dispatch check join request
                                tmp_connectState(i).connectState = TmpConnectState_PendingLobbyJoinCheckResult
                                tmp_connectState(i).connectTime = steadyNowMs()     ' reset connect time
                            Else
                                ' failure to check join request, despite having lobby connection - internal logic error
                                ASSERT_FAILURE "false", "lib/netplay/netplay.cpp:4208", "NETallowJoining", "Failure to issue check join request"
                                NETrejectTempSocketClient i, ERROR_WRONGDATA, False
                            End If
                        Else
                            ' No lobbyHostingHandler
                            NETtempSocketBeginHandleJoinApproval i
                        End If
                    End If
                Else
                    ' unexpected message type at this time
                    ' reject the bad client
                    NETpop NETnetTmpQueue(i)
                    NETrejectTempSocketClient i, ERROR_WRONGDATA, True
                End If
                GoTo continue1
            End If
continue1:
        Next
    End If

    currentSteadTime = steadyNowMs()

    For i = 0 To MAX_TMP_SOCKETS - 1
        If tmp_socket(i) Is Nothing Then
            GoTo continue2
        End If

        If tmp_connectState(i).connectState = TmpConnectState_PendingAsyncApproval Then
            ' check if async approval result has been set
            If tmp_connectState(i).has_asyncJoinApprovalResult Then
                If tmp_connectState(i).asyncJoinApprovalResult <> AsyncJoinApprovalAction_Reject Then
                    ' async join approved - process the join
                    tmp_connectState(i).connectState = TmpConnectState_ProcessJoin
                    tmp_connectState(i).connectTime = steadyNowMs()     ' reset connect time
                    If tmp_connectState(i).asyncJoinApprovalResult = AsyncJoinApprovalAction_ApproveSpectators Then
                        ' change the player join request to spectators
                        tmp_connectState(i).receivedJoinInfo.playerType = NET_JOIN_SPECTATOR
                        ' enforce spectator state for this player
                        n = EcKey_toBytes(tmp_connectState(i).receivedJoinInfo.identity, EcKey_Public, b)
                        playerManagementRecord_movedPlayerToSpectators tmp_connectState(i).ip, b, n, True
                    ElseIf tmp_connectState(i).asyncJoinApprovalResult = AsyncJoinApprovalAction_Approve Then
                        ' clear any enforced spectator state for this player
                        n = EcKey_toBytes(tmp_connectState(i).receivedJoinInfo.identity, EcKey_Public, b)
                        playerManagementRecord_movedSpectatorToPlayers tmp_connectState(i).ip, b, n, True
                    End If
                    ' deliberately fall-through to the TmpSocketInfo::TmpConnectState::ProcessJoin condition further below
                Else
                    ' async join rejected
                    rejected = tmp_connectState(i).asyncJoinRejectCode And &HFF&
                    buf = sstrcpyN(va("**Rejecting player(%s), due to async approval rejection, reason (%u).", tmp_connectState(i).ip, rejected), 256)
                    debug_ 4270, LOG_INFO, "NETallowJoining", "%s", buf
                    NETlogEntry buf, SYNC_FLAG, i
                    Set w = NETbeginEncode(NETnetTmpQueue(i), NET_REJECTED)
                    NETuint8_t_w w, rejected
                    maxrejectlen = Len(tmp_connectState(i).asyncJoinRejectCustomMessage) + 1
                    If maxrejectlen > MAX_JOIN_REJECT_REASON Then maxrejectlen = MAX_JOIN_REJECT_REASON
                    NETstring_w w, tmp_connectState(i).asyncJoinRejectCustomMessage, maxrejectlen

                    If rejected = ERROR_REDIRECT Then
                        ' construct + send encrypted client challenge response
                        Erase encryptedHostChallengeResponse
                        If tmp_connectState(i).receivedJoinInfo.challengeForHostSize > 0 Then
                            n = getLocalSharedIdentityKey(localKey)
                            sigLen = EcKey_sign(localKey, tmp_connectState(i).receivedJoinInfo.challengeForHost, 0, tmp_connectState(i).receivedJoinInfo.challengeForHostSize, hostChallengeResponse)
                            encLen = SessionKeys_encryptMessageForOther(tmp_connectState(i).receivedJoinInfo.connectionAuthSessionKeys, hostChallengeResponse, 0, sigLen, encryptedHostChallengeResponse)
                            If encLen = 0 Then
                                Erase encryptedHostChallengeResponse
                                debug_ 4288, LOG_INFO, "NETallowJoining", "Failed to encrypt response?"
                            End If
                        End If
                        NETbytes_w w, encryptedHostChallengeResponse
                    End If

                    NETend_w w
                    NETflush
                    tmpQueue = NETnetTmpQueue(i)
                    If NETisMessageReady(tmpQueue) Then
                        NETpop tmpQueue
                    End If

                    NETcloseTempSocket i
                    sync_counter.cantjoin = (sync_counter.cantjoin + 1) And &HFFFF&
                    GoTo continue2 ' continue to next tmp_socket
                End If
            Else
                ' if no async approval, do a timeout check
                timeout = 8000 ' although the client no longer blocks on join, a response should still be relatively prompt (the client has its own timeout)
                If currentSteadTime - tmp_connectState(i).connectTime > timeout Then
                    debug_ 4313, LOG_INFO, "NETallowJoining", "Freeing temp socket %p due to async connection approval timeout (IP: %s)", ObjPtr(tmp_socket(i)), tmp_connectState(i).ip

                    rejected = ERROR_HOSTDROPPED
                    Set w = NETbeginEncode(NETnetTmpQueue(i), NET_REJECTED)
                    NETuint8_t_w w, rejected
                    NETstring_w w, "", 0
                    NETend_w w
                    NETflush
                    tmpQueue = NETnetTmpQueue(i)
                    If NETisMessageReady(tmpQueue) Then
                        NETpop tmpQueue
                    End If
                    NETcloseTempSocket i
                End If
                GoTo continue2 ' continue to next tmp_socket
            End If
        End If
        ' note: *not* an "else if" because the condition above may transition the state to ProcessJoin!
        If tmp_connectState(i).connectState = TmpConnectState_ProcessJoin Then
            previousPlayersShouldCheckReadyValue = netGame.multiplayPlayersShouldCheckReady()
            hasTmp = False

            If tmp_connectState(i).has_asyncJoinApprovalExplicitPlayerIdx Then
                ' Try to allocate the new player at the explicit index
                If NET_IsSlotOpenForPlayerJoin(tmp_connectState(i).asyncJoinApprovalExplicitPlayerIdx) Then
                    If NET_CreatePlayerAtIdx(tmp_connectState(i).asyncJoinApprovalExplicitPlayerIdx, tmp_connectState(i).receivedJoinInfo.name) Then
                        tmp = tmp_connectState(i).asyncJoinApprovalExplicitPlayerIdx
                        hasTmp = True
                    End If
                End If
                If Not hasTmp Then
                    debug_ 4347, LOG_INFO, "NETallowJoining", "Async join approval provided explicit player index (%u), but it's occupied", tmp_connectState(i).asyncJoinApprovalExplicitPlayerIdx
                End If
            Else
                ' Try to allocate the new player an open slot
                If (tmp_connectState(i).receivedJoinInfo.playerType = NET_JOIN_SPECTATOR) Or ULT(NetPlay.playercount, gameDetails.players_max) Then
                    hasTmp = NET_CreatePlayer(tmp, tmp_connectState(i).receivedJoinInfo.name, False, True, (tmp_connectState(i).receivedJoinInfo.playerType = NET_JOIN_SPECTATOR))
                End If
            End If

            If Not hasTmp Or tmp > 255 Then
                If hasTmp Then
                    If Not (tmp <= 255) Then ASSERT_FAILURE "tmp.value_or(0) <= static_cast<uint32_t>(std::numeric_limits<uint8_t>::max())", "lib/netplay/netplay.cpp:4361", "NETallowJoining", "Currently limited to uint8_t"
                End If
                debug_ 4362, LOG_INFO, "NETallowJoining", "freeing temp socket %p, couldn't create slot", ObjPtr(tmp_socket(i))

                ' Tell the player that we are full.
                NETrejectTempSocketClient i, ERROR_FULL, False
                GoTo continue2 ' continue to next tmp_socket
            End If

            index = tmp

            joinRequestInfo = tmp_connectState(i).receivedJoinInfo ' keep the join info
            TmpSocketInfo_reset i

            NEThostPromoteTempSocketToPermanentPlayerConnection i, index

            ' construct encrypted client challenge response
            Erase encryptedHostChallengeResponse
            If joinRequestInfo.challengeForHostSize > 0 Then
                n = getLocalSharedIdentityKey(localKey)
                sigLen = EcKey_sign(localKey, joinRequestInfo.challengeForHost, 0, joinRequestInfo.challengeForHostSize, hostChallengeResponse)
                encLen = SessionKeys_encryptMessageForOther(joinRequestInfo.connectionAuthSessionKeys, hostChallengeResponse, 0, sigLen, encryptedHostChallengeResponse)
                If encLen = 0 Then
                    Erase encryptedHostChallengeResponse
                    debug_ 4384, LOG_INFO, "NETallowJoining", "Failed to encrypt response?"
                End If
            End If

            ' Send NET_ACCEPTED
            Set w = NETbeginEncode(NETnetQueue(index), NET_ACCEPTED)
            NETuint8_t_w w, index
            NETuint32_t_w w, NetPlay.hostPlayer
            blindModeVal = netGame.gameBlindMode() And &HFF&
            NETuint8_t_w w, blindModeVal
            NETbytes_w w, encryptedHostChallengeResponse
            NETend_w w

            ' First send info about players to newcomer.
            NETSendAllPlayerInfoTo index
            ' then send info about newcomer to all players.
            NETBroadcastPlayerInfo index

            If NetPlay.players(index).isSpectator Then pPlayerType = "Spectator" Else pPlayerType = "Player"
            buf = sstrcpyN(va("%s[%u] %s has joined, IP is: %s", pPlayerType, index, netGame.getPlayerName(index), NetPlay.players(index).IPtextAddress), 250)
            debug_ 4407, LOG_INFO, "NETallowJoining", "%s", buf
            NETlogEntry buf, SYNC_FLAG, index

            debug_ 4410, LOG_NET, "NETallowJoining", "%s, %s, with index of %u has joined using socket %p", pPlayerType, netGame.getPlayerName(index), index, ObjPtr(connected_bsocket(index))

            ' Increment player count
            If Not NetPlay.players(index).isSpectator Then
                gameDetails.players_current = gameDetails.players_current + 1
            Else
                gameDetails.spectators_current = gameDetails.spectators_current + 1
            End If

            n = EcKey_toBytes(joinRequestInfo.identity, EcKey_Public, b)
            netGame.MultiPlayerJoin index, True, b, n

            netGame.handlePossiblePlayersShouldCheckReadyChange previousPlayersShouldCheckReadyValue

            If netGame.wzCommandInterfaceEnabled() Then
                n = EcKey_toBytes(joinRequestInfo.identity, EcKey_Public, b)
                joinerPublicKeyB64 = base64Encode(b, n)
                joinerIdentityHash = EcKey_publicHashString(joinRequestInfo.identity)
                joinerNameB64 = base64EncodeString(joinRequestInfo.name)
                netGame.wzCommandInterfaceOutput va("WZEVENT: player join: %u %s %s %s %s" & vbLf, index, joinerPublicKeyB64, joinerIdentityHash, NetPlay.players(index).IPtextAddress, joinerNameB64)
            End If

            ' Broadcast to everyone that a new player has joined
            Set w = NETbeginEncode(NETbroadcastQueue(), NET_PLAYER_JOINED)
            NETuint8_t_w w, index
            NETend_w w

            For j = 0 To MAX_CONNECTED_PLAYERS - 1
                NETBroadcastPlayerInfo j
            Next
            NETfixDuplicatePlayerNames

            ' Send the updated gameListing to the masterserver
            NETupdateLobbyGameListing

            ' reset flags for new players
            If Not NetPlay.players(index).wzFiles Is Nothing Then
                Set NetPlay.players(index).wzFiles = New Collection
            Else
                ASSERT_FAILURE "false", "lib/netplay/netplay.cpp:4458", "NETallowJoining", "wzFiles is uninitialized?? (Player: %u)", index
            End If

            netGame.wzCommandInterfaceOutputRoomStatusJson False

            GoTo continue2 ' continue to next tmp_socket
        End If

        ' in all other cases, do a timeout check
        If tmp_connectState(i).connectState = TmpConnectState_PendingInitialConnect Then timeout = 2500 Else timeout = 3000
        If tmp_connectState(i).connectState = TmpConnectState_PendingLobbyJoinCheckResult Then
            If firstAvailableTempSocketIdx(j) Then
                timeout = 10000
            ElseIf timeout < 6000 Then
                timeout = 6000
            End If
        End If

        If currentSteadTime - tmp_connectState(i).connectTime > timeout Then
            debug_ 4475, LOG_INFO, "NETallowJoining", "Freeing temp socket %p due to initial connection timeout (IP: %s)", ObjPtr(tmp_socket(i)), tmp_connectState(i).ip

            rejected = ERROR_WRONGDATA
            Set w = NETbeginEncode(NETnetTmpQueue(i), NET_REJECTED)
            NETuint8_t_w w, rejected
            NETstring_w w, "", 0
            NETend_w w
            NETflush
            tmpQueue = NETnetTmpQueue(i)
            If NETisMessageReady(tmpQueue) Then
                NETpop tmpQueue
            End If

            rIP = tmp_socket(i).textAddress()
            If tmp_connectState(i).connectState <> TmpConnectState_PendingLobbyJoinCheckResult Then
                NETaddSessionBanBadIP rIP
            End If

            NETcloseTempSocket i
            sync_counter.cantjoin = (sync_counter.cantjoin + 1) And &HFFFF&
        End If
continue2:
    Next
End Sub

' (VB) strcmp(tmp_connectState[i].buffer, "list") == 0
Private Function bufferIsList(ByVal i As Long) As Boolean
    bufferIsList = tmp_connectState(i).buffer(0) = 108 And tmp_connectState(i).buffer(1) = 105 And tmp_connectState(i).buffer(2) = 115 _
                   And tmp_connectState(i).buffer(3) = 116 And tmp_connectState(i).buffer(4) = 0
End Function

' (VB) getLocalSharedIdentity(): the game's local shared identity, with its private key (it signs)
Public Function getLocalSharedIdentityKey(k As EcKey) As Long
    Dim key() As Byte, n As Long
    EcKey_clear k
    n = netGame.getLocalSharedIdentity(key)
    If n > 0 Then
        If Not EcKey_fromBytes(k, key, n, EcKey_Private) Then EcKey_fromBytes k, key, n, EcKey_Public
    End If
    getLocalSharedIdentityKey = n
End Function

Private Sub SetupKnownExternalIPv4Info(ByVal externalIp As String, ByVal extPort As Long)
    Dim connType As String, outputExternalIP As String, outputGamePassword As String
    knownExternalIPv4Info.host = externalIp
    knownExternalIPv4Info.port = extPort And &HFFFF&

    If Not activeConnProvider Is Nothing Then connType = ConnectionProviderType_to_string(activeConnProvider.type_()) Else connType = ""
    If externalIp <> "" Then outputExternalIP = externalIp Else outputExternalIP = "unknown"
    If NETGameIsLocked() Then outputGamePassword = NetPlay.gamePassword Else outputGamePassword = ""
    netGame.wzCommandInterfaceOutput va("WZEVENT: readyForJoins: %s %u %s %s" & vbLf, connType, extPort, outputExternalIP, outputGamePassword)
End Sub

' Set the fields for GAMESTRUCT structure to announce it to the lobby server and other players.
Private Sub SetupGameStructInfo(ByVal SessionName As String, ByVal PlayerName As String, playerIdentity As EcKey, ByVal spectatorHost As Boolean, _
        ByVal maxPlayers As Long, ByVal gameType As Long, ByVal blindMode As Long, ByVal alliancesType As Long, ByVal techLevel As Long, _
        ByVal powerLevel As Long, ByVal basesLevel As Long)
    Dim fresh As GameDetails_, key() As Byte, n As Long, currentSpectatorInfo As Long, modNames() As String, modHashes() As String
    gameDetails = fresh

    gameDetails.versionStr = netGame.versionGetVersionString()
    gameDetails.netcodeVer_major = NETCODE_VERSION_MAJOR
    gameDetails.netcodeVer_minor = NETCODE_VERSION_MINOR

    gameDetails.name = SessionName
    gameDetails.host_name = PlayerName
    n = EcKey_toBytes(playerIdentity, EcKey_Public, key)
    gameDetails.host_publicIdentity = base64Encode(key, n)

    gameDetails.map_name = netGame.gameMap()
    gameDetails.map_hash = netGame.gameHashString()

    ' getLoadedMods(): each mod's name, and its hash when not zero (netlobby::ModDetails)
    gameDetails.modsCount = netGame.getLoadedMods(modNames, modHashes)
    If gameDetails.modsCount > 0 Then
        gameDetails.mods_name = modNames
        gameDetails.mods_hash = modHashes
    End If

    If spectatorHost Then gameDetails.players_current = 0 Else gameDetails.players_current = 1
    gameDetails.players_max = maxPlayers

    currentSpectatorInfo = SpectatorInfo_currentNetPlayState()
    gameDetails.spectators_current = SpectatorInfo_spectatorsJoined(currentSpectatorInfo)
    gameDetails.spectators_max = SpectatorInfo_totalSpectatorSlots(currentSpectatorInfo)

    gameDetails.alliancesMode = alliancesType And &HFF&
    gameDetails.blindMode = blindMode

    NETsetLobbyLimitFlags netGame.ingameFlags()

    gameDetails.isPrivate = NETGameIsLocked()
End Sub

Private Sub NETEnableAllowJoining()
    If Not NetPlay.isHost Then ASSERT_FAILURE "false", "lib/netplay/netplay.cpp:4570", "NETEnableAllowJoining", "Host only routine detected for client!": Exit Sub
    If Not (Not allow_joining) Then ASSERT_FAILURE "!allow_joining", "lib/netplay/netplay.cpp:4571", "NETEnableAllowJoining", "allow_joining already true!": Exit Sub

    ' Once this is true, we are able to connect to the lobby server and announce to other players that
    ' this game session is available to join to.
    allow_joining = True

    debug_ 4577, LOG_NET, "NETEnableAllowJoining", "Hosting a server. We are player %d.", selectedPlayer
End Sub

' playerIdentity: the host's EcKey as bytes (its public key: EcKey_toBytes(..., EcKey_Public))
Public Function NEThostGame(ByVal SessionName As String, ByVal PlayerName As String, playerIdentityKey() As Byte, ByVal playerIdentityKeySize As Long, _
                            ByVal spectatorHost As Boolean, ByVal gameType As Long, ByVal blindMode As Long, _
                            ByVal plyrs As Long, _
                            ByVal defaultOpenSpectatorSlots As Long, _
                            ByVal alliancesType As Long, ByVal techLevel As Long, ByVal powerLevel As Long, ByVal basesLevel As Long) As Boolean
    Dim playerIndex As Long, ec As Long, sockErrMsg As String, sockErrMsgB64 As String, i As Long, newHostPlayerIdx As Long, ok As Boolean
    Dim currentSpectatorSlotInfo As Long, desiredDefaultOpenSpectatorSlots As Long, addedSpecSlots As Long, specSlot As Long
    Dim playerIdentity As EcKey, listenResult As IListenSocket
    debug_ 4586, LOG_NET, "NEThostGame", "NEThostGame(%s, %s, %u, %u, %u)", SessionName, PlayerName, _
          gameType, blindMode, plyrs

    netPlayersUpdated = True

    For playerIndex = 0 To MAX_CONNECTED_PLAYERS - 1
        initPlayerNetworkProps playerIndex
    Next
    If Not NetPlay.bComms Then
        If Not (Not spectatorHost) Then ASSERT_FAILURE "!spectatorHost", "lib/netplay/netplay.cpp:4597", "NEThostGame", "spectatorHost flag will be ignored"
        selectedPlayer = 0
        realSelectedPlayer = selectedPlayer
        NetPlay.isHost = True
        NetPlay.hostPlayer = selectedPlayer
        NetPlay.players(0).allocated = True
        NetPlay.players(0).difficulty = AIDifficulty_HUMAN
        NetPlay.playercount = 1
        debug_ 4605, LOG_NET, "NEThostGame", "Hosting but no comms"
        ' Now switch player color of the host to what they normally use for MP games
        If netGame.warGetMPcolour() >= 0 Then
            netGame.changeColour NetPlay.hostPlayer, netGame.warGetMPcolour(), realSelectedPlayer
        End If
        NEThostGame = True
        Exit Function
    End If

    If activeConnProvider Is Nothing Then ASSERT_FAILURE "activeConnProvider != nullptr", "lib/netplay/netplay.cpp:4614", "NEThostGame", "Active connection provider is not set!": NEThostGame = False: Exit Function

    ' Start listening for client connections on `gameserver_port`.
    ' These will initially be assigned to `tmp_socket[i]` until accepted in the game session,
    ' in which case `tmp_socket[i]` will be assigned to `connected_bsocket[i]` and `tmp_socket[i]`
    ' will become nullptr.
    ec = 0
    If server_listen_socket Is Nothing Then
        debug_ 4623, LOG_INFO, "NEThostGame", "Opening listening socket (backend: %s)", ConnectionProviderType_to_string(activeConnProvider.type_())
        Set listenResult = activeConnProvider.openListenSocket(gameserver_port, ec)
        If ec = 0 Then Set server_listen_socket = listenResult Else Set server_listen_socket = Nothing
    End If
    If server_listen_socket Is Nothing Then
        sockErrMsg = net_errorMessage(ec)
        debug_ 4630, LOG_ERROR, "NEThostGame", "Cannot open listening connection: %s", sockErrMsg
        If netGame.wzCommandInterfaceEnabled() Then
            sockErrMsgB64 = base64EncodeString(sockErrMsg)
            netGame.wzCommandInterfaceOutput va("WZEVENT: serverListenSocketFailed: %d %s" & vbLf, net_errorValue(ec), sockErrMsgB64)
        End If
        NEThostGame = False
        Exit Function
    End If
    debug_ 4638, LOG_NET, "NEThostGame", "New server_listen_socket = %p", ObjPtr(server_listen_socket)
    ' Host needs to create a socket set for MAX_PLAYERS
    If server_socket_set Is Nothing Then
        Set server_socket_set = activeConnProvider.newConnectionPollGroup()
    End If
    ' allocate socket storage for all possible players
    For i = 0 To MAX_CONNECTED_PLAYERS - 1
        Set connected_bsocket(i) = Nothing
        NETinitQueue NETnetQueue(i)
    Next

    NetPlay.isHost = True
    NETlogEntry "Hosting game, resetting ban list.", SYNC_FLAG, 0
    NETpermissionsInit
    playerManagementRecord_clear

    If spectatorHost Then
        ' open one spectator slot for the host
        If Not NET_openNewSpectatorSlot_internal(False, specSlot) Then ASSERT_FAILURE "_NET_openNewSpectatorSlot_internal(false).has_value()", "lib/netplay/netplay.cpp:4659", "NEThostGame", "Unable to open spectator slot for host??"
    End If
    ok = NET_CreatePlayer(newHostPlayerIdx, PlayerName, (netGame.getHostLaunch() = HostLaunch_Autohost), True, spectatorHost)
    If ok Then ok = (newHostPlayerIdx < MAX_PLAYERS Or (spectatorHost And newHostPlayerIdx < MAX_CONNECTED_PLAYERS))
    If Not ok Then ASSERT_FAILURE "newHostPlayerIdx.has_value() && (newHostPlayerIdx.value() < MAX_PLAYERS || (spectatorHost && newHostPlayerIdx.value() < MAX_CONNECTED_PLAYERS))", _
        "lib/netplay/netplay.cpp:4663", "NEThostGame", "Failed to create player": NEThostGame = False: Exit Function
    selectedPlayer = newHostPlayerIdx
    realSelectedPlayer = selectedPlayer
    NetPlay.isHost = True
    NetPlay.isHostAlive = True
    NetPlay.hostPlayer = selectedPlayer

    netGame.MultiPlayerJoin selectedPlayer, False, playerIdentityKey, 0

    ' Now switch player color of the host to what they normally use for SP games
    If NetPlay.hostPlayer < MAX_PLAYERS And netGame.warGetMPcolour() >= 0 Then
        netGame.changeColour NetPlay.hostPlayer, netGame.warGetMPcolour(), realSelectedPlayer
    End If

    ' Enable spectator slots as desired
    currentSpectatorSlotInfo = SpectatorInfo_currentNetPlayState()
    desiredDefaultOpenSpectatorSlots = defaultOpenSpectatorSlots
    If desiredDefaultOpenSpectatorSlots > MAX_SPECTATOR_SLOTS Then desiredDefaultOpenSpectatorSlots = MAX_SPECTATOR_SLOTS
    If SpectatorInfo_availableSpectatorSlots(currentSpectatorSlotInfo) < desiredDefaultOpenSpectatorSlots Then
        For addedSpecSlots = SpectatorInfo_availableSpectatorSlots(currentSpectatorSlotInfo) To desiredDefaultOpenSpectatorSlots - 1
            If Not NETopenNewSpectatorSlot() Then
                Exit For
            End If
        Next
    End If

    knownExternalIPv4Info.host = ""     ' knownExternalIPv4Info = KnownExternalConnInfo()
    knownExternalIPv4Info.port = 0
    If lobbyHostingHandler_exists() Then
        lobbyHostingHandler_removeGameListing False
        lobbyHostingHandler_reset
    End If

    If playerIdentityKeySize > 0 Then EcKey_fromBytes playerIdentity, playerIdentityKey, playerIdentityKeySize, EcKey_Public
    SetupGameStructInfo SessionName, PlayerName, playerIdentity, spectatorHost, plyrs, gameType, blindMode, alliancesType, techLevel, powerLevel, basesLevel

    If Not ipv4MappingRequest Is Nothing Then
        ' Do not allow others to join and delay announcing the game session to the lobby server
        ' until we manage to setup (successfully or not) the port mapping rule for the `gameserver_port`.
        PortMappingManager_attach_callback ipv4MappingRequest, PMCB_NEThostGame
    Else
        If Not (Not NetPlay.isPortMappingEnabled) Then ASSERT_FAILURE "!NetPlay.isPortMappingEnabled", "lib/netplay/netplay.cpp:4728", "NEThostGame", "Expected to have an in-flight port mapping request to attach to"

        ' Allow joining with the default gameserver host + port combination and proceed as usual in the hope
        ' that others will still be able to connect to us.
        SetupKnownExternalIPv4Info "", gameserver_port

        NETEnableAllowJoining
    End If

    NEThostGame = True
End Function

' NEThostGame's success callback: [](std::string externalIp, uint16_t extPort) { ... }
Public Sub NEThostGame_successCb(ByVal externalIp As String, ByVal extPort As Long)
    ' Setup gamestruct with the external ip + port combination received from the LibPlum.
    SetupKnownExternalIPv4Info externalIp, extPort

    ' Only allow joining the game once the server has successfully discovered its external IP + port combination.
    '
    ' Once this is true, we are able to connect to the lobby server and announce to other players that
    ' this game session is available to join to.
    NETEnableAllowJoining
End Sub

' NEThostGame's failure callback: [](PortMappingDiscoveryStatus /*status*/) { ... }
Public Sub NEThostGame_failureCb(ByVal status As PortMappingDiscoveryStatus)
    ' Allow joining with the default gameserver host + port combination and proceed as usual in the hope
    ' that others will still be able to connect to us.
    SetupKnownExternalIPv4Info "", gameserver_port

    NETEnableAllowJoining
End Sub

' ////////////////////////////////////////////////////////////////////////
' Stop the dplay interface from accepting more players.
Public Function NEThaltJoining(ByVal matchStarted As Boolean) As Boolean
    debug_ 4744, LOG_NET, "NEThaltJoining", "temporarily locking game to prevent more players"

    allow_joining = False
    ' disconnect from the master server
    If lobbyHostingHandler_exists() Then
        lobbyHostingHandler_removeGameListing matchStarted
        lobbyHostingHandler_reset
    End If
    NEThaltJoining = True
End Function

' "consumes" the sockets and related info: client_joining_socket and client_joining_socket_set are set to Nothing
Public Function NETpromoteJoinAttemptToEstablishedConnectionToHost(ByVal hostPlayer As Long, ByVal index As Long, ByVal playername As String, _
        ByVal joiningQUEUEInfo As Long, client_joining_socket As IClientConnection, client_joining_socket_set As IConnectionPollGroup) As Boolean
    If Not ULT(hostPlayer, MAX_CONNECTED_PLAYERS) Then
        debug_ 4761, LOG_ERROR, "NETpromoteJoinAttemptToEstablishedConnectionToHost", "Bad host player number (%u) received from host!", hostPlayer
        NETpromoteJoinAttemptToEstablishedConnectionToHost = False
        Exit Function
    End If

    If Not ULT(index, MAX_CONNECTED_PLAYERS) Then
        debug_ 4767, LOG_ERROR, "NETpromoteJoinAttemptToEstablishedConnectionToHost", "Bad player number (%u) received from host!", index
        NETpromoteJoinAttemptToEstablishedConnectionToHost = False
        Exit Function
    End If

    NetPlay.hostPlayer = hostPlayer

    ' On success, promote the temporary socket / socketset / queuepair to their permanent (stable) locations
    NETinitQueue NETnetQueue(NetPlay.hostPlayer)
    NETmoveQueue joiningQUEUEInfo, NETnetQueue(NetPlay.hostPlayer)

    ' Promote transient socket to stable client connection.
    ' bsocket = client_joining_socket, client_joining_socket = nullptr now!
    Set bsocket = client_joining_socket
    Set client_joining_socket = Nothing
    Set client_socket_set = client_joining_socket_set
    Set client_joining_socket_set = Nothing

    debug_ 4784, LOG_NET, "NETpromoteJoinAttemptToEstablishedConnectionToHost", "NET_ACCEPTED received. Accepted into the game - I'm player %u using bsocket %p", index, ObjPtr(bsocket)

    selectedPlayer = index
    realSelectedPlayer = selectedPlayer
    NetPlay.isHost = False
    NetPlay.isHostAlive = True

    NetPlay.players(index).allocated = True
    netGame.setPlayerName index, playername
    NetPlay.players(index).heartbeat = True

    ' Store the host's address
    has_lastHostAddress = True
    lastHostAddress = bsocket.textAddress()

    NETpromoteJoinAttemptToEstablishedConnectionToHost = True
End Function

Public Sub NETacceptIncomingConnections()
    Dim i As Long, rIP As String, rIPLogEntry As String, n As Long
    If Not NetPlay.isHost Then ASSERT_FAILURE "false", "lib/netplay/netplay.cpp:5036", "NETacceptIncomingConnections", "Host only routine detected for client!": Exit Sub
    If Not allow_joining Then
        Exit Sub
    End If

    ' First-time initialize `tmp_socket_set` if needed.
    If tmp_socket_set Is Nothing Then
        ' initialize temporary server socket set
        ' FIXME: why is this not done in NETinit()?? - Per
        If activeConnProvider Is Nothing Then ASSERT_FAILURE "activeConnProvider != nullptr", "lib/netplay/netplay.cpp:5047", "NETacceptIncomingConnections", "Active connection provider is not set!": Exit Sub
        Set tmp_socket_set = activeConnProvider.newConnectionPollGroup()
        ' FIXME: I guess initialization of allowjoining is here now... - FlexCoral
        For i = 0 To MAX_TMP_SOCKETS - 1
            TmpSocketInfo_reset i
        Next
        pendingIPs_clear
        ' NOTE: Do *NOT* call tmp_badIPs.clear() here - we want to preserve it until quit
    End If

    ' Find the first empty socket slot
    If Not firstAvailableTempSocketIdx(i) Then
        debug_ 5062, LOG_NET, "NETacceptIncomingConnections", "all temp sockets are currently used up!"
        Exit Sub
    End If

    ' See if there's an incoming connection
    Set tmp_socket(i) = server_listen_socket.accept()
    If tmp_socket(i) Is Nothing Then
        Exit Sub
    End If
    rIP = tmp_socket(i).textAddress()
    If quickRejectConnection(rIP) Then
        debug_ 5076, LOG_NET, "NETacceptIncomingConnections", "freeing temp socket %p (%d)", ObjPtr(tmp_socket(i)), 5078
        NETcloseTempSocket i
        Exit Sub
    End If

    NETinitQueue NETnetTmpQueue(i)
    tmp_socket_set.add tmp_socket(i)

    pendingIPs_increment rIP

    rIPLogEntry = "Incoming connection from:"
    rIPLogEntry = rIPLogEntry & rIP
    NETlogEntry rIPLogEntry, SYNC_FLAG, i

    tmp_connectState(i).ip = rIP
    tmp_connectState(i).connectTime = steadyNowMs()
    tmp_connectState(i).connectState = TmpConnectState_PendingInitialConnect

    If bEnableTCPNoDelay Then
        ' Disable use of Nagle Algorithm for the TCP socket (i.e. enable TCP_NODELAY option in case of TCP connection)
        tmp_socket(i).useNagleAlgorithm False
    End If
End Sub

Public Sub NETadjustConnectedTimeoutForClients()
    Const IN_GAME_CONNECTED_TIMEOUT As Long = 60000     ' std::chrono::milliseconds
    Dim i As Long

    If NetPlay.isHost Then
        For i = 0 To MAX_CONNECTED_PLAYERS - 1
            If Not connected_bsocket(i) Is Nothing Then
                connected_bsocket(i).setConnectedTimeout IN_GAME_CONNECTED_TIMEOUT
            End If
        Next
        Exit Sub
    End If
    If Not bsocket Is Nothing Then
        bsocket.setConnectedTimeout IN_GAME_CONNECTED_TIMEOUT
    End If
End Sub

' optional<std::string> NET_getCurrentHostTextAddress(): has_value() is the result, the address in result
Public Function NET_getCurrentHostTextAddress(result As String) As Boolean
    If bsocket Is Nothing Then
        result = lastHostAddress        ' lastHostAddress.value() (throws without a value in C++)
        NET_getCurrentHostTextAddress = True
        Exit Function
    End If
    result = bsocket.textAddress()
    NET_getCurrentHostTextAddress = True
End Function

'--------------------------------------------------------------------------
' (VB) tmp_pendingIPs

Private Function pendingIPs_find(ByVal ip As String) As Long
    Dim n As Long
    For n = 0 To pendingIPN - 1
        If pendingIPKey(n) = ip Then
            pendingIPs_find = n
            Exit Function
        End If
    Next
    pendingIPs_find = -1
End Function

' tmp_pendingIPs[ip]++
Private Sub pendingIPs_increment(ByVal ip As String)
    Dim n As Long
    n = pendingIPs_find(ip)
    If n < 0 Then
        If pendingIPN = 0 Then
            ReDim pendingIPKey(0 To 15)
            ReDim pendingIPCount(0 To 15)
        ElseIf pendingIPN > UBound(pendingIPKey) Then
            ReDim Preserve pendingIPKey(0 To 2 * pendingIPN - 1)
            ReDim Preserve pendingIPCount(0 To 2 * pendingIPN - 1)
        End If
        n = pendingIPN
        pendingIPKey(n) = ip
        pendingIPCount(n) = 0
        pendingIPN = pendingIPN + 1
    End If
    pendingIPCount(n) = pendingIPCount(n) + 1
End Sub

Private Sub pendingIPs_erase(ByVal n As Long)
    Dim k As Long
    For k = n To pendingIPN - 2
        pendingIPKey(k) = pendingIPKey(k + 1)
        pendingIPCount(k) = pendingIPCount(k + 1)
    Next
    pendingIPN = pendingIPN - 1
End Sub

Private Sub pendingIPs_clear()
    pendingIPN = 0
End Sub

'--------------------------------------------------------------------------
' (VB) tmp_badIPs: an LRU cache of MAX_TMP_BAD_IPS entries that prunes to that many when it holds
' TMP_BAD_IPS_ELASTICITY more (lru11::Cache)

' tryGetPt: the entry's index (it becomes the most recently used), -1 when none
Private Function badIPs_tryGet(ByVal ip As String) As Long
    Dim n As Long
    For n = 0 To badIPN - 1
        If badIPKey(n) = ip Then
            badIPUseCounter = badIPUseCounter + 1
            badIPUse(n) = badIPUseCounter
            badIPs_tryGet = n
            Exit Function
        End If
    Next
    badIPs_tryGet = -1
End Function

Private Sub badIPs_insert(ByVal ip As String, ByVal expiry As Double, ByVal attempts As Double)
    Dim n As Long, oldest As Long, k As Long
    If badIPN = 0 Then
        ReDim badIPKey(0 To 63)
        ReDim badIPExpiry(0 To 63)
        ReDim badIPAttempts(0 To 63)
        ReDim badIPUse(0 To 63)
    ElseIf badIPN > UBound(badIPKey) Then
        ReDim Preserve badIPKey(0 To 2 * badIPN - 1)
        ReDim Preserve badIPExpiry(0 To 2 * badIPN - 1)
        ReDim Preserve badIPAttempts(0 To 2 * badIPN - 1)
        ReDim Preserve badIPUse(0 To 2 * badIPN - 1)
    End If
    n = badIPN
    badIPKey(n) = ip
    badIPExpiry(n) = expiry
    badIPAttempts(n) = attempts
    badIPUseCounter = badIPUseCounter + 1
    badIPUse(n) = badIPUseCounter
    badIPN = badIPN + 1
    If badIPN > MAX_TMP_BAD_IPS + TMP_BAD_IPS_ELASTICITY Then
        ' prune(): remove the least recently used entries until maxSize remain
        Do While badIPN > MAX_TMP_BAD_IPS
            oldest = 0
            For k = 1 To badIPN - 1
                If badIPUse(k) < badIPUse(oldest) Then oldest = k
            Next
            badIPs_remove oldest
        Loop
    End If
End Sub

Private Sub badIPs_remove(ByVal n As Long)
    Dim k As Long
    For k = n To badIPN - 2
        badIPKey(k) = badIPKey(k + 1)
        badIPExpiry(k) = badIPExpiry(k + 1)
        badIPAttempts(k) = badIPAttempts(k + 1)
        badIPUse(k) = badIPUse(k + 1)
    Next
    badIPN = badIPN - 1
End Sub
