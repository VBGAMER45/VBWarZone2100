Attribute VB_Name = "nettypes_c"
' lib/netplay/nettypes.cpp, lib/netplay/nettypes.h: the queues (game queues, net
' queues, temporary queues, the broadcast queue), NETbeginEncode /
' NETbeginDecode / NETend and the serialisation of each type, the session
' keys and the secured messages (SessionKeys: src/shared/sodium's crc_sodium_c).
'
' A NETQUEUE is a Long: queueType * &H10000 + index * &H100 + exclude. Its
' queue pointer and isPair follow from queueType and index (TMP, NET and
' TRANSIENT_JOIN are pairs; GAME and GAME_FORCED use gameQueues(index),
' BROADCAST broadcastQueue), looked up when the queue is used.
' TRANSIENT_JOIN's pair is the joining screen's tmpJoiningQueuePair (a
' NetQueuePair* it owns in C++): it lives here as transientJoinQueuePair, which
' the screen reads and deletes through NETtransientJoin*.
'
' nettypes.h's overload pairs NETx(MessageReader&, T&) / NETx(MessageWriter&, T)
' are NETx_r and NETx_w here, and NETx (the C++ name, as NetApi exposes it)
' takes either As Object and calls the right one; the value is ByRef (a
' reader sets it).
Option Explicit

' There is a game queue representing each player. The game queues are synchronised among all players, so that all players process the same game queue
' messages at the same game time. The game queues should be used, even in single-player. Players should write to their own queue, not to other player's
' queues, and should read messages from all queues including their own.
' (Note: Extra +1 for the added replay spectator)
Private gameQueues(0 To MAX_GAMEQUEUE_SLOTS - 1) As NetQueue

' There is a bidirectional net queue for communicating with each client or host. Each queue corresponds either to a real socket, or a virtual socket
' which routes via the host.
Private netQueues(0 To MAX_CONNECTED_PLAYERS - 1) As NetQueuePair

' These queues are for clients which just connected, but haven't yet been assigned a player number.
Private tmpQueues(0 To MAX_TMP_SOCKETS - 1) As NetQueuePair

' Sending a message to the broadcast queue is equivalent to sending the message to the net queues of all other players.
Private broadcastQueue As NetQueue

' joiningscreen.cpp's NetQueuePair *tmpJoiningQueuePair (QUEUE_TRANSIENT_JOIN's pair)
Private transientJoinQueuePair As NetQueuePair

' std::array<std::unique_ptr<SessionKeys>, MAX_CONNECTED_PLAYERS> netSessionKeys: hasSessionKeys True where the
' unique_ptr is set
Private netSessionKeys(0 To MAX_CONNECTED_PLAYERS - 1) As SessionKeys
Private hasSessionKeys(0 To MAX_CONNECTED_PLAYERS - 1) As Boolean

' static std::vector<uint8_t> tmpMessageRawDataBuffer (NETend's)
Private tmpMessageRawDataBuffer() As Byte

Private bIsReplay As Boolean

Private numInvalidMessageReads As Long          ' size_t

'--------------------------------------------------------------------------
' (VB) NETQUEUE's fields

Public Function NETQUEUE_make(ByVal queueType_ As Long, ByVal index As Long, ByVal exclude As Long) As Long
    NETQUEUE_make = queueType_ * &H10000 + (index And &HFF&) * &H100 + (exclude And &HFF&)
End Function

' queue.index
Public Function NETQUEUE_index(ByVal queue As Long) As Long
    NETQUEUE_index = (queue And &HFF00&) \ &H100&
End Function

' queue.queueType
Public Function NETQUEUE_queueType(ByVal queue As Long) As Long
    NETQUEUE_queueType = (queue And &H7FFF0000) \ &H10000
End Function

' queue.exclude
Public Function NETQUEUE_exclude(ByVal queue As Long) As Long
    NETQUEUE_exclude = queue And &HFF&
End Function

' queue.isPair
Public Function NETQUEUE_isPair(ByVal queue As Long) As Boolean
    Select Case NETQUEUE_queueType(queue)
    Case QUEUE_TMP, QUEUE_NET, QUEUE_TRANSIENT_JOIN
        NETQUEUE_isPair = True
    End Select
End Function

' (VB) queue.queue == nullptr: for a pair, queue.queue points at its slot
' (never nullptr); otherwise it is the NetQueue itself
Public Function NETQUEUE_queueIsNull(ByVal queue As Long) As Boolean
    If NETQUEUE_isPair(queue) Then
        NETQUEUE_queueIsNull = False
    Else
        NETQUEUE_queueIsNull = queueOf(queue) Is Nothing
    End If
End Function

' (VB) the number of bytes in a dynamic array (0 when not allocated)
Private Function byteArraySize(b() As Byte) As Long
    On Error GoTo none
    byteArraySize = UBound(b) - LBound(b) + 1
    Exit Function
none:
    byteArraySize = 0
End Function

'--------------------------------------------------------------------------
' Queue selection functions

' Gets the &NetQueuePair::send or NetQueue *, corresponding to queue.
' (A pair that is not there gives Nothing, as &(nullptr)->send is nullptr.)
Private Function sendQueue(ByVal queue As Long) As NetQueue
    Dim p As NetQueuePair
    If NETQUEUE_isPair(queue) Then
        Set p = pairQueue(queue)
        If Not p Is Nothing Then Set sendQueue = p.send
    Else
        Set sendQueue = queueOf(queue)
    End If
End Function

' Gets the &NetQueuePair::receive or NetQueue *, corresponding to queue.
Private Function receiveQueue(ByVal queue As Long) As NetQueue
    If NETQUEUE_isPair(queue) Then
        Set receiveQueue = pairQueue(queue).receive
    Else
        Set receiveQueue = queueOf(queue)
    End If
End Function

' (VB) queue.queue of a queue that is no pair
Private Function queueOf(ByVal queue As Long) As NetQueue
    Dim index As Long
    index = NETQUEUE_index(queue)
    Select Case NETQUEUE_queueType(queue)
    Case QUEUE_GAME, QUEUE_GAME_FORCED
        If index < MAX_GAMEQUEUE_SLOTS Then Set queueOf = gameQueues(index)
    Case QUEUE_BROADCAST
        Set queueOf = broadcastQueue
    End Select
End Function

' Gets the &NetQueuePair, corresponding to queue.
Private Function pairQueue(ByVal queue As Long) As NetQueuePair
    Dim index As Long
    If Not NETQUEUE_isPair(queue) Then ASSERT_FAILURE "queue.isPair", "lib/netplay/nettypes.cpp:88", "pairQueue", "Huh?"
    index = NETQUEUE_index(queue)
    Select Case NETQUEUE_queueType(queue)
    Case QUEUE_NET
        If index < MAX_CONNECTED_PLAYERS Then Set pairQueue = netQueues(index)
    Case QUEUE_TMP
        If index < MAX_TMP_SOCKETS Then Set pairQueue = tmpQueues(index)
    Case QUEUE_TRANSIENT_JOIN
        Set pairQueue = transientJoinQueuePair
    End Select
End Function

' (VB) pairQueue(queue) = p
Private Sub setPairQueue(ByVal queue As Long, ByVal p As NetQueuePair)
    Dim index As Long
    If Not NETQUEUE_isPair(queue) Then ASSERT_FAILURE "queue.isPair", "lib/netplay/nettypes.cpp:88", "pairQueue", "Huh?"
    index = NETQUEUE_index(queue)
    Select Case NETQUEUE_queueType(queue)
    Case QUEUE_NET
        If index < MAX_CONNECTED_PLAYERS Then Set netQueues(index) = p
    Case QUEUE_TMP
        If index < MAX_TMP_SOCKETS Then Set tmpQueues(index) = p
    Case QUEUE_TRANSIENT_JOIN
        Set transientJoinQueuePair = p
    End Select
End Sub

' (VB) the joining screen's tmpJoiningQueuePair: &tmpJoiningQueuePair->send's numMessagesForNet, getMessageForNet,
' popMessageForNet; delete tmpJoiningQueuePair (also its "= nullptr" once NETpromoteJoinAttemptToEstablishedConnectionToHost
' moved it to the host's net queue: NETmoveQueue left Nothing here then)
Public Function NETtransientJoinSendQueue_numMessagesForNet() As Long
    NETtransientJoinSendQueue_numMessagesForNet = transientJoinQueuePair.send.numMessagesForNet()
End Function

Public Function NETtransientJoinSendQueue_getMessageForNet() As NetMessage
    Set NETtransientJoinSendQueue_getMessageForNet = transientJoinQueuePair.send.getMessageForNet()
End Function

Public Sub NETtransientJoinSendQueue_popMessageForNet()
    transientJoinQueuePair.send.popMessageForNet
End Sub

Public Sub NETtransientJoinQueuePair_delete()
    Set transientJoinQueuePair = Nothing
End Sub

Public Function NETtransientJoinQueuePair_exists() As Boolean
    NETtransientJoinQueuePair_exists = Not transientJoinQueuePair Is Nothing
End Function

' One of the temp queues from before a client has joined the game. (See comments on tmpQueues.)
Public Function NETnetTmpQueue(ByVal tmpPlayer As Long) As Long
    If Not ULT(tmpPlayer, MAX_TMP_SOCKETS) Then ASSERT_FAILURE "tmpPlayer < 16", "lib/netplay/nettypes.cpp:95", "NETnetTmpQueue", "Huh?"
    NETnetTmpQueue = NETQUEUE_make(QUEUE_TMP, tmpPlayer, NET_NO_EXCLUDE)
End Function

' The queue pair used for sending and receiving data directly from another client. (See comments on netQueues.)
Public Function NETnetQueue(ByVal player As Long, Optional ByVal excludePlayer As Long = 255) As Long
    If player = NET_ALL_PLAYERS Then
        NETnetQueue = NETbroadcastQueue(excludePlayer)
        Exit Function
    End If

    If Not ULT(player, MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "player < ((11 + 2) + 10)", "lib/netplay/nettypes.cpp:114", "NETnetQueue", "Huh?"
    NETnetQueue = NETQUEUE_make(QUEUE_NET, player, excludePlayer)
End Function

' The game action queue. (See comments on gameQueues.)
Public Function NETgameQueue(ByVal player As Long) As Long
    If Not ULT(player, MAX_GAMEQUEUE_SLOTS) Then
        ' found one
        debug_ 130, LOG_ERROR, "NETgameQueue", "Found the call"
    End If
    If Not ULT(player, MAX_GAMEQUEUE_SLOTS) Then ASSERT_FAILURE "player < (((11 + 2) + 10) + 1)", "lib/netplay/nettypes.cpp:132", "NETgameQueue", "Huh?"
    NETgameQueue = NETQUEUE_make(QUEUE_GAME, player, NET_NO_EXCLUDE)
End Function

Public Function NETgameIsBehindPlayersByAtLeast(Optional ByVal numGameTimeUpdates As Long = 2) As Boolean
    Dim player As Long, pPlayerGameQueue As NetQueue
    ' if we should be waited on, then there's no reason we should be behind other players
    If netGame.gtimeShouldWaitForPlayer(realSelectedPlayer) Then
        NETgameIsBehindPlayersByAtLeast = False
        Exit Function
    End If

    For player = 0 To MAX_CONNECTED_PLAYERS - 1
        Set pPlayerGameQueue = gameQueues(player)
        If pPlayerGameQueue Is Nothing Then
            GoTo continue_
        End If

        If netGame.gtimeShouldWaitForPlayer(player) Then
            If ULT(pPlayerGameQueue.numPendingGameTimeUpdateMessages(), numGameTimeUpdates) Then
                NETgameIsBehindPlayersByAtLeast = False  ' Do not have enough pending game time updates for this player
                Exit Function
            End If
        End If
continue_:
    Next

    NETgameIsBehindPlayersByAtLeast = True      ' Have enough pending game time updates from all players that should be waited on
End Function

' Only used by the host, to force-feed a GAME_PLAYER_LEFT message into someone's game queue.
Public Function NETgameQueueForced(ByVal player As Long) As Long
    If Not ULT(player, MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "player < ((11 + 2) + 10)", "lib/netplay/nettypes.cpp:172", "NETgameQueueForced", "Huh?"
    NETgameQueueForced = NETQUEUE_make(QUEUE_GAME_FORCED, player, NET_NO_EXCLUDE)
End Function

' The queue for sending data directly to the netQueues of all clients, not just a specific one. (See comments on broadcastQueue.)
Public Function NETbroadcastQueue(Optional ByVal excludePlayer As Long = 255) As Long
    NETbroadcastQueue = NETQUEUE_make(QUEUE_BROADCAST, NET_ALL_PLAYERS, excludePlayer)
End Function

' Dump raw data from sockets and raw data sent via host here: data(0 .. dataLen - 1).
Public Sub NETinsertRawData(ByVal queue As Long, data() As Byte, ByVal dataLen As Long)
    receiveQueue(queue).writeRawData data, dataLen
End Sub

' Dump whole NetMessages into the queue.
Public Sub NETinsertMessageFromNet(ByVal queue As Long, ByVal newMessage As NetMessage)
    receiveQueue(queue).pushMessage newMessage
End Sub

' Returns true if there is a complete message ready to deserialise in this queue.
Public Function NETisMessageReady(ByVal queue As Long) As Boolean
    NETisMessageReady = receiveQueue(queue).haveMessage()
End Function

Public Function NETincompleteMessageDataBuffered(ByVal queue As Long) As Long
    NETincompleteMessageDataBuffered = receiveQueue(queue).currentIncompleteDataBuffered()
End Function

' Returns the current message in the queue which is ready to be deserialised. Do not delete the message.
Public Function NETgetMessage(ByVal queue As Long) As NetMessage
    Set NETgetMessage = receiveQueue(queue).getMessage()
End Function

'--------------------------------------------------------------------------
' Begin & End functions

' Allocates the queue. Deletes the old queue, if there was one. Avoids a crash on NULL pointer deference when trying to use the queue.
Public Sub NETinitQueue(ByVal queue As Long)
    If NETQUEUE_queueType(queue) = QUEUE_BROADCAST Then
        Set broadcastQueue = New NetQueue       ' (the old one deleted)
        broadcastQueue.setWillNeverGetMessages
        Exit Sub
    ElseIf NETQUEUE_queueType(queue) = QUEUE_GAME Then
        Set gameQueues(NETQUEUE_index(queue)) = New NetQueue
        Exit Sub
    Else
        pairQueue queue                         ' delete pairQueue(queue);
        setPairQueue queue, New NetQueuePair
    End If
End Sub

' Delete queues for cleanup
Public Sub NETdeleteQueue()
    Dim i As Long
    For i = 0 To MAX_CONNECTED_PLAYERS - 1
        pairQueue NETnetQueue(i)                ' delete pairQueue(NETnetQueue(i));
        setPairQueue NETnetQueue(i), Nothing
        Set gameQueues(i) = Nothing
    Next

    ' extra replay spectator gamequeue
    Set gameQueues(MAX_CONNECTED_PLAYERS) = Nothing

    Set broadcastQueue = Nothing

    bIsReplay = False

    If numInvalidMessageReads <> 0 Then
        debug_ 266, LOG_NET, "NETdeleteQueue", "Experienced %zu invalid message reads", numInvalidMessageReads
    End If
    numInvalidMessageReads = 0
End Sub

' Used to mark that a game queue should not be sent over the network (for example, if it is being sent to us, instead).
Public Sub NETsetNoSendOverNetwork(ByVal queue As Long)
    sendQueue(queue).setWillNeverGetMessagesForNet  ' Will not be sending over net.
End Sub

' Used for moving the tmpQueue to a netQueue, once a newly-connected client is assigned a player number.
Public Sub NETmoveQueue(ByVal src As Long, ByVal dst As Long)
    Dim p As NetQueuePair
    If Not NETQUEUE_isPair(src) Then ASSERT_FAILURE "src.isPair", "lib/netplay/nettypes.cpp:278", "NETmoveQueue", "Huh?"
    If Not NETQUEUE_isPair(dst) Then ASSERT_FAILURE "dst.isPair", "lib/netplay/nettypes.cpp:279", "NETmoveQueue", "Huh?"
    pairQueue dst                               ' delete pairQueue(dst);
    setPairQueue dst, Nothing
    ' std::swap(pairQueue(src), pairQueue(dst));
    Set p = pairQueue(src)
    setPairQueue src, pairQueue(dst)
    setPairQueue dst, p
End Sub

' Used for swapping two netQueues, when swapping a player index (i.e. player <-> spectator-only slot). Dangerous and rare.
Public Sub NETswapQueues(ByVal src As Long, ByVal dst As Long)
    Dim p As NetQueuePair
    If Not NETQUEUE_isPair(src) Then ASSERT_FAILURE "src.isPair", "lib/netplay/nettypes.cpp:287", "NETswapQueues", "Huh?"
    If Not NETQUEUE_isPair(dst) Then ASSERT_FAILURE "dst.isPair", "lib/netplay/nettypes.cpp:288", "NETswapQueues", "Huh?"
    ' std::swap(pairQueue(src), pairQueue(dst));
    Set p = pairQueue(src)
    setPairQueue src, pairQueue(dst)
    setPairQueue dst, p
End Sub

Public Function NETbeginEncode(ByVal queue As Long, ByVal typ As Long) As MessageWriter
    Set NETbeginEncode = MessageWriter_new(queue, typ)
End Function

Public Function NETbeginDecode(ByVal queue As Long, ByVal typ As Long) As MessageReader
    Dim res As MessageReader
    Set res = MessageReader_new(receiveQueue(queue).getMessage())
    ' assert(type == res.msgType): a C assert, compiled out in a release build (NDEBUG)
    Set NETbeginDecode = res
End Function

' NETsetSessionKeys(uint8_t player, SessionKeys&& keys): the keys as their bytes (receiveKey, sendKey: whole arrays)
Public Sub NETsetSessionKeys(ByVal player As Long, receiveKey() As Byte, sendKey() As Byte)
    Dim keys As SessionKeys
    If Not ULT(player, MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "player < netSessionKeys.size()", "lib/netplay/nettypes.cpp:306", "NETsetSessionKeys", "Invalid player: %u", player: Exit Sub
    keys.receiveKey = receiveKey
    keys.receiveKeySize = byteArraySize(receiveKey)
    keys.sendKey = sendKey
    keys.sendKeySize = byteArraySize(sendKey)
    netSessionKeys(player) = keys               ' std::make_unique<SessionKeys>(std::move(keys))
    hasSessionKeys(player) = True
End Sub

' NETclearSessionKeys()
Public Sub NETclearSessionKeys()
    Dim i As Long, blank As SessionKeys
    For i = 0 To MAX_CONNECTED_PLAYERS - 1
        netSessionKeys(i) = blank               ' key.reset()
        hasSessionKeys(i) = False
    Next
End Sub

' NETclearSessionKeys(uint8_t player)
Public Sub NETclearSessionKeys_player(ByVal player As Long)
    Dim blank As SessionKeys
    If Not ULT(player, MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "player < netSessionKeys.size()", "lib/netplay/nettypes.cpp:320", "NETclearSessionKeys", "Invalid player: %u", player: Exit Sub
    netSessionKeys(player) = blank
    hasSessionKeys(player) = False
End Sub

Public Function NETisExpectedSecuredMessageType(ByVal typ As Long) As Boolean
    Select Case typ
    Case NET_TEAM_STRATEGY
        NETisExpectedSecuredMessageType = True
    Case NET_QUICK_CHAT_MSG
        NETisExpectedSecuredMessageType = True
    Case Else
        NETisExpectedSecuredMessageType = False
    End Select
End Function

' For encoding a secured net message, for a *specific player*
' Returns Nothing (nullopt) on failure
' Notes:
'   - Only for NET_* messages
Public Function NETbeginEncodeSecured(ByVal queue As Long, ByVal typ As Long) As MessageWriter
    Dim w As MessageWriter, index As Long
    index = NETQUEUE_index(queue)
    If Not (typ < NET_MAX_TYPE) Then ASSERT_FAILURE "type < NET_MAX_TYPE", "lib/netplay/nettypes.cpp:343", "NETbeginEncodeSecured", "Message type %u is >= NET_MAX_TYPE", typ: Exit Function
    If Not (index <> realSelectedPlayer) Then ASSERT_FAILURE "queue.index != realSelectedPlayer", "lib/netplay/nettypes.cpp:344", "NETbeginEncodeSecured", "Secured messages are for other players, not ourselves.": Exit Function
    If Not (index < MAX_PLAYERS Or index = NetPlay.hostPlayer) Then ASSERT_FAILURE "queue.index < MAX_PLAYERS || queue.index == NetPlay.hostPlayer", "lib/netplay/nettypes.cpp:345", "NETbeginEncodeSecured", "Invalid recipient (queue.index == %u)", index: Exit Function
    If Not hasSessionKeys(index) Then ASSERT_FAILURE "netSessionKeys[queue.index] != nullptr", "lib/netplay/nettypes.cpp:346", "NETbeginEncodeSecured", "Lacking session key for recipient: %u", index: Exit Function
    If Not NETisExpectedSecuredMessageType(typ) Then ASSERT_FAILURE "NETisExpectedSecuredMessageType(type)", "lib/netplay/nettypes.cpp:347", "NETbeginEncodeSecured", "Message type is not expected to be secured, and will be ignored on receipt"

    Set w = NETbeginEncode(queue, typ)
    w.bSecretMessageWrap = True

    Set NETbeginEncodeSecured = w
End Function

' For decoding what is expected to have been a secured message
' Returns Nothing (nullopt) on failure
' Notes:
'   - *DO NOT CALL NETend() if this returns Nothing!!*
'   - Only for NET_* messages
Public Function NETbeginDecodeSecured(ByVal queue As Long, ByVal typ As Long) As MessageReader
    Dim index As Long
    index = NETQUEUE_index(queue)
    If Not (typ < NET_MAX_TYPE) Then ASSERT_FAILURE "type < NET_MAX_TYPE", "lib/netplay/nettypes.cpp:362", "NETbeginDecodeSecured", "Message type %u is >= NET_MAX_TYPE", typ: Exit Function
    If Not (index <> realSelectedPlayer) Then ASSERT_FAILURE "queue.index != realSelectedPlayer", "lib/netplay/nettypes.cpp:363", "NETbeginDecodeSecured", "Secured messages are for other players, not ourselves.": Exit Function
    If Not (index < MAX_PLAYERS Or index = NetPlay.hostPlayer) Then ASSERT_FAILURE "queue.index < MAX_PLAYERS || queue.index == NetPlay.hostPlayer", "lib/netplay/nettypes.cpp:364", "NETbeginDecodeSecured", "Invalid sender (queue.index == %u)", index: Exit Function
    If Not receiveQueue(queue).currentMessageWasDecrypted() Then ASSERT_FAILURE "receiveQueue(queue)->currentMessageWasDecrypted()", "lib/netplay/nettypes.cpp:365", "NETbeginDecodeSecured", "Message was not sent secured (type: %s)", messageTypeToString(typ): Exit Function

    Set NETbeginDecodeSecured = NETbeginDecode(queue, typ)
End Function

' Decrypts a secured net message in a queue *and replaces it with the decrypted message*
' If message is successfully decrypted:
'   - Returns true, updates the current message in queue to be the decrypted message, updates `type`
Public Function NETdecryptSecuredNetMessage(ByVal queue As Long, typ As Long) As Boolean
    Dim index As Long, pReceiveQueue As NetQueue, encryptedMessage As NetMessage
    Dim payload() As Byte, decryptedMessageRawData() As Byte, decryptedSize As Long, decryptedMessage As NetMessage
    index = NETQUEUE_index(queue)
    If Not (index < MAX_PLAYERS Or index = NetPlay.hostPlayer) Then ASSERT_FAILURE "queue.index < MAX_PLAYERS || queue.index == NetPlay.hostPlayer", "lib/netplay/nettypes.cpp:375", "NETdecryptSecuredNetMessage", "Invalid sender (queue.index == %u", index: NETdecryptSecuredNetMessage = False: Exit Function
    If Not hasSessionKeys(index) Then ASSERT_FAILURE "netSessionKeys[queue.index] != nullptr", "lib/netplay/nettypes.cpp:376", "NETdecryptSecuredNetMessage", "Lacking session key for player: %u", index: NETdecryptSecuredNetMessage = False: Exit Function
    If Not (typ = NET_SECURED_NET_MESSAGE) Then ASSERT_FAILURE "type == NET_SECURED_NET_MESSAGE", "lib/netplay/nettypes.cpp:377", "NETdecryptSecuredNetMessage", "Not a secured message?": NETdecryptSecuredNetMessage = False: Exit Function

    Set pReceiveQueue = receiveQueue(queue)

    ' Get & decrypt message raw data
    Set encryptedMessage = pReceiveQueue.getMessage()
    If Not (encryptedMessage.type_() = NET_SECURED_NET_MESSAGE) Then ASSERT_FAILURE "encryptedMessage.type() == NET_SECURED_NET_MESSAGE", "lib/netplay/nettypes.cpp:383", "NETdecryptSecuredNetMessage", "Not a secured message?": NETdecryptSecuredNetMessage = False: Exit Function
    payload = encryptedMessage.payload()
    If Not SessionKeys_decryptMessageFromOther(netSessionKeys(index), payload, 0, encryptedMessage.payloadSize(), decryptedMessageRawData, decryptedSize) Then
        debug_ 387, LOG_INFO, "NETdecryptSecuredNetMessage", "Invalid encrypted message from player: %u", index
        NETdecryptSecuredNetMessage = False
        Exit Function
    End If

    Set decryptedMessage = NetMessage_tryFromRawData(decryptedMessageRawData, decryptedSize)
    If decryptedMessage Is Nothing Then
        debug_ 394, LOG_INFO, "NETdecryptSecuredNetMessage", "Failed to parse decrypted data from player: %u", index
        NETdecryptSecuredNetMessage = False
        Exit Function
    End If

    If Not (decryptedMessage.type_() > NET_MIN_TYPE And decryptedMessage.type_() < NET_MAX_TYPE) Then
        debug_ 400, LOG_NET, "NETdecryptSecuredNetMessage", "Not a secured NET_* message? (type: %s) - ignoring", messageTypeToString(decryptedMessage.type_())
        NETdecryptSecuredNetMessage = False
        Exit Function
    End If

    If Not NETisExpectedSecuredMessageType(decryptedMessage.type_()) Then
        ' Ignore message types that aren't expected to be secured
        debug_ 407, LOG_NET, "NETdecryptSecuredNetMessage", "Not a message type that's expected to be secured: (type: %s) - ignoring", messageTypeToString(decryptedMessage.type_())
        NETdecryptSecuredNetMessage = False
        Exit Function
    End If

    NETlogPacket NET_SECURED_NET_MESSAGE, encryptedMessage.rawDataSize(), True

    typ = decryptedMessage.type_()              ' must update type!
    pReceiveQueue.replaceCurrentWithDecrypted decryptedMessage
    NETdecryptSecuredNetMessage = True
End Function

' NETend(MessageReader& r)
Public Function NETend_r(ByVal r As MessageReader) As Boolean
    Dim result As Boolean
    result = r.valid()
    If Not result Then
        numInvalidMessageReads = numInvalidMessageReads + 1
    End If
    NETend_r = result
End Function

' NETend(MessageWriter& w)
Public Function NETend_w(ByVal w As MessageWriter) As Boolean
    Dim shouldWrapSecretMessage As Boolean, queue As NetQueue, msg As NetMessage, player As Long, allPlayers As Long
    Dim shareQueueWriter As MessageWriter, builtShareGameQueueMessage As NetMessage
    Dim sendToPlayerWriter As MessageWriter, sendToHostSelfWriter As MessageWriter
    Dim msgType As Long, msgRawDataSize As Long, qt As Long
    Dim tmpMessageRawDataSize As Long, encryptedData() As Byte, encryptedDataSize As Long, encryptedNetMessage As NetMessageBuilder
    shouldWrapSecretMessage = w.bSecretMessageWrap
    w.bSecretMessageWrap = False                ' reset, always

    If bIsReplay And NETQUEUE_index(w.queueInfo) <> realSelectedPlayer Then
        ' don't bother adding to the send queue if we're playing a replay
        NETend_w = True
        Exit Function
    End If

    ' Push the message onto the list.
    Set queue = sendQueue(w.queueInfo)
    If queue Is Nothing Then
        debug_ 445, LOG_WARNING, "NETend", "Sending %s to null queue, type %d.", messageTypeToString(w.msgBuilder.type_()), NETQUEUE_queueType(w.queueInfo)
        NETend_w = True
        Exit Function
    End If

    Set msg = w.msgBuilder.build()
    If shouldWrapSecretMessage Then
        ' Need to actually encrypt and wrap the current net message in a NET_SECURED_NET_MESSAGE message
        If Not (msg.type_() < NET_MAX_TYPE) Then ASSERT_FAILURE "msg.type() < NET_MAX_TYPE", "lib/netplay/nettypes.cpp:453", "NETend", "Message type %u is >= NET_MAX_TYPE", msg.type_()
        If Not (NETQUEUE_index(w.queueInfo) < MAX_PLAYERS Or NETQUEUE_index(w.queueInfo) = NetPlay.hostPlayer) Then ASSERT_FAILURE "w.queueInfo.index < 11 || w.queueInfo.index == NetPlay.hostPlayer", "lib/netplay/nettypes.cpp:454", "NETend", "Invalid recipient (queue.index == %u)", NETQUEUE_index(w.queueInfo)
        If Not hasSessionKeys(NETQUEUE_index(w.queueInfo)) Then ASSERT_FAILURE "netSessionKeys[w.queueInfo.index] != nullptr", "lib/netplay/nettypes.cpp:455", "NETend", "Lacking session key for recipient: %u", NETQUEUE_index(w.queueInfo)

        ' Decoded in NETprocessSystemMessage in netplay.cpp.
        ' Encrypt the serialized message (including type and size)
        tmpMessageRawDataSize = 0               ' tmpMessageRawDataBuffer.clear()
        msg.rawDataAppendToVector tmpMessageRawDataBuffer, tmpMessageRawDataSize

        encryptedDataSize = SessionKeys_encryptMessageForOther(netSessionKeys(NETQUEUE_index(w.queueInfo)), tmpMessageRawDataBuffer, 0, tmpMessageRawDataSize, encryptedData)
        Set encryptedNetMessage = NetMessageBuilder_new(NET_SECURED_NET_MESSAGE, encryptedDataSize)
        If encryptedDataSize > 0 Then encryptedNetMessage.append_bytes encryptedData, 0, encryptedDataSize
        Set msg = encryptedNetMessage.build()
    End If

    qt = NETQUEUE_queueType(w.queueInfo)
    If qt = QUEUE_GAME_FORCED Then              ' If true, we must be the host, inserting a GAME_PLAYER_LEFT into the other player's game queue. Since they left, they're not around to complain about us messing with their queue, which would normally cause a desynch.
        ' Almost duplicate code from NETflushGameQueues() in here.
        ' Message must be sent as message in a NET_SHARE_GAME_QUEUE in a NET_SEND_TO_PLAYER.
        ' If not sent inside a NET_SEND_TO_PLAYER, and the message arrives at the same time as a real message from that player, then the messages may be processed in an unexpected order.
        ' See NETrecvNet, in the for (current = 0; current < MAX_CONNECTED_PLAYERS; ++current) loop.
        ' Assume dropped client sends a NET_SENT_TO_PLAYERS(broadcast, NET_SHARE_GAME_QUEUE(GAME_REALMESSAGE)), and then drops.
        ' The host then sends the NET_SEND_TO_PLAYER(broadcast, NET_SHARE_GAME_QUEUE(GAME_REALMESSAGE)) to everyone. If the host then spoofs a NET_SHARE_GAME_QUEUE(GAME_PLAYER_LEFT) without
        ' wrapping it in a NET_SEND_TO_PLAYER from that player, then the client may sometimes unwrap the real NET_SEND_TO_PLAYER message, then process the host's spoofed NET_SHARE_GAME_QUEUE
        ' message, and then after that process the previously unwrapped NET_SHARE_GAME_QUEUE(GAME_REALMESSAGE) message, such that the GAME_PLAYER_LEFT appears on the queue before the
        ' GAME_REALMESSAGE.

        ' Decoded in NETprocessSystemMessage in netplay.cpp.
        player = NETQUEUE_index(w.queueInfo)
        Set shareQueueWriter = NETbeginEncode(NETbroadcastQueue(), NET_SHARE_GAME_QUEUE)
        NETuint8_t_w shareQueueWriter, player
        NETuint32_t_w shareQueueWriter, 1
        NETnetMessage_w shareQueueWriter, msg

        Set builtShareGameQueueMessage = shareQueueWriter.msgBuilder.build()

        allPlayers = NET_ALL_PLAYERS
        Set sendToPlayerWriter = NETbeginEncode(NETbroadcastQueue(), NET_SEND_TO_PLAYER)
        NETuint8_t_w sendToPlayerWriter, player
        NETuint8_t_w sendToPlayerWriter, allPlayers
        NETnetMessage_w sendToPlayerWriter, builtShareGameQueueMessage
        NETend_w sendToPlayerWriter             ' This time we actually send it.

        ' Also insert the NET_SEND_TO_PLAYER into the ** host queue **
        ' - The broadcast above doesn't do this, since broadcasts don't get sent to self
        ' - Insert into the host queue so that it gets processed just like on the clients
        Set sendToHostSelfWriter = NETbeginEncode(NETnetQueue(NetPlay.hostPlayer), NET_SEND_TO_PLAYER)
        NETuint8_t_w sendToHostSelfWriter, player
        NETuint8_t_w sendToHostSelfWriter, allPlayers
        NETnetMessage_w sendToHostSelfWriter, builtShareGameQueueMessage
        NETinsertMessageFromNet NETnetQueue(NetPlay.hostPlayer), sendToHostSelfWriter.msgBuilder.build()

        NETend_w = True
        Exit Function
    End If

    msgType = msg.type_()
    msgRawDataSize = msg.rawDataSize()

    NETlogPacket msgType, msgRawDataSize, False
    queue.pushMessage msg

    If qt = QUEUE_GAME Or qt = QUEUE_GAME_FORCED Then
        If Not (msgType > GAME_MIN_TYPE And msgType < GAME_MAX_TYPE) Then ASSERT_FAILURE "msgType > GAME_MIN_TYPE && msgType < GAME_MAX_TYPE", "lib/netplay/nettypes.cpp:516", "NETend", "Inserting %s into game queue.", messageTypeToString(msgType)
    Else
        If Not (msgType > NET_MIN_TYPE And msgType < NET_MAX_TYPE) Then ASSERT_FAILURE "msgType > NET_MIN_TYPE && msgType < NET_MAX_TYPE", "lib/netplay/nettypes.cpp:520", "NETend", "Inserting %s into net queue.", messageTypeToString(msgType)
    End If

    If qt = QUEUE_NET Or qt = QUEUE_BROADCAST Or qt = QUEUE_TMP Then
        NETsend w.queueInfo, queue.getMessageForNet()
        queue.popMessageForNet
        If Not (queue.numMessagesForNet() = 0) Then ASSERT_FAILURE "queue->numMessagesForNet() == 0", "lib/netplay/nettypes.cpp:527", "NETend", _
            "Queue not empty (%u messages remaining). (message = type: %u, size: %zu), (queue = index: %u; queueType: %u; exclude: %u; isPair: %d)", _
            queue.numMessagesForNet(), msgType, msgRawDataSize, NETQUEUE_index(w.queueInfo), qt, NETQUEUE_exclude(w.queueInfo), Abs(CLng(NETQUEUE_isPair(w.queueInfo)))
    End If

    ' Process any delayed actions from the NETsend call
    NETsendProcessDelayedActions

    NETend_w = True                             ' Serialising never fails.
End Function

' NETend(r) or NETend(w)
Public Function NETend(ByVal c As Object) As Boolean
    If TypeOf c Is MessageReader Then NETend = NETend_r(c) Else NETend = NETend_w(c)
End Function

Public Sub NETflushGameQueues()
    Dim player As Long, queue As NetQueue, num As Long, n As Long, w As MessageWriter
    For player = 0 To MAX_GAMEQUEUE_SLOTS - 1
        Set queue = gameQueues(player)

        If queue Is Nothing Then
            GoTo continue_                      ' Can't send for this player.
        End If

        num = queue.numMessagesForNet()

        If num <= 0 Then
            GoTo continue_                      ' Nothing to send for this player.
        End If

        If bIsReplay Then ASSERT_FAILURE "!bIsReplay", "lib/netplay/nettypes.cpp:554", "NETflushGameQueues", "Where are we sending this if it's a replay?"

        ' Decoded in NETprocessSystemMessage in netplay.cpp.
        Set w = NETbeginEncode(NETbroadcastQueue(), NET_SHARE_GAME_QUEUE)
        NETuint8_t_w w, player
        NETuint32_t_w w, num
        For n = 0 To num - 1
            NETnetMessage_w w, queue.getMessageForNet()
            queue.popMessageForNet
        Next
        NETend_w w
continue_:
    Next
End Sub

Public Sub NETpop(ByVal queue As Long)
    receiveQueue(queue).popMessage
End Sub

' same logic as using NETbytes() for a write, except written to the `output` vector (output(0 .. outputSize - 1))
Public Sub NETbytesOutputToVector(data() As Byte, ByVal dataSize As Long, output() As Byte, outputSize As Long)
    Dim v As Long, moreBytes As Boolean, n As Long, b As Long, i As Long

    ' same as queueAuto(uint32_t) - write the length
    v = dataSize
    moreBytes = True
    n = 0
    Do While moreBytes
        moreBytes = encode_uint32_t(b, v, n)
        pushByte output, outputSize, b
        n = n + 1
    Loop

    ' write all the data bytes
    For i = 0 To dataSize - 1
        pushByte output, outputSize, data(i)
    Next
End Sub

' (VB) output.push_back(b)
Private Sub pushByte(output() As Byte, outputSize As Long, ByVal b As Long)
    If outputSize = 0 Then
        ReDim output(0 To 15)
    ElseIf outputSize > UBound(output) Then
        ReDim Preserve output(0 To outputSize * 2 - 1)
    End If
    output(outputSize) = b
    outputSize = outputSize + 1
End Sub

' TODO Call this function somewhere.
Public Function NETloadReplay(ByVal filename As String, ByVal optionsHandler As ReplayOptionsHandler) As Boolean
    Dim replayFormatVer As Long, newMessage As NetMessage, player As Long, gotReplayEnded As Boolean, skip As Boolean
    Dim level As Long
    replayFormatVer = 0
    If Not NETreplayLoadStart(filename, optionsHandler, replayFormatVer) Then
        NETloadReplay = False
        Exit Function
    End If
    gotReplayEnded = False
    Do While NETreplayLoadNetMessage(newMessage, player)
        ' (VB) the || short-circuited: gameQueues[player] is only read for a player below MAX_PLAYERS or the host
        skip = player >= MAX_PLAYERS And player <> NetPlay.hostPlayer
        If Not skip Then
            If player < MAX_GAMEQUEUE_SLOTS Then skip = gameQueues(player) Is Nothing Else skip = True
        End If
        If skip Then
            If newMessage.type_() <> GAME_GAME_TIME Then level = LOG_ERROR Else level = LOG_INFO
            debug_ 610, level, "NETloadReplay", "Skipping message to player %d in replay.", player
            GoTo continue_
        End If
        If newMessage.type_() = REPLAY_ENDED Then
            gotReplayEnded = True
            Exit Do
        End If
        gameQueues(player).pushMessage newMessage
continue_:
    Loop
    If Not gotReplayEnded And replayFormatVer >= 2 Then
        debug_ 622, LOG_POPUP, "NETloadReplay", "Unable to load replay: The replay file is incomplete or corrupted."
        bIsReplay = True
        NETshutdownReplay
        NETloadReplay = False
        Exit Function
    End If
    ' Add special REPLAY_ENDED message to the end of the host's gameQueue
    gameQueues(NetPlay.hostPlayer).pushMessage NetMessageBuilder_new(REPLAY_ENDED, 0).build()
    NETreplayLoadStop
    bIsReplay = True
    NETloadReplay = True
End Function

Public Function NETisReplay() As Boolean
    NETisReplay = bIsReplay
End Function

Public Sub NETshutdownReplay()
    If bIsReplay Then
        ' extra replay spectator gamequeue
        Set gameQueues(MAX_CONNECTED_PLAYERS) = Nothing
    End If

    bIsReplay = False
End Sub

'--------------------------------------------------------------------------
' Readers

Public Sub NETuint8_t_r(ByVal r As MessageReader, v As Long)
    r.byte_ v
End Sub

Public Sub NETint8_t_r(ByVal r As MessageReader, v As Long)
    Dim b As Long
    NETuint8_t_r r, b                           ' reinterpret_cast<uint8_t&>(val)
    If b >= 128 Then v = b - 256 Else v = b
End Sub

Public Sub NETuint16_t_r(ByVal r As MessageReader, v As Long)
    Dim b0 As Long, b1 As Long
    r.byte_ b0
    r.byte_ b1
    v = b0 * 256 + b1                           ' wz_ntohs_load_unaligned
End Sub

Public Sub NETint16_t_r(ByVal r As MessageReader, v As Long)
    Dim u As Long
    NETuint16_t_r r, u                          ' reinterpret_cast<uint16_t&>(val)
    If u >= 32768 Then v = u - 65536 Else v = u
End Sub

Public Sub NETuint32_t_r(ByVal r As MessageReader, v As Long)
    Dim acc As Long, moreBytes As Boolean, n As Long, b As Long
    acc = 0
    moreBytes = True
    n = 0
    Do While moreBytes
        b = 0
        r.byte_ b
        moreBytes = decode_uint32_t(b, acc, n)
        n = n + 1
    Loop
    v = acc
End Sub

Public Sub NETint32_t_r(ByVal r As MessageReader, v As Long)
    ' Non-negative values: value*2
    ' Negative values:     -value*2 - 1
    ' Example: int32_t -5 -4 -3 -2 -1  0  1  2  3  4  5
    ' becomes uint32_t  9  7  5  3  1  0  2  4  6  8 10
    Dim b As Long
    b = 0
    NETuint32_t_r r, b
    v = SHRU(b, 1) Xor -(b And 1)
End Sub

Public Sub NETuint64_t_r(ByVal r As MessageReader, v As Double)
    Dim b0 As Long, b1 As Long
    NETuint32_t_r r, b0
    NETuint32_t_r r, b1
    v = U32(b0) * TWO_32 + U32(b1)              ' (uint64_t)b[0] << 32 | b[1]
End Sub

Public Sub NETint64_t_r(ByVal r As MessageReader, v As Double)
    Dim b0 As Long, b1 As Long
    NETuint32_t_r r, b0                         ' NETuint64_t(r, reinterpret_cast<uint64_t&>(val))
    NETuint32_t_r r, b1
    v = CDbl(b0) * TWO_32 + U32(b1)
End Sub

Public Sub NETbool_r(ByVal r As MessageReader, v As Boolean)
    Dim b As Long
    NETuint8_t_r r, b
    v = b <> 0
End Sub

Public Sub NETwzstring_r(ByVal r As MessageReader, str_ As String)
    Dim b() As Byte, n As Long
    n = NETstring_std_bytes_r(r, b, 65536)      ' std::string s; NETstring(r, s);
    str_ = utf8_decode(b, 0, n)                 ' WzString::fromUtf8(s)
End Sub

' Receives a string from the current network package.
'  str    The string decoded from the network package (the text before its first NUL).
'  maxlen The buffer size of str (sizeof(str) for static buffers).
'  If while decoding maxlen is smaller than the actual length of the
'  string being decoded, the resulting string will be truncated.
Public Sub NETstring_r(ByVal r As MessageReader, str_ As String, ByVal maxlen As Long)
    Dim length As Long, b() As Byte, n As Long
    NETuint16_t_r r, length
    maxlen = maxlen And &HFFFF&
    If maxlen < length Then length = maxlen

    If length > 0 Then
        ReDim b(0 To length - 1)
        r.bytes b, 0, length
    End If
    ' str[len] = '\0' (the text ends at the first NUL)
    n = 0
    Do While n < length
        If b(n) = 0 Then Exit Do
        n = n + 1
    Loop
    str_ = utf8_decode(b, 0, n)
End Sub

' NETstring(MessageReader &r, std::string& s, uint32_t maxLen = 65536)
Public Sub NETstring_std_r(ByVal r As MessageReader, s As String, Optional ByVal maxLen As Long = 65536)
    Dim b() As Byte, n As Long
    n = NETstring_std_bytes_r(r, b, maxLen)
    s = utf8_decode(b, 0, n)
End Sub

' (VB) NETstring(r, std::string&, maxLen) with the string's bytes in b; returns its length
Private Function NETstring_std_bytes_r(ByVal r As MessageReader, b() As Byte, ByVal maxLen As Long) As Long
    Dim length As Long
    length = 0
    NETuint32_t_r r, length
    If ULT(maxLen, length) Then length = maxLen  ' std::min(len, maxLen)
    If length > 0 Then ReDim b(0 To length - 1) ' s.resize(len) (NULs)
    If r.valid() Then
        If length > 0 Then r.bytes b, 0, length
    End If
    NETstring_std_bytes_r = length
End Function

' NETbin(MessageReader &r, uint8_t *str, uint32_t len): into str(0 .. len_ - 1) (grown if smaller)
Public Sub NETbin_r(ByVal r As MessageReader, str_() As Byte, ByVal len_ As Long)
    If len_ <= 0 Then Exit Sub
    If byteArraySize(str_) = 0 Then
        ReDim str_(0 To len_ - 1)
    ElseIf UBound(str_) < len_ - 1 Then
        ReDim Preserve str_(0 To len_ - 1)
    End If
    r.bytes str_, 0, len_
End Sub

' template NETbytes(MessageReader& r, VecT& vec, unsigned maxLen = 10000) (nettypes.h): vec is the whole array (Erased when empty)
Public Sub NETbytes_r(ByVal r As MessageReader, vec() As Byte, Optional ByVal maxLen As Long = 10000)
    Dim length As Long, vecSize As Long
    '
    ' Strings sent over the network are prefixed with their length, sent as an
    ' unsigned 16-bit integer, not including \0 termination.
    '

    length = 0
    NETuint32_t_r r, length
    If ULT(maxLen, length) Then
        debug_ 112, LOG_ERROR, "NETbytes", "NETbytes: Decoding packet, length %u truncated at %u", length, maxLen
    End If
    If ULT(maxLen, length) Then length = maxLen ' Truncate length if necessary.
    Erase vec                                   ' vec.clear()
    vecSize = 0
    If r.valid() Then
        r.bytesVector vec, vecSize, length
    End If
End Sub

Public Sub NETnetMessage_r(ByVal r As MessageReader, msg As NetMessage)
    Dim rawData() As Byte
    NETbytes_r r, rawData, -1                   ' std::numeric_limits<uint32_t>::max()
    Set msg = NetMessageBuilder_new_raw(rawData, byteArraySize(rawData)).build()
End Sub

'--------------------------------------------------------------------------
' Writers

Public Sub NETuint8_t_w(ByVal w As MessageWriter, ByVal v As Long)
    w.byte_ v And &HFF&
End Sub

Public Sub NETint8_t_w(ByVal w As MessageWriter, ByVal v As Long)
    NETuint8_t_w w, v And &HFF&                 ' static_cast<uint8_t>(val)
End Sub

Public Sub NETuint16_t_w(ByVal w As MessageWriter, ByVal v As Long)
    ' wz_htons_store_unaligned(b, val): network byte order
    NETuint8_t_w w, (v And &HFF00&) \ &H100&
    NETuint8_t_w w, v And &HFF&
End Sub

Public Sub NETint16_t_w(ByVal w As MessageWriter, ByVal v As Long)
    NETuint16_t_w w, v And &HFFFF&              ' static_cast<uint16_t>(val)
End Sub

Public Sub NETuint32_t_w(ByVal w As MessageWriter, ByVal v As Long)
    Dim moreBytes As Boolean, n As Long, b As Long
    moreBytes = True
    n = 0
    Do While moreBytes
        b = 0
        moreBytes = encode_uint32_t(b, v, n)
        NETuint8_t_w w, b
        n = n + 1
    Loop
End Sub

Public Sub NETint32_t_w(ByVal w As MessageWriter, ByVal v As Long)
    ' Non-negative values: value*2
    ' Negative values:     -value*2 - 1
    ' Example: int32_t -5 -4 -3 -2 -1  0  1  2  3  4  5
    ' becomes uint32_t  9  7  5  3  1  0  2  4  6  8 10
    Dim b As Long
    b = SHL(v, 1) Xor -SHRU(v, 31)              ' (uint32_t)val << 1 ^ (-((uint32_t)val >> 31))
    NETuint32_t_w w, b
End Sub

' NETuint64_t(MessageWriter&, uint64_t) and NETint64_t (static_cast<uint64_t>): the value's two 32-bit halves
Public Sub NETuint64_t_w(ByVal w As MessageWriter, ByVal v As Double)
    Dim hi As Double
    hi = Int(v / TWO_32)
    NETuint32_t_w w, FromU32(hi)                ' uint32_t(val >> 32)
    NETuint32_t_w w, FromU32(v - hi * TWO_32)   ' uint32_t(val)
End Sub

Public Sub NETint64_t_w(ByVal w As MessageWriter, ByVal v As Double)
    NETuint64_t_w w, v
End Sub

Public Sub NETbool_w(ByVal w As MessageWriter, ByVal v As Boolean)
    Dim i As Long
    If v Then i = 1 Else i = 0                  ' !!val
    NETuint8_t_w w, i
End Sub

Public Sub NETwzstring_w(ByVal w As MessageWriter, ByVal str_ As String)
    Dim b() As Byte, n As Long
    n = utf8_encode(str_, b)                    ' str.toUtf8()
    If Not (n <= 65535) Then ASSERT_FAILURE "utf8_string.size() <= static_cast<size_t>(std::numeric_limits<uint16_t>::max())", "lib/netplay/nettypes.cpp:879", "NETwzstring", "utf8_string.size() exceeds uint16_t max"
    NETstring_std_bytes_w w, b, n, 65536
End Sub

' NETstring(MessageWriter& w, const char* str, uint16_t maxlen): str's text up to its first NUL
Public Sub NETstring_w(ByVal w As MessageWriter, ByVal str_ As String, ByVal maxlen As Long)
    '
    ' Strings sent over the network are prefixed with their length, sent as an
    ' unsigned 16-bit integer, not including \0 termination.
    '
    Dim length As Long, cappedStrLen As Long, maxReadLen As Long, b() As Byte, n As Long, p As Long
    maxlen = maxlen And &HFFFF&
    p = InStr(str_, vbNullChar)
    If p > 0 Then str_ = Left$(str_, p - 1)
    n = utf8_encode(str_, b)

    length = 0
    If n < maxlen Then cappedStrLen = n + 1 Else cappedStrLen = maxlen  ' strnlen1(str, maxlen)
    If cappedStrLen > 0 Then length = cappedStrLen - 1 Else length = 0
    NETuint16_t_w w, length

    ' Truncate length if necessary
    If maxlen > 0 Then maxReadLen = maxlen - 1 Else maxReadLen = 0
    If length > maxReadLen Then
        debug_ 899, LOG_ERROR, "NETstring", "NETstring: Encoding packet, length %u truncated at %u", length, maxlen
        length = maxReadLen
    End If
    If length > 0 Then w.bytes b, 0, length
End Sub

' NETstring(MessageWriter& w, const std::string& s, uint32_t maxLen = 65536)
Public Sub NETstring_std_w(ByVal w As MessageWriter, ByVal s As String, Optional ByVal maxLen As Long = 65536)
    Dim b() As Byte, n As Long
    n = utf8_encode(s, b)
    NETstring_std_bytes_w w, b, n, maxLen
End Sub

' (VB) NETstring(w, std::string, maxLen) for the string's bytes b(0 .. n - 1)
Private Sub NETstring_std_bytes_w(ByVal w As MessageWriter, b() As Byte, ByVal n As Long, ByVal maxLen As Long)
    Dim length As Long
    If ULT(maxLen, n) Then length = maxLen Else length = n  ' std::min<size_t>(s.size(), maxLen)
    NETuint32_t_w w, length
    NETbin_w w, b, length
End Sub

' NETbin(MessageWriter& w, const uint8_t* str, uint32_t len)
Public Sub NETbin_w(ByVal w As MessageWriter, str_() As Byte, ByVal len_ As Long)
    If len_ > 0 Then w.bytes str_, 0, len_
End Sub

' template NETbytes(MessageWriter& w, const VecT& vec, unsigned maxLen = 10000) (nettypes.h): vec is the whole array
Public Sub NETbytes_w(ByVal w As MessageWriter, vec() As Byte, Optional ByVal maxLen As Long = 10000)
    NETbytes_n_w w, vec, byteArraySize(vec), maxLen
End Sub

' (VB) NETbytes(w, vec, maxLen) for vec(0 .. vecSize - 1)
Private Sub NETbytes_n_w(ByVal w As MessageWriter, vec() As Byte, ByVal vecSize As Long, ByVal maxLen As Long)
    Dim length As Long
    '
    ' Strings sent over the network are prefixed with their length, sent as an
    ' unsigned 16-bit integer, not including \0 termination.
    '

    If Not (vecSize <= 4294967295#) Then ASSERT_FAILURE "vec.size() <= static_cast<size_t>(std::numeric_limits<uint32_t>::max())", "lib/netplay/nettypes.h:162", "NETbytes", "vec.size() exceeds uint32_t max"
    length = vecSize
    If ULT(maxLen, length) Then
        debug_ 166, LOG_ERROR, "NETbytes", "NETbytes: Encoding packet, length %u truncated at %u", length, maxLen
    End If
    If ULT(maxLen, length) Then length = maxLen ' Truncate length if necessary.
    NETuint32_t_w w, length
    If length > 0 Then w.bytes vec, 0, length
End Sub

Public Sub NETnetMessage_w(ByVal w As MessageWriter, ByVal msg As NetMessage)
    Dim rawData() As Byte
    rawData = msg.rawData()
    NETbytes_n_w w, rawData, msg.rawDataSize(), -1   ' std::numeric_limits<uint32_t>::max()
End Sub

'--------------------------------------------------------------------------
' The overload pairs: a MessageReader or a MessageWriter

Public Sub NETuint8_t(ByVal c As Object, v As Long)
    If TypeOf c Is MessageReader Then NETuint8_t_r c, v Else NETuint8_t_w c, v
End Sub

Public Sub NETint8_t(ByVal c As Object, v As Long)
    If TypeOf c Is MessageReader Then NETint8_t_r c, v Else NETint8_t_w c, v
End Sub

Public Sub NETuint16_t(ByVal c As Object, v As Long)
    If TypeOf c Is MessageReader Then NETuint16_t_r c, v Else NETuint16_t_w c, v
End Sub

Public Sub NETint16_t(ByVal c As Object, v As Long)
    If TypeOf c Is MessageReader Then NETint16_t_r c, v Else NETint16_t_w c, v
End Sub

Public Sub NETuint32_t(ByVal c As Object, v As Long)
    If TypeOf c Is MessageReader Then NETuint32_t_r c, v Else NETuint32_t_w c, v
End Sub

Public Sub NETint32_t(ByVal c As Object, v As Long)
    If TypeOf c Is MessageReader Then NETint32_t_r c, v Else NETint32_t_w c, v
End Sub

Public Sub NETuint64_t(ByVal c As Object, v As Double)
    If TypeOf c Is MessageReader Then NETuint64_t_r c, v Else NETuint64_t_w c, v
End Sub

Public Sub NETint64_t(ByVal c As Object, v As Double)
    If TypeOf c Is MessageReader Then NETint64_t_r c, v Else NETint64_t_w c, v
End Sub

Public Sub NETbool(ByVal c As Object, v As Boolean)
    If TypeOf c Is MessageReader Then NETbool_r c, v Else NETbool_w c, v
End Sub

' template NETenum: the enum as a uint32_t
Public Sub NETenum(ByVal c As Object, v As Long)
    If TypeOf c Is MessageReader Then NETuint32_t_r c, v Else NETuint32_t_w c, v
End Sub

Public Sub NETwzstring(ByVal c As Object, s As String)
    If TypeOf c Is MessageReader Then NETwzstring_r c, s Else NETwzstring_w c, s
End Sub

' NETstring(r, char *str, uint16_t maxlen) / NETstring(w, const char *str, uint16_t maxlen)
Public Sub NETstring(ByVal c As Object, s As String, ByVal maxlen As Long)
    If TypeOf c Is MessageReader Then NETstring_r c, s, maxlen Else NETstring_w c, s, maxlen
End Sub

' NETstring(r, std::string &s, uint32_t maxLen) / NETstring(w, const std::string &s, uint32_t maxLen)
Public Sub NETstring_std(ByVal c As Object, s As String, Optional ByVal maxLen As Long = 65536)
    If TypeOf c Is MessageReader Then NETstring_std_r c, s, maxLen Else NETstring_std_w c, s, maxLen
End Sub

Public Sub NETbin(ByVal c As Object, b() As Byte, ByVal len_ As Long)
    If TypeOf c Is MessageReader Then NETbin_r c, b, len_ Else NETbin_w c, b, len_
End Sub

Public Sub NETbytes(ByVal c As Object, b() As Byte, Optional ByVal maxLen As Long = 10000)
    If TypeOf c Is MessageReader Then NETbytes_r c, b, maxLen Else NETbytes_w c, b, maxLen
End Sub

' NETnetMessage(r, NetMessage **msg) / NETnetMessage(w, const NetMessage &msg)
Public Sub NETnetMessage(ByVal c As Object, msg As NetMessage)
    If TypeOf c Is MessageReader Then NETnetMessage_r c, msg Else NETnetMessage_w c, msg
End Sub
