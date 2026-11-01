Attribute VB_Name = "multisync_c"
' src/multisync.cpp: the score broadcasts and the pings (NET_PING with its challenge: a signature by the player's
' shared identity proves it; the host verifies each player's identity this way).
' (VB) std::array<optional<PingChallengeBytes>, MAX_CONNECTED_PLAYERS> pingChallenges: one Byte array per player
' (PING_CHALLENGE_BYTES bytes from pingChallenges(i * PING_CHALLENGE_BYTES)) with has_pingChallenge(i).
Option Explicit
Private Const MAX_CONNECTED_PLAYERS As Long = MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS

' ////////////////////////////////////////////////////////////////////////////
' function definitions

Private Const AV_PING_FREQUENCY As Long = 20000     ' how often to update average pingtimes. in approx millisecs.
Private Const PING_FREQUENCY As Long = 4000         ' how often to update pingtimes. in approx millisecs.

Private PingSend(0 To MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS - 1) As Long   ' UDWORD: stores the time the ping was called.

Private Const PING_CHALLENGE_BYTES As Long = 32
' static_assert(PING_CHALLENGE_BYTES % 8 == 0, "Must be a multiple of 8 bytes");
Private pingChallenges((MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS) * 32 - 1) As Byte  ' Random data sent with the last ping.
Private has_pingChallenge(0 To MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS - 1) As Boolean

' sendPing's statics
Private lastPing As Long                ' static UDWORD lastPing = 0: Last time we sent a ping
Private lastav As Long                  ' static UDWORD lastav = 0: Last time we updated average

' ////////////////////////////////////////////////////////////////////////
' ////////////////////////////////////////////////////////////////////////
' Score
' We use setMultiStats() to broadcast the score when needed.
Public Function sendScoreCheck() As Boolean
    Dim i As Long
    If game.blindMode <> BLIND_MODE_NONE Then
        ' no-op in blind mode
        sendScoreCheck = False
        Exit Function
    End If

    ' Broadcast any changes in other players, but not in FRONTEND!!!
    ' Detection for this no longer uses title mode, but instead game mode, because that makes more sense
    If GetGameMode() = GS_NORMAL Then
        For i = 0 To MAX_CONNECTED_PLAYERS - 1
            ' Host controls AI's scores + his own...
            If myResponsibility(i) Then
                ' Send score to everyone else
                sendMultiStatsScoreUpdates i
            End If
        Next
    End If

    sendScoreCheck = True
End Function

' ////////////////////////////////////////////////////////////////////////
' ////////////////////////////////////////////////////////////////////////
' Pings

Private Function averagePing() As Long
    Dim i As Long, count As Long, total As Long
    count = 0
    total = 0
    For i = 0 To MAX_CONNECTED_PLAYERS - 1
        If isHumanPlayer(i) And (i <> realSelectedPlayer) And ULT(ingame.PingTimes(i), PING_LIMIT) Then
            total = total + ingame.PingTimes(i)
            count = count + 1
        End If
    Next
    If count < 1 Then count = 1                 ' MAX(count, 1)
    averagePing = FromU32(Fix(U32(total) / U32(count)))
End Function

Public Sub multiSyncResetAllChallenges()
    Dim i As Long
    For i = 0 To MAX_CONNECTED_PLAYERS - 1
        has_pingChallenge(i) = False
    Next
End Sub

Public Sub multiSyncResetPlayerChallenge(ByVal playerIdx As Long)
    If Not NetPlay_isHost Then Exit Sub
    has_pingChallenge(playerIdx) = False
End Sub

Public Sub multiSyncPlayerSwap(ByVal playerIndexA As Long, ByVal playerIndexB As Long)
    Dim i As Long, b As Byte, h As Boolean
    If Not NetPlay_isHost Then Exit Sub
    For i = 0 To PING_CHALLENGE_BYTES - 1
        b = pingChallenges(playerIndexA * PING_CHALLENGE_BYTES + i)
        pingChallenges(playerIndexA * PING_CHALLENGE_BYTES + i) = pingChallenges(playerIndexB * PING_CHALLENGE_BYTES + i)
        pingChallenges(playerIndexB * PING_CHALLENGE_BYTES + i) = b
    Next
    h = has_pingChallenge(playerIndexA)
    has_pingChallenge(playerIndexA) = has_pingChallenge(playerIndexB)
    has_pingChallenge(playerIndexB) = h
End Sub

' generatePingChallenge(playerIdx) into pingChallenges' slot (the optional set)
Private Sub generatePingChallenge(ByVal playerIdx As Long, ByVal slot As Long)
    Dim bytes(0 To PING_CHALLENGE_BYTES - 1) As Byte, pBytesDst As Long, hi As Long, lo As Long, i As Long
    If NetPlay_isHost And Not ingame.VerifiedIdentity(playerIdx) Then
        ' generate secure random challenges until the player verifies their identity
        genSecRandomBytes bytes, 0, PING_CHALLENGE_BYTES
    Else
        pBytesDst = 0
        Do While pBytesDst <= PING_CHALLENGE_BYTES - 8
            ' uint64_t pingChallengei = (uint64_t)rand() << 32 | rand(); (MSVC: the left rand() first) memcpy'd little-endian
            hi = rand()
            lo = rand()
            For i = 0 To 3
                bytes(pBytesDst + i) = (lo \ BIT_(8 * i)) And &HFF&
                bytes(pBytesDst + 4 + i) = (hi \ BIT_(8 * i)) And &HFF&
            Next
            pBytesDst = pBytesDst + 8
        Loop
    End If
    For i = 0 To PING_CHALLENGE_BYTES - 1
        pingChallenges(slot * PING_CHALLENGE_BYTES + i) = bytes(i)
    Next
    has_pingChallenge(slot) = True
End Sub

' (VB) pingChallenges[slot].value() as an array of its own
Private Sub challengeBytes(ByVal slot As Long, out() As Byte)
    Dim i As Long
    ReDim out(0 To PING_CHALLENGE_BYTES - 1)
    For i = 0 To PING_CHALLENGE_BYTES - 1
        out(i) = pingChallenges(slot * PING_CHALLENGE_BYTES + i)
    Next
End Sub

Public Function sendPing() As Boolean
    Dim isNew As Boolean, player As Long, i As Long, w As MessageWriter, ch() As Byte
    isNew = True
    player = selectedPlayer

    ' Only ping every so often
    If ULT(realTime, lastPing) Then
        lastPing = 0
    End If

    If ULT(realTime - lastPing, PING_FREQUENCY) Then
        sendPing = True
        Exit Function
    End If

    lastPing = realTime

    ' If host, also update the average ping stat for joiners
    If NetPlay_isHost Then
        If ULT(realTime, lastav) Then
            lastav = 0
        End If

        If ULT(AV_PING_FREQUENCY, realTime - lastav) Then
            NETsetGameFlags 2, averagePing()
            lastav = realTime
        End If
    End If

    '
    ' Before we send the ping, if any player failed to respond to the last one
    ' we should re-enumerate the players.
    '

    For i = 0 To MAX_CONNECTED_PLAYERS - 1
        If isHumanPlayer(i) And PingSend(i) <> 0 And ingame.PingTimes(i) <> 0 And i <> selectedPlayer Then
            ingame.PingTimes(i) = PING_LIMIT
        ElseIf Not isHumanPlayer(i) And PingSend(i) <> 0 And ingame.PingTimes(i) <> 0 And i <> selectedPlayer Then
            ingame.PingTimes(i) = 0
        End If
    Next

    If NetPlay_isHost Then
        ' Generate a unique ping challenge for each player, and send them individually
        For i = 0 To MAX_CONNECTED_PLAYERS - 1
            If isHumanPlayer(i) And i <> selectedPlayer Then
                generatePingChallenge i, i
                challengeBytes i, ch
                Set w = NETbeginEncode(NETnetQueue(i), NET_PING)
                NETuint8_t w, player
                NETbool w, isNew
                NETbin w, ch, PING_CHALLENGE_BYTES
                NETend w

                ' Note when we sent the ping
                PingSend(i) = realTime
            End If
        Next
    Else
        ' Just generate and broadcast the same ping challenge to all other players
        generatePingChallenge 0, 0
        challengeBytes 0, ch
        Set w = NETbeginEncode(NETbroadcastQueue(), NET_PING)
        NETuint8_t w, player
        NETbool w, isNew
        NETbin w, ch, PING_CHALLENGE_BYTES
        NETend w

        ' Note when we sent the ping
        For i = 0 To MAX_CONNECTED_PLAYERS - 1
            PingSend(i) = realTime
        Next
    End If

    sendPing = True
End Function

' inline optional<PingChallengeBytes>& expectedPingChallengeBytes(uint8_t playerIdx): the slot
Private Function expectedPingChallengeSlot(ByVal playerIdx As Long) As Long
    If NetPlay_isHost Then expectedPingChallengeSlot = playerIdx Else expectedPingChallengeSlot = 0
End Function

' accept and process incoming ping messages.
Public Function recvPing(ByVal queue As Long) As Boolean
    Dim isNew As Boolean, sender As Long, us As Long, challenge() As Byte, challengeResponse() As Byte, sigLen As Long
    Dim r As MessageReader, w As MessageWriter, slot As Long, verifiedResponse As Boolean, senderIdentity As EcKey
    Dim oldIsAdminStatus As Boolean, keyBytes() As Byte, keyLen As Long, senderPublicKeyB64 As String, senderIdentityHash As String
    Dim expected() As Byte, localIdentity As EcKey
    isNew = False
    us = selectedPlayer
    ReDim challenge(0 To PING_CHALLENGE_BYTES - 1)

    Set r = NETbeginDecode(queue, NET_PING)
    NETuint8_t r, sender
    NETbool r, isNew
    If isNew Then
        NETbin r, challenge, PING_CHALLENGE_BYTES
    Else
        NETbytes r, challengeResponse
    End If
    NETend r

    If sender >= MAX_CONNECTED_PLAYERS Then
        debug_ 284, LOG_ERROR, "recvPing", "Bad NET_PING packet, sender is %d", sender
        recvPing = False
        Exit Function
    End If

    If whosResponsible(sender) <> NETQUEUE_index(queue) Then
        HandleBadParam "NET_PING given incorrect params.", sender, NETQUEUE_index(queue)
        recvPing = False
        Exit Function
    End If

    ' If this is a new ping, respond to it
    If isNew Then
        localIdentity = getLocalSharedIdentity()
        sigLen = EcKey_sign(localIdentity, challenge, 0, PING_CHALLENGE_BYTES, challengeResponse)

        Set w = NETbeginEncode(NETnetQueue(sender), NET_PING)
        ' We are responding to a new ping
        isNew = False

        NETuint8_t w, us
        NETbool w, isNew
        NETbytes w, challengeResponse
        NETend w
    ' They are responding to one of our pings
    Else
        slot = expectedPingChallengeSlot(sender)
        If Not has_pingChallenge(slot) Then
            ' Someone is sending a NET_PING but we do not have a valid challenge for them... ignore
            recvPing = False
            Exit Function
        End If

        verifiedResponse = False
        senderIdentity = getMultiStats(sender).identity
        If Not EcKey_empty(senderIdentity) Then
            challengeBytes slot, expected
            verifiedResponse = EcKey_verify(senderIdentity, challengeResponse, byteArraySize(challengeResponse), expected, 0, PING_CHALLENGE_BYTES)
        End If
        If Not verifiedResponse Then
            ' Either bad signature, or we sent more than one ping packet and this response is to an older one than the latest.
            debug_ 327, LOG_NEVER, "recvPing", "Bad and/or old NET_PING packet, alleged sender is %d", sender
            recvPing = False
            Exit Function
        End If

        ' Work out how long it took them to respond
        ingame.PingTimes(sender) = FromU32(Fix(U32(realTime - PingSend(sender)) / 2))

        If Not ingame.VerifiedIdentity(sender) Then
            ' Record that they've verified the identity
            ingame.VerifiedIdentity(sender) = True

            If NetPlay_isHost And game.blindMode <> BLIND_MODE_NONE Then
                ' check if verified identity is an admin, and handle changes to admin status
                oldIsAdminStatus = NetPlay_players_isAdmin(sender)
                NetPlay_players_isAdmin(sender) = identityMatchesAdmin(senderIdentity)
                If oldIsAdminStatus <> NetPlay_players_isAdmin(sender) Then
                    ' then send info about admin status changes to all players
                    NETBroadcastPlayerInfo sender
                End If
            End If

            ' Output to stdinterface, if enabled
            keyLen = EcKey_toBytes(senderIdentity, EcKey_Public, keyBytes)
            senderPublicKeyB64 = base64Encode(keyBytes, keyLen)
            senderIdentityHash = EcKey_publicHashString(senderIdentity)
            wz_command_interface_output "WZEVENT: player identity VERIFIED: %u %s %s %s" & vbLf, sender, senderPublicKeyB64, senderIdentityHash, NetPlay_players_IPtextAddress(sender)
        End If

        ' Note that we have received it
        PingSend(sender) = 0
        multiSyncResetPlayerChallenge sender
    End If

    recvPing = True
End Function

' (VB) the size of a byte array (0 when unallocated)
Private Function byteArraySize(b() As Byte) As Long
    On Error GoTo none
    byteArraySize = UBound(b) - LBound(b) + 1
    Exit Function
none:
    byteArraySize = 0
End Function
