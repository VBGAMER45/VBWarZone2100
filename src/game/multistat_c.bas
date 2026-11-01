Attribute VB_Name = "multistat_c"
' src/multistat.cpp, multistat.h: the players' statistics (PLAYERSTATS, the type PLAYERSTATS_ since the variable
' playerStats has its name in VB), their identities (the shared, blind and host-verified join identities, the
' session keys made from them), the profile files (WZ.STA.v3), the NET_PLAYER_STATS messages, the score updates
' and the JSON of the replays and saves. EcKey, SessionKeys and base64 are crc.cpp's (src/shared/sodium).
' The known players' database (KnownPlayersDB: initKnownPlayers, shutdownKnownPlayers, isLocallyKnownPlayer,
' addKnownPlayer, getStoredPlayerOptions, storePlayerMuteOption): SQLite's tables as a JSON file (below).
' C++'s to_json(j, x) is <type>_to_json(j, x) (j becomes a new value) and from_json(j, x) is <type>_from_json(j, x).
Option Explicit
Private Const MAX_CONNECTED_PLAYERS As Long = MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS

' class KnownPlayersDB: the known players (a name and the public key last seen with it) and the player options
' (muted / banned times per key and name). 4.7.0 keeps them in an SQLite database, <write dir>/knownPlayers.db:
' known_players(local_id INTEGER PRIMARY KEY, name TEXT UNIQUE, pk TEXT) and player_options(local_id INTEGER
' PRIMARY KEY, pk TEXT NOT NULL, name TEXT NOT NULL, muted INTEGER, banned INTEGER) with a unique index on (pk, name),
' PRAGMA user_version 1. (VB) SQLite is not ported: the same tables are the JSON file knownPlayers.json in the write
' dir ({"user_version": 1, "known_players": [{"local_id", "name", "pk"}, ...], "player_options": [{"local_id", "pk",
' "name", "muted", "banned"}, ...]}), read when the database opens and written at each transaction's commit; a
' failed write raises an error where SQLite would throw. INSERT OR IGNORE / UPDATE return the rows they changed.
' muted and banned are int64s, kept as decimal text (exact). 4.7.0 writes into muted system_clock's count (MSVC:
' 100 ns ticks since 1970) and reads it back as std::chrono::seconds, converted to a system_clock time point: the
' multiplication by 10^7 overflows int64 and wraps (chronoSecondsToTicks), so whether a stored mute comes back
' (count() > 0) depends on the wrapped value's sign: as the official build does. local_id: SQLite's rowid (the
' largest + 1).

Private Const KNOWN_PLAYERS_FILE As String = "knownPlayers.json"

Private knownPlayersDB As Boolean       ' std::unique_ptr<KnownPlayersDB> knownPlayersDB: made
Private kp_id() As Double, kp_name() As String, kp_pk() As String, kpCount As Long
Private po_id() As Double, po_pk() As String, po_name() As String, po_muted() As Variant, po_banned() As Variant, poCount As Long
Private userVersion As Long
' std::unordered_map<std::string, std::pair<optional<PlayerInfo>, UDWORD>> findPlayerCache
Private fpc_name() As String, fpc_has() As Boolean, fpc_pk() As String, fpc_time() As Long, fpcCount As Long
Private lastCacheClean As Long


Public Type PLAYERSTATS_
    played As Long                      ' uint32_t: propagated stats.
    wins As Long                        ' uint32_t
    losses As Long                      ' uint32_t
    totalKills As Long                  ' uint32_t
    totalScore As Long                  ' uint32_t
    recentKills As Long                 ' uint32_t: score/kills in last game. (total kills - both droids and structures)
    recentDroidsKilled As Long          ' uint32_t
    recentDroidsLost As Long            ' uint32_t
    recentDroidsBuilt As Long           ' uint32_t
    recentStructuresKilled As Long      ' uint32_t
    recentStructuresLost As Long        ' uint32_t
    recentStructuresBuilt As Long       ' uint32_t
    recentScore As Long                 ' uint32_t
    recentResearchComplete As Long      ' uint32_t
    recentPowerLost As Double           ' uint64_t: power lost in last game (i.e. from droids / structures being killed by other players)
    recentDroidPowerLost As Double      ' uint64_t: power lost in last game (from droids being killed by other players)
    recentStructurePowerLost As Double  ' uint64_t: power lost in last game (from structures being killed by other players)
    recentPowerWon As Double            ' uint64_t: power that was destroyed in last game (i.e. from droids / structures being killed)
    recentResearchPotential As Double   ' uint64_t: how many labs were ticking
    recentResearchPerformance As Double ' uint64_t: how many labs were ticking with objective (researching)
    identity As EcKey
End Type

' multistat.h: struct TrueIdentity { const EcKey& identity; bool verified; } (the key copied)
Public Type TrueIdentity
    identity As EcKey
    verified As Boolean
End Type

' ////////////////////////////////////////////////////////////////////////////
' STATS STUFF
' ////////////////////////////////////////////////////////////////////////////
Public playerStats(0 To MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS - 1) As PLAYERSTATS_  ' static PLAYERSTATS playerStats[MAX_CONNECTED_PLAYERS]
Private zeroStats As PLAYERSTATS_
Private blindIdentity As EcKey          ' a freshly-generated identity used for the local client in the current blind room
Private hostVerifiedJoinIdentities(0 To MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS - 1) As EcKey

' structure.h: static inline bool isStructure(SIMPLE_OBJECT const *psObject)
Private Function isStructure(ByVal psObject As Long) As Boolean
    If psObject <> 0 Then isStructure = SIMPLE_OBJECT_type(psObject) = OBJ_STRUCTURE
End Function

' ////////////////////////////////////////////////////////////////////////////
' Get Player's stats
Public Function getMultiStats(ByVal player As Long) As PLAYERSTATS_
    getMultiStats = playerStats(player)
End Function

Public Function generateBlindIdentity() As Boolean
    EcKey_generate blindIdentity           ' Generate new identity
    If EcKey_empty(blindIdentity) Then ASSERT_FAILURE "!blindIdentity.empty()", "src/multistat.cpp:71", "generateBlindIdentity", "Failed to generate new blind identity"
    generateBlindIdentity = Not EcKey_empty(blindIdentity)
End Function

Public Function getLocalSharedIdentity() As EcKey
    If NetPlay_isHost And realSelectedPlayer >= MAX_PLAYER_SLOTS Then
        ' If spectator host, send real identity
        getLocalSharedIdentity = playerStats(realSelectedPlayer).identity
        Exit Function
    End If

    If game.blindMode <> BLIND_MODE_NONE Then
        ' In blind mode, share the blind identity
        getLocalSharedIdentity = blindIdentity
    Else
        ' In regular mode, share the actual identity
        getLocalSharedIdentity = playerStats(realSelectedPlayer).identity
    End If
End Function

Public Function getVerifiedJoinIdentity(ByVal player As Long) As EcKey
    If Not ULT(player, MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "player < ((11 + 2) + 10)", "src/multistat.cpp:97", "getVerifiedJoinIdentity", "Invalid player: %u", player
    If NetPlay_isHost Then
        getVerifiedJoinIdentity = hostVerifiedJoinIdentities(player)
    Else
        If game.blindMode = BLIND_MODE_NONE And ingame.VerifiedIdentity(player) Then
            getVerifiedJoinIdentity = playerStats(player).identity
        Else
            getVerifiedJoinIdentity = hostVerifiedJoinIdentities(player)
        End If
    End If
End Function

' In blind games, it returns the verified join identity (if executed on the host, or on all clients after the game has ended)
' In regular games, it returns the current player identity
Public Function getTruePlayerIdentity(ByVal player As Long) As TrueIdentity
    If game.blindMode <> BLIND_MODE_NONE Then
        ' In blind games, always output the join identity (since the internal shared identity is a random one)
        getTruePlayerIdentity.identity = hostVerifiedJoinIdentities(player)
        getTruePlayerIdentity.verified = Not EcKey_empty(hostVerifiedJoinIdentities(player))
    Else
        ' Otherwise, always return the *current* player identity (which may differ from the identity used to join the game if the player switched to a different profile)
        getTruePlayerIdentity.identity = playerStats(player).identity
        getTruePlayerIdentity.verified = ingame.VerifiedIdentity(player)
    End If
End Function

' Should be used when a player identity is to be output (in logs, or otherwise accessible to the user)
Public Function getOutputPlayerIdentity(ByVal player As Long) As EcKey
    getOutputPlayerIdentity = getTruePlayerIdentity(player).identity
End Function

' (VB) NETsetSessionKeys(playerIndex, SessionKeys(me, me_playerIdx, other, other_playerIdx)): the constructor's
' std::invalid_argument is a False return, its what() in errorText; the keys cross to WzNet as their bytes
Private Function makeAndSetSessionKeys(ByVal playerIndex As Long, me_ As EcKey, ByVal me_playerIdx As Long, other As EcKey, _
        errorText As String) As Boolean
    Dim sk As SessionKeys
    If Not SessionKeys_init(sk, me_, me_playerIdx, other, playerIndex) Then
        errorText = SessionKeys_lastError()
        makeAndSetSessionKeys = False
        Exit Function
    End If
    NETsetSessionKeys playerIndex, sk.receiveKey, sk.sendKey
    makeAndSetSessionKeys = True
End Function

' (VB) generateSessionKeysWithPlayer's throw: False with the exception's what() in errorText
Private Function generateSessionKeysWithPlayer(ByVal playerIndex As Long, errorText As String) As Boolean
    Dim localIdentity As EcKey
    errorText = ""
    If EcKey_empty(playerStats(playerIndex).identity) Then
        NETclearSessionKeys_player playerIndex
        generateSessionKeysWithPlayer = False
        Exit Function
    End If

    ' generate session keys
    localIdentity = getLocalSharedIdentity()
    If Not makeAndSetSessionKeys(playerIndex, localIdentity, realSelectedPlayer, playerStats(playerIndex).identity, errorText) Then
        NETclearSessionKeys_player playerIndex
        generateSessionKeysWithPlayer = False      ' (the throw, which the callers catch)
        Exit Function
    End If
    generateSessionKeysWithPlayer = True
End Function

Public Function swapPlayerMultiStatsLocal(ByVal playerIndexA As Long, ByVal playerIndexB As Long) As Boolean
    Dim t As PLAYERSTATS_, k As EcKey, e As String
    If Not ULT(playerIndexA, MAX_CONNECTED_PLAYERS) Or Not ULT(playerIndexB, MAX_CONNECTED_PLAYERS) Then
        swapPlayerMultiStatsLocal = False
        Exit Function
    End If
    t = playerStats(playerIndexA): playerStats(playerIndexA) = playerStats(playerIndexB): playerStats(playerIndexB) = t
    k = hostVerifiedJoinIdentities(playerIndexA)
    hostVerifiedJoinIdentities(playerIndexA) = hostVerifiedJoinIdentities(playerIndexB)
    hostVerifiedJoinIdentities(playerIndexB) = k

    ' NOTE: We can't just swap session keys - we have to re-generate to be sure they are correct
    ' (since client / server determinism can also be based on the playerIdx relative to the realSelectedPlayer - see SessionKeys constructor)
    If playerIndexA <> realSelectedPlayer And (playerIndexA < MAX_PLAYERS Or playerIndexA = NetPlay_hostPlayer) Then
        If Not generateSessionKeysWithPlayer(playerIndexA, e) Then
            If Len(e) > 0 Then debug_ 174, LOG_INFO, "swapPlayerMultiStatsLocal", "Cannot create session keys: (self: %u), (other: %u, name: ""%s""), with error: %s", realSelectedPlayer, playerIndexA, getPlayerName(playerIndexA), e
        End If
    End If
    If playerIndexB <> realSelectedPlayer And (playerIndexB < MAX_PLAYERS Or playerIndexB = NetPlay_hostPlayer) Then
        If Not generateSessionKeysWithPlayer(playerIndexB, e) Then
            If Len(e) > 0 Then debug_ 183, LOG_INFO, "swapPlayerMultiStatsLocal", "Cannot create session keys: (self: %u), (other: %u, name: ""%s""), with error: %s", realSelectedPlayer, playerIndexB, getPlayerName(playerIndexB), e
        End If
    End If

    swapPlayerMultiStatsLocal = True
End Function

' recipientPlayerIndex: -1 for nullopt
Private Function sendMultiStatsInternal(ByVal playerIndex As Long, Optional ByVal recipientPlayerIndex As Long = -1, _
        Optional ByVal sendHostVerifiedJoinIdentity As Boolean = False) As Boolean
    Dim queue As Long, w As MessageWriter, pStatsToSend As PLAYERSTATS_, identityPublicKey() As Byte, n As Long
    Dim isHostVerifiedIdentity As Boolean, identity As EcKey
    If Not (NetPlay_isHost Or playerIndex = realSelectedPlayer) Then ASSERT_FAILURE "NetPlay.isHost || playerIndex == realSelectedPlayer", "src/multistat.cpp:191", "sendMultiStatsInternal", "Huh?"

    If recipientPlayerIndex < 0 Then
        queue = NETbroadcastQueue()
    Else
        queue = NETnetQueue(recipientPlayerIndex)
    End If

    ' Now send it to all other players
    Set w = NETbeginEncode(queue, NET_PLAYER_STATS)
    ' Send the ID of the player's stats we're updating
    NETuint32_t w, playerIndex

    pStatsToSend = playerStats(playerIndex)
    If game.blindMode <> BLIND_MODE_NONE Then
        ' In blind mode, always send zeroed stats
        pStatsToSend = zeroStats
    End If

    NETuint32_t w, pStatsToSend.played
    NETuint32_t w, pStatsToSend.wins
    NETuint32_t w, pStatsToSend.losses
    NETuint32_t w, pStatsToSend.totalKills
    NETuint32_t w, pStatsToSend.totalScore

    isHostVerifiedIdentity = False

    ' Choose the identity to send
    If Not sendHostVerifiedJoinIdentity Or Not NetPlay_isHost Then
        If playerIndex = realSelectedPlayer Then
            ' Local client sending its own identity
            identity = getLocalSharedIdentity()
            If Not EcKey_empty(identity) Then
                n = EcKey_toBytes(identity, EcKey_Public, identityPublicKey)
            End If
        Else
            ' Host sending other player details - relay client provided details
            If Not EcKey_empty(playerStats(playerIndex).identity) Then
                n = EcKey_toBytes(playerStats(playerIndex).identity, EcKey_Public, identityPublicKey)
            End If
        End If
    Else
        ' Once game has begun or ended (depending on settings), if we're the host, send the hostVerifiedJoinIdentity
        If Not (NetPlay_isHost And Not isBlindPlayerInfoState()) Then ASSERT_FAILURE "NetPlay.isHost && !isBlindPlayerInfoState()", "src/multistat.cpp:246", "sendMultiStatsInternal", "Not time to send host verified identity yet?"
        isHostVerifiedIdentity = True
        If Not EcKey_empty(hostVerifiedJoinIdentities(playerIndex)) Then
            n = EcKey_toBytes(hostVerifiedJoinIdentities(playerIndex), EcKey_Public, identityPublicKey)
        End If
    End If

    NETbool w, isHostVerifiedIdentity
    NETbytes w, identityPublicKey
    NETend w
    sendMultiStatsInternal = True
End Function

' recipientPlayerIndex: -1 for nullopt
Public Function sendMultiStats(ByVal playerIndex As Long, Optional ByVal recipientPlayerIndex As Long = -1) As Boolean
    sendMultiStats = sendMultiStatsInternal(playerIndex, recipientPlayerIndex, False)
End Function

Public Function sendMultiStatsHostVerifiedIdentities(ByVal playerIndex As Long) As Boolean
    If Not NetPlay_isHost Then
        ASSERT_FAILURE "false", "src/multistat.cpp:268", "sendMultiStatsHostVerifiedIdentities", "Host only routine detected for client!"
        sendMultiStatsHostVerifiedIdentities = False
        Exit Function
    End If
    If isBlindPlayerInfoState() Then ASSERT_FAILURE "!isBlindPlayerInfoState()", "src/multistat.cpp:269", "sendMultiStatsHostVerifiedIdentities", "Not time to send host verified identities yet": sendMultiStatsHostVerifiedIdentities = False: Exit Function
    sendMultiStatsHostVerifiedIdentities = sendMultiStatsInternal(playerIndex, -1, True)
End Function

' ////////////////////////////////////////////////////////////////////////////
' Set Player's stats
' send stats to all players when bLocal is false
Public Function setMultiStats(ByVal playerIndex As Long, plStats As PLAYERSTATS_, ByVal bLocal As Boolean) As Boolean
    Dim localIdentity As EcKey, i As Long, e As String
    If Not ULT(playerIndex, MAX_CONNECTED_PLAYERS) Then
        setMultiStats = True
        Exit Function
    End If

    ' First copy over the data into our local array
    playerStats(playerIndex) = plStats

    If Not bLocal And (NetPlay_isHost Or playerIndex = realSelectedPlayer) Then
        sendMultiStats playerIndex

        If playerIndex = realSelectedPlayer Then
            ' need to clear and re-generate any session keys for communication between us and other players
            NETclearSessionKeys
            localIdentity = getLocalSharedIdentity()
            If EcKey_hasPrivate(localIdentity) Then
                For i = 0 To MAX_CONNECTED_PLAYERS - 1
                    If i = realSelectedPlayer Then GoTo continue_
                    If i >= MAX_PLAYERS And i <> NetPlay_hostPlayer Then
                        ' Don't bother creating SessionKeys with non-host spectator slots
                        GoTo continue_
                    End If
                    If EcKey_empty(playerStats(i).identity) Then
                        GoTo continue_
                    End If
                    If Not makeAndSetSessionKeys(i, localIdentity, realSelectedPlayer, playerStats(i).identity, e) Then
                        debug_ 313, LOG_INFO, "setMultiStats", "One or both identities can't be used for session keys (self: %u, other: %u), with error: %s", realSelectedPlayer, i, e
                    End If
continue_:
                Next
            Else
                ASSERT_FAILURE "false", "src/multistat.cpp:319", "setMultiStats", "Local identity is missing key pair?"
            End If
        End If
    End If

    setMultiStats = True
End Function

Public Function clearPlayerMultiStats(ByVal playerIndex As Long) As Boolean
    Dim blank As PLAYERSTATS_
    setMultiStats playerIndex, blank, True  ' local only
    If NetPlay_isHost Then
        EcKey_clear hostVerifiedJoinIdentities(playerIndex)
    End If
    clearPlayerMultiStats = True
End Function

Public Function sendMultiStatsScoreUpdates(ByVal playerIndex As Long) As Boolean
    If game.blindMode <> BLIND_MODE_NONE Then
        ' No-op if in blind mode
        sendMultiStatsScoreUpdates = False
        Exit Function
    End If

    If NetPlay_isHost Or playerIndex = realSelectedPlayer Then
        sendMultiStatsScoreUpdates = sendMultiStats(playerIndex)
        Exit Function
    End If
    sendMultiStatsScoreUpdates = False
End Function

' (VB) EcKey::Key == EcKey::Key: the sizes and the bytes
Private Function keyBytesEqual(a() As Byte, ByVal aLen As Long, b() As Byte, ByVal bLen As Long) As Boolean
    Dim i As Long
    If aLen <> bLen Then Exit Function
    For i = 0 To aLen - 1
        If a(i) <> b(i) Then Exit Function
    Next
    keyBytesEqual = True
End Function

' identity: the public key's bytes identity(0 .. identityLen - 1)
Private Function multiStatsSetIdentity(ByVal playerIndex As Long, identity() As Byte, ByVal identityLen As Long, _
        Optional ByVal verifiedFromJoin As Boolean = False) As Boolean
    Dim prevIdentity() As Byte, prevLen As Long, senderPublicKey() As Byte, n As Long, senderPublicKeyB64 As String
    Dim senderIdentityHash As String, sendername As String, nameBytes() As Byte, nameLen As Long, senderNameB64 As String
    Dim hasOptions As Boolean, mutedTimeHas As Boolean, mutedTimeCount As Double, e As String

    If Not EcKey_empty(playerStats(playerIndex).identity) Then
        prevLen = EcKey_toBytes(playerStats(playerIndex).identity, EcKey_Public, prevIdentity)
    End If

    ' If game hasn't actually started, process potential identity changes
    If Not ingame.has_TimeEveryoneIsInGame Then
        EcKey_clear playerStats(playerIndex).identity
        If identityLen > 0 Then
            If Not EcKey_fromBytes(playerStats(playerIndex).identity, identity, identityLen, EcKey_Public) Then
                debug_ 368, LOG_INFO, "multiStatsSetIdentity", "Player sent invalid identity: (player: %u, name: ""%s"", IP: %s)", playerIndex, getPlayerName(playerIndex), NetPlay_players_IPtextAddress(playerIndex)
            End If
        Else
            debug_ 373, LOG_INFO, "multiStatsSetIdentity", "Player sent empty identity: (player: %u, name: ""%s"", IP: %s)", playerIndex, getPlayerName(playerIndex), NetPlay_players_IPtextAddress(playerIndex)
        End If

        If Not keyBytesEqual(identity, identityLen, prevIdentity, prevLen) Or identityLen = 0 Then
            If GetGameMode() = GS_NORMAL Then
                debug_ 379, LOG_INFO, "multiStatsSetIdentity", "Unexpected identity change after NET_FIREUP for: (player: %u, name: ""%s"", IP: %s)", playerIndex, getPlayerName(playerIndex), NetPlay_players_IPtextAddress(playerIndex)
            End If
            ingame.PingTimes(playerIndex) = PING_LIMIT
            If verifiedFromJoin And identityLen > 0 Then
                ' Record that the identity has been verified
                ingame.VerifiedIdentity(playerIndex) = True
                If NetPlay_isHost Then
                    NetPlay_players_isAdmin(playerIndex) = identityMatchesAdmin(playerStats(playerIndex).identity)
                    ' Do not broadcast player info here, as it's assumed caller will do it
                End If
                ' Store the verified join identity
                EcKey_fromBytes hostVerifiedJoinIdentities(playerIndex), identity, identityLen, EcKey_Public
                ' Do *not* output to stdinterface here - the join event is still being processed
            Else
                ingame.VerifiedIdentity(playerIndex) = False
                If NetPlay_isHost Then
                    ' when verified identity is cleared, so is admin status (until new identity is verified)
                    If NetPlay_players_isAdmin(playerIndex) Then
                        NetPlay_players_isAdmin(playerIndex) = False
                        NETBroadcastPlayerInfo playerIndex
                    End If
                End If
                ' Output to stdinterface, if enabled
                If identityLen > 0 Then
                    n = EcKey_toBytes(playerStats(playerIndex).identity, EcKey_Public, senderPublicKey)
                    senderPublicKeyB64 = base64Encode(senderPublicKey, n)
                    senderIdentityHash = EcKey_publicHashString(playerStats(playerIndex).identity)
                    sendername = getPlayerName(playerIndex)
                    nameLen = utf8_encode(sendername, nameBytes)
                    senderNameB64 = base64Encode(nameBytes, nameLen)
                    wz_command_interface_output "WZEVENT: player identity UNVERIFIED: %u %s %s %s %s" & vbLf, playerIndex, senderPublicKeyB64, senderIdentityHash, senderNameB64, NetPlay_players_IPtextAddress(playerIndex)
                Else
                    wz_command_interface_output "WZEVENT: player identity EMPTY: %u" & vbLf, playerIndex
                End If
            End If

            If Not ingame.muteChat(playerIndex) Then
                ' check if the new identity was previously muted
                hasOptions = getStoredPlayerOptions_mutedTime(NetPlay_players_name(playerIndex), playerStats(playerIndex).identity, mutedTimeHas, mutedTimeCount)
                If hasOptions And mutedTimeHas Then
                    ingame.muteChat(playerIndex) = (mutedTimeCount > 0)
                End If
            End If

            If playerIndex < MAX_PLAYERS Or playerIndex = NetPlay_hostPlayer Then
                If Not EcKey_empty(playerStats(playerIndex).identity) Then
                    ' generate session keys
                    If Not generateSessionKeysWithPlayer(playerIndex, e) Then
                        If Len(e) > 0 Then debug_ 446, LOG_INFO, "multiStatsSetIdentity", "Cannot create session keys: (self: %u), (other: %u, name: ""%s"", IP: %s), with error: %s", realSelectedPlayer, playerIndex, getPlayerName(playerIndex), NetPlay_players_IPtextAddress(playerIndex), e
                    End If
                Else
                    NETclearSessionKeys_player playerIndex
                End If
            End If
            multiStatsSetIdentity = True
            Exit Function
        End If
    Else
        ' Changing an identity should not happen once a game starts
        If Not keyBytesEqual(identity, identityLen, prevIdentity, prevLen) Or identityLen = 0 Then
            If Not ingame.has_endTime Then
                ASSERT_FAILURE "false", "src/multistat.cpp:465", "multiStatsSetIdentity", "Cannot change identity for player %u after game has started", playerIndex
            End If
        End If
    End If

    multiStatsSetIdentity = False
End Function

Public Function recvMultiStats(ByVal queue As Long) As Boolean
    Dim playerIndex As Long, r As MessageReader, identity() As Byte, identityLen As Long, isHostVerifiedIdentity As Boolean
    Set r = NETbeginDecode(queue, NET_PLAYER_STATS)
    ' Retrieve the ID number of the player for which we need to
    ' update the stats
    NETuint32_t r, playerIndex

    If Not ULT(playerIndex, MAX_CONNECTED_PLAYERS) Then
        NETend r
        recvMultiStats = False
        Exit Function
    End If

    If playerIndex <> NETQUEUE_index(queue) And NETQUEUE_index(queue) <> NetPlay_hostPlayer Then
        HandleBadParam "NET_PLAYER_STATS given incorrect params.", playerIndex, NETQUEUE_index(queue)
        NETend r
        recvMultiStats = False
        Exit Function
    End If

    ' we don't what to update ourselves, we already know our score (FIXME: rewrite setMultiStats())
    If Not myResponsibility(playerIndex) Then
        ' Retrieve the actual stats
        NETuint32_t r, playerStats(playerIndex).played
        NETuint32_t r, playerStats(playerIndex).wins
        NETuint32_t r, playerStats(playerIndex).losses
        NETuint32_t r, playerStats(playerIndex).totalKills
        NETuint32_t r, playerStats(playerIndex).totalScore
        isHostVerifiedIdentity = False
        NETbool r, isHostVerifiedIdentity
        NETbytes r, identity
        identityLen = byteArraySize(identity)
        NETend r

        If Not isHostVerifiedIdentity Then
            multiStatsSetIdentity playerIndex, identity, identityLen, False
        Else
            If NETQUEUE_index(queue) = NetPlay_hostPlayer Then
                EcKey_clear hostVerifiedJoinIdentities(playerIndex)
                If identityLen > 0 Then
                    If Not EcKey_fromBytes(hostVerifiedJoinIdentities(playerIndex), identity, identityLen, EcKey_Public) Then
                        debug_ 523, LOG_INFO, "recvMultiStats", "Host sent invalid host-verified join identity for: (player: %u, name: ""%s"")", playerIndex, getPlayerName(playerIndex)
                    End If
                End If
            End If
        End If
    Else
        NETend r
    End If

    recvMultiStats = True
End Function

' (VB) the size of a byte array (0 when unallocated)
Private Function byteArraySize(b() As Byte) As Long
    On Error GoTo none
    byteArraySize = UBound(b) - LBound(b) + 1
    Exit Function
none:
    byteArraySize = 0
End Function

' identity: the public key's bytes identity(0 .. identityLen - 1) (identityLen omitted: the whole array)
Public Sub multiStatsSetVerifiedIdentityFromJoin(ByVal playerIndex As Long, identity() As Byte, Optional ByVal identityLen As Long = -1)
    If identityLen = -1 Then identityLen = byteArraySize(identity)
    If Not NetPlay_isHost Then
        ASSERT_FAILURE "false", "src/multistat.cpp:538", "multiStatsSetVerifiedIdentityFromJoin", "Host only routine detected for client!"
        Exit Sub
    End If
    If playerIndex = NetPlay_hostPlayer Then ASSERT_FAILURE "playerIndex != NetPlay.hostPlayer", "src/multistat.cpp:539", "multiStatsSetVerifiedIdentityFromJoin", "playerIndex is hostPlayer? (%u)", NetPlay_hostPlayer: Exit Sub
    multiStatsSetIdentity playerIndex, identity, identityLen, True
End Sub

' identity: the public key's bytes identity(0 .. identityLen - 1) (identityLen omitted: the whole array)
Public Sub multiStatsSetVerifiedHostIdentityFromJoin(identity() As Byte, Optional ByVal identityLen As Long = -1)
    If identityLen = -1 Then identityLen = byteArraySize(identity)
    If Not (NetPlay_isHost Or NetPlay_isHostAlive) Then ASSERT_FAILURE "NetPlay.isHost || NetPlay.isHostAlive", "src/multistat.cpp:545", "multiStatsSetVerifiedHostIdentityFromJoin", "Unexpected state when called": Exit Sub
    EcKey_fromBytes hostVerifiedJoinIdentities(NetPlay_hostPlayer), identity, identityLen, EcKey_Public
End Sub

' ////////////////////////////////////////////////////////////////////////////
' Load Player Stats

' (VB) sscanf's %u: the unsigned decimal at s's position p (skipping white space); False when there is none
Private Function scanUnsigned(ByVal s As String, p As Long, value As Long) As Boolean
    Dim v As Double, digits As Long, c As String
    Do While p <= Len(s)
        c = Mid$(s, p, 1)
        If c <> " " And c <> vbLf And c <> vbCr And c <> vbTab And c <> vbVerticalTab And c <> vbFormFeed Then Exit Do
        p = p + 1
    Loop
    Do While p <= Len(s)
        c = Mid$(s, p, 1)
        If c < "0" Or c > "9" Then Exit Do
        v = v * 10 + (AscW(c) - 48)
        If v >= 4294967296# Then v = v - 4294967296#
        digits = digits + 1
        p = p + 1
    Loop
    If digits = 0 Then Exit Function
    value = FromU32(v)
    scanUnsigned = True
End Function

Private Function loadMultiStatsFile(ByVal fileName As String, st As PLAYERSTATS_, Optional ByVal skipLoadingIdentity As Boolean = False) As Boolean
    Dim pFileData() As Byte, size As Long, s As String, identity As String, p As Long, c As String, ok As Boolean
    Dim keyBytes() As Byte, keyLen As Long
    If loadFile(fileName, pFileData, size) Then
        s = StrConv(LeftB$(pFileData, size), vbUnicode)
        If Left$(s, 9) <> "WZ.STA.v3" Then
            Erase pFileData
            loadMultiStatsFile = False ' wrong version or not a stats file
            Exit Function
        End If

        identity = ""
        ' sscanf(pFileData, "WZ.STA.v3\n%u %u %u %u %u\n%1000[A-Za-z0-9+/=]", ...) (the \n's match any white space)
        p = 10
        ok = scanUnsigned(s, p, st.wins)
        If ok Then ok = scanUnsigned(s, p, st.losses)
        If ok Then ok = scanUnsigned(s, p, st.totalKills)
        If ok Then ok = scanUnsigned(s, p, st.totalScore)
        If ok Then ok = scanUnsigned(s, p, st.played)
        If ok And Not skipLoadingIdentity Then
            Do While p <= Len(s)
                c = Mid$(s, p, 1)
                If c <> " " And c <> vbLf And c <> vbCr And c <> vbTab Then Exit Do
                p = p + 1
            Loop
            Do While p <= Len(s) And Len(identity) < 1000
                c = Mid$(s, p, 1)
                If Not ((c >= "A" And c <= "Z") Or (c >= "a" And c <= "z") Or (c >= "0" And c <= "9") Or c = "+" Or c = "/" Or c = "=") Then Exit Do
                identity = identity & c
                p = p + 1
            Loop
        End If
        Erase pFileData

        If Len(identity) > 0 Then
            keyLen = base64Decode(identity, keyBytes)
            If Not EcKey_fromBytes(st.identity, keyBytes, keyLen, EcKey_Private) Then
                debug_ 583, LOG_INFO, "loadMultiStatsFile", "Failed to load profile identity"
            End If
        End If
    End If
    loadMultiStatsFile = True
End Function

' sPlayerName: char * (an empty name becomes "Player", in the caller's variable as in C++)
Public Function loadMultiStats(sPlayerName As String, st As PLAYERSTATS_) As Boolean
    Dim currentIdentity As EcKey, blank As PLAYERSTATS_, fileName As String
    ' preserve current player identity (if loaded)
    currentIdentity = st.identity
    st = blank  ' clear in case we don't get to load

    ' Prevent an empty player name (where the first byte is a 0x0 terminating char already)
    If Len(sPlayerName) = 0 Then
        sPlayerName = "Player"
    End If

    fileName = MultiPlayersPath & sPlayerName & ".sta2"

    debug_ 606, LOG_WZ, "loadMultiStats", "loadMultiStats: %s", fileName

    ' check player .sta2 already exists
    If PHYSFS_exists(fileName) <> 0 Then
        If Not loadMultiStatsFile(fileName, st) Then
            loadMultiStats = False
            Exit Function
        End If
    Else
        ' one-time porting of old .sta player files to .sta2
        fileName = MultiPlayersPath & sPlayerName & ".sta"
        If PHYSFS_exists(fileName) <> 0 Then
            If Not loadMultiStatsFile(fileName, st, True) Then
                loadMultiStats = False
                Exit Function
            End If
        End If
    End If

    If EcKey_empty(st.identity) Or Not EcKey_hasPrivate(st.identity) Then
        If Not EcKey_empty(currentIdentity) Then
            st.identity = currentIdentity       ' Preserve existing identity when creating a new profile
        Else
            EcKey_generate st.identity          ' Generate new identity
        End If
        saveMultiStats sPlayerName, sPlayerName, st  ' Save new profile
    End If

    ' reset recent scores
    st.recentKills = 0
    st.recentDroidsKilled = 0
    st.recentDroidsLost = 0
    st.recentDroidsBuilt = 0
    st.recentStructuresKilled = 0
    st.recentStructuresLost = 0
    st.recentStructuresBuilt = 0
    st.recentScore = 0
    st.recentResearchComplete = 0
    st.recentPowerLost = 0
    st.recentDroidPowerLost = 0
    st.recentStructurePowerLost = 0
    st.recentPowerWon = 0
    st.recentResearchPerformance = 0
    st.recentResearchPotential = 0

    loadMultiStats = True
End Function

' ////////////////////////////////////////////////////////////////////////////
' Save Player Stats
Public Function saveMultiStats(ByVal sFileName As String, ByVal sPlayerName As String, st As PLAYERSTATS_) As Boolean
    Dim buffer As String, keyBytes() As Byte, keyLen As Long, fileName As String, data() As Byte, n As Long
    If Cheated Then
        saveMultiStats = False
        Exit Function
    End If
    If EcKey_empty(st.identity) Then
        debug_ 675, LOG_INFO, "saveMultiStats", "Refusing to save profile with empty identity: %s", sFileName
        saveMultiStats = False
        Exit Function
    End If
    keyLen = EcKey_toBytes(st.identity, EcKey_Private, keyBytes)
    buffer = "WZ.STA.v3" & vbLf & CStr(U32(st.wins)) & " " & CStr(U32(st.losses)) & " " & CStr(U32(st.totalKills)) & " " & _
        CStr(U32(st.totalScore)) & " " & CStr(U32(st.played)) & vbLf & base64Encode(keyBytes, keyLen) & vbLf
    fileName = MultiPlayersPath & sFileName & ".sta2"
    n = utf8_encode(buffer, data)
    saveFile fileName, data, n
    saveMultiStats = True
End Function

' ////////////////////////////////////////////////////////////////////////////
' score update functions

' update players damage stats.
Public Sub updateMultiStatsDamage(ByVal attacker As Long, ByVal defender As Long, ByVal inflicted As Long)
    If defender = PLAYER_FEATURE Then
        ' damaging features like skyscrapers does not count
        Exit Sub
    End If

    If Not ULT(attacker, MAX_PLAYERS) Then ASSERT_FAILURE "attacker < MAX_PLAYERS", "src/multistat.cpp:701", "updateMultiStatsDamage", "invalid attacker: %u", attacker: Exit Sub
    If Not ULT(defender, MAX_PLAYERS) Then ASSERT_FAILURE "defender < MAX_PLAYERS", "src/multistat.cpp:702", "updateMultiStatsDamage", "invalid defender: %u", defender: Exit Sub

    If NetPlay_bComms Then
        If (attacker = scavengerSlot()) Or (defender = scavengerSlot()) Then
            Exit Sub ' damaging and getting damaged by scavengers does not influence scores in MP games
        End If
        playerStats(attacker).totalScore = playerStats(attacker).totalScore + 2 * inflicted
        playerStats(defender).totalScore = playerStats(defender).totalScore - inflicted
    End If

    playerStats(attacker).recentScore = playerStats(attacker).recentScore + 2 * inflicted
    playerStats(defender).recentScore = playerStats(defender).recentScore - inflicted
End Sub

' update games played.
Public Sub updateMultiStatsGames()
    If selectedPlayer >= MAX_PLAYERS Then
        Exit Sub
    End If
    playerStats(selectedPlayer).played = playerStats(selectedPlayer).played + 1
End Sub

' games won
Public Sub updateMultiStatsWins()
    If selectedPlayer >= MAX_PLAYERS Then
        Exit Sub
    End If
    playerStats(selectedPlayer).wins = playerStats(selectedPlayer).wins + 1
End Sub

'games lost.
Public Sub updateMultiStatsLoses()
    If selectedPlayer >= MAX_PLAYERS Then
        Exit Sub
    End If
    playerStats(selectedPlayer).losses = playerStats(selectedPlayer).losses + 1
End Sub

Private Function calcObjectCost(ByVal psObj As Long) As Long
    Select Case SIMPLE_OBJECT_type(psObj)
    Case OBJ_DROID
        calcObjectCost = calcDroidPower(psObj)
        Exit Function
    Case OBJ_STRUCTURE
        If aStruct(psObj).pStructureType < 0 Then ASSERT_FAILURE "psStruct->pStructureType != nullptr", "src/multistat.cpp:757", "calcObjectCost", "pStructureType is null?": calcObjectCost = 0: Exit Function
        calcObjectCost = asStructureStats(aStruct(psObj).pStructureType).powerToBuild
        Exit Function
    Case OBJ_FEATURE
        calcObjectCost = 0
        Exit Function
    Case Else
        ASSERT_FAILURE "false", "src/multistat.cpp:763", "calcObjectCost", "No such supported object type: %d", CLng(SIMPLE_OBJECT_type(psObj))
    End Select
    calcObjectCost = 0
End Function

Public Sub incrementMultiStatsResearchPerformance(ByVal player As Long)
    If Not ULT(player, MAX_PLAYERS) Then
        Exit Sub
    End If
    playerStats(player).recentResearchPerformance = playerStats(player).recentResearchPerformance + 1
End Sub

Public Sub incrementMultiStatsResearchPotential(ByVal player As Long)
    If Not ULT(player, MAX_PLAYERS) Then
        Exit Sub
    End If
    playerStats(player).recentResearchPotential = playerStats(player).recentResearchPotential + 1
End Sub

' update kills
Public Sub updateMultiStatsKills(ByVal psKilled As Long, ByVal player As Long)
    Dim pwrCost As Double, killedPlayer As Long
    If Not ULT(player, MAX_PLAYERS) Then
        Exit Sub
    End If

    If bMultiPlayer Then
        If psKilled <> 0 Then
            killedPlayer = SIMPLE_OBJECT_player(psKilled)
            If killedPlayer = scavengerSlot() Then
                Exit Sub ' killing scavengers does not count in MP games
            End If
            If killedPlayer < MAX_PLAYERS Then
                pwrCost = U32(calcObjectCost(psKilled))
                playerStats(killedPlayer).recentPowerLost = playerStats(killedPlayer).recentPowerLost + pwrCost
                playerStats(player).recentPowerWon = playerStats(player).recentPowerWon + pwrCost
                If isDroid(psKilled) Then
                    playerStats(killedPlayer).recentDroidPowerLost = playerStats(killedPlayer).recentDroidPowerLost + pwrCost
                    playerStats(killedPlayer).recentDroidsLost = playerStats(killedPlayer).recentDroidsLost + 1
                    playerStats(player).recentDroidsKilled = playerStats(player).recentDroidsKilled + 1
                ElseIf isStructure(psKilled) Then
                    playerStats(killedPlayer).recentStructurePowerLost = playerStats(killedPlayer).recentStructurePowerLost + pwrCost
                    playerStats(killedPlayer).recentStructuresLost = playerStats(killedPlayer).recentStructuresLost + 1
                    playerStats(player).recentStructuresKilled = playerStats(player).recentStructuresKilled + 1
                End If
            End If
            If NetPlay_bComms Then
                playerStats(player).totalKills = playerStats(player).totalKills + 1
            End If
            playerStats(player).recentKills = playerStats(player).recentKills + 1
        End If
        Exit Sub
    End If
    playerStats(player).recentKills = playerStats(player).recentKills + 1
End Sub

Public Sub updateMultiStatsBuilt(ByVal psBuilt As Long)
    Dim pl As Long
    pl = SIMPLE_OBJECT_player(psBuilt)
    If Not ULT(pl, MAX_PLAYERS) Then
        Exit Sub
    End If

    If isDroid(psBuilt) Then
        playerStats(pl).recentDroidsBuilt = playerStats(pl).recentDroidsBuilt + 1
    ElseIf isStructure(psBuilt) Then
        playerStats(pl).recentStructuresBuilt = playerStats(pl).recentStructuresBuilt + 1
    End If
End Sub

Public Sub updateMultiStatsResearchComplete(ByVal psResearch As Long, ByVal player As Long)
    If Not ULT(player, MAX_PLAYERS) Then
        Exit Sub
    End If
    playerStats(player).recentResearchComplete = playerStats(player).recentResearchComplete + 1
End Sub

' (VB) multistat.cpp's per-player recent statistics' getters and setters (game.cpp's save and load)
Public Function getMultiPlayUnitsKilled(ByVal player As Long) As Long
    getMultiPlayUnitsKilled = getMultiStats(player).recentKills
End Function

Public Sub setMultiPlayUnitsKilled(ByVal player As Long, ByVal kills As Long)
    playerStats(player).recentKills = kills
End Sub

Public Function getSelectedPlayerUnitsKilled() As Long
    If ActivityManager_getCurrentGameMode() <> ActivitySink_GameMode_CAMPAIGN Then
        getSelectedPlayerUnitsKilled = getMultiPlayUnitsKilled(selectedPlayer)
    Else
        getSelectedPlayerUnitsKilled = missionData.unitsKilled
    End If
End Function

Public Function getMultiPlayRecentScore(ByVal player As Long) As Long
    getMultiPlayRecentScore = getMultiStats(player).recentScore
End Function

Public Sub setMultiPlayRecentScore(ByVal player As Long, ByVal score As Long)
    playerStats(player).recentScore = score
End Sub

Public Sub setMultiPlayRecentDroidsKilled(ByVal player As Long, ByVal value As Long)
    playerStats(player).recentDroidsKilled = value
End Sub

Public Sub setMultiPlayRecentDroidsLost(ByVal player As Long, ByVal value As Long)
    playerStats(player).recentDroidsLost = value
End Sub

Public Sub setMultiPlayRecentDroidsBuilt(ByVal player As Long, ByVal value As Long)
    playerStats(player).recentDroidsBuilt = value
End Sub

Public Sub setMultiPlayRecentStructuresKilled(ByVal player As Long, ByVal value As Long)
    playerStats(player).recentStructuresKilled = value
End Sub

Public Sub setMultiPlayRecentStructuresLost(ByVal player As Long, ByVal value As Long)
    playerStats(player).recentStructuresLost = value
End Sub

Public Sub setMultiPlayRecentStructuresBuilt(ByVal player As Long, ByVal value As Long)
    playerStats(player).recentStructuresBuilt = value
End Sub

' powerLost: uint64_t (a Double)
Public Sub setMultiPlayRecentPowerLost(ByVal player As Long, ByVal powerLost As Double)
    playerStats(player).recentPowerLost = powerLost
End Sub

Public Sub setMultiPlayRecentDroidPowerLost(ByVal player As Long, ByVal powerLost As Double)
    playerStats(player).recentDroidPowerLost = powerLost
End Sub

Public Sub setMultiPlayRecentStructurePowerLost(ByVal player As Long, ByVal powerLost As Double)
    playerStats(player).recentStructurePowerLost = powerLost
End Sub

Public Sub setMultiPlayRecentPowerWon(ByVal player As Long, ByVal powerWon As Double)
    playerStats(player).recentPowerWon = powerWon
End Sub

Public Sub setMultiPlayRecentResearchComplete(ByVal player As Long, ByVal value As Long)
    playerStats(player).recentResearchComplete = value
End Sub

Public Sub setMultiPlayRecentResearchPotential(ByVal player As Long, ByVal value As Double)
    playerStats(player).recentResearchPotential = value
End Sub

Public Sub setMultiPlayRecentResearchPerformance(ByVal player As Long, ByVal value As Double)
    playerStats(player).recentResearchPerformance = value
End Sub

Public Sub resetRecentScoreData()
    Dim i As Long
    For i = 0 To MAX_CONNECTED_PLAYERS - 1
        playerStats(i).recentKills = 0
        playerStats(i).recentDroidsKilled = 0
        playerStats(i).recentDroidsLost = 0
        playerStats(i).recentDroidsBuilt = 0
        playerStats(i).recentStructuresKilled = 0
        playerStats(i).recentStructuresLost = 0
        playerStats(i).recentStructuresBuilt = 0
        playerStats(i).recentScore = 0
        playerStats(i).recentResearchComplete = 0
        playerStats(i).recentPowerLost = 0
        playerStats(i).recentDroidPowerLost = 0
        playerStats(i).recentStructurePowerLost = 0
        playerStats(i).recentPowerWon = 0
        playerStats(i).recentResearchPerformance = 0
        playerStats(i).recentResearchPotential = 0
        EcKey_clear playerStats(i).identity
        EcKey_clear hostVerifiedJoinIdentities(i)
    Next
End Sub

' MARK: -

Public Sub EcKey_to_json(j As nlohmann_json, k As EcKey)
    Dim publicKeyBytes() As Byte, n As Long, publicKeyB64Str As String
    If EcKey_empty(k) Then
        Set j = json_value("")
        Exit Sub
    End If
    n = EcKey_toBytes(k, EcKey_Public, publicKeyBytes)
    publicKeyB64Str = base64Encode(publicKeyBytes, n)
    Set j = json_value(publicKeyB64Str)
End Sub

Public Sub EcKey_from_json(j As nlohmann_json, k As EcKey)
    Dim publicKeyB64Str As String, publicKeyBytes() As Byte, n As Long
    publicKeyB64Str = j.get_string()
    If Len(publicKeyB64Str) = 0 Then
        EcKey_clear k
        Exit Sub
    End If
    n = base64Decode(publicKeyB64Str, publicKeyBytes)
    EcKey_fromBytes k, publicKeyBytes, n, EcKey_Public
End Sub

Public Sub PLAYERSTATS_to_json(j As nlohmann_json, p As PLAYERSTATS_)
    Dim id As nlohmann_json
    Set j = json_object()
    j.set_item "played", json_value_unsigned(p.played)
    j.set_item "wins", json_value_unsigned(p.wins)
    j.set_item "losses", json_value_unsigned(p.losses)
    j.set_item "totalKills", json_value_unsigned(p.totalKills)
    j.set_item "totalScore", json_value_unsigned(p.totalScore)
    j.set_item "recentKills", json_value_unsigned(p.recentKills)
    j.set_item "recentDroidsKilled", json_value_unsigned(p.recentDroidsKilled)
    j.set_item "recentDroidsLost", json_value_unsigned(p.recentDroidsLost)
    j.set_item "recentDroidsBuilt", json_value_unsigned(p.recentDroidsBuilt)
    j.set_item "recentStructuresKilled", json_value_unsigned(p.recentStructuresKilled)
    j.set_item "recentStructuresLost", json_value_unsigned(p.recentStructuresLost)
    j.set_item "recentStructuresBuilt", json_value_unsigned(p.recentStructuresBuilt)
    j.set_item "recentScore", json_value_unsigned(p.recentScore)
    j.set_item "recentResearchComplete", json_value_unsigned(p.recentResearchComplete)
    j.set_item "recentPowerLost", json_value_unsigned(p.recentPowerLost)
    j.set_item "recentDroidPowerLost", json_value_unsigned(p.recentDroidPowerLost)
    j.set_item "recentStructurePowerLost", json_value_unsigned(p.recentStructurePowerLost)
    j.set_item "recentPowerWon", json_value_unsigned(p.recentPowerWon)
    j.set_item "recentResearchPotential", json_value_unsigned(p.recentResearchPotential)
    j.set_item "recentResearchPerformance", json_value_unsigned(p.recentResearchPerformance)
    EcKey_to_json id, p.identity
    j.set_item "identity", id
End Sub

' template <typename T> optional<T> optGetJSONValue(j, key) for T = uint32_t: hasValue False for nullopt
Private Function optGetJSONValue_uint32(j As nlohmann_json, ByVal key As String, value As Long) As Boolean
    Dim it As nlohmann_json
    Set it = j.find(key)
    If it Is Nothing Then
        optGetJSONValue_uint32 = False
        Exit Function
    End If
    value = it.get_uint()
    optGetJSONValue_uint32 = True
End Function

' optGetJSONValue for T = uint64_t
Private Function optGetJSONValue_uint64(j As nlohmann_json, ByVal key As String, value As Double) As Boolean
    Dim it As nlohmann_json
    Set it = j.find(key)
    If it Is Nothing Then
        optGetJSONValue_uint64 = False
        Exit Function
    End If
    value = it.get_uint64()
    optGetJSONValue_uint64 = True
End Function

' (VB) optGetJSONValue<uint32_t>(j, key).value_or(0)
Private Function optUint32(j As nlohmann_json, ByVal key As String) As Long
    Dim v As Long
    If optGetJSONValue_uint32(j, key, v) Then optUint32 = v Else optUint32 = 0
End Function

' (VB) optGetJSONValue<uint64_t>(j, key).value_or(0)
Private Function optUint64(j As nlohmann_json, ByVal key As String) As Double
    Dim v As Double
    If optGetJSONValue_uint64(j, key, v) Then optUint64 = v Else optUint64 = 0
End Function

Public Sub PLAYERSTATS_from_json(j As nlohmann_json, k As PLAYERSTATS_)
    k.played = j.at("played").get_uint()
    k.wins = j.at("wins").get_uint()
    k.losses = j.at("losses").get_uint()
    k.totalKills = j.at("totalKills").get_uint()
    k.totalScore = j.at("totalScore").get_uint()
    k.recentKills = j.at("recentKills").get_uint()
    k.recentScore = j.at("recentScore").get_uint()
    k.recentPowerLost = j.at("recentPowerLost").get_uint64()
    EcKey_from_json j.at("identity"), k.identity
    ' WZ 4.4.0+:
    k.recentDroidsKilled = optUint32(j, "recentDroidsKilled")
    k.recentDroidsLost = optUint32(j, "recentDroidsLost")
    k.recentDroidsBuilt = optUint32(j, "recentDroidsBuilt")
    k.recentStructuresKilled = optUint32(j, "recentStructuresKilled")
    k.recentStructuresLost = optUint32(j, "recentStructuresLost")
    k.recentStructuresBuilt = optUint32(j, "recentStructuresBuilt")
    k.recentResearchComplete = optUint32(j, "recentResearchComplete")
    k.recentDroidPowerLost = optUint64(j, "recentDroidPowerLost")
    k.recentStructurePowerLost = optUint64(j, "recentStructurePowerLost")
    k.recentPowerWon = optUint64(j, "recentPowerWon")
    k.recentResearchPotential = optUint64(j, "recentResearchPotential")
    k.recentResearchPerformance = optUint64(j, "recentResearchPerformance")
End Sub

' json: an array is made of it (the caller's object becomes the array)
Public Function saveMultiStatsToJSON(json As nlohmann_json) As Boolean
    Dim idx As Long, s As nlohmann_json
    Set json = json_array()

    For idx = 0 To MAX_CONNECTED_PLAYERS - 1
        PLAYERSTATS_to_json s, playerStats(idx)
        json.push_back s
    Next

    saveMultiStatsToJSON = True
End Function

Public Function loadMultiStatsFromJSON(ByVal json As nlohmann_json) As Boolean
    Dim idx As Long, fresh As PLAYERSTATS_, blank As PLAYERSTATS_
    If Not json.is_array() Then
        debug_ 1422, LOG_ERROR, "loadMultiStatsFromJSON", "Expecting an array"
        loadMultiStatsFromJSON = False
        Exit Function
    End If
    If json.size() > MAX_CONNECTED_PLAYERS Then
        debug_ 1427, LOG_ERROR, "loadMultiStatsFromJSON", "Array size is too large: %zu", json.size()
        loadMultiStatsFromJSON = False
        Exit Function
    End If

    For idx = 0 To json.size() - 1
        fresh = blank                       ' get<PLAYERSTATS>(): a value-initialised PLAYERSTATS, filled
        PLAYERSTATS_from_json json.at_idx(idx), fresh
        playerStats(idx) = fresh
    Next

    loadMultiStatsFromJSON = True
End Function

Public Function updateMultiStatsIdentitiesInJSON(ByVal json As nlohmann_json, ByVal useVerifiedJoinIdentity As Boolean) As Boolean
    Dim idx As Long, stats As PLAYERSTATS_, blank As PLAYERSTATS_, s As nlohmann_json
    If Not json.is_array() Then
        debug_ 1443, LOG_ERROR, "updateMultiStatsIdentitiesInJSON", "Expecting an array"
        updateMultiStatsIdentitiesInJSON = False
        Exit Function
    End If
    If json.size() > MAX_CONNECTED_PLAYERS Then
        debug_ 1448, LOG_ERROR, "updateMultiStatsIdentitiesInJSON", "Array size is too large: %zu", json.size()
        updateMultiStatsIdentitiesInJSON = False
        Exit Function
    End If

    For idx = 0 To json.size() - 1
        stats = blank
        PLAYERSTATS_from_json json.at_idx(idx), stats
        If useVerifiedJoinIdentity Then
            stats.identity = getVerifiedJoinIdentity(idx)
        Else
            stats.identity = playerStats(idx).identity
        End If
        PLAYERSTATS_to_json s, stats
        json.set_item_idx idx, s
    Next

    updateMultiStatsIdentitiesInJSON = True
End Function

'--------------------------------------------------------------------------
' class KnownPlayersDB (its tables and state: the declarations above)

' (VB) system_clock::time_point(std::chrono::seconds(secs)).time_since_epoch().count(): secs * 10^7 (100 ns ticks),
' wrapped to int64 as the overflowing multiplication does (a Decimal)
Private Function chronoSecondsToTicks(ByVal secs As Variant) As Variant
    Dim x As Variant, m As Variant
    m = CDec("18446744073709551616")
    x = CDec(secs) * CDec(10000000)
    x = x - Int(x / m) * m
    If x >= CDec("9223372036854775808") Then x = x - m
    chronoSecondsToTicks = x
End Function

' (VB) SQLite's rowid for a new row: the largest + 1
Private Function nextRowId(ids() As Double, ByVal n As Long) As Double
    Dim i As Long, m As Double
    For i = 0 To n - 1
        If ids(i) > m Then m = ids(i)
    Next
    nextRowId = m + 1
End Function

' (VB) the database read from its file (an absent file: empty tables); False when the file is no database
Private Function knownPlayersDB_load() As Boolean
    Dim data() As Byte, size As Long, j As nlohmann_json, t As nlohmann_json, r As nlohmann_json, i As Long
    kpCount = 0: poCount = 0: userVersion = 0
    If PHYSFS_exists(KNOWN_PLAYERS_FILE) = 0 Then
        knownPlayersDB_load = True
        Exit Function
    End If
    If Not loadFile(KNOWN_PLAYERS_FILE, data, size, False) Then Err.Raise vbObjectError + 14, "SQLite", "unable to open database file"
    Set j = json_parse_bytes(data, 0, size)
    If Not j.is_object() Then Err.Raise vbObjectError + 26, "SQLite", "file is not a database"
    userVersion = j.value_int("user_version", 0)
    If j.contains("known_players") Then
        Set t = j.at("known_players")
        For i = 0 To t.size() - 1
            Set r = t.at_idx(i)
            ReDim Preserve kp_id(0 To kpCount): ReDim Preserve kp_name(0 To kpCount): ReDim Preserve kp_pk(0 To kpCount)
            kp_id(kpCount) = r.at("local_id").get_int64()
            kp_name(kpCount) = r.at("name").get_string()
            kp_pk(kpCount) = r.at("pk").get_string()
            kpCount = kpCount + 1
        Next
    End If
    If j.contains("player_options") Then
        Set t = j.at("player_options")
        For i = 0 To t.size() - 1
            Set r = t.at_idx(i)
            ReDim Preserve po_id(0 To poCount): ReDim Preserve po_pk(0 To poCount): ReDim Preserve po_name(0 To poCount)
            ReDim Preserve po_muted(0 To poCount): ReDim Preserve po_banned(0 To poCount)
            po_id(poCount) = r.at("local_id").get_int64()
            po_pk(poCount) = r.at("pk").get_string()
            po_name(poCount) = r.at("name").get_string()
            po_muted(poCount) = CDec(r.at("muted").get_string())
            po_banned(poCount) = CDec(r.at("banned").get_string())
            poCount = poCount + 1
        Next
    End If
    knownPlayersDB_load = True
End Function

' (VB) transaction.commit(): the tables written to the file
Private Sub knownPlayersDB_commit()
    Dim j As nlohmann_json, t As nlohmann_json, r As nlohmann_json, i As Long, s As String, data() As Byte, n As Long
    Set j = json_object()
    j.set_item "user_version", json_value(userVersion)
    Set t = json_array()
    For i = 0 To kpCount - 1
        Set r = json_object()
        r.set_item "local_id", json_value(CDec(kp_id(i)))
        r.set_item "name", json_value(kp_name(i))
        r.set_item "pk", json_value(kp_pk(i))
        t.push_back r
    Next
    j.set_item "known_players", t
    Set t = json_array()
    For i = 0 To poCount - 1
        Set r = json_object()
        r.set_item "local_id", json_value(CDec(po_id(i)))
        r.set_item "pk", json_value(po_pk(i))
        r.set_item "name", json_value(po_name(i))
        r.set_item "muted", json_value(CStr(po_muted(i)))      ' (VB) the int64 as decimal text: exact
        r.set_item "banned", json_value(CStr(po_banned(i)))
        t.push_back r
    Next
    j.set_item "player_options", t
    s = j.dump(1, vbTab)
    n = utf8_encode(s, data)
    If Not saveFile(KNOWN_PLAYERS_FILE, data, n) Then Err.Raise vbObjectError + 10, "SQLite", "disk I/O error"
End Sub

' KnownPlayersDB(const std::string& knownPlayersDBPath) (the caller handles the errors raised)
Private Sub knownPlayersDB_new()
    If Not knownPlayersDB_load() Then Exit Sub
    createKnownPlayersDBTables
End Sub

' optional<PlayerInfo> findPlayerIdentityByName(const std::string& name): the key (pk) when it has a value
Private Function findPlayerIdentityByName(ByVal name_ As String, pk() As Byte, pkLen As Long) As Boolean
    Dim i As Long, has As Boolean, pkB64 As String
    If Len(name_) = 0 Then
        findPlayerIdentityByName = False
        Exit Function
    End If
    cleanupCache
    For i = 0 To fpcCount - 1
        If fpc_name(i) = name_ Then
            If fpc_has(i) Then pkLen = base64Decode(fpc_pk(i), pk)
            findPlayerIdentityByName = fpc_has(i)
            Exit Function
        End If
    Next
    ' SELECT local_id, name, pk FROM known_players WHERE name = ?
    For i = 0 To kpCount - 1
        If kp_name(i) = name_ Then
            has = True
            pkB64 = kp_pk(i)
            pkLen = base64Decode(pkB64, pk)
            Exit For
        End If
    Next
    ' add to the current in-memory cache
    ReDim Preserve fpc_name(0 To fpcCount): ReDim Preserve fpc_has(0 To fpcCount)
    ReDim Preserve fpc_pk(0 To fpcCount): ReDim Preserve fpc_time(0 To fpcCount)
    fpc_name(fpcCount) = name_
    fpc_has(fpcCount) = has
    fpc_pk(fpcCount) = pkB64
    fpc_time(fpcCount) = realTime
    fpcCount = fpcCount + 1
    findPlayerIdentityByName = has
End Function

' (VB) findPlayerCache.erase(name)
Private Sub findPlayerCache_erase(ByVal name_ As String)
    Dim i As Long
    For i = 0 To fpcCount - 1
        If fpc_name(i) = name_ Then
            fpcCount = fpcCount - 1
            fpc_name(i) = fpc_name(fpcCount): fpc_has(i) = fpc_has(fpcCount)
            fpc_pk(i) = fpc_pk(fpcCount): fpc_time(i) = fpc_time(fpcCount)
            Exit Sub
        End If
    Next
End Sub

' Note: May throw on database error!
Private Sub addKnownPlayerIdentity(ByVal name_ As String, key As EcKey, ByVal overrideCurrentKey As Boolean)
    Dim publicKeyb64 As String, bytes() As Byte, n As Long, i As Long, inserted As Boolean, updated As Boolean
    If EcKey_empty(key) Then
        Exit Sub
    End If

    n = EcKey_toBytes(key, EcKey_Public, bytes)
    publicKeyb64 = base64Encode(bytes, n)

    ' Begin transaction
    ' INSERT OR IGNORE INTO known_players(name, pk) VALUES(?, ?) (name is UNIQUE)
    inserted = True
    For i = 0 To kpCount - 1
        If kp_name(i) = name_ Then inserted = False: Exit For
    Next
    If inserted Then
        ReDim Preserve kp_id(0 To kpCount): ReDim Preserve kp_name(0 To kpCount): ReDim Preserve kp_pk(0 To kpCount)
        kp_id(kpCount) = nextRowId(kp_id, kpCount)
        kp_name(kpCount) = name_
        kp_pk(kpCount) = publicKeyb64
        kpCount = kpCount + 1
    ElseIf overrideCurrentKey Then
        ' UPDATE known_players SET pk = ? WHERE name = ?
        For i = 0 To kpCount - 1
            If kp_name(i) = name_ Then kp_pk(i) = publicKeyb64: updated = True
        Next
        If Not updated Then
            debug_ 953, LOG_WARNING, "addKnownPlayerIdentity", "Failed to update known_player (%s)", name_
        End If
    End If

    ' remove from the current in-memory cache
    findPlayerCache_erase name_

    ' Commit transaction
    knownPlayersDB_commit
End Sub

' std::vector<PlayerOptions> findPlayerOptions(const std::string& name, EcKey const &key): the rows' indexes in
' rows(0 .. n - 1), n returned
Private Function findPlayerOptions(ByVal name_ As String, key As EcKey, rows() As Long) As Long
    Dim publicKeyb64 As String, bytes() As Byte, n As Long, i As Long, cnt As Long
    If EcKey_empty(key) Then
        findPlayerOptions = 0
        Exit Function
    End If

    n = EcKey_toBytes(key, EcKey_Public, bytes)
    publicKeyb64 = base64Encode(bytes, n)

    ' SELECT name, muted, banned FROM player_options WHERE pk = ?
    For i = 0 To poCount - 1
        If po_pk(i) = publicKeyb64 Then
            If po_name(i) = name_ Then
                ' if we find an exact match (for both pk *and* name), use it!
                ReDim rows(0 To 0)
                rows(0) = i
                findPlayerOptions = 1
                Exit Function
            End If
            ReDim Preserve rows(0 To cnt)
            rows(cnt) = i
            cnt = cnt + 1
        End If
    Next
    findPlayerOptions = cnt
End Function

' Note: May throw on database error! mutedTimeHas: mutedTime.has_value(); mutedTimeValue: its time_since_epoch().count()
Private Sub knownPlayersDB_setPlayerMuted(ByVal name_ As String, key As EcKey, ByVal mutedTimeHas As Boolean, ByVal mutedTimeValueIn As Variant)
    Dim publicKeyb64 As String, bytes() As Byte, n As Long, i As Long, inserted As Boolean, updated As Boolean
    Dim mutedTimeValue As Variant
    If Not EcKey_empty(key) Then
        n = EcKey_toBytes(key, EcKey_Public, bytes)
        publicKeyb64 = base64Encode(bytes, n)
    End If

    ' Begin transaction

    If mutedTimeHas Then mutedTimeValue = CDec(mutedTimeValueIn) Else mutedTimeValue = CDec(0)

    ' INSERT OR IGNORE INTO player_options(pk, name, muted, banned) VALUES(?, ?, ?, ?) (unique on (pk, name))
    inserted = True
    For i = 0 To poCount - 1
        If po_pk(i) = publicKeyb64 And po_name(i) = name_ Then inserted = False: Exit For
    Next
    If inserted Then
        ReDim Preserve po_id(0 To poCount): ReDim Preserve po_pk(0 To poCount): ReDim Preserve po_name(0 To poCount)
        ReDim Preserve po_muted(0 To poCount): ReDim Preserve po_banned(0 To poCount)
        po_id(poCount) = nextRowId(po_id, poCount)
        po_pk(poCount) = publicKeyb64
        po_name(poCount) = name_
        po_muted(poCount) = mutedTimeValue
        po_banned(poCount) = CDec(-1)    ' default "never set" is -1, not 0
        poCount = poCount + 1
    Else
        ' UPDATE player_options SET muted = ? WHERE pk = ? AND name = ?
        For i = 0 To poCount - 1
            If po_pk(i) = publicKeyb64 And po_name(i) = name_ Then po_muted(i) = mutedTimeValue: updated = True
        Next
        If Not updated Then
            debug_ 1043, LOG_WARNING, "setPlayerMuted", "Failed to update player_options (%s)", name_
        End If
    End If

    ' Commit transaction
    knownPlayersDB_commit
End Sub

Private Sub createKnownPlayersDBTables()
    ' (known_players exists: the file's table, or a new empty one)
    ' database schema upgrades
    Select Case userVersion
    Case 0
        ' used to store player_options that may be associated with a player identity
        ' NOTE: In this case, "name" does *NOT* have a unique constraint
        poCount = 0
        userVersion = 1
        ' fall-through
    End Select
    knownPlayersDB_commit
End Sub

Private Sub cleanupCache()
    Const CACHE_CLEAN_INTERAL As Long = 10 * GAME_TICKS_PER_SEC
    If (realTime - lastCacheClean) > CACHE_CLEAN_INTERAL Then
        fpcCount = 0
        Erase fpc_name: Erase fpc_has: Erase fpc_pk: Erase fpc_time
        lastCacheClean = realTime
    End If
End Sub

Public Sub initKnownPlayers()
    Dim pWriteDir As String, knownPlayersDBPath As String
    If Not knownPlayersDB Then
        pWriteDir = PHYSFS_getWriteDir()
        If Len(pWriteDir) = 0 Then ASSERT_FAILURE "pWriteDir", "src/multistat.cpp:1110", "initKnownPlayers", "PHYSFS_getWriteDir returned null": Exit Sub
        knownPlayersDBPath = pWriteDir & "/" & KNOWN_PLAYERS_FILE
        On Error GoTo failed
        knownPlayersDB_new
        knownPlayersDB = True
    End If
    Exit Sub
failed:
    ' error loading SQLite database
    debug_ 1117, LOG_ERROR, "initKnownPlayers", "Unable to load or initialize SQLite3 database (%s); error: %s", knownPlayersDBPath, Err.Description
    kpCount = 0: poCount = 0: fpcCount = 0
End Sub

Public Sub shutdownKnownPlayers()
    knownPlayersDB = False
    kpCount = 0: poCount = 0: fpcCount = 0
    Erase kp_id: Erase kp_name: Erase kp_pk
    Erase po_id: Erase po_pk: Erase po_name: Erase po_muted: Erase po_banned
    Erase fpc_name: Erase fpc_has: Erase fpc_pk: Erase fpc_time
    lastCacheClean = 0
End Sub

Public Function isLocallyKnownPlayer(ByVal name_ As String, key As EcKey) As Boolean
    Dim pk() As Byte, pkLen As Long, mine() As Byte, n As Long, i As Long
    If Not knownPlayersDB Then ASSERT_FAILURE "knownPlayersDB.operator bool()", "src/multistat.cpp:1130", "isLocallyKnownPlayer", "knownPlayersDB is uninitialized": Exit Function
    If EcKey_empty(key) Then
        isLocallyKnownPlayer = False
        Exit Function
    End If
    If Not findPlayerIdentityByName(name_, pk, pkLen) Then
        isLocallyKnownPlayer = False
        Exit Function
    End If
    ' result.value().pk == key.toBytes(EcKey::Public)
    n = EcKey_toBytes(key, EcKey_Public, mine)
    If pkLen <> n Then Exit Function
    For i = 0 To n - 1
        If pk(i) <> mine(i) Then Exit Function
    Next
    isLocallyKnownPlayer = True
End Function

' void addKnownPlayer(std::string const &name, EcKey const &key, bool override) (VB: the key is
' getMultiStats(identityPlayer).identity, as the callers pass it)
Public Sub addKnownPlayer(ByVal name_ As String, ByVal identityPlayer As Long, Optional ByVal override As Boolean = False)
    Dim key As EcKey
    key = getMultiStats(identityPlayer).identity
    If EcKey_empty(key) Then
        Exit Sub
    End If
    If Not knownPlayersDB Then ASSERT_FAILURE "knownPlayersDB.operator bool()", "src/multistat.cpp:1149", "addKnownPlayer", "knownPlayersDB is uninitialized": Exit Sub

    On Error GoTo failed
    addKnownPlayerIdentity name_, key, override
    Exit Sub
failed:
    debug_ 1155, LOG_ERROR, "addKnownPlayer", "Failed to add known_player with error: %s", Err.Description
End Sub

' optional<StoredPlayerOptions> getStoredPlayerOptions(std::string const &name, EcKey const &key): has_value() is the
' result; its mutedTime: has_value() in mutedTimeHas, time_since_epoch().count() (100 ns ticks) in mutedTimeCount
Public Function getStoredPlayerOptions_mutedTime(ByVal name_ As String, key As EcKey, mutedTimeHas As Boolean, mutedTimeCount As Double) As Boolean
    Dim rows() As Long, n As Long, i As Long, secs As Variant, best As Variant, ticks As Variant
    mutedTimeHas = False
    mutedTimeCount = 0
    If Not knownPlayersDB Then ASSERT_FAILURE "knownPlayersDB.operator bool()", "src/multistat.cpp:1161", "getStoredPlayerOptions", "knownPlayersDB is uninitialized": Exit Function
    n = findPlayerOptions(name_, key, rows)
    If n = 0 Then
        getStoredPlayerOptions_mutedTime = False
        Exit Function
    End If
    ' findPlayerOptions may have returned more than one result, so sort through and find the "newest" for each setting
    ' (std::max with value_or(time_point()): at least 0; bannedTime: nobody reads it)
    best = CDec(0)
    For i = 0 To n - 1
        secs = po_muted(rows(i))         ' std::chrono::seconds(getInt64()): a time point when > 0
        If secs > 0 Then
            ticks = chronoSecondsToTicks(secs)
            If ticks > best Then best = ticks
            mutedTimeHas = True
        End If
    Next
    If mutedTimeHas Then mutedTimeCount = CDbl(best)
    getStoredPlayerOptions_mutedTime = True
End Function

Public Sub storePlayerMuteOption(ByVal name_ As String, key As EcKey, ByVal muted As Boolean)
    If Not knownPlayersDB Then ASSERT_FAILURE "knownPlayersDB.operator bool()", "src/multistat.cpp:1185", "storePlayerMuteOption", "knownPlayersDB is uninitialized": Exit Sub
    ' mutedTime = std::chrono::system_clock::now() when muted: its count, 100 ns ticks since 1970 (MSVC)
    On Error GoTo failed
    knownPlayersDB_setPlayerMuted name_, key, muted, CDec(time_()) * CDec(10000000)
    Exit Sub
failed:
    debug_ 1195, LOG_ERROR, "storePlayerMuteOption", "Failed to store player mute option with error: %s", Err.Description
End Sub
