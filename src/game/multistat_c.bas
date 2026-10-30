Attribute VB_Name = "multistat_c"
' src/multistat.cpp, multistat.h: so far the players' statistics as the
' replays save and restore them (PLAYERSTATS, the type PLAYERSTATS_ since
' the variable playerStats has its name in VB, and their JSON:
' saveMultiStatsToJSON, loadMultiStatsFromJSON,
' updateMultiStatsIdentitiesInJSON), getMultiStats and
' getVerifiedJoinIdentity. Their updates during a game, the files and the
' messages come with the network (stubs_game; loadMultiStats and
' setMultiStats are stand-ins here, as their PLAYERSTATS_ in stubs_game
' would make the two modules depend on each other).
' An EcKey (crc.h: libsodium's ed25519 key) is its public key's bytes: no
' libsodium, so a key read from JSON is kept as given (fromBytes checks its
' length; libsodium's curve point check is not done), and none is ever
' generated, signed with or verified. lib/framework/crc.cpp's EcKey::clear,
' empty, toBytes, fromBytes (the public key) and base64Encode, base64Decode
' are here with it (crc_c is shared by every project).
' C++'s to_json(j, x) is <type>_to_json(j, x) (j becomes a new value) and
' from_json(j, x) is <type>_from_json(j, x).
Option Explicit

Private Const MAX_CONNECTED_PLAYERS As Long = MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS
Private Const crypto_sign_ed25519_PUBLICKEYBYTES As Long = 32

' crc.h: class EcKey (vKey: the public key, or nullptr)
Public Type EcKey
    hasKey As Boolean                   ' vKey != nullptr
    publicKey() As Byte                 ' EC_KEY's publicKey: publicKey(0 .. publicKeySize - 1)
    publicKeySize As Long
End Type

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

' ////////////////////////////////////////////////////////////////////////////
' STATS STUFF
' ////////////////////////////////////////////////////////////////////////////
Public playerStats(0 To MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS - 1) As PLAYERSTATS_  ' static PLAYERSTATS playerStats[MAX_CONNECTED_PLAYERS]

Private hostVerifiedJoinIdentities(0 To MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS - 1) As EcKey

'--------------------------------------------------------------------------
' crc.cpp: EcKey (the public key only) and base64

Public Sub EcKey_clear(k As EcKey)
    k.hasKey = False
    Erase k.publicKey
    k.publicKeySize = 0
End Sub

Public Function EcKey_empty(k As EcKey) As Boolean
    EcKey_empty = Not k.hasKey
End Function

' toBytes(EcKey::Public): the key's bytes into key(0 .. returned size - 1)
Public Function EcKey_toBytes_public(k As EcKey, key() As Byte) As Long
    Dim n As Long
    If EcKey_empty(k) Then
        debug_ 321, LOG_WZ, "toBytes", "No key"
        EcKey_toBytes_public = 0
        Exit Function
    End If
    If k.publicKeySize > 0 Then
        ReDim key(0 To k.publicKeySize - 1)
        For n = 0 To k.publicKeySize - 1
            key(n) = k.publicKey(n)
        Next
    End If
    EcKey_toBytes_public = k.publicKeySize
End Function

' fromBytes(key, EcKey::Public): key(0 .. keySize - 1). (VB) crypto_core_ed25519_is_valid_point (libsodium) is not
' checked: a key of the right length is taken as valid.
Public Function EcKey_fromBytes_public(k As EcKey, key() As Byte, ByVal keySize As Long) As Boolean
    Dim n As Long
    EcKey_clear k

    If keySize <> crypto_sign_ed25519_PUBLICKEYBYTES Then
        debug_ 356, LOG_ERROR, "fromBytes", "Invalid public key"
        EcKey_fromBytes_public = False
        Exit Function
    End If

    ReDim k.publicKey(0 To keySize - 1)
    For n = 0 To keySize - 1
        k.publicKey(n) = key(n)
    Next
    k.publicKeySize = keySize
    k.hasKey = True
    EcKey_fromBytes_public = True
End Function

' crc.cpp: base64Encode(bytes): bytes(0 .. size - 1)
Public Function base64Encode(bytes() As Byte, ByVal size As Long) As String
    Const alphabet As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    Dim str_ As String, n As Long, rem_ As Long, block As Long, i As Long, b1 As Long, b2 As Long
    str_ = String$(((size + 2) \ 3) * 4, vbNullChar)
    n = 0
    Do While n * 3 < size
        rem_ = size - n * 3
        If rem_ > 1 Then b1 = bytes(1 + n * 3) Else b1 = 0
        If rem_ > 2 Then b2 = bytes(2 + n * 3) Else b2 = 0
        block = CLng(bytes(0 + n * 3)) * 65536 Or b1 * 256 Or b2
        For i = 0 To 3
            Mid$(str_, i + n * 4 + 1, 1) = Mid$(alphabet, ((block \ BIT_(6 * (3 - i))) And &H3F) + 1, 1)
        Next
        If rem_ <= 2 Then
            Mid$(str_, 3 + n * 4 + 1, 1) = "="
            If rem_ <= 1 Then
                Mid$(str_, 2 + n * 4 + 1, 1) = "="
            End If
        End If
        n = n + 1
    Loop
    base64Encode = str_
End Function

' crc.cpp: base64Decode(str): the bytes into bytes(0 .. returned size - 1). (VB) str is the std::string's bytes: its
' characters as UTF-8 (base64 text is ASCII)
Public Function base64Decode(ByVal str_ As String, bytes() As Byte) As Long
    Dim s() As Byte, sLen As Long, size As Long, n As Long, i As Long, block As Long, ch As Long, val_ As Long
    sLen = utf8_encode(str_, s)
    size = (sLen \ 4) * 3
    If size > 0 Then ReDim bytes(0 To size - 1)
    For n = 0 To (sLen \ 4) - 1
        block = 0
        For i = 0 To 3
            ch = s(i + n * 4)
            If ch >= 65 And ch <= 90 Then           ' 'A'..'Z'
                val_ = ch - 65 + 0
            ElseIf ch >= 97 And ch <= 122 Then      ' 'a'..'z'
                val_ = ch - 97 + 26
            ElseIf ch >= 48 And ch <= 57 Then       ' '0'..'9'
                val_ = ch - 48 + 52
            ElseIf ch = 43 Then                     ' '+'
                val_ = 62
            ElseIf ch = 47 Then                     ' '/'
                val_ = 63
            Else
                val_ = 0
            End If
            block = block Or val_ * BIT_(6 * (3 - i))
        Next
        bytes(0 + n * 3) = (block \ 65536) And &HFF&
        bytes(1 + n * 3) = (block \ 256) And &HFF&
        bytes(2 + n * 3) = block And &HFF&
    Next
    If sLen >= 4 Then
        If s((sLen \ 4) * 4 - 2) = 61 Then            ' '='
            size = size - 2
        ElseIf s((sLen \ 4) * 4 - 1) = 61 Then
            size = size - 1
        End If
    End If
    base64Decode = size
End Function

' multistat.cpp (stand-in: the player's .sta and .json stats files, file output, later; here for its PLAYERSTATS_):
' bool saveMultiStats(const char *sFileName, const char *sPlayerName, const PLAYERSTATS *st)
Public Function saveMultiStats(ByVal sFileName As String, ByVal sPlayerName As String, st As PLAYERSTATS_) As Boolean
    saveMultiStats = True
End Function

' multistat.h: bool loadMultiStats(char *sPlayerName, PLAYERSTATS *playerStats) (stand-in: the players' stats files, cluster J)
Public Function loadMultiStats(ByVal sPlayerName As String, st As PLAYERSTATS_) As Boolean
    loadMultiStats = True
End Function

' multistat.h: bool setMultiStats(uint32_t player, PLAYERSTATS plStats, bool bLocal) (stand-in: set + send to net, cluster J)
Public Function setMultiStats(ByVal player As Long, plStats As PLAYERSTATS_, ByVal bLocal As Boolean) As Boolean
    setMultiStats = True
End Function

' ////////////////////////////////////////////////////////////////////////////
' Get Player's stats
Public Function getMultiStats(ByVal player As Long) As PLAYERSTATS_
    getMultiStats = playerStats(player)
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

' MARK: -

Public Sub EcKey_to_json(j As nlohmann_json, k As EcKey)
    Dim publicKeyBytes() As Byte, n As Long, publicKeyB64Str As String
    If EcKey_empty(k) Then
        Set j = json_value("")
        Exit Sub
    End If
    n = EcKey_toBytes_public(k, publicKeyBytes)
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
    EcKey_fromBytes_public k, publicKeyBytes, n
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

' (VB) multistat.cpp's per-player recent statistics' getters and setters (game.cpp's save and load)
Public Function getMultiPlayUnitsKilled(ByVal player As Long) As Long
    getMultiPlayUnitsKilled = getMultiStats(player).recentKills
End Function

Public Sub setMultiPlayUnitsKilled(ByVal player As Long, ByVal kills As Long)
    playerStats(player).recentKills = kills
End Sub

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
