Attribute VB_Name = "multivote_c"
' src/multivote.cpp, multivote.h: the lobby's votes (the host's map change or randomisation), the kick votes and
' the players' preferences for the lobby options (NET_VOTE, NET_VOTE_REQUEST; the .multiprefs files).
' (VB) PlayerPreferences::BuiltinPreferences' std::set<ScavType / AllianceType / PowerSetting / CampType /
' TechLevel>s are bit masks (bit v set: v in the set; iterating the bits upwards is std::set's order), one per
' player in the pref arrays; std::vector<PendingVoteKick> pendingKickVotes is the pkv arrays (pkvCount of them).
' The notifications' lambdas are WzFunc ids (multivote_c_fn).
Option Explicit
Private Const MAX_CONNECTED_PLAYERS As Long = MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS

' multiplaydefs.h's enum maxima (as multiint_lobby_c's)
Private Const SCAV_TYPE_MAX As Long = ULTIMATE_SCAVENGERS
Private Const ALLIANCE_TYPE_MAX As Long = ALLIANCES_UNSHARED
Private Const POWER_SETTING_MAX As Long = LEV_HI
Private Const CAMP_TYPE_MAX As Long = CAMP_WALLS
Private Const TECH_LEVEL_MIN As Long = 1                    ' TechLevel::TECH_1
Private Const TECH_LEVEL_MAX As Long = 4                    ' TechLevel::TECH_4

Public Const FN_lobbyChangeVote_allow As Long = 17000       ' [](const WZ_Notification&) { sendLobbyChangeVoteData(1); }
Public Const FN_kickVote_yes As Long = 17001                ' [voteID](const WZ_Notification&) { sendPlayerKickedVote(voteID, 1); }
Public Const FN_kickVote_onDismissed As Long = 17002        ' [voteID](const WZ_Notification&, WZ_Notification_Dismissal_Reason reason)

' MARK: -

Private Const VOTE_TAG As String = "voting"
Private Const VOTE_KICK_TAG_PREFIX As String = "votekick::"

Private playerVotes(0 To MAX_PLAYERS - 1) As Long           ' uint8_t

' struct PlayerPreferences::BuiltinPreferences (the sets as bit masks)
Private Type BuiltinPreferences
    scavengers As Long
    alliances As Long
    power As Long
    base As Long
    techLevel As Long
    minPlayers As Long                                      ' uint8_t
    maxPlayers As Long                                      ' uint8_t
End Type

' static std::array<PlayerPreferences, MAX_CONNECTED_PLAYERS> playerPreferences
Private playerPreferences(0 To MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS - 1) As BuiltinPreferences
Private playerPreferencesReady As Boolean                   ' (VB) the constructors have run

' enum class NetVoteType
Private Const NetVoteType_LOBBY_SETTING_CHANGE As Long = 0
Private Const NetVoteType_KICK_PLAYER As Long = 1
Private Const NetVoteType_LOBBY_OPTION_PREFERENCES_BUILTIN As Long = 2

' struct PendingVoteKick: pendingKickVotes[i] is pkv...(i)
Private pkvCount As Long
Private pkvUniqueVoteId() As Long
Private pkvRequesterPlayerId() As Long
Private pkvTargetPlayerId() As Long
Private pkvHasVote() As Boolean                             ' votes[p].has_value(): (i, p)
Private pkvVote() As Boolean                                ' votes[p].value()
Private pkvStartTime() As Long

Private Const PENDING_KICK_VOTE_TIMEOUT_MS As Long = 20000
Private Const MIN_INTERVAL_BETWEEN_PLAYER_KICK_VOTES_MS As Long = 60000
Private has_lastKickVoteForEachPlayer(0 To MAX_PLAYERS - 1) As Boolean
Private lastKickVoteForEachPlayer(0 To MAX_PLAYERS - 1) As Long

Private last_vote_id As Long                                ' startKickVote's static uint32_t last_vote_id = 0

Private Const MaxSupportedSetMembers As Long = 255          ' std::numeric_limits<uint8_t>::max()

' (VB) bit v of a set mask
Private Function setHas(ByVal mask As Long, ByVal v As Long) As Boolean
    setHas = (mask And BIT_(v)) <> 0
End Function

' MARK: - PlayerPreferences

' PlayerPreferences::BuiltinPreferences::BuiltinPreferences() and reset()
Private Sub BuiltinPreferences_reset(p As BuiltinPreferences)
    p.scavengers = 0
    p.alliances = 0
    p.power = 0
    p.base = 0
    p.techLevel = 0
    p.minPlayers = 2
    p.maxPlayers = MAX_PLAYERS_IN_GUI
End Sub

' (VB) the array's constructors
Private Sub ensurePlayerPreferences()
    Dim i As Long
    If playerPreferencesReady Then Exit Sub
    For i = 0 To MAX_CONNECTED_PLAYERS - 1
        BuiltinPreferences_reset playerPreferences(i)
    Next
    playerPreferencesReady = True
End Sub

Private Function BuiltinPreferences_equals(a As BuiltinPreferences, b As BuiltinPreferences) As Boolean
    BuiltinPreferences_equals = a.scavengers = b.scavengers And a.alliances = b.alliances And a.power = b.power And _
        a.base = b.base And a.techLevel = b.techLevel And a.minPlayers = b.minPlayers And a.maxPlayers = b.maxPlayers
End Function

Private Function validateNumToScavTypeEnum(ByVal val_ As Long) As Boolean
    validateNumToScavTypeEnum = val_ <= SCAV_TYPE_MAX
End Function

Private Function validateNumToAlliancesTypeEnum(ByVal val_ As Long) As Boolean
    validateNumToAlliancesTypeEnum = val_ <= ALLIANCE_TYPE_MAX
End Function

Private Function validateNumToPowerSettingEnum(ByVal val_ As Long) As Boolean
    validateNumToPowerSettingEnum = val_ <= POWER_SETTING_MAX
End Function

Private Function validateNumToCampTypeEnum(ByVal val_ As Long) As Boolean
    validateNumToCampTypeEnum = val_ <= CAMP_TYPE_MAX
End Function

Private Function validateNumToTechLevelEnum(ByVal val_ As Long) As Boolean
    validateNumToTechLevelEnum = val_ >= TECH_LEVEL_MIN And val_ <= TECH_LEVEL_MAX
End Function

' (VB) the validate function of a set: 0 scavengers, 1 alliances, 2 power, 3 base, 4 techLevel
Private Function validateNumToEnum(ByVal which As Long, ByVal val_ As Long) As Boolean
    Select Case which
    Case 0: validateNumToEnum = validateNumToScavTypeEnum(val_)
    Case 1: validateNumToEnum = validateNumToAlliancesTypeEnum(val_)
    Case 2: validateNumToEnum = validateNumToPowerSettingEnum(val_)
    Case 3: validateNumToEnum = validateNumToCampTypeEnum(val_)
    Case 4: validateNumToEnum = validateNumToTechLevelEnum(val_)
    End Select
End Function

Private Function validateMinMaxPlayersPrefValue(ByVal val_ As Long) As Boolean
    validateMinMaxPlayersPrefValue = val_ >= 2 And val_ <= MAX_PLAYERS_IN_GUI
End Function

Private Function clampMinMaxPlayersPrefValue(ByVal val_ As Long) As Long
    If val_ < 2 Then val_ = 2
    If val_ > MAX_PLAYERS_IN_GUI Then val_ = MAX_PLAYERS_IN_GUI
    clampMinMaxPlayersPrefValue = val_
End Function

' (VB) a std::set<enum> as a JSON array of its numbers
Private Function setToJSON(ByVal mask As Long) As nlohmann_json
    Dim a As nlohmann_json, v As Long
    Set a = json_array()
    For v = 0 To 30
        If setHas(mask, v) Then a.push_back json_value_unsigned(v)
    Next
    Set setToJSON = a
End Function

' inline void to_json(nlohmann::json& j, const PlayerPreferences::BuiltinPreferences& builtinPreferences)
Private Sub BuiltinPreferences_to_json(j As nlohmann_json, p As BuiltinPreferences)
    Dim playersObj As nlohmann_json
    Set j = json_object()
    If p.scavengers <> 0 Then
        j.set_item "scavs", setToJSON(p.scavengers)
    End If
    If p.alliances <> 0 Then
        j.set_item "allies", setToJSON(p.alliances)
    End If
    If p.power <> 0 Then
        j.set_item "power", setToJSON(p.power)
    End If
    If p.base <> 0 Then
        j.set_item "base", setToJSON(p.base)
    End If
    If p.techLevel <> 0 Then
        j.set_item "tech", setToJSON(p.techLevel)
    End If
    Set playersObj = json_object()
    playersObj.set_item "min", json_value_unsigned(p.minPlayers)
    playersObj.set_item "max", json_value_unsigned(p.maxPlayers)
    j.set_item "players", playersObj
End Sub

' setUint8ToSetEnumT<T>(it->get<std::set<uint8_t>>(), validate...): get<std::set<uint8_t>> throws (a VB error) for a non-array
Private Function setUint8ToSetEnumT(ByVal arr As nlohmann_json, ByVal which As Long) As Long
    Dim i As Long, el As Long, result As Long
    If Not arr.is_array() Then Err.Raise vbObjectError + 302, "nlohmann::json", "[json.exception.type_error.302] type must be array, but is " & arr.type_name()   ' (json_hpp's json_throw)
    For i = 0 To arr.size() - 1
        el = arr.at_idx(i).get_uint() And &HFF&         ' get<uint8_t>
        If validateNumToEnum(which, el) Then
            result = result Or BIT_(el)
        End If
    Next
    setUint8ToSetEnumT = result
End Function

' inline void from_json(const nlohmann::json& j, PlayerPreferences::BuiltinPreferences& builtinPreferences)
Private Sub BuiltinPreferences_from_json(ByVal j As nlohmann_json, p As BuiltinPreferences)
    Dim it As nlohmann_json, playersObj As nlohmann_json, players_it As nlohmann_json
    Set it = j.find("scavs")
    If Not it Is Nothing Then
        p.scavengers = setUint8ToSetEnumT(it, 0)
    End If
    Set it = j.find("allies")
    If Not it Is Nothing Then
        p.alliances = setUint8ToSetEnumT(it, 1)
    End If
    Set it = j.find("power")
    If Not it Is Nothing Then
        p.power = setUint8ToSetEnumT(it, 2)
    End If
    Set it = j.find("base")
    If Not it Is Nothing Then
        p.base = setUint8ToSetEnumT(it, 3)
    End If
    Set it = j.find("tech")
    If Not it Is Nothing Then
        p.techLevel = setUint8ToSetEnumT(it, 4)
    End If
    Set it = j.find("players")
    If Not it Is Nothing Then
        Set playersObj = it
        Set players_it = playersObj.find("min")
        If Not players_it Is Nothing Then
            p.minPlayers = clampMinMaxPlayersPrefValue(players_it.get_uint() And &HFF&)
        End If
        Set players_it = playersObj.find("max")
        If Not players_it Is Nothing Then
            p.maxPlayers = clampMinMaxPlayersPrefValue(players_it.get_uint() And &HFF&)
        End If
    End If
End Sub

' nlohmann::json PlayerPreferences::prefValuesToJSON()
Private Function prefValuesToJSON(ByVal playerIndex As Long) As nlohmann_json
    Dim result As nlohmann_json
    BuiltinPreferences_to_json result, playerPreferences(playerIndex)
    Set prefValuesToJSON = result
End Function

' bool PlayerPreferences::setPrefValuesFromJSON(const nlohmann::json& j): j.get<BuiltinPreferences>() starts from a new one
Private Function setPrefValuesFromJSON(ByVal playerIndex As Long, ByVal j As nlohmann_json) As Boolean
    Dim fresh As BuiltinPreferences
    On Error GoTo threw
    BuiltinPreferences_reset fresh
    BuiltinPreferences_from_json j, fresh
    playerPreferences(playerIndex) = fresh
    setPrefValuesFromJSON = True
    Exit Function
threw:
    debug_ 336, LOG_ERROR, "setPrefValuesFromJSON", "Failed to load pref json with error: %s", Err.Description
    setPrefValuesFromJSON = False
End Function

' bool PlayerPreferences::setBuiltinPreferences(BuiltinPreferences&& newPrefs)
Private Function setBuiltinPreferences(ByVal playerIndex As Long, newPrefs As BuiltinPreferences) As Boolean
    If BuiltinPreferences_equals(playerPreferences(playerIndex), newPrefs) Then
        ' no change
        setBuiltinPreferences = False
        Exit Function
    End If
    playerPreferences(playerIndex) = newPrefs
    setBuiltinPreferences = True
End Function

' (VB) PlayerPreferences::getPrefValue / setPrefValue of one of the sets (which: 0 scavengers .. 4 techLevel)
Private Function getPrefValue(ByVal playerIndex As Long, ByVal which As Long, ByVal val_ As Long) As Boolean
    Dim mask As Long
    Select Case which
    Case 0: mask = playerPreferences(playerIndex).scavengers
    Case 1: mask = playerPreferences(playerIndex).alliances
    Case 2: mask = playerPreferences(playerIndex).power
    Case 3: mask = playerPreferences(playerIndex).base
    Case 4: mask = playerPreferences(playerIndex).techLevel
    End Select
    getPrefValue = setHas(mask, val_)
End Function

' returns "true" if preference value changed, false if unchanged (insert(val).second / erase(val) != 0)
Private Function setPrefValue(ByVal playerIndex As Long, ByVal which As Long, ByVal val_ As Long, ByVal pref As Boolean) As Boolean
    Dim mask As Long, newMask As Long
    Select Case which
    Case 0: mask = playerPreferences(playerIndex).scavengers
    Case 1: mask = playerPreferences(playerIndex).alliances
    Case 2: mask = playerPreferences(playerIndex).power
    Case 3: mask = playerPreferences(playerIndex).base
    Case 4: mask = playerPreferences(playerIndex).techLevel
    End Select
    If pref Then newMask = mask Or BIT_(val_) Else newMask = mask And Not BIT_(val_)
    Select Case which
    Case 0: playerPreferences(playerIndex).scavengers = newMask
    Case 1: playerPreferences(playerIndex).alliances = newMask
    Case 2: playerPreferences(playerIndex).power = newMask
    Case 3: playerPreferences(playerIndex).base = newMask
    Case 4: playerPreferences(playerIndex).techLevel = newMask
    End Select
    setPrefValue = newMask <> mask
End Function

' void PlayerPreferences::setAllBuiltin(bool pref) (setAllPrefValues<T>(maxVal, pref, minVal))
Private Sub setAllBuiltin(ByVal playerIndex As Long, ByVal pref As Boolean)
    Dim i As Long
    For i = 0 To SCAV_TYPE_MAX: setPrefValue playerIndex, 0, i, pref: Next
    For i = 0 To ALLIANCE_TYPE_MAX: setPrefValue playerIndex, 1, i, pref: Next
    For i = 0 To POWER_SETTING_MAX: setPrefValue playerIndex, 2, i, pref: Next
    For i = 0 To CAMP_TYPE_MAX: setPrefValue playerIndex, 3, i, pref: Next
    For i = TECH_LEVEL_MIN To TECH_LEVEL_MAX: setPrefValue playerIndex, 4, i, pref: Next
End Sub

Public Sub loadMultiOptionPrefValues(ByVal sPlayerName As String, ByVal playerIndex As Long)
    Dim fileName As String, result As nlohmann_json, it As nlohmann_json
    ensurePlayerPreferences
    If Not ULT(playerIndex, MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "playerIndex < playerPreferences.size()", "src/multivote.cpp:434", "loadMultiOptionPrefValues", "Invalid player idx: %u", playerIndex: Exit Sub
    BuiltinPreferences_reset playerPreferences(playerIndex)

    ' Stored alongside the profile, with the same filename but different extension
    fileName = MultiPlayersPath & sPlayerName & ".multiprefs"
    Set result = wzLoadJsonObjectFromFile(fileName, True)
    If result Is Nothing Then
        Exit Sub
    End If

    ' Load last used prefs set
    Set it = result.find("last")
    If Not it Is Nothing Then
        setPrefValuesFromJSON playerIndex, it
    End If
End Sub

Public Function saveMultiOptionPrefValues(ByVal sPlayerName As String, ByVal playerIndex As Long) As Boolean
    Dim fileName As String, rootObj As nlohmann_json, result As nlohmann_json, dumpedJSON As String, data() As Byte, n As Long
    ensurePlayerPreferences
    If Not ULT(playerIndex, MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "playerIndex < playerPreferences.size()", "src/multivote.cpp:456", "saveMultiOptionPrefValues", "Invalid player idx: %u", playerIndex: saveMultiOptionPrefValues = False: Exit Function

    ' Stored alongside the profile, with the same filename but different extension
    fileName = MultiPlayersPath & sPlayerName & ".multiprefs"

    Set rootObj = json_object()

    ' Load existing file (to preserve any future saved pref sets)
    Set result = wzLoadJsonObjectFromFile(fileName, True)
    If Not result Is Nothing Then
        Set rootObj = result
    End If

    rootObj.set_item "last", prefValuesToJSON(playerIndex)

    dumpedJSON = rootObj.dump(4)
    n = utf8_encode(dumpedJSON, data)
    saveMultiOptionPrefValues = saveFile(fileName, data, n)
End Function

Public Function getMultiOptionPrefValuesJSON(ByVal playerIndex As Long) As nlohmann_json
    ensurePlayerPreferences
    If Not ULT(playerIndex, MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "playerIndex < playerPreferences.size()", "src/multivote.cpp:476", "getMultiOptionPrefValuesJSON", "Invalid player idx: %u", playerIndex: Set getMultiOptionPrefValuesJSON = json_parse("null"): Exit Function
    Set getMultiOptionPrefValuesJSON = prefValuesToJSON(playerIndex)
End Function

Public Sub resetMultiOptionPrefValues(ByVal player As Long)
    ensurePlayerPreferences
    If Not ULT(player, MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "player < playerPreferences.size()", "src/multivote.cpp:482", "resetMultiOptionPrefValues", "Invalid player idx: %u", player: Exit Sub
    BuiltinPreferences_reset playerPreferences(player)
    If NetPlay_bComms And ingame.side = InGameSide_MULTIPLAYER_CLIENT And Not NetPlay_isHost And player = selectedPlayer Then
        ' Notify host of change
        sendPlayerMultiOptPreferencesBuiltin_player selectedPlayer
    End If
End Sub

Public Sub resetAllMultiOptionPrefValues()
    Dim i As Long
    ensurePlayerPreferences
    For i = 0 To MAX_CONNECTED_PLAYERS - 1
        BuiltinPreferences_reset playerPreferences(i)
    Next
End Sub

Public Sub multiOptionPrefValuesSwap(ByVal playerIndexA As Long, ByVal playerIndexB As Long)
    Dim t As BuiltinPreferences
    If Not NetPlay_isHost Then Exit Sub
    ensurePlayerPreferences
    t = playerPreferences(playerIndexA)
    playerPreferences(playerIndexA) = playerPreferences(playerIndexB)
    playerPreferences(playerIndexB) = t
End Sub

Public Sub sendPlayerMultiOptPreferencesBuiltin()
    sendPlayerMultiOptPreferencesBuiltin_player selectedPlayer
End Sub

' template <typename T> void setMultiOptionPrefValueT(T val, bool pref)
Private Sub setMultiOptionPrefValueT(ByVal which As Long, ByVal val_ As Long, ByVal pref As Boolean)
    ensurePlayerPreferences
    If setPrefValue(selectedPlayer, which, val_, pref) Then
        If NetPlay_bComms And ingame.side = InGameSide_MULTIPLAYER_CLIENT And Not NetPlay_isHost Then
            ' Notify host of change
            sendPlayerMultiOptPreferencesBuiltin_player selectedPlayer
        End If
    End If
End Sub

Public Sub setMultiOptionPrefValue_ScavType(ByVal val_ As Long, ByVal pref As Boolean)
    setMultiOptionPrefValueT 0, val_, pref
End Sub

Public Sub setMultiOptionPrefValue_AllianceType(ByVal val_ As Long, ByVal pref As Boolean)
    setMultiOptionPrefValueT 1, val_, pref
End Sub

Public Sub setMultiOptionPrefValue_PowerSetting(ByVal val_ As Long, ByVal pref As Boolean)
    setMultiOptionPrefValueT 2, val_, pref
End Sub

Public Sub setMultiOptionPrefValue_CampType(ByVal val_ As Long, ByVal pref As Boolean)
    setMultiOptionPrefValueT 3, val_, pref
End Sub

Public Sub setMultiOptionPrefValue_TechLevel(ByVal val_ As Long, ByVal pref As Boolean)
    setMultiOptionPrefValueT 4, val_, pref
End Sub

Public Sub setMultiOptionBuiltinPrefValues(ByVal pref As Boolean)
    ensurePlayerPreferences
    setAllBuiltin selectedPlayer, pref
    If NetPlay_bComms And ingame.side = InGameSide_MULTIPLAYER_CLIENT And Not NetPlay_isHost Then
        ' Notify host of change
        sendPlayerMultiOptPreferencesBuiltin_player selectedPlayer
    End If
End Sub

Public Function getMultiOptionPrefValue_ScavType(ByVal val_ As Long) As Boolean
    ensurePlayerPreferences
    getMultiOptionPrefValue_ScavType = getPrefValue(selectedPlayer, 0, val_)
End Function

Public Function getMultiOptionPrefValue_AllianceType(ByVal val_ As Long) As Boolean
    ensurePlayerPreferences
    getMultiOptionPrefValue_AllianceType = getPrefValue(selectedPlayer, 1, val_)
End Function

Public Function getMultiOptionPrefValue_PowerSetting(ByVal val_ As Long) As Boolean
    ensurePlayerPreferences
    getMultiOptionPrefValue_PowerSetting = getPrefValue(selectedPlayer, 2, val_)
End Function

Public Function getMultiOptionPrefValue_CampType(ByVal val_ As Long) As Boolean
    ensurePlayerPreferences
    getMultiOptionPrefValue_CampType = getPrefValue(selectedPlayer, 3, val_)
End Function

Public Function getMultiOptionPrefValue_TechLevel(ByVal val_ As Long) As Boolean
    ensurePlayerPreferences
    getMultiOptionPrefValue_TechLevel = getPrefValue(selectedPlayer, 4, val_)
End Function

' template <typename T> size_t getMultiOptionPrefValueTotalT(T val, bool playersOnly)
Private Function getMultiOptionPrefValueTotalT(ByVal which As Long, ByVal val_ As Long, ByVal playersOnly As Boolean) As Long
    Dim idx As Long, result As Long
    ensurePlayerPreferences
    For idx = 0 To MAX_CONNECTED_PLAYERS - 1
        If playersOnly And idx >= MAX_PLAYERS Then Exit For
        If getPrefValue(idx, which, val_) Then result = result + 1
    Next
    getMultiOptionPrefValueTotalT = result
End Function

Public Function getMultiOptionPrefValueTotal_ScavType(ByVal val_ As Long, Optional ByVal playersOnly As Boolean = True) As Long
    getMultiOptionPrefValueTotal_ScavType = getMultiOptionPrefValueTotalT(0, val_, playersOnly)
End Function

Public Function getMultiOptionPrefValueTotal_AllianceType(ByVal val_ As Long, Optional ByVal playersOnly As Boolean = True) As Long
    getMultiOptionPrefValueTotal_AllianceType = getMultiOptionPrefValueTotalT(1, val_, playersOnly)
End Function

Public Function getMultiOptionPrefValueTotal_PowerSetting(ByVal val_ As Long, Optional ByVal playersOnly As Boolean = True) As Long
    getMultiOptionPrefValueTotal_PowerSetting = getMultiOptionPrefValueTotalT(2, val_, playersOnly)
End Function

Public Function getMultiOptionPrefValueTotal_CampType(ByVal val_ As Long, Optional ByVal playersOnly As Boolean = True) As Long
    getMultiOptionPrefValueTotal_CampType = getMultiOptionPrefValueTotalT(3, val_, playersOnly)
End Function

Public Function getMultiOptionPrefValueTotal_TechLevel(ByVal val_ As Long, Optional ByVal playersOnly As Boolean = True) As Long
    getMultiOptionPrefValueTotal_TechLevel = getMultiOptionPrefValueTotalT(4, val_, playersOnly)
End Function


' template<typename T> bool NETEnumTSet_uint8(MessageReader& r, std::set<T>& val, validateValueFunc)
Private Function NETEnumTSet_uint8_r(ByVal r As MessageReader, val_ As Long, ByVal which As Long) As Boolean
    Dim retVal As Boolean, numElements As Long, i As Long, el As Long
    retVal = True
    numElements = 0
    val_ = 0
    NETuint32_t r, numElements
    If ULT(MaxSupportedSetMembers, numElements) Then
        debug_ 646, LOG_NET, "NETEnumTSet_uint8", "Invalid number of set members: %u", numElements
        ' will skip extras in the loop, set return value to false
        retVal = False
    End If
    i = 0
    Do While ULT(i, numElements)
        el = 0
        NETuint8_t r, el
        If ULT(i, numElements) Then
            If Not validateNumToEnum(which, el) Then
                retVal = False
            Else
                val_ = val_ Or BIT_(el)
            End If
        End If
        i = i + 1
    Loop
    NETEnumTSet_uint8_r = retVal
End Function

' template<typename T> bool NETEnumTSet_uint8(MessageWriter& w, std::set<T>& val, validateValueFunc)
Private Function NETEnumTSet_uint8_w(ByVal w As MessageWriter, ByVal val_ As Long) As Boolean
    Dim numElements As Long, i As Long, el As Long
    For el = 0 To 30
        If setHas(val_, el) Then numElements = numElements + 1
    Next
    If numElements > MaxSupportedSetMembers Then
        ASSERT_FAILURE "false", "src/multivote.cpp:688", "NETEnumTSet_uint8", "Invalid number of set members: %u", numElements
        numElements = MaxSupportedSetMembers
    End If
    NETuint32_t w, numElements
    i = 0
    For el = 0 To 30
        If setHas(val_, el) Then
            If i >= numElements Then Exit For
            NETuint8_t w, el
            i = i + 1
        End If
    Next
    NETEnumTSet_uint8_w = True
End Function

' template <typename SerdeContext> bool NETbuiltinPlayerPreferences(SerdeContext& c, BuiltinPreferences& builtinPrefs)
Private Function NETbuiltinPlayerPreferences(ByVal c As Object, builtinPrefs As BuiltinPreferences) As Boolean
    Dim retSuccess As Boolean, r As MessageReader, w As MessageWriter
    retSuccess = True
    If TypeOf c Is MessageReader Then
        Set r = c
        retSuccess = NETEnumTSet_uint8_r(r, builtinPrefs.scavengers, 0) And retSuccess
        retSuccess = NETEnumTSet_uint8_r(r, builtinPrefs.alliances, 1) And retSuccess
        retSuccess = NETEnumTSet_uint8_r(r, builtinPrefs.power, 2) And retSuccess
        retSuccess = NETEnumTSet_uint8_r(r, builtinPrefs.base, 3) And retSuccess
        retSuccess = NETEnumTSet_uint8_r(r, builtinPrefs.techLevel, 4) And retSuccess
    Else
        Set w = c
        retSuccess = NETEnumTSet_uint8_w(w, builtinPrefs.scavengers) And retSuccess
        retSuccess = NETEnumTSet_uint8_w(w, builtinPrefs.alliances) And retSuccess
        retSuccess = NETEnumTSet_uint8_w(w, builtinPrefs.power) And retSuccess
        retSuccess = NETEnumTSet_uint8_w(w, builtinPrefs.base) And retSuccess
        retSuccess = NETEnumTSet_uint8_w(w, builtinPrefs.techLevel) And retSuccess
    End If
    NETuint8_t c, builtinPrefs.minPlayers
    NETuint8_t c, builtinPrefs.maxPlayers
    If Not validateMinMaxPlayersPrefValue(builtinPrefs.minPlayers) Then
        builtinPrefs.minPlayers = 2
        retSuccess = False
    End If
    If Not validateMinMaxPlayersPrefValue(builtinPrefs.maxPlayers) Then
        builtinPrefs.maxPlayers = MAX_PLAYERS_IN_GUI
        retSuccess = False
    End If
    NETbuiltinPlayerPreferences = retSuccess
End Function

' MARK: -

Public Sub resetKickVoteData()
    Dim i As Long
    pkvCount = 0                    ' pendingKickVotes.clear();
    For i = 0 To MAX_PLAYERS - 1
        has_lastKickVoteForEachPlayer(i) = False
    Next
End Sub

Public Sub resetLobbyChangeVoteData()
    Dim i As Long
    For i = 0 To MAX_PLAYERS - 1
        playerVotes(i) = 0
    Next
End Sub

Public Sub resetLobbyChangePlayerVote(ByVal player As Long)
    If Not ULT(player, MAX_PLAYERS) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/multivote.cpp:756", "resetLobbyChangePlayerVote", "Invalid player: %u", player: Exit Sub
    playerVotes(player) = 0
End Sub

Public Sub sendLobbyChangeVoteData(ByVal currentVote As Long)
    Dim w As MessageWriter, voteType As Long
    Set w = NETbeginEncode(NETbroadcastQueue(), NET_VOTE)
    NETuint32_t w, selectedPlayer
    voteType = NetVoteType_LOBBY_SETTING_CHANGE
    NETuint8_t w, voteType
    NETuint8_t w, currentVote
    NETend w
End Sub

Public Function getLobbyChangeVoteTotal() As Long
    Dim total As Long, i As Long
    If Not NetPlay_isHost Then
        ASSERT_FAILURE "false", "src/multivote.cpp:772", "getLobbyChangeVoteTotal", "Host only routine detected for client!"
        getLobbyChangeVoteTotal = 1         ' return true (uint8_t)
        Exit Function
    End If

    total = 0
    For i = 0 To MAX_PLAYERS - 1
        If isHumanPlayer(i) Then
            If selectedPlayer = i Then
                ' always count the host as a "yes" vote.
                playerVotes(i) = 1
            End If
            total = (total + playerVotes(i)) And &HFF&
        Else
            playerVotes(i) = 0
        End If
    Next
    getLobbyChangeVoteTotal = total
End Function

Private Sub recvLobbyChangeVote(ByVal player As Long, ByVal newVote As Long)
    Dim msg As String
    If Not ULT(player, MAX_PLAYERS) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/multivote.cpp:798", "recvLobbyChangeVote", "Invalid sender: %u", player: Exit Sub

    If newVote = 1 Then playerVotes(player) = 1 Else playerVotes(player) = 0

    debug_ 802, LOG_NET, "recvLobbyChangeVote", "total votes: %d/%d", getLobbyChangeVoteTotal(), NET_numHumanPlayers()

    ' there is no "votes" that disallows map change so assume they are all allowing
    If newVote = 1 Then
        msg = Left$(va("%s (%d) allowed map change. Total: %d/%d", getPlayerName(player, True), player, getLobbyChangeVoteTotal(), NET_numHumanPlayers()), 127)
        sendRoomSystemMessage msg
    End If
End Sub

Public Sub sendPlayerKickedVote(ByVal voteID As Long, ByVal newVote As Long)
    Dim w As MessageWriter, voteType As Long
    Set w = NETbeginEncode(NETbroadcastQueue(), NET_VOTE)
    NETuint32_t w, selectedPlayer
    voteType = NetVoteType_KICK_PLAYER
    NETuint8_t w, voteType
    NETuint32_t w, voteID
    NETuint8_t w, newVote
    NETend w
End Sub

' (VB) the index of the pending vote with this id, -1 (pendingKickVotes.end()) if none
Private Function findPendingKickVote(ByVal voteID As Long) As Long
    Dim i As Long
    For i = 0 To pkvCount - 1
        If pkvUniqueVoteId(i) = voteID Then
            findPendingKickVote = i
            Exit Function
        End If
    Next
    findPendingKickVote = -1
End Function

' bool PendingVoteKick::setPlayerVote(uint32_t sender, bool vote)
Private Function PendingVoteKick_setPlayerVote(ByVal it As Long, ByVal sender As Long, ByVal vote As Boolean) As Boolean
    If Not ULT(sender, MAX_PLAYERS) Then ASSERT_FAILURE "sender < votes.size()", "src/multivote.cpp:154", "setPlayerVote", "Invalid sender: %u", sender: Exit Function
    If pkvHasVote(it, sender) Then
        ' ignore double-votes
        PendingVoteKick_setPlayerVote = False
        Exit Function
    End If
    pkvHasVote(it, sender) = True
    pkvVote(it, sender) = vote
    PendingVoteKick_setPlayerVote = True
End Function

' std::vector::erase of pending vote it (the later ones move down)
Private Sub erasePendingKickVote(ByVal it As Long)
    Dim i As Long, p As Long
    For i = it To pkvCount - 2
        pkvUniqueVoteId(i) = pkvUniqueVoteId(i + 1)
        pkvRequesterPlayerId(i) = pkvRequesterPlayerId(i + 1)
        pkvTargetPlayerId(i) = pkvTargetPlayerId(i + 1)
        pkvStartTime(i) = pkvStartTime(i + 1)
        For p = 0 To MAX_PLAYERS - 1
            pkvHasVote(i, p) = pkvHasVote(i + 1, p)
            pkvVote(i, p) = pkvVote(i + 1, p)
        Next
    Next
    pkvCount = pkvCount - 1
End Sub

Private Sub recvPlayerKickVote(ByVal voteID As Long, ByVal sender As Long, ByVal newVote As Long)
    Dim it As Long, voteToKick As Boolean, outputMsg As String
    If Not ULT(sender, MAX_PLAYERS) Then ASSERT_FAILURE "sender < MAX_PLAYERS", "src/multivote.cpp:825", "recvPlayerKickVote", "Invalid sender: %u", sender: Exit Sub

    it = findPendingKickVote(voteID)
    If it < 0 Then
        ' didn't find the vote? may have already ended
        Exit Sub
    End If

    voteToKick = (newVote = 1)
    If PendingVoteKick_setPlayerVote(it, sender, voteToKick) Then
        If voteToKick Then
            outputMsg = va("A player voted FOR kicking: %s", getPlayerName(pkvTargetPlayerId(it), True))
            sendInGameSystemMessage outputMsg
            debug_ 844, LOG_INFO, "recvPlayerKickVote", "Player [%u] %s voted FOR kicking player: %s", sender, getPlayerName(sender, True), getPlayerName(pkvTargetPlayerId(it), True)
        Else
            If newVote = 0 Then
                outputMsg = va("A player voted AGAINST kicking: %s", getPlayerName(pkvTargetPlayerId(it), True))
                sendInGameSystemMessage outputMsg
                debug_ 852, LOG_INFO, "recvPlayerKickVote", "Player [%u] %s voted AGAINST kicking player: %s", sender, getPlayerName(sender, True), getPlayerName(pkvTargetPlayerId(it), True)
            Else
                outputMsg = va("A player's client ignored your vote to kick request (too frequent): %s", getPlayerName(pkvTargetPlayerId(it), True))
                addConsoleMessage outputMsg, DEFAULT_JUSTIFY, SYSTEM_MESSAGE, False ' only display to the host
                debug_ 858, LOG_INFO, "recvPlayerKickVote", "Player [%u] %s ignored vote to kick request for player: %s - (too frequent)", sender, getPlayerName(sender, True), getPlayerName(pkvTargetPlayerId(it), True)
            End If
        End If
    End If

    If handleVoteKickResult(it) Then
        ' got a result and handled it
        erasePendingKickVote it
    End If
End Sub

Private Sub sendPlayerMultiOptPreferencesBuiltin_player(ByVal playerIdx As Long)
    Dim w As MessageWriter, voteType As Long
    ensurePlayerPreferences
    If Not ULT(playerIdx, MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "playerIdx < MAX_CONNECTED_PLAYERS", "src/multivote.cpp:872", "sendPlayerMultiOptPreferencesBuiltin", "Invalid player idx: %u", playerIdx: Exit Sub
    If Not (whosResponsible(playerIdx) = selectedPlayer Or NetPlay_isHost) Then ASSERT_FAILURE "whosResponsible(playerIdx) == selectedPlayer || NetPlay.isHost", "src/multivote.cpp:873", "sendPlayerMultiOptPreferencesBuiltin", "Sending unexpected player prefs: %u", playerIdx: Exit Sub
    If GetGameMode() = GS_NORMAL Then ASSERT_FAILURE "GetGameMode() != GS_NORMAL", "src/multivote.cpp:874", "sendPlayerMultiOptPreferencesBuiltin", "Trying to send multiopt preferences after game started?": Exit Sub

    Set w = NETbeginEncode(NETnetQueue(NetPlay_hostPlayer), NET_VOTE)
    NETuint32_t w, selectedPlayer
    voteType = NetVoteType_LOBBY_OPTION_PREFERENCES_BUILTIN
    NETuint8_t w, voteType
    NETbuiltinPlayerPreferences w, playerPreferences(playerIdx)
    NETend w
End Sub

Private Function recvPlayerMultiOptPreferencesBuiltin(ByVal sender As Long, builtinPrefs As BuiltinPreferences) As Boolean
    Dim prefsChanged As Boolean
    If Not ULT(sender, MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "sender < MAX_CONNECTED_PLAYERS", "src/multivote.cpp:886", "recvPlayerMultiOptPreferencesBuiltin", "Invalid sender: %u", sender: Exit Function
    prefsChanged = setBuiltinPreferences(sender, builtinPrefs)
    If prefsChanged Then
        wz_command_interface_output_room_status_json True
    End If
    recvPlayerMultiOptPreferencesBuiltin = prefsChanged
End Function

Public Function recvVote(ByVal queue As Long, Optional ByVal inLobby As Boolean = True) As Boolean
    Dim player As Long, voteID As Long, voteType As Long, newVote As Long, builtinPrefs As BuiltinPreferences
    Dim senderIsSpectator As Boolean, validPrefs As Boolean, r As MessageReader, qi As Long
    If Not NetPlay_isHost Then
        ASSERT_FAILURE "false", "src/multivote.cpp:897", "recvVote", "Host only routine detected for client!"
        recvVote = False
        Exit Function
    End If
    ensurePlayerPreferences

    player = MAX_PLAYERS
    voteID = 0
    voteType = 0
    newVote = 0
    BuiltinPreferences_reset builtinPrefs
    qi = NETQUEUE_index(queue)
    If qi < NetPlay_players_size() Then senderIsSpectator = NetPlay_players_isSpectator(qi) Else senderIsSpectator = True
    validPrefs = False

    Set r = NETbeginDecode(queue, NET_VOTE)
    NETuint32_t r, player
    NETuint8_t r, voteType
    Select Case voteType
    Case NetVoteType_LOBBY_SETTING_CHANGE
        NETuint8_t r, newVote
    Case NetVoteType_KICK_PLAYER
        NETuint32_t r, voteID
        NETuint8_t r, newVote
    Case NetVoteType_LOBBY_OPTION_PREFERENCES_BUILTIN
        If inLobby Then
            validPrefs = NETbuiltinPlayerPreferences(r, builtinPrefs)
        End If
    End Select
    NETend r

    Select Case voteType
    Case NetVoteType_LOBBY_SETTING_CHANGE, NetVoteType_KICK_PLAYER
        If senderIsSpectator Then
            ' Silently ignore LOBBY_SETTING_CHANGE and KICK_PLAYER votes from spectators
            recvVote = False
            Exit Function
        End If
        If Not ULT(player, MAX_PLAYERS) Then
            debug_ 941, LOG_NET, "recvVote", "Invalid NET_VOTE from player %d: player id = %d", qi, player
            recvVote = False
            Exit Function
        End If
    Case NetVoteType_LOBBY_OPTION_PREFERENCES_BUILTIN
        If Not inLobby Or Not ULT(player, MAX_CONNECTED_PLAYERS) Then
            debug_ 948, LOG_NET, "recvVote", "Invalid NET_VOTE from player %d: player id = %d", qi, player
            recvVote = False
            Exit Function
        End If
    End Select

    If whosResponsible(player) <> qi Then
        debug_ 956, LOG_NET, "recvVote", "Invalid NET_VOTE from player %d: (for player id = %d)", qi, player
        recvVote = False
        Exit Function
    End If

    Select Case voteType
    Case NetVoteType_LOBBY_SETTING_CHANGE
        recvLobbyChangeVote player, newVote
        recvVote = True
        Exit Function
    Case NetVoteType_KICK_PLAYER
        recvPlayerKickVote voteID, player, newVote
        recvVote = True
        Exit Function
    Case NetVoteType_LOBBY_OPTION_PREFERENCES_BUILTIN
        If Not validPrefs Then
            debug_ 971, LOG_NET, "recvVote", "Invalid NET_VOTE from player %d: (for player id = %d)", qi, player
            recvVote = False
            Exit Function
        End If
        recvVote = recvPlayerMultiOptPreferencesBuiltin(player, builtinPrefs)
        Exit Function
    End Select

    recvVote = False
End Function

' Show a vote popup to allow changing maps or using the randomization feature.
Private Sub setupLobbyChangeVoteChoice()
    Dim notification As WZ_Notification
    'This shouldn't happen...
    If NetPlay_isHost Then
        ASSERT_FAILURE "false", "src/multivote.cpp:986", "setupLobbyChangeVoteChoice", "Host tried to send vote data to themself"
        Exit Sub
    End If

    If Not hasNotificationsWithTag(VOTE_TAG) Then
        notification.duration = 0
        notification.contentTitle = "Vote"
        notification.contentText = "Allow host to change map or randomize?"
        notification.action_title = "Allow"
        Set notification.action_onAction = WzFunc_new(FN_lobbyChangeVote_allow)
        notification.tag = VOTE_TAG
        addNotification notification, GAME_TICKS_PER_SEC * 1
    End If
End Sub

' Show a kick vote popup
Private Sub setupKickVoteChoice(ByVal voteID As Long, ByVal targetPlayer As Long)
    Dim targetIsActiveAIPlayer As Boolean, notificationTag As String, outputMsg As String, notification As WZ_Notification
    Dim pPlayerName As String, playerDisplayName As String
    'This shouldn't happen...
    If NetPlay_isHost Then
        ASSERT_FAILURE "false", "src/multivote.cpp:1012", "setupKickVoteChoice", "Host tried to send vote data to themself?"
        Exit Sub
    End If

    If Not ULT(targetPlayer, MAX_PLAYERS) Then
        ' Invalid targetPlayer
        Exit Sub
    End If

    targetIsActiveAIPlayer = NetPlay_players_allocated(targetPlayer) = False And NetPlay_players_ai(targetPlayer) >= 0 And Not NetPlay_players_isSpectator(targetPlayer)
    If Not NetPlay_players_allocated(targetPlayer) And Not targetIsActiveAIPlayer Then
        ' no active player to vote on
        Exit Sub
    End If

    If has_lastKickVoteForEachPlayer(targetPlayer) Then
        If ULT(realTime - lastKickVoteForEachPlayer(targetPlayer), MIN_INTERVAL_BETWEEN_PLAYER_KICK_VOTES_MS) Then
            ' Host is spamming kick requests - deny it automatically
            If targetPlayer <> selectedPlayer Then
                sendPlayerKickedVote voteID, 2
            End If
            Exit Sub
        End If
    End If

    If targetPlayer = selectedPlayer Then
        ' The vote is for the current player - just display a local console message and return
        addConsoleMessage "A vote was started to kick you from the game.", DEFAULT_JUSTIFY, SYSTEM_MESSAGE, False
        debug_ 1046, LOG_INFO, "setupKickVoteChoice", "A vote was started to kick you from the game."
        Exit Sub
    End If

    notificationTag = VOTE_KICK_TAG_PREFIX & CStr(targetPlayer)
    If hasNotificationsWithTag(notificationTag) Then
        ' dismiss existing notification targeting this player
        cancelOrDismissNotificationsWithTag notificationTag
    End If

    lastKickVoteForEachPlayer(targetPlayer) = realTime
    has_lastKickVoteForEachPlayer(targetPlayer) = True

    outputMsg = va("A vote was started to kick %s from the game.", getPlayerName(targetPlayer))
    addConsoleMessage outputMsg, DEFAULT_JUSTIFY, SYSTEM_MESSAGE, False
    debug_ 1062, LOG_INFO, "setupKickVoteChoice", "A vote was started to kick %s from the game.", getPlayerName(targetPlayer)

    notification.duration = 0
    pPlayerName = getPlayerName(targetPlayer)
    playerDisplayName = pPlayerName             ' (getPlayerName never returns nullptr here)
    notification.contentTitle = va("Vote To Kick: %s", playerDisplayName)
    notification.contentText = va("Should player %s be kicked from the game?", playerDisplayName)
    notification.action_title = "Yes, Kick Them"
    Set notification.action_onAction = WzFunc_new(FN_kickVote_yes, voteID)
    Set notification.onDismissed = WzFunc_new(FN_kickVote_onDismissed, voteID)
    notification.tag = notificationTag
    addNotification notification, GAME_TICKS_PER_SEC * 1
End Sub

' voteID, targetPlayer: 0 by default
Private Function sendVoteRequest(ByVal typ As Long, Optional ByVal voteID As Long = 0, Optional ByVal targetPlayer As Long = 0) As Boolean
    Dim w As MessageWriter, voteType As Long
    If Not NetPlay_isHost Then
        ASSERT_FAILURE "false", "src/multivote.cpp:1084", "sendVoteRequest", "Host only routine detected for client!"
        sendVoteRequest = False
        Exit Function
    End If

    'setup a vote popup for the clients
    Set w = NETbeginEncode(NETbroadcastQueue(), NET_VOTE_REQUEST)
    NETuint32_t w, selectedPlayer
    NETuint32_t w, targetPlayer
    NETuint32_t w, voteID
    voteType = typ
    NETuint8_t w, voteType
    NETend w

    sendVoteRequest = True
End Function

Public Function recvVoteRequest(ByVal queue As Long) As Boolean
    Dim sender As Long, targetPlayer As Long, voteID As Long, voteType As Long, r As MessageReader
    sender = MAX_PLAYERS
    targetPlayer = MAX_PLAYERS
    voteID = 0
    voteType = 0

    Set r = NETbeginDecode(queue, NET_VOTE_REQUEST)
    NETuint32_t r, sender
    NETuint32_t r, targetPlayer
    NETuint32_t r, voteID
    NETuint8_t r, voteType
    NETend r

    If Not ULT(sender, MAX_PLAYERS) Then
        debug_ 1113, LOG_NET, "recvVoteRequest", "Invalid NET_VOTE_REQUEST from player %d: player id = %d", NETQUEUE_index(queue), sender
        recvVoteRequest = False
        Exit Function
    End If

    If whosResponsible(sender) <> NETQUEUE_index(queue) Then
        debug_ 1119, LOG_NET, "recvVoteRequest", "Invalid NET_VOTE_REQUEST from player %d: (for player id = %d)", NETQUEUE_index(queue), sender
        recvVoteRequest = False
        Exit Function
    End If

    Select Case voteType
    Case NetVoteType_LOBBY_SETTING_CHANGE
        setupLobbyChangeVoteChoice
        recvVoteRequest = True
        Exit Function
    Case NetVoteType_KICK_PLAYER
        setupKickVoteChoice voteID, targetPlayer
        recvVoteRequest = True
        Exit Function
    Case NetVoteType_LOBBY_OPTION_PREFERENCES_BUILTIN
        recvVoteRequest = False
        Exit Function
    End Select

    recvVoteRequest = False
End Function

Public Sub startLobbyChangeVote()
    If Not NetPlay_isHost Then
        ASSERT_FAILURE "false", "src/multivote.cpp:1140", "startLobbyChangeVote", "Host only routine detected for client!"
        Exit Sub
    End If
    sendVoteRequest NetVoteType_LOBBY_SETTING_CHANGE, 0, NetPlay_hostPlayer
End Sub

' std::vector<uint32_t> getPlayersWhoCanVoteToKick(): the count, the players in result
Private Function getPlayersWhoCanVoteToKick(result() As Long) As Long
    Dim player As Long, n As Long, lim As Long
    lim = game.maxPlayers
    If lim > MAX_PLAYERS Then lim = MAX_PLAYERS
    ReDim result(0 To MAX_PLAYERS)
    For player = 0 To lim - 1
        If isHumanPlayer(player) And Not NetPlay_players_isSpectator(player) Then
            ' is an active (connected) human player who is *not* a spectator
            result(n) = player
            n = n + 1
        End If
    Next
    getPlayersWhoCanVoteToKick = n
End Function

' filterPlayersByTeam (sameTeam True) / getPlayersNotOnTeam (sameTeam False): the count, the players in result
Private Function filterPlayersByTeamT(players() As Long, ByVal count As Long, ByVal specifiedTeam As Long, ByVal sameTeam As Boolean, _
        result() As Long) As Long
    Dim i As Long, n As Long
    ReDim result(0 To MAX_PLAYERS)
    For i = 0 To count - 1
        If (checkedGetPlayerTeam(players(i)) = specifiedTeam) = sameTeam Then
            result(n) = players(i)
            n = n + 1
        End If
    Next
    filterPlayersByTeamT = n
End Function

' static_cast<size_t>(ceilf(x))
Private Function ceilfSize(ByVal x As Single) As Long
    ceilfSize = -Int(-x)
End Function

' optional<bool> PendingVoteKick::getKickVoteResult(): 0 false, 1 true, -1 nullopt
Private Function PendingVoteKick_getKickVoteResult(ByVal it As Long) As Long
    Dim targetIsActiveAIPlayer As Boolean, eligible_voters() As Long, nEligible As Long, target_player_team As Long, tgt As Long
    Dim target_player_teammembers() As Long, nTeam As Long, target_player_is_solo As Boolean, team_voteout_still_possible As Boolean
    Dim team_votes_for_kick As Long, team_votes_against_kick As Long, i As Long, player As Long, team_threshold As Long
    Dim team_votes_outstanding As Long, overall_votes_for_kick As Long, overall_votes_against_kick As Long, num_voting As Long
    Dim num_not_voting As Long, players_on_other_teams() As Long, nOther As Long, overall_vote_threshold As Long
    Dim overall_voteout_still_possible As Boolean
    tgt = pkvTargetPlayerId(it)
    targetIsActiveAIPlayer = NetPlay_players_allocated(tgt) = False And NetPlay_players_ai(tgt) >= 0 And Not NetPlay_players_isSpectator(tgt)
    If Not NetPlay_players_allocated(tgt) And Not targetIsActiveAIPlayer Then
        ' target player has already left / AI has lost
        PendingVoteKick_getKickVoteResult = 0
        Exit Function
    End If

    nEligible = getPlayersWhoCanVoteToKick(eligible_voters)

    ' Special case:
    ' - (If there are only two eligible voters): If they are on separate teams, vote fails immediately. If they are on the *same* team, just allow it (presumably they are against bots or the game would have already ended)
    If nEligible <= 2 Then
        If nEligible <= 1 Then
            ' If there's only 1 eligible voter, let them kick (presumably kicking an AI)
            PendingVoteKick_getKickVoteResult = 1
            Exit Function
        End If
        If checkedGetPlayerTeam(eligible_voters(0)) = checkedGetPlayerTeam(eligible_voters(1)) Then
            ' If they are on the *same* team, just allow it (presumably they are against bots or the game would have already ended)
            PendingVoteKick_getKickVoteResult = 1
            Exit Function
        End If
        ' If only two human players, but not on the same team, vote to kick can't succeed (the host can always quit, though)
        PendingVoteKick_getKickVoteResult = 0
        Exit Function
    End If

    target_player_team = checkedGetPlayerTeam(tgt)
    nTeam = filterPlayersByTeamT(eligible_voters, nEligible, target_player_team, True, target_player_teammembers)
    target_player_is_solo = nTeam <= 1
    team_voteout_still_possible = False

    ' Check if vote reached a team threshold
    If Not target_player_is_solo Then
        team_votes_for_kick = 0
        team_votes_against_kick = 0
        For i = 0 To nTeam - 1
            player = target_player_teammembers(i)
            If pkvHasVote(it, player) Then
                If pkvVote(it, player) Then
                    team_votes_for_kick = team_votes_for_kick + 1
                Else
                    team_votes_against_kick = team_votes_against_kick + 1
                End If
            End If
        Next
        ' If 50%+ of the target player's team agrees to kick, that's a success
        team_threshold = ceilfSize(CSng(nTeam) / 2!)
        If team_votes_for_kick >= team_threshold Then
            PendingVoteKick_getKickVoteResult = 1
            Exit Function
        End If
        team_votes_outstanding = nTeam - (team_votes_for_kick + team_votes_against_kick)
        If team_votes_for_kick + team_votes_outstanding >= team_threshold Then
            team_voteout_still_possible = True
        End If
        ' Otherwise, fall-through to overall thresholds
    End If

    overall_votes_for_kick = 0
    overall_votes_against_kick = 0
    For i = 0 To nEligible - 1
        player = eligible_voters(i)
        If pkvHasVote(it, player) Then
            If pkvVote(it, player) Then
                overall_votes_for_kick = overall_votes_for_kick + 1
            Else
                overall_votes_against_kick = overall_votes_against_kick + 1
            End If
        End If
    Next
    num_voting = overall_votes_for_kick + overall_votes_against_kick
    num_not_voting = nEligible - num_voting

    nOther = filterPlayersByTeamT(eligible_voters, nEligible, target_player_team, False, players_on_other_teams)
    If target_player_is_solo Then
        ' If target player is the only human player on their team: 50%+ of all other eligible voters vote to kick
        overall_vote_threshold = ceilfSize(CSng(nEligible) / 2!)
        If overall_vote_threshold < 2 Then overall_vote_threshold = 2
    Else
        ' If target player is *not* the only human player on their team, the smaller of:
        ' - 2/3 of all eligible voters
        ' or
        ' - num eligible voters on other teams
        ' but must be at least 2
        overall_vote_threshold = ceilfSize(CSng(nEligible) * 2! / 3!)
        If nOther < overall_vote_threshold Then overall_vote_threshold = nOther
        If overall_vote_threshold < 2 Then overall_vote_threshold = 2
    End If

    ' Check if vote reached an overall threshold
    If overall_votes_for_kick >= overall_vote_threshold Then
        PendingVoteKick_getKickVoteResult = 1
        Exit Function
    End If

    overall_voteout_still_possible = (overall_votes_for_kick + num_not_voting >= overall_vote_threshold)
    If (Not overall_voteout_still_possible And Not team_voteout_still_possible) Or num_not_voting = 0 Then
        ' didn't (and can't) hit any threshold
        PendingVoteKick_getKickVoteResult = 0
        Exit Function
    End If

    ' waiting for more votes
    PendingVoteKick_getKickVoteResult = -1
End Function

' Returns: true if PendingVote has a result and was handled, false if still waiting for results
Private Function handleVoteKickResult(ByVal it As Long) As Boolean
    Dim currentResult As Long, outputMsg As String, logMsg As String, tgt As Long
    If Not NetPlay_isHost Then
        ASSERT_FAILURE "false", "src/multivote.cpp:1317", "handleVoteKickResult", "Host only routine detected for client!"
        handleVoteKickResult = False
        Exit Function
    End If

    currentResult = PendingVoteKick_getKickVoteResult(it)
    If currentResult < 0 Then
        handleVoteKickResult = False
        Exit Function
    End If

    tgt = pkvTargetPlayerId(it)
    If currentResult = 1 Then
        outputMsg = va("The vote to kick player %s succeeded (sufficient votes in favor) - kicking", getPlayerName(tgt, True))
        sendInGameSystemMessage outputMsg
        logMsg = "kicked " & getPlayerName(tgt) & " : " & NetPlay_players_IPtextAddress(tgt) & " from the game"
        NETlogEntry logMsg, SYNC_FLAG, tgt
        kickPlayer tgt, "The players have voted to kick you from the game.", ERROR_KICKED, False
    Else
        ' Vote failed - message all players
        outputMsg = va("The vote to kick player %s failed (insufficient votes in favor)", getPlayerName(tgt, True))
        sendInGameSystemMessage outputMsg
    End If

    handleVoteKickResult = True
End Function

Public Sub processPendingKickVotes()
    Dim it As Long, outputMsg As String
    If Not NetPlay_isHost Then
        Exit Sub
    End If

    it = 0
    Do While it < pkvCount
        If Not ULT(realTime - pkvStartTime(it), PENDING_KICK_VOTE_TIMEOUT_MS) Then
            ' pending vote timed-out
            ' double-check if pending vote has a result (this might have changed if other players left)
            If Not handleVoteKickResult(it) Then
                ' dismiss the pending vote
                outputMsg = va("The vote to kick player %s failed (insufficient votes before timeout)", getPlayerName(pkvTargetPlayerId(it), True))
                sendInGameSystemMessage outputMsg
                debug_ 1363, LOG_INFO, "processPendingKickVotes", "%s", outputMsg
            End If
            erasePendingKickVote it
        Else
            it = it + 1
        End If
    Loop
End Sub

' (VB) pendingKickVotes.push_back(PendingVoteKick(unique_vote_id, requester_player_id, target_player_id)): the index
Private Function newPendingVoteKick(ByVal unique_vote_id As Long, ByVal requester_player_id As Long, ByVal target_player_id As Long) As Long
    Dim i As Long, p As Long
    i = pkvCount
    If i = 0 Then
        ReDim pkvUniqueVoteId(0 To 0): ReDim pkvRequesterPlayerId(0 To 0): ReDim pkvTargetPlayerId(0 To 0)
        ReDim pkvStartTime(0 To 0)
        ReDim pkvHasVote(0 To 0, 0 To MAX_PLAYERS - 1): ReDim pkvVote(0 To 0, 0 To MAX_PLAYERS - 1)
    ElseIf i > UBound(pkvUniqueVoteId) Then
        growPendingKickVotes i * 2
    End If
    pkvUniqueVoteId(i) = unique_vote_id
    pkvRequesterPlayerId(i) = requester_player_id
    pkvTargetPlayerId(i) = target_player_id
    For p = 0 To MAX_PLAYERS - 1
        pkvHasVote(i, p) = False
        pkvVote(i, p) = False
    Next
    If ULT(requester_player_id, MAX_PLAYERS) Then
        pkvHasVote(i, requester_player_id) = True
        pkvVote(i, requester_player_id) = True
    End If
    pkvHasVote(i, target_player_id) = True
    pkvVote(i, target_player_id) = False
    pkvStartTime(i) = realTime
    pkvCount = i + 1
    newPendingVoteKick = i
End Function

' (VB) the pending votes' arrays to n slots (the 2D arrays copied: ReDim Preserve keeps only the last dimension)
Private Sub growPendingKickVotes(ByVal n As Long)
    Dim hv() As Boolean, vv() As Boolean, i As Long, p As Long
    ReDim Preserve pkvUniqueVoteId(0 To n - 1)
    ReDim Preserve pkvRequesterPlayerId(0 To n - 1)
    ReDim Preserve pkvTargetPlayerId(0 To n - 1)
    ReDim Preserve pkvStartTime(0 To n - 1)
    ReDim hv(0 To n - 1, 0 To MAX_PLAYERS - 1)
    ReDim vv(0 To n - 1, 0 To MAX_PLAYERS - 1)
    For i = 0 To pkvCount - 1
        For p = 0 To MAX_PLAYERS - 1
            hv(i, p) = pkvHasVote(i, p)
            vv(i, p) = pkvVote(i, p)
        Next
    Next
    pkvHasVote = hv
    pkvVote = vv
End Sub

' requester_player_id: -1 for nullopt (selectedPlayer)
Public Function startKickVote(ByVal targetPlayer As Long, Optional ByVal requester_player_id As Long = -1) As Boolean
    Dim it As Long, currentStatus As Long, outputMsg As String, requester As Long
    If Not NetPlay_isHost Then
        ASSERT_FAILURE "false", "src/multivote.cpp:1377", "startKickVote", "Host only routine detected for client!"
        startKickVote = False
        Exit Function
    End If

    If requester_player_id < 0 Then requester = selectedPlayer Else requester = requester_player_id
    ' auto pendingVote = PendingVoteKick(last_vote_id++, ...): (VB) made in the vector's next slot, kept there only by push_back
    it = newPendingVoteKick(last_vote_id, requester, targetPlayer)
    pkvCount = pkvCount - 1                 ' not pushed yet
    last_vote_id = last_vote_id + 1

    currentStatus = PendingVoteKick_getKickVoteResult(it)
    If currentStatus >= 0 Then
        ' Vote either isn't possible or is a special case
        If currentStatus = 1 Then
            kickPlayer targetPlayer, "The host has kicked you from the game.", ERROR_KICKED, False
            startKickVote = True
            Exit Function
        Else
            ' message to the requester that player kick vote failed
            outputMsg = va("The vote to kick player %s failed", getPlayerName(targetPlayer, True))
            addConsoleMessage outputMsg, DEFAULT_JUSTIFY, SYSTEM_MESSAGE, False
            startKickVote = False
            Exit Function
        End If
    End If

    If has_lastKickVoteForEachPlayer(targetPlayer) Then
        If ULT(realTime - lastKickVoteForEachPlayer(targetPlayer), MIN_INTERVAL_BETWEEN_PLAYER_KICK_VOTES_MS + 5000) Then ' + extra on the sender side
            ' Prevent spamming kick votes
            outputMsg = va("Cannot request vote to kick player %s yet - please wait a bit longer", getPlayerName(targetPlayer, True))
            addConsoleMessage outputMsg, DEFAULT_JUSTIFY, SYSTEM_MESSAGE, False
            startKickVote = False
            Exit Function
        End If
    End If

    ' Store in the list of pending votes
    pkvCount = pkvCount + 1                 ' pendingKickVotes.push_back(pendingVote);

    ' Initiate a network vote
    sendVoteRequest NetVoteType_KICK_PLAYER, pkvUniqueVoteId(it), pkvTargetPlayerId(it)

    outputMsg = va("Starting vote to kick player: %s", getPlayerName(targetPlayer, True))
    addConsoleMessage outputMsg, DEFAULT_JUSTIFY, SYSTEM_MESSAGE, False
    debug_ 1418, LOG_INFO, "startKickVote", "Starting vote to kick player: %s", getPlayerName(targetPlayer, True)

    lastKickVoteForEachPlayer(targetPlayer) = realTime
    has_lastKickVoteForEachPlayer(targetPlayer) = True
    startKickVote = True
End Function

Public Sub cancelOrDismissVoteNotifications()
    cancelOrDismissNotificationsWithTag VOTE_TAG
    cancelOrDismissNotificationIfTagHasPrefix VOTE_KICK_TAG_PREFIX
End Sub

Public Sub cancelOrDismissKickVote(ByVal targetPlayer As Long)
    cancelOrDismissNotificationsWithTag VOTE_KICK_TAG_PREFIX & CStr(targetPlayer)
End Sub

Public Function multivote_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Select Case f.id
    Case FN_lobbyChangeVote_allow       ' [](const WZ_Notification&)
        sendLobbyChangeVoteData 1
    Case FN_kickVote_yes                ' [voteID](const WZ_Notification&)
        sendPlayerKickedVote f.cap(0), 1
    Case FN_kickVote_onDismissed        ' [voteID](const WZ_Notification&, WZ_Notification_Dismissal_Reason reason)
        If args(1) <> WZ_Notification_Dismissal_Reason_USER_DISMISSED Then Exit Function
        sendPlayerKickedVote f.cap(0), 0
    End Select
End Function
