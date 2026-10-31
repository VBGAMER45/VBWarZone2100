Attribute VB_Name = "quickchat_c"
' src/hci/quickchat.cpp, quickchat.h, whole: the quick chat (the lobby's and the game's quick chat forms, the
' messages and their texts, sending and receiving them, the spam mute). Its widget classes' fields are the data
' classes named as the classes; two names are over VB's 31 characters: WzQuickChatToggleMixedCheckButton's data
' class is QCToggleMixedCheckButtonData, WzQuickChatSendToPlayersSelector's QCSendToPlayersSelectorData.
' (VB) A class's nested enum is named with the class (WzQuickChatPanel::WzQuickChatSendToMode is
' WzQuickChatPanel_WzQuickChatSendToMode, its members WzQuickChatSendToMode_<name>); an enum class's members carry
' the enum's name (WzQuickChatMessage_REQUEST_PLEASE_WAIT); WzQuickChatDataContexts::X::Y is
' WzQuickChatDataContexts_X_Y. A std::vector<uint32_t> is a UInt32Vector (v(0 .. n - 1)); a
' std::unordered_set<uint32_t> is one too, with set semantics (UInt32Set_insert / _erase / _count) and the
' insertion order as its iteration order: MSVC's hash order would only show in formatReceivers' text, the order of
' the specific players in a network message, and the order the AI scripts get eventQuickChat (no script of 4.7.0
' handles it). std::sort of the players by position is a stable insertion sort (MSVC's for up to 32 elements).
' optional<WzQuickChatMessageData> is has_messageData and the data; sendQuickChat with the data is sendQuickChat_data.
' std::chrono::steady_clock's time points are QueryPerformanceCounter counts (as chat_c's; MSVC's steady_clock is
' the same counter), the per-player std::deque of times a ring of MAX_RECORDED_MESSAGE_TIMES + 1 slots.
' The std::weak_ptr captures of this module's widgets are widgweak_vb serials (WzQuickChatSection,
' WzQuickChatSendToSelector and WzQuickChatForm register themselves). to_display_string(WzQuickChatMessage) is
' to_display_string_WzQuickChatMessage (terrain_c has to_display_string(TerrainShaderQuality)).
' NETbeginEncodeSecured / NETbeginDecodeSecured are WzNet's (nettypes.cpp). checkedGetPlayerTeam is
' teamstrategy.cpp's (cluster J). Translations (_()) are not ported: the English text is used.
' (VB) widget class WzQuickChatButton : W_BUTTON
' (VB) widget class WzQuickChatSection : W_CLICKFORM
' (VB) widget class WzQuickChatToggleMixedCheckButton : W_BUTTON
' (VB) widget class WzQuickChatPlayerButton : W_BUTTON
' (VB) widget class WzQuickChatSendToPlayersSelector : WIDGET
' (VB) widget class WzQuickChatSendToSelector : WIDGET
' (VB) widget class ChooserOverlay : WIDGET
' (VB) widget class WzQuickChatPanel : WIDGET
' (VB) widget class WzQuickChatPanelTabs : MultichoiceWidget
' (VB) widget class WzQuickChatForm : W_CLICKFORM
Option Explicit

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Private Const MAX_CONNECTED_PLAYERS As Long = MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS

' quickchat.h: FOREACH_QUICKCHATMSG(GENERATE_ENUM): enum class WzQuickChatMessage : uint32_t
Public Enum WzQuickChatMessage
    ' LOBBY ONLY
    WzQuickChatMessage_LOBBY_GREETING_WAVE
    WzQuickChatMessage_LOBBY_GREETING_WELCOME
    WzQuickChatMessage_LOBBY_DESIRE_TO_SPECTATE
    WzQuickChatMessage_LOBBY_DESIRE_TO_PLAY
    WzQuickChatMessage_LOBBY_DESIRE_TO_SWITCH_TEAMS
    WzQuickChatMessage_LOBBY_REQUEST_CHECK_READY
    ' ALL (lobby + in-game global *and* team)
    WzQuickChatMessage_REQUEST_PLEASE_WAIT
    WzQuickChatMessage_REQUEST_LETS_GO
    WzQuickChatMessage_NOTICE_ALMOST_READY
    WzQuickChatMessage_NOTICE_READY
    WzQuickChatMessage_NOTICE_BRB
    WzQuickChatMessage_NOTICE_I_AM_BACK
    WzQuickChatMessage_REACTIONS_YES
    WzQuickChatMessage_REACTIONS_NO
    WzQuickChatMessage_REACTIONS_MAYBE
    WzQuickChatMessage_REACTIONS_NOT_YET
    WzQuickChatMessage_REACTIONS_SOON
    WzQuickChatMessage_REACTIONS_THANK_YOU
    WzQuickChatMessage_REACTIONS_NO_PROBLEM
    ' In-game reactions (global + team)
    WzQuickChatMessage_REACTIONS_WELL_PLAYED
    WzQuickChatMessage_REACTIONS_CENSORED
    WzQuickChatMessage_REACTIONS_SORRY_DO_NOT_UNDERSTAND_PLEASE_USE_QUICK_CHAT
    ' Global Taunts
    WzQuickChatMessage_REACTIONS_TAUNT_GET_READY
    WzQuickChatMessage_REACTIONS_TAUNT_YOURE_GOING_TO_REGRET_THAT
    WzQuickChatMessage_REACTIONS_TAUNT_BARELY_A_SCRATCH
    ' Team communication
    WzQuickChatMessage_TEAM_REQUEST_ATTACK_NOW
    WzQuickChatMessage_TEAM_REQUEST_GROUP_UP
    WzQuickChatMessage_TEAM_REQUEST_SPLIT_UP
    WzQuickChatMessage_TEAM_REQUEST_FOCUS_ATTACKS_WHERE_MARKED
    WzQuickChatMessage_TEAM_REQUEST_RETREAT
    WzQuickChatMessage_TEAM_REQUEST_NEED_HELP
    WzQuickChatMessage_TEAM_REQUEST_LEFT_SIDE
    WzQuickChatMessage_TEAM_REQUEST_RIGHT_SIDE
    WzQuickChatMessage_TEAM_REQUEST_TRANSFER_UNITS
    WzQuickChatMessage_TEAM_REQUEST_TRANSFER_REBUILD
    WzQuickChatMessage_TEAM_NOTICE_ATTACKING_NOW
    WzQuickChatMessage_TEAM_NOTICE_ON_MY_WAY
    WzQuickChatMessage_TEAM_NOTICE_OPPONENTS_COMING
    WzQuickChatMessage_TEAM_NOTICE_BEING_ATTACKED
    WzQuickChatMessage_TEAM_NOTICE_RUSHING_OILS
    WzQuickChatMessage_TEAM_NOTICE_OPPONENTS_RUSHING_OILS
    WzQuickChatMessage_TEAM_NOTICE_DONT_HAVE_POWER
    WzQuickChatMessage_TEAM_SUGGESTION_BUILD_MORE_UNITS
    WzQuickChatMessage_TEAM_SUGGESTION_BUILD_DIFFERENT_UNITS
    WzQuickChatMessage_TEAM_SUGGESTION_CHECK_TEAM_STRATEGY
    WzQuickChatMessage_TEAM_SUGGESTION_RESEARCH_DIFFERENT_TECH
    WzQuickChatMessage_TEAM_SUGGESTION_KEEP_RESEARCH_BUSY
    WzQuickChatMessage_TEAM_SUGGESTION_BUILD_ANTI_AIR
    WzQuickChatMessage_TEAM_SUGGESTION_REPAIR_UNITS
    WzQuickChatMessage_TEAM_SUGGESTION_BUILD_REPAIR_FACILITIES
    WzQuickChatMessage_TEAM_SUGGESTION_BUILD_POWER_GEN
    WzQuickChatMessage_TEAM_SUGGESTION_BUILD_CAPTURE_OILS
    ' Team reactions
    WzQuickChatMessage_REACTIONS_TEAM_THAT_DIDNT_GO_WELL
    WzQuickChatMessage_REACTIONS_TEAM_I_HAVE_ANOTHER_PLAN
    ' End game reactions
    WzQuickChatMessage_REACTIONS_ENDGAME_GOOD_GAME
    WzQuickChatMessage_REACTIONS_ENDGAME_I_GIVE_UP
    WzQuickChatMessage_REACTIONS_ENDGAME_SORRY_HAVE_TO_LEAVE
    ' FROM THIS POINT ON - ONLY INTERNAL MESSAGES!
    ' WZ-generated internal messages - not for users to deliberately send
    WzQuickChatMessage_INTERNAL_MSG_DELIVERY_FAILURE_TRY_AGAIN    ' This should always be the first internal message!
    WzQuickChatMessage_INTERNAL_LOBBY_NOTICE_MAP_DOWNLOADED
    WzQuickChatMessage_INTERNAL_ADMIN_ACTION_NOTICE
    WzQuickChatMessage_INTERNAL_LOCALIZED_LOBBY_NOTICE
    ' Always last
    WzQuickChatMessage_MESSAGE_COUNT
End Enum

Public Const WzQuickChatMessage_FIRST_INTERNAL_MSG_VALUE As Long = WzQuickChatMessage_INTERNAL_MSG_DELIVERY_FAILURE_TRY_AGAIN

Public Enum WzQuickChatContext
    WzQuickChatContext_Lobby
    WzQuickChatContext_InGame
End Enum

Public Enum WzQuickChatMode
    WzQuickChatMode_Global
    WzQuickChatMode_Team
    WzQuickChatMode_Cheats
    WzQuickChatMode_EndGame
End Enum

' (VB) std::vector<uint32_t>: v(0 .. n - 1); also a std::unordered_set<uint32_t> (UInt32Set_*)
Public Type UInt32Vector
    v() As Long
    n As Long
End Type

' struct WzQuickChatTargeting; all is all_ (a VB word)
Public Type WzQuickChatTargeting
    all_ As Boolean                     ' = false
    humanTeammates As Boolean           ' = false
    aiTeammates As Boolean              ' = false
    specificPlayers As UInt32Vector     ' std::unordered_set<uint32_t>
End Type

Public Type WzQuickChatMessageData
    dataContext As Long                 ' uint32_t = 0
    dataA As Long                       ' uint32_t = 0
    dataB As Long                       ' uint32_t = 0
End Type

' Begin: DataContext Enums for specific messages
' - INTERNAL_ADMIN_ACTION_NOTICE
Public Enum WzQuickChatDataContexts_INTERNAL_ADMIN_ACTION_NOTICE_Context
    WzQuickChatDataContexts_INTERNAL_ADMIN_ACTION_NOTICE_Context_Invalid = 0
    WzQuickChatDataContexts_INTERNAL_ADMIN_ACTION_NOTICE_Context_Team
    WzQuickChatDataContexts_INTERNAL_ADMIN_ACTION_NOTICE_Context_Position
    WzQuickChatDataContexts_INTERNAL_ADMIN_ACTION_NOTICE_Context_Color
    WzQuickChatDataContexts_INTERNAL_ADMIN_ACTION_NOTICE_Context_Faction
End Enum

' - INTERNAL_LOCALIZED_LOBBY_NOTICE
Public Enum WzQuickChatDataContexts_INTERNAL_LOCALIZED_LOBBY_NOTICE_Context
    WzQuickChatDataContexts_INTERNAL_LOCALIZED_LOBBY_NOTICE_Context_Invalid = 0
    WzQuickChatDataContexts_INTERNAL_LOCALIZED_LOBBY_NOTICE_Context_NotReadyKickWarning
    WzQuickChatDataContexts_INTERNAL_LOCALIZED_LOBBY_NOTICE_Context_NotReadyKicked
    WzQuickChatDataContexts_INTERNAL_LOCALIZED_LOBBY_NOTICE_Context_PlayerShouldCheckReadyNotice
End Enum

' class WzQuickChatToggleMixedCheckButton: enum class MixedCheckState
Public Enum WzQuickChatToggleMixedCheckButton_MixedCheckState
    MixedCheckState_Unchecked
    MixedCheckState_Mixed
    MixedCheckState_Checked
End Enum

' class WzQuickChatSendToSelector: enum class PlayerSelectionState
Public Enum WzQuickChatSendToSelector_PlayerSelectionState
    PlayerSelectionState_NotAvailable
    PlayerSelectionState_Unselected
    PlayerSelectionState_Selected
End Enum

' class WzQuickChatPanel: enum class WzQuickChatSendToMode
Public Enum WzQuickChatPanel_WzQuickChatSendToMode
    WzQuickChatSendToMode_GlobalSendOptions
    WzQuickChatSendToMode_TeamSendOptions
End Enum

' the lambdas (WzFunc: quickchat_c_fn)
Public Const FN_qcSection_messagesList_calcLayout As Long = 15000      ' WzQuickChatSection::initialize: messagesList->setCalcLayout(LAMBDA_CALCLAYOUT_SIMPLE({...}))
Public Const FN_qcSection_chatOption_onClick As Long = 15001           ' addChatOption: [message, psWeakSection](W_BUTTON& widg)
Public Const FN_qcSendTo_clicked_task As Long = 15002                  ' WzQuickChatSendToSelector::clicked: widgScheduleTask([weakThis](){...})
Public Const FN_qcSendTo_openChooser_task As Long = 15003              ' openChooser: widgScheduleTask([pWeakThis]() {...})
Public Const FN_qcSendTo_closeChooser_task As Long = 15004             ' closeChooser: widgScheduleTask([pWeakThis]() {...})
Public Const FN_qcChooserOverlay_calcLayout As Long = 15005            ' ChooserOverlay(): setCalcLayout(LAMBDA_CALCLAYOUT_SIMPLE({...}))
Public Const FN_qcChooserOverlay_onClicked As Long = 15006             ' openChooser: ChooserOverlay's [pWeakThis]() {...}
Public Const FN_qcPlayersSel_humanTeammates_onClick As Long = 15007    ' WzQuickChatSendToPlayersSelector::initialize: [psWeakParent](W_BUTTON& widg)
Public Const FN_qcPlayersSel_humanTeammates_calcLayout As Long = 15008 ' humanTeammates->setCalcLayout(LAMBDA_CALCLAYOUT_SIMPLE({...}))
Public Const FN_qcPlayersSel_aiTeammates_onClick As Long = 15009       ' [psWeakParent](W_BUTTON& widg)
Public Const FN_qcPlayersSel_aiTeammates_calcLayout As Long = 15010    ' aiTeammates->setCalcLayout(LAMBDA_CALCLAYOUT_SIMPLE({...}))
Public Const FN_qcPlayersSel_player_onClick As Long = 15011            ' [psWeakParent, playerIdx](W_BUTTON& widg)
Public Const FN_qcPlayersSel_playersList_calcLayout As Long = 15012    ' [maxPlayersListHeight](WIDGET *psWidget)
Public Const FN_qcPanel_tabPress_task As Long = 15013                  ' WzQuickChatPanel::handleTabPress: widgScheduleTask([targetWidgetToTakeFocus](){...})
Public Const FN_qcForm_panelSwitcher_onChoose As Long = 15014          ' WzQuickChatForm::initialize: [](MultibuttonWidget& widget, int newValue)
Public Const FN_qcForm_panelSwitcher_task As Long = 15015              ' widgScheduleTask([newValue, weakParent](){...})
Public Const FN_qcForm_panelSwitcher_calcLayout As Long = 15016        ' panelSwitcher->setCalcLayout(LAMBDA_CALCLAYOUT_SIMPLE({...}))

' MARK: - Globals

Private defaultTeamChatTargeting As WzQuickChatTargeting
Private lastTeamChatTargeting As WzQuickChatTargeting
' static std::array<std::deque<std::chrono::steady_clock::time_point>, MAX_CONNECTED_PLAYERS> lastQuickChatMessageTimes:
' (VB) per player a ring of QC_TIMES_SLOTS QueryPerformanceCounter counts, its front and size
Private lastQuickChatMessageTimes() As Currency         ' (player, slot)
Private lastQuickChatMessageTimes_first() As Long
Private lastQuickChatMessageTimes_size() As Long
Private lastQuickChatMessageTimes_allocated As Boolean
Private perfFrequency As Currency                       ' (VB) QueryPerformanceFrequency: counts per second

' MARK: - Message throttling

Private Const DEFAULT_MAX_MESSAGES_IN_INTERVAL As Long = 8
Private Const MAX_RECORDED_MESSAGE_TIMES As Long = 200
Private Const MESSAGE_INTERVAL As Long = 5000                   ' std::chrono::milliseconds
Private Const MESSAGE_THROTTLE_TIMEOUT_INTERVAL As Long = 5000  ' std::chrono::milliseconds
Private Const QC_TIMES_SLOTS As Long = MAX_RECORDED_MESSAGE_TIMES + 1   ' (VB) push_back happens before the pop_fronts

' MARK: - Classes

Private Const WZCOL_QUICKCHAT_TABS_FILL_COLOR As Long = &HDC6E0019         ' pal_RGBA(25, 0, 110, 220)
Private Const WZCOL_QUICKCHAT_TABS_FILL_COLOR_DARK As Long = &HFA46000A    ' pal_RGBA(10, 0, 70, 250)
Private Const WZCOL_QUICKCHAT_TABS_BORDER_LIGHT As Long = &H96FFFFFF       ' pal_RGBA(255, 255, 255, 150)
Private Const WZCOL_QUICKCHAT_BUTTON_FILL_COLOR As Long = &HFF430000       ' pal_RGBA(0, 0, 67, 255)

Private Const QuickChatButtonVerticalPadding As Long = 5
Private Const QuickChatButtonHorizontalPadding As Long = 5

Private Const QC_PLAYER_BUTTON_PADDING As Long = 20
Private Const QC_PLAYER_BUTTON_HEIGHT As Long = 20

Private Const QCFORM_INTERNAL_PADDING As Long = 10
Private Const QCFORM_PANEL_TABS_HEIGHT As Long = 20

Private Const QC_SENDTO_SELECTOR_INTERNAL_SPACING As Long = 4

'--------------------------------------------------------------------------
' (VB) helpers

' a PIELIGHT's alpha (its byte.a: the top byte)
Private Function pielightAlpha(ByVal c As Long) As Long
    pielightAlpha = SHRU(c, 24)
End Function

Private Function pielightWithAlpha(ByVal c As Long, ByVal a As Long) As Long
    pielightWithAlpha = (c And &HFFFFFF) Or SHL(a And &HFF&, 24)
End Function

' std::vector<uint32_t>::push_back
Public Sub UInt32Vector_push_back(vec As UInt32Vector, ByVal value As Long)
    If vec.n = 0 Then
        ReDim vec.v(0 To 7)
    ElseIf vec.n > UBound(vec.v) Then
        ReDim Preserve vec.v(0 To 2 * vec.n - 1)
    End If
    vec.v(vec.n) = value
    vec.n = vec.n + 1
End Sub

Public Sub UInt32Vector_clear(vec As UInt32Vector)
    Erase vec.v
    vec.n = 0
End Sub

' std::unordered_set<uint32_t>::count
Public Function UInt32Set_count(s As UInt32Vector, ByVal value As Long) As Long
    Dim i As Long
    For i = 0 To s.n - 1
        If s.v(i) = value Then
            UInt32Set_count = 1
            Exit Function
        End If
    Next
End Function

' std::unordered_set<uint32_t>::insert (a new element goes last)
Public Sub UInt32Set_insert(s As UInt32Vector, ByVal value As Long)
    If UInt32Set_count(s, value) = 0 Then
        UInt32Vector_push_back s, value
    End If
End Sub

' std::unordered_set<uint32_t>::erase (the others keep their order)
Public Sub UInt32Set_erase(s As UInt32Vector, ByVal value As Long)
    Dim i As Long, j As Long
    For i = 0 To s.n - 1
        If s.v(i) = value Then
            For j = i To s.n - 2
                s.v(j) = s.v(j + 1)
            Next
            s.n = s.n - 1
            If s.n = 0 Then Erase s.v
            Exit Sub
        End If
    Next
End Sub

' std::sort(v.begin(), v.end(), [](uint32_t playerIdxA, uint32_t playerIdxB) { return
' NetPlay.players[playerIdxA].position < NetPlay.players[playerIdxB].position; }): MSVC's insertion sort (stable)
Public Sub UInt32Vector_sortByPlayerPosition(vec As UInt32Vector)
    Dim i As Long, j As Long, value As Long, pos As Long
    For i = 1 To vec.n - 1
        value = vec.v(i)
        pos = NetPlay_players_position(value)
        j = i
        Do While j > 0
            If Not (pos < NetPlay_players_position(vec.v(j - 1))) Then Exit Do
            vec.v(j) = vec.v(j - 1)
            j = j - 1
        Loop
        vec.v(j) = value
    Next
End Sub

' std::all_of(v.begin(), v.end(), [&s](uint32_t playerIdx) -> bool { return s.count(playerIdx) > 0; })
Private Function UInt32Vector_allIn(vec As UInt32Vector, s As UInt32Vector) As Boolean
    Dim i As Long
    For i = 0 To vec.n - 1
        If Not (UInt32Set_count(s, vec.v(i)) > 0) Then Exit Function
    Next
    UInt32Vector_allIn = True
End Function

' std::any_of(v.begin(), v.end(), [&s](uint32_t playerIdx) -> bool { return s.count(playerIdx) > 0; })
Private Function UInt32Vector_anyIn(vec As UInt32Vector, s As UInt32Vector) As Boolean
    Dim i As Long
    For i = 0 To vec.n - 1
        If UInt32Set_count(s, vec.v(i)) > 0 Then
            UInt32Vector_anyIn = True
            Exit Function
        End If
    Next
End Function

Private Sub lastQuickChatMessageTimes_alloc()
    If lastQuickChatMessageTimes_allocated Then Exit Sub
    ReDim lastQuickChatMessageTimes(0 To MAX_CONNECTED_PLAYERS - 1, 0 To QC_TIMES_SLOTS - 1)
    ReDim lastQuickChatMessageTimes_first(0 To MAX_CONNECTED_PLAYERS - 1)
    ReDim lastQuickChatMessageTimes_size(0 To MAX_CONNECTED_PLAYERS - 1)
    QueryPerformanceFrequency perfFrequency
    lastQuickChatMessageTimes_allocated = True
End Sub

' lastQuickChatMessageTimes[p][i] (0 = the front)
Private Function lastQuickChatMessageTimes_at(ByVal p As Long, ByVal i As Long) As Currency
    lastQuickChatMessageTimes_at = lastQuickChatMessageTimes(p, (lastQuickChatMessageTimes_first(p) + i) Mod QC_TIMES_SLOTS)
End Function

Private Sub lastQuickChatMessageTimes_clear(ByVal p As Long)
    lastQuickChatMessageTimes_first(p) = 0
    lastQuickChatMessageTimes_size(p) = 0
End Sub

' (a duration in steady_clock counts as duration_cast<std::chrono::milliseconds>: truncated)
Private Function countsToMs(ByVal d As Currency) As Double
    countsToMs = Fix(CDbl(d) * 1000# / CDbl(perfFrequency))
End Function

' MARK: - Message throttling

' optional<std::chrono::steady_clock::time_point> playerSpamMutedUntil(uint32_t playerIdx): has_value(), the value
' (a QueryPerformanceCounter count) in mutedUntil
Public Function playerSpamMutedUntil(ByVal playerIdx As Long, mutedUntil As Currency) As Boolean
    Dim maxMessageInLastInterval As Long, currentTime As Currency, newestRecordedMessageTime As Currency
    Dim numMessagesInLastInterval As Long, it As Long
    lastQuickChatMessageTimes_alloc
    If Not ULT(playerIdx, MAX_CONNECTED_PLAYERS) Then   ' playerIdx >= lastQuickChatMessageTimes.size()
        playerSpamMutedUntil = False
        Exit Function
    End If

    If lastQuickChatMessageTimes_size(playerIdx) = 0 Then
        playerSpamMutedUntil = False
        Exit Function
    End If

    maxMessageInLastInterval = DEFAULT_MAX_MESSAGES_IN_INTERVAL
    QueryPerformanceCounter currentTime
    newestRecordedMessageTime = lastQuickChatMessageTimes_at(playerIdx, lastQuickChatMessageTimes_size(playerIdx) - 1)
    numMessagesInLastInterval = 0
    For it = lastQuickChatMessageTimes_size(playerIdx) - 1 To 0 Step -1  ' rbegin() .. rend()
        ' (duration_cast<std::chrono::seconds>(...) <= MESSAGE_INTERVAL: whole seconds compared as milliseconds)
        If Fix(CDbl(newestRecordedMessageTime - lastQuickChatMessageTimes_at(playerIdx, it)) / CDbl(perfFrequency)) * 1000# <= MESSAGE_INTERVAL Then
            numMessagesInLastInterval = numMessagesInLastInterval + 1
        Else
            Exit For
        End If
    Next

    If numMessagesInLastInterval >= maxMessageInLastInterval And countsToMs(currentTime - newestRecordedMessageTime) < MESSAGE_THROTTLE_TIMEOUT_INTERVAL Then
        ' throttle incoming messages
        mutedUntil = newestRecordedMessageTime + perfFrequency * (MESSAGE_THROTTLE_TIMEOUT_INTERVAL \ 1000)
        playerSpamMutedUntil = True
        Exit Function
    End If

    playerSpamMutedUntil = False
End Function

Public Sub playerSpamMuteNotifyIndexSwap(ByVal playerIndexA As Long, ByVal playerIndexB As Long)
    Dim i As Long, t As Currency
    lastQuickChatMessageTimes_alloc
    If Not ULT(playerIndexA, MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "playerIndexA < lastQuickChatMessageTimes.size()", "src/hci/quickchat.cpp:109", "playerSpamMuteNotifyIndexSwap", "playerIndexA invalid: %u", playerIndexA: Exit Sub
    If Not ULT(playerIndexB, MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "playerIndexB < lastQuickChatMessageTimes.size()", "src/hci/quickchat.cpp:110", "playerSpamMuteNotifyIndexSwap", "playerIndexB invalid: %u", playerIndexB: Exit Sub
    ' std::swap(lastQuickChatMessageTimes[playerIndexA], lastQuickChatMessageTimes[playerIndexB]);
    For i = 0 To QC_TIMES_SLOTS - 1
        t = lastQuickChatMessageTimes(playerIndexA, i)
        lastQuickChatMessageTimes(playerIndexA, i) = lastQuickChatMessageTimes(playerIndexB, i)
        lastQuickChatMessageTimes(playerIndexB, i) = t
    Next
    i = lastQuickChatMessageTimes_first(playerIndexA)
    lastQuickChatMessageTimes_first(playerIndexA) = lastQuickChatMessageTimes_first(playerIndexB)
    lastQuickChatMessageTimes_first(playerIndexB) = i
    i = lastQuickChatMessageTimes_size(playerIndexA)
    lastQuickChatMessageTimes_size(playerIndexA) = lastQuickChatMessageTimes_size(playerIndexB)
    lastQuickChatMessageTimes_size(playerIndexB) = i
End Sub

Public Sub playerSpamMuteReset(ByVal playerIndex As Long)
    lastQuickChatMessageTimes_alloc
    If Not ULT(playerIndex, MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "playerIndex < lastQuickChatMessageTimes.size()", "src/hci/quickchat.cpp:116", "playerSpamMuteReset", "playerIndex invalid: %u", playerIndex: Exit Sub
    lastQuickChatMessageTimes_clear playerIndex
End Sub

Public Sub recordPlayerMessageSent(ByVal playerIdx As Long)
    Dim mutedUntil As Currency, now_ As Currency
    lastQuickChatMessageTimes_alloc
    If playerSpamMutedUntil(playerIdx, mutedUntil) Then
        ' ignore incoming message
        Exit Sub
    End If
    ' lastQuickChatMessageTimes[playerIdx].push_back(std::chrono::steady_clock::now());
    QueryPerformanceCounter now_
    lastQuickChatMessageTimes(playerIdx, (lastQuickChatMessageTimes_first(playerIdx) + lastQuickChatMessageTimes_size(playerIdx)) Mod QC_TIMES_SLOTS) = now_
    lastQuickChatMessageTimes_size(playerIdx) = lastQuickChatMessageTimes_size(playerIdx) + 1
    Do While lastQuickChatMessageTimes_size(playerIdx) > MAX_RECORDED_MESSAGE_TIMES
        ' pop_front()
        lastQuickChatMessageTimes_first(playerIdx) = (lastQuickChatMessageTimes_first(playerIdx) + 1) Mod QC_TIMES_SLOTS
        lastQuickChatMessageTimes_size(playerIdx) = lastQuickChatMessageTimes_size(playerIdx) - 1
    Loop
End Sub

Private Function isInternalMessage(ByVal msg As WzQuickChatMessage) As Boolean
    isInternalMessage = msg >= WzQuickChatMessage_FIRST_INTERNAL_MSG_VALUE
End Function

' MARK: - Helper functions

' static optional<size_t> numberButtonPressed(): has_value(), the value in keyPressForNumber
Private Function numberButtonPressed(keyPressForNumber As Long) As Boolean
    numberButtonPressed = True
    If keyPressed(KEY_1) Or keyPressed(KEY_KP_1) Then
        keyPressForNumber = 0
    ElseIf keyPressed(KEY_2) Or keyPressed(KEY_KP_2) Then
        keyPressForNumber = 1
    ElseIf keyPressed(KEY_3) Or keyPressed(KEY_KP_3) Then
        keyPressForNumber = 2
    ElseIf keyPressed(KEY_4) Or keyPressed(KEY_KP_4) Then
        keyPressForNumber = 3
    ElseIf keyPressed(KEY_5) Or keyPressed(KEY_KP_5) Then
        keyPressForNumber = 4
    ElseIf keyPressed(KEY_6) Or keyPressed(KEY_KP_6) Then
        keyPressForNumber = 5
    ElseIf keyPressed(KEY_7) Or keyPressed(KEY_KP_7) Then
        keyPressForNumber = 6
    ElseIf keyPressed(KEY_8) Or keyPressed(KEY_KP_8) Then
        keyPressForNumber = 7
    ElseIf keyPressed(KEY_9) Or keyPressed(KEY_KP_9) Then
        keyPressForNumber = 8
    ElseIf keyPressed(KEY_0) Or keyPressed(KEY_KP_0) Then
        keyPressForNumber = 9
    Else
        numberButtonPressed = False
    End If
End Function

Private Function getPlayerTeammates(ByVal playerIdx As Long, Optional ByVal includeAIs As Boolean = False, Optional ByVal includeAlliances As Boolean = False) As UInt32Vector
    Dim inputPlayerTeam As Long, playersOnSameTeamAsInputPlayer As UInt32Vector, player As Long
    inputPlayerTeam = checkedGetPlayerTeam(playerIdx)
    For player = 0 To MIN_(game.maxPlayers, MAX_PLAYERS) - 1
        If player = playerIdx Then
            GoTo continue_
        End If
        ' ((checkedGetPlayerTeam(player) != inputPlayerTeam) && (!includeAlliances || !aiCheckAlliances(playerIdx, player)))
        If checkedGetPlayerTeam(player) <> inputPlayerTeam Then
            If Not includeAlliances Then
                GoTo continue_
            ElseIf Not aiCheckAlliances(playerIdx, player) Then
                GoTo continue_
            End If
        End If
        If isHumanPlayer(player) Or NetPlay_players_difficulty(player) = AIDifficulty_HUMAN Or (includeAIs And NetPlay_players_ai(player) >= 0 And NetPlay_players_difficulty(player) <> AIDifficulty_DISABLED) Then
            ' is an active (connected) human player, was a human player (probably disconnected), or an AI
            UInt32Vector_push_back playersOnSameTeamAsInputPlayer, player
        End If
continue_:
    Next

    getPlayerTeammates = playersOnSameTeamAsInputPlayer
End Function

' MARK: - WzQuickChatButton

Public Sub WzQuickChatButton_construct(ByVal self As WIDGET)
    W_BUTTON_construct self
    self.kind = K_WzQuickChatButton ' (VB) virtual calls from here on reach WzQuickChatButton's overrides
    WzQuickChatButton_(self).lastWidgetWidth = 0
    WzQuickChatButton_(self).isTruncated = False
End Sub

Public Function WzQuickChatButton_make(ByVal messageText As String, Optional ByVal disabled As Boolean = False) As WIDGET
    Dim result As WIDGET
    Set result = WIDGET_alloc(K_WzQuickChatButton)
    WzQuickChatButton_construct result
    W_BUTTON_(result).pText = messageText
    Set WzQuickChatButton_make = result
End Function

Public Function WzQuickChatButton_getIdealBaseButtonHeight() As Long
    WzQuickChatButton_getIdealBaseButtonHeight = iV_GetTextLineSize(font_regular) + (QuickChatButtonVerticalPadding * 2)
End Function

' Display the text, shrunk to fit, with a faint border around it (that gets bright when highlighted)
Public Sub WzQuickChatButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As WzQuickChatButton, b As W_BUTTON, pParent As WIDGET, xPos As Long, yPos As Long, h As Long, w As Long
    Dim isDisabled As Boolean, isHighlight As Boolean, borderColor As Long, msgColour As Long, fontID As iV_fonts
    Dim availableButtonTextWidth As Long, result As iV_fonts, textX0 As Long, textY0 As Long, maxTextDisplayableWidth As Long
    Set d = WzQuickChatButton_(self)
    Set b = W_BUTTON_(self)
    Set pParent = self.parent()
    If Not (Not pParent Is Nothing) Then ASSERT_FAILURE "pParent != nullptr", "src/hci/quickchat.cpp:244", "display", "Keymap buttons should have a parent container!": Exit Sub

    xPos = xOffset + self.x()
    yPos = yOffset + self.y()
    h = self.height()
    w = self.width()

    isDisabled = (b.state And WBUT_DISABLE) <> 0
    isHighlight = (b.state And WBUT_HIGHLIGHT) <> 0

    ' Draw background
    pie_BoxFill xPos, yPos, xPos + w, yPos + h, WZCOL_QUICKCHAT_BUTTON_FILL_COLOR
    If isHighlight And Not isDisabled Then borderColor = pal_RGBA(255, 255, 255, 200) Else borderColor = pal_RGBA(255, 255, 255, 75)
    iV_Box xPos + 1, yPos + 1, xPos + w, yPos + h, borderColor

    If isHighlight And Not isDisabled Then
        pie_UniTransBoxFill xPos, yPos, xPos + w, yPos + h, pal_RGBA(0, 0, 255, 70)
    End If

    ' Select label text and color
    If Not isDisabled Then msgColour = psPalette(WZCOL_FORM_TEXT) Else msgColour = psPalette(WZCOL_FORM_DISABLE)

    fontID = d.wzMessageText.getFontID()
    If fontID = font_count Or d.lastWidgetWidth <> self.width() Or d.wzMessageText.getText() <> b.pText Then
        fontID = b.FontID
    End If
    d.wzMessageText.setText b.pText, fontID

    availableButtonTextWidth = w - (QuickChatButtonHorizontalPadding * 2)
    If d.wzMessageText.width() > availableButtonTextWidth Then
        ' text would exceed the bounds of the button area
        ' try to shrink font so it fits
        Do
            If fontID = font_small Or fontID = font_bar Then
                Exit Do
            End If
            If Not iV_ShrinkFont(fontID, result) Then
                Exit Do
            End If
            fontID = result
            d.wzMessageText.setText b.pText, fontID
        Loop While d.wzMessageText.width() > availableButtonTextWidth
    End If
    d.lastWidgetWidth = self.width()

    textX0 = xPos + QuickChatButtonHorizontalPadding
    ' (static_cast<int>(yPos + (h - lineSize()) / 2 - float(aboveBase())): integers throughout)
    textY0 = yPos + (h - d.wzMessageText.lineSize()) \ 2 - d.wzMessageText.aboveBase()

    maxTextDisplayableWidth = w - (QuickChatButtonHorizontalPadding * 2)
    d.isTruncated = maxTextDisplayableWidth < d.wzMessageText.width()
    If d.isTruncated Then
        maxTextDisplayableWidth = maxTextDisplayableWidth - (iV_GetEllipsisWidth(d.wzMessageText.getFontID()) + 2)
    End If
    d.wzMessageText.render textX0, textY0, msgColour, 0!, maxTextDisplayableWidth

    If d.isTruncated Then
        ' Render ellipsis
        iV_DrawEllipsis d.wzMessageText.getFontID(), textX0 + maxTextDisplayableWidth + 2, textY0, msgColour
    End If
End Sub

Public Function WzQuickChatButton_getTip(ByVal self As WIDGET) As String
    Dim d As WzQuickChatButton
    Set d = WzQuickChatButton_(self)
    If Len(W_BUTTON_(self).pTip) = 0 And Not d.isTruncated Then
        WzQuickChatButton_getTip = ""
        Exit Function
    End If

    If Not d.isTruncated Then
        WzQuickChatButton_getTip = W_BUTTON_(self).pTip
        Exit Function
    End If

    WzQuickChatButton_getTip = W_BUTTON_(self).pText
End Function

Public Sub WzQuickChatButton_simulateClick(ByVal self As WIDGET)
    Dim context As W_CONTEXT    ' W_CONTEXT::ZeroContext()
    context.mx = self.screenPosX()
    context.my = self.screenPosY()

    self.clicked context, WKEY_PRIMARY
    self.released context, WKEY_PRIMARY
End Sub

' MARK: - WzQuickChatToggleMixedCheckButton

' (currentChecked has no initialiser in C++: 0, Unchecked, until updateDisplay sets it)
Public Sub WzQuickChatToggleMixedCheckButton_construct(ByVal self As WIDGET)
    W_BUTTON_construct self
    self.kind = K_WzQuickChatToggleMixedCheckButton ' (VB) virtual calls from here on reach WzQuickChatToggleMixedCheckButton's overrides
End Sub

Public Function WzQuickChatToggleMixedCheckButton_make(ByVal text As String) As WIDGET
    Dim button As WIDGET, minButtonWidthForText As Long
    Set button = WIDGET_alloc(K_WzQuickChatToggleMixedCheckButton)
    WzQuickChatToggleMixedCheckButton_construct button
    button.setString text
    W_BUTTON_(button).FontID = font_regular
    minButtonWidthForText = iV_GetTextWidth(text, W_BUTTON_(button).FontID)
    button.setGeometry 0, 0, minButtonWidthForText + QC_PLAYER_BUTTON_PADDING, QC_PLAYER_BUTTON_HEIGHT
    Set WzQuickChatToggleMixedCheckButton_make = button
End Function

Public Sub WzQuickChatToggleMixedCheckButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As QCToggleMixedCheckButtonData, x0 As Long, y0 As Long, x1 As Long, y1 As Long, isDown As Boolean
    Dim isDisabled As Boolean, isHighlight As Boolean, isChecked As Boolean, light_border As Long, fill_color As Long
    Dim border As Long, fw As Long, fx As Long, fy As Long, textColor As Long
    Set d = WzQuickChatToggleMixedCheckButton_(self)
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    x1 = x0 + self.width()
    y1 = y0 + self.height()

    isDown = (self.getState() And (WBUT_DOWN Or WBUT_LOCK Or WBUT_CLICKLOCK)) <> 0
    isDisabled = (self.getState() And WBUT_DISABLE) <> 0
    isHighlight = Not isDisabled And ((self.getState() And WBUT_HIGHLIGHT) <> 0)
    isChecked = d.currentChecked = MixedCheckState_Checked

    ' Display the button.
    light_border = WZCOL_QUICKCHAT_TABS_BORDER_LIGHT
    If isDown Or isDisabled Then fill_color = WZCOL_QUICKCHAT_TABS_FILL_COLOR_DARK Else fill_color = WZCOL_QUICKCHAT_TABS_FILL_COLOR
    If isDown Or isHighlight Or isChecked Then border = light_border Else border = psPalette(WZCOL_FORM_DARK)
    iV_ShadowBox x0, y0, x1, y1, 0, border, border, fill_color

    d.wzText.setText W_BUTTON_(self).pText, W_BUTTON_(self).FontID
    Select Case d.currentChecked
    Case MixedCheckState_Unchecked
        d.wzStatusText.setText "", font_regular
    Case MixedCheckState_Mixed
        d.wzStatusText.setText "-", font_regular
    Case MixedCheckState_Checked
        d.wzStatusText.setText ChrW$(&H2713), font_regular    ' "(U+2713: the check mark)"
    End Select

    fw = d.wzText.width()
    fx = x0 + (self.width() - fw) \ 2
    fy = y0 + (self.height() - d.wzText.lineSize()) \ 2 - d.wzText.aboveBase()
    If isDown Or isChecked Then textColor = psPalette(WZCOL_TEXT_BRIGHT) Else textColor = psPalette(WZCOL_FORM_TEXT)
    If isDisabled Then
        textColor = pielightWithAlpha(textColor, pielightAlpha(textColor) \ 2)
    End If
    d.wzText.render fx, fy, textColor

    d.wzStatusText.render x0 + 5, fy, textColor
End Sub

Public Sub WzQuickChatToggleMixedCheckButton_setCheckedState(ByVal self As WIDGET, ByVal newState As WzQuickChatToggleMixedCheckButton_MixedCheckState)
    WzQuickChatToggleMixedCheckButton_(self).currentChecked = newState
End Sub

Public Function WzQuickChatToggleMixedCheckButton_getCheckedState(ByVal self As WIDGET) As WzQuickChatToggleMixedCheckButton_MixedCheckState
    WzQuickChatToggleMixedCheckButton_getCheckedState = WzQuickChatToggleMixedCheckButton_(self).currentChecked
End Function

Public Sub WzQuickChatToggleMixedCheckButton_toggle(ByVal self As WIDGET)
    Dim d As QCToggleMixedCheckButtonData
    Set d = WzQuickChatToggleMixedCheckButton_(self)
    Select Case d.currentChecked
    Case MixedCheckState_Unchecked
        d.currentChecked = MixedCheckState_Checked
    Case MixedCheckState_Mixed
        d.currentChecked = MixedCheckState_Checked
    Case MixedCheckState_Checked
        d.currentChecked = MixedCheckState_Unchecked
    End Select
End Sub

Public Sub WzQuickChatToggleMixedCheckButton_clicked(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    W_BUTTON_clicked self, psContext, key

    If (W_BUTTON_(self).state And WBUT_DOWN) <> 0 Then
        WzQuickChatToggleMixedCheckButton_toggle self
    End If
End Sub

' MARK: - WzQuickChatPlayerButton

Public Sub WzQuickChatPlayerButton_construct(ByVal self As WIDGET)
    W_BUTTON_construct self
    self.kind = K_WzQuickChatPlayerButton ' (VB) virtual calls from here on reach WzQuickChatPlayerButton's overrides
End Sub

Public Function WzQuickChatPlayerButton_make(ByVal playerIdx As Long) As WIDGET
    Dim button As WIDGET, d As WzQuickChatPlayerButton, playerButtonText As String, disabledOption As Boolean
    Dim minButtonWidthForText As Long
    Set button = WIDGET_alloc(K_WzQuickChatPlayerButton)
    WzQuickChatPlayerButton_construct button
    Set d = WzQuickChatPlayerButton_(button)
    d.playerIdx = playerIdx
    d.playerPos_ = NetPlay_players_position(playerIdx)
    playerButtonText = ""
    If Not ULT(10, d.playerPos_) Then    ' (uint32_t) playerPos <= 10
        playerButtonText = CStr(d.playerPos_) & ": "
    End If
    playerButtonText = playerButtonText & getPlayerName(playerIdx)
    button.setString playerButtonText
    W_BUTTON_(button).FontID = font_regular
    disabledOption = (NetPlay_players_isSpectator(playerIdx) <> NetPlay_players_isSpectator(realSelectedPlayer))
    If disabledOption Then
        button.setState WBUT_DISABLE
    End If
    minButtonWidthForText = iV_GetTextWidth(playerButtonText, W_BUTTON_(button).FontID)
    button.setGeometry 0, 0, minButtonWidthForText + QC_PLAYER_BUTTON_PADDING, QC_PLAYER_BUTTON_HEIGHT
    Set WzQuickChatPlayerButton_make = button
End Function

Public Sub WzQuickChatPlayerButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As WzQuickChatPlayerButton, x0 As Long, y0 As Long, x1 As Long, y1 As Long, haveText As Boolean, isDown As Boolean
    Dim isDisabled As Boolean, isHighlight As Boolean, light_border As Long, fill_color As Long, border As Long
    Dim fw As Long, fx As Long, fy As Long, textColor As Long
    Set d = WzQuickChatPlayerButton_(self)
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    x1 = x0 + self.width()
    y1 = y0 + self.height()

    haveText = Len(W_BUTTON_(self).pText) <> 0

    isDown = (self.getState() And (WBUT_DOWN Or WBUT_LOCK Or WBUT_CLICKLOCK)) <> 0
    isDisabled = (self.getState() And WBUT_DISABLE) <> 0
    isHighlight = Not isDisabled And ((self.getState() And WBUT_HIGHLIGHT) <> 0)

    ' Display the button.
    light_border = WZCOL_QUICKCHAT_TABS_BORDER_LIGHT
    If isDown Or isDisabled Then fill_color = WZCOL_QUICKCHAT_TABS_FILL_COLOR_DARK Else fill_color = WZCOL_QUICKCHAT_TABS_FILL_COLOR
    If isDown Or isHighlight Then border = light_border Else border = psPalette(WZCOL_FORM_DARK)
    iV_ShadowBox x0, y0, x1, y1, 0, border, border, fill_color

    If haveText Then
        d.wzText.setText W_BUTTON_(self).pText, W_BUTTON_(self).FontID
        fw = d.wzText.width()
        fx = x0 + (self.width() - fw) \ 2
        fy = y0 + (self.height() - d.wzText.lineSize()) \ 2 - d.wzText.aboveBase()
        If isDown Then textColor = psPalette(WZCOL_TEXT_BRIGHT) Else textColor = psPalette(WZCOL_FORM_TEXT)
        If isDisabled Then
            textColor = pielightWithAlpha(textColor, pielightAlpha(textColor) \ 2)
        End If
        d.wzText.render fx, fy, textColor
    End If
End Sub

Public Sub WzQuickChatPlayerButton_setIsChosen(ByVal self As WIDGET, ByVal chosen As Boolean)
    Dim b As W_BUTTON
    Set b = W_BUTTON_(self)
    If chosen Then
        b.state = b.state Or WBUT_CLICKLOCK
    Else
        b.state = b.state And Not WBUT_CLICKLOCK
    End If
End Sub

Public Sub WzQuickChatPlayerButton_toggle(ByVal self As WIDGET)
    Dim b As W_BUTTON
    Set b = W_BUTTON_(self)
    If (b.state And WBUT_CLICKLOCK) = 0 Then
        b.state = b.state Or WBUT_CLICKLOCK
    Else
        b.state = b.state And Not WBUT_CLICKLOCK
    End If
End Sub

Public Function WzQuickChatPlayerButton_isChosen(ByVal self As WIDGET) As Boolean
    WzQuickChatPlayerButton_isChosen = (W_BUTTON_(self).state And (WBUT_LOCK Or WBUT_CLICKLOCK)) <> 0
End Function

Public Function WzQuickChatPlayerButton_playerIndex(ByVal self As WIDGET) As Long
    WzQuickChatPlayerButton_playerIndex = WzQuickChatPlayerButton_(self).playerIdx
End Function

Public Function WzQuickChatPlayerButton_playerPosNumber(ByVal self As WIDGET) As Long
    WzQuickChatPlayerButton_playerPosNumber = WzQuickChatPlayerButton_(self).playerPos_
End Function

Public Sub WzQuickChatPlayerButton_clicked(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    W_BUTTON_clicked self, psContext, key

    If (W_BUTTON_(self).state And WBUT_DOWN) <> 0 Then
        WzQuickChatPlayerButton_toggle self
    End If
End Sub

' MARK: - WzQuickChatTargeting

Public Sub WzQuickChatTargeting_reset(t As WzQuickChatTargeting)
    t.all_ = False
    t.humanTeammates = False
    t.aiTeammates = False
    UInt32Vector_clear t.specificPlayers
End Sub

Public Function WzQuickChatTargeting_noTargets(t As WzQuickChatTargeting) As Boolean
    WzQuickChatTargeting_noTargets = Not t.all_ And Not t.humanTeammates And Not t.aiTeammates And t.specificPlayers.n = 0
End Function

Public Function WzQuickChatTargeting_targetAll() As WzQuickChatTargeting
    Dim targeting As WzQuickChatTargeting
    targeting.all_ = True
    WzQuickChatTargeting_targetAll = targeting
End Function

' targeting.specificPlayers.insert(player)
Public Sub WzQuickChatTargeting_insertSpecificPlayer(t As WzQuickChatTargeting, ByVal player As Long)
    UInt32Set_insert t.specificPlayers, player
End Sub

' MARK: - WzQuickChatSendToSelector

Public Sub WzQuickChatSendToSelector_construct(ByVal self As WIDGET)
    Dim d As WzQuickChatSendToSelector
    WIDGET_construct self
    self.kind = K_WzQuickChatSendToSelector ' (VB) virtual calls from here on reach WzQuickChatSendToSelector's overrides
    Set d = WzQuickChatSendToSelector_(self)
    d.teamOnly = False
    d.hasHighlight = False
    d.weakSerial = WeakWidget_register(self)    ' (VB) its std::weak_ptrs
End Sub

Public Sub WzQuickChatSendToSelector_destruct(ByVal self As WIDGET)
    WeakWidget_unregister WzQuickChatSendToSelector_(self).weakSerial
End Sub

Public Function WzQuickChatSendToSelector_make(ByVal teamOnly As Boolean, initialTargeting As WzQuickChatTargeting) As WIDGET
    Dim result As WIDGET
    Set result = WIDGET_alloc(K_WzQuickChatSendToSelector)
    WzQuickChatSendToSelector_construct result
    WzQuickChatSendToSelector_initialize result, teamOnly, initialTargeting
    Set WzQuickChatSendToSelector_make = result
End Function

' std::unordered_set<uint32_t> currentlySelectedPlayers() const
Public Function WzQuickChatSendToSelector_currentlySelectedPlayers(ByVal self As WIDGET) As UInt32Vector
    Dim d As WzQuickChatSendToSelector, result As UInt32Vector, playerIdx As Long
    Set d = WzQuickChatSendToSelector_(self)
    For playerIdx = 0 To MAX_CONNECTED_PLAYERS - 1
        If d.playerSelections(playerIdx) = PlayerSelectionState_Selected Then
            UInt32Set_insert result, playerIdx
        End If
    Next
    WzQuickChatSendToSelector_currentlySelectedPlayers = result
End Function

Private Sub WzQuickChatSendToSelector_updateTargeting(ByVal self As WIDGET)
    Dim d As WzQuickChatSendToSelector, specificPlayers As UInt32Vector, t As WzQuickChatTargeting
    Dim h As UInt32Vector, a As UInt32Vector, pts As UInt32Vector, i As Long
    Set d = WzQuickChatSendToSelector_(self)
    specificPlayers = WzQuickChatSendToSelector_currentlySelectedPlayers(self)
    t = d.targeting
    h = d.playerHumanTeammates
    a = d.playerAITeammates

    WzQuickChatTargeting_reset t

    t.humanTeammates = (h.n <> 0) And UInt32Vector_allIn(h, specificPlayers)
    t.aiTeammates = (a.n <> 0) And UInt32Vector_allIn(a, specificPlayers)

    If d.teamOnly Then
        t.all_ = False
    Else
        ' "global" targeting mode
        pts = d.playersToShow
        t.all_ = (pts.n <> 0) And UInt32Vector_allIn(pts, specificPlayers)
    End If

    If t.all_ Then
        UInt32Vector_clear specificPlayers
    End If
    If t.humanTeammates Then
        For i = 0 To h.n - 1
            UInt32Set_erase specificPlayers, h.v(i)
        Next
    End If
    If t.aiTeammates Then
        For i = 0 To a.n - 1
            UInt32Set_erase specificPlayers, a.v(i)
        Next
    End If
    t.specificPlayers = specificPlayers
    d.targeting = t
End Sub

' const WzQuickChatTargeting& currentTargeting() const (VB: a copy)
Public Function WzQuickChatSendToSelector_currentTargeting(ByVal self As WIDGET) As WzQuickChatTargeting
    WzQuickChatSendToSelector_currentTargeting = WzQuickChatSendToSelector_(self).targeting
End Function

Public Function WzQuickChatSendToSelector_getPlayerHumanTeammates(ByVal self As WIDGET) As UInt32Vector
    WzQuickChatSendToSelector_getPlayerHumanTeammates = WzQuickChatSendToSelector_(self).playerHumanTeammates
End Function

Public Function WzQuickChatSendToSelector_getPlayerAITeammates(ByVal self As WIDGET) As UInt32Vector
    WzQuickChatSendToSelector_getPlayerAITeammates = WzQuickChatSendToSelector_(self).playerAITeammates
End Function

Private Sub WzQuickChatSendToSelector_initialize(ByVal self As WIDGET, ByVal teamOnly_ As Boolean, initialTargeting As WzQuickChatTargeting)
    Dim d As WzQuickChatSendToSelector, teammates As UInt32Vector, h As UInt32Vector, a As UInt32Vector
    Dim pts As UInt32Vector, i As Long, playerIdx As Long
    Set d = WzQuickChatSendToSelector_(self)
    d.teamOnly = teamOnly_

    teammates = getPlayerTeammates(selectedPlayer, True, True)
    d.playerTeammates = teammates

    ' filter playerTeammates for humans and AIs
    For i = 0 To teammates.n - 1
        playerIdx = teammates.v(i)
        If NetPlay_players_allocated(playerIdx) And Not NetPlay_players_isSpectator(playerIdx) Then
            UInt32Vector_push_back h, playerIdx
        End If
    Next
    For i = 0 To teammates.n - 1
        playerIdx = teammates.v(i)
        If NetPlay_players_ai(playerIdx) >= 0 And Not NetPlay_players_isSpectator(playerIdx) Then
            UInt32Vector_push_back a, playerIdx
        End If
    Next
    d.playerHumanTeammates = h
    d.playerAITeammates = a

    If Not d.teamOnly Then
        For playerIdx = 0 To MAX_PLAYERS - 1
            If playerIdx = selectedPlayer Then
                GoTo continue_
            End If

            If isHumanPlayer(playerIdx) Or (NetPlay_players_ai(playerIdx) >= 0) Then
                ' is an active (connected) human player, or an AI
                UInt32Vector_push_back pts, playerIdx
            End If
continue_:
        Next
    Else
        pts = teammates
    End If

    ' Sort playersToShow by position
    UInt32Vector_sortByPlayerPosition pts
    d.playersToShow = pts

    For i = 0 To MAX_CONNECTED_PLAYERS - 1
        d.playerSelections(i) = PlayerSelectionState_NotAvailable
    Next

    For i = 0 To pts.n - 1
        d.playerSelections(pts.v(i)) = PlayerSelectionState_Unselected
    Next

    d.targeting = initialTargeting
    WzQuickChatSendToSelector_setPlayerSelectionStateFromTargeting self, initialTargeting

    Set d.playersSelectorWidget = WzQuickChatSendToPlayersSelector_make(pts, d.teamOnly And h.n <> 0, d.teamOnly And a.n <> 0, self)

'    updateTargeting();
    WzQuickChatSendToSelector_updateSelectedPlayersString self

    d.sendToText.setText "To:", font_regular
End Sub

Private Sub WzQuickChatSendToSelector_updatePlayerSelectionState(ByVal self As WIDGET, ByVal playerIdx As Long, ByVal selected As Boolean)
    Dim d As WzQuickChatSendToSelector
    Set d = WzQuickChatSendToSelector_(self)
    If d.playerSelections(playerIdx) = PlayerSelectionState_NotAvailable Then
        Exit Sub
    End If
    If selected Then d.playerSelections(playerIdx) = PlayerSelectionState_Selected Else d.playerSelections(playerIdx) = PlayerSelectionState_Unselected
    WzQuickChatSendToSelector_updateTargeting self
    WzQuickChatSendToSelector_updateSelectedPlayersString self
    WzQuickChatSendToPlayersSelector_updateDisplay d.playersSelectorWidget, self
End Sub

' void updateTargetingByCategory(optional<bool> all = nullopt, optional<bool> humanTeammates = nullopt,
' optional<bool> aiTeammates = nullopt)
Private Sub WzQuickChatSendToSelector_updateTargetingByCategory(ByVal self As WIDGET, ByVal has_all As Boolean, ByVal all_ As Boolean, ByVal has_humanTeammates As Boolean, ByVal humanTeammates As Boolean, ByVal has_aiTeammates As Boolean, ByVal aiTeammates As Boolean)
    Dim d As WzQuickChatSendToSelector, t As WzQuickChatTargeting
    Set d = WzQuickChatSendToSelector_(self)
    t = d.targeting
    If has_all Then
        t.all_ = all_
        UInt32Vector_clear t.specificPlayers
    End If
    If has_humanTeammates Then
        t.humanTeammates = humanTeammates
        UInt32Vector_clear t.specificPlayers
    End If
    If has_aiTeammates Then
        t.aiTeammates = aiTeammates
        UInt32Vector_clear t.specificPlayers
    End If
    d.targeting = t
    WzQuickChatSendToSelector_setPlayerSelectionStateFromTargeting self, t
    WzQuickChatSendToSelector_updateSelectedPlayersString self
    WzQuickChatSendToPlayersSelector_updateDisplay d.playersSelectorWidget, self
End Sub

Private Sub WzQuickChatSendToSelector_setPlayerSelectionStateFromTargeting(ByVal self As WIDGET, newTargeting As WzQuickChatTargeting)
    Dim d As WzQuickChatSendToSelector, i As Long, h As UInt32Vector, a As UInt32Vector, playerIdx As Long
    Set d = WzQuickChatSendToSelector_(self)
    For i = 0 To MAX_CONNECTED_PLAYERS - 1
        If d.playerSelections(i) = PlayerSelectionState_NotAvailable Then
            GoTo continue1
        End If
        d.playerSelections(i) = PlayerSelectionState_Unselected
continue1:
    Next

    If newTargeting.all_ Then
        For i = 0 To MAX_CONNECTED_PLAYERS - 1
            If d.playerSelections(i) = PlayerSelectionState_NotAvailable Then
                GoTo continue2
            End If
            d.playerSelections(i) = PlayerSelectionState_Selected
continue2:
        Next
        Exit Sub
    End If

    If newTargeting.humanTeammates Then
        h = d.playerHumanTeammates
        For i = 0 To h.n - 1
            playerIdx = h.v(i)
            If d.playerSelections(playerIdx) = PlayerSelectionState_NotAvailable Then
                GoTo continue3
            End If
            d.playerSelections(playerIdx) = PlayerSelectionState_Selected
continue3:
        Next
    End If

    If newTargeting.aiTeammates Then
        a = d.playerAITeammates
        For i = 0 To a.n - 1
            playerIdx = a.v(i)
            If d.playerSelections(playerIdx) = PlayerSelectionState_NotAvailable Then
                GoTo continue4
            End If
            d.playerSelections(playerIdx) = PlayerSelectionState_Selected
continue4:
        Next
    End If

    For i = 0 To newTargeting.specificPlayers.n - 1
        playerIdx = newTargeting.specificPlayers.v(i)
        If d.playerSelections(playerIdx) = PlayerSelectionState_NotAvailable Then
            GoTo continue5
        End If
        d.playerSelections(playerIdx) = PlayerSelectionState_Selected
continue5:
    Next
End Sub

Private Sub WzQuickChatSendToSelector_updateSelectedPlayersString(ByVal self As WIDGET)
    Dim d As WzQuickChatSendToSelector, t As WzQuickChatTargeting, selectedPlayersString As String
    Dim currentPlayersSet As UInt32Vector, v As UInt32Vector, i As Long, additionalPlayers As UInt32Vector
    Set d = WzQuickChatSendToSelector_(self)
    t = d.targeting
    If t.all_ Then
        d.cachedPlayerText.setText "All", font_regular
        Exit Sub
    End If

    selectedPlayersString = ""
    currentPlayersSet = WzQuickChatSendToSelector_currentlySelectedPlayers(self)

    If t.humanTeammates And t.aiTeammates Then
        selectedPlayersString = selectedPlayersString & "Team"
        v = d.playerTeammates
        For i = 0 To v.n - 1
            UInt32Set_erase currentPlayersSet, v.v(i)
        Next
    ElseIf t.humanTeammates Then
        If d.playerAITeammates.n = 0 Then
            selectedPlayersString = selectedPlayersString & "Team"
        Else
            selectedPlayersString = selectedPlayersString & "Human Teammates"
        End If
        v = d.playerHumanTeammates
        For i = 0 To v.n - 1
            UInt32Set_erase currentPlayersSet, v.v(i)
        Next
    ElseIf t.aiTeammates Then
        If d.playerHumanTeammates.n = 0 Then
            selectedPlayersString = selectedPlayersString & "Team"
        Else
            selectedPlayersString = selectedPlayersString & "Bot Teammates"
        End If
        v = d.playerAITeammates
        For i = 0 To v.n - 1
            UInt32Set_erase currentPlayersSet, v.v(i)
        Next
    End If

    ' additionalPlayers.insert(additionalPlayers.end(), currentPlayersSet.begin(), currentPlayersSet.end());
    additionalPlayers = currentPlayersSet
    UInt32Vector_sortByPlayerPosition additionalPlayers

    If Len(selectedPlayersString) = 0 And additionalPlayers.n = 1 Then
        selectedPlayersString = selectedPlayersString & getPlayerName(additionalPlayers.v(0))
    ElseIf additionalPlayers.n <> 0 Then
        If Len(selectedPlayersString) <> 0 Then
            selectedPlayersString = selectedPlayersString & " + "
        End If
        selectedPlayersString = selectedPlayersString & va("%u players", additionalPlayers.n)
    End If

    If Len(selectedPlayersString) = 0 Then
        selectedPlayersString = "Choose a recipient"
    End If

    d.cachedPlayerText.setText selectedPlayersString, font_regular
End Sub

Public Sub WzQuickChatSendToSelector_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As WzQuickChatSendToSelector, x0 As Long, y0 As Long, toTextX0 As Long, toTextY0 As Long, fx As Long, fy As Long
    Dim textColor As Long, t As WzQuickChatTargeting
    Set d = WzQuickChatSendToSelector_(self)
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset

    If Not d.overlayScreen Is Nothing Then
        pie_UniTransBoxFill x0, y0, x0 + self.width(), y0 + self.height(), WZCOL_QUICKCHAT_TABS_FILL_COLOR_DARK
    End If

    toTextX0 = x0
    toTextY0 = y0 + (self.height() - d.sendToText.lineSize()) \ 2 - d.sendToText.aboveBase()
    d.sendToText.render toTextX0, toTextY0, psPalette(WZCOL_TEXT_BRIGHT)

    fx = toTextX0 + d.sendToText.width() + 4
    fy = y0 + (self.height() - d.cachedPlayerText.lineSize()) \ 2 - d.cachedPlayerText.aboveBase()
    t = d.targeting
    If WzQuickChatTargeting_noTargets(t) Then textColor = psPalette(WZCOL_RED) Else textColor = psPalette(WZCOL_TEXT_MEDIUM)
    d.cachedPlayerText.render fx, fy, textColor

    If d.hasHighlight Then
        iV_Box x0, y0, x0 + self.width(), y0 + self.height(), WZCOL_QUICKCHAT_TABS_BORDER_LIGHT
    End If
End Sub

Public Function WzQuickChatSendToSelector_idealWidth(ByVal self As WIDGET) As Long
    WzQuickChatSendToSelector_idealWidth = 200 ' TODO: Calculate
End Function

Public Function WzQuickChatSendToSelector_idealHeight(ByVal self As WIDGET) As Long
    WzQuickChatSendToSelector_idealHeight = WzQuickChatSendToSelector_(self).cachedPlayerText.lineSize()
End Function

Public Sub WzQuickChatSendToSelector_highlight(ByVal self As WIDGET, psContext As W_CONTEXT)
    WzQuickChatSendToSelector_(self).hasHighlight = True
End Sub

Public Sub WzQuickChatSendToSelector_highlightLost(ByVal self As WIDGET)
    WzQuickChatSendToSelector_(self).hasHighlight = False
End Sub

Public Sub WzQuickChatSendToSelector_clicked(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    Dim weakThis As Long
    weakThis = WzQuickChatSendToSelector_(self).weakSerial
    widgScheduleTask WzFunc_new(FN_qcSendTo_clicked_task, weakThis)
End Sub

' MARK: - ChooserOverlay

' ChooserOverlay(std::function<void ()> onClickedFunc)
Public Sub ChooserOverlay_construct(ByVal self As WIDGET, ByVal onClickedFunc As WzFunc)
    WIDGET_construct self
    self.kind = K_ChooserOverlay ' (VB) virtual calls from here on reach ChooserOverlay's overrides
    Set ChooserOverlay_(self).onClickedFunc = onClickedFunc
    self.setCalcLayout WzFunc_new(FN_qcChooserOverlay_calcLayout)
End Sub

' std::make_shared<ChooserOverlay>(onClickedFunc)
Public Function ChooserOverlay_new(ByVal onClickedFunc As WzFunc) As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_ChooserOverlay)
    ChooserOverlay_construct w, onClickedFunc
    Set ChooserOverlay_new = w
End Function

Public Sub ChooserOverlay_clicked(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    WzFunc_call ChooserOverlay_(self).onClickedFunc
End Sub

Private Function WzQuickChatSendToSelector_calculateDropdownListScreenPosY(ByVal self As WIDGET) As Long
    Dim d As WzQuickChatSendToSelector, dropDownOverlayPosY As Long
    Set d = WzQuickChatSendToSelector_(self)
    dropDownOverlayPosY = self.screenPosY() + self.height() + 1
    If dropDownOverlayPosY + d.playersSelectorWidget.idealHeight() > screenHeight Then
        ' Positioning the dropdown below would cause it to appear partially or fully offscreen
        ' So, instead, position it above
        dropDownOverlayPosY = self.screenPosY() - (d.playersSelectorWidget.idealHeight() + 1)
        If dropDownOverlayPosY < 0 Then
            ' Well, this would be off-screen too...
            ' For now: Just position it at the top
            dropDownOverlayPosY = 0
        End If
    End If
    WzQuickChatSendToSelector_calculateDropdownListScreenPosY = dropDownOverlayPosY
End Function

Private Sub WzQuickChatSendToSelector_openChooser(ByVal self As WIDGET)
    Dim pWeakThis As Long
    If Not WzQuickChatSendToSelector_(self).overlayScreen Is Nothing Then Exit Sub
    pWeakThis = WzQuickChatSendToSelector_(self).weakSerial
    widgScheduleTask WzFunc_new(FN_qcSendTo_openChooser_task, pWeakThis)
End Sub

Private Sub WzQuickChatSendToSelector_closeChooser(ByVal self As WIDGET)
    Dim pWeakThis As Long
    pWeakThis = WzQuickChatSendToSelector_(self).weakSerial
    widgScheduleTask WzFunc_new(FN_qcSendTo_closeChooser_task, pWeakThis)
End Sub

Public Sub WzQuickChatSendToSelector_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    Dim d As WzQuickChatSendToSelector
    Set d = WzQuickChatSendToSelector_(self)
    If Not d.overlayScreen Is Nothing Then
        d.playersSelectorWidget.setGeometry self.screenPosX(), WzQuickChatSendToSelector_calculateDropdownListScreenPosY(self), self.width(), d.playersSelectorWidget.height()

        If keyPressed(KEY_ESC) Then
            inputLoseFocus  ' clear the input buffer.
            WzQuickChatSendToSelector_closeChooser self
        End If
    End If
End Sub

' MARK: - WzQuickChatSendToPlayersSelector

Public Sub WzQuickChatSendToPlayersSelector_construct(ByVal self As WIDGET)
    WIDGET_construct self
    self.kind = K_WzQuickChatSendToPlayersSelector ' (VB) virtual calls from here on reach WzQuickChatSendToPlayersSelector's overrides
End Sub

Public Function WzQuickChatSendToPlayersSelector_make(playersToShow As UInt32Vector, ByVal showHumanTeammatesButton As Boolean, ByVal showAITeammatesButton As Boolean, ByVal selectorParent As WIDGET) As WIDGET
    Dim result As WIDGET
    Set result = WIDGET_alloc(K_WzQuickChatSendToPlayersSelector)
    WzQuickChatSendToPlayersSelector_construct result
    WzQuickChatSendToPlayersSelector_initialize result, playersToShow, showHumanTeammatesButton, showAITeammatesButton, selectorParent
    Set WzQuickChatSendToPlayersSelector_make = result
End Function

Private Sub WzQuickChatSendToPlayersSelector_initialize(ByVal self As WIDGET, playersToShow As UInt32Vector, ByVal showHumanTeammatesButton As Boolean, ByVal showAITeammatesButton As Boolean, ByVal selectorParent As WIDGET)
    Dim d As QCSendToPlayersSelectorData, psWeakParent As Long, i As Long, playerIdx As Long, playerButton As WIDGET
    Dim maxPlayersListHeight As Long
    Set d = WzQuickChatSendToPlayersSelector_(self)
    psWeakParent = WzQuickChatSendToSelector_(selectorParent).weakSerial

    If showHumanTeammatesButton Then
        Set d.humanTeammates = WzQuickChatToggleMixedCheckButton_make("Human Teammates")
        W_BUTTON_addOnClickHandler d.humanTeammates, WzFunc_new(FN_qcPlayersSel_humanTeammates_onClick, psWeakParent)
        self.attach d.humanTeammates
        d.humanTeammates.setCalcLayout WzFunc_new(FN_qcPlayersSel_humanTeammates_calcLayout)
    End If

    If showAITeammatesButton Then
        Set d.aiTeammates = WzQuickChatToggleMixedCheckButton_make("Bot Teammates")
        W_BUTTON_addOnClickHandler d.aiTeammates, WzFunc_new(FN_qcPlayersSel_aiTeammates_onClick, psWeakParent)
        self.attach d.aiTeammates
        d.aiTeammates.setCalcLayout WzFunc_new(FN_qcPlayersSel_aiTeammates_calcLayout)
    End If

    Set d.playersList = ScrollableListWidget_make()
    For i = 0 To playersToShow.n - 1
        playerIdx = playersToShow.v(i)
        Set playerButton = WzQuickChatPlayerButton_make(playerIdx)
        W_BUTTON_addOnClickHandler playerButton, WzFunc_new(FN_qcPlayersSel_player_onClick, psWeakParent, playerIdx)
        ScrollableListWidget_addItem d.playersList, playerButton
        d.playerButtons.Add playerButton
    Next
    self.attach d.playersList

    ' set playersList max height
    maxPlayersListHeight = MIN_(100, playersToShow.n * QC_PLAYER_BUTTON_HEIGHT)
    d.playersList.setCalcLayout WzFunc_new(FN_qcPlayersSel_playersList_calcLayout, maxPlayersListHeight)

    WzQuickChatSendToPlayersSelector_updateDisplay self, selectorParent
End Sub

Public Sub WzQuickChatSendToPlayersSelector_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x0 As Long, y0 As Long, x1 As Long, y1 As Long, borderColor As Long
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    x1 = x0 + self.width()
    y1 = y0 + self.height()

    borderColor = pal_RGBA(0, 0, 0, 255)

    ' draw background + border
    iV_TransBoxFill x0, y0, x1, y1
    iV_Box x0, y0, x1, y1, borderColor
End Sub

Public Sub WzQuickChatSendToPlayersSelector_updateDisplay(ByVal self As WIDGET, ByVal selectorParent As WIDGET)
    Dim d As QCSendToPlayersSelectorData, targeting As WzQuickChatTargeting, selectedPlayers As UInt32Vector
    Dim possible As UInt32Vector, i As Long, playerButton As WIDGET
    Set d = WzQuickChatSendToPlayersSelector_(self)
    targeting = WzQuickChatSendToSelector_currentTargeting(selectorParent)
    selectedPlayers = WzQuickChatSendToSelector_currentlySelectedPlayers(selectorParent)
    If Not d.humanTeammates Is Nothing Then
        If targeting.humanTeammates Then
            WzQuickChatToggleMixedCheckButton_setCheckedState d.humanTeammates, MixedCheckState_Checked
        Else
            ' distinguish between mixed or unchecked states
            possible = WzQuickChatSendToSelector_getPlayerHumanTeammates(selectorParent)
            If UInt32Vector_anyIn(possible, selectedPlayers) Then
                WzQuickChatToggleMixedCheckButton_setCheckedState d.humanTeammates, MixedCheckState_Mixed
            Else
                WzQuickChatToggleMixedCheckButton_setCheckedState d.humanTeammates, MixedCheckState_Unchecked
            End If
        End If
    End If
    If Not d.aiTeammates Is Nothing Then
        If targeting.aiTeammates Then
            WzQuickChatToggleMixedCheckButton_setCheckedState d.aiTeammates, MixedCheckState_Checked
        Else
            ' distinguish between mixed or unchecked states
            possible = WzQuickChatSendToSelector_getPlayerAITeammates(selectorParent)
            If UInt32Vector_anyIn(possible, selectedPlayers) Then
                WzQuickChatToggleMixedCheckButton_setCheckedState d.aiTeammates, MixedCheckState_Mixed
            Else
                WzQuickChatToggleMixedCheckButton_setCheckedState d.aiTeammates, MixedCheckState_Unchecked
            End If
        End If
    End If

    For i = 1 To d.playerButtons.Count
        Set playerButton = d.playerButtons(i)
        If UInt32Set_count(selectedPlayers, WzQuickChatPlayerButton_playerIndex(playerButton)) > 0 Then
            WzQuickChatPlayerButton_setIsChosen playerButton, True
        Else
            WzQuickChatPlayerButton_setIsChosen playerButton, False
        End If
    Next
End Sub

Public Sub WzQuickChatSendToPlayersSelector_geometryChanged(ByVal self As WIDGET)
    Dim d As QCSendToPlayersSelectorData
    Set d = WzQuickChatSendToPlayersSelector_(self)
    If Not d.humanTeammates Is Nothing Then
        d.humanTeammates.callCalcLayout
    End If
    If Not d.aiTeammates Is Nothing Then
        d.aiTeammates.callCalcLayout
    End If
    d.playersList.callCalcLayout
End Sub

Public Function WzQuickChatSendToPlayersSelector_idealWidth(ByVal self As WIDGET) As Long
    WzQuickChatSendToPlayersSelector_idealWidth = 200 ' TODO: Calculate
End Function

Public Function WzQuickChatSendToPlayersSelector_idealHeight(ByVal self As WIDGET) As Long
    Dim d As QCSendToPlayersSelectorData, result As Long
    Set d = WzQuickChatSendToPlayersSelector_(self)
    result = d.playersList.height()
    If Not d.humanTeammates Is Nothing Then result = result + d.humanTeammates.idealHeight()
    If Not d.aiTeammates Is Nothing Then result = result + d.aiTeammates.idealHeight()
    WzQuickChatSendToPlayersSelector_idealHeight = result
End Function

' MARK: - WzQuickChatSection

Public Sub WzQuickChatSection_construct(ByVal self As WIDGET)
    Dim d As WzQuickChatSection
    W_CLICKFORM_construct self
    self.kind = K_WzQuickChatSection ' (VB) virtual calls from here on reach WzQuickChatSection's overrides
    Set d = WzQuickChatSection_(self)
    d.hasFocus = False
    d.hasMouseOver = False
    d.weakSerial = WeakWidget_register(self)    ' (VB) its std::weak_ptrs
End Sub

Public Sub WzQuickChatSection_destruct(ByVal self As WIDGET)
    WeakWidget_unregister WzQuickChatSection_(self).weakSerial
End Sub

Public Function WzQuickChatSection_make(ByVal sectionTitle As String) As WIDGET
    Dim result As WIDGET
    Set result = WIDGET_alloc(K_WzQuickChatSection)
    WzQuickChatSection_construct result
    WzQuickChatSection_initialize result, sectionTitle
    Set WzQuickChatSection_make = result
End Function

Public Sub WzQuickChatSection_displayRecursive(ByVal self As WIDGET, context As WidgetGraphicsContext)
    Dim d As WzQuickChatSection, x0 As Long, y0 As Long, lineSize As Long, aboveBase As Long, widgetDisplayIndex As Long
    Dim i As Long, lw As WIDGET, widget As WIDGET, widgetX0 As Long, widgetY0 As Long, widgetX1 As Long
    Dim overlayKeyX0 As Long, overlayKeyY0 As Long, numberStr As String, textX0 As Long, textY0 As Long
    Set d = WzQuickChatSection_(self)
    ' call parent displayRecursive
    WIDGET_displayRecursive self, context

    If d.hasFocus Then
        ' "over-draw" with keyboard shortcut identifiers for all of the messages (and a border to make it clear it's selected)

        x0 = WidgetGraphicsContext_getXOffset(context) + self.x()
        y0 = WidgetGraphicsContext_getYOffset(context) + self.y()

        iV_Box x0, y0, x0 + self.width(), y0 + self.height(), pal_RGBA(255, 255, 255, 150)

        lineSize = iV_GetTextLineSize(font_bar)
        aboveBase = iV_GetTextAboveBase(font_bar)

        widgetDisplayIndex = 0
        Set lw = ListTabWidget_listWidget(d.messagesList)
        For i = ListWidget_firstWidgetShownIndex(lw) To ListWidget_lastWidgetShownIndex(lw)
            Set widget = ListWidget_getWidgetAtIndex(lw, i)
            widgetX0 = widget.screenPosX()
            widgetY0 = widget.screenPosY()
            widgetX1 = widgetX0 + widget.width()

            overlayKeyX0 = widgetX1 - lineSize
            overlayKeyY0 = widgetY0
            pie_UniTransBoxFill overlayKeyX0, overlayKeyY0, overlayKeyX0 + lineSize, overlayKeyY0 + lineSize, pal_RGBA(0, 0, 255, 150)
            If widgetDisplayIndex < 9 Then numberStr = CStr(widgetDisplayIndex + 1) Else numberStr = "0"
            textX0 = overlayKeyX0 + (lineSize - iV_GetTextWidth(numberStr, font_bar)) \ 2
            textY0 = overlayKeyY0 - aboveBase
            iV_DrawText numberStr, textX0, textY0, font_bar

            widgetDisplayIndex = widgetDisplayIndex + 1
        Next
    End If
End Sub

Public Sub WzQuickChatSection_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    Dim d As WzQuickChatSection, mouseIsOverWidgetOrChildren As Boolean, psParent As WIDGET, hasKey As Boolean
    Dim keyPressForMessageButton As Long, displayedMessagesFirstIndex As Long, targetedMessageIndex As Long
    Dim lw As WIDGET, widget As WIDGET
    Set d = WzQuickChatSection_(self)
    If Not d.hasFocus Then
        ' the only thing we handle if we don't have focus is mousewheel events *if* mouse is over the widget (or a widget child)
        mouseIsOverWidgetOrChildren = d.hasMouseOver
        If Not mouseIsOverWidgetOrChildren Then
            If self.hitTest(psContext.mx, psContext.my) Then mouseIsOverWidgetOrChildren = isMouseOverScreen(self.screenPointer)
        End If
        If mouseIsOverWidgetOrChildren Then
            If mousePressed(MOUSE_WUP) Then
                If ListTabWidget_currentPage(d.messagesList) = 0 Then
                    ' could add a sound effect here
                ElseIf Not ListTabWidget_setCurrentPage(d.messagesList, ListTabWidget_currentPage(d.messagesList) - 1) Then
                    ' could add a sound effect here
                End If
                Exit Sub
            End If
            If mousePressed(MOUSE_WDN) Then
                If Not ListTabWidget_setCurrentPage(d.messagesList, ListTabWidget_currentPage(d.messagesList) + 1) Then
                    ' could add sound effect here
                End If
                Exit Sub
            End If
        End If
        Exit Sub ' nothing else to do
    End If

    ' handle keypresses when has focus
    If keyPressed(KEY_ESC) Then
        ' immediately return - parent should handle?
        Exit Sub
    End If

    If keyPressed(KEY_TAB) Then
        ' parent should handle switching
        Set psParent = self.parent()
        If WIDGET_isKindOf(psParent, K_WzQuickChatPanel) Then
            WzQuickChatPanel_handleTabPress psParent
        End If
        inputLoseFocus
        Exit Sub
    End If

    If keyPressed(KEY_MINUS) Or keyPressed(KEY_KP_MINUS) Or mousePressed(MOUSE_WUP) Then
        If ListTabWidget_currentPage(d.messagesList) = 0 Then
            ' could add a sound effect here
        ElseIf Not ListTabWidget_setCurrentPage(d.messagesList, ListTabWidget_currentPage(d.messagesList) - 1) Then
            ' could add a sound effect here
        End If
        inputLoseFocus
        Exit Sub
    End If
    If keyPressed(KEY_EQUALS) Or keyPressed(KEY_KP_PLUS) Or mousePressed(MOUSE_WDN) Then
        If Not ListTabWidget_setCurrentPage(d.messagesList, ListTabWidget_currentPage(d.messagesList) + 1) Then
            ' could add sound effect here
        End If
        inputLoseFocus
        Exit Sub
    End If

    hasKey = numberButtonPressed(keyPressForMessageButton)

    inputLoseFocus

    If hasKey Then
        Set lw = ListTabWidget_listWidget(d.messagesList)
        displayedMessagesFirstIndex = ListWidget_firstWidgetShownIndex(lw)
        targetedMessageIndex = displayedMessagesFirstIndex + keyPressForMessageButton
        If targetedMessageIndex <= ListWidget_lastWidgetShownIndex(lw) Then
            Set widget = ListWidget_getWidgetAtIndex(lw, targetedMessageIndex)
            If WIDGET_isKindOf(widget, K_WzQuickChatButton) Then
                WzQuickChatButton_simulateClick widget
            End If
        End If
    End If
End Sub

Public Sub WzQuickChatSection_highlight(ByVal self As WIDGET, psContext As W_CONTEXT)
    WzQuickChatSection_(self).hasMouseOver = True
End Sub

Public Sub WzQuickChatSection_highlightLost(ByVal self As WIDGET)
    WzQuickChatSection_(self).hasMouseOver = False
End Sub

Public Sub WzQuickChatSection_released(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    Dim clickAndReleaseOnThisButton As Boolean
    clickAndReleaseOnThisButton = ((W_CLICKFORM_(self).state And WBUT_DOWN) <> 0) ' relies on W_CLICKFORM handling to properly set WBUT_DOWN

    W_CLICKFORM_released self, psContext, key

    If Not clickAndReleaseOnThisButton Then
        Exit Sub ' do nothing
    End If

    If key = WKEY_PRIMARY Then
        WzQuickChatSection_clickPrimary self
    ElseIf key = WKEY_SECONDARY Then
        WzQuickChatSection_clickSecondary self, False
    End If
End Sub

Public Sub WzQuickChatSection_clickPrimary(ByVal self As WIDGET)
    Dim lockedScreen As W_SCREEN, psParent As WIDGET
    ' Tell the form that the panel has focus
    Set lockedScreen = self.screenPointer
    If Not lockedScreen Is Nothing Then
        WzQuickChatSection_(self).hasFocus = True
        lockedScreen.setFocus self
        Set psParent = self.parent()
        If WIDGET_isKindOf(psParent, K_WzQuickChatPanel) Then
            WzQuickChatPanel_handleSelfOrChildFocusChange psParent
        End If
    End If
End Sub

Public Function WzQuickChatSection_clickSecondary(ByVal self As WIDGET, ByVal synthesizedFromHold As Boolean) As Boolean
    WzQuickChatSection_clickSecondary = False
End Function

Public Sub WzQuickChatSection_focusLost(ByVal self As WIDGET)
    Dim psParent As WIDGET
    WzQuickChatSection_(self).hasFocus = False
    Set psParent = self.parent()
    If WIDGET_isKindOf(psParent, K_WzQuickChatPanel) Then
        WzQuickChatPanel_handleSelfOrChildFocusChange psParent
    End If
End Sub

Public Function WzQuickChatSection_idealWidth(ByVal self As WIDGET) As Long
    WzQuickChatSection_idealWidth = WzQuickChatSection_(self).messagesList.idealWidth()
End Function

Public Function WzQuickChatSection_idealHeight(ByVal self As WIDGET) As Long
    Dim d As WzQuickChatSection, displayedChildren As Long
    Set d = WzQuickChatSection_(self)
    displayedChildren = MIN_(10, ListTabWidget_childrenSize(d.messagesList))
    WzQuickChatSection_idealHeight = ListTabWidget_heightOfTabsLabel(d.messagesList) + (WzQuickChatButton_getIdealBaseButtonHeight() * displayedChildren) + (OBJ_GAP * (displayedChildren - 1))
End Function

Public Sub WzQuickChatSection_geometryChanged(ByVal self As WIDGET)
    Dim d As WzQuickChatSection
    Set d = WzQuickChatSection_(self)
    If Not d.messagesList Is Nothing Then
        d.messagesList.callCalcLayout
    End If
End Sub

Public Sub WzQuickChatSection_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    ' currently, no-op
End Sub

Private Sub WzQuickChatSection_initialize(ByVal self As WIDGET, ByVal sectionTitle As String)
    Dim d As WzQuickChatSection, titleStr As String
    Set d = WzQuickChatSection_(self)
    Set d.messagesList = IntListTabWidget_make(TabAlignment_RightAligned)
    self.attach d.messagesList
    ListTabWidget_setChildSpacing d.messagesList, OBJ_GAP, OBJ_GAP
    d.messagesList.setCalcLayout WzFunc_new(FN_qcSection_messagesList_calcLayout)
    titleStr = sectionTitle
    titleStr = titleStr & " "
    ListTabWidget_setTitle d.messagesList, titleStr
End Sub

' void addChatOption(WzQuickChatMessage message, ShouldEnableOptionFunc shouldEnableOption = nullptr)
Public Sub WzQuickChatSection_addChatOption(ByVal self As WIDGET, ByVal message As WzQuickChatMessage, Optional ByVal shouldEnableOption As WzFunc = Nothing)
    Dim enabledOption As Boolean, messageButton As WIDGET, psWeakSection As Long
    enabledOption = True
    If Not shouldEnableOption Is Nothing Then
        enabledOption = CBool(WzFunc_call(shouldEnableOption, message))
    End If
    Set messageButton = WzQuickChatButton_make(to_display_string_WzQuickChatMessage(message), Not enabledOption)
    psWeakSection = WzQuickChatSection_(self).weakSerial
    W_BUTTON_addOnClickHandler messageButton, WzFunc_new(FN_qcSection_chatOption_onClick, message, psWeakSection)
    ListTabWidget_addWidgetToLayout WzQuickChatSection_(self).messagesList, messageButton
End Sub

Public Sub WzQuickChatSection_addCheatOption(ByVal self As WIDGET, ByVal cheatCommand As String, Optional ByVal enabledOption As Boolean = True)
    Dim messageButton As WIDGET
    Set messageButton = WzQuickChatButton_make(cheatCommand, Not enabledOption)
    ListTabWidget_addWidgetToLayout WzQuickChatSection_(self).messagesList, messageButton
End Sub

' MARK: - WzQuickChatPanel

Public Sub WzQuickChatPanel_construct(ByVal self As WIDGET)
    Dim d As WzQuickChatPanel
    WIDGET_construct self
    self.kind = K_WzQuickChatPanel ' (VB) virtual calls from here on reach WzQuickChatPanel's overrides
    Set d = WzQuickChatPanel_(self)
    d.layoutMinSectionsPerRow = 1
    d.hasFocus = False
End Sub

' static std::shared_ptr<WzQuickChatPanel> make(const WzString& panelName, optional<WzQuickChatSendToMode> sendToMode)
Public Function WzQuickChatPanel_make(ByVal panelName As String, ByVal has_sendToMode As Boolean, Optional ByVal sendToMode As WzQuickChatPanel_WzQuickChatSendToMode = WzQuickChatSendToMode_GlobalSendOptions) As WIDGET
    Dim result As WIDGET
    Set result = WIDGET_alloc(K_WzQuickChatPanel)
    WzQuickChatPanel_construct result
    WzQuickChatPanel_(result).panelName = panelName
    WzQuickChatPanel_(result).has_sendToMode = has_sendToMode
    WzQuickChatPanel_(result).sendToMode = sendToMode
    Set WzQuickChatPanel_make = result
End Function

Public Sub WzQuickChatPanel_addSection(ByVal self As WIDGET, ByVal section As WIDGET, Optional ByVal row As Long = 0)
    Dim d As WzQuickChatPanel, rowSections As Collection
    Set d = WzQuickChatPanel_(self)
    If row >= d.sectionRows.Count Then
        ' sectionRows.resize(row + 1);
        Do While d.sectionRows.Count < row + 1
            d.sectionRows.Add New Collection
        Loop
    End If
    Set rowSections = d.sectionRows(row + 1)
    rowSections.Add section
    d.sections.Add section
    self.attach section
    WzQuickChatPanel_recalcLayout self
End Sub

Public Sub WzQuickChatPanel_showSendTo(ByVal self As WIDGET)
    Dim d As WzQuickChatPanel, initialTargeting As WzQuickChatTargeting
    Set d = WzQuickChatPanel_(self)
    If Not d.has_sendToMode Then
        If Not d.sendToWidget Is Nothing Then
            self.detach d.sendToWidget
            Set d.sendToWidget = Nothing
        End If
        Exit Sub
    End If
    Select Case d.sendToMode
    Case WzQuickChatSendToMode_GlobalSendOptions
        initialTargeting.all_ = True
    Case WzQuickChatSendToMode_TeamSendOptions
        initialTargeting = lastTeamChatTargeting
    End Select

    Set d.sendToWidget = WzQuickChatSendToSelector_make((d.sendToMode = WzQuickChatSendToMode_TeamSendOptions), initialTargeting)
    self.attach d.sendToWidget
End Sub

Public Function WzQuickChatPanel_getTargeting(ByVal self As WIDGET) As WzQuickChatTargeting
    Dim d As WzQuickChatPanel, initialTargeting As WzQuickChatTargeting
    Set d = WzQuickChatPanel_(self)
    If Not d.sendToWidget Is Nothing Then
        WzQuickChatPanel_getTargeting = WzQuickChatSendToSelector_currentTargeting(d.sendToWidget)
    Else
        If d.has_sendToMode Then
            Select Case d.sendToMode
            Case WzQuickChatSendToMode_GlobalSendOptions
                initialTargeting.all_ = True
            Case WzQuickChatSendToMode_TeamSendOptions
                initialTargeting = defaultTeamChatTargeting
            End Select
        End If
        WzQuickChatPanel_getTargeting = initialTargeting
    End If
End Function

Public Sub WzQuickChatPanel_setMinSectionsPerRowForLayout(ByVal self As WIDGET, ByVal minSectionsPerRow As Long)
    Dim d As WzQuickChatPanel
    Set d = WzQuickChatPanel_(self)
    If minSectionsPerRow = d.layoutMinSectionsPerRow Then
        Exit Sub
    End If
    d.layoutMinSectionsPerRow = MAX_(minSectionsPerRow, 1)
    WzQuickChatPanel_recalcLayout self
End Sub

Private Sub WzQuickChatPanel_recalcLayout(ByVal self As WIDGET)
    Dim d As WzQuickChatPanel, sectionsTopY0 As Long, r As Long, rowSections As Collection, layoutSectionsPerRow As Long
    Dim sectionWidth As Long, lastSectionX1 As Long, maxSectionIdealHeight As Long, i As Long, section As WIDGET
    Dim sectionX0 As Long
    Set d = WzQuickChatPanel_(self)
    ' Send-To widget is always on top, if set
    sectionsTopY0 = 0
    If Not d.sendToWidget Is Nothing Then
        d.sendToWidget.setGeometry 10, sectionsTopY0, d.sendToWidget.idealWidth(), d.sendToWidget.idealHeight()
        sectionsTopY0 = sectionsTopY0 + d.sendToWidget.y() + d.sendToWidget.height() + 10
    End If

    If d.sectionRows.Count = 0 Then
        Exit Sub
    End If

    ' Sections row contents are laid out left-to-right
    For r = 1 To d.sectionRows.Count
        Set rowSections = d.sectionRows(r)
        layoutSectionsPerRow = MAX_(d.layoutMinSectionsPerRow, rowSections.Count)
        sectionWidth = (self.width() - ((layoutSectionsPerRow + 1) * 10)) \ layoutSectionsPerRow
        lastSectionX1 = 0
        maxSectionIdealHeight = 0
        For i = 1 To rowSections.Count
            Set section = rowSections(i)
            sectionX0 = lastSectionX1 + 10
            section.setGeometry sectionX0, sectionsTopY0, sectionWidth, section.idealHeight()
            lastSectionX1 = sectionX0 + sectionWidth
            maxSectionIdealHeight = MAX_(maxSectionIdealHeight, section.idealHeight())
        Next

        sectionsTopY0 = sectionsTopY0 + maxSectionIdealHeight + 10
    Next
End Sub

Public Sub WzQuickChatPanel_takeFocus(ByVal self As WIDGET)
    Dim lockedScreen As W_SCREEN, psParent As WIDGET
    ' Tell the form that the panel has focus
    Set lockedScreen = self.screenPointer
    If Not lockedScreen Is Nothing Then
        lockedScreen.setFocus self
        WzQuickChatPanel_(self).hasFocus = True
        Set psParent = self.parent()
        If WIDGET_isKindOf(psParent, K_WzQuickChatForm) Then
            WzQuickChatForm_setIsHighlightedMode psParent, True
        End If
    End If
End Sub

Public Sub WzQuickChatPanel_clicked(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    WzQuickChatPanel_takeFocus self
End Sub

Public Sub WzQuickChatPanel_handleSelfOrChildFocusChange(ByVal self As WIDGET)
    Dim childHasFocus As Boolean, psParent As WIDGET, lockedScreen As W_SCREEN, newFocusWidget As WIDGET
    childHasFocus = False
    Set psParent = self.parent()
    If Not WIDGET_isKindOf(psParent, K_WzQuickChatForm) Then Set psParent = Nothing
    Set lockedScreen = self.screenPointer
    If Not psParent Is Nothing And Not lockedScreen Is Nothing Then
        Set newFocusWidget = lockedScreen.getWidgetWithFocus()
        If Not newFocusWidget Is Nothing Then
            childHasFocus = newFocusWidget.hasAncestor(self)
        End If
        WzQuickChatForm_setIsHighlightedMode psParent, childHasFocus
    End If
End Sub

Public Sub WzQuickChatPanel_focusLost(ByVal self As WIDGET)
    WzQuickChatPanel_(self).hasFocus = False
    WzQuickChatPanel_handleSelfOrChildFocusChange self
End Sub

Public Function WzQuickChatPanel_idealWidth(ByVal self As WIDGET) As Long
    Dim d As WzQuickChatPanel, maxIdealWidth As Long, r As Long, rowSections As Collection, maxSectionIdealWidthTotal As Long
    Dim i As Long, section As WIDGET
    Set d = WzQuickChatPanel_(self)
    maxIdealWidth = 0
    If Not d.sendToWidget Is Nothing Then
        maxIdealWidth = MAX_(maxIdealWidth, d.sendToWidget.idealWidth())
    End If
    For r = 1 To d.sectionRows.Count
        Set rowSections = d.sectionRows(r)
        maxSectionIdealWidthTotal = 10
        For i = 1 To rowSections.Count
            Set section = rowSections(i)
            maxSectionIdealWidthTotal = maxSectionIdealWidthTotal + section.idealWidth() + 10
        Next
        maxIdealWidth = MAX_(maxIdealWidth, maxSectionIdealWidthTotal)
    Next
    WzQuickChatPanel_idealWidth = maxIdealWidth
End Function

Public Function WzQuickChatPanel_idealHeight(ByVal self As WIDGET) As Long
    Dim d As WzQuickChatPanel, result As Long, r As Long, rowSections As Collection, maxSectionIdealHeight As Long
    Dim i As Long, section As WIDGET
    Set d = WzQuickChatPanel_(self)
    result = 10
    If Not d.sendToWidget Is Nothing Then
        result = result + d.sendToWidget.idealHeight() + 10
    End If
    For r = 1 To d.sectionRows.Count
        Set rowSections = d.sectionRows(r)
        maxSectionIdealHeight = 0
        For i = 1 To rowSections.Count
            Set section = rowSections(i)
            maxSectionIdealHeight = MAX_(maxSectionIdealHeight, section.idealHeight())
        Next
        result = result + maxSectionIdealHeight + 10
    Next
    WzQuickChatPanel_idealHeight = result
End Function

Public Sub WzQuickChatPanel_geometryChanged(ByVal self As WIDGET)
    WzQuickChatPanel_recalcLayout self
End Sub

Public Sub WzQuickChatPanel_displayRecursive(ByVal self As WIDGET, context As WidgetGraphicsContext)
    Dim d As WzQuickChatPanel, x0 As Long, y0 As Long, lineSize As Long, aboveBase As Long, i As Long, section As WIDGET
    Dim sectionX0 As Long, sectionY0 As Long, overlayKeyX0 As Long, overlayKeyX1 As Long, numberStr As String
    Dim textX0 As Long, textY0 As Long
    Set d = WzQuickChatPanel_(self)
    ' call parent displayRecursive
    WIDGET_displayRecursive self, context

    If d.hasFocus Then
        ' "over-draw" with keyboard shortcut identifiers for all of the sections

        x0 = WidgetGraphicsContext_getXOffset(context) + self.x()
        y0 = WidgetGraphicsContext_getYOffset(context) + self.y()

        lineSize = iV_GetTextLineSize(font_bar)
        aboveBase = iV_GetTextAboveBase(font_bar)

        For i = 0 To d.sections.Count - 1
            Set section = d.sections(i + 1)
            sectionX0 = section.x() + x0
            sectionY0 = section.y() + y0

            overlayKeyX0 = sectionX0 + (section.width() - lineSize) \ 2
            overlayKeyX1 = sectionY0
            pie_UniTransBoxFill overlayKeyX0, overlayKeyX1, overlayKeyX0 + lineSize, overlayKeyX1 + lineSize, pal_RGBA(0, 0, 255, 150)
            numberStr = CStr(i + 1)
            textX0 = overlayKeyX0 + (lineSize - iV_GetTextWidth(numberStr, font_bar)) \ 2
            textY0 = overlayKeyX1 - aboveBase
            iV_DrawText numberStr, textX0, textY0, font_bar
        Next
    End If
End Sub

Public Function WzQuickChatPanel_handleTabPress(ByVal self As WIDGET) As Boolean
    Dim d As WzQuickChatPanel, lockedScreen As W_SCREEN, widgetWithFocus As WIDGET, it As Long, i As Long
    Dim targetWidgetToTakeFocus As WIDGET, section As WIDGET
    Set d = WzQuickChatPanel_(self)
    ' find current section with focus
    Set lockedScreen = self.screenPointer
    If Not lockedScreen Is Nothing Then
        Set widgetWithFocus = lockedScreen.getWidgetWithFocus()
        If widgetWithFocus Is Nothing Then
            WzQuickChatPanel_handleTabPress = False
            Exit Function
        End If
        ' auto it = std::find(sections.begin(), sections.end(), widgetWithFocus); (VB: its 1-based item, 0 = end())
        it = 0
        For i = 1 To d.sections.Count
            Set section = d.sections(i)
            If section Is widgetWithFocus Then
                it = i
                Exit For
            End If
        Next
        If it = 0 Then
            WzQuickChatPanel_handleTabPress = False
            Exit Function
        End If
'        if (keyDown(KEY_LSHIFT) || keyDown(KEY_RSHIFT))
'        {
'            if (it != sections.begin())
'            {
'                --it;
'            }
'            else
'            {
'                it = sections.end() - 1;
'            }
'        }
'        else
'        {
            it = it + 1
            If it > d.sections.Count Then
                it = 1
            End If
'        }
        inputLoseFocus
        Set targetWidgetToTakeFocus = d.sections(it)
        widgScheduleTask WzFunc_new(FN_qcPanel_tabPress_task, targetWidgetToTakeFocus)
        WzQuickChatPanel_handleTabPress = True
        Exit Function
    End If
    WzQuickChatPanel_handleTabPress = False
End Function

Private Sub WzQuickChatPanel_sendQuickChat(ByVal self As WIDGET, ByVal message As WzQuickChatMessage, ByVal fromPlayer As Long)
    Dim psParent As WIDGET, d As WzQuickChatForm
    ' ::sendQuickChat(message, selectedPlayer, getTargeting());
    sendQuickChat message, selectedPlayer, WzQuickChatPanel_getTargeting(self)
    Set psParent = self.parent()
    If Not WIDGET_isKindOf(psParent, K_WzQuickChatForm) Then Set psParent = Nothing
    If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/hci/quickchat.cpp:1983", "sendQuickChat", "No parent?": Exit Sub
    Set d = WzQuickChatForm_(psParent)
    If Not d.onQuickChatSent Is Nothing Then
        WzFunc_call d.onQuickChatSent
    End If
End Sub

Public Sub WzQuickChatPanel_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    Dim d As WzQuickChatPanel, lockedScreen As W_SCREEN, widgetWithFocus As WIDGET, hasKey As Boolean
    Dim keyPressForSection As Long, section As WIDGET
    Set d = WzQuickChatPanel_(self)
    If keyPressed(KEY_ESC) Then
        Set lockedScreen = self.screenPointer
        If Not lockedScreen Is Nothing Then
            Set widgetWithFocus = lockedScreen.getWidgetWithFocus()
            If Not widgetWithFocus Is Nothing Then
                If Not widgetWithFocus Is self Then
                    If widgetWithFocus.hasAncestor(self) Then
                        ' the focused widget is a child widget - take focus
                        WzQuickChatPanel_takeFocus self
                        inputLoseFocus ' clear input buffer
                        Exit Sub
                    End If
                End If
            End If
        End If
        ' else, immediately return - parent should handle?
        Exit Sub
    End If

    If Not d.hasFocus Then
        Exit Sub ' nothing to do
    End If

    ' handle keypresses when has focus

    If keyPressed(KEY_TAB) Then
        If Not WzQuickChatPanel_handleTabPress(self) Then
            ' let the parent handle it
        End If
        Exit Sub
    End If

    hasKey = numberButtonPressed(keyPressForSection)

    inputLoseFocus

    If hasKey Then
        If keyPressForSection < d.sections.Count Then
            Set section = d.sections(keyPressForSection + 1)
            WzQuickChatSection_clickPrimary section
        End If
    End If
End Sub

' MARK: - WzQuickChatPanelTabs

' WzQuickChatPanelTabs(int value = -1) : MultichoiceWidget(value) { }
Public Sub WzQuickChatPanelTabs_construct(ByVal self As WIDGET, Optional ByVal value As Long = -1)
    MultichoiceWidget_construct self, value
    self.kind = K_WzQuickChatPanelTabs ' (VB) virtual calls from here on reach WzQuickChatPanelTabs's overrides
End Sub

' std::make_shared<WzQuickChatPanelTabs>(value)
Public Function WzQuickChatPanelTabs_new(Optional ByVal value As Long = -1) As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_WzQuickChatPanelTabs)
    WzQuickChatPanelTabs_construct w, value
    Set WzQuickChatPanelTabs_new = w
End Function

Public Sub WzQuickChatPanelTabs_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
End Sub

' MARK: - WzQuickChatForm

Public Sub WzQuickChatForm_construct(ByVal self As WIDGET)
    Dim d As WzQuickChatForm
    W_CLICKFORM_construct self
    self.kind = K_WzQuickChatForm ' (VB) virtual calls from here on reach WzQuickChatForm's overrides
    Set d = WzQuickChatForm_(self)
    d.showHighlighted = False
    d.weakSerial = WeakWidget_register(self)    ' (VB) its std::weak_ptrs
End Sub

Public Sub WzQuickChatForm_destruct(ByVal self As WIDGET)
    WeakWidget_unregister WzQuickChatForm_(self).weakSerial
End Sub

Public Sub WzQuickChatForm_released(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    Dim clickAndReleaseOnThisForm As Boolean
    clickAndReleaseOnThisForm = ((W_CLICKFORM_(self).state And WBUT_DOWN) <> 0) ' relies on W_CLICKFORM handling to properly set WBUT_DOWN

    W_CLICKFORM_released self, psContext, key

    If Not clickAndReleaseOnThisForm Then
        Exit Sub ' do nothing
    End If

    If key = WKEY_PRIMARY Then
        ' give the current panel focus
        WzQuickChatForm_panelTakeFocus self
    End If
End Sub

Public Sub WzQuickChatForm_panelTakeFocus(ByVal self As WIDGET)
    Dim d As WzQuickChatForm
    Set d = WzQuickChatForm_(self)
    ' give the current panel focus
    If Not d.currentlyAttachedPanel Is Nothing Then
        WzQuickChatPanel_takeFocus d.currentlyAttachedPanel
    End If
End Sub

Public Sub WzQuickChatForm_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x0 As Long, y0 As Long, x1 As Long, y1 As Long, borderColor As Long
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    x1 = x0 + self.width()
    y1 = y0 + self.height()

    If Not WzQuickChatForm_(self).showHighlighted Then borderColor = pal_RGBA(0, 0, 0, 255) Else borderColor = pal_RGBA(255, 255, 255, 100)

    ' draw background + border
    iV_TransBoxFill x0, y0, x1, y1
    iV_Box x0, y0, x1, y1, borderColor
End Sub

Public Function WzQuickChatForm_idealHeight(ByVal self As WIDGET) As Long
    Dim d As WzQuickChatForm, suggestedHeight As Long, maxPanelHeight As Long, i As Long, panel As WIDGET
    Set d = WzQuickChatForm_(self)
    suggestedHeight = QCFORM_INTERNAL_PADDING
    If Not d.panelSwitcher Is Nothing Then
        suggestedHeight = suggestedHeight + MAX_(d.panelSwitcher.idealHeight(), W_LABEL_requiredHeight(d.quickChatLabel)) + QCFORM_INTERNAL_PADDING
    Else
        suggestedHeight = suggestedHeight + W_LABEL_requiredHeight(d.quickChatLabel) + QCFORM_INTERNAL_PADDING
    End If
    maxPanelHeight = 0
    For i = 1 To d.panels.Count
        Set panel = d.panels(i)
        maxPanelHeight = MAX_(maxPanelHeight, panel.idealHeight())
    Next
    suggestedHeight = suggestedHeight + maxPanelHeight
    WzQuickChatForm_idealHeight = suggestedHeight
End Function

Public Sub WzQuickChatForm_geometryChanged(ByVal self As WIDGET)
    WzQuickChatForm_recalcLayout self
End Sub

Private Sub WzQuickChatForm_recalcLayout(ByVal self As WIDGET)
    Dim d As WzQuickChatForm, panelsTopY0 As Long, i As Long, panel As WIDGET
    Set d = WzQuickChatForm_(self)
    d.quickChatLabel.callCalcLayout
    panelsTopY0 = QCFORM_INTERNAL_PADDING + W_LABEL_requiredHeight(d.quickChatLabel) + QCFORM_INTERNAL_PADDING
    If Not d.panelSwitcher Is Nothing Then
        d.panelSwitcher.callCalcLayout
        panelsTopY0 = MAX_(panelsTopY0, d.panelSwitcher.y() + d.panelSwitcher.height() + QCFORM_INTERNAL_PADDING)
    End If
    For i = 1 To d.panels.Count
        Set panel = d.panels(i)
        panel.setGeometry 0, panelsTopY0, self.width(), self.height() - panelsTopY0
    Next
End Sub

' static std::shared_ptr<WzQuickChatForm> make(WzQuickChatContext context, optional<WzQuickChatMode> startingPanel = nullopt)
Public Function WzQuickChatForm_make(ByVal context As WzQuickChatContext, Optional ByVal has_startingPanel As Boolean = False, Optional ByVal startingPanel As WzQuickChatMode = WzQuickChatMode_Global) As WIDGET
    Dim result As WIDGET
    Set result = WIDGET_alloc(K_WzQuickChatForm)
    WzQuickChatForm_construct result
    WzQuickChatForm_initialize result, context, has_startingPanel, startingPanel
    Set WzQuickChatForm_make = result
End Function

Private Sub WzQuickChatForm_initialize(ByVal self As WIDGET, ByVal context As WzQuickChatContext, ByVal has_startingMode As Boolean, ByVal startingMode As WzQuickChatMode)
    Dim d As WzQuickChatForm, idxSelectedPanel As Long, i As Long, panel As WIDGET
    Set d = WzQuickChatForm_(self)
    Set d.quickChatLabel = W_LABEL_new()
    W_LABEL_setFont d.quickChatLabel, font_regular, psPalette(WZCOL_TEXT_MEDIUM)
    d.quickChatLabel.setString "Quick Chat:"
    self.attach d.quickChatLabel
    d.quickChatLabel.setGeometry QCFORM_INTERNAL_PADDING, QCFORM_INTERNAL_PADDING, W_LABEL_getMaxLineWidth(d.quickChatLabel), W_LABEL_requiredHeight(d.quickChatLabel)
    d.quickChatLabel.setTransparentToMouse True

    Select Case context
    Case WzQuickChatContext_Lobby
        WzQuickChatForm_initializeLobby self
    Case WzQuickChatContext_InGame
        WzQuickChatForm_initializeInGame self, has_startingMode, startingMode
    End Select
    If d.panels.Count > 1 Then
        ' Display a panel switcher
        Set d.panelSwitcher = WzQuickChatPanelTabs_new(0)
        self.attach d.panelSwitcher
        MultibuttonWidget_setButtonAlignment d.panelSwitcher, MultibuttonWidget_ButtonAlignment_CENTER_ALIGN
        idxSelectedPanel = 0
        For i = 0 To d.panels.Count - 1
            Set panel = d.panels(i + 1)
            MultibuttonWidget_addButton d.panelSwitcher, i, WzPanelTabButton_make(WzQuickChatPanel_(panel).panelName)
            If d.currentlyAttachedPanel Is panel Then
                idxSelectedPanel = i
            End If
        Next
        MultibuttonWidget_choose d.panelSwitcher, idxSelectedPanel
        MultibuttonWidget_addOnChooseHandler d.panelSwitcher, WzFunc_new(FN_qcForm_panelSwitcher_onChoose)
        d.panelSwitcher.setCalcLayout WzFunc_new(FN_qcForm_panelSwitcher_calcLayout)
    End If

    WzQuickChatForm_recalcLayout self
End Sub

Public Sub WzQuickChatForm_setIsHighlightedMode(ByVal self As WIDGET, ByVal childIsFocused As Boolean)
    Dim d As WzQuickChatForm
    Set d = WzQuickChatForm_(self)
    d.showHighlighted = childIsFocused
    If childIsFocused Then
        W_LABEL_setFontColour d.quickChatLabel, psPalette(WZCOL_TEXT_BRIGHT)
    Else
        W_LABEL_setFontColour d.quickChatLabel, psPalette(WZCOL_TEXT_MEDIUM)
    End If
End Sub

Private Sub WzQuickChatForm_initializeLobby(ByVal self As WIDGET)
    Dim d As WzQuickChatForm, panel As WIDGET, requests As WIDGET, notices As WIDGET, reactions As WIDGET
    Set d = WzQuickChatForm_(self)
    Set panel = WzQuickChatPanel_make("Lobby Chat", True, WzQuickChatSendToMode_GlobalSendOptions)

    Set requests = WzQuickChatSection_make("Requests")
    WzQuickChatSection_addChatOption requests, WzQuickChatMessage_LOBBY_DESIRE_TO_PLAY                    ' I want to play
    WzQuickChatSection_addChatOption requests, WzQuickChatMessage_LOBBY_DESIRE_TO_SPECTATE                ' I want to spectate
    WzQuickChatSection_addChatOption requests, WzQuickChatMessage_LOBBY_DESIRE_TO_SWITCH_TEAMS            ' I want to switch teams
    WzQuickChatSection_addChatOption requests, WzQuickChatMessage_LOBBY_REQUEST_CHECK_READY               ' Please check ready so we can start
    WzQuickChatSection_addChatOption requests, WzQuickChatMessage_REQUEST_PLEASE_WAIT                     ' Please wait
    WzQuickChatSection_addChatOption requests, WzQuickChatMessage_REQUEST_LETS_GO                         ' Let's go!
    WzQuickChatPanel_addSection panel, requests

    ' Notices
    Set notices = WzQuickChatSection_make("Notices")
    WzQuickChatSection_addChatOption notices, WzQuickChatMessage_NOTICE_ALMOST_READY                      ' Almost ready
    WzQuickChatSection_addChatOption notices, WzQuickChatMessage_NOTICE_READY                             ' Ready
    WzQuickChatSection_addChatOption notices, WzQuickChatMessage_NOTICE_BRB                               ' Be right back
    WzQuickChatSection_addChatOption notices, WzQuickChatMessage_NOTICE_I_AM_BACK                         ' I'm back
    WzQuickChatPanel_addSection panel, notices

    ' Reactions
    Set reactions = WzQuickChatSection_make("Reactions")
    WzQuickChatSection_addChatOption reactions, WzQuickChatMessage_REACTIONS_YES                          ' Yes
    WzQuickChatSection_addChatOption reactions, WzQuickChatMessage_REACTIONS_NO                           ' No
    WzQuickChatSection_addChatOption reactions, WzQuickChatMessage_REACTIONS_MAYBE                        ' Maybe
    WzQuickChatSection_addChatOption reactions, WzQuickChatMessage_REACTIONS_NOT_YET                      ' Not Yet
    WzQuickChatSection_addChatOption reactions, WzQuickChatMessage_REACTIONS_SOON                         ' Soon
    WzQuickChatSection_addChatOption reactions, WzQuickChatMessage_REACTIONS_THANK_YOU                    ' Thank you
    WzQuickChatSection_addChatOption reactions, WzQuickChatMessage_REACTIONS_NO_PROBLEM                   ' No problem
    WzQuickChatSection_addChatOption reactions, WzQuickChatMessage_LOBBY_GREETING_WAVE                    ' Hey everyone!
    WzQuickChatSection_addChatOption reactions, WzQuickChatMessage_LOBBY_GREETING_WELCOME                 ' Welcome!
    WzQuickChatSection_addChatOption reactions, WzQuickChatMessage_REACTIONS_SORRY_DO_NOT_UNDERSTAND_PLEASE_USE_QUICK_CHAT ' Sorry, I don't understand. Please use Quick Chat?
    WzQuickChatPanel_addSection panel, reactions

    d.panels.Add panel

    If d.currentlyAttachedPanel Is Nothing Then
        ' default to global, which is always available
        WzQuickChatForm_switchAttachedPanel self, panel
    End If
End Sub

Private Sub WzQuickChatForm_switchAttachedPanel(ByVal self As WIDGET, ByVal newPanel As WIDGET)
    Dim d As WzQuickChatForm
    Set d = WzQuickChatForm_(self)
    If Not d.currentlyAttachedPanel Is Nothing Then
        self.detach d.currentlyAttachedPanel
    End If
    Set d.currentlyAttachedPanel = newPanel
    self.attach newPanel
End Sub

Private Sub WzQuickChatForm_initializeInGame(ByVal self As WIDGET, ByVal has_startingMode As Boolean, ByVal startingMode As WzQuickChatMode)
    Dim d As WzQuickChatForm, firstVisiblePanel As WzQuickChatMode, playerTeammates As UInt32Vector, hasTeammates As Boolean
    Dim globalPanel As WIDGET, teamPanel As WIDGET, endGamePanel As WIDGET
    Set d = WzQuickChatForm_(self)
    ' startingMode.value_or(WzQuickChatMode::Global)
    If has_startingMode Then firstVisiblePanel = startingMode Else firstVisiblePanel = WzQuickChatMode_Global

    ' Initialize 1-3 panels
    playerTeammates = getPlayerTeammates(selectedPlayer, True, True)
    hasTeammates = playerTeammates.n <> 0

    ' Global Panel
    Set globalPanel = WzQuickChatForm_createInGameGlobalPanel(self)
    d.panels.Add globalPanel
    If firstVisiblePanel = WzQuickChatMode_Global Then
        WzQuickChatForm_switchAttachedPanel self, globalPanel
    End If

    ' Team Panel (only if we have teammates to send anything to!)
    If hasTeammates Then
        Set teamPanel = WzQuickChatForm_createInGameTeamPanel(self, playerTeammates)
        d.panels.Add teamPanel
        If firstVisiblePanel = WzQuickChatMode_Team Then
            WzQuickChatForm_switchAttachedPanel self, teamPanel
        End If
    End If

'    // Cheats Panel (only if debug mode is enabled)
'    const DebugInputManager& dbgInputManager = gInputManager.debugManager();
'    if (dbgInputManager.debugMappingsAllowed())
'    {
'        auto cheatsPanel = createInGameCheatPanel();
'        panels.push_back(cheatsPanel);
'        if (firstVisiblePanel == WzQuickChatMode::Cheats)
'        {
'            switchAttachedPanel(cheatsPanel);
'        }
'    }

    ' End-Game reactions panel
    Set endGamePanel = WzQuickChatForm_createInGameEndGamePanel(self)
    d.panels.Add endGamePanel
    If firstVisiblePanel = WzQuickChatMode_EndGame Then
        WzQuickChatForm_switchAttachedPanel self, endGamePanel
    End If

    If d.currentlyAttachedPanel Is Nothing Then
        ' default to global, which is always available
        WzQuickChatForm_switchAttachedPanel self, globalPanel
    End If
End Sub

Private Function WzQuickChatForm_createInGameGlobalPanel(ByVal self As WIDGET) As WIDGET
    Dim panel As WIDGET, requests As WIDGET, notices As WIDGET, taunts As WIDGET, reactions As WIDGET
    Set panel = WzQuickChatPanel_make("Global", True, WzQuickChatSendToMode_GlobalSendOptions)

    Set requests = WzQuickChatSection_make("Requests")
    WzQuickChatSection_addChatOption requests, WzQuickChatMessage_REQUEST_PLEASE_WAIT                     ' Please wait
    WzQuickChatSection_addChatOption requests, WzQuickChatMessage_REQUEST_LETS_GO                         ' Let's go!
    WzQuickChatPanel_addSection panel, requests

    ' Notices
    Set notices = WzQuickChatSection_make("Notices")
    WzQuickChatSection_addChatOption notices, WzQuickChatMessage_NOTICE_ALMOST_READY                      ' Almost ready
    WzQuickChatSection_addChatOption notices, WzQuickChatMessage_NOTICE_READY                             ' Ready
    WzQuickChatSection_addChatOption notices, WzQuickChatMessage_NOTICE_BRB                               ' Be right back
    WzQuickChatSection_addChatOption notices, WzQuickChatMessage_NOTICE_I_AM_BACK                         ' I'm back
    WzQuickChatPanel_addSection panel, notices

    ' Taunts
    Set taunts = WzQuickChatSection_make("Taunts")
    WzQuickChatSection_addChatOption taunts, WzQuickChatMessage_REACTIONS_TAUNT_GET_READY                 ' Get ready...
    WzQuickChatSection_addChatOption taunts, WzQuickChatMessage_REACTIONS_TAUNT_YOURE_GOING_TO_REGRET_THAT ' You're going to regret that
    WzQuickChatSection_addChatOption taunts, WzQuickChatMessage_REACTIONS_TAUNT_BARELY_A_SCRATCH          ' Barely a scratch!
    WzQuickChatPanel_addSection panel, taunts

    ' Reactions
    Set reactions = WzQuickChatSection_make("Reactions")
    WzQuickChatSection_addChatOption reactions, WzQuickChatMessage_REACTIONS_YES                          ' Yes
    WzQuickChatSection_addChatOption reactions, WzQuickChatMessage_REACTIONS_NO                           ' No
    WzQuickChatSection_addChatOption reactions, WzQuickChatMessage_REACTIONS_MAYBE                        ' Maybe
    WzQuickChatSection_addChatOption reactions, WzQuickChatMessage_REACTIONS_NOT_YET                      ' Not Yet
    WzQuickChatSection_addChatOption reactions, WzQuickChatMessage_REACTIONS_SOON                         ' Soon
    WzQuickChatSection_addChatOption reactions, WzQuickChatMessage_REACTIONS_THANK_YOU                    ' Thank you
    WzQuickChatSection_addChatOption reactions, WzQuickChatMessage_REACTIONS_NO_PROBLEM                   ' No problem
    WzQuickChatSection_addChatOption reactions, WzQuickChatMessage_REACTIONS_WELL_PLAYED                  ' Well-played
    WzQuickChatSection_addChatOption reactions, WzQuickChatMessage_REACTIONS_CENSORED                     ' @#%*!
    WzQuickChatSection_addChatOption reactions, WzQuickChatMessage_REACTIONS_SORRY_DO_NOT_UNDERSTAND_PLEASE_USE_QUICK_CHAT ' Sorry, I don't understand. Please use Quick Chat?
    WzQuickChatPanel_addSection panel, reactions

    Set WzQuickChatForm_createInGameGlobalPanel = panel
End Function

Private Function WzQuickChatForm_createInGameTeamPanel(ByVal self As WIDGET, playerTeammates As UInt32Vector) As WIDGET
    Dim panel As WIDGET, requests As WIDGET, notices As WIDGET, suggestions As WIDGET, reactions As WIDGET
    Set panel = WzQuickChatPanel_make("Team", True, WzQuickChatSendToMode_TeamSendOptions)

    ' Coordination
    Set requests = WzQuickChatSection_make("Coordination")
    WzQuickChatSection_addChatOption requests, WzQuickChatMessage_TEAM_REQUEST_ATTACK_NOW                 ' Attack now?
    WzQuickChatSection_addChatOption requests, WzQuickChatMessage_TEAM_REQUEST_GROUP_UP                   ' Group up
    WzQuickChatSection_addChatOption requests, WzQuickChatMessage_TEAM_REQUEST_SPLIT_UP                   ' Split up
    WzQuickChatSection_addChatOption requests, WzQuickChatMessage_TEAM_REQUEST_FOCUS_ATTACKS_WHERE_MARKED ' Focus attacks where I've marked
    WzQuickChatSection_addChatOption requests, WzQuickChatMessage_TEAM_REQUEST_RETREAT                    ' Retreat!
    WzQuickChatSection_addChatOption requests, WzQuickChatMessage_TEAM_REQUEST_NEED_HELP                  ' I need help!
    WzQuickChatSection_addChatOption requests, WzQuickChatMessage_REQUEST_PLEASE_WAIT                     ' Please wait
    WzQuickChatSection_addChatOption requests, WzQuickChatMessage_REQUEST_LETS_GO                         ' Let's go!
    WzQuickChatSection_addChatOption requests, WzQuickChatMessage_TEAM_REQUEST_LEFT_SIDE                  ' Left side
    WzQuickChatSection_addChatOption requests, WzQuickChatMessage_TEAM_REQUEST_RIGHT_SIDE                 ' Right side
    WzQuickChatSection_addChatOption requests, WzQuickChatMessage_TEAM_REQUEST_TRANSFER_UNITS             ' **COND** Can someone please transfer me some units?
    WzQuickChatSection_addChatOption requests, WzQuickChatMessage_TEAM_REQUEST_TRANSFER_REBUILD           ' **COND** Can someone please transfer me trucks so I can rebuild?
    WzQuickChatPanel_addSection panel, requests

    ' Notices
    Set notices = WzQuickChatSection_make("Notices")
    WzQuickChatSection_addChatOption notices, WzQuickChatMessage_NOTICE_ALMOST_READY                      ' Almost ready
    WzQuickChatSection_addChatOption notices, WzQuickChatMessage_NOTICE_READY                             ' Ready
    WzQuickChatSection_addChatOption notices, WzQuickChatMessage_NOTICE_BRB                               ' Be right back
    WzQuickChatSection_addChatOption notices, WzQuickChatMessage_NOTICE_I_AM_BACK                         ' I'm back
    WzQuickChatSection_addChatOption notices, WzQuickChatMessage_TEAM_NOTICE_ATTACKING_NOW                ' Attacking now
    WzQuickChatSection_addChatOption notices, WzQuickChatMessage_TEAM_NOTICE_ON_MY_WAY                    ' On my way!
    WzQuickChatSection_addChatOption notices, WzQuickChatMessage_TEAM_NOTICE_OPPONENTS_COMING             ' They're coming!
    WzQuickChatSection_addChatOption notices, WzQuickChatMessage_TEAM_NOTICE_BEING_ATTACKED               ' I'm being attacked!
    WzQuickChatSection_addChatOption notices, WzQuickChatMessage_TEAM_NOTICE_RUSHING_OILS                 ' I'm rushing oils
    WzQuickChatSection_addChatOption notices, WzQuickChatMessage_TEAM_NOTICE_OPPONENTS_RUSHING_OILS       ' They're rushing oils!
    WzQuickChatSection_addChatOption notices, WzQuickChatMessage_TEAM_NOTICE_DONT_HAVE_POWER              ' **COND** I don't have enough power
    WzQuickChatPanel_addSection panel, notices

    ' Suggestions
    Set suggestions = WzQuickChatSection_make("Suggestions")
    WzQuickChatSection_addChatOption suggestions, WzQuickChatMessage_TEAM_SUGGESTION_BUILD_MORE_UNITS     ' I suggest building more units
    WzQuickChatSection_addChatOption suggestions, WzQuickChatMessage_TEAM_SUGGESTION_BUILD_DIFFERENT_UNITS ' I suggest building different units
    WzQuickChatSection_addChatOption suggestions, WzQuickChatMessage_TEAM_SUGGESTION_CHECK_TEAM_STRATEGY  ' I suggest checking the team strategy table
    WzQuickChatSection_addChatOption suggestions, WzQuickChatMessage_TEAM_SUGGESTION_RESEARCH_DIFFERENT_TECH ' I suggest researching different tech
    WzQuickChatSection_addChatOption suggestions, WzQuickChatMessage_TEAM_SUGGESTION_KEEP_RESEARCH_BUSY   ' Be sure to keep research centers busy
    WzQuickChatSection_addChatOption suggestions, WzQuickChatMessage_TEAM_SUGGESTION_BUILD_ANTI_AIR       ' I suggest building anti-air
    WzQuickChatSection_addChatOption suggestions, WzQuickChatMessage_TEAM_SUGGESTION_REPAIR_UNITS         ' I suggest repairing your units
    WzQuickChatSection_addChatOption suggestions, WzQuickChatMessage_TEAM_SUGGESTION_BUILD_REPAIR_FACILITIES ' I suggest building repair facilities
    WzQuickChatSection_addChatOption suggestions, WzQuickChatMessage_TEAM_SUGGESTION_BUILD_POWER_GEN      ' I suggest building more power generators
    WzQuickChatSection_addChatOption suggestions, WzQuickChatMessage_TEAM_SUGGESTION_BUILD_CAPTURE_OILS   ' I suggest capturing oil resources
    WzQuickChatPanel_addSection panel, suggestions

    ' Reactions
    Set reactions = WzQuickChatSection_make("Reactions")
    WzQuickChatSection_addChatOption reactions, WzQuickChatMessage_REACTIONS_YES                          ' Yes
    WzQuickChatSection_addChatOption reactions, WzQuickChatMessage_REACTIONS_NO                           ' No
    WzQuickChatSection_addChatOption reactions, WzQuickChatMessage_REACTIONS_MAYBE                        ' Maybe
    WzQuickChatSection_addChatOption reactions, WzQuickChatMessage_REACTIONS_NOT_YET                      ' Not Yet
    WzQuickChatSection_addChatOption reactions, WzQuickChatMessage_REACTIONS_SOON                         ' Soon
    WzQuickChatSection_addChatOption reactions, WzQuickChatMessage_REACTIONS_THANK_YOU                    ' Thank you
    WzQuickChatSection_addChatOption reactions, WzQuickChatMessage_REACTIONS_NO_PROBLEM                   ' No problem
    WzQuickChatSection_addChatOption reactions, WzQuickChatMessage_REACTIONS_WELL_PLAYED                  ' Well-played
    WzQuickChatSection_addChatOption reactions, WzQuickChatMessage_REACTIONS_CENSORED                     ' @#%*!
    WzQuickChatSection_addChatOption reactions, WzQuickChatMessage_REACTIONS_SORRY_DO_NOT_UNDERSTAND_PLEASE_USE_QUICK_CHAT ' Sorry, I don't understand. Please use Quick Chat?
    WzQuickChatSection_addChatOption reactions, WzQuickChatMessage_REACTIONS_TEAM_THAT_DIDNT_GO_WELL      ' That didn't go well...
    WzQuickChatSection_addChatOption reactions, WzQuickChatMessage_REACTIONS_TEAM_I_HAVE_ANOTHER_PLAN     ' I have another plan
    WzQuickChatPanel_addSection panel, reactions

    If playerTeammates.n > 1 Then
        ' Show Send-To options
        WzQuickChatPanel_showSendTo panel
    End If

    Set WzQuickChatForm_createInGameTeamPanel = panel
End Function

Private Function WzQuickChatForm_createInGameCheatPanel(ByVal self As WIDGET) As WIDGET
    Dim panel As WIDGET
    Set panel = WzQuickChatPanel_make("Cheats", False)

    ' TODO: Implement

    Set WzQuickChatForm_createInGameCheatPanel = panel
End Function

Private Function WzQuickChatForm_createInGameEndGamePanel(ByVal self As WIDGET) As WIDGET
    Dim panel As WIDGET, endReactions As WIDGET
    Set panel = WzQuickChatPanel_make("End-Game", True, WzQuickChatSendToMode_GlobalSendOptions)

    ' End-Game Reactions
    Set endReactions = WzQuickChatSection_make("Reactions")
    WzQuickChatSection_addChatOption endReactions, WzQuickChatMessage_REACTIONS_ENDGAME_GOOD_GAME         ' Good game
    WzQuickChatSection_addChatOption endReactions, WzQuickChatMessage_REACTIONS_ENDGAME_I_GIVE_UP         ' I give up
    WzQuickChatSection_addChatOption endReactions, WzQuickChatMessage_REACTIONS_ENDGAME_SORRY_HAVE_TO_LEAVE ' Sorry, I have to leave
    WzQuickChatPanel_addSection panel, endReactions

    WzQuickChatPanel_setMinSectionsPerRowForLayout panel, 4
    Set WzQuickChatForm_createInGameEndGamePanel = panel
End Function

' MARK: -

' namespace WzQuickChatDataContexts

' - INTERNAL_ADMIN_ACTION_NOTICE
Public Function WzQuickChatDataContexts_INTERNAL_ADMIN_ACTION_NOTICE_constructMessageData(ByVal ctx As WzQuickChatDataContexts_INTERNAL_ADMIN_ACTION_NOTICE_Context, ByVal responsiblePlayerIdx As Long, ByVal targetPlayerIdx As Long) As WzQuickChatMessageData
    Dim result As WzQuickChatMessageData
    result.dataContext = ctx
    result.dataA = responsiblePlayerIdx
    result.dataB = targetPlayerIdx
    WzQuickChatDataContexts_INTERNAL_ADMIN_ACTION_NOTICE_constructMessageData = result
End Function

Private Function WzQuickChatDataContexts_INTERNAL_ADMIN_ACTION_NOTICE_to_output_string(messageData As WzQuickChatMessageData) As String
    Dim responsiblePlayerIdx As Long, targetPlayerIdx As Long, responsiblePlayerName As String, targetPlayerName As String
    Dim responsiblePlayerType As String
    responsiblePlayerIdx = messageData.dataA
    targetPlayerIdx = messageData.dataB

    If Not ULT(responsiblePlayerIdx, MAX_CONNECTED_PLAYERS) Then
        WzQuickChatDataContexts_INTERNAL_ADMIN_ACTION_NOTICE_to_output_string = ""
        Exit Function
    End If
    If Not ULT(targetPlayerIdx, MAX_CONNECTED_PLAYERS) Then
        WzQuickChatDataContexts_INTERNAL_ADMIN_ACTION_NOTICE_to_output_string = ""
        Exit Function
    End If

    responsiblePlayerName = getPlayerName(responsiblePlayerIdx)
    targetPlayerName = getPlayerName(targetPlayerIdx)

    responsiblePlayerType = "Player"
    If responsiblePlayerIdx = NetPlay_hostPlayer Then
        responsiblePlayerType = "Host"
    ElseIf NetPlay_players_isAdmin(responsiblePlayerIdx) Then
        responsiblePlayerType = "Admin"
    End If

    Select Case messageData.dataContext
    Case WzQuickChatDataContexts_INTERNAL_ADMIN_ACTION_NOTICE_Context_Invalid
        WzQuickChatDataContexts_INTERNAL_ADMIN_ACTION_NOTICE_to_output_string = ""
        Exit Function
    Case WzQuickChatDataContexts_INTERNAL_ADMIN_ACTION_NOTICE_Context_Team
        WzQuickChatDataContexts_INTERNAL_ADMIN_ACTION_NOTICE_to_output_string = va("%s (%s) changed team of player (%s) to: %d", responsiblePlayerType, responsiblePlayerName, targetPlayerName, NetPlay_players_team(targetPlayerIdx))
        Exit Function
    Case WzQuickChatDataContexts_INTERNAL_ADMIN_ACTION_NOTICE_Context_Position
        WzQuickChatDataContexts_INTERNAL_ADMIN_ACTION_NOTICE_to_output_string = va("%s (%s) changed position of player (%s) to: %d", responsiblePlayerType, responsiblePlayerName, targetPlayerName, NetPlay_players_position(targetPlayerIdx))
        Exit Function
    Case WzQuickChatDataContexts_INTERNAL_ADMIN_ACTION_NOTICE_Context_Color
        WzQuickChatDataContexts_INTERNAL_ADMIN_ACTION_NOTICE_to_output_string = va("%s (%s) changed color of player (%s) to: %s", responsiblePlayerType, responsiblePlayerName, targetPlayerName, getPlayerColourName(targetPlayerIdx))
        Exit Function
    Case WzQuickChatDataContexts_INTERNAL_ADMIN_ACTION_NOTICE_Context_Faction
        WzQuickChatDataContexts_INTERNAL_ADMIN_ACTION_NOTICE_to_output_string = va("%s (%s) changed faction of player (%s) to: %s", responsiblePlayerType, responsiblePlayerName, targetPlayerName, to_localized_string_FactionID(NetPlay_players_faction(targetPlayerIdx)))
        Exit Function
    End Select

    WzQuickChatDataContexts_INTERNAL_ADMIN_ACTION_NOTICE_to_output_string = "" ' Silence compiler warning
End Function

' - INTERNAL_LOCALIZED_LOBBY_NOTICE
Public Function WzQuickChatDataContexts_INTERNAL_LOCALIZED_LOBBY_NOTICE_constructMessageData(ByVal ctx As WzQuickChatDataContexts_INTERNAL_LOCALIZED_LOBBY_NOTICE_Context, ByVal targetPlayerIdx As Long, ByVal additionalData As Long) As WzQuickChatMessageData
    Dim result As WzQuickChatMessageData
    result.dataContext = ctx
    result.dataA = additionalData
    result.dataB = targetPlayerIdx
    WzQuickChatDataContexts_INTERNAL_LOCALIZED_LOBBY_NOTICE_constructMessageData = result
End Function

Private Function WzQuickChatDataContexts_INTERNAL_LOCALIZED_LOBBY_NOTICE_to_output_string(messageData As WzQuickChatMessageData) As String
    Dim additionalData As Long, targetPlayerIdx As Long, targetPlayerName As String, result As String
    additionalData = messageData.dataA
    targetPlayerIdx = messageData.dataB

    If Not ULT(targetPlayerIdx, MAX_CONNECTED_PLAYERS) Then
        WzQuickChatDataContexts_INTERNAL_LOCALIZED_LOBBY_NOTICE_to_output_string = ""
        Exit Function
    End If

    targetPlayerName = getPlayerName(targetPlayerIdx)

    Select Case messageData.dataContext
    Case WzQuickChatDataContexts_INTERNAL_LOCALIZED_LOBBY_NOTICE_Context_Invalid
        WzQuickChatDataContexts_INTERNAL_LOCALIZED_LOBBY_NOTICE_to_output_string = ""
        Exit Function
    Case WzQuickChatDataContexts_INTERNAL_LOCALIZED_LOBBY_NOTICE_Context_NotReadyKickWarning
        If targetPlayerIdx = selectedPlayer Then
            audio_PlayTrack ID_SOUND_ZOOM_ON_RADAR
            WzQuickChatDataContexts_INTERNAL_LOCALIZED_LOBBY_NOTICE_to_output_string = "NOTICE: If you don't check Ready soon, you will be kicked from the room"
            Exit Function
        Else
            ' Not intended for this player, but show a single message when 5 seconds remaining
            ' so other players know a kick is coming
            If additionalData = 5 And Not isBlindSimpleLobby(game.blindMode) Then
                WzQuickChatDataContexts_INTERNAL_LOCALIZED_LOBBY_NOTICE_to_output_string = va("Player will be kicked if they don't check Ready soon: %s", targetPlayerName)
                Exit Function
            End If
            WzQuickChatDataContexts_INTERNAL_LOCALIZED_LOBBY_NOTICE_to_output_string = ""
            Exit Function
        End If
    Case WzQuickChatDataContexts_INTERNAL_LOCALIZED_LOBBY_NOTICE_Context_NotReadyKicked
        WzQuickChatDataContexts_INTERNAL_LOCALIZED_LOBBY_NOTICE_to_output_string = va("Auto-kicking player (%s) because they waited too long to check Ready", targetPlayerName)
        Exit Function
    Case WzQuickChatDataContexts_INTERNAL_LOCALIZED_LOBBY_NOTICE_Context_PlayerShouldCheckReadyNotice
        audio_PlayTrack ID_SOUND_ZOOM_ON_RADAR
        If isBlindSimpleLobby(game.blindMode) Then
            result = "NOTICE: Please check Ready"
        Else
            result = "NOTICE: Please check Ready so the game can begin"
        End If
        If additionalData <> 0 Then     ' (uint32_t) additionalData > 0
            ' auto-not-ready-kick is enabled
            result = result & vbLf
            result = result & "Players who don't check Ready in time will be kicked."
        End If
        WzQuickChatDataContexts_INTERNAL_LOCALIZED_LOBBY_NOTICE_to_output_string = result
        Exit Function
    End Select

    WzQuickChatDataContexts_INTERNAL_LOCALIZED_LOBBY_NOTICE_to_output_string = "" ' Silence compiler warning
End Function

' MARK: - Public functions

Public Function quickChatMessageExpectsExtraData(ByVal msg As WzQuickChatMessage) As Boolean
    Select Case msg
    ' WZ-generated internal messages which require extra data
    Case WzQuickChatMessage_INTERNAL_ADMIN_ACTION_NOTICE, WzQuickChatMessage_INTERNAL_LOCALIZED_LOBBY_NOTICE
        quickChatMessageExpectsExtraData = True

    Case Else
        quickChatMessageExpectsExtraData = False
    End Select
End Function

Public Function to_output_sender(ByVal msg As WzQuickChatMessage, ByVal sender As Long) As Long
    ' Certain internal messages override the sender for display purposes
    Select Case msg
    ' WZ-generated internal messages - not for users to deliberately send
    Case WzQuickChatMessage_INTERNAL_ADMIN_ACTION_NOTICE
        ' override the "sender" to SYSTEM_MESSAGE type
        to_output_sender = SYSTEM_MESSAGE
    Case WzQuickChatMessage_INTERNAL_LOCALIZED_LOBBY_NOTICE
        to_output_sender = NOTIFY_MESSAGE

    Case Else
        to_output_sender = sender
    End Select
End Function

' std::string to_output_string(WzQuickChatMessage msg, const optional<WzQuickChatMessageData>& messageData)
' (messageData.value() of an empty optional would throw: the callers pass the data with the messages that expect it)
Public Function to_output_string(ByVal msg As WzQuickChatMessage, ByVal has_messageData As Boolean, messageData As WzQuickChatMessageData) As String
    Select Case msg
    ' These messages have a long form
    Case WzQuickChatMessage_TEAM_REQUEST_TRANSFER_UNITS
        to_output_string = "Can someone please transfer me some units?"
    Case WzQuickChatMessage_TEAM_REQUEST_TRANSFER_REBUILD
        to_output_string = "Can someone please transfer me trucks so I can rebuild?"
    Case WzQuickChatMessage_REACTIONS_SORRY_DO_NOT_UNDERSTAND_PLEASE_USE_QUICK_CHAT
        to_output_string = "Sorry, I don't understand. (Please use Quick Chat?)"

    ' All suggestions use the "I suggest: " prefix
    Case WzQuickChatMessage_TEAM_SUGGESTION_BUILD_MORE_UNITS
        ' TRANSLATORS: A suggestion to other player(s)
        to_output_string = "I suggest: Building more units"
    Case WzQuickChatMessage_TEAM_SUGGESTION_BUILD_DIFFERENT_UNITS
        ' TRANSLATORS: A suggestion to other player(s)
        to_output_string = "I suggest: Building different units"
    Case WzQuickChatMessage_TEAM_SUGGESTION_CHECK_TEAM_STRATEGY
        ' TRANSLATORS: A suggestion to other player(s)
        to_output_string = "I suggest: Checking the team strategy view"
    Case WzQuickChatMessage_TEAM_SUGGESTION_RESEARCH_DIFFERENT_TECH
        ' TRANSLATORS: A suggestion to other player(s)
        to_output_string = "I suggest: Researching different tech"
    Case WzQuickChatMessage_TEAM_SUGGESTION_KEEP_RESEARCH_BUSY
        ' TRANSLATORS: A suggestion to other player(s)
        to_output_string = "I suggest: Keeping research centers busy"
    Case WzQuickChatMessage_TEAM_SUGGESTION_BUILD_ANTI_AIR
        ' TRANSLATORS: A suggestion to other player(s)
        to_output_string = "I suggest: Building anti-air"
    Case WzQuickChatMessage_TEAM_SUGGESTION_REPAIR_UNITS
        ' TRANSLATORS: A suggestion to other player(s)
        to_output_string = "I suggest: Repairing your units"
    Case WzQuickChatMessage_TEAM_SUGGESTION_BUILD_REPAIR_FACILITIES
        ' TRANSLATORS: A suggestion to other player(s)
        to_output_string = "I suggest: Building repair facilities"
    Case WzQuickChatMessage_TEAM_SUGGESTION_BUILD_POWER_GEN
        ' TRANSLATORS: A suggestion to other player(s)
        to_output_string = "I suggest: Building power generators"
    Case WzQuickChatMessage_TEAM_SUGGESTION_BUILD_CAPTURE_OILS
        ' TRANSLATORS: A suggestion to other player(s)
        to_output_string = "I suggest: Capturing oil resources"

    ' WZ-generated internal messages - not for users to deliberately send
    Case WzQuickChatMessage_INTERNAL_ADMIN_ACTION_NOTICE
        to_output_string = WzQuickChatDataContexts_INTERNAL_ADMIN_ACTION_NOTICE_to_output_string(messageData)
    Case WzQuickChatMessage_INTERNAL_LOCALIZED_LOBBY_NOTICE
        to_output_string = WzQuickChatDataContexts_INTERNAL_LOCALIZED_LOBBY_NOTICE_to_output_string(messageData)

    Case Else
        to_output_string = to_display_string_WzQuickChatMessage(msg)
    End Select
End Function

' const char* to_display_string(WzQuickChatMessage msg)
Public Function to_display_string_WzQuickChatMessage(ByVal msg As WzQuickChatMessage) As String
    Select Case msg
    ' LOBBY ONLY
    Case WzQuickChatMessage_LOBBY_GREETING_WAVE
        to_display_string_WzQuickChatMessage = "Hey everyone!"
    Case WzQuickChatMessage_LOBBY_GREETING_WELCOME
        to_display_string_WzQuickChatMessage = "Welcome!"
    Case WzQuickChatMessage_LOBBY_DESIRE_TO_SPECTATE
        to_display_string_WzQuickChatMessage = "I want to spectate"
    Case WzQuickChatMessage_LOBBY_DESIRE_TO_PLAY
        to_display_string_WzQuickChatMessage = "I want to play"
    Case WzQuickChatMessage_LOBBY_DESIRE_TO_SWITCH_TEAMS
        to_display_string_WzQuickChatMessage = "I want to switch teams"
    Case WzQuickChatMessage_LOBBY_REQUEST_CHECK_READY
        to_display_string_WzQuickChatMessage = "Please check Ready so we can start"

    ' ALL (lobby + in-game global *and* team)

    Case WzQuickChatMessage_REQUEST_PLEASE_WAIT
        to_display_string_WzQuickChatMessage = "Please wait"
    Case WzQuickChatMessage_REQUEST_LETS_GO
        to_display_string_WzQuickChatMessage = "Let's go!"

    Case WzQuickChatMessage_NOTICE_ALMOST_READY
        ' TRANSLATORS: As in "I am almost ready" (but shorter, more informal)
        to_display_string_WzQuickChatMessage = "Almost ready"
    Case WzQuickChatMessage_NOTICE_READY
        ' TRANSLATORS: As in "I am ready" (but shorter, more informal)
        to_display_string_WzQuickChatMessage = "Ready"
    Case WzQuickChatMessage_NOTICE_BRB
        to_display_string_WzQuickChatMessage = "Be right back"
    Case WzQuickChatMessage_NOTICE_I_AM_BACK
        to_display_string_WzQuickChatMessage = "I'm back"

    Case WzQuickChatMessage_REACTIONS_YES
        to_display_string_WzQuickChatMessage = "Yes"
    Case WzQuickChatMessage_REACTIONS_NO
        to_display_string_WzQuickChatMessage = "No"
    Case WzQuickChatMessage_REACTIONS_MAYBE
        to_display_string_WzQuickChatMessage = "Maybe"
    Case WzQuickChatMessage_REACTIONS_NOT_YET
        to_display_string_WzQuickChatMessage = "Not yet"
    Case WzQuickChatMessage_REACTIONS_SOON
        to_display_string_WzQuickChatMessage = "Soon"
    Case WzQuickChatMessage_REACTIONS_THANK_YOU
        to_display_string_WzQuickChatMessage = "Thank you"
    Case WzQuickChatMessage_REACTIONS_NO_PROBLEM
        to_display_string_WzQuickChatMessage = "No problem"

    ' In-game reactions (global + team)
    Case WzQuickChatMessage_REACTIONS_WELL_PLAYED
        to_display_string_WzQuickChatMessage = "Well-played"
    Case WzQuickChatMessage_REACTIONS_CENSORED
        ' TRANSLATORS: Should probably be left as-is, unless there is a better way of denoting a censored outburst in your language
        to_display_string_WzQuickChatMessage = "@#%*!"
    Case WzQuickChatMessage_REACTIONS_SORRY_DO_NOT_UNDERSTAND_PLEASE_USE_QUICK_CHAT
        to_display_string_WzQuickChatMessage = "Sorry, I don't understand. (Use Quick Chat?)"

    ' Global reactions (taunts)
    Case WzQuickChatMessage_REACTIONS_TAUNT_GET_READY
        to_display_string_WzQuickChatMessage = "Get ready..."
    Case WzQuickChatMessage_REACTIONS_TAUNT_YOURE_GOING_TO_REGRET_THAT
        to_display_string_WzQuickChatMessage = "You're going to regret that"
    Case WzQuickChatMessage_REACTIONS_TAUNT_BARELY_A_SCRATCH
        to_display_string_WzQuickChatMessage = "Barely a scratch!"

    ' Team communication
    Case WzQuickChatMessage_TEAM_REQUEST_ATTACK_NOW
        to_display_string_WzQuickChatMessage = "Attack now?"
    Case WzQuickChatMessage_TEAM_REQUEST_GROUP_UP
        to_display_string_WzQuickChatMessage = "Group up"
    Case WzQuickChatMessage_TEAM_REQUEST_SPLIT_UP
        to_display_string_WzQuickChatMessage = "Split up"
    Case WzQuickChatMessage_TEAM_REQUEST_FOCUS_ATTACKS_WHERE_MARKED
        to_display_string_WzQuickChatMessage = "Focus attacks where marked"
    Case WzQuickChatMessage_TEAM_REQUEST_RETREAT
        to_display_string_WzQuickChatMessage = "Retreat!"
    Case WzQuickChatMessage_TEAM_REQUEST_NEED_HELP
        to_display_string_WzQuickChatMessage = "I need help!"
    Case WzQuickChatMessage_TEAM_REQUEST_LEFT_SIDE
        to_display_string_WzQuickChatMessage = "Left side"
    Case WzQuickChatMessage_TEAM_REQUEST_RIGHT_SIDE
        to_display_string_WzQuickChatMessage = "Right side"
    Case WzQuickChatMessage_TEAM_REQUEST_TRANSFER_UNITS
        to_display_string_WzQuickChatMessage = "Transfer Request: Units"
    Case WzQuickChatMessage_TEAM_REQUEST_TRANSFER_REBUILD
        to_display_string_WzQuickChatMessage = "Transfer Request: Trucks"

    Case WzQuickChatMessage_TEAM_NOTICE_ATTACKING_NOW
        ' TRANSLATORS: As in "I am attacking now!" (but shorter, more informal)
        to_display_string_WzQuickChatMessage = "Attacking now!"
    Case WzQuickChatMessage_TEAM_NOTICE_ON_MY_WAY
        ' TRANSLATORS: As in "I am on my way!" (but shorter, more informal)
        to_display_string_WzQuickChatMessage = "On my way!"
    Case WzQuickChatMessage_TEAM_NOTICE_OPPONENTS_COMING
        to_display_string_WzQuickChatMessage = "They're coming!"
    Case WzQuickChatMessage_TEAM_NOTICE_BEING_ATTACKED
        to_display_string_WzQuickChatMessage = "I'm being attacked!"
    Case WzQuickChatMessage_TEAM_NOTICE_RUSHING_OILS
        to_display_string_WzQuickChatMessage = "I'm rushing oils"
    Case WzQuickChatMessage_TEAM_NOTICE_OPPONENTS_RUSHING_OILS
        to_display_string_WzQuickChatMessage = "They're rushing oils!"
    Case WzQuickChatMessage_TEAM_NOTICE_DONT_HAVE_POWER
        to_display_string_WzQuickChatMessage = "I don't have enough power"

    Case WzQuickChatMessage_TEAM_SUGGESTION_BUILD_MORE_UNITS
        ' TRANSLATORS: A suggestion to other player(s), as in "Suggestion: %s"
        to_display_string_WzQuickChatMessage = "Build more units"
    Case WzQuickChatMessage_TEAM_SUGGESTION_BUILD_DIFFERENT_UNITS
        ' TRANSLATORS: A suggestion to other player(s), as in "Suggestion: %s"
        to_display_string_WzQuickChatMessage = "Build different units"
    Case WzQuickChatMessage_TEAM_SUGGESTION_CHECK_TEAM_STRATEGY
        ' TRANSLATORS: A suggestion to other player(s), as in "Suggestion: %s"
        to_display_string_WzQuickChatMessage = "Check team strategy view"
    Case WzQuickChatMessage_TEAM_SUGGESTION_RESEARCH_DIFFERENT_TECH
        ' TRANSLATORS: A suggestion to other player(s), as in "Suggestion: %s"
        to_display_string_WzQuickChatMessage = "Research different tech"
    Case WzQuickChatMessage_TEAM_SUGGESTION_KEEP_RESEARCH_BUSY
        ' TRANSLATORS: A suggestion to other player(s), as in "Suggestion: %s"
        to_display_string_WzQuickChatMessage = "Keep research centers busy"
    Case WzQuickChatMessage_TEAM_SUGGESTION_BUILD_ANTI_AIR
        ' TRANSLATORS: A suggestion to other player(s), as in "Suggestion: %s"
        to_display_string_WzQuickChatMessage = "Build anti-air"
    Case WzQuickChatMessage_TEAM_SUGGESTION_REPAIR_UNITS
        ' TRANSLATORS: A suggestion to other player(s), as in "Suggestion: %s"
        to_display_string_WzQuickChatMessage = "Repair your units"
    Case WzQuickChatMessage_TEAM_SUGGESTION_BUILD_REPAIR_FACILITIES
        ' TRANSLATORS: A suggestion to other player(s), as in "Suggestion: %s"
        to_display_string_WzQuickChatMessage = "Build repair facilities"
    Case WzQuickChatMessage_TEAM_SUGGESTION_BUILD_POWER_GEN
        ' TRANSLATORS: A suggestion to other player(s), as in "Suggestion: %s"
        to_display_string_WzQuickChatMessage = "Build power generators"
    Case WzQuickChatMessage_TEAM_SUGGESTION_BUILD_CAPTURE_OILS
        ' TRANSLATORS: A suggestion to other player(s), as in "Suggestion: %s"
        to_display_string_WzQuickChatMessage = "Capture oil resources"

    ' Team reactions
    Case WzQuickChatMessage_REACTIONS_TEAM_THAT_DIDNT_GO_WELL
        to_display_string_WzQuickChatMessage = "That didn't go well..."
    Case WzQuickChatMessage_REACTIONS_TEAM_I_HAVE_ANOTHER_PLAN
        to_display_string_WzQuickChatMessage = "I have another plan"

    ' End game reactions
    Case WzQuickChatMessage_REACTIONS_ENDGAME_GOOD_GAME
        to_display_string_WzQuickChatMessage = "Good game"
    Case WzQuickChatMessage_REACTIONS_ENDGAME_I_GIVE_UP
        to_display_string_WzQuickChatMessage = "I give up"
    Case WzQuickChatMessage_REACTIONS_ENDGAME_SORRY_HAVE_TO_LEAVE
        to_display_string_WzQuickChatMessage = "Sorry, I have to leave"

    ' WZ-generated internal messages - not for users to deliberately send
    Case WzQuickChatMessage_INTERNAL_MSG_DELIVERY_FAILURE_TRY_AGAIN
        to_display_string_WzQuickChatMessage = "Message delivery failure - try again"
    Case WzQuickChatMessage_INTERNAL_LOBBY_NOTICE_MAP_DOWNLOADED
        to_display_string_WzQuickChatMessage = "Map Downloaded"
    Case WzQuickChatMessage_INTERNAL_ADMIN_ACTION_NOTICE
        to_display_string_WzQuickChatMessage = "Admin modified a setting"
    Case WzQuickChatMessage_INTERNAL_LOCALIZED_LOBBY_NOTICE
        to_display_string_WzQuickChatMessage = ""

    ' not a valid message
    Case WzQuickChatMessage_MESSAGE_COUNT
        ASSERT_FAILURE "false", "src/hci/quickchat.cpp:2821", "to_display_string", "Invalid message"
        to_display_string_WzQuickChatMessage = ""

    Case Else
        to_display_string_WzQuickChatMessage = ""
    End Select
End Function

' std::shared_ptr<W_FORM> createQuickChatForm(WzQuickChatContext context, const std::function<void ()>& onQuickChatSent,
' optional<WzQuickChatMode> startingPanel = nullopt)
Public Function createQuickChatForm(ByVal context As WzQuickChatContext, ByVal onQuickChatSent As WzFunc, Optional ByVal has_startingPanel As Boolean = False, Optional ByVal startingPanel As WzQuickChatMode = WzQuickChatMode_Global) As WIDGET
    Dim result As WIDGET
    Set result = WzQuickChatForm_make(context, has_startingPanel, startingPanel)
    Set WzQuickChatForm_(result).onQuickChatSent = onQuickChatSent
    Set createQuickChatForm = result
End Function

' bool to_WzQuickChatMessage(uint32_t value, WzQuickChatMessage& output)
Public Function to_WzQuickChatMessage(ByVal value As Long, ByRef output As Long) As Boolean
    If ULT(value, WzQuickChatMessage_MESSAGE_COUNT) Then
        output = value
        to_WzQuickChatMessage = True
        Exit Function
    End If
    to_WzQuickChatMessage = False
End Function

' (VB) FOREACH_QUICKCHATMSG's #ENUM: the names of the WzQuickChatMessage values, in the enum's order
Private Function WzQuickChatMessage_names() As String()
    Dim s As String
    s = "LOBBY_GREETING_WAVE LOBBY_GREETING_WELCOME LOBBY_DESIRE_TO_SPECTATE LOBBY_DESIRE_TO_PLAY"
    s = s & " LOBBY_DESIRE_TO_SWITCH_TEAMS LOBBY_REQUEST_CHECK_READY REQUEST_PLEASE_WAIT REQUEST_LETS_GO NOTICE_ALMOST_READY"
    s = s & " NOTICE_READY NOTICE_BRB NOTICE_I_AM_BACK REACTIONS_YES REACTIONS_NO REACTIONS_MAYBE REACTIONS_NOT_YET"
    s = s & " REACTIONS_SOON REACTIONS_THANK_YOU REACTIONS_NO_PROBLEM REACTIONS_WELL_PLAYED REACTIONS_CENSORED"
    s = s & " REACTIONS_SORRY_DO_NOT_UNDERSTAND_PLEASE_USE_QUICK_CHAT REACTIONS_TAUNT_GET_READY"
    s = s & " REACTIONS_TAUNT_YOURE_GOING_TO_REGRET_THAT REACTIONS_TAUNT_BARELY_A_SCRATCH TEAM_REQUEST_ATTACK_NOW"
    s = s & " TEAM_REQUEST_GROUP_UP TEAM_REQUEST_SPLIT_UP TEAM_REQUEST_FOCUS_ATTACKS_WHERE_MARKED TEAM_REQUEST_RETREAT"
    s = s & " TEAM_REQUEST_NEED_HELP TEAM_REQUEST_LEFT_SIDE TEAM_REQUEST_RIGHT_SIDE TEAM_REQUEST_TRANSFER_UNITS"
    s = s & " TEAM_REQUEST_TRANSFER_REBUILD TEAM_NOTICE_ATTACKING_NOW TEAM_NOTICE_ON_MY_WAY TEAM_NOTICE_OPPONENTS_COMING"
    s = s & " TEAM_NOTICE_BEING_ATTACKED TEAM_NOTICE_RUSHING_OILS TEAM_NOTICE_OPPONENTS_RUSHING_OILS"
    s = s & " TEAM_NOTICE_DONT_HAVE_POWER TEAM_SUGGESTION_BUILD_MORE_UNITS TEAM_SUGGESTION_BUILD_DIFFERENT_UNITS"
    s = s & " TEAM_SUGGESTION_CHECK_TEAM_STRATEGY TEAM_SUGGESTION_RESEARCH_DIFFERENT_TECH TEAM_SUGGESTION_KEEP_RESEARCH_BUSY"
    s = s & " TEAM_SUGGESTION_BUILD_ANTI_AIR TEAM_SUGGESTION_REPAIR_UNITS TEAM_SUGGESTION_BUILD_REPAIR_FACILITIES"
    s = s & " TEAM_SUGGESTION_BUILD_POWER_GEN TEAM_SUGGESTION_BUILD_CAPTURE_OILS REACTIONS_TEAM_THAT_DIDNT_GO_WELL"
    s = s & " REACTIONS_TEAM_I_HAVE_ANOTHER_PLAN REACTIONS_ENDGAME_GOOD_GAME REACTIONS_ENDGAME_I_GIVE_UP"
    s = s & " REACTIONS_ENDGAME_SORRY_HAVE_TO_LEAVE INTERNAL_MSG_DELIVERY_FAILURE_TRY_AGAIN"
    s = s & " INTERNAL_LOBBY_NOTICE_MAP_DOWNLOADED INTERNAL_ADMIN_ACTION_NOTICE INTERNAL_LOCALIZED_LOBBY_NOTICE"
    WzQuickChatMessage_names = Split(s, " ")
End Function

' std::unordered_map<std::string, uint32_t> getWzQuickChatMessageValueMap() (VB: the JSON object that
' wzapi::getUsefulConstants makes of it; its keys are sorted, as nlohmann::json's)
Public Function getWzQuickChatMessageValueMap() As nlohmann_json
    Dim result As nlohmann_json, names() As String, i As Long
    Set result = json_object()
    ' #define QC_ENUM_ADD_TO_MAP(ENUM) result[#ENUM] = static_cast<uint32_t>(WzQuickChatMessage::ENUM);
    ' FOREACH_QUICKCHATMSG(QC_ENUM_ADD_TO_MAP)
    names = WzQuickChatMessage_names()
    For i = 0 To WzQuickChatMessage_MESSAGE_COUNT - 1
        result.set_item_ref names(i), json_value_unsigned(i)
    Next

    ' filter out anything starting with "INTERNAL_"
    For i = 0 To WzQuickChatMessage_MESSAGE_COUNT - 1
        If Left$(names(i), 9) = "INTERNAL_" Then
            result.erase_ names(i)
        End If
    Next

    Set getWzQuickChatMessageValueMap = result
End Function

Public Sub quickChatInitInGame()
    Dim blank As WzQuickChatTargeting, playerMessageTimes As Long
    If NetPlay_bComms Then
        ' for true multiplayer games, always default team targeting to human allies
        defaultTeamChatTargeting = blank
        defaultTeamChatTargeting.humanTeammates = True
    Else
        ' for local skirmish, always default team targeting to computer allies
        defaultTeamChatTargeting = blank
        defaultTeamChatTargeting.aiTeammates = True
    End If

    lastTeamChatTargeting = defaultTeamChatTargeting

    lastQuickChatMessageTimes_alloc
    For playerMessageTimes = 0 To MAX_CONNECTED_PLAYERS - 1
        lastQuickChatMessageTimes_clear playerMessageTimes
    Next
End Sub

Private Function playerHasHumanAndAITeammates(ByVal playerIdx As Long) As Boolean
    Dim playerTeammates As UInt32Vector, firstIsOrWasHuman As Boolean, thisIsOrWasHuman As Boolean, i As Long, p As Long
    If Not ULT(playerIdx, MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "playerIdx < MAX_CONNECTED_PLAYERS", "src/hci/quickchat.cpp:2892", "playerHasHumanAndAITeammates", "Invalid player: %u", playerIdx: playerHasHumanAndAITeammates = False: Exit Function
    playerTeammates = getPlayerTeammates(playerIdx, True, True)
    If playerTeammates.n = 0 Then
        playerHasHumanAndAITeammates = False
        Exit Function
    End If
    p = playerTeammates.v(0)
    firstIsOrWasHuman = isHumanPlayer(p) Or NetPlay_players_difficulty(p) = AIDifficulty_HUMAN ' is an active (connected) human player, or was a human player (probably disconnected)
    ' std::any_of(std::next(playerTeammates.begin()), playerTeammates.end(), ...)
    For i = 1 To playerTeammates.n - 1
        p = playerTeammates.v(i)
        thisIsOrWasHuman = isHumanPlayer(p) Or NetPlay_players_difficulty(p) = AIDifficulty_HUMAN ' is an active (connected) human player, or was a human player (probably disconnected)
        If thisIsOrWasHuman <> firstIsOrWasHuman Then
            playerHasHumanAndAITeammates = True
            Exit Function
        End If
    Next
    playerHasHumanAndAITeammates = False
End Function

' formatReceivers' lambda alliesStr ([&]: targeting, senderIdx): vbNullString = nullptr
Private Function formatReceivers_alliesStr(ByVal senderIdx As Long, targeting As WzQuickChatTargeting) As String
    If targeting.aiTeammates And targeting.humanTeammates Then
        formatReceivers_alliesStr = "Allies"
        Exit Function
    End If
    If targeting.aiTeammates Then
        ' only need to distinguish if the sending player has both human and AI teammates
        If playerHasHumanAndAITeammates(senderIdx) Then
            formatReceivers_alliesStr = "Bot Allies"
        Else
            formatReceivers_alliesStr = "Allies"
        End If
        Exit Function
    End If
    If targeting.humanTeammates Then
        ' only need to distinguish if the sending player has both human and AI teammates
        If playerHasHumanAndAITeammates(senderIdx) Then
            formatReceivers_alliesStr = "Human Allies"
        Else
            formatReceivers_alliesStr = "Allies"
        End If
        Exit Function
    End If
    formatReceivers_alliesStr = vbNullString
End Function

Private Function formatReceivers(ByVal senderIdx As Long, targeting As WzQuickChatTargeting) As String
    Dim toAllies As Boolean, pRetString As String, directs As Long, ss As String, nextName As String
    If targeting.all_ Then
        formatReceivers = "Global"
        Exit Function
    End If

    toAllies = (targeting.aiTeammates Or targeting.humanTeammates)

    If toAllies And targeting.specificPlayers.n = 0 Then
        pRetString = formatReceivers_alliesStr(senderIdx, targeting)
        If StrPtr(pRetString) <> 0 Then formatReceivers = pRetString Else formatReceivers = ""
        Exit Function
    End If

    If targeting.specificPlayers.n = 0 Then
        formatReceivers = ""
        Exit Function
    End If

    directs = 0     ' targeting.specificPlayers.begin()
    ss = ""
    If targeting.aiTeammates Or targeting.humanTeammates Then
        pRetString = formatReceivers_alliesStr(senderIdx, targeting)
        ss = ss & pRetString
        ss = ss & " & "
    Else
        ss = ss & "private to "
        ss = ss & getPlayerName(targeting.specificPlayers.v(directs))
        directs = directs + 1
    End If

    Do While directs <> targeting.specificPlayers.n
        nextName = getPlayerName(targeting.specificPlayers.v(directs))
        directs = directs + 1
        If StrPtr(nextName) = 0 Then
            GoTo continue_
        End If
        If directs = targeting.specificPlayers.n Then ss = ss & " and " Else ss = ss & ", "
        ss = ss & nextName
continue_:
    Loop

    formatReceivers = ss
End Function

Public Sub addQuickChatMessageToConsole(ByVal message As WzQuickChatMessage, ByVal sender As Long, targeting As WzQuickChatTargeting, ByVal has_messageData As Boolean, messageData As WzQuickChatMessageData)
    Dim formatted As String, teamSpecific As Boolean
    formatted = Left$(va("[%s] %s (%s): %s", formatLocalDateTime("%H:%M"), getPlayerName(sender), formatReceivers(sender, targeting), to_output_string(message, has_messageData, messageData)), MAX_CONSOLE_STRING_LENGTH - 1)   ' char formatted[MAX_CONSOLE_STRING_LENGTH]
    teamSpecific = Not targeting.all_ And (targeting.humanTeammates Or targeting.aiTeammates)
    addConsoleMessage formatted, DEFAULT_JUSTIFY, to_output_sender(message, sender), teamSpecific
End Sub

Public Sub addLobbyQuickChatMessageToConsole(ByVal message As WzQuickChatMessage, ByVal sender As Long, targeting As WzQuickChatTargeting, ByVal has_messageData As Boolean, messageData As WzQuickChatMessageData)
    Dim teamSpecific As Boolean, outputMsg As String
    teamSpecific = Not targeting.all_ And (targeting.humanTeammates Or targeting.aiTeammates)
    outputMsg = to_output_string(message, has_messageData, messageData)
    If Len(outputMsg) = 0 Then
        Exit Sub
    End If
    addConsoleMessage outputMsg, DEFAULT_JUSTIFY, to_output_sender(message, sender), teamSpecific
End Sub

Public Function shouldHideQuickChatMessageFromLocalDisplay(ByVal message As WzQuickChatMessage) As Boolean
    If Not isInternalMessage(message) Then
        shouldHideQuickChatMessageFromLocalDisplay = False
        Exit Function
    End If
    ' special cases for internal messages that should be displayed locally
    Select Case message
    Case WzQuickChatMessage_INTERNAL_ADMIN_ACTION_NOTICE
        shouldHideQuickChatMessageFromLocalDisplay = False
        Exit Function
    Case Else
    End Select

    ' hide local display of this message
    shouldHideQuickChatMessageFromLocalDisplay = True
End Function

' sendQuickChat's lambda isPotentiallyValidTarget ([&]: fromPlayer, isInGame, senderIsSpectator)
Private Function sendQuickChat_isPotentiallyValidTarget(ByVal playerIdx As Long, ByVal fromPlayer As Long, ByVal isInGame As Boolean, ByVal senderIsSpectator As Boolean) As Boolean
    If playerIdx = fromPlayer Then
        sendQuickChat_isPotentiallyValidTarget = False
        Exit Function
    End If

    If isInGame And senderIsSpectator And Not NetPlay_players_isSpectator(playerIdx) Then
        sendQuickChat_isPotentiallyValidTarget = False
        Exit Function
    End If

    If NetPlay_players_allocated(playerIdx) Then
        ' human player that's connected
        sendQuickChat_isPotentiallyValidTarget = True
    ElseIf isInGame And playerIdx < game.maxPlayers And NetPlay_players_ai(playerIdx) >= 0 And NetPlay_players_difficulty(playerIdx) <> AIDifficulty_DISABLED And Not NetPlay_players_isSpectator(playerIdx) Then
        ' AI player that's still in the game
        sendQuickChat_isPotentiallyValidTarget = True
    Else
        ' not a valid target player
        sendQuickChat_isPotentiallyValidTarget = False
    End If
End Function

' void sendQuickChat(WzQuickChatMessage message, uint32_t fromPlayer, WzQuickChatTargeting targeting,
' optional<WzQuickChatMessageData> messageData = nullopt): without the data
Public Sub sendQuickChat(ByVal message As WzQuickChatMessage, ByVal fromPlayer As Long, targeting As WzQuickChatTargeting)
    Dim nullopt As WzQuickChatMessageData
    sendQuickChat_impl message, fromPlayer, targeting, False, nullopt
End Sub

' (VB) sendQuickChat with the message data
Public Sub sendQuickChat_data(ByVal message As WzQuickChatMessage, ByVal fromPlayer As Long, targeting As WzQuickChatTargeting, messageData As WzQuickChatMessageData)
    sendQuickChat_impl message, fromPlayer, targeting, True, messageData
End Sub

Private Sub sendQuickChat_impl(ByVal message As WzQuickChatMessage, ByVal fromPlayer As Long, targeting As WzQuickChatTargeting, ByVal has_messageData As Boolean, messageData As WzQuickChatMessageData)
    Dim hasDataText As String, internalMessage As Boolean, hasMutedUntil As Boolean, mutedUntil As Currency, currentTime As Currency
    Dim isInGame As Boolean, senderTeam As Long, senderIsSpectator As Boolean, senderCanUseSecuredMessages As Boolean
    Dim recipients As UInt32Vector, i As Long, j As Long, playerIdx As Long, messageValue As Long, teamSpecific As Boolean
    Dim recipient As Long, queue As Long, sendSecured As Boolean, w As MessageWriter, numSpecificRecipients As Long
    Dim value As Long
    If Not myResponsibility(fromPlayer) Then ASSERT_FAILURE "myResponsibility(fromPlayer)", "src/hci/quickchat.cpp:3024", "sendQuickChat", "We are not responsible for player: %u", fromPlayer: Exit Sub
    If Not (quickChatMessageExpectsExtraData(message) = has_messageData) Then
        If has_messageData Then hasDataText = "has message data" Else hasDataText = "lacks message data"
        ASSERT_FAILURE "quickChatMessageExpectsExtraData(message) == messageData.has_value()", "src/hci/quickchat.cpp:3025", "sendQuickChat", "Message [%d] unexpectedly %s", CLng(message), hasDataText
        Exit Sub
    End If

    internalMessage = isInternalMessage(message)

    hasMutedUntil = playerSpamMutedUntil(fromPlayer, mutedUntil)
    If hasMutedUntil And Not internalMessage Then
        QueryPerformanceCounter currentTime
        addConsoleMessage "You have sent too many messages in the last few seconds. Please wait and try again.", DEFAULT_JUSTIFY, INFO_MESSAGE, False, FromU32(countsToMs(mutedUntil - currentTime))
        Exit Sub
    End If

    isInGame = (GetGameMode() = GS_NORMAL)
    senderTeam = checkedGetPlayerTeam(fromPlayer)
    senderIsSpectator = NetPlay_players_isSpectator(fromPlayer)
    senderCanUseSecuredMessages = realSelectedPlayer = NetPlay_hostPlayer Or ULT(realSelectedPlayer, MAX_PLAYERS) ' non-host spectator slots don't currently support / send secured messages

    ' build a list of players to whom this message should be sent
    For i = 0 To targeting.specificPlayers.n - 1
        playerIdx = targeting.specificPlayers.v(i)
        If sendQuickChat_isPotentiallyValidTarget(playerIdx, fromPlayer, isInGame, senderIsSpectator) Then
            UInt32Set_insert recipients, playerIdx
        End If
    Next

    For playerIdx = 0 To MAX_CONNECTED_PLAYERS - 1
        If Not sendQuickChat_isPotentiallyValidTarget(playerIdx, fromPlayer, isInGame, senderIsSpectator) Then
            GoTo continue1
        End If

        If NetPlay_players_allocated(playerIdx) Then
            ' human player that's connected
            If targeting.all_ Or (targeting.humanTeammates And checkedGetPlayerTeam(playerIdx) = senderTeam) Then
                UInt32Set_insert recipients, playerIdx
                GoTo continue1
            End If
        ElseIf playerIdx < game.maxPlayers And NetPlay_players_ai(playerIdx) >= 0 And NetPlay_players_difficulty(playerIdx) <> AIDifficulty_DISABLED And Not NetPlay_players_isSpectator(playerIdx) Then
            ' AI player that's still in the game
            If targeting.all_ Or (targeting.aiTeammates And checkedGetPlayerTeam(playerIdx) = senderTeam) Then
                UInt32Set_insert recipients, playerIdx
                GoTo continue1
            End If
        End If
continue1:
    Next

    messageValue = message
    teamSpecific = Not targeting.all_ And (targeting.humanTeammates Or targeting.aiTeammates)

    For i = 0 To recipients.n - 1
        recipient = recipients.v(i)
        If myResponsibility(recipient) Then
            ' don't bother sending a net message
            If recipient = selectedPlayer Then
                ' Targetted at us - Output to console
                addQuickChatMessageToConsole message, fromPlayer, targeting, has_messageData, messageData
            Else
                ' this is presumably a host sending a message to an AI
                If Not (NetPlay_players_ai(recipient) >= 0) Then
                    ASSERT_FAILURE "false", "src/hci/quickchat.cpp:3133", "sendQuickChat", "Not an AI??"
                    GoTo continue2
                End If
                If Not NetPlay_isHost Then ASSERT_FAILURE "NetPlay.isHost", "src/hci/quickchat.cpp:3136", "sendQuickChat", "Not the host?"
                triggerEventQuickChatMessage fromPlayer, recipient, message, teamSpecific
            End If
            GoTo continue2
        End If

        If ULT(recipient, MAX_PLAYERS) Then queue = NETnetQueue(whosResponsible(recipient)) Else queue = NETnetQueue(recipient)
        sendSecured = isInGame And (NETQUEUE_index(queue) = NetPlay_hostPlayer Or NETQUEUE_index(queue) < MAX_PLAYERS) And senderCanUseSecuredMessages
        Set w = Nothing     ' optional<MessageWriter> w;
        If sendSecured Then
            Set w = NETbeginEncodeSecured(queue, NET_QUICK_CHAT_MSG)
            If w Is Nothing Then
                debug_ 3150, LOG_NET, "sendQuickChat", "Failed to encode secured message for queue.index: %u", NETQUEUE_index(queue)
                GoTo continue2
            End If
        Else
            Set w = NETbeginEncode(queue, NET_QUICK_CHAT_MSG)
        End If

        NETuint32_t w, fromPlayer
        NETuint32_t w, recipient
        NETuint32_t w, messageValue
        ' send targeting structure
        NETbool w, targeting.all_
        NETbool w, targeting.humanTeammates
        NETbool w, targeting.aiTeammates
        numSpecificRecipients = targeting.specificPlayers.n
        NETuint32_t w, numSpecificRecipients
        For j = 0 To targeting.specificPlayers.n - 1
            value = targeting.specificPlayers.v(j)
            NETuint32_t w, value
        Next
        If quickChatMessageExpectsExtraData(message) Then
            If has_messageData Then
                NETuint32_t w, messageData.dataContext
                NETuint32_t w, messageData.dataA
                NETuint32_t w, messageData.dataB
            End If
        End If
        NETend w
continue2:
    Next

    If fromPlayer = selectedPlayer And (recipients.n <> 0 Or Not isInGame) And Not shouldHideQuickChatMessageFromLocalDisplay(message) Then
        If isInGame Then
            addQuickChatMessageToConsole message, fromPlayer, targeting, has_messageData, messageData
        Else
            addLobbyQuickChatMessageToConsole message, fromPlayer, targeting, has_messageData, messageData
        End If
    End If

    If recipients.n <> 0 And Not internalMessage Then
        recordPlayerMessageSent fromPlayer
    End If
End Sub

Public Function shouldProcessQuickChatMessage(ByVal queue As Long, ByVal isInGame As Boolean, ByVal message As WzQuickChatMessage, ByVal sender As Long, ByVal recipient As Long, targeting As WzQuickChatTargeting, ByVal has_messageData As Boolean, messageData As WzQuickChatMessageData) As Boolean
    Dim internalMessage As Boolean, senderIsSpectator As Boolean, senderSpamMute As Boolean, mutedUntil As Currency
    internalMessage = isInternalMessage(message)

    If Not ULT(sender, MAX_CONNECTED_PLAYERS) Or Not ULT(recipient, MAX_CONNECTED_PLAYERS) Then
        shouldProcessQuickChatMessage = False
        Exit Function
    End If

    If whosResponsible(sender) <> NETQUEUE_index(queue) Then
        shouldProcessQuickChatMessage = False
        Exit Function
    End If

    If Not myResponsibility(recipient) Then
        shouldProcessQuickChatMessage = False
        Exit Function
    End If

    senderIsSpectator = NetPlay_players_isSpectator(sender)
    If isInGame And senderIsSpectator And Not NetPlay_players_isSpectator(recipient) Then
        ' spectators can't talk to players in-game
        shouldProcessQuickChatMessage = False
        Exit Function
    End If

    senderSpamMute = playerSpamMutedUntil(sender, mutedUntil)
    If senderSpamMute And Not internalMessage Then
        ' ignore message sent while player send was throttled
        shouldProcessQuickChatMessage = False
        Exit Function
    End If

    ' begin: message-specific checks
    Select Case message
    Case WzQuickChatMessage_INTERNAL_ADMIN_ACTION_NOTICE, WzQuickChatMessage_INTERNAL_LOCALIZED_LOBBY_NOTICE
        ' should only ever be sent by the host
        If NETQUEUE_index(queue) <> NetPlay_hostPlayer Then
            shouldProcessQuickChatMessage = False
            Exit Function
        End If
    Case Else
        ' no special handling
    End Select
    ' end: message-specific checks

    ' passes checks
    shouldProcessQuickChatMessage = True
End Function

Public Function recvQuickChat(ByVal queue As Long) As Boolean
    Dim isInGame As Boolean, senderCanUseSecuredMessages As Boolean, expectingSecuredMessage As Boolean, sender As Long
    Dim recipient As Long, messageValue As Long, targeting As WzQuickChatTargeting, has_messageData As Boolean
    Dim messageData As WzQuickChatMessageData, msgEnumVal As Long, r As MessageReader, numSpecificRecipients As Long
    Dim i As Long, tmp_playerIdx As Long, validMessageEnumValue As Boolean, internalMessage As Boolean
    Dim teamSpecific As Boolean
    isInGame = (GetGameMode() = GS_NORMAL)
    senderCanUseSecuredMessages = NETQUEUE_index(queue) = NetPlay_hostPlayer Or NETQUEUE_index(queue) < MAX_PLAYERS
    expectingSecuredMessage = isInGame And (realSelectedPlayer = NetPlay_hostPlayer Or ULT(realSelectedPlayer, MAX_PLAYERS)) And senderCanUseSecuredMessages ' spectator slots do not expect secured messages

    sender = MAX_CONNECTED_PLAYERS
    recipient = MAX_CONNECTED_PLAYERS
    messageValue = WzQuickChatMessage_MESSAGE_COUNT

    If expectingSecuredMessage Then
        Set r = NETbeginDecodeSecured(queue, NET_QUICK_CHAT_MSG)
        If r Is Nothing Then
            recvQuickChat = False
            Exit Function
        End If
    Else
        Set r = NETbeginDecode(queue, NET_QUICK_CHAT_MSG)
    End If
    NETuint32_t r, sender
    NETuint32_t r, recipient
    NETuint32_t r, messageValue
    ' receive targeting structure
    NETbool r, targeting.all_
    NETbool r, targeting.humanTeammates
    NETbool r, targeting.aiTeammates
    numSpecificRecipients = 0
    NETuint32_t r, numSpecificRecipients
    i = 0
    Do While ULT(i, numSpecificRecipients)
        tmp_playerIdx = -1     ' std::numeric_limits<uint32_t>::max()
        NETuint32_t r, tmp_playerIdx
        If ULT(tmp_playerIdx, MAX_CONNECTED_PLAYERS) Then
            UInt32Set_insert targeting.specificPlayers, tmp_playerIdx
        End If
        i = i + 1
    Loop
    validMessageEnumValue = to_WzQuickChatMessage(messageValue, msgEnumVal)
    If validMessageEnumValue Then
        If quickChatMessageExpectsExtraData(msgEnumVal) Then
            has_messageData = True     ' messageData = WzQuickChatMessageData();
            NETuint32_t r, messageData.dataContext
            NETuint32_t r, messageData.dataA
            NETuint32_t r, messageData.dataB
        End If
    End If
    NETend r

    If Not validMessageEnumValue Then
        recvQuickChat = False
        Exit Function
    End If

    If Not shouldProcessQuickChatMessage(queue, isInGame, msgEnumVal, sender, recipient, targeting, has_messageData, messageData) Then
        recvQuickChat = False
        Exit Function
    End If

    internalMessage = isInternalMessage(msgEnumVal)
    If Not internalMessage And recipient = selectedPlayer Then
        recordPlayerMessageSent sender
    End If

    If recipient = selectedPlayer Or recipient = realSelectedPlayer Then
        ' Output to the console
        If isInGame Then
            addQuickChatMessageToConsole msgEnumVal, sender, targeting, has_messageData, messageData
            audio_PlayTrack ID_SOUND_MESSAGEEND
        Else
            addLobbyQuickChatMessageToConsole msgEnumVal, sender, targeting, has_messageData, messageData
            audio_PlayTrack FE_AUDIO_MESSAGEEND
        End If
    Else
        If Not isInGame Then
            ' silently discard messages to AIs when not in the game
            recvQuickChat = True
            Exit Function
        End If

        If Not (NetPlay_players_ai(recipient) >= 0) Then ASSERT_FAILURE "NetPlay.players[recipient].ai >= 0", "src/hci/quickchat.cpp:3351", "recvQuickChat", "Recipient is not an AI?: %u", recipient: recvQuickChat = False: Exit Function
        teamSpecific = Not targeting.all_ And (targeting.humanTeammates Or targeting.aiTeammates)
        triggerEventQuickChatMessage sender, recipient, msgEnumVal, teamSpecific
    End If

    recvQuickChat = True
End Function

'--------------------------------------------------------------------------
' the lambdas

Public Function quickchat_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim psWidget As WIDGET, psParent As WIDGET, psMessagesList As WIDGET, desiredChildWidth As Long, psStrongSection As WIDGET
    Dim strongSendToSelector As WIDGET, strongParent As WIDGET, playerSelectorWidget As WIDGET, d As WzQuickChatSendToSelector
    Dim widg As WIDGET, psStrongToggleMixedButton As WIDGET, humanTeammatesValue As Boolean, aiTeammatesValue As Boolean
    Dim isChosen As Boolean, y0 As Long, p As QCSendToPlayersSelectorData, targetWidgetToTakeFocus As WIDGET
    Dim widget As WIDGET, newValue As Long, weakParent As Long, quickChatLabelWidth As Long, maxPanelSwitcherWidth As Long
    Dim panelSwitcherWidth As Long, panelSwitcherX0 As Long, q As WzQuickChatForm
    Select Case f.id
    Case FN_qcSection_messagesList_calcLayout       ' LAMBDA_CALCLAYOUT_SIMPLE (WIDGET *psWidget)
        Set psWidget = args(0)
        Set psParent = psWidget.parent()
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/hci/quickchat.cpp:1664", "initialize", "No parent": Exit Function
        Set psMessagesList = psWidget
        If Not WIDGET_isKindOf(psMessagesList, K_IntListTabWidget) Then Set psMessagesList = Nothing
        If Not (Not psMessagesList Is Nothing) Then ASSERT_FAILURE "psMessagesList != nullptr", "src/hci/quickchat.cpp:1666", "initialize", "Not expected type?": Exit Function
        psMessagesList.setGeometry 0, 0, psParent.width(), psParent.height()
        desiredChildWidth = psParent.width()
        ListTabWidget_setChildSize psMessagesList, desiredChildWidth, WzQuickChatButton_getIdealBaseButtonHeight()

    Case FN_qcSection_chatOption_onClick            ' [message, psWeakSection](W_BUTTON& widg)
        Set psStrongSection = WeakWidget_lock(f.cap(1))
        If Not (Not psStrongSection Is Nothing) Then ASSERT_FAILURE "psStrongSection != nullptr", "src/hci/quickchat.cpp:1687", "addChatOption", "No valid section?": Exit Function
        Set psParent = psStrongSection.parent()
        If Not WIDGET_isKindOf(psParent, K_WzQuickChatPanel) Then Set psParent = Nothing
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/hci/quickchat.cpp:1689", "addChatOption", "No valid parent?": Exit Function
        WzQuickChatPanel_sendQuickChat psParent, f.cap(0), selectedPlayer

    Case FN_qcSendTo_clicked_task                   ' [weakThis]()
        Set strongSendToSelector = WeakWidget_lock(f.cap(0))
        If Not (Not strongSendToSelector Is Nothing) Then ASSERT_FAILURE "strongSendToSelector != nullptr", "src/hci/quickchat.cpp:1149", "clicked", "Widget no longer exists?": Exit Function
        Set strongParent = strongSendToSelector.parent()
        If Not WIDGET_isKindOf(strongParent, K_WzQuickChatPanel) Then Set strongParent = Nothing
        If Not (Not strongParent Is Nothing) Then ASSERT_FAILURE "strongParent != nullptr", "src/hci/quickchat.cpp:1151", "clicked", "Parent doesn't exist?": Exit Function
        ' give the current panel focus
        WzQuickChatPanel_takeFocus strongParent

        ' open chooser
        WzQuickChatSendToSelector_openChooser strongSendToSelector

    Case FN_qcSendTo_openChooser_task               ' [pWeakThis]()
        Set playerSelectorWidget = WeakWidget_lock(f.cap(0))
        If Not playerSelectorWidget Is Nothing Then
            Set d = WzQuickChatSendToSelector_(playerSelectorWidget)
            Set d.overlayScreen = W_SCREEN_make()
            widgRegisterOverlayScreenOnTopOfScreen d.overlayScreen, playerSelectorWidget.screenPointer
            d.playersSelectorWidget.setGeometry playerSelectorWidget.screenPosX(), WzQuickChatSendToSelector_calculateDropdownListScreenPosY(playerSelectorWidget), playerSelectorWidget.idealWidth(), d.playersSelectorWidget.idealHeight()
            d.overlayScreen.psForm.attach d.playersSelectorWidget
            d.overlayScreen.psForm.attach ChooserOverlay_new(WzFunc_new(FN_qcChooserOverlay_onClicked, f.cap(0)))
        End If

    Case FN_qcSendTo_closeChooser_task              ' [pWeakThis]()
        Set playerSelectorWidget = WeakWidget_lock(f.cap(0))
        If Not playerSelectorWidget Is Nothing Then
            Set d = WzQuickChatSendToSelector_(playerSelectorWidget)
            If Not d.overlayScreen Is Nothing Then
                widgRemoveOverlayScreen d.overlayScreen
                If Not d.overlayScreen.psForm Is Nothing Then
                    d.overlayScreen.psForm.detach d.playersSelectorWidget
                End If
                Set d.overlayScreen = Nothing
            End If
        End If

    Case FN_qcChooserOverlay_calcLayout             ' LAMBDA_CALCLAYOUT_SIMPLE
        Set psWidget = args(0)
        psWidget.setGeometry 0, 0, screenWidth - 1, screenHeight - 1

    Case FN_qcChooserOverlay_onClicked              ' [pWeakThis]()
        Set playerSelectorWidget = WeakWidget_lock(f.cap(0))
        If Not playerSelectorWidget Is Nothing Then
            WzQuickChatSendToSelector_closeChooser playerSelectorWidget
        End If

    Case FN_qcPlayersSel_humanTeammates_onClick     ' [psWeakParent](W_BUTTON& widg)
        Set widg = args(0)
        Set strongParent = WeakWidget_lock(f.cap(0))
        If Not (Not strongParent Is Nothing) Then ASSERT_FAILURE "strongParent != nullptr", "src/hci/quickchat.cpp:1259", "initialize", "No parent?": Exit Function
        Set psStrongToggleMixedButton = widg
        If Not WIDGET_isKindOf(psStrongToggleMixedButton, K_WzQuickChatToggleMixedCheckButton) Then Set psStrongToggleMixedButton = Nothing
        If Not (Not psStrongToggleMixedButton Is Nothing) Then ASSERT_FAILURE "psStrongToggleMixedButton != nullptr", "src/hci/quickchat.cpp:1261", "initialize", "Unable to get proper access to self?": Exit Function
        ' (WzQuickChatTargeting modifiedTargeting = strongParent->currentTargeting(); is not used)
        humanTeammatesValue = False
        Select Case WzQuickChatToggleMixedCheckButton_getCheckedState(psStrongToggleMixedButton)
        Case MixedCheckState_Mixed
            ' do nothing
            Exit Function
        Case MixedCheckState_Unchecked
            humanTeammatesValue = False
        Case MixedCheckState_Checked
            humanTeammatesValue = True
        End Select
        WzQuickChatSendToSelector_updateTargetingByCategory strongParent, False, False, True, humanTeammatesValue, False, False

    Case FN_qcPlayersSel_humanTeammates_calcLayout  ' LAMBDA_CALCLAYOUT_SIMPLE
        Set psWidget = args(0)
        Set psParent = psWidget.parent()
        If Not WIDGET_isKindOf(psParent, K_WzQuickChatSendToPlayersSelector) Then Set psParent = Nothing
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/hci/quickchat.cpp:1281", "initialize", "Parent is null?": Exit Function
        psWidget.setGeometry 0, 0, psParent.width(), psWidget.idealHeight()

    Case FN_qcPlayersSel_aiTeammates_onClick        ' [psWeakParent](W_BUTTON& widg)
        Set widg = args(0)
        Set strongParent = WeakWidget_lock(f.cap(0))
        If Not (Not strongParent Is Nothing) Then ASSERT_FAILURE "strongParent != nullptr", "src/hci/quickchat.cpp:1291", "initialize", "No parent?": Exit Function
        Set psStrongToggleMixedButton = widg
        If Not WIDGET_isKindOf(psStrongToggleMixedButton, K_WzQuickChatToggleMixedCheckButton) Then Set psStrongToggleMixedButton = Nothing
        If Not (Not psStrongToggleMixedButton Is Nothing) Then ASSERT_FAILURE "psStrongToggleMixedButton != nullptr", "src/hci/quickchat.cpp:1293", "initialize", "Unable to get proper access to self?": Exit Function
        aiTeammatesValue = False
        Select Case WzQuickChatToggleMixedCheckButton_getCheckedState(psStrongToggleMixedButton)
        Case MixedCheckState_Mixed
            ' do nothing
            Exit Function
        Case MixedCheckState_Unchecked
            aiTeammatesValue = False
        Case MixedCheckState_Checked
            aiTeammatesValue = True
        End Select
        WzQuickChatSendToSelector_updateTargetingByCategory strongParent, False, False, False, False, True, aiTeammatesValue

    Case FN_qcPlayersSel_aiTeammates_calcLayout     ' LAMBDA_CALCLAYOUT_SIMPLE
        Set psWidget = args(0)
        Set psParent = psWidget.parent()
        If Not WIDGET_isKindOf(psParent, K_WzQuickChatSendToPlayersSelector) Then Set psParent = Nothing
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/hci/quickchat.cpp:1312", "initialize", "Parent is null?": Exit Function
        Set p = WzQuickChatSendToPlayersSelector_(psParent)
        If Not p.humanTeammates Is Nothing Then y0 = p.humanTeammates.y() + p.humanTeammates.height() + QC_SENDTO_SELECTOR_INTERNAL_SPACING Else y0 = 0
        psWidget.setGeometry 0, y0, psParent.width(), psWidget.idealHeight()

    Case FN_qcPlayersSel_player_onClick             ' [psWeakParent, playerIdx](W_BUTTON& widg)
        Set widg = args(0)
        Set strongParent = WeakWidget_lock(f.cap(0))
        If Not (Not strongParent Is Nothing) Then ASSERT_FAILURE "strongParent != nullptr", "src/hci/quickchat.cpp:1324", "initialize", "No parent?": Exit Function
        isChosen = (widg.getState() And WBUT_CLICKLOCK) <> 0
        WzQuickChatSendToSelector_updatePlayerSelectionState strongParent, f.cap(1), isChosen

    Case FN_qcPlayersSel_playersList_calcLayout     ' [maxPlayersListHeight](WIDGET *psWidget)
        Set psWidget = args(0)
        Set psParent = psWidget.parent()
        If Not WIDGET_isKindOf(psParent, K_WzQuickChatSendToPlayersSelector) Then Set psParent = Nothing
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/hci/quickchat.cpp:1337", "initialize", "Parent is null?": Exit Function
        Set p = WzQuickChatSendToPlayersSelector_(psParent)
        y0 = 0
        If Not p.aiTeammates Is Nothing Then
            y0 = p.aiTeammates.y() + p.aiTeammates.height() + QC_SENDTO_SELECTOR_INTERNAL_SPACING
        ElseIf Not p.humanTeammates Is Nothing Then
            y0 = p.humanTeammates.y() + p.humanTeammates.height() + QC_SENDTO_SELECTOR_INTERNAL_SPACING
        End If
        psWidget.setGeometry 0, y0, psParent.width(), f.cap(0)

    Case FN_qcPanel_tabPress_task                   ' [targetWidgetToTakeFocus]()
        Set targetWidgetToTakeFocus = f.cap(0)
        WzQuickChatSection_clickPrimary targetWidgetToTakeFocus

    Case FN_qcForm_panelSwitcher_onChoose           ' [](MultibuttonWidget& widget, int newValue)
        Set widget = args(0)
        newValue = args(1)
        ' Switch actively-displayed panel
        ' (auto weakParent = std::weak_ptr<WIDGET>(widget.parent()): the form's serial, 0 for another parent, which the
        ' task's dynamic_pointer_cast would refuse)
        Set psParent = widget.parent()
        weakParent = 0
        If WIDGET_isKindOf(psParent, K_WzQuickChatForm) Then weakParent = WzQuickChatForm_(psParent).weakSerial
        widgScheduleTask WzFunc_new(FN_qcForm_panelSwitcher_task, newValue, weakParent)

    Case FN_qcForm_panelSwitcher_task               ' [newValue, weakParent]()
        Set strongParent = WeakWidget_lock(f.cap(1))
        If Not (Not strongParent Is Nothing) Then ASSERT_FAILURE "strongParent != nullptr", "src/hci/quickchat.cpp:2167", "initialize", "Parent doesn't exist?": Exit Function
        Set q = WzQuickChatForm_(strongParent)
        Set widget = q.panels(CLng(f.cap(0)) + 1)     ' strongParent->panels[newValue]
        WzQuickChatForm_switchAttachedPanel strongParent, widget
        ' give the new panel focus
        If Not q.currentlyAttachedPanel Is Nothing Then
            WzQuickChatPanel_takeFocus q.currentlyAttachedPanel
        End If

    Case FN_qcForm_panelSwitcher_calcLayout         ' LAMBDA_CALCLAYOUT_SIMPLE
        Set psWidget = args(0)
        Set psParent = psWidget.parent()
        If Not WIDGET_isKindOf(psParent, K_WzQuickChatForm) Then Set psParent = Nothing
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/hci/quickchat.cpp:2178", "initialize", "No parent": Exit Function
        quickChatLabelWidth = WzQuickChatForm_(psParent).quickChatLabel.width()
        maxPanelSwitcherWidth = psParent.width() - ((QCFORM_INTERNAL_PADDING + quickChatLabelWidth) * 2)
        panelSwitcherWidth = MIN_(maxPanelSwitcherWidth, psWidget.idealWidth())
        panelSwitcherX0 = (psParent.width() - panelSwitcherWidth) \ 2
        psWidget.setGeometry panelSwitcherX0, QCFORM_INTERNAL_PADDING, panelSwitcherWidth, QCFORM_PANEL_TABS_HEIGHT
    End Select
End Function
