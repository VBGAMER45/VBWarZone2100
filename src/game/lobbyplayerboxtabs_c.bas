Attribute VB_Name = "lobbyplayerboxtabs_c"
' src/titleui/widgets/lobbyplayerboxtabs.cpp, lobbyplayerboxtabs.h: the lobby's
' Players / Spectators tabs (WzPlayerBoxTabs, WzPlayerBoxTabButton), the host's
' options button and pop-over (WzPlayerBoxOptionsButton, LobbyHostOptionsForm)
' and their helpers. The widget object model is docs/CONVENTIONS.md's
' "Widgets": the classes' fields are their data classes (same names), the
' lambdas WzFuncs run by lobbyplayerboxtabs_c_fn; a std::weak_ptr is an
' address (ObjPtr).
' (VB) widget class LobbyHostOptionsForm : WIDGET
' (VB) widget class WzPlayerBoxTabButton : W_BUTTON
' (VB) widget class WzPlayerBoxOptionsButton : W_BUTTON
' (VB) widget class WzPlayerBoxTabs : WIDGET
Option Explicit

Private Const MAX_CONNECTED_PLAYERS As Long = MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

' WzPlayerBoxTabs::PlayerDisplayView
Public Enum WzPlayerBoxTabs_PlayerDisplayView
    PlayerDisplayView_Players
    PlayerDisplayView_Spectators
End Enum

' LobbyHostOptionsForm's const int members
Private Const innerVerticalSpacing As Long = 15
Private Const outerPaddingX As Long = 20
Private Const optionsListPaddingX As Long = 0
Private Const outerPaddingY As Long = 20
Private Const betweenButtonPadding As Long = 15
Private Const closeButtonVerticalPadding As Long = 5

' WzPlayerBoxTabButton's const int members
Private Const elementPadding As Long = 5
Private Const borderWidth As Long = 1

' the lambdas and function pointers
Public Const FN_LHOF_CLOSE_CLICK As Long = 6100         ' closeButton->addOnClickHandler([onClose](W_BUTTON&) {...})
Public Const FN_LHOF_CLOSE_TASK As Long = 6101          ' widgScheduleTask([onClose]() { onClose(); })
Public Const FN_IsGameHostedInLobby As Long = 6102      ' OptionInfo::AvailabilityResult IsGameHostedInLobby(const OptionInfo&)
Public Const FN_LOBBYHOSTJOINOPT_SUPPORTED As Long = 6103 ' MakeIsLobbyHostJoinOptSupportedCondition's [optType](const OptionInfo&)
Public Const FN_LHOF_SPECJOIN_POPULATE As Long = 6104
Public Const FN_LHOF_SPECJOIN_SELECT As Long = 6105
Public Const FN_LHOF_LOCKTEAMS_POPULATE As Long = 6106
Public Const FN_LHOF_LOCKTEAMS_SELECT As Long = 6107
Public Const FN_LHOF_LOCKPOS_POPULATE As Long = 6108
Public Const FN_LHOF_LOCKPOS_SELECT As Long = 6109
Public Const FN_LHOF_INACTIVITY_SELECT As Long = 6110
Public Const FN_LHOF_LAGKICK_SELECT As Long = 6111
Public Const FN_LHOF_TIMELIMIT_SELECT As Long = 6112
Public Const FN_LHOF_HOSTJOINOPT_POPULATE As Long = 6113 ' the three host join options' populate functions: cap(0) the HostJoinOptionType
Public Const FN_LHOF_HOSTJOINOPT_SELECT As Long = 6114
Public Const FN_LHOF_DIRECTJOIN_POPULATE As Long = 6115
Public Const FN_LHOF_DIRECTJOIN_SELECT As Long = 6116
Public Const FN_PBOB_TITLE_CALCLAYOUT As Long = 6117    ' WzPlayerBoxOptionsButton's titleLabel->setCalcLayout([psWeakParent](WIDGET *psWidget) {...})
Public Const FN_PBT_PLAYERS_CLICK As Long = 6118        ' playersButton->addOnClickHandler([weakSelf](W_BUTTON& button) {...})
Public Const FN_PBT_SPECTATORS_CLICK As Long = 6119
Public Const FN_PBT_SETVIEW_TASK As Long = 6120         ' widgScheduleTask([weakSelf]() {... setDisplayView(cap(1)) })
Public Const FN_PBT_OPTIONS_CLICK As Long = 6121        ' optionsButton->addOnClickHandler([](W_BUTTON& button) {...})
Public Const FN_PBT_CALCLAYOUT As Long = 6122           ' widget->setCalcLayout(LAMBDA_CALCLAYOUT_SIMPLE({...}))
Public Const FN_PBT_POPOVER_CLOSE As Long = 6123        ' LobbyHostOptionsForm::make(weakTitleUI, [psWeakSelf]() {...})
Public Const FN_PBT_OVERLAY_CLICKED As Long = 6124      ' newRootFrm->onClickedFunc
Public Const FN_PBT_POPOVER_CALCLAYOUT As Long = 6125   ' optionsPopOver->setCalcLayout([](WIDGET *psWidget) {...})
Public Const FN_PBT_CLEAR_FOCUS As Long = 6126          ' widgScheduleTask([lockedScreen]() {...})

' (VB) std::weak_ptr<WzMultiplayerOptionsTitleUI>::lock(): the title UI at the address (0: nullptr)
Private Function titleUIFromPtr(ByVal p As Long) As WzMultiplayerOptionsTitleUI
    Dim o As WzMultiplayerOptionsTitleUI
    If p = 0 Then Exit Function
    CopyMemory o, p, 4
    Set titleUIFromPtr = o
    CopyMemory o, 0&, 4
End Function

' (VB) PIELIGHT's byte.a, read and set
Private Function pielightA(ByVal c As Long) As Long
    pielightA = SHRU(c, 24) And &HFF&
End Function

Private Function pielightWithA(ByVal c As Long, ByVal a As Long) As Long
    pielightWithA = (c And &HFFFFFF) Or SHL(a And &HFF&, 24)
End Function

' (VB) OptionInfo::AvailabilityResult result; result.available_ = ...; result.localizedUnavailabilityReason = ...;
Private Function makeAvailabilityResult(ByVal isAvailable As Boolean, ByVal localizedUnavailabilityReason As String) As OptionInfo_AvailabilityResult
    Dim result As OptionInfo_AvailabilityResult
    Set result = New OptionInfo_AvailabilityResult
    result.available_ = isAvailable
    result.localizedUnavailabilityReason = localizedUnavailabilityReason
    Set makeAvailabilityResult = result
End Function

' MARK: - Helper functions

Private Function hasOpenSpectatorOnlySlots() As Boolean
    Dim i As Long
    ' Look for a spectator slot that's available
    For i = 0 To MAX_CONNECTED_PLAYERS - 1
        If Not isSpectatorOnlySlot(i) Then
            GoTo continue_
        End If
        If game.mapHasScavengers And NetPlay_players_position(i) = scavengerSlot() Then
            GoTo continue_ ' skip it
        End If
        If i = PLAYER_FEATURE Then
            GoTo continue_ ' skip it
        End If
        If NetPlay_players_isSpectator(i) And NetPlay_players_ai(i) = AI_OPEN And Not NetPlay_players_allocated(i) Then
            ' found available spectator-only slot
            hasOpenSpectatorOnlySlots = True
            Exit Function
        End If
continue_:
    Next
    hasOpenSpectatorOnlySlots = False
End Function

Private Sub closeAllOpenSpectatorOnlySlots()
    Dim i As Long
    If Not NetPlay_isHost Then ASSERT_FAILURE "false", "src/titleui/widgets/lobbyplayerboxtabs.cpp:76", "closeAllOpenSpectatorOnlySlots", "Host only routine detected for client!": Exit Sub
    For i = 0 To MAX_CONNECTED_PLAYERS - 1
        If Not isSpectatorOnlySlot(i) Then
            GoTo continue_
        End If
        If game.mapHasScavengers And NetPlay_players_position(i) = scavengerSlot() Then
            GoTo continue_ ' skip it
        End If
        If i = PLAYER_FEATURE Then
            GoTo continue_ ' skip it
        End If
        If NetPlay_players_isSpectator(i) And NetPlay_players_ai(i) = AI_OPEN And Not NetPlay_players_allocated(i) Then
            ' found available spectator-only slot
            ' close it
            NetPlay_players_ai(i) = AI_CLOSED
            NetPlay_players_isSpectator(i) = False
            NETBroadcastPlayerInfo i
        End If
continue_:
    Next
End Sub

Private Sub enableSpectatorJoin(ByVal enabled As Boolean)
    Dim success As Boolean
    If enabled Then
        ' add max spectator slots
        success = False
        Do
            success = NETopenNewSpectatorSlot()
        Loop While success
    Else
        ' disable any open / unoccupied spectator slots
        closeAllOpenSpectatorOnlySlots
    End If
    netPlayersUpdated = True
End Sub

' MARK: - LobbyHostOptionsForm

Public Sub LobbyHostOptionsForm_construct(ByVal self As WIDGET)
    WIDGET_construct self, WIDG_UNSPECIFIED_TYPE
    self.kind = K_LobbyHostOptionsForm ' (VB) virtual calls from here on reach LobbyHostOptionsForm's overrides
End Sub

' static std::shared_ptr<LobbyHostOptionsForm> make(const std::weak_ptr<WzMultiplayerOptionsTitleUI>& weakTitleUI, const std::function<void ()>& onCloseFunc)
' (weakTitleUI: the title UI's address)
Public Function LobbyHostOptionsForm_make(ByVal weakTitleUI As Long, ByVal onCloseFunc As WzFunc) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_LobbyHostOptionsForm)
    LobbyHostOptionsForm_construct widget
    LobbyHostOptionsForm_initialize widget, weakTitleUI, onCloseFunc
    Set LobbyHostOptionsForm_make = widget
End Function

Private Sub LobbyHostOptionsForm_initialize(ByVal self As WIDGET, ByVal weakTitleUI_ As Long, ByVal onClose As WzFunc)
    Dim d As LobbyHostOptionsForm
    Set d = LobbyHostOptionsForm_(self)
    d.weakTitleUI = weakTitleUI_

    Set d.formTitle = W_LABEL_new()
    W_LABEL_setFont d.formTitle, font_medium_bold, psPalette(WZCOL_TEXT_BRIGHT)
    d.formTitle.setString "Host Lobby Options"
    W_LABEL_setCanTruncate d.formTitle, True
    d.formTitle.setTransparentToMouse True
    d.formTitle.setGeometry outerPaddingX, outerPaddingY, W_LABEL_getMaxLineWidth(d.formTitle), iV_GetTextLineSize(font_regular_bold)
    self.attach d.formTitle

    ' Add "Close" button
    Set d.closeButton = WzFrontendImageButton_make(False, 0)
    d.closeButton.setString "Close"
    W_BUTTON_addOnClickHandler d.closeButton, WzFunc_new(FN_LHOF_CLOSE_CLICK, onClose)
    WzFrontendImageButton_setPadding d.closeButton, 7, closeButtonVerticalPadding
    self.attach d.closeButton

    Set d.hostOptionsForm = LobbyHostOptionsForm_makeHostOptionsForm(self)
    self.attach d.hostOptionsForm
End Sub

Public Function LobbyHostOptionsForm_idealWidth(ByVal self As WIDGET) As Long
    LobbyHostOptionsForm_idealWidth = (outerPaddingX * 2) + optionsListPaddingX + LobbyHostOptionsForm_(self).hostOptionsForm.idealWidth()
End Function

Public Function LobbyHostOptionsForm_idealHeight(ByVal self As WIDGET) As Long
    Dim d As LobbyHostOptionsForm, result As Long, a As Long, b As Long
    Set d = LobbyHostOptionsForm_(self)
    a = d.formTitle.idealHeight()
    b = d.closeButton.idealHeight() - closeButtonVerticalPadding
    If b > a Then a = b
    result = (outerPaddingY * 2) + a
    result = result + innerVerticalSpacing
    result = result + d.hostOptionsForm.idealHeight()
    result = result + innerVerticalSpacing
    LobbyHostOptionsForm_idealHeight = result
End Function

Private Function hasGameCurrentlyHostedInLobby() As Boolean
    hasGameCurrentlyHostedInLobby = Len(NET_getCurrentHostedLobbyGameId()) <> 0
End Function

' OptionInfo::AvailabilityResult IsGameHostedInLobby(const OptionInfo&) (a WzFunc: FN_IsGameHostedInLobby)
Public Function IsGameHostedInLobby(ByVal info As OptionInfo) As OptionInfo_AvailabilityResult
    Set IsGameHostedInLobby = makeAvailabilityResult(hasGameCurrentlyHostedInLobby(), "Game must be listed in lobby to support this option.")
End Function

' OptionInfo::OptionAvailabilityCondition MakeIsLobbyHostJoinOptSupportedCondition(netlobby::HostJoinOptionType optType)
Public Function MakeIsLobbyHostJoinOptSupportedCondition(ByVal optType As HostJoinOptionType) As WzFunc
    Set MakeIsLobbyHostJoinOptSupportedCondition = WzFunc_new(FN_LOBBYHOSTJOINOPT_SUPPORTED, optType)
End Function

Private Function lobbyHostJoinOptionEnabled(ByVal optType As HostJoinOptionType) As Boolean
    Dim supported As Boolean
    If hasGameCurrentlyHostedInLobby() Then
        ' NET_getHostJoinOptionSupportedByLobby(optType).value_or(false)
        If NET_getHostJoinOptionSupportedByLobby(optType, supported) Then lobbyHostJoinOptionEnabled = supported
    End If
End Function

' (VB) NETgetLobbyHostJoinOptions()'s field for optType
Private Function hostJoinOptionValue(ByVal optType As HostJoinOptionType) As Long
    Dim botProt As Long, proxyIPs As Long, hostingIPs As Long
    NETgetLobbyHostJoinOptions botProt, proxyIPs, hostingIPs
    Select Case optType
    Case HostJoinOptionType_BotProt: hostJoinOptionValue = botProt
    Case HostJoinOptionType_ProxyIPs: hostJoinOptionValue = proxyIPs
    Case HostJoinOptionType_HostingIPs: hostJoinOptionValue = hostingIPs
    End Select
End Function

Private Function LobbyHostOptionsForm_makeHostOptionsForm(ByVal self As WIDGET) As WIDGET
    Dim weakTitleUICopy As Long, psWeakParent As Long, result As WIDGET, optionInfo As OptionInfo, valueChanger As WIDGET
    weakTitleUICopy = LobbyHostOptionsForm_(self).weakTitleUI
    psWeakParent = ObjPtr(self)
    Set result = OptionsForm_make()

    ' Slot Options
    OptionsForm_addSection result, OptionsSection_("Slot Options", ""), True
    Set optionInfo = OptionInfo_new("mp.spectatorjoin", "Spectator Join", "")
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_LHOF_SPECJOIN_POPULATE), WzFunc_new(FN_LHOF_SPECJOIN_SELECT, weakTitleUICopy), True)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set optionInfo = OptionInfo_new("mp.locked.teams", "Teams", "")
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_LHOF_LOCKTEAMS_POPULATE), WzFunc_new(FN_LHOF_LOCKTEAMS_SELECT, weakTitleUICopy), False)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set optionInfo = OptionInfo_new("mp.locked.position", "Position", "")
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_LHOF_LOCKPOS_POPULATE), WzFunc_new(FN_LHOF_LOCKPOS_SELECT, weakTitleUICopy), False)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    ' Game Options
    OptionsForm_addSection result, OptionsSection_("Game Options", ""), True
    Set optionInfo = OptionInfo_new("defaults.hosting.inactivityTimeout", "Inactivity Timeout", "")
    optionInfo.addAvailabilityCondition WzFunc_new(FN_IsNotInGame)
    Set valueChanger = OptionsDropdown_make(makeInactivityTimeoutOptionsDropdownPopulateFunc(), WzFunc_new(FN_LHOF_INACTIVITY_SELECT), False)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set optionInfo = OptionInfo_new("defaults.hosting.lagKick", "Lag Kick", "")
    Set valueChanger = OptionsDropdown_make(makeLagKickOptionsDropdownPopulateFunc(), WzFunc_new(FN_LHOF_LAGKICK_SELECT), False)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set optionInfo = OptionInfo_new("defaults.hosting.gameTimeLimit", "Game Time Limit", "")
    optionInfo.addAvailabilityCondition WzFunc_new(FN_IsNotInGame)
    Set valueChanger = OptionsDropdown_make(makeGameTimeLimitOptionsDropdownPopulateFunc(), WzFunc_new(FN_LHOF_TIMELIMIT_SELECT), False)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    ' Host Join Options
    OptionsForm_addSection result, OptionsSection_("Lobby Join Options", ""), True
    Set optionInfo = OptionInfo_new("mp.hostjoinoptions.botprot", "Bot Protection", "Whether automated bots or tools can join or gather host connection info (including your IP address). If you are encountering undesirable joins, spammers, or other abuse, you may wish to enable this option.")
    optionInfo.addAvailabilityCondition WzFunc_new(FN_IsGameHostedInLobby)
    optionInfo.addAvailabilityCondition MakeIsLobbyHostJoinOptSupportedCondition(HostJoinOptionType_BotProt)
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_LHOF_HOSTJOINOPT_POPULATE, HostJoinOptionType_BotProt), WzFunc_new(FN_LHOF_HOSTJOINOPT_SELECT, HostJoinOptionType_BotProt, weakTitleUICopy), False)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set optionInfo = OptionInfo_new("mp.hostjoinoptions.proxyips", "Proxy / VPN Connections", "Proxies / VPNs hide the true source of connections. May be used by players trying to evade bans or hide their identity.")
    optionInfo.addAvailabilityCondition WzFunc_new(FN_IsGameHostedInLobby)
    optionInfo.addAvailabilityCondition MakeIsLobbyHostJoinOptSupportedCondition(HostJoinOptionType_ProxyIPs)
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_LHOF_HOSTJOINOPT_POPULATE, HostJoinOptionType_ProxyIPs), WzFunc_new(FN_LHOF_HOSTJOINOPT_SELECT, HostJoinOptionType_ProxyIPs, weakTitleUICopy), False)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set optionInfo = OptionInfo_new("mp.hostjoinoptions.hostingips", "Hosting Provider IPs", "Server hosting providers. May be used by automated tools, port scanners, or players trying to evade bans.")
    optionInfo.addAvailabilityCondition WzFunc_new(FN_IsGameHostedInLobby)
    optionInfo.addAvailabilityCondition MakeIsLobbyHostJoinOptSupportedCondition(HostJoinOptionType_HostingIPs)
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_LHOF_HOSTJOINOPT_POPULATE, HostJoinOptionType_HostingIPs), WzFunc_new(FN_LHOF_HOSTJOINOPT_SELECT, HostJoinOptionType_HostingIPs, weakTitleUICopy), False)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set optionInfo = OptionInfo_new("mp.hostjoinoptions.lobbyjoinbypass", "Lobby Join Bypass", "Allow clients to bypass lobby join via direct port:ip connection. (May enable trolls to bypass lobby anti-abuse protections.)")
    optionInfo.addAvailabilityCondition WzFunc_new(FN_IsGameHostedInLobby)
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_LHOF_DIRECTJOIN_POPULATE), WzFunc_new(FN_LHOF_DIRECTJOIN_SELECT), False)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set LobbyHostOptionsForm_makeHostOptionsForm = result
End Function

Public Sub LobbyHostOptionsForm_geometryChanged(ByVal self As WIDGET)
    Dim d As LobbyHostOptionsForm, lastButtonX0 As Long, buttonY0 As Long, topButtonHeight As Long, buttonX0 As Long
    Dim maxAvailableOptionsFormHeight As Long, hostOptionsFormY0 As Long, a As Long, b As Long, h As Long
    Set d = LobbyHostOptionsForm_(self)
    If self.width() <= 1 Or self.height() <= 1 Then
        Exit Sub
    End If

    ' position top button(s)
    lastButtonX0 = self.width() - outerPaddingX
    buttonY0 = outerPaddingY - closeButtonVerticalPadding
    topButtonHeight = d.closeButton.idealHeight()
    If Not d.closeButton Is Nothing Then
        buttonX0 = lastButtonX0 - d.closeButton.idealWidth()
        d.closeButton.setGeometry buttonX0, buttonY0, d.closeButton.idealWidth(), topButtonHeight
        lastButtonX0 = buttonX0 - betweenButtonPadding
    End If

    a = d.formTitle.height()
    b = d.closeButton.height() - closeButtonVerticalPadding
    If b > a Then a = b
    maxAvailableOptionsFormHeight = self.height() - (outerPaddingY * 2) - a - (innerVerticalSpacing * 2)
    a = d.formTitle.idealHeight()
    b = d.closeButton.idealHeight()
    If b > a Then a = b
    hostOptionsFormY0 = outerPaddingY + a + innerVerticalSpacing
    h = d.hostOptionsForm.idealHeight()
    If maxAvailableOptionsFormHeight < h Then h = maxAvailableOptionsFormHeight
    d.hostOptionsForm.setGeometry outerPaddingX + optionsListPaddingX, hostOptionsFormY0, self.width() - (outerPaddingX * 2) - optionsListPaddingX, h

End Sub

Public Sub LobbyHostOptionsForm_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x0 As Long, y0 As Long, x1 As Long, y1 As Long, backColor As Long
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    x1 = x0 + self.width()
    y1 = y0 + self.height()

    backColor = psPalette(WZCOL_TRANSPARENT_BOX)
    If pielightA(backColor) < 220 Then backColor = pielightWithA(backColor, 220)
    pie_UniTransBoxFill x0, y0, x1, y1, backColor

    iV_Box x0, y0, x1, y1, pal_RGBA(255, 255, 255, 175)
End Sub


' MARK: - WzPlayerBoxTabButton

Public Sub WzPlayerBoxTabButton_construct(ByVal self As WIDGET)
    W_BUTTON_construct self
    self.kind = K_WzPlayerBoxTabButton ' (VB) virtual calls from here on reach WzPlayerBoxTabButton's overrides
End Sub

Public Function WzPlayerBoxTabButton_make(ByVal title_ As String) As WIDGET
    Dim widget As WIDGET, d As WzPlayerBoxTabButton
    Set widget = WIDGET_alloc(K_WzPlayerBoxTabButton)
    WzPlayerBoxTabButton_construct widget
    Set d = WzPlayerBoxTabButton_(widget)

    ' add the titleLabel
    Set d.titleLabel = W_LABEL_new()
    W_LABEL_setFont d.titleLabel, font_regular, psPalette(WZCOL_TEXT_BRIGHT)
    d.titleLabel.setString title_
    d.titleLabel.setGeometry 0, 0, W_LABEL_getMaxLineWidth(d.titleLabel), d.titleLabel.idealHeight()
    W_LABEL_setCanTruncate d.titleLabel, True

    ' add the slot counts label
    Set d.slotsCountsLabel = W_LABEL_new()
    W_LABEL_setFont d.slotsCountsLabel, font_small, psPalette(WZCOL_TEXT_BRIGHT)

    ' refresh tooltip
    WzPlayerBoxTabButton_refreshTip widget

    Set WzPlayerBoxTabButton_make = widget
End Function

Private Function WzPlayerBoxTabButton_getImageForSlotsReadyStatus(ByVal self As WIDGET) As AtlasImage
    Dim d As WzPlayerBoxTabButton
    Set d = WzPlayerBoxTabButton_(self)
    If d.totalAvailableSlots = 0 Then
        WzPlayerBoxTabButton_getImageForSlotsReadyStatus = AtlasImage_()
        Exit Function
    End If
    If d.numTakenSlots = 0 Then
        WzPlayerBoxTabButton_getImageForSlotsReadyStatus = AtlasImage_()
        Exit Function
    End If
    If d.numTakenSlots <> d.numTakenSlotsReady Then
        WzPlayerBoxTabButton_getImageForSlotsReadyStatus = AtlasImage_(FrontImages, IMAGE_LAMP_AMBER)
        Exit Function
    End If
    WzPlayerBoxTabButton_getImageForSlotsReadyStatus = AtlasImage_(FrontImages, IMAGE_LAMP_GREEN)
End Function

' setSlotCounts(uint8_t numTakenSlots, uint8_t totalAvailableSlots)
Public Sub WzPlayerBoxTabButton_setSlotCounts(ByVal self As WIDGET, ByVal new_numTakenSlots As Long, ByVal new_totalAvailableSlots As Long)
    Dim d As WzPlayerBoxTabButton, slotCountsStr As String
    Set d = WzPlayerBoxTabButton_(self)
    new_numTakenSlots = new_numTakenSlots And &HFF&
    new_totalAvailableSlots = new_totalAvailableSlots And &HFF&
    If new_numTakenSlots <> d.numTakenSlots Or new_totalAvailableSlots <> d.totalAvailableSlots Then
        d.numTakenSlots = new_numTakenSlots
        d.totalAvailableSlots = new_totalAvailableSlots

        slotCountsStr = CStr(d.numTakenSlots) & " / " & CStr(d.totalAvailableSlots)
        d.slotsCountsLabel.setString slotCountsStr
        WzPlayerBoxTabButton_recalculateSlotsLabel self
        WzPlayerBoxTabButton_recalculateTitleLabel self
        WzPlayerBoxTabButton_refreshTip self
    End If
End Sub

Private Sub WzPlayerBoxTabButton_refreshTip(ByVal self As WIDGET)
    Dim d As WzPlayerBoxTabButton, tipStr As String
    Set d = WzPlayerBoxTabButton_(self)
    tipStr = d.titleLabel.getString()
    tipStr = tipStr & vbLf
    tipStr = tipStr & "Joined:"
    tipStr = tipStr & " " & CStr(d.numTakenSlots) & "/" & CStr(d.totalAvailableSlots)
    tipStr = tipStr & " - "
    tipStr = tipStr & "Ready:"
    tipStr = tipStr & " " & CStr(d.numTakenSlotsReady) & "/" & CStr(d.numTakenSlots)
    self.setTip tipStr
End Sub

Private Sub WzPlayerBoxTabButton_recalculateSlotsLabel(ByVal self As WIDGET)
    Dim d As WzPlayerBoxTabButton, slotCountsLabelWidth As Long, slotCountsLabelHeight As Long, slotCountsLabelX0 As Long, slotCountsLabelY0 As Long
    Set d = WzPlayerBoxTabButton_(self)
    slotCountsLabelWidth = W_LABEL_getMaxLineWidth(d.slotsCountsLabel)
    slotCountsLabelHeight = d.slotsCountsLabel.idealHeight()
    slotCountsLabelX0 = self.width() - borderWidth - d.rightPadding - elementPadding - slotCountsLabelWidth
    slotCountsLabelY0 = (self.height() - slotCountsLabelHeight) \ 2
    d.slotsCountsLabel.setGeometry slotCountsLabelX0, slotCountsLabelY0, slotCountsLabelWidth, slotCountsLabelHeight
End Sub

Private Sub WzPlayerBoxTabButton_recalculateTitleLabel(ByVal self As WIDGET)
    Dim d As WzPlayerBoxTabButton, slotCountsLabelX0 As Long, titleLabelX0 As Long, titleLabelY0 As Long, titleLabelWidth As Long
    Set d = WzPlayerBoxTabButton_(self)
    ' size titleLabel to take up available space in the middle
    slotCountsLabelX0 = d.slotsCountsLabel.x()
    titleLabelX0 = borderWidth + elementPadding + (AtlasImage_width(FrontImages, IMAGE_LAMP_GREEN) \ 2) + elementPadding
    titleLabelY0 = (self.height() - d.titleLabel.height()) \ 2
    titleLabelWidth = slotCountsLabelX0 - elementPadding - titleLabelX0
    d.titleLabel.setGeometry titleLabelX0, titleLabelY0, titleLabelWidth, d.titleLabel.height()
End Sub

' setSlotReadyStatus(uint8_t numTakenSlotsReady)
Public Sub WzPlayerBoxTabButton_setSlotReadyStatus(ByVal self As WIDGET, ByVal new_numTakenSlotsReady As Long)
    WzPlayerBoxTabButton_(self).numTakenSlotsReady = new_numTakenSlotsReady And &HFF&
    WzPlayerBoxTabButton_refreshTip self
End Sub

Public Sub WzPlayerBoxTabButton_setSelected(ByVal self As WIDGET, ByVal new_selected As Boolean)
    WzPlayerBoxTabButton_(self).selected = new_selected
End Sub

Public Sub WzPlayerBoxTabButton_geometryChanged(ByVal self As WIDGET)
    WzPlayerBoxTabButton_recalculateSlotsLabel self
    WzPlayerBoxTabButton_recalculateTitleLabel self
End Sub

Public Sub WzPlayerBoxTabButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As WzPlayerBoxTabButton, x0 As Long, y0 As Long, w As Long, h As Long, highlight As Boolean, down As Boolean
    Dim boxBorder As Long, boxBackground As Long, readyStatusImage As AtlasImage
    Dim imageDisplayWidth As Long, imageDisplayHeight As Long, imageX0 As Long, imageY0 As Long
    Set d = WzPlayerBoxTabButton_(self)
    x0 = xOffset + self.x()
    y0 = yOffset + self.y()
    w = self.width()
    h = self.height()
    highlight = (self.getState() And WBUT_HIGHLIGHT) <> 0
    down = (self.getState() And (WBUT_DOWN Or WBUT_LOCK Or WBUT_CLICKLOCK)) <> 0

    ' draw box
    If Not d.selected Then boxBorder = psPalette(WZCOL_MENU_BACKGROUND) Else boxBorder = psPalette(WZCOL_MENU_BORDER)
    If highlight Then
        boxBorder = pal_RGBA(255, 255, 255, 255)
    End If
    If Not d.selected Then boxBackground = psPalette(WZCOL_MENU_BACKGROUND) Else boxBackground = psPalette(WZCOL_MENU_BORDER)
    pie_BoxFill x0, y0, x0 + w, y0 + h, boxBorder
    pie_BoxFill x0 + 1, y0 + 1, x0 + w - 1, y0 + h - 1, boxBackground
    If Not d.selected And (Not highlight Or down) Then
        pie_UniTransBoxFill x0 + 1, y0 + 1, x0 + w - 1, y0 + h - 1, pal_RGBA(0, 0, 0, 80)
    End If

    ' draw ready status image
    readyStatusImage = WzPlayerBoxTabButton_getImageForSlotsReadyStatus(self)
    If Not AtlasImage_isNull(readyStatusImage) Then
        imageDisplayWidth = AtlasImage_width(readyStatusImage.images, readyStatusImage.id) \ 2
        imageDisplayHeight = AtlasImage_height(readyStatusImage.images, readyStatusImage.id) \ 2
        imageX0 = x0 + borderWidth + elementPadding
        imageY0 = y0 + ((h - imageDisplayHeight) \ 2)
        iV_DrawImageFileAnisotropic readyStatusImage.images, readyStatusImage.id, imageX0, imageY0, imageDisplayWidth, imageDisplayHeight, 0, 255
    End If

    ' label drawing is handled by the embedded W_LABEL
    d.titleLabel.display x0, y0

    ' slot count drawing is handled by the embedded W_LABEL
    d.slotsCountsLabel.display x0, y0
End Sub

' MARK: - WzPlayerBoxOptionsButton

Public Sub WzPlayerBoxOptionsButton_construct(ByVal self As WIDGET)
    W_BUTTON_construct self
    self.kind = K_WzPlayerBoxOptionsButton ' (VB) virtual calls from here on reach WzPlayerBoxOptionsButton's overrides
End Sub

Public Function WzPlayerBoxOptionsButton_make() As WIDGET
    Dim widget As WIDGET, d As WzPlayerBoxOptionsButton, psWeakParent As Long
    Set widget = WIDGET_alloc(K_WzPlayerBoxOptionsButton)
    WzPlayerBoxOptionsButton_construct widget
    Set d = WzPlayerBoxOptionsButton_(widget)

    ' add the titleLabel
    Set d.titleLabel = W_LABEL_new()
    W_LABEL_setFont d.titleLabel, font_regular, psPalette(WZCOL_TEXT_BRIGHT)
    d.titleLabel.setString ChrW$(&H2699) ' "(gear)"
    psWeakParent = ObjPtr(widget)
    d.titleLabel.setCalcLayout WzFunc_new(FN_PBOB_TITLE_CALCLAYOUT, psWeakParent)
    W_LABEL_setTextAlignment d.titleLabel, WLAB_ALIGNCENTRE

    Set WzPlayerBoxOptionsButton_make = widget
End Function

Public Sub WzPlayerBoxOptionsButton_geometryChanged(ByVal self As WIDGET)
    If Not WzPlayerBoxOptionsButton_(self).titleLabel Is Nothing Then
        WzPlayerBoxOptionsButton_(self).titleLabel.callCalcLayout
    End If
End Sub

Public Sub WzPlayerBoxOptionsButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x0 As Long, y0 As Long, w As Long, h As Long, highlight As Boolean, down As Boolean, selected As Boolean
    Dim boxBorder As Long, boxBackground As Long
    x0 = xOffset + self.x()
    y0 = yOffset + self.y()
    w = self.width()
    h = self.height()
    highlight = (self.getState() And WBUT_HIGHLIGHT) <> 0
    down = (self.getState() And (WBUT_DOWN Or WBUT_LOCK Or WBUT_CLICKLOCK)) <> 0
    selected = False

    ' draw box
    If Not selected Then boxBorder = psPalette(WZCOL_MENU_BACKGROUND) Else boxBorder = psPalette(WZCOL_MENU_BORDER)
    If highlight Or down Then
        boxBorder = pal_RGBA(255, 255, 255, 255)
    End If
    If Not selected Then boxBackground = psPalette(WZCOL_MENU_BACKGROUND) Else boxBackground = psPalette(WZCOL_MENU_BORDER)
    pie_BoxFill x0, y0, x0 + w, y0 + h, boxBorder
    pie_BoxFill x0 + 1, y0 + 1, x0 + w - 1, y0 + h - 1, boxBackground
    If Not selected And (Not highlight Or down) Then
        pie_UniTransBoxFill x0 + 1, y0 + 1, x0 + w - 1, y0 + h - 1, pal_RGBA(0, 0, 0, 80)
    End If

    ' label drawing is handled by the embedded W_LABEL
    WzPlayerBoxOptionsButton_(self).titleLabel.display x0, y0
End Sub

' MARK: - WzPlayerBoxTabs

Public Sub WzPlayerBoxTabs_construct(ByVal self As WIDGET, ByVal titleUI As WzMultiplayerOptionsTitleUI)
    WIDGET_construct self, WIDG_UNSPECIFIED_TYPE
    self.kind = K_WzPlayerBoxTabs ' (VB) virtual calls from here on reach WzPlayerBoxTabs's overrides
    WzPlayerBoxTabs_(self).weakTitleUI = ObjPtr(titleUI)
    WzPlayerBoxTabs_(self).playerDisplayView_ = PlayerDisplayView_Players
End Sub

Public Sub WzPlayerBoxTabs_destruct(ByVal self As WIDGET)
    WzPlayerBoxTabs_closeOptionsOverlay self
End Sub

' static std::shared_ptr<WzPlayerBoxTabs> make(bool displayHostOptions, const std::shared_ptr<WzMultiplayerOptionsTitleUI>& titleUI)
Public Function WzPlayerBoxTabs_make(ByVal displayHostOptions As Boolean, ByVal titleUI As WzMultiplayerOptionsTitleUI) As WIDGET
    Dim widget As WIDGET, d As WzPlayerBoxTabs, weakSelf As Long, playersButton As WIDGET, spectatorsButton As WIDGET
    Set widget = WIDGET_alloc(K_WzPlayerBoxTabs)
    WzPlayerBoxTabs_construct widget, titleUI
    Set d = WzPlayerBoxTabs_(widget)

    weakSelf = ObjPtr(widget)

    ' add the "players" button
    Set playersButton = WzPlayerBoxTabButton_make("Players")
    WzPlayerBoxTabButton_setSelected playersButton, True
    d.tabButtons_push_back playersButton
    widget.attach playersButton
    W_BUTTON_addOnClickHandler playersButton, WzFunc_new(FN_PBT_PLAYERS_CLICK, weakSelf)

    ' add the "spectators" button
    Set spectatorsButton = WzPlayerBoxTabButton_make("Spectators")
    d.tabButtons_push_back spectatorsButton
    widget.attach spectatorsButton
    W_BUTTON_addOnClickHandler spectatorsButton, WzFunc_new(FN_PBT_SPECTATORS_CLICK, weakSelf)

    If displayHostOptions Then
        ' Add "gear" / "Host Options" button
        Set d.optionsButton = WzPlayerBoxOptionsButton_make() ' "(gear)"
        d.optionsButton.setTip "Host Options"
        widget.attach d.optionsButton
        W_BUTTON_addOnClickHandler d.optionsButton, WzFunc_new(FN_PBT_OPTIONS_CLICK)
    End If

    widget.setCalcLayout WzFunc_new(FN_PBT_CALCLAYOUT)

    WzPlayerBoxTabs_refreshData widget

    Set WzPlayerBoxTabs_make = widget
End Function

Public Sub WzPlayerBoxTabs_geometryChanged(ByVal self As WIDGET)
    WzPlayerBoxTabs_recalculateTabLayout self
End Sub

Public Sub WzPlayerBoxTabs_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    WzPlayerBoxTabs_refreshData self
End Sub

Public Function WzPlayerBoxTabs_getDisplayView(ByVal self As WIDGET) As WzPlayerBoxTabs_PlayerDisplayView
    WzPlayerBoxTabs_getDisplayView = WzPlayerBoxTabs_(self).playerDisplayView_
End Function

Public Sub WzPlayerBoxTabs_setDisplayView(ByVal self As WIDGET, ByVal view As WzPlayerBoxTabs_PlayerDisplayView)
    Dim strongTitleUI As WzMultiplayerOptionsTitleUI
    WzPlayerBoxTabs_(self).playerDisplayView_ = view

    Set strongTitleUI = titleUIFromPtr(WzPlayerBoxTabs_(self).weakTitleUI)
    If Not (Not strongTitleUI Is Nothing) Then ASSERT_FAILURE "strongTitleUI != nullptr", "src/titleui/widgets/lobbyplayerboxtabs.cpp:846", "setDisplayView", "No parent title UI": Exit Sub
    strongTitleUI.updatePlayers
End Sub

Private Sub WzPlayerBoxTabs_setSelectedTab(ByVal self As WIDGET, ByVal index As Long)
    Dim d As WzPlayerBoxTabs, i As Long
    Set d = WzPlayerBoxTabs_(self)
    For i = 0 To d.tabButtons_size() - 1
        WzPlayerBoxTabButton_setSelected d.tabButtons(i), i = index
    Next
End Sub

Private Sub WzPlayerBoxTabs_refreshData(ByVal self As WIDGET)
    Dim d As WzPlayerBoxTabs, i As Long
    Dim takenPlayerSlots As Long, totalPlayerSlots As Long, readyPlayers As Long
    Dim takenSpectatorOnlySlots As Long, totalSpectatorOnlySlots As Long, readySpectatorOnlySlots As Long
    Set d = WzPlayerBoxTabs_(self)
    ' update currently-selected tab
    Select Case d.playerDisplayView_
    Case PlayerDisplayView_Players
        WzPlayerBoxTabs_setSelectedTab self, 0
    Case PlayerDisplayView_Spectators
        WzPlayerBoxTabs_setSelectedTab self, 1
    End Select

    ' update counts of players & spectators (uint8_t)
    For i = 0 To MAX_CONNECTED_PLAYERS - 1
        ' PLAYER const &p = NetPlay.players[i];
        If NetPlay_players_ai(i) = AI_CLOSED Then
            ' closed slot - skip
            GoTo continue_
        End If
        If isSpectatorOnlySlot(i) Then
            If (i >= NetPlay_players_size()) Then
                ' the only slots displayable beyond game.maxPlayers are spectator slots
                GoTo continue_
            ElseIf Not (NetPlay_players_isSpectator(i) And NetPlay_players_ai(i) = AI_OPEN) Then
                GoTo continue_
            End If
        End If
        If isSpectatorOnlySlot(i) Then
            If NetPlay_players_ai(i) = AI_OPEN Then
                totalSpectatorOnlySlots = (totalSpectatorOnlySlots + 1) And &HFF&
                If NetPlay_players_allocated(i) Then
                    takenSpectatorOnlySlots = (takenSpectatorOnlySlots + 1) And &HFF&
                    If NetPlay_players_ready(i) Then
                        readySpectatorOnlySlots = (readySpectatorOnlySlots + 1) And &HFF&
                    End If
                End If
            End If
            GoTo continue_
        End If
        totalPlayerSlots = (totalPlayerSlots + 1) And &HFF&
        If NetPlay_players_ai(i) = AI_OPEN Then
            If NetPlay_players_allocated(i) Then
                takenPlayerSlots = (takenPlayerSlots + 1) And &HFF&
                If NetPlay_players_ready(i) Then
                    readyPlayers = (readyPlayers + 1) And &HFF&
                End If
            End If
        Else
            If Not (Not NetPlay_players_allocated(i)) Then ASSERT_FAILURE "!p.allocated", "src/titleui/widgets/lobbyplayerboxtabs.cpp:922", "refreshData", "Expecting AI bots to not be flagged as allocated?"
            takenPlayerSlots = (takenPlayerSlots + 1) And &HFF&
            readyPlayers = (readyPlayers + 1) And &HFF& ' AI slots are always "ready"
        End If
continue_:
    Next

    ' players button
    WzPlayerBoxTabButton_setSlotCounts d.tabButtons(0), takenPlayerSlots, totalPlayerSlots
    WzPlayerBoxTabButton_setSlotReadyStatus d.tabButtons(0), readyPlayers

    ' spectators button
    WzPlayerBoxTabButton_setSlotCounts d.tabButtons(1), takenSpectatorOnlySlots, totalSpectatorOnlySlots
    WzPlayerBoxTabButton_setSlotReadyStatus d.tabButtons(1), readySpectatorOnlySlots
End Sub

Private Sub WzPlayerBoxTabs_recalculateTabLayout(ByVal self As WIDGET)
    Dim d As WzPlayerBoxTabs, optionsButtonSize As Long, rightPaddingForAllButtons As Long, availableWidthForButtons As Long
    Dim baseWidthPerButton As Long, x0 As Long, i As Long, button As WIDGET, buttonNewWidth As Long, optionsButtonX0 As Long
    Set d = WzPlayerBoxTabs_(self)
    If Not d.optionsButton Is Nothing Then optionsButtonSize = self.height() Else optionsButtonSize = 0
    rightPaddingForAllButtons = optionsButtonSize
    availableWidthForButtons = self.width() - rightPaddingForAllButtons ' width minus internal and external button padding
    baseWidthPerButton = availableWidthForButtons \ d.tabButtons_size()
    x0 = 0
    For i = 0 To d.tabButtons_size() - 1
        Set button = d.tabButtons(i)
        buttonNewWidth = baseWidthPerButton
        button.setGeometry x0, 0, buttonNewWidth, self.height()
        x0 = button.x() + button.width()
    Next

    If Not d.optionsButton Is Nothing Then
        optionsButtonX0 = self.width() - optionsButtonSize
        d.optionsButton.setGeometry optionsButtonX0, 0, optionsButtonSize, optionsButtonSize
    End If
End Sub

Private Sub WzPlayerBoxTabs_closeOptionsOverlay(ByVal self As WIDGET)
    If Not WzPlayerBoxTabs_(self).optionsOverlayScreen Is Nothing Then
        widgRemoveOverlayScreen WzPlayerBoxTabs_(self).optionsOverlayScreen
        Set WzPlayerBoxTabs_(self).optionsOverlayScreen = Nothing
    End If
End Sub

Private Function WzPlayerBoxTabs_createOptionsPopoverForm(ByVal self As WIDGET) As WIDGET
    Dim psWeakSelf As Long, optionsPopOver As WIDGET
    psWeakSelf = ObjPtr(self)
    Set optionsPopOver = LobbyHostOptionsForm_make(WzPlayerBoxTabs_(self).weakTitleUI, WzFunc_new(FN_PBT_POPOVER_CLOSE, psWeakSelf))
    optionsPopOver.setGeometry 0, 0, optionsPopOver.idealWidth(), optionsPopOver.idealHeight()
    Set WzPlayerBoxTabs_createOptionsPopoverForm = optionsPopOver
End Function

Private Sub WzPlayerBoxTabs_displayOptionsOverlay(ByVal self As WIDGET, ByVal psParent As WIDGET)
    Dim d As WzPlayerBoxTabs, lockedScreen As W_SCREEN, newRootFrm As WIDGET, psWeakOptionsOverlayScreen As Long, psWeakParent As Long
    Dim optionsPopOver As WIDGET
    Set d = WzPlayerBoxTabs_(self)
    Set lockedScreen = self.screenPointer
    If Not (Not lockedScreen Is Nothing) Then ASSERT_FAILURE "lockedScreen != nullptr", "src/titleui/widgets/lobbyplayerboxtabs.cpp:983", "displayOptionsOverlay", "The WzPlayerBoxTabs does not have an associated screen pointer?"

    ' Initialize the options overlay screen
    Set d.optionsOverlayScreen = W_SCREEN_make()
    Set newRootFrm = W_FULLSCREENOVERLAY_CLICKFORM_make()
    psWeakOptionsOverlayScreen = ObjPtr(d.optionsOverlayScreen)
    psWeakParent = ObjPtr(self)
    Set W_FULLSCREENOVERLAY_CLICKFORM_(newRootFrm).onClickedFunc = WzFunc_new(FN_PBT_OVERLAY_CLICKED, psWeakOptionsOverlayScreen, psWeakParent)
    Set W_FULLSCREENOVERLAY_CLICKFORM_(newRootFrm).onCancelPressed = W_FULLSCREENOVERLAY_CLICKFORM_(newRootFrm).onClickedFunc
    d.optionsOverlayScreen.psForm.attach newRootFrm

    ' Create the pop-over form
    Set optionsPopOver = WzPlayerBoxTabs_createOptionsPopoverForm(self)
    newRootFrm.attach optionsPopOver

    ' Position the pop-over form
    optionsPopOver.setCalcLayout WzFunc_new(FN_PBT_POPOVER_CALCLAYOUT)

    widgRegisterOverlayScreenOnTopOfScreen d.optionsOverlayScreen, lockedScreen
    widgScheduleTask WzFunc_new(FN_PBT_CLEAR_FOCUS, lockedScreen)
End Sub

' MARK: - the lambdas

Public Function lobbyplayerboxtabs_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim result As OptionChoices, strongTitleUI As WzMultiplayerOptionsTitleUI, newLocked As MultiplayOptionsLocked
    Dim optAvailable As Boolean, hasValue As Boolean, reason As String, susHelpDescription As String, allowAllDescription As String
    Dim botProt As Long, proxyIPs As Long, hostingIPs As Long, psParent As WIDGET, psWidget As WIDGET, button As WIDGET
    Dim strongSelf As WIDGET, psOverlayScreen As W_SCREEN, strongParent As WIDGET, lockedScreen As W_SCREEN
    Dim popOverWidth As Long, popOverHeight As Long, popOverX0 As Long, popOverY0 As Long, v As Long
    Select Case f.id
    Case FN_LHOF_CLOSE_CLICK                    ' [onClose](W_BUTTON&)
        widgScheduleTask WzFunc_new(FN_LHOF_CLOSE_TASK, f.cap(0))

    Case FN_LHOF_CLOSE_TASK                     ' [onClose]()
        WzFunc_call f.cap(0)

    Case FN_IsGameHostedInLobby                 ' IsGameHostedInLobby(const OptionInfo&)
        Set lobbyplayerboxtabs_c_fn = IsGameHostedInLobby(args(0))

    Case FN_LOBBYHOSTJOINOPT_SUPPORTED          ' [optType](const OptionInfo&) -> OptionInfo::AvailabilityResult
        hasValue = NET_getHostJoinOptionSupportedByLobby(f.cap(0), optAvailable)
        ' result.available_ = optAvailable.value_or(false);
        If Not hasValue Then optAvailable = False
        If hasValue Then
            reason = "This option is not supported by the current lobby."
        Else
            reason = "Awaiting lobby information..."
        End If
        Set lobbyplayerboxtabs_c_fn = makeAvailabilityResult(optAvailable, reason)

    Case FN_LHOF_SPECJOIN_POPULATE              ' []() (Spectator Join)
        Set result = New OptionChoices
        result.addChoice "Off", "", False
        result.addChoice "Allow", "", True
        result.setCurrentIdxForValue hasOpenSpectatorOnlySlots()
        Set lobbyplayerboxtabs_c_fn = result

    Case FN_LHOF_SPECJOIN_SELECT                ' [weakTitleUICopy](const auto& newValue) -> bool
        Set strongTitleUI = titleUIFromPtr(f.cap(0))
        If Not (Not strongTitleUI Is Nothing) Then ASSERT_FAILURE "strongTitleUI != nullptr", "src/titleui/widgets/lobbyplayerboxtabs.cpp:260", "makeHostOptionsForm", "No Title UI?": lobbyplayerboxtabs_c_fn = False: Exit Function
        enableSpectatorJoin CBool(args(0))
        strongTitleUI.updatePlayers
        lobbyplayerboxtabs_c_fn = True

    Case FN_LHOF_LOCKTEAMS_POPULATE             ' []() (Teams)
        Set result = New OptionChoices
        result.addChoice "Locked", "", True
        result.addChoice "Allow Change", "", False
        result.setCurrentIdxForValue getLockedOptions().teams
        Set lobbyplayerboxtabs_c_fn = result

    Case FN_LHOF_LOCKTEAMS_SELECT               ' [weakTitleUICopy](const auto& newValue) -> bool
        newLocked = getLockedOptions()
        newLocked.teams = CBool(args(0))
        updateLockedOptionsOnHost newLocked
        lobbyplayerboxtabs_c_fn = True

    Case FN_LHOF_LOCKPOS_POPULATE               ' []() (Position)
        Set result = New OptionChoices
        result.addChoice "Locked", "", True
        result.addChoice "Allow Change", "", False
        result.setCurrentIdxForValue getLockedOptions().position
        Set lobbyplayerboxtabs_c_fn = result

    Case FN_LHOF_LOCKPOS_SELECT                 ' [weakTitleUICopy](const auto& newValue) -> bool
        newLocked = getLockedOptions()
        newLocked.position = CBool(args(0))
        updateLockedOptionsOnHost newLocked
        lobbyplayerboxtabs_c_fn = True

    Case FN_LHOF_INACTIVITY_SELECT              ' [](const auto& newValue) -> bool
        war_setMPInactivityMinutes CLng(args(0))
        game.inactivityMinutes = war_getMPInactivityMinutes()
        sendOptions
        lobbyplayerboxtabs_c_fn = True

    Case FN_LHOF_LAGKICK_SELECT                 ' [](const auto& newValue) -> bool
        war_setAutoLagKickSeconds CLng(args(0))
        lobbyplayerboxtabs_c_fn = True

    Case FN_LHOF_TIMELIMIT_SELECT               ' [](const auto& newValue) -> bool
        war_setMPGameTimeLimitMinutes CLng(args(0))
        game.gameTimeLimitMinutes = war_getMPGameTimeLimitMinutes()
        sendOptions
        lobbyplayerboxtabs_c_fn = True

    Case FN_LHOF_HOSTJOINOPT_POPULATE           ' []() (Bot Protection, Proxy / VPN Connections, Hosting Provider IPs)
        Set result = New OptionChoices
        Select Case f.cap(0)
        Case HostJoinOptionType_BotProt
            result.addChoice "Block Known Bots", "Block known bots or automated tools.", HostJoinOptionValue_BlockAll
            result.addChoice "Disabled", "", HostJoinOptionValue_Allow
        Case HostJoinOptionType_ProxyIPs
            susHelpDescription = "Block suspicious connections from proxies / VPNs."
            susHelpDescription = susHelpDescription & " "
            susHelpDescription = susHelpDescription & "(Recommended)"
            result.addChoice "Block All", "Block all connections from proxies / VPNs.", HostJoinOptionValue_BlockAll
            result.addChoice "Block Suspicious", susHelpDescription, HostJoinOptionValue_BlockSuspicious
            result.addChoice "Allow", "", HostJoinOptionValue_Allow
        Case HostJoinOptionType_HostingIPs
            result.addChoice "Block All", "Block all connections from hosting provider IPs.", HostJoinOptionValue_BlockAll
            result.addChoice "Block Suspicious", "Block suspicious connections from hosting provider IPs.", HostJoinOptionValue_BlockSuspicious
            result.addChoice "Allow", "", HostJoinOptionValue_Allow
        End Select
        If lobbyHostJoinOptionEnabled(f.cap(0)) Then v = hostJoinOptionValue(f.cap(0)) Else v = HostJoinOptionValue_Allow
        result.setCurrentIdxForValue v
        Set lobbyplayerboxtabs_c_fn = result

    Case FN_LHOF_HOSTJOINOPT_SELECT             ' [weakTitleUICopy](const auto& newValue) -> bool
        If Not lobbyHostJoinOptionEnabled(f.cap(0)) Then lobbyplayerboxtabs_c_fn = False: Exit Function
        NETgetLobbyHostJoinOptions botProt, proxyIPs, hostingIPs
        Select Case f.cap(0)
        Case HostJoinOptionType_BotProt: botProt = CLng(args(0))
        Case HostJoinOptionType_ProxyIPs: proxyIPs = CLng(args(0))
        Case HostJoinOptionType_HostingIPs: hostingIPs = CLng(args(0))
        End Select
        NETsetLobbyHostJoinOptions botProt, proxyIPs, hostingIPs
        lobbyplayerboxtabs_c_fn = True

    Case FN_LHOF_DIRECTJOIN_POPULATE            ' []() (Lobby Join Bypass)
        Set result = New OptionChoices
        allowAllDescription = "Allow any IP to direct join. Bypasses lobby anti-abuse protections."
        allowAllDescription = allowAllDescription & " "
        allowAllDescription = allowAllDescription & "(Not Recommended)"
        result.addChoice "Off", "Do not allow direct joins. Require joining through the lobby.", LobbyHostDirectJoinOption_None
        result.addChoice "Local", "Allow local / LAN IPs to direct join. Require others to join through the lobby.", LobbyHostDirectJoinOption_Local
        result.addChoice "Allow All", allowAllDescription, LobbyHostDirectJoinOption_All
        If hasGameCurrentlyHostedInLobby() Then v = NETgetLobbyHostDirectJoinOption() Else v = LobbyHostDirectJoinOption_All
        result.setCurrentIdxForValue v
        Set lobbyplayerboxtabs_c_fn = result

    Case FN_LHOF_DIRECTJOIN_SELECT              ' [](const auto& newValue) -> bool
        If Not hasGameCurrentlyHostedInLobby() Then lobbyplayerboxtabs_c_fn = False: Exit Function
        NETsetLobbyHostDirectJoinOption CLng(args(0))
        lobbyplayerboxtabs_c_fn = True

    Case FN_PBOB_TITLE_CALCLAYOUT               ' [psWeakParent](WIDGET *psWidget)
        Set psWidget = args(0)
        Set psParent = WIDGET_fromPtr(f.cap(0))
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/titleui/widgets/lobbyplayerboxtabs.cpp:705", "make", "Parent is null": Exit Function
        psWidget.setGeometry 0, 0, psParent.width(), psParent.height()

    Case FN_PBT_PLAYERS_CLICK                   ' [weakSelf](W_BUTTON& button)
        widgScheduleTask WzFunc_new(FN_PBT_SETVIEW_TASK, f.cap(0), PlayerDisplayView_Players)

    Case FN_PBT_SPECTATORS_CLICK                ' [weakSelf](W_BUTTON& button)
        widgScheduleTask WzFunc_new(FN_PBT_SETVIEW_TASK, f.cap(0), PlayerDisplayView_Spectators)

    Case FN_PBT_SETVIEW_TASK                    ' [weakSelf]() (cap(1): the view the button stands for)
        Set strongSelf = WIDGET_fromPtr(f.cap(0))
        If f.cap(1) = PlayerDisplayView_Players Then
            If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/titleui/widgets/lobbyplayerboxtabs.cpp:784", "make", "Self is null?": Exit Function
        Else
            If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/titleui/widgets/lobbyplayerboxtabs.cpp:796", "make", "Self is null?": Exit Function
        End If
        WzPlayerBoxTabs_setDisplayView strongSelf, f.cap(1)

    Case FN_PBT_OPTIONS_CLICK                   ' [](W_BUTTON& button)
        Set button = args(0)
        Set psParent = button.parent()
        If Not WIDGET_isKindOf(psParent, K_WzPlayerBoxTabs) Then Set psParent = Nothing
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/titleui/widgets/lobbyplayerboxtabs.cpp:809", "make", "No parent": Exit Function
        ' Display a "pop-over" options menu
        WzPlayerBoxTabs_displayOptionsOverlay psParent, button

    Case FN_PBT_CALCLAYOUT                      ' LAMBDA_CALCLAYOUT_SIMPLE({...})
        Set psWidget = args(0)
        If Not WIDGET_isKindOf(psWidget, K_WzPlayerBoxTabs) Then Set psWidget = Nothing
        If Not (Not psWidget Is Nothing) Then ASSERT_FAILURE "psPlayerBoxTabs.get() != nullptr", "src/titleui/widgets/lobbyplayerboxtabs.cpp:817", "make", "Wrong type of psWidget??": Exit Function
        WzPlayerBoxTabs_recalculateTabLayout psWidget

    Case FN_PBT_POPOVER_CLOSE                   ' [psWeakSelf]()
        Set strongSelf = WIDGET_fromPtr(f.cap(0))
        If Not strongSelf Is Nothing Then
            WzPlayerBoxTabs_closeOptionsOverlay strongSelf
        End If

    Case FN_PBT_OVERLAY_CLICKED                 ' [psWeakOptionsOverlayScreen, psWeakParent]()
        Set psOverlayScreen = W_SCREEN_fromPtr(f.cap(0))
        If Not psOverlayScreen Is Nothing Then
            widgRemoveOverlayScreen psOverlayScreen
        End If
        ' Destroy Options overlay / overlay screen
        Set strongParent = WIDGET_fromPtr(f.cap(1))
        If Not strongParent Is Nothing Then
            Set WzPlayerBoxTabs_(strongParent).optionsOverlayScreen = Nothing
        End If

    Case FN_PBT_POPOVER_CALCLAYOUT              ' [](WIDGET *psWidget)
        Set psWidget = args(0)
        popOverWidth = psWidget.idealWidth()
        If popOverWidth < 400 Then popOverWidth = 400
        If screenWidth - 40 < popOverWidth Then popOverWidth = screenWidth - 40
        popOverHeight = psWidget.idealHeight()
        If screenHeight - 40 < popOverHeight Then popOverHeight = screenHeight - 40

        ' Centered in screen
        popOverX0 = (screenWidth - popOverWidth) \ 2
        If popOverX0 < 0 Then
            popOverX0 = 0
        End If
        popOverY0 = (screenHeight - popOverHeight) \ 2
        If popOverY0 < 0 Then
            popOverY0 = 0
        End If

        psWidget.setGeometry popOverX0, popOverY0, popOverWidth, popOverHeight

    Case FN_PBT_CLEAR_FOCUS                     ' [lockedScreen]()
        ' clear focus from the underlying chatbox, which aggressively takes focus
        Set lockedScreen = f.cap(0)
        lockedScreen.setFocus Nothing
    End Select
End Function
