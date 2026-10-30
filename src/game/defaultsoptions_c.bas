Attribute VB_Name = "defaultsoptions_c"
' src/titleui/options/defaultsoptions.cpp: the Defaults Options form (makeDefaultsOptionsForm), the
' hosting options' populate functions (make...OptionsDropdownPopulateFunc) and the widget class
' PlayerColorOptionsSelector, whole. The widget class follows docs/CONVENTIONS.md, "Widgets" (its
' fields: PlayerColorOptionsSelector.cls). (VB) An optional<unsigned> tc is a Variant: Empty is
' nullopt.
Option Explicit

' (VB) widget class PlayerColorOptionsSelector : WIDGET

Public Const FN_IsLobbyIPv6Enabled As Long = 4550           ' OptionInfo::AvailabilityResult IsLobbyIPv6Enabled(const OptionInfo&)
Public Const FN_PCOS_colorClick As Long = 4551              ' PlayerColorOptionsSelector::addColor, [weakSelf, tc](W_BUTTON&)
Public Const FN_inactivityTimeout_populate As Long = 4552   ' makeInactivityTimeoutOptionsDropdownPopulateFunc, []()
Public Const FN_lagKick_populate As Long = 4553             ' makeLagKickOptionsDropdownPopulateFunc, []()
Public Const FN_gameTimeLimit_populate As Long = 4554       ' makeGameTimeLimitOptionsDropdownPopulateFunc, []()
Public Const FN_defaults_spColor_get As Long = 4555         ' makeDefaultsOptionsForm, "defaults.campaign.unitColor": [validSPColors]() -> optional<unsigned>
Public Const FN_defaults_spColor_select As Long = 4556      ' [](optional<unsigned> tc) -> bool
Public Const FN_defaults_difficulty_populate As Long = 4557 ' "defaults.campaign.difficultyLevel"
Public Const FN_defaults_difficulty_select As Long = 4558
Public Const FN_defaults_mpColor_get As Long = 4559         ' "defaults.multiplayer.unitColor"
Public Const FN_defaults_mpColor_select As Long = 4560
Public Const FN_defaults_defaultAI_populate As Long = 4561  ' "defaults.multiplayer.defaultAI"
Public Const FN_defaults_defaultAI_select As Long = 4562
Public Const FN_defaults_enableipv6_populate As Long = 4563 ' "defaults.lobby.enableipv6"
Public Const FN_defaults_enableipv6_select As Long = 4564
Public Const FN_defaults_preferipv6_populate As Long = 4565 ' "defaults.lobby.preferipv6"
Public Const FN_defaults_preferipv6_select As Long = 4566
Public Const FN_defaults_portMapping_populate As Long = 4567 ' "defaults.hosting.portMapping"
Public Const FN_defaults_portMapping_select As Long = 4568
Public Const FN_defaults_chat_populate As Long = 4569       ' "defaults.hosting.chat"
Public Const FN_defaults_chat_select As Long = 4570
Public Const FN_defaults_inactivityTimeout_select As Long = 4571 ' "defaults.hosting.inactivityTimeout"
Public Const FN_defaults_lagKick_select As Long = 4572      ' "defaults.hosting.lagKick"
Public Const FN_defaults_spectatorSlots_populate As Long = 4573 ' "defaults.hosting.spectatorSlots"
Public Const FN_defaults_spectatorSlots_select As Long = 4574
Public Const FN_defaults_onPlayerLeave_populate As Long = 4575 ' "defaults.hosting.onPlayerLeave"
Public Const FN_defaults_onPlayerLeave_select As Long = 4576
Public Const FN_defaults_gameTimeLimit_select As Long = 4577 ' "defaults.hosting.gameTimeLimit"

' PlayerColorOptionsSelector's const members
Private Const PlayerColorOptionsSelector_outerPaddingX As Long = 8
Private Const PlayerColorOptionsSelector_outerPaddingY As Long = 5

' MARK: -

Public Sub PlayerColorOptionsSelector_construct(ByVal self As WIDGET)
    Dim d As PlayerColorOptionsSelector
    WIDGET_construct self
    self.kind = K_PlayerColorOptionsSelector
    Set d = PlayerColorOptionsSelector_(self)
    Set d.onChangeHandlers = New Collection
    Set d.colorButtons = New Collection
    d.buttonWidth = 0
    d.buttonHeight = 0
    d.buttonMargin = 5
End Sub

' static std::shared_ptr<PlayerColorOptionsSelector> make(GetCurrentValueFunc getCurrVal, OnSelectColorFunc onSelect)
' (GetCurrentValueFunc: optional<unsigned>(); OnSelectColorFunc: bool (optional<unsigned> tc))
Public Function PlayerColorOptionsSelector_make(ByVal getCurrVal As WzFunc, ByVal onSelect As WzFunc) As WIDGET
    Dim result As WIDGET, d As PlayerColorOptionsSelector
    Set result = WIDGET_alloc(K_PlayerColorOptionsSelector)
    PlayerColorOptionsSelector_construct result
    Set d = PlayerColorOptionsSelector_(result)
    Set d.getCurrentValueFunc = getCurrVal
    Set d.onSelectFunc = onSelect
    d.buttonWidth = iV_GetImageWidth(FrontImages, IMAGE_PLAYERN)
    d.buttonHeight = iV_GetImageHeight(FrontImages, IMAGE_PLAYERN)
    Set PlayerColorOptionsSelector_make = result
End Function

Private Function PlayerColorOptionsSelector_optTcToMultiButTc(ByVal self As WIDGET, ByVal tc As Variant) As Long
    If IsEmpty(tc) Then PlayerColorOptionsSelector_optTcToMultiButTc = MAX_PLAYERS + 1 Else PlayerColorOptionsSelector_optTcToMultiButTc = CLng(tc)
End Function

' void addColor(optional<unsigned> tc)
Public Sub PlayerColorOptionsSelector_addColor(ByVal self As WIDGET, ByVal tc As Variant)
    Dim d As PlayerColorOptionsSelector, button As WIDGET, weakSelf As Long
    Set d = PlayerColorOptionsSelector_(self)
    Set button = makeMultiBut(0, d.buttonWidth, d.buttonHeight, vbNullString, IMAGE_PLAYERN, IMAGE_PLAYERX, 1, PlayerColorOptionsSelector_optTcToMultiButTc(self, tc))
    self.attach button
    weakSelf = ObjPtr(self)
    W_BUTTON_addOnClickHandler button, WzFunc_new(FN_PCOS_colorClick, weakSelf, tc)
    d.colorButtons.Add button
    PlayerColorOptionsSelector_updateSelectedItem self, WzFunc_call(d.getCurrentValueFunc)
End Sub

Public Sub PlayerColorOptionsSelector_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    ' no-op
End Sub

Public Sub PlayerColorOptionsSelector_geometryChanged(ByVal self As WIDGET)
    Dim d As PlayerColorOptionsSelector, w As Long, h As Long, availableWidth As Long, availableHeight As Long
    Dim rowsRequired As Long, neededHeight As Long, buttonX0 As Long, buttonY0 As Long, b As WIDGET, i As Long
    Set d = PlayerColorOptionsSelector_(self)
    w = self.width()
    h = self.height()

    If w = 0 Or h = 0 Then
        Exit Sub
    End If

    availableWidth = w - (PlayerColorOptionsSelector_outerPaddingX * 2)
    availableHeight = h - (PlayerColorOptionsSelector_outerPaddingY * 2)

    rowsRequired = PlayerColorOptionsSelector_rowsRequiredForInnerWidth(self, w)
    neededHeight = PlayerColorOptionsSelector_innerHeightRequiredForNumRows(self, rowsRequired)

    buttonX0 = 0    ' TODO: Right-align everything?
    buttonY0 = MAX_(0, (availableHeight - neededHeight) \ 2)
    For i = 1 To d.colorButtons.Count
        Set b = d.colorButtons(i)
        b.setGeometry PlayerColorOptionsSelector_outerPaddingX + buttonX0, PlayerColorOptionsSelector_outerPaddingY + buttonY0, d.buttonWidth, d.buttonHeight
        buttonX0 = buttonX0 + d.buttonWidth + d.buttonMargin
        If buttonX0 + d.buttonWidth > availableWidth Then
            buttonY0 = buttonY0 + d.buttonHeight + d.buttonMargin
            buttonX0 = 0
        End If
    Next
End Sub

Public Function PlayerColorOptionsSelector_idealWidth(ByVal self As WIDGET) As Long
    PlayerColorOptionsSelector_idealWidth = (PlayerColorOptionsSelector_outerPaddingX * 2) + 157
End Function

Public Function PlayerColorOptionsSelector_idealHeight(ByVal self As WIDGET) As Long
    PlayerColorOptionsSelector_idealHeight = (PlayerColorOptionsSelector_outerPaddingY * 2) + PlayerColorOptionsSelector_heightRequiredForWidth(self, 157)
End Function

Private Function PlayerColorOptionsSelector_rowsRequiredForInnerWidth(ByVal self As WIDGET, ByVal w As Long) As Long
    Dim d As PlayerColorOptionsSelector, totalRequiredWidthForButtons As Long, n As Long
    Set d = PlayerColorOptionsSelector_(self)
    totalRequiredWidthForButtons = 0
    n = d.colorButtons.Count
    If n <> 0 Then
        totalRequiredWidthForButtons = (d.buttonWidth * n) + (d.buttonMargin * (n - 1))
    End If
    PlayerColorOptionsSelector_rowsRequiredForInnerWidth = ((totalRequiredWidthForButtons - 1) \ w) + 1
End Function

Private Function PlayerColorOptionsSelector_innerHeightRequiredForNumRows(ByVal self As WIDGET, ByVal rows As Long) As Long
    Dim d As PlayerColorOptionsSelector, result As Long
    Set d = PlayerColorOptionsSelector_(self)
    result = (rows * d.buttonHeight)
    If rows > 1 Then
        result = result + (rows - 1) * d.buttonMargin
    End If
    PlayerColorOptionsSelector_innerHeightRequiredForNumRows = result
End Function

Public Function PlayerColorOptionsSelector_heightRequiredForWidth(ByVal self As WIDGET, ByVal w As Long) As Long
    PlayerColorOptionsSelector_heightRequiredForWidth = PlayerColorOptionsSelector_innerHeightRequiredForNumRows(self, PlayerColorOptionsSelector_rowsRequiredForInnerWidth(self, w))
End Function

Private Sub PlayerColorOptionsSelector_updateSelectedItem(ByVal self As WIDGET, ByVal tc As Variant)
    Dim d As PlayerColorOptionsSelector, b As WIDGET, i As Long
    Set d = PlayerColorOptionsSelector_(self)
    For i = 1 To d.colorButtons.Count
        Set b = d.colorButtons(i)
        If WzMultiButton_(b).tc = PlayerColorOptionsSelector_optTcToMultiButTc(self, tc) Then b.setState WBUT_LOCK Else b.setState 0
    Next
End Sub

Private Sub PlayerColorOptionsSelector_handleColorPressed(ByVal self As WIDGET, ByVal tc As Variant)
    Dim d As PlayerColorOptionsSelector, onChangeFunc As WzFunc, i As Long
    Set d = PlayerColorOptionsSelector_(self)
    If OptionValue_equal(tc, WzFunc_call(d.getCurrentValueFunc)) Then
        Exit Sub
    End If

    If Not CBool(WzFunc_call(d.onSelectFunc, tc)) Then
        Exit Sub
    End If
    PlayerColorOptionsSelector_updateSelectedItem self, tc

    For i = 1 To d.onChangeHandlers.Count
        Set onChangeFunc = d.onChangeHandlers(i)
        WzFunc_call onChangeFunc, self
    Next
End Sub

Public Function PlayerColorOptionsSelector_ovc_is(ByVal self As WIDGET) As Boolean
    PlayerColorOptionsSelector_ovc_is = True
End Function

Public Sub PlayerColorOptionsSelector_ovc_update(ByVal self As WIDGET, Optional ByVal force As Boolean = False)
    PlayerColorOptionsSelector_updateSelectedItem self, WzFunc_call(PlayerColorOptionsSelector_(self).getCurrentValueFunc)
End Sub

Public Sub PlayerColorOptionsSelector_ovc_informAvailable(ByVal self As WIDGET, ByVal isAvailable As Boolean)
    ' future todo: could disable the buttons
End Sub

Public Sub PlayerColorOptionsSelector_ovc_addOnChangeHandler(ByVal self As WIDGET, ByVal handler As WzFunc)
    PlayerColorOptionsSelector_(self).onChangeHandlers.Add handler
End Sub

' MARK: -

Public Function IsLobbyIPv6Enabled(ByVal info As OptionInfo) As OptionInfo_AvailabilityResult
    Dim result As OptionInfo_AvailabilityResult
    Set result = New OptionInfo_AvailabilityResult
    result.available_ = Not war_getLobbyDisableIPv6()
    result.localizedUnavailabilityReason = "Only useful if Enable IPv6 is On."
    Set IsLobbyIPv6Enabled = result
End Function

' MARK: -

' OptionsDropdown<uint32_t>::PopulateFunc makeInactivityTimeoutOptionsDropdownPopulateFunc()
Public Function makeInactivityTimeoutOptionsDropdownPopulateFunc() As WzFunc
    Set makeInactivityTimeoutOptionsDropdownPopulateFunc = WzFunc_new(FN_inactivityTimeout_populate)
End Function

' OptionsDropdown<int>::PopulateFunc makeLagKickOptionsDropdownPopulateFunc()
Public Function makeLagKickOptionsDropdownPopulateFunc() As WzFunc
    Set makeLagKickOptionsDropdownPopulateFunc = WzFunc_new(FN_lagKick_populate)
End Function

' OptionsDropdown<uint32_t>::PopulateFunc makeGameTimeLimitOptionsDropdownPopulateFunc()
Public Function makeGameTimeLimitOptionsDropdownPopulateFunc() As WzFunc
    Set makeGameTimeLimitOptionsDropdownPopulateFunc = WzFunc_new(FN_gameTimeLimit_populate)
End Function

' (VB) fmt::format("{:#}", value): the shortest form, always with a decimal point (2 -> "2.0")
Private Function fmtAlternateFloat(ByVal value As Single) As String
    Dim s As String
    s = LTrim$(Str$(value))
    If Left$(s, 1) = "." Then s = "0" & s
    If InStr(s, ".") = 0 And InStr(s, "E") = 0 Then s = s & ".0"
    fmtAlternateFloat = s
End Function

' makeGameTimeLimitOptionsDropdownPopulateFunc's addGameTimeLimitMinutesOption: [&](uint32_t gameTimeLimitMinutes)
Private Sub addGameTimeLimitMinutesOption(ByVal result As OptionChoices, ByVal gameTimeLimitMinutes As Long)
    Dim buttonText As String, hoursFloatStr As String
    If gameTimeLimitMinutes >= 120 Then
        hoursFloatStr = fmtAlternateFloat(CSng(CSng(gameTimeLimitMinutes) / 60!))
        buttonText = va("%s hours", hoursFloatStr)
    Else
        buttonText = va("%u minutes", gameTimeLimitMinutes)
    End If
    result.addChoice buttonText, "", gameTimeLimitMinutes
End Sub

Public Function makeDefaultsOptionsForm() As WIDGET
    Dim result As WIDGET, optionInfo As OptionInfo, valueChanger As WIDGET, validSPColors As Variant, i As Long, tc As Long
    Set result = OptionsForm_make()

    ' Campaign:
    OptionsForm_addSection result, OptionsSection_("Campaign", ""), True
    Set optionInfo = OptionInfo_new("defaults.campaign.unitColor", "Unit Color", "")
    ' Valid colors for SP are 0, 4-7. (Some are reserved for AI players.)
    validSPColors = Array(0&, 4&, 5&, 6&, 7&)
    Set valueChanger = PlayerColorOptionsSelector_make(WzFunc_new(FN_defaults_spColor_get, validSPColors), WzFunc_new(FN_defaults_spColor_select))
    For i = 0 To UBound(validSPColors)
        PlayerColorOptionsSelector_addColor valueChanger, validSPColors(i)
    Next
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set optionInfo = OptionInfo_new("defaults.campaign.difficultyLevel", "Difficulty Level", "")
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_defaults_difficulty_populate), WzFunc_new(FN_defaults_difficulty_select), True)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    ' Multiplayer:
    OptionsForm_addSection result, OptionsSection_("Multiplayer", ""), True
    Set optionInfo = OptionInfo_new("defaults.multiplayer.unitColor", "Unit Color", "")
    Set valueChanger = PlayerColorOptionsSelector_make(WzFunc_new(FN_defaults_mpColor_get), WzFunc_new(FN_defaults_mpColor_select))
    PlayerColorOptionsSelector_addColor valueChanger, Empty   ' nullopt
    For tc = 0 To MAX_PLAYERS_IN_GUI - 1
        PlayerColorOptionsSelector_addColor valueChanger, tc
    Next
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set optionInfo = OptionInfo_new("defaults.multiplayer.defaultAI", "Default AI", "")
    optionInfo.setRequiresRestart True
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_defaults_defaultAI_populate), WzFunc_new(FN_defaults_defaultAI_select), True)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    ' Lobby Options:
    OptionsForm_addSection result, OptionsSection_("Lobby Options", ""), True
    Set optionInfo = OptionInfo_new("defaults.lobby.enableipv6", "Enable IPv6", "Whether to list and try connecting via IPv6 (if available).")
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_defaults_enableipv6_populate), WzFunc_new(FN_defaults_enableipv6_select), False)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set optionInfo = OptionInfo_new("defaults.lobby.preferipv6", "Prefer IPv6 over IPv4", "Prefer connecting via IPv6 over IPv4 when a game is available via both.")
    optionInfo.addAvailabilityCondition WzFunc_new(FN_IsLobbyIPv6Enabled)
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_defaults_preferipv6_populate), WzFunc_new(FN_defaults_preferipv6_select), True)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    ' Hosting Options:
    OptionsForm_addSection result, OptionsSection_("Hosting Options", ""), True
    Set optionInfo = OptionInfo_new("defaults.hosting.portMapping", "Port Mapping", "Use PCP, NAT-PMP, or UPnP to help configure your router / firewall to allow connections while hosting.")
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_defaults_portMapping_populate), WzFunc_new(FN_defaults_portMapping_select), False)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set optionInfo = OptionInfo_new("defaults.hosting.chat", "Chat", "")
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_defaults_chat_populate), WzFunc_new(FN_defaults_chat_select), False)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set optionInfo = OptionInfo_new("defaults.hosting.inactivityTimeout", "Inactivity Timeout", "")
    optionInfo.addAvailabilityCondition WzFunc_new(FN_IsNotInGame)
    Set valueChanger = OptionsDropdown_make(makeInactivityTimeoutOptionsDropdownPopulateFunc(), WzFunc_new(FN_defaults_inactivityTimeout_select), False)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set optionInfo = OptionInfo_new("defaults.hosting.lagKick", "Lag Kick", "")
    Set valueChanger = OptionsDropdown_make(makeLagKickOptionsDropdownPopulateFunc(), WzFunc_new(FN_defaults_lagKick_select), False)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set optionInfo = OptionInfo_new("defaults.hosting.spectatorSlots", "Spectator Slots", "")
    optionInfo.addAvailabilityCondition WzFunc_new(FN_IsNotInGame)
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_defaults_spectatorSlots_populate), WzFunc_new(FN_defaults_spectatorSlots_select), False)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set optionInfo = OptionInfo_new("defaults.hosting.onPlayerLeave", "On Player Leave", "")
    optionInfo.addAvailabilityCondition WzFunc_new(FN_IsNotInGame)
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_defaults_onPlayerLeave_populate), WzFunc_new(FN_defaults_onPlayerLeave_select), False)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set optionInfo = OptionInfo_new("defaults.hosting.gameTimeLimit", "Game Time Limit", "")
    optionInfo.addAvailabilityCondition WzFunc_new(FN_IsNotInGame)
    Set valueChanger = OptionsDropdown_make(makeGameTimeLimitOptionsDropdownPopulateFunc(), WzFunc_new(FN_defaults_gameTimeLimit_select), False)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set makeDefaultsOptionsForm = result
End Function

' MARK: - the lambdas

Public Function defaultsoptions_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim result As OptionChoices, strongSelf As WIDGET, currValue As Long, inactivityMinutes As Long, lagKickSeconds As Long
    Dim gameTimeLimitMinutes As Long, validSPColors As Variant, isFound As Boolean, i As Long, level As Long
    Dim openSpectatorSlots As Long, levelV As Variant
    Select Case f.id
    Case FN_IsLobbyIPv6Enabled
        Set defaultsoptions_c_fn = IsLobbyIPv6Enabled(args(0))

    Case FN_PCOS_colorClick                     ' [weakSelf, tc](W_BUTTON&)
        Set strongSelf = WIDGET_fromPtr(f.cap(0))
        If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/titleui/options/defaultsoptions.cpp:104", "addColor", "Parent gone?": Exit Function
        PlayerColorOptionsSelector_handleColorPressed strongSelf, f.cap(1)

    Case FN_inactivityTimeout_populate          ' OptionChoices<uint32_t> ()
        Set result = New OptionChoices
        result.addChoice "Off", "", 0
        For inactivityMinutes = MIN_MPINACTIVITY_MINUTES To (MIN_MPINACTIVITY_MINUTES + 6)
            result.addChoice va("%u minutes", inactivityMinutes), "", inactivityMinutes
        Next
        currValue = war_getMPInactivityMinutes()
        If Not result.setCurrentIdxForValue(currValue) Then
            ' add "Custom" item
            result.addChoice va("%u minutes", currValue), "", currValue
            result.currentIdx = result.count - 1
        End If
        Set defaultsoptions_c_fn = result

    Case FN_lagKick_populate                    ' OptionChoices<int> ()
        Set result = New OptionChoices
        result.addChoice "Off", "", 0
        For lagKickSeconds = 60 To 120 Step 30
            result.addChoice va("%u seconds", lagKickSeconds), "", lagKickSeconds
        Next
        currValue = war_getAutoLagKickSeconds()
        If Not result.setCurrentIdxForValue(currValue) Then
            ' add "Custom" item
            result.addChoice va("%u seconds", currValue), "", currValue
            result.currentIdx = result.count - 1
        End If
        Set defaultsoptions_c_fn = result

    Case FN_gameTimeLimit_populate              ' OptionChoices<uint32_t> ()
        Set result = New OptionChoices
        result.addChoice "Off", "", 0
        For gameTimeLimitMinutes = MIN_MPGAMETIMELIMIT_MINUTES To (MIN_MPGAMETIMELIMIT_MINUTES + (15 * 10)) Step 15
            addGameTimeLimitMinutesOption result, gameTimeLimitMinutes
        Next
        currValue = war_getMPGameTimeLimitMinutes()
        If Not result.setCurrentIdxForValue(currValue) Then
            ' add custom value
            addGameTimeLimitMinutesOption result, currValue
            result.currentIdx = result.count - 1
        End If
        Set defaultsoptions_c_fn = result

    Case FN_defaults_spColor_get                ' [validSPColors]() -> optional<unsigned>
        validSPColors = f.cap(0)
        currValue = war_GetSPcolor()
        isFound = False
        For i = 0 To UBound(validSPColors)
            If validSPColors(i) = currValue Then isFound = True: Exit For
        Next
        If Not isFound Then
            currValue = 0
        End If
        defaultsoptions_c_fn = currValue

    Case FN_defaults_spColor_select             ' [](optional<unsigned> tc) -> bool
        If IsEmpty(args(0)) Then
            defaultsoptions_c_fn = False
            Exit Function
        End If
        war_SetSPcolor CLng(args(0))
        defaultsoptions_c_fn = True

    Case FN_defaults_difficulty_populate        ' OptionChoices<DIFFICULTY_LEVEL> ()
        Set result = New OptionChoices
        For Each levelV In Array(DL_SUPER_EASY, DL_EASY, DL_NORMAL, DL_HARD, DL_INSANE)   ' availableDifficultyLevels
            level = levelV
            result.addChoice difficultyLevelToString(level), getCampaignDifficultyDescriptionString(level), level
        Next
        result.setCurrentIdxForValue CLng(getDifficultyLevel())
        Set defaultsoptions_c_fn = result

    Case FN_defaults_difficulty_select
        setDifficultyLevel CLng(args(0))
        defaultsoptions_c_fn = True

    Case FN_defaults_mpColor_get                ' []() -> optional<unsigned>
        currValue = war_getMPcolour()
        If currValue < 0 Then
            defaultsoptions_c_fn = Empty    ' nullopt
            Exit Function
        End If
        defaultsoptions_c_fn = currValue

    Case FN_defaults_mpColor_select             ' [](optional<unsigned> tc) -> bool
        If IsEmpty(args(0)) Then war_setMPcolour -1 Else war_setMPcolour CLng(args(0))
        defaultsoptions_c_fn = True

    Case FN_defaults_defaultAI_populate         ' OptionChoices<WzString> ()
        Set result = New OptionChoices
        For i = 0 To getAIData_size() - 1
            result.addChoice getAIData_name(i), "", getAIData_js(i)
        Next
        result.setCurrentIdxForValue getDefaultSkirmishAI()
        Set defaultsoptions_c_fn = result

    Case FN_defaults_defaultAI_select
        setDefaultSkirmishAI CStr(args(0))
        defaultsoptions_c_fn = True

    Case FN_defaults_enableipv6_populate        ' OptionChoices<bool> ()
        Set result = New OptionChoices
        result.addChoice "Off", "Disable connecting via IPv6, and listing IPv6-only games. Only recommended if your system has issues with IPv6 connectivity.", False
        result.addChoice "On", "Enable IPv6. If your system supports it, this is the recommended option.", True
        result.setCurrentIdxForValue Not war_getLobbyDisableIPv6()
        Set defaultsoptions_c_fn = result
    Case FN_defaults_enableipv6_select
        war_setLobbyDisableIPv6 Not CBool(args(0))
        defaultsoptions_c_fn = True

    Case FN_defaults_preferipv6_populate        ' OptionChoices<bool> ()
        Set result = New OptionChoices
        result.addChoice "Off", "", False
        result.addChoice "On", "", True
        result.setCurrentIdxForValue NETgetJoinPreferenceIPv6()
        Set defaultsoptions_c_fn = result
    Case FN_defaults_preferipv6_select
        NETsetJoinPreferenceIPv6 CBool(args(0))
        defaultsoptions_c_fn = True

    Case FN_defaults_portMapping_populate       ' OptionChoices<bool> ()
        Set result = New OptionChoices
        result.addChoice "Off", "", False
        result.addChoice "On", "", True
        result.setCurrentIdxForValue NetPlay_isPortMappingEnabled
        Set defaultsoptions_c_fn = result
    Case FN_defaults_portMapping_select
        NetPlay_isPortMappingEnabled = CBool(args(0))
        defaultsoptions_c_fn = True

    Case FN_defaults_chat_populate              ' OptionChoices<bool> ()
        Set result = New OptionChoices
        result.addChoice "Quick Chat Only", "", False
        result.addChoice "Allow All", "", True
        result.setCurrentIdxForValue NETgetDefaultMPHostFreeChatPreference()
        Set defaultsoptions_c_fn = result
    Case FN_defaults_chat_select
        NETsetDefaultMPHostFreeChatPreference CBool(args(0))
        defaultsoptions_c_fn = True

    Case FN_defaults_inactivityTimeout_select   ' [](const auto& newValue) -> bool
        war_setMPInactivityMinutes CLng(args(0))
        game.inactivityMinutes = war_getMPInactivityMinutes()
        defaultsoptions_c_fn = True

    Case FN_defaults_lagKick_select
        war_setAutoLagKickSeconds CLng(args(0))
        defaultsoptions_c_fn = True

    Case FN_defaults_spectatorSlots_populate    ' OptionChoices<uint32_t> ()
        Set result = New OptionChoices
        result.addChoice "None", "", 0
        For openSpectatorSlots = 1 To MAX_SPECTATOR_SLOTS
            result.addChoice CStr(openSpectatorSlots), "", openSpectatorSlots
        Next
        result.setCurrentIdxForValue war_getMPopenSpectatorSlots()
        Set defaultsoptions_c_fn = result
    Case FN_defaults_spectatorSlots_select
        war_setMPopenSpectatorSlots CLng(args(0))
        defaultsoptions_c_fn = True

    Case FN_defaults_onPlayerLeave_populate     ' OptionChoices<PLAYER_LEAVE_MODE> ()
        Set result = New OptionChoices
        result.addChoice "Distribute to Team", "", PLAYER_LEAVE_MODE_SPLIT_WITH_TEAM
        result.addChoice "Destroy (Classic)", "", PLAYER_LEAVE_MODE_DESTROY_RESOURCES
        result.setCurrentIdxForValue war_getMPPlayerLeaveMode()
        Set defaultsoptions_c_fn = result
    Case FN_defaults_onPlayerLeave_select
        war_setMPPlayerLeaveMode CLng(args(0))
        defaultsoptions_c_fn = True

    Case FN_defaults_gameTimeLimit_select
        war_setMPGameTimeLimitMinutes CLng(args(0))
        game.gameTimeLimitMinutes = war_getMPGameTimeLimitMinutes()
        defaultsoptions_c_fn = True
    End Select
End Function
