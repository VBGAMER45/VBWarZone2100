Attribute VB_Name = "starthostingbutton_c"
' src/titleui/widgets/starthostingbutton.cpp, starthostingbutton.h: the lobby's
' "Start Hosting" button (WzStartHostingButton) with its advanced hosting
' options pop-over (HostingAdvancedOptionsForm). The widget object model is
' docs/CONVENTIONS.md's "Widgets": the classes' fields are the data classes
' WzStartHostingButton.cls and HostingAdvancedOptionsForm.cls; the lambdas are
' WzFuncs run by starthostingbutton_c_fn. A std::weak_ptr to a widget or to the
' title UI is its address (ObjPtr), read back with WIDGET_fromPtr /
' titleUIFromPtr.
' (VB) widget class HostingAdvancedOptionsForm : WIDGET
' (VB) widget class WzStartHostingButton : W_BUTTON
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

' HostingAdvancedOptionsForm's const int members
Private Const innerVerticalSpacing As Long = 15
Private Const outerPaddingX As Long = 20
Private Const optionsListPaddingX As Long = 10
Private Const outerPaddingY As Long = 20
Private Const betweenButtonPadding As Long = 15
Private Const closeButtonVerticalPadding As Long = 5

' the lambdas
Public Const FN_HAOF_CLOSE_CLICK As Long = 6000         ' closeButton->addOnClickHandler([onClose](W_BUTTON&) {...})
Public Const FN_HAOF_CLOSE_TASK As Long = 6001          ' widgScheduleTask([onClose]() { onClose(); })
Public Const FN_HAOF_HOSTROLE_POPULATE As Long = 6002   ' the Host Role dropdown's populate function
Public Const FN_HAOF_HOSTROLE_SELECT As Long = 6003     ' its select function
Public Const FN_HAOF_BLINDMODE_AVAILABLE As Long = 6004 ' the Blind Mode option's availability condition
Public Const FN_HAOF_BLINDMODE_POPULATE As Long = 6005
Public Const FN_HAOF_BLINDMODE_SELECT As Long = 6006
Public Const FN_SHB_OPTIONS_CLICK As Long = 6007        ' optionsButton->addOnClickHandler([weakSelf](W_BUTTON&) {...})
Public Const FN_SHB_OPTIONS_TASK As Long = 6008         ' widgScheduleTask([weakSelf]() {...})
Public Const FN_SHB_POPOVER_CLOSE As Long = 6009        ' HostingAdvancedOptionsForm::make(..., [psWeakSelf]() {...})
Public Const FN_SHB_OVERLAY_CLICKED As Long = 6010      ' newRootFrm->onClickedFunc
Public Const FN_SHB_POPOVER_CALCLAYOUT As Long = 6011   ' optionsPopOver->setCalcLayout([](WIDGET *psWidget) {...})
Public Const FN_SHB_CLEAR_FOCUS As Long = 6012          ' widgScheduleTask([lockedScreen]() {...})

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

' MARK: - HostingAdvancedOptionsForm

Public Sub HostingAdvancedOptionsForm_construct(ByVal self As WIDGET)
    WIDGET_construct self, WIDG_UNSPECIFIED_TYPE
    self.kind = K_HostingAdvancedOptionsForm ' (VB) virtual calls from here on reach HostingAdvancedOptionsForm's overrides
End Sub

' static std::shared_ptr<HostingAdvancedOptionsForm> make(const std::shared_ptr<WzStartHostingButton>& parent, const std::function<void ()>& onCloseFunc)
Public Function HostingAdvancedOptionsForm_make(ByVal parent As WIDGET, ByVal onCloseFunc As WzFunc) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_HostingAdvancedOptionsForm)
    HostingAdvancedOptionsForm_construct widget
    HostingAdvancedOptionsForm_initialize widget, parent, onCloseFunc
    Set HostingAdvancedOptionsForm_make = widget
End Function

Private Sub HostingAdvancedOptionsForm_initialize(ByVal self As WIDGET, ByVal parent As WIDGET, ByVal onClose As WzFunc)
    Dim d As HostingAdvancedOptionsForm
    Set d = HostingAdvancedOptionsForm_(self)
    d.multiOptionsTitleUI = ObjPtr(WzStartHostingButton_getMultiOptionsTitleUI(parent))

    Set d.formTitle = W_LABEL_new()
    W_LABEL_setFont d.formTitle, font_regular_bold, psPalette(WZCOL_TEXT_MEDIUM)
    d.formTitle.setString "Advanced Hosting Options:"
    W_LABEL_setCanTruncate d.formTitle, True
    d.formTitle.setTransparentToMouse True
    d.formTitle.setGeometry outerPaddingX, outerPaddingY, W_LABEL_getMaxLineWidth(d.formTitle), iV_GetTextLineSize(font_regular_bold)
    self.attach d.formTitle

    ' Add "Close" button
    Set d.closeButton = WzFrontendImageButton_make(False, 0)
    d.closeButton.setString "Close"
    W_BUTTON_addOnClickHandler d.closeButton, WzFunc_new(FN_HAOF_CLOSE_CLICK, onClose)
    WzFrontendImageButton_setPadding d.closeButton, 7, closeButtonVerticalPadding
    self.attach d.closeButton

    Set d.hostOptionsForm = HostingAdvancedOptionsForm_makeHostOptionsForm(self)
    self.attach d.hostOptionsForm
End Sub

Public Function HostingAdvancedOptionsForm_idealWidth(ByVal self As WIDGET) As Long
    HostingAdvancedOptionsForm_idealWidth = (outerPaddingX * 2) + optionsListPaddingX + HostingAdvancedOptionsForm_(self).hostOptionsForm.idealWidth()
End Function

Public Function HostingAdvancedOptionsForm_idealHeight(ByVal self As WIDGET) As Long
    Dim d As HostingAdvancedOptionsForm, result As Long, a As Long, b As Long
    Set d = HostingAdvancedOptionsForm_(self)
    a = d.formTitle.idealHeight()
    b = d.closeButton.idealHeight() - closeButtonVerticalPadding
    If b > a Then a = b
    result = (outerPaddingY * 2) + a
    result = result + innerVerticalSpacing
    result = result + d.hostOptionsForm.idealHeight()
    HostingAdvancedOptionsForm_idealHeight = result
End Function

Private Function HostingAdvancedOptionsForm_makeHostOptionsForm(ByVal self As WIDGET) As WIDGET
    Dim psWeakParent As Long, weakMultiOptionsTitleUI As Long, result As WIDGET
    Dim optionInfo As OptionInfo, valueChanger As WIDGET
    psWeakParent = ObjPtr(self)
    weakMultiOptionsTitleUI = HostingAdvancedOptionsForm_(self).multiOptionsTitleUI
    Set result = OptionsForm_make()
    Set optionInfo = OptionInfo_new("mp.hostRole", "Host Role", "")
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_HAOF_HOSTROLE_POPULATE, weakMultiOptionsTitleUI), WzFunc_new(FN_HAOF_HOSTROLE_SELECT, psWeakParent), False)
    OptionsForm_addOption result, optionInfo, valueChanger

    Set optionInfo = OptionInfo_new("mp.blindMode", "Blind Mode", "Whether and when players are blinded to other players' identities.")
    optionInfo.addAvailabilityCondition WzFunc_new(FN_HAOF_BLINDMODE_AVAILABLE, weakMultiOptionsTitleUI)
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_HAOF_BLINDMODE_POPULATE, weakMultiOptionsTitleUI), WzFunc_new(FN_HAOF_BLINDMODE_SELECT, psWeakParent), True)
    OptionsForm_addOption result, optionInfo, valueChanger
    Set HostingAdvancedOptionsForm_makeHostOptionsForm = result
End Function

Private Function HostingAdvancedOptionsForm_setSpectatorHost(ByVal self As WIDGET, ByVal value As Boolean) As Boolean
    Dim strongMultiOptionsTitleUI As WzMultiplayerOptionsTitleUI
    Set strongMultiOptionsTitleUI = titleUIFromPtr(HostingAdvancedOptionsForm_(self).multiOptionsTitleUI)
    If Not (Not strongMultiOptionsTitleUI Is Nothing) Then ASSERT_FAILURE "strongMultiOptionsTitleUI != nullptr", "src/titleui/widgets/starthostingbutton.cpp:198", "setSpectatorHost", "No parent multi options title ui?": Exit Function
    strongMultiOptionsTitleUI.setOption_SpectatorHost value
    HostingAdvancedOptionsForm_setSpectatorHost = True
End Function

Private Function HostingAdvancedOptionsForm_setBlindMode(ByVal self As WIDGET, ByVal value As Long) As Boolean
    Dim strongMultiOptionsTitleUI As WzMultiplayerOptionsTitleUI
    Set strongMultiOptionsTitleUI = titleUIFromPtr(HostingAdvancedOptionsForm_(self).multiOptionsTitleUI)
    If Not (Not strongMultiOptionsTitleUI Is Nothing) Then ASSERT_FAILURE "strongMultiOptionsTitleUI != nullptr", "src/titleui/widgets/starthostingbutton.cpp:206", "setBlindMode", "No parent multi options title ui?": Exit Function
    If value <> BLIND_MODE_NONE Then
        If Not strongMultiOptionsTitleUI.getOption_SpectatorHost() Then ASSERT_FAILURE "value == BLIND_MODE::NONE || strongMultiOptionsTitleUI->getOption_SpectatorHost()", "src/titleui/widgets/starthostingbutton.cpp:207", "setBlindMode", "Blind modes are only available when host is a spectator": Exit Function
    End If
    game.blindMode = value
    HostingAdvancedOptionsForm_setBlindMode = True
End Function

Public Sub HostingAdvancedOptionsForm_geometryChanged(ByVal self As WIDGET)
    Dim d As HostingAdvancedOptionsForm, lastButtonX0 As Long, buttonY0 As Long, topButtonHeight As Long, buttonX0 As Long
    Dim maxAvailableOptionsFormHeight As Long, hostOptionsFormY0 As Long, a As Long, b As Long, h As Long
    Set d = HostingAdvancedOptionsForm_(self)
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
    maxAvailableOptionsFormHeight = self.height() - (outerPaddingY * 2) - a - innerVerticalSpacing
    a = d.formTitle.idealHeight()
    b = d.closeButton.idealHeight()
    If b > a Then a = b
    hostOptionsFormY0 = outerPaddingY + a + innerVerticalSpacing
    h = d.hostOptionsForm.idealHeight()
    If maxAvailableOptionsFormHeight < h Then h = maxAvailableOptionsFormHeight
    d.hostOptionsForm.setGeometry outerPaddingX + optionsListPaddingX, hostOptionsFormY0, self.width() - (outerPaddingX * 2) - optionsListPaddingX, h

End Sub

Public Sub HostingAdvancedOptionsForm_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
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


' MARK: - WzStartHostingButton

' static std::shared_ptr<WzStartHostingButton> make(const std::shared_ptr<WzMultiplayerOptionsTitleUI>& parentTitleUI)
Public Function WzStartHostingButton_make(ByVal parentTitleUI As WzMultiplayerOptionsTitleUI) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_WzStartHostingButton)
    WzStartHostingButton_construct widget
    WzStartHostingButton_(widget).multiOptionsTitleUI = ObjPtr(parentTitleUI)
    WzStartHostingButton_initialize widget
    Set WzStartHostingButton_make = widget
End Function

Public Sub WzStartHostingButton_construct(ByVal self As WIDGET)
    W_BUTTON_construct self
    self.kind = K_WzStartHostingButton ' (VB) virtual calls from here on reach WzStartHostingButton's overrides
    Set WzStartHostingButton_(self).wzText_ = New WzText
End Sub

Public Sub WzStartHostingButton_destruct(ByVal self As WIDGET)
    WzStartHostingButton_closeOptionsOverlay self
End Sub

Private Sub WzStartHostingButton_initialize(ByVal self As WIDGET)
    Dim minButtonWidthForText As Long, showOptionsButton As Boolean, weakSelf As Long, but As WIDGET
    self.setString "Start Hosting"
    W_BUTTON_(self).FontID = font_medium
    minButtonWidthForText = iV_GetTextWidth(W_BUTTON_(self).pText, W_BUTTON_(self).FontID)
    self.setGeometry 0, 0, minButtonWidthForText + 10, iV_GetTextLineSize(W_BUTTON_(self).FontID) + 16

    showOptionsButton = (NetPlay_bComms And ingame.side = InGameSide_HOST_OR_SINGLEPLAYER)
    If showOptionsButton Then
        weakSelf = ObjPtr(self)
        ' Initialize an embedded options button to the right
        Set but = WzFrontendImageButton_make(True, IMAGE_GEAR)
        WzFrontendImageButton_setBorderDrawMode but, BorderDrawMode_Highlighted
        WzFrontendImageButton_setPadding but, 0, 0
        WzFrontendImageButton_setImageDimensions but, 16
        but.setTip "Configure Advanced Hosting Options"
        Set WzStartHostingButton_(self).optionsButton = but
        self.attach WzStartHostingButton_(self).optionsButton
        W_BUTTON_addOnClickHandler WzStartHostingButton_(self).optionsButton, WzFunc_new(FN_SHB_OPTIONS_CLICK, weakSelf)
    End If
End Sub

Public Function WzStartHostingButton_getMultiOptionsTitleUI(ByVal self As WIDGET) As WzMultiplayerOptionsTitleUI
    Set WzStartHostingButton_getMultiOptionsTitleUI = titleUIFromPtr(WzStartHostingButton_(self).multiOptionsTitleUI)
End Function

Public Sub WzStartHostingButton_geometryChanged(ByVal self As WIDGET)
    Dim butWidth As Long, x0 As Long
    If Not WzStartHostingButton_(self).optionsButton Is Nothing Then
        butWidth = self.height()
        If butWidth < 16 Then butWidth = 16
        x0 = self.width() - butWidth
        WzStartHostingButton_(self).optionsButton.setGeometry x0, 0, butWidth, self.height()
    End If
End Sub

Public Sub WzStartHostingButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x0 As Long, y0 As Long, w As Long, h As Long, haveText As Boolean
    Dim isDown As Boolean, isDisabled As Boolean, isHighlight As Boolean
    Dim border_color As Long, background_color As Long, backX0 As Long, backY0 As Long
    Dim txt As WzText, fw As Long, fx As Long, fy As Long, textColor As Long
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    w = self.width()
    h = self.height()

    haveText = Len(W_BUTTON_(self).pText) <> 0

    isDown = (self.getState() And (WBUT_DOWN Or WBUT_LOCK Or WBUT_CLICKLOCK)) <> 0
    isDisabled = (self.getState() And WBUT_DISABLE) <> 0
    isHighlight = Not isDisabled And ((self.getState() And WBUT_HIGHLIGHT) <> 0)

    ' Display the button.
    If isHighlight Then border_color = psPalette(WZCOL_TEXT_BRIGHT) Else border_color = psPalette(WZCOL_TEXT_MEDIUM)
    If isDown Or isDisabled Then background_color = pal_RGBA(10, 0, 70, 250) Else background_color = psPalette(WZCOL_MENU_BACKGROUND)
    If isDisabled Then
        border_color = psPalette(WZCOL_TEXT_DARK)
        background_color = pielightWithA(background_color, pielightA(background_color) \ 2)
    ElseIf isDown Then
        background_color = psPalette(WZCOL_FORM_DARK)
    End If

    backX0 = x0
    backY0 = y0
    pie_BoxFill backX0, backY0, backX0 + w, backY0 + h, border_color
    pie_BoxFill backX0 + 1, backY0 + 1, backX0 + w - 1, backY0 + h - 1, background_color

    If haveText Then
        Set txt = WzStartHostingButton_(self).wzText_
        txt.setText W_BUTTON_(self).pText, W_BUTTON_(self).FontID
        fw = txt.width()
        fx = x0 + (self.width() - fw) \ 2
        fy = y0 + (self.height() - txt.lineSize()) \ 2 - txt.aboveBase()
        textColor = psPalette(WZCOL_TEXT_BRIGHT)
        If isDisabled Then
            textColor = psPalette(WZCOL_TEXT_DARK)
            textColor = pielightWithA(textColor, pielightA(textColor) \ 2)
        ElseIf isHighlight Then
            textColor = psPalette(WZCOL_TEXT_BRIGHT)
        End If
        txt.render fx, fy, textColor
    End If
End Sub

Private Function WzStartHostingButton_createOptionsPopoverForm(ByVal self As WIDGET) As WIDGET
    Dim psWeakSelf As Long, optionsPopOver As WIDGET, maxPopoverHeight As Long, h As Long
    psWeakSelf = ObjPtr(self)
    Set optionsPopOver = HostingAdvancedOptionsForm_make(self, WzFunc_new(FN_SHB_POPOVER_CLOSE, psWeakSelf))
    maxPopoverHeight = 350
    optionsPopOver.setGeometry 0, 0, FRONTEND_BUTWIDTH - 100, maxPopoverHeight ' set once to set width
    h = optionsPopOver.idealHeight()
    If maxPopoverHeight < h Then h = maxPopoverHeight
    optionsPopOver.setGeometry 0, 0, optionsPopOver.width(), h ' set again to set height based on idealHeight for that width (if smaller)
    Set WzStartHostingButton_createOptionsPopoverForm = optionsPopOver
End Function

Private Sub WzStartHostingButton_closeOptionsOverlay(ByVal self As WIDGET)
    If Not WzStartHostingButton_(self).optionsOverlayScreen Is Nothing Then
        widgRemoveOverlayScreen WzStartHostingButton_(self).optionsOverlayScreen
        Set WzStartHostingButton_(self).optionsOverlayScreen = Nothing
    End If
End Sub

Private Sub WzStartHostingButton_displayOptionsOverlay(ByVal self As WIDGET, ByVal psParent As WIDGET)
    Dim lockedScreen As W_SCREEN, newRootFrm As WIDGET, psWeakOptionsOverlayScreen As Long, psWeakHostingButton As Long
    Dim optionsPopOver As WIDGET, d As WzStartHostingButton
    Set d = WzStartHostingButton_(self)
    Set lockedScreen = self.screenPointer
    If Not (Not lockedScreen Is Nothing) Then ASSERT_FAILURE "lockedScreen != nullptr", "src/titleui/widgets/starthostingbutton.cpp:392", "displayOptionsOverlay", "The WzPlayerBoxTabs does not have an associated screen pointer?"

    ' Initialize the options overlay screen
    Set d.optionsOverlayScreen = W_SCREEN_make()
    Set newRootFrm = W_FULLSCREENOVERLAY_CLICKFORM_make()
    psWeakOptionsOverlayScreen = ObjPtr(d.optionsOverlayScreen)
    psWeakHostingButton = ObjPtr(self)
    Set W_FULLSCREENOVERLAY_CLICKFORM_(newRootFrm).onClickedFunc = WzFunc_new(FN_SHB_OVERLAY_CLICKED, psWeakOptionsOverlayScreen, psWeakHostingButton)
    Set W_FULLSCREENOVERLAY_CLICKFORM_(newRootFrm).onCancelPressed = W_FULLSCREENOVERLAY_CLICKFORM_(newRootFrm).onClickedFunc
    d.optionsOverlayScreen.psForm.attach newRootFrm

    ' Create the pop-over form
    Set optionsPopOver = WzStartHostingButton_createOptionsPopoverForm(self)
    newRootFrm.attach optionsPopOver

    ' Position the pop-over form
    optionsPopOver.setCalcLayout WzFunc_new(FN_SHB_POPOVER_CALCLAYOUT)

    widgRegisterOverlayScreenOnTopOfScreen d.optionsOverlayScreen, lockedScreen
    widgScheduleTask WzFunc_new(FN_SHB_CLEAR_FOCUS, lockedScreen)
End Sub

' MARK: - the lambdas

Public Function starthostingbutton_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim strongMultiOptionsTitleUI As WzMultiplayerOptionsTitleUI, strongParent As WIDGET, strongParentForm As WIDGET
    Dim result As OptionChoices, current_SpectatorHostValue As Boolean, strongSelf As WIDGET, psWidget As WIDGET
    Dim psOverlayScreen As W_SCREEN, strongHostingButton As WIDGET, popOverX0 As Long, popOverY0 As Long, lockedScreen As W_SCREEN
    Select Case f.id
    Case FN_HAOF_CLOSE_CLICK                    ' [onClose](W_BUTTON&)
        widgScheduleTask WzFunc_new(FN_HAOF_CLOSE_TASK, f.cap(0))

    Case FN_HAOF_CLOSE_TASK                     ' [onClose]()
        WzFunc_call f.cap(0)

    Case FN_HAOF_HOSTROLE_POPULATE              ' [weakMultiOptionsTitleUI]()
        Set result = New OptionChoices
        result.addChoice "Player Host", "Host is a player in the game (the default).", False
        result.addChoice "Spectator Host", "Host is a spectator in the game.", True
        current_SpectatorHostValue = False
        Set strongMultiOptionsTitleUI = titleUIFromPtr(f.cap(0))
        If Not (Not strongMultiOptionsTitleUI Is Nothing) Then ASSERT_FAILURE "strongMultiOptionsTitleUI != nullptr", "src/titleui/widgets/starthostingbutton.cpp:134", "makeHostOptionsForm", "No parent multi options title ui?"
        If Not strongMultiOptionsTitleUI Is Nothing Then
            current_SpectatorHostValue = strongMultiOptionsTitleUI.getOption_SpectatorHost()
        End If
        If current_SpectatorHostValue Then result.currentIdx = 1 Else result.currentIdx = 0
        Set starthostingbutton_c_fn = result

    Case FN_HAOF_HOSTROLE_SELECT                ' [psWeakParent](const auto& newValue) -> bool
        Set strongParent = WIDGET_fromPtr(f.cap(0))
        If Not (Not strongParent Is Nothing) Then ASSERT_FAILURE "strongParent != nullptr", "src/titleui/widgets/starthostingbutton.cpp:144", "makeHostOptionsForm", "No parent?": starthostingbutton_c_fn = False: Exit Function
        starthostingbutton_c_fn = HostingAdvancedOptionsForm_setSpectatorHost(strongParent, CBool(args(0)))

    Case FN_HAOF_BLINDMODE_AVAILABLE            ' [weakMultiOptionsTitleUI](const OptionInfo&) -> OptionInfo::AvailabilityResult
        Set strongMultiOptionsTitleUI = titleUIFromPtr(f.cap(0))
        If Not (Not strongMultiOptionsTitleUI Is Nothing) Then ASSERT_FAILURE "strongMultiOptionsTitleUI != nullptr", "src/titleui/widgets/starthostingbutton.cpp:154", "makeHostOptionsForm", "No parent multi options title ui?": Set starthostingbutton_c_fn = makeAvailabilityResult(False, ""): Exit Function
        Set starthostingbutton_c_fn = makeAvailabilityResult(strongMultiOptionsTitleUI.getOption_SpectatorHost(), "Blind modes are only available when host is a Spectator Host.")

    Case FN_HAOF_BLINDMODE_POPULATE             ' [weakMultiOptionsTitleUI]()
        Set strongMultiOptionsTitleUI = titleUIFromPtr(f.cap(0))
        If Not (Not strongMultiOptionsTitleUI Is Nothing) Then ASSERT_FAILURE "strongMultiOptionsTitleUI != nullptr", "src/titleui/widgets/starthostingbutton.cpp:160", "makeHostOptionsForm", "No parent multi options title ui?"
        current_SpectatorHostValue = strongMultiOptionsTitleUI.getOption_SpectatorHost()
        Set result = New OptionChoices
        result.addChoice "None", "Players' identities are visible, in lobby and in game (the default).", BLIND_MODE_NONE
        result.addChoice "Blind Lobby", "Players' identities are hidden from everyone (except the spectator host) until the game starts.", BLIND_MODE_BLIND_LOBBY, Not current_SpectatorHostValue
        result.addChoice "Blind Game", "Players' identities are hidden from everyone (except the spectator host) until the game ends.", BLIND_MODE_BLIND_GAME, Not current_SpectatorHostValue
        Select Case game.blindMode
        Case BLIND_MODE_NONE
            result.currentIdx = 0
        Case BLIND_MODE_BLIND_LOBBY, BLIND_MODE_BLIND_LOBBY_SIMPLE_LOBBY
            result.currentIdx = 1
        Case BLIND_MODE_BLIND_GAME, BLIND_MODE_BLIND_GAME_SIMPLE_LOBBY
            result.currentIdx = 2
        End Select
        Set starthostingbutton_c_fn = result

    Case FN_HAOF_BLINDMODE_SELECT               ' [psWeakParent](const auto& newValue) -> bool
        Set strongParent = WIDGET_fromPtr(f.cap(0))
        If Not (Not strongParent Is Nothing) Then ASSERT_FAILURE "strongParent != nullptr", "src/titleui/widgets/starthostingbutton.cpp:186", "makeHostOptionsForm", "No parent?": starthostingbutton_c_fn = False: Exit Function
        starthostingbutton_c_fn = HostingAdvancedOptionsForm_setBlindMode(strongParent, CLng(args(0)))

    Case FN_SHB_OPTIONS_CLICK                   ' [weakSelf](W_BUTTON&)
        widgScheduleTask WzFunc_new(FN_SHB_OPTIONS_TASK, f.cap(0))

    Case FN_SHB_OPTIONS_TASK                    ' [weakSelf]()
        Set strongParentForm = WIDGET_fromPtr(f.cap(0))
        If Not (Not strongParentForm Is Nothing) Then ASSERT_FAILURE "strongParentForm != nullptr", "src/titleui/widgets/starthostingbutton.cpp:292", "initialize", "No parent form": Exit Function
        WzStartHostingButton_displayOptionsOverlay strongParentForm, WzStartHostingButton_(strongParentForm).optionsButton

    Case FN_SHB_POPOVER_CLOSE                   ' [psWeakSelf]()
        Set strongSelf = WIDGET_fromPtr(f.cap(0))
        If Not strongSelf Is Nothing Then
            WzStartHostingButton_closeOptionsOverlay strongSelf
        End If

    Case FN_SHB_OVERLAY_CLICKED                 ' [psWeakOptionsOverlayScreen, psWeakHostingButton]()
        Set psOverlayScreen = W_SCREEN_fromPtr(f.cap(0))
        If Not psOverlayScreen Is Nothing Then
            widgRemoveOverlayScreen psOverlayScreen
        End If
        ' Destroy Options overlay / overlay screen
        Set strongHostingButton = WIDGET_fromPtr(f.cap(1))
        If Not strongHostingButton Is Nothing Then
            Set WzStartHostingButton_(strongHostingButton).optionsOverlayScreen = Nothing
        End If

    Case FN_SHB_POPOVER_CALCLAYOUT              ' [](WIDGET *psWidget)
        Set psWidget = args(0)
        ' Centered in screen
        popOverX0 = (screenWidth - psWidget.width()) \ 2
        If popOverX0 < 0 Then
            popOverX0 = 0
        End If
        popOverY0 = (screenHeight - psWidget.height()) \ 2
        If popOverY0 < 0 Then
            popOverY0 = 0
        End If
        psWidget.move popOverX0, popOverY0

    Case FN_SHB_CLEAR_FOCUS                     ' [lockedScreen]()
        ' clear focus from the underlying chatbox, which aggressively takes focus
        Set lockedScreen = f.cap(0)
        lockedScreen.setFocus Nothing
    End Select
End Function
