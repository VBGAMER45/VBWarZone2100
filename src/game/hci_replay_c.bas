Attribute VB_Name = "hci_replay_c"
' src/hci.cpp's replay controller: class ReplayControllerWidget (with its button class W_REPLAY_CONTROL_BUTTON) and
' createReplayControllerOverlay, the pause / resume / fast-forward buttons over the game while a replay plays.
' replayOverlayScreen is here (hci_c's intShutDown and its screen size changes use it). The rest of hci.cpp is
' hci_c's. WZ_Notification_Image(path).loadImageToTexture() is gfx_api_context_loadTextureFromFile. Translations
' (_()) are not ported: the English text is used.
' (VB) widget class ReplayControllerWidget : W_FORM
' (VB) widget class W_REPLAY_CONTROL_BUTTON : W_BUTTON
Option Explicit

Private Const REPLAY_ACTION_BUTTONS_PADDING As Long = 5
Private Const REPLAY_ACTION_BUTTONS_SPACING As Long = 5
Private Const REPLAY_ACTION_BUTTONS_IMAGE_SIZE As Long = 16
Private Const REPLAY_ACTION_BUTTONS_WIDTH As Long = (REPLAY_ACTION_BUTTONS_IMAGE_SIZE + (REPLAY_ACTION_BUTTONS_PADDING * 2))
Private Const REPLAY_ACTION_BUTTONS_HEIGHT As Long = REPLAY_ACTION_BUTTONS_WIDTH

Private Const WZ_MAX_REPLAY_FASTFORWARD_TICKS As Long = 3   ' constexpr size_t

' W_REPLAY_CONTROL_BUTTON's enum class OutlineState
Private Const OutlineState_None As Long = 0
Private Const OutlineState_Single As Long = 1
Private Const OutlineState_Double As Long = 2

' ReplayControllerWidget's enum class ReplayGameStatus
Private Const ReplayGameStatus_PLAY As Long = 0
Private Const ReplayGameStatus_PAUSE As Long = 1

' the lambdas (WzFunc: hci_replay_c_fn)
Public Const FN_replayTitleLabel_calcLayout As Long = 18830
Public Const FN_replayPause_onClick As Long = 18831
Public Const FN_replayPause_calcLayout As Long = 18832
Public Const FN_replayPlay_onClick As Long = 18833
Public Const FN_replayPlay_calcLayout As Long = 18834
Public Const FN_replayFastForward_onClick As Long = 18835
Public Const FN_replayFastForward_calcLayout As Long = 18836
Public Const FN_replayRefreshButtonStates_task As Long = 18837 ' widgScheduleTask([sharedThis](){ sharedThis->refreshButtonStates(); })
Public Const FN_replayController_calcLayout As Long = 18838    ' createReplayControllerOverlay's setCalcLayout

Public replayOverlayScreen As W_SCREEN          ' static std::shared_ptr<W_SCREEN> replayOverlayScreen

' MARK: - W_REPLAY_CONTROL_BUTTON

' W_REPLAY_CONTROL_BUTTON(): W_BUTTON() (~W_REPLAY_CONTROL_BUTTON deletes pButtonImage: released with the data object)
Public Sub W_REPLAY_CONTROL_BUTTON_construct(ByVal self As WIDGET)
    W_BUTTON_construct self
    self.kind = K_W_REPLAY_CONTROL_BUTTON ' (VB) virtual calls from here on reach W_REPLAY_CONTROL_BUTTON's overrides
End Sub

Public Sub W_REPLAY_CONTROL_BUTTON_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As W_REPLAY_CONTROL_BUTTON, x0 As Long, y0 As Long, x1 As Long, y1 As Long, state As Long
    Dim isDown As Boolean, isDisabled As Boolean, isHighlight As Boolean, light_border As Long, fill_color As Long
    Dim firstColour As Long, secondColour As Long, imageLeft As Long, imageTop As Long
    Set d = W_REPLAY_CONTROL_BUTTON_(self)

    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    x1 = x0 + self.width()
    y1 = y0 + self.height()

    state = self.getState()
    isDown = (state And (WBUT_DOWN Or WBUT_LOCK Or WBUT_CLICKLOCK)) <> 0
    isDisabled = (state And WBUT_DISABLE) <> 0
    isHighlight = Not isDisabled And ((state And WBUT_HIGHLIGHT) <> 0)

    ' Display the button.
    light_border = pal_RGBA(255, 255, 255, 100)
    If isDown Or isDisabled Then fill_color = pal_RGBA(10, 0, 70, 200) Else fill_color = pal_RGBA(25, 0, 110, 175)
    If isDown Then firstColour = pal_RGBA(0, 0, 0, 0) Else firstColour = light_border
    If isDisabled Then secondColour = light_border Else secondColour = psPalette(WZCOL_FORM_DARK)
    iV_ShadowBox x0, y0, x1, y1, 0, firstColour, secondColour, fill_color

    ' Highlight
    If isHighlight Then
        iV_Box x0 + 2, y0 + 2, x1 - 2, y1 - 2, psPalette(WZCOL_FORM_HILITE)
    End If

    ' Outline
    Select Case d.outline
    Case OutlineState_None
    Case OutlineState_Double
        iV_Box x0 + 3, y0 + 3, x1 - 3, y1 - 3, pal_RGBA(218, 207, 255, 200)
        ' fall-through
        iV_Box x0 + 1, y0 + 1, x1 - 1, y1 - 1, pal_RGBA(218, 207, 255, 200)
    Case OutlineState_Single
        iV_Box x0 + 1, y0 + 1, x1 - 1, y1 - 1, pal_RGBA(218, 207, 255, 200)
    End Select

    ' Display the image, if present
    imageLeft = x1 - REPLAY_ACTION_BUTTONS_PADDING - REPLAY_ACTION_BUTTONS_IMAGE_SIZE
    imageTop = (self.y() + yOffset) + REPLAY_ACTION_BUTTONS_PADDING

    If Not d.pButtonImage Is Nothing Then
        iV_DrawImageAnisotropic d.pButtonImage, imageLeft, imageTop, 0!, 0!, REPLAY_ACTION_BUTTONS_IMAGE_SIZE, REPLAY_ACTION_BUTTONS_IMAGE_SIZE, 0!, psPalette(WZCOL_WHITE)
    End If
End Sub

' MARK: - ReplayControllerWidget

' ReplayControllerWidget(): W_FORM()
Public Sub ReplayControllerWidget_construct(ByVal self As WIDGET)
    W_FORM_construct self
    self.kind = K_ReplayControllerWidget ' (VB) virtual calls from here on reach ReplayControllerWidget's overrides
    ReplayControllerWidget_(self).autoStateRefreshInterval = GAME_TICKS_PER_SEC * 3
End Sub

Private Function ReplayControllerWidget_getReplayGameStatus() As Long
    Dim status As Long
    If Not gameTimeIsStopped() Then status = ReplayGameStatus_PLAY Else status = ReplayGameStatus_PAUSE
    ' sanity check
    ReplayControllerWidget_getReplayGameStatus = status
End Function

Private Sub ReplayControllerWidget_setReplayGameStatus(ByVal self As WIDGET, ByVal newStatus As Long)
    Select Case newStatus
    Case ReplayGameStatus_PLAY
        setAudioPause False
        setConsolePause False
        If gameTimeIsStopped() Then
            gameTimeStart
        End If
    Case ReplayGameStatus_PAUSE
        If Not gameTimeIsStopped() Then
            gameTimeStop
        End If
        setAudioPause True
        setConsolePause True
    End Select

    widgScheduleTask WzFunc_new(FN_replayRefreshButtonStates_task, self)
End Sub

Private Sub ReplayControllerWidget_toggleFastForward(ByVal self As WIDGET, Optional ByVal reverseDirection As Boolean = False)
    Dim currFastForward As Long, newFastForward As Long
    currFastForward = getMaxFastForwardTicks()
    newFastForward = 0
    If currFastForward = 0 Then
        If Not reverseDirection Then newFastForward = 1 Else newFastForward = WZ_MAX_REPLAY_FASTFORWARD_TICKS
    ElseIf currFastForward = 1 Then
        If Not reverseDirection Then newFastForward = WZ_MAX_REPLAY_FASTFORWARD_TICKS Else newFastForward = 0
    ElseIf reverseDirection And currFastForward = WZ_MAX_REPLAY_FASTFORWARD_TICKS Then
        newFastForward = 1
    End If
    setMaxFastForwardTicks True, newFastForward, True

    widgScheduleTask WzFunc_new(FN_replayRefreshButtonStates_task, self)
End Sub

Private Sub ReplayControllerWidget_refreshButtonStates(ByVal self As WIDGET)
    Dim d As ReplayControllerWidget, maxFastForwardTicks As Long, outline As Long
    Set d = ReplayControllerWidget_(self)
    Select Case ReplayControllerWidget_getReplayGameStatus()
    Case ReplayGameStatus_PLAY
        ' play button shoud be disabled
        d.playButton.setState WBUT_DISABLE
        ' pause + fast-forward buttons should be enabled
        d.pauseButton.setState 0
    Case ReplayGameStatus_PAUSE
        ' play button should be enabled
        d.playButton.setState 0
        ' pause + fast-forward buttons should be disabled
        d.pauseButton.setState WBUT_DISABLE
    End Select

    ' fast-forward state is based on getMaxFastForwardTicks
    maxFastForwardTicks = getMaxFastForwardTicks()
    If maxFastForwardTicks > 0 Then
        ' fast-forward is enabled
        If maxFastForwardTicks >= WZ_MAX_REPLAY_FASTFORWARD_TICKS Then outline = OutlineState_Double Else outline = OutlineState_Single
    Else
        ' normal speed
        outline = OutlineState_None
    End If
    W_REPLAY_CONTROL_BUTTON_(d.fastForwardButton).outline = outline
    d.lastUpdateTime = realTime
End Sub

Private Function ReplayControllerWidget_makeReplayActionButton(ByVal imagePath As String) As WIDGET
    Dim button As WIDGET
    Set button = WIDGET_alloc(K_W_REPLAY_CONTROL_BUTTON)
    W_REPLAY_CONTROL_BUTTON_construct button
    button.style = button.style Or WBUT_SECONDARY
    W_BUTTON_(button).FontID = font_regular
    Set W_REPLAY_CONTROL_BUTTON_(button).pButtonImage = gfx_api_context_loadTextureFromFile(imagePath, texture_type_user_interface)
    button.setGeometry 0, 0, REPLAY_ACTION_BUTTONS_WIDTH, REPLAY_ACTION_BUTTONS_HEIGHT
    Set ReplayControllerWidget_makeReplayActionButton = button
End Function

Public Function ReplayControllerWidget_make() As WIDGET
    Dim result As WIDGET, d As ReplayControllerWidget, buttonsY0 As Long, neededWidth As Long, neededHeight As Long
    Set result = WIDGET_alloc(K_ReplayControllerWidget)
    ReplayControllerWidget_construct result
    Set d = ReplayControllerWidget_(result)

    ' Add title text
    Set d.titleLabel = W_LABEL_new()
    W_LABEL_setFont d.titleLabel, font_regular_bold, psPalette(WZCOL_FORM_TEXT)
    d.titleLabel.setString "Replay"
    d.titleLabel.setGeometry REPLAY_ACTION_BUTTONS_SPACING, REPLAY_ACTION_BUTTONS_SPACING, W_LABEL_getMaxLineWidth(d.titleLabel), iV_GetTextLineSize(font_regular_bold)
    W_LABEL_setCacheNeverExpires d.titleLabel, True
    result.attach d.titleLabel
    d.titleLabel.setCalcLayout WzFunc_new(FN_replayTitleLabel_calcLayout)

    buttonsY0 = d.titleLabel.y() + d.titleLabel.height() + REPLAY_ACTION_BUTTONS_SPACING

    ' Add pauseButton button
    Set d.pauseButton = ReplayControllerWidget_makeReplayActionButton("images/replay/pause.png")
    d.pauseButton.setTip "Pause"
    result.attach d.pauseButton
    W_BUTTON_addOnClickHandler d.pauseButton, WzFunc_new(FN_replayPause_onClick)
    d.pauseButton.move 0, buttonsY0
    d.pauseButton.setCalcLayout WzFunc_new(FN_replayPause_calcLayout)

    ' Add "play" button
    Set d.playButton = ReplayControllerWidget_makeReplayActionButton("images/replay/play.png")
    d.playButton.setTip "Resume"
    result.attach d.playButton
    W_BUTTON_addOnClickHandler d.playButton, WzFunc_new(FN_replayPlay_onClick)
    d.playButton.move 0, buttonsY0
    d.playButton.setCalcLayout WzFunc_new(FN_replayPlay_calcLayout)

    ' Add "fast-forwards" button
    Set d.fastForwardButton = ReplayControllerWidget_makeReplayActionButton("images/replay/fast_forward.png")
    d.fastForwardButton.setTip "Fast-Forward"
    result.attach d.fastForwardButton
    W_BUTTON_addOnClickHandler d.fastForwardButton, WzFunc_new(FN_replayFastForward_onClick)
    d.fastForwardButton.move 0, buttonsY0
    d.fastForwardButton.setCalcLayout WzFunc_new(FN_replayFastForward_calcLayout)

    ' set overall width
    neededWidth = (REPLAY_ACTION_BUTTONS_SPACING * 4) + d.pauseButton.width() + d.playButton.width() + d.fastForwardButton.width()
    neededWidth = MAX_(neededWidth, REPLAY_ACTION_BUTTONS_SPACING + d.titleLabel.width() + REPLAY_ACTION_BUTTONS_SPACING)
    neededHeight = d.pauseButton.y() + d.pauseButton.height() + REPLAY_ACTION_BUTTONS_SPACING
    result.setGeometry result.x(), result.y(), neededWidth, neededHeight

    ReplayControllerWidget_refreshButtonStates result

    Set ReplayControllerWidget_make = result
End Function

Public Sub ReplayControllerWidget_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    ' currently, no-op
End Sub

Public Sub ReplayControllerWidget_geometryChanged(ByVal self As WIDGET)
    Dim d As ReplayControllerWidget
    Set d = ReplayControllerWidget_(self)
    If Not d.titleLabel Is Nothing Then
        d.titleLabel.callCalcLayout
    End If
    If Not d.pauseButton Is Nothing Then
        d.pauseButton.callCalcLayout
    End If
    If Not d.playButton Is Nothing Then
        d.playButton.callCalcLayout
    End If
    If Not d.fastForwardButton Is Nothing Then
        d.fastForwardButton.callCalcLayout
    End If
End Sub

' (screenSizeDidChange: W_FORM's, as the override only calls it)

Public Sub ReplayControllerWidget_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    Dim d As ReplayControllerWidget
    Set d = ReplayControllerWidget_(self)
    If Not ULT(realTime - d.lastUpdateTime, d.autoStateRefreshInterval) Then   ' realTime - lastUpdateTime >= autoStateRefreshInterval
        ' trigger update call
        widgScheduleTask WzFunc_new(FN_replayRefreshButtonStates_task, self)
        d.lastUpdateTime = realTime
    End If
End Sub

Public Function createReplayControllerOverlay() As Boolean
    Dim replayControllerForm As WIDGET
    If Not (Not psWScreen Is Nothing) Then ASSERT_FAILURE "psWScreen != nullptr", "src/hci.cpp:879", "createReplayControllerOverlay", "psWScreen is not initialized yet?"
    If Not (replayOverlayScreen Is Nothing) Then ASSERT_FAILURE "replayOverlayScreen == nullptr", "src/hci.cpp:880", "createReplayControllerOverlay", "Already initialized": createReplayControllerOverlay = False: Exit Function

    ' Initialize the replay overlay screen
    Set replayOverlayScreen = W_SCREEN_make()
    replayOverlayScreen.psForm.hide

    ' Create the Replay Controller form
    Set replayControllerForm = ReplayControllerWidget_make()
    replayOverlayScreen.psForm.attach replayControllerForm

    ' Position the Replay Controller form
    replayControllerForm.setCalcLayout WzFunc_new(FN_replayController_calcLayout)

    widgRegisterOverlayScreenOnTopOfScreen replayOverlayScreen, psWScreen

    createReplayControllerOverlay = True
End Function

' (VB) std::dynamic_pointer_cast<ReplayControllerWidget>(w): w when it is one, else Nothing
Private Function asReplayControllerWidget(ByVal w As WIDGET) As WIDGET
    If w Is Nothing Then Exit Function
    If WIDGET_isKindOf(w, K_ReplayControllerWidget) Then Set asReplayControllerWidget = w
End Function

' the lambdas
Public Function hci_replay_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim psWidget As WIDGET, psParent As WIDGET, pd As ReplayControllerWidget, x0 As Long, psOptionsButton As WIDGET
    Dim visibleOptionsbutton As Boolean, optionsWidth As Long
    Select Case f.id
    Case FN_replayTitleLabel_calcLayout             ' LAMBDA_CALCLAYOUT_SIMPLE (WIDGET *psWidget)
        Set psWidget = args(0)
        Set psParent = asReplayControllerWidget(psWidget.parent())
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/hci.cpp:764", "make", "No parent": Exit Function
        x0 = psParent.width() - (psWidget.width() + REPLAY_ACTION_BUTTONS_SPACING)
        psWidget.setGeometry x0, REPLAY_ACTION_BUTTONS_SPACING, psWidget.width(), psWidget.height()

    Case FN_replayPause_onClick                     ' [](W_BUTTON& button)
        Set psWidget = args(0)
        Set psParent = asReplayControllerWidget(psWidget.parent())
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/hci.cpp:777", "make", "No parent": Exit Function
        ReplayControllerWidget_setReplayGameStatus psParent, ReplayGameStatus_PAUSE

    Case FN_replayPause_calcLayout                  ' LAMBDA_CALCLAYOUT_SIMPLE (WIDGET *psWidget)
        Set psWidget = args(0)
        x0 = REPLAY_ACTION_BUTTONS_SPACING
        psWidget.setGeometry x0, psWidget.y(), psWidget.width(), psWidget.height()

    Case FN_replayPlay_onClick                      ' [](W_BUTTON& button)
        Set psWidget = args(0)
        Set psParent = asReplayControllerWidget(psWidget.parent())
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/hci.cpp:792", "make", "No parent": Exit Function
        ReplayControllerWidget_setReplayGameStatus psParent, ReplayGameStatus_PLAY

    Case FN_replayPlay_calcLayout                   ' LAMBDA_CALCLAYOUT_SIMPLE (WIDGET *psWidget)
        Set psWidget = args(0)
        Set psParent = asReplayControllerWidget(psWidget.parent())
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/hci.cpp:798", "make", "No parent": Exit Function
        Set pd = ReplayControllerWidget_(psParent)
        x0 = pd.pauseButton.x() + pd.pauseButton.width() + REPLAY_ACTION_BUTTONS_SPACING
        psWidget.setGeometry x0, psWidget.y(), psWidget.width(), psWidget.height()

    Case FN_replayFastForward_onClick               ' [](W_BUTTON& button)
        Set psWidget = args(0)
        Set psParent = asReplayControllerWidget(psWidget.parent())
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/hci.cpp:809", "make", "No parent": Exit Function
        ReplayControllerWidget_toggleFastForward psParent, W_BUTTON_getOnClickButtonPressed(psWidget) = WKEY_SECONDARY

    Case FN_replayFastForward_calcLayout            ' LAMBDA_CALCLAYOUT_SIMPLE (WIDGET *psWidget)
        Set psWidget = args(0)
        Set psParent = asReplayControllerWidget(psWidget.parent())
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/hci.cpp:815", "make", "No parent": Exit Function
        Set pd = ReplayControllerWidget_(psParent)
        x0 = pd.playButton.x() + pd.playButton.width() + REPLAY_ACTION_BUTTONS_SPACING
        psWidget.setGeometry x0, psWidget.y(), psWidget.width(), psWidget.height()

    Case FN_replayRefreshButtonStates_task          ' [sharedThis]() / [replayController]()
        ReplayControllerWidget_refreshButtonStates f.cap(0)

    Case FN_replayController_calcLayout             ' [](WIDGET *psWidget)
        Set psWidget = args(0)
        Set psOptionsButton = widgFormGetFromID(psWScreen.psForm, IDRET_OPTIONS)
        If Not psOptionsButton Is Nothing Then
            If Not WIDGET_isKindOf(psOptionsButton, K_W_INGAMEOPTIONS_BUTTON) Then Set psOptionsButton = Nothing
        End If
        If Not psOptionsButton Is Nothing Then visibleOptionsbutton = psOptionsButton.visible()
        If visibleOptionsbutton Then optionsWidth = W_INGAMEOPTIONS_BUTTON_getMarginX(psOptionsButton) + psOptionsButton.width()
        x0 = screenWidth - optionsWidth - psWidget.width() - (REPLAY_ACTION_BUTTONS_SPACING * 2)
        psWidget.move MAX_(0, x0), 0
    End Select
End Function
