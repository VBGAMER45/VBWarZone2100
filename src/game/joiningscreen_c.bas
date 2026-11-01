Attribute VB_Name = "joiningscreen_c"
' src/screens/joiningscreen.cpp, joiningscreen.h, whole: the overlay screen that joins a game (the connection
' attempts, the netcode handshake, the join messages, the password prompt, the failure and rejection displays) and
' its public API (startJoiningAttempt, startLobbyJoiningAttempt, startJoinRedirectAttempt, shutdownJoiningAttempt,
' resetJoinRedirectTracking, createJoiningIndeterminateProgressWidget). Its widget classes' fields are the data
' classes named as the classes (WzJoiningIndeterminateIndicatorWidget's is WzJoiningIndicatorData: VB's 31
' characters). struct WzJoiningGameScreen : public W_SCREEN is a W_SCREEN (W_SCREEN_make with its root form), as
' chatscreen_c's; ~WzJoiningGameScreen's setFocus(nullptr) is on the dying screen and has no effect left;
' psCurrentJoiningAttemptScreen is a strong reference that the shutdown paths drop (an address could outlive the
' screen). The weak pointers that a scheduled task, wzAsyncExecOnMainThread or the connection's callback can outlive
' are widgweak_vb serials. The connection is J_DESIGN's Contract C (IClientConnection, IConnectionPollGroup,
' WzConnectionProvider's openClientConnectionAsync, whose callback is JoiningOpenConnectionCallback, the exe's
' IOpenConnectionCallback; a std::error_code is a Long, read with net_error*); the keys are Contract B's EcKey and
' SessionKeys (UDTs; EcKey::Key and Sig are byte arrays). tmpJoiningQueuePair (the NetQueuePair* this screen owns in
' C++) is WzNet's, the pair QUEUE_TRANSIENT_JOIN resolves to (NETtransientJoin*); has_tmpJoiningQueuePair says it is
' there. netlobby (curl) is not ported: its stand-ins answer as when the lobby cannot be reached; their results
' reach the lambdas as store handles, an errors map as an array (in its order: 4.7.0's is the hash's). A lambda's
' captured PLAYERSTATS (a UDT, no Variant) is a handle into joinPlayerStatsStore; a JoinConnectionDescription
' passed to onSuccessFunc is its host, port and type. Translations (_()) are not ported: the English text is used.
' (VB) widget class WzJoiningActionButton : W_BUTTON
' (VB) widget class WzJoiningPasswordPrompt : WIDGET
' (VB) widget class WzJoiningIndeterminateIndicatorWidget : WIDGET
' (VB) widget class WzJoiningStatusForm : W_FORM
' (VB) widget class WzJoiningGameScreen_HandlerRoot : W_CLICKFORM
Option Explicit

Private Const MAX_CONNECTED_PLAYERS As Long = MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS
Private Const StringSize As Long = 64                   ' netplay.h: size of strings used.
Private Const modlist_string_size As Long = 255         ' netplay.h: For a concatenated list of mods
Private Const MaxMsgSize As Long = 32768                ' netplay.h: max size of a message in bytes.
Private Const MAX_JOIN_REJECT_REASON As Long = 2048     ' multiplay.h
Private Const NET_JOIN_PLAYER As Long = 0               ' netplay.h: enum NET_JOIN_PLAYERTYPE
Private Const NET_JOIN_SPECTATOR As Long = 1
Private Const BLIND_MODE_MAX As Long = BLIND_MODE_BLIND_GAME_SIMPLE_LOBBY  ' multiplaydefs.h

Private Const NET_READ_TIMEOUT As Long = 0              ' constexpr std::chrono::milliseconds NET_READ_TIMEOUT{ 0 };
Private Const HOST_RESPONSE_TIMEOUT As Long = 10000     ' constexpr uint32_t

' WzJoiningActionButton's const int InternalPadding = 12;
Private Const ActionButton_InternalPadding As Long = 12
' WzJoiningPasswordPrompt's const int32_t minPasswordBoxHeight = 20;
Private Const minPasswordBoxHeight As Long = 20
' WzJoiningIndeterminateIndicatorWidget's static constexprs
Private Const Indicator_PaddingBetween As Long = 2
Private Const Indicator_InstanceCount As Long = 3
Private Const Indicator_AnimTimePerInstance As Long = 750
Private Const Indicator_TotalAnimTime As Long = Indicator_AnimTimePerInstance * Indicator_InstanceCount
' WzJoiningStatusForm's static constexprs
Private Const TitleContentsPadding As Long = 20
Private Const StatusForm_InternalPadding As Long = 20
Private Const DetailsLabelParagraphPadding As Long = 5
Private Const MaxParagraphHeight As Long = 185
' WzJoiningGameScreen_HandlerRoot's static constexpr int32_t INTERNAL_PADDING = 15; (unused in the C++)
' and its const members
Private Const expectedInitialAckSize As Long = 4        ' sizeof(uint32_t)
Private Const minimumTimeBeforeAutoClose As Long = 300  ' std::chrono::milliseconds(300)
Private Const CLIENT_OPEN_ASYNC_TIMEOUT As Long = 15000 ' Default timeout of 15s

' enum class WzJoiningGameScreen_HandlerRoot::JoiningState
Public Enum JoiningState
    JoiningState_NeedsLobbyErrorResolve                 ' Waiting for user to resolve the error
    JoiningState_FetchingLobbyJoinInfo                  ' Waiting for lobby to return join info
    JoiningState_NeedsPassword                          ' Waiting for user to enter a password
    JoiningState_AwaitingConnection                     ' Waiting for background thread to (hopefully) yield an open connection (socket)
    JoiningState_AwaitingInitialNetcodeHandshakeAck     ' Waiting for response to initial netcode version handshake
    JoiningState_ProcessingJoinMessages                 ' Waiting for initial join-related net messages, and responding as needed
    JoiningState_Failure                                ' Join attempt failed
    JoiningState_SuccessPendingClose                    ' Join attempt was successful - handed off to onSuccessFunc, but still waiting to close display
    JoiningState_Success                                ' Success
End Enum

' enum class WzJoiningGameScreen_HandlerRoot::FailureDetails::Source
Public Enum FailureDetails_Source
    FailureDetails_Source_NetLobby
    FailureDetails_Source_Host
    FailureDetails_Source_Other
End Enum

' joiningscreen.h: struct ExpectedHostProperties
Public Type ExpectedHostProperties
    has_hostPublicKey As Boolean        ' optional<EcKey::Key> hostPublicKey
    hostPublicKey() As Byte
    has_gamePassword As Boolean         ' optional<std::string> gamePassword
    gamePassword As String
End Type

' struct WzJoiningGameScreen_HandlerRoot::FailureDetails
Public Type JoiningFailureDetails
    context As String
    resultMessage As String
    source As FailureDetails_Source     ' = Source::Other
End Type

' struct WzJoiningGameScreen_HandlerRoot::InternalConnectionInfo
Public Type InternalConnectionInfo
    conn As JoinConnectionDescription
    joinToken As String
End Type

' the lambdas (WzFunc: joiningscreen_c_fn)
Public Const FN_joiningPasswordPrompt_focusTask As Long = 16000      ' submitPassword: widgScheduleTask([weakSelf]() {...})
Public Const FN_joiningPasswordPrompt_onReturn As Long = 16001       ' initialize: passwordBox->setOnReturnHandler([weakSelf](W_EDITBOX& widg) {...})
Public Const FN_joiningPasswordPrompt_onEditingStopped As Long = 16002 ' initialize: passwordBox->setOnEditingStoppedHandler([weakSelf](W_EDITBOX&) {...})
Public Const FN_joiningPasswordPrompt_okClick As Long = 16003        ' initialize: okayButton->addOnClickHandler([weakSelf](W_BUTTON&) {...})
Public Const FN_joiningStatusForm_actionClick As Long = 16004        ' initialize: actionButton->addOnClickHandler([](W_BUTTON& but) {...})
Public Const FN_joiningStatusForm_wrappedSubmit As Long = 16005      ' displayPasswordPrompt: [weakSelf, submitFunc](WzString str) {...}
Public Const FN_joiningHandlerRoot_lobbyJoinResult As Long = 16006   ' requestLobbyJoin: [weakSelf, lobbyAddress, gameId, specificIpVersion](netlobby::JoinResult result) {...}
Public Const FN_joiningHandlerRoot_lobbyErrorResolved As Long = 16007 ' handleLobbyRequestJoinResult: [weakSelf, lobbyAddressCopy, gameIdCopy, specificIpVersion](resolutionData) {...}
Public Const FN_joiningHandlerRoot_formInitCalcLayout As Long = 16008 ' makeInternal: sInit.calcLayout = LAMBDA_CALCLAYOUT_SIMPLE({...})
Public Const FN_joiningHandlerRoot_progressFormCalcLayout As Long = 16009 ' initialize: joiningProgressForm->setCalcLayout(LAMBDA_CALCLAYOUT_SIMPLE({...}))
Public Const FN_joiningHandlerRoot_passwordSubmitted As Long = 16010 ' promptForPassword: [weakSelf](WzString newPassword) {...}
Public Const FN_joiningHandlerRoot_passwordTask As Long = 16011      ' promptForPassword: widgScheduleTask([weakSelf, newPassword]() {...})
Public Const FN_joiningHandlerRoot_resultOnMainThread As Long = 16012 ' processOpenConnectionResultOnMainThread: wzAsyncExecOnMainThread([weakSelf, connectionIdx, ppResult]() {...})
Public Const FN_joiningScreen_cancelPressed As Long = 16013          ' WzJoiningGameScreen::make(...): newRootFrm->onCancelPressed = [psWeakHelpOverlayScreen]() {...}
Public Const FN_joining_onSuccess As Long = 16014                    ' startJoiningAttemptInternal: [playerStats](const JoinConnectionDescription& connection) {...}
Public Const FN_joining_onFailure As Long = 16015                    ' startJoiningAttemptInternal / startLobbyJoiningAttemptInternal: []() {...}
Public Const FN_joining_registerOverlayTask As Long = 16016          ' widgScheduleTask([screen]() {...})
Public Const FN_joining_lobbyOnSuccess As Long = 16017               ' startLobbyJoiningAttemptInternal: [lobbyAddressCopy, gameIdCopy, playerStats](connection) {...}
Public Const FN_joining_lobbyGameIdMsgTask As Long = 16018           ' widgScheduleTask([gameIdCopy]() {...})
Public Const FN_joining_changeTitleUITask As Long = 16019            ' handleJoinSuccess: widgScheduleTask([]() {...})

' MARK: - Globals

Private psCurrentJoiningAttemptScreen As W_SCREEN  ' static std::weak_ptr<WzJoiningGameScreen> (VB: a strong reference, see above)
Private currentNumRedirectAttempts As Long          ' static size_t currentNumRedirectAttempts = 0;

Private joinPlayerStatsStore() As PLAYERSTATS_      ' (VB) the lambdas' captured PLAYERSTATS
Private joinPlayerStatsCount As Long

Private readBuffer(0 To MaxMsgSize - 1) As Byte     ' processJoining's uint8_t readBuffer[MaxMsgSize]

'--------------------------------------------------------------------------
' (VB) helpers

Private Function jmax(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then jmax = a Else jmax = b
End Function

Private Function jmin(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then jmin = a Else jmin = b
End Function

' a PIELIGHT with its byte.a replaced
Private Function pielightWithAlpha(ByVal c As Long, ByVal a As Long) As Long
    pielightWithAlpha = (c And &HFFFFFF) Or SHL(a And &HFF&, 24)
End Function

' PIELIGHT::isTransparent(): byte.a == 0
Private Function pielightIsTransparent(ByVal c As Long) As Boolean
    pielightIsTransparent = (SHRU(c, 24) = 0)
End Function

' the number of bytes in a dynamic array (0 when not allocated)
Private Function bytesCount(b() As Byte) As Long
    On Error GoTo none
    bytesCount = UBound(b) - LBound(b) + 1
    Exit Function
none:
    bytesCount = 0
End Function

' std::vector<uint8_t> == (the same bytes)
Private Function bytesEqual(a() As Byte, b() As Byte) As Boolean
    Dim n As Long, i As Long
    n = bytesCount(a)
    If n <> bytesCount(b) Then Exit Function
    For i = 0 To n - 1
        If a(i) <> b(i) Then Exit Function
    Next
    bytesEqual = True
End Function

' sstrcpy(dest, src) into a char[size]: at most size - 1 characters
Private Function sstrcpyTo(ByVal src As String, ByVal size As Long) As String
    If Len(src) > size - 1 Then sstrcpyTo = Left$(src, size - 1) Else sstrcpyTo = src
End Function

' the length of a string's UTF-8 (WzString::toUtf8().size())
Private Function utf8Size(ByVal s As String) As Long
    Dim b() As Byte
    utf8Size = utf8_encode(s, b)
End Function

' activity.cpp: std::string to_string(const ActivitySink::GameMode& mode)
Private Function gameModeToString(ByVal mode As ActivitySink_GameMode) As String
    Select Case mode
    Case ActivitySink_GameMode_MENUS: gameModeToString = "Menus"
    Case ActivitySink_GameMode_TUTORIAL: gameModeToString = "Tutorial"
    Case ActivitySink_GameMode_CAMPAIGN: gameModeToString = "Campaign"
    Case ActivitySink_GameMode_CHALLENGE: gameModeToString = "Challenge"
    Case ActivitySink_GameMode_SKIRMISH: gameModeToString = "Skirmish"
    Case ActivitySink_GameMode_HOSTING_IN_LOBBY: gameModeToString = "Hosting::InLobby"
    Case ActivitySink_GameMode_JOINING_IN_PROGRESS: gameModeToString = "Joining::InProgress"
    Case ActivitySink_GameMode_JOINING_IN_LOBBY: gameModeToString = "Joining::InLobby"
    Case ActivitySink_GameMode_MULTIPLAYER: gameModeToString = "Multiplayer"
    Case Else: gameModeToString = ""    ' silence warning
    End Select
End Function

Private Function storeJoinPlayerStats(st As PLAYERSTATS_) As Long
    ReDim Preserve joinPlayerStatsStore(0 To joinPlayerStatsCount)
    joinPlayerStatsStore(joinPlayerStatsCount) = st
    joinPlayerStatsCount = joinPlayerStatsCount + 1
    storeJoinPlayerStats = joinPlayerStatsCount - 1
End Function

' MARK: - WzJoiningStatusForm

' WzJoiningActionButton() {}
Public Sub WzJoiningActionButton_construct(ByVal self As WIDGET)
    W_BUTTON_construct self
    self.kind = K_WzJoiningActionButton ' (VB) virtual calls from here on reach WzJoiningActionButton's overrides
End Sub

' void WzJoiningActionButton::initialize(const WzString& text)
Private Sub WzJoiningActionButton_initialize(ByVal self As WIDGET, ByVal text As String)
    Dim b As W_BUTTON, minButtonWidthForText As Long
    self.setString text
    Set b = W_BUTTON_(self)
    b.FontID = font_regular_bold
    minButtonWidthForText = iV_GetTextWidth(b.pText, b.FontID)
    self.setGeometry 0, 0, minButtonWidthForText + ActionButton_InternalPadding, iV_GetTextLineSize(b.FontID) + ActionButton_InternalPadding
End Sub

' static std::shared_ptr<WzJoiningActionButton> make(const WzString& text)
Public Function WzJoiningActionButton_make(ByVal text As String) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_WzJoiningActionButton)
    WzJoiningActionButton_construct widget
    WzJoiningActionButton_initialize widget, text
    Set WzJoiningActionButton_make = widget
End Function

Public Function WzJoiningActionButton_idealHeight(ByVal self As WIDGET) As Long
    WzJoiningActionButton_idealHeight = iV_GetTextLineSize(W_BUTTON_(self).FontID) + ActionButton_InternalPadding
End Function

Public Sub WzJoiningActionButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x0 As Long, y0 As Long, x1 As Long, y1 As Long, haveText As Boolean, isDown As Boolean, isDisabled As Boolean
    Dim isHighlight As Boolean, colour As Long, fillClr As Long, d As WzJoiningActionButton, b As W_BUTTON
    Dim fw As Long, fx As Long, fy As Long
    Set b = W_BUTTON_(self)
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    x1 = x0 + self.width()
    y1 = y0 + self.height()

    haveText = Len(b.pText) <> 0

    isDown = (self.getState() And (WBUT_DOWN Or WBUT_LOCK Or WBUT_CLICKLOCK)) <> 0
    isDisabled = (self.getState() And WBUT_DISABLE) <> 0
    isHighlight = (Not isDisabled) And ((self.getState() And WBUT_HIGHLIGHT) <> 0)

    If isDisabled Then                  ' unavailable
        colour = psPalette(WZCOL_TEXT_DARK)
    Else                                ' available
        If isHighlight Then             ' hilight
            colour = psPalette(WZCOL_TEXT_BRIGHT)
        Else                            ' don't highlight
            colour = psPalette(WZCOL_TEXT_MEDIUM)
        End If
    End If

    If isHighlight Or isDown Then
        fillClr = pal_RGBA(255, 255, 255, 30)
        pie_UniTransBoxFill x0, y0, x1, y1, fillClr
    End If
    iV_Box x0, y0, x1, y1, colour

    If haveText Then
        Set d = WzJoiningActionButton_(self)
        d.wzText.setText b.pText, b.FontID
        fw = d.wzText.width()
        fx = x0 + (self.width() - fw) \ 2
        fy = y0 + (self.height() - d.wzText.lineSize()) \ 2 - d.wzText.aboveBase()
        d.wzText.render fx, fy, colour
    End If
End Sub

' WzJoiningPasswordPrompt() { }
Public Sub WzJoiningPasswordPrompt_construct(ByVal self As WIDGET)
    Dim d As WzJoiningPasswordPrompt
    WIDGET_construct self
    self.kind = K_WzJoiningPasswordPrompt ' (VB) virtual calls from here on reach WzJoiningPasswordPrompt's overrides
    Set d = WzJoiningPasswordPrompt_(self)
    d.invalidPasswordPlaceholder = False
    d.weakSerial = WeakWidget_register(self)
End Sub

Public Sub WzJoiningPasswordPrompt_destruct(ByVal self As WIDGET)
    WeakWidget_unregister WzJoiningPasswordPrompt_(self).weakSerial
End Sub

Public Sub WzJoiningPasswordPrompt_setOnSubmit(ByVal self As WIDGET, ByVal onSubmit_ As WzFunc)
    Set WzJoiningPasswordPrompt_(self).onSubmit = onSubmit_
End Sub

Public Function WzJoiningPasswordPrompt_idealHeight(ByVal self As WIDGET) As Long
    WzJoiningPasswordPrompt_idealHeight = jmax(minPasswordBoxHeight, WzJoiningPasswordPrompt_(self).okayButton.idealHeight())
End Function

Public Sub WzJoiningPasswordPrompt_givePasswordBoxFocus(ByVal self As WIDGET)
    Dim context As W_CONTEXT    ' W_CONTEXT::ZeroContext(): a new W_CONTEXT is zeroed
    ' give the password box the focus
    W_EDITBOX_simulateClick WzJoiningPasswordPrompt_(self).passwordBox, context, True
End Sub

Public Sub WzJoiningPasswordPrompt_clearAndSetCurrentEntryAsInvalid(ByVal self As WIDGET)
    Dim d As WzJoiningPasswordPrompt, oldPass As String
    Set d = WzJoiningPasswordPrompt_(self)
    oldPass = W_EDITBOX_getString(d.passwordBox)
    If Len(oldPass) = 0 Then
        Exit Sub
    End If

    W_EDITBOX_setString d.passwordBox, ""
    W_EDITBOX_setPlaceholder d.passwordBox, oldPass
    W_EDITBOX_setPlaceholderTextColor d.passwordBox, True, pal_RGBA(255, 0, 0, 170)
    d.invalidPasswordPlaceholder = True
End Sub

' static std::shared_ptr<WzJoiningPasswordPrompt> make(const OnSubmitFunc& onSubmit)
Public Function WzJoiningPasswordPrompt_make(ByVal onSubmit As WzFunc) As WIDGET
    Dim result As WIDGET
    Set result = WIDGET_alloc(K_WzJoiningPasswordPrompt)
    WzJoiningPasswordPrompt_construct result
    WzJoiningPasswordPrompt_initialize result, onSubmit
    Set WzJoiningPasswordPrompt_make = result
End Function

Private Sub WzJoiningPasswordPrompt_submitPassword(ByVal self As WIDGET, ByVal str_ As String)
    Dim d As WzJoiningPasswordPrompt, weakSelf As Long
    Set d = WzJoiningPasswordPrompt_(self)
    If Len(str_) = 0 Then
        ' don't send empty password
        weakSelf = d.weakSerial
        widgScheduleTask WzFunc_new(FN_joiningPasswordPrompt_focusTask, weakSelf)
        Exit Sub
    End If

    If Not d.onSubmit Is Nothing Then
        WzFunc_call d.onSubmit, str_
    End If
End Sub

Public Sub WzJoiningPasswordPrompt_setDisabled(ByVal self As WIDGET, ByVal disabled As Boolean)
    Dim d As WzJoiningPasswordPrompt, lockedScreen As W_SCREEN
    Set d = WzJoiningPasswordPrompt_(self)
    If disabled Then
        ' ensure passwordBox does not have focus
        Set lockedScreen = self.screenPointer
        If Not lockedScreen Is Nothing Then
            If lockedScreen.hasFocus(d.passwordBox) Then
                lockedScreen.setFocus Nothing
            End If
        End If
    End If
    If disabled Then W_EDITBOX_setState d.passwordBox, WEDBS_DISABLE Else W_EDITBOX_setState d.passwordBox, 0
    If disabled Then d.okayButton.setState WBUT_DISABLE Else d.okayButton.setState 0
End Sub

' void WzJoiningPasswordPrompt::initialize(const OnSubmitFunc& _onSubmit)
Private Sub WzJoiningPasswordPrompt_initialize(ByVal self As WIDGET, ByVal onSubmit_ As WzFunc)
    Dim d As WzJoiningPasswordPrompt, weakSelf As Long
    Set d = WzJoiningPasswordPrompt_(self)
    Set d.onSubmit = onSubmit_

    weakSelf = d.weakSerial

    Set d.passwordBox = W_EDITBOX_new()
    self.attach d.passwordBox
    d.passwordBox.setGeometry 0, 0, 280, minPasswordBoxHeight
    W_EDITBOX_setBoxColours d.passwordBox, psPalette(WZCOL_MENU_BORDER), psPalette(WZCOL_MENU_BORDER), psPalette(WZCOL_MENU_BACKGROUND)
    W_EDITBOX_setMaxStringSize d.passwordBox, password_string_size
    W_EDITBOX_setPlaceholder d.passwordBox, "Enter password"

    W_EDITBOX_setOnReturnHandler d.passwordBox, WzFunc_new(FN_joiningPasswordPrompt_onReturn, weakSelf)
    W_EDITBOX_setOnEditingStoppedHandler d.passwordBox, WzFunc_new(FN_joiningPasswordPrompt_onEditingStopped, weakSelf)

    Set d.okayButton = WzJoiningActionButton_make("OK")
    self.attach d.okayButton
    W_BUTTON_addOnClickHandler d.okayButton, WzFunc_new(FN_joiningPasswordPrompt_okClick, weakSelf)
End Sub

Public Sub WzJoiningPasswordPrompt_geometryChanged(ByVal self As WIDGET)
    WzJoiningPasswordPrompt_recalcLayout self
End Sub

Private Sub WzJoiningPasswordPrompt_recalcLayout(ByVal self As WIDGET)
    Dim d As WzJoiningPasswordPrompt, w As Long, h As Long, contentsIdealHeight As Long, contentsActualHeight As Long
    Dim buttonX0 As Long, buttonY0 As Long, passwordBoxWidth As Long, passwordBoxY0 As Long
    Set d = WzJoiningPasswordPrompt_(self)
    w = self.width()
    h = self.height()

    contentsIdealHeight = jmax(d.okayButton.idealHeight(), minPasswordBoxHeight)
    contentsActualHeight = jmin(h, contentsIdealHeight)

    buttonX0 = w - d.okayButton.width()
    buttonY0 = jmax((h - contentsActualHeight) \ 2, 0)
    d.okayButton.setGeometry buttonX0, buttonY0, d.okayButton.width(), contentsActualHeight

    passwordBoxWidth = jmax(buttonX0 - 5, 0)
    passwordBoxY0 = jmax((h - contentsActualHeight) \ 2, 0)
    d.passwordBox.setGeometry 0, passwordBoxY0, passwordBoxWidth, contentsActualHeight
End Sub

' WzJoiningIndeterminateIndicatorWidget(iV_fonts fontID) : WIDGET()
Public Sub WzJoiningIndeterminateIndicatorWidget_construct(ByVal self As WIDGET, ByVal fontID As iV_fonts)
    WIDGET_construct self
    self.kind = K_WzJoiningIndeterminateIndicatorWidget ' (VB) virtual calls from here on reach its overrides
    WzJoiningIndeterminateIndicatorWidget_(self).periodText.setText ".", fontID
End Sub

' std::make_shared<WzJoiningIndeterminateIndicatorWidget>(fontID)
Public Function WzJoiningIndeterminateIndicatorWidget_new(ByVal fontID As iV_fonts) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_WzJoiningIndeterminateIndicatorWidget)
    WzJoiningIndeterminateIndicatorWidget_construct widget, fontID
    Set WzJoiningIndeterminateIndicatorWidget_new = widget
End Function

Public Function WzJoiningIndeterminateIndicatorWidget_idealWidth(ByVal self As WIDGET) As Long
    Dim t As WzText
    Set t = WzJoiningIndeterminateIndicatorWidget_(self).periodText
    WzJoiningIndeterminateIndicatorWidget_idealWidth = (t.width() * Indicator_InstanceCount) + (Indicator_PaddingBetween * (Indicator_InstanceCount - 1))
End Function

Public Function WzJoiningIndeterminateIndicatorWidget_idealHeight(ByVal self As WIDGET) As Long
    WzJoiningIndeterminateIndicatorWidget_idealHeight = WzJoiningIndeterminateIndicatorWidget_(self).periodText.lineSize()
End Function

Public Sub WzJoiningIndeterminateIndicatorWidget_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim t As WzText, x0 As Long, y0 As Long, timeFactor As Long, perInstanceAnim As Long, textY0 As Long
    Dim periodTextWidth As Long, i As Long, offsetX0 As Long, bckColor As Long, a As Long, c As Long
    Set t = WzJoiningIndeterminateIndicatorWidget_(self).periodText
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset

    ' static_cast<int>(realTime % TotalAnimTime) (realTime is a uint32_t)
    timeFactor = CLng(U32(realTime) - Int(U32(realTime) / Indicator_TotalAnimTime) * Indicator_TotalAnimTime)
    perInstanceAnim = Indicator_AnimTimePerInstance

    textY0 = y0 + (self.height() - t.lineSize()) \ 2 - t.aboveBase()

    periodTextWidth = t.width()
    For i = 0 To Indicator_InstanceCount - 1
        offsetX0 = i * (periodTextWidth + Indicator_PaddingBetween)
        bckColor = psPalette(WZCOL_TEXT_MEDIUM)
        c = timeFactor - (i * perInstanceAnim)
        If c < 0 Then c = 0
        If c > perInstanceAnim Then c = perInstanceAnim     ' clip<int>(..., 0, perInstanceAnim)
        a = 100 + (155 * c \ perInstanceAnim)
        bckColor = pielightWithAlpha(bckColor, a)
        t.render x0 + offsetX0, textY0, bckColor
    Next
End Sub

' WzJoiningStatusForm() { }
Public Sub WzJoiningStatusForm_construct(ByVal self As WIDGET)
    W_FORM_construct self
    self.kind = K_WzJoiningStatusForm ' (VB) virtual calls from here on reach WzJoiningStatusForm's overrides
    WzJoiningStatusForm_(self).weakSerial = WeakWidget_register(self)
End Sub

Public Sub WzJoiningStatusForm_destruct(ByVal self As WIDGET)
    WeakWidget_unregister WzJoiningStatusForm_(self).weakSerial
End Sub

Public Function WzJoiningStatusForm_idealWidth(ByVal self As WIDGET) As Long
    WzJoiningStatusForm_idealWidth = 540
End Function

Public Function WzJoiningStatusForm_idealHeight(ByVal self As WIDGET) As Long
    WzJoiningStatusForm_idealHeight = WzJoiningStatusForm_calculateNeededHeight(self, Not WzJoiningStatusForm_(self).detailsParagraph Is Nothing)
End Function

Public Function WzJoiningStatusForm_maximumHeight(ByVal self As WIDGET) As Long
    WzJoiningStatusForm_maximumHeight = WzJoiningStatusForm_calculateNeededHeight(self, True)
End Function

Private Function WzJoiningStatusForm_calculateNeededHeight(ByVal self As WIDGET, ByVal withDetailsParagraph As Boolean) As Long
    Dim d As WzJoiningStatusForm, result As Long
    Set d = WzJoiningStatusForm_(self)
    result = (StatusForm_InternalPadding * 2)
    result = result + d.titleLabel.height()
    result = result + TitleContentsPadding
    result = result + d.statusTitle.height()
    result = result + DetailsLabelParagraphPadding
    If withDetailsParagraph Then
        result = result + MaxParagraphHeight
    Else
        result = result + d.statusDetails.idealHeight()
    End If
    result = result + StatusForm_InternalPadding
    If Not d.resolvableJoinErrorWidget Is Nothing Then
        If d.resolvableJoinErrorWidget.visible() Then
            result = result + d.resolvableJoinErrorWidget.idealHeight() + StatusForm_InternalPadding
        End If
    End If
    result = result + d.actionButton.height()
    WzJoiningStatusForm_calculateNeededHeight = result
End Function

' static std::shared_ptr<WzJoiningStatusForm> make()
Public Function WzJoiningStatusForm_make() As WIDGET
    Dim result As WIDGET
    Set result = WIDGET_alloc(K_WzJoiningStatusForm)
    WzJoiningStatusForm_construct result
    WzJoiningStatusForm_initialize result
    Set WzJoiningStatusForm_make = result
End Function

Private Sub WzJoiningStatusForm_initialize(ByVal self As WIDGET)
    Dim d As WzJoiningStatusForm
    Set d = WzJoiningStatusForm_(self)
    d.startTime = widgSteadyNowMs()

    Set d.titleLabel = W_LABEL_new()
    self.attach d.titleLabel
    W_LABEL_setFont d.titleLabel, font_medium_bold, psPalette(WZCOL_TEXT_BRIGHT)
    d.titleLabel.setString "Connecting to Game"
    d.titleLabel.setGeometry 0, 0, W_LABEL_getMaxLineWidth(d.titleLabel), iV_GetTextLineSize(font_medium_bold)
    W_LABEL_setCanTruncate d.titleLabel, True

    Set d.progressIndicator = WzJoiningIndeterminateIndicatorWidget_new(font_medium_bold)
    self.attach d.progressIndicator
    d.progressIndicator.setGeometry 0, 0, d.progressIndicator.idealWidth(), d.progressIndicator.idealHeight()

    Set d.statusTitle = W_LABEL_new()
    self.attach d.statusTitle
    W_LABEL_setFont d.statusTitle, font_regular_bold, psPalette(WZCOL_TEXT_MEDIUM)
    d.statusTitle.setString "Status:"
    d.statusTitle.setGeometry 0, 0, W_LABEL_getMaxLineWidth(d.statusTitle), iV_GetTextLineSize(font_regular_bold)
    W_LABEL_setCanTruncate d.statusTitle, True

    Set d.statusDetails = W_LABEL_new()
    self.attach d.statusDetails
    W_LABEL_setFont d.statusDetails, font_regular, psPalette(WZCOL_TEXT_BRIGHT)
    d.statusDetails.setString ""
    d.statusDetails.setGeometry 0, 0, W_LABEL_getMaxLineWidth(d.statusDetails), iV_GetTextLineSize(font_regular)
    W_LABEL_setCanTruncate d.statusDetails, True

    Set d.actionButton = WzJoiningActionButton_make("Cancel")
    self.attach d.actionButton
    W_BUTTON_addOnClickHandler d.actionButton, WzFunc_new(FN_joiningStatusForm_actionClick)

    Set d.scrollableParagraphContainer = ScrollableListWidget_make()
    d.scrollableParagraphContainer.setGeometry 0, 0, 400, 45
    self.attach d.scrollableParagraphContainer
End Sub

Public Sub WzJoiningStatusForm_geometryChanged(ByVal self As WIDGET)
    WzJoiningStatusForm_recalcLayout self
End Sub

Private Sub WzJoiningStatusForm_recalcLayout(ByVal self As WIDGET)
    Dim d As WzJoiningStatusForm, w As Long, h As Long, usableWidth As Long, usableHeight As Long, indicatorWidth As Long
    Dim indicatorX0 As Long, titleUsableWidth As Long, lastLineY1 As Long, actionButtonX0 As Long, actionButtonY0 As Long
    Dim availableResolveErrorWidgetHeight As Long, resolveErrorWidgetHeight As Long, passwordPromptY0 As Long
    Dim statusDetailsX0 As Long, statusDetailsWidth As Long, paragraphHeight As Long
    Set d = WzJoiningStatusForm_(self)
    w = self.width()
    h = self.height()

    If w <= 0 And h <= 0 Then
        Exit Sub
    End If

    usableWidth = w - (StatusForm_InternalPadding * 2)
    usableHeight = h - (StatusForm_InternalPadding * 2)

    indicatorWidth = d.progressIndicator.idealWidth()
    indicatorX0 = (w - StatusForm_InternalPadding - indicatorWidth)
    d.progressIndicator.setGeometry indicatorX0, StatusForm_InternalPadding, indicatorWidth, d.progressIndicator.idealHeight()

    titleUsableWidth = usableWidth - indicatorWidth - StatusForm_InternalPadding
    d.titleLabel.setGeometry StatusForm_InternalPadding, StatusForm_InternalPadding, titleUsableWidth, iV_GetTextLineSize(font_medium_bold)
    lastLineY1 = d.titleLabel.y() + d.titleLabel.height()
    usableHeight = usableHeight - d.titleLabel.height()

    d.statusTitle.setGeometry StatusForm_InternalPadding, lastLineY1 + TitleContentsPadding, usableWidth, d.statusTitle.height()
    lastLineY1 = d.statusTitle.y() + d.statusTitle.height()
    usableHeight = usableHeight - (TitleContentsPadding + d.statusTitle.height())

    ' position button at bottom
    actionButtonX0 = (self.width() - d.actionButton.width()) \ 2
    actionButtonY0 = self.height() - StatusForm_InternalPadding - d.actionButton.height()
    d.actionButton.setGeometry actionButtonX0, actionButtonY0, d.actionButton.width(), d.actionButton.height()
    usableHeight = usableHeight - (d.actionButton.height() + StatusForm_InternalPadding)

    If Not d.resolvableJoinErrorWidget Is Nothing Then
        If d.resolvableJoinErrorWidget.visible() Then
            ' determine available height for resolvableJoinErrorWidget
            availableResolveErrorWidgetHeight = usableHeight - (DetailsLabelParagraphPadding + d.statusDetails.idealHeight() + StatusForm_InternalPadding)
            resolveErrorWidgetHeight = jmin(availableResolveErrorWidgetHeight, d.resolvableJoinErrorWidget.idealHeight())

            passwordPromptY0 = d.actionButton.y() - (StatusForm_InternalPadding + resolveErrorWidgetHeight)
            d.resolvableJoinErrorWidget.setGeometry StatusForm_InternalPadding, passwordPromptY0, usableWidth, resolveErrorWidgetHeight
            usableHeight = usableHeight - (d.resolvableJoinErrorWidget.height() + StatusForm_InternalPadding)
        End If
    End If

    ' both of these start at the same place, taking up the remaining vertical space, but only one is displayed at a time

    statusDetailsX0 = StatusForm_InternalPadding
    statusDetailsWidth = usableWidth
    d.statusDetails.setGeometry statusDetailsX0, lastLineY1 + DetailsLabelParagraphPadding, statusDetailsWidth, d.statusDetails.idealHeight()

    paragraphHeight = usableHeight - DetailsLabelParagraphPadding
    d.scrollableParagraphContainer.setGeometry StatusForm_InternalPadding, lastLineY1 + DetailsLabelParagraphPadding, usableWidth, paragraphHeight
End Sub

Private Sub WzJoiningStatusForm_updateTitle(ByVal self As WIDGET, ByVal title_ As String)
    WzJoiningStatusForm_(self).titleLabel.setString title_
End Sub

Private Sub WzJoiningStatusForm_updateStatusTitle(ByVal self As WIDGET, ByVal title_ As String)
    WzJoiningStatusForm_(self).statusTitle.setString title_
End Sub

Private Sub WzJoiningStatusForm_displayStatus(ByVal self As WIDGET, ByVal statusDescription As String)
    Dim d As WzJoiningStatusForm
    Set d = WzJoiningStatusForm_(self)
    If Not d.detailsParagraph Is Nothing Then
        ScrollableListWidget_clear d.scrollableParagraphContainer
        Set d.detailsParagraph = Nothing
    End If
    d.scrollableParagraphContainer.hide

    W_LABEL_setFormattedString d.statusDetails, statusDescription, &HFFFFFFFF, font_regular     ' std::numeric_limits<uint32_t>::max()
    d.statusDetails.show
End Sub

Private Sub WzJoiningStatusForm_displayDetailsParagraph(ByVal self As WIDGET, ByVal messageContents As String)
    Dim d As WzJoiningStatusForm, usableWidth As Long, messageContentsStr As String, maxLinePos As Long
    Set d = WzJoiningStatusForm_(self)
    If Not d.detailsParagraph Is Nothing Then
        ScrollableListWidget_clear d.scrollableParagraphContainer
    End If
    Set d.detailsParagraph = Paragraph_new()
    usableWidth = self.width() - (StatusForm_InternalPadding * 2)
    d.detailsParagraph.setGeometry StatusForm_InternalPadding, StatusForm_InternalPadding, usableWidth, MaxParagraphHeight
    Paragraph_setFontColour d.detailsParagraph, psPalette(WZCOL_TEXT_BRIGHT)
    Paragraph_setLineSpacing d.detailsParagraph, 5
    Paragraph_setFont d.detailsParagraph, font_regular

    messageContentsStr = messageContents
    maxLinePos = nthOccurrenceOfChar(messageContentsStr, vbLf, 10)
    If maxLinePos <> -1 Then        ' std::string::npos
        messageContentsStr = Left$(messageContentsStr, maxLinePos)
    End If
    Paragraph_addText d.detailsParagraph, messageContentsStr

    ScrollableListWidget_addItem d.scrollableParagraphContainer, d.detailsParagraph

    d.scrollableParagraphContainer.show
    d.statusDetails.hide
End Sub

Public Sub WzJoiningStatusForm_displayProgressStatus(ByVal self As WIDGET, ByVal statusDescription As String)
    WzJoiningStatusForm_displayStatus self, statusDescription
    WzJoiningStatusForm_(self).progressIndicator.show
End Sub

Public Sub WzJoiningStatusForm_displaySuccessStatus(ByVal self As WIDGET, ByVal statusDescription As String)
    Dim d As WzJoiningStatusForm
    Set d = WzJoiningStatusForm_(self)
    WzJoiningStatusForm_displayStatus self, statusDescription
    d.actionButton.setState WBUT_DISABLE
    d.actionButton.hide
End Sub

' void displayPasswordPrompt(const PasswordSubmitFunc& submitFunc) (PasswordSubmitFunc: void (WzString))
Public Sub WzJoiningStatusForm_displayPasswordPrompt(ByVal self As WIDGET, ByVal submitFunc As WzFunc)
    Dim d As WzJoiningStatusForm, weakSelf As Long, wrappedSubmitFunc As WzFunc, passwordPrompt As WIDGET
    Set d = WzJoiningStatusForm_(self)
    weakSelf = d.weakSerial
    Set wrappedSubmitFunc = WzFunc_new(FN_joiningStatusForm_wrappedSubmit, weakSelf, submitFunc)

    WzJoiningStatusForm_displayStatus self, "Game requires a password to join"

    Set passwordPrompt = d.resolvableJoinErrorWidget
    If Not passwordPrompt Is Nothing Then
        If Not WIDGET_isKindOf(passwordPrompt, K_WzJoiningPasswordPrompt) Then Set passwordPrompt = Nothing
    End If
    If passwordPrompt Is Nothing Then
        Set passwordPrompt = WzJoiningPasswordPrompt_make(wrappedSubmitFunc)
        WzJoiningStatusForm_displayResolvableJoinErrorWidget self, passwordPrompt
    Else
        WzJoiningPasswordPrompt_setOnSubmit passwordPrompt, wrappedSubmitFunc
        WzJoiningPasswordPrompt_clearAndSetCurrentEntryAsInvalid passwordPrompt
        passwordPrompt.show
        WzJoiningStatusForm_recalcLayout self
    End If
    WzJoiningPasswordPrompt_setDisabled passwordPrompt, False
    WzJoiningPasswordPrompt_givePasswordBoxFocus passwordPrompt
End Sub

Public Sub WzJoiningStatusForm_displayResolvableJoinErrorWidget(ByVal self As WIDGET, ByVal newResolvableJoinErrorWidget As WIDGET)
    Dim d As WzJoiningStatusForm
    Set d = WzJoiningStatusForm_(self)
    If Not (newResolvableJoinErrorWidget Is d.resolvableJoinErrorWidget) Then
        If Not d.resolvableJoinErrorWidget Is Nothing Then
            self.detach d.resolvableJoinErrorWidget
            Set d.resolvableJoinErrorWidget = Nothing
        End If
        Set d.resolvableJoinErrorWidget = newResolvableJoinErrorWidget
        self.attach d.resolvableJoinErrorWidget
    End If

    d.resolvableJoinErrorWidget.show
    WzJoiningStatusForm_recalcLayout self
End Sub

Public Sub WzJoiningStatusForm_displayUnableToJoinError(ByVal self As WIDGET, ByVal errorDetails As String)
    Dim d As WzJoiningStatusForm
    Set d = WzJoiningStatusForm_(self)
    If Not d.resolvableJoinErrorWidget Is Nothing Then
        d.resolvableJoinErrorWidget.hide
    End If
    d.progressIndicator.hide
    d.titleLabel.setString "Unable to Join Game"
    d.statusTitle.setString "Error:"
    WzJoiningStatusForm_displayStatus self, errorDetails
    d.actionButton.setString "Close"
End Sub

Public Sub WzJoiningStatusForm_displayNetLobbyError(ByVal self As WIDGET, ByVal errorDetails As String)
    Dim d As WzJoiningStatusForm
    Set d = WzJoiningStatusForm_(self)
    If Not d.resolvableJoinErrorWidget Is Nothing Then
        d.resolvableJoinErrorWidget.hide
    End If
    d.progressIndicator.hide
    d.titleLabel.setString "Unable to Join Game"
    d.statusTitle.setString "Message from Lobby:"
    WzJoiningStatusForm_displayDetailsParagraph self, errorDetails
    d.actionButton.setString "Close"
End Sub

Public Sub WzJoiningStatusForm_displayRejectionMessage(ByVal self As WIDGET, ByVal rejectionMessageFromHost As String)
    Dim d As WzJoiningStatusForm
    Set d = WzJoiningStatusForm_(self)
    If Not d.resolvableJoinErrorWidget Is Nothing Then
        d.resolvableJoinErrorWidget.hide
    End If
    d.progressIndicator.hide
    d.titleLabel.setString "Join Attempt Rejected"
    d.statusTitle.setString "Message from Host:"
    WzJoiningStatusForm_displayDetailsParagraph self, rejectionMessageFromHost
    d.actionButton.setString "Close"
End Sub

Public Sub WzJoiningStatusForm_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    ' currently, no-op
End Sub

Public Sub WzJoiningStatusForm_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x0 As Long, y0 As Long, x1 As Long, y1 As Long
    Dim dropShadowX0 As Long, dropShadowY0 As Long, dropShadowX1 As Long, dropShadowY1 As Long
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    x1 = x0 + self.width()
    y1 = y0 + self.height()

    ' draw "drop-shadow"
    dropShadowX0 = jmax(x0 - 6, 0)
    dropShadowY0 = jmax(y0 - 6, 0)
    dropShadowX1 = jmin(x1 + 6, pie_GetVideoBufferWidth())
    dropShadowY1 = jmin(y1 + 6, pie_GetVideoBufferHeight())
    pie_UniTransBoxFill CSng(dropShadowX0), CSng(dropShadowY0), CSng(dropShadowX1), CSng(dropShadowY1), pal_RGBA(0, 0, 0, 40)

    pie_UniTransBoxFill x0, y0, x1, y1, pal_RGBA(255, 255, 255, 50)
    pie_UniTransBoxFill x0, y0, x1, y1, pal_RGBA(0, 0, 0, 70)

    pie_UniTransBoxFill x0, y0, x1, y1, psPalette(WZCOL_NOTIFICATION_BOX)
    iV_Box2 x0, y0, x1, y1, psPalette(WZCOL_FORM_DARK), psPalette(WZCOL_FORM_DARK)
    iV_Box2 x0 - 1, y0 - 1, x1 + 1, y1 + 1, pal_RGBA(255, 255, 255, 50), pal_RGBA(255, 255, 255, 50)
End Sub

' MARK: - WzJoiningGameScreen_HandlerRoot

' FailureDetails FailureDetails::makeFromNetLobbyError(optional<netlobby::IPVersion> ipVer, const netlobby::LobbyError& error)
Private Function FailureDetails_makeFromNetLobbyError(ByVal has_ipVer As Boolean, ByVal ipVer As netlobby_IPVersion, error_ As netlobby_LobbyError) As JoiningFailureDetails
    Dim has_localized As Boolean, localizedLobbyErrorDescription As String, result As JoiningFailureDetails
    has_localized = netlobby_GetLocalizedLobbyErrorMessage(error_, localizedLobbyErrorDescription)

    If has_ipVer Then
        result.context = netlobby_to_string(ipVer)
    End If
    result.resultMessage = error_.errCode & ": "
    If has_localized Then
        result.resultMessage = result.resultMessage & localizedLobbyErrorDescription
    Else
        result.resultMessage = result.resultMessage & error_.errMessage
    End If

    result.source = FailureDetails_Source_NetLobby
    FailureDetails_makeFromNetLobbyError = result
End Function

' FailureDetails FailureDetails::makeFromLobbyError(LOBBY_ERROR_TYPES resultCode)
Private Function FailureDetails_makeFromLobbyError(ByVal resultCode As Long) As JoiningFailureDetails
    Dim txt As String, result As JoiningFailureDetails
    txt = ""                                ' const char* txt = nullptr (a WzString of it: empty)
    Select Case resultCode
    Case ERROR_NOERROR
    Case ERROR_FULL
        txt = "Game is full"
    Case ERROR_KICKED
        txt = "You were kicked!"
    Case ERROR_WRONGVERSION
        txt = "Your game version does not match the host"
    Case ERROR_WRONGDATA
        txt = "The host rejected your connection due to invalid data"
    Case ERROR_HOSTDROPPED
        txt = "Host has dropped connection!"
    ' error codes not currently used by the host on join:
    Case ERROR_UNKNOWNFILEISSUE, ERROR_WRONGPASSWORD
        ASSERT_FAILURE "false", "src/screens/joiningscreen.cpp:963", "makeFromLobbyError", "Unexpected lobby error from host join: %d", resultCode
        ' fall-through
        txt = "Connection Error"
    Case Else                               ' ERROR_CONNECTION, default
        txt = "Connection Error"
    End Select
    result.resultMessage = txt
    result.source = FailureDetails_Source_Other
    FailureDetails_makeFromLobbyError = result
End Function

Private Function FailureDetails_makeFromInternalError(ByVal details As String) As JoiningFailureDetails
    Dim result As JoiningFailureDetails
    result.resultMessage = details
    result.source = FailureDetails_Source_Other
    FailureDetails_makeFromInternalError = result
End Function

Private Function FailureDetails_makeFromHostRejectionMessage(ByVal message As String) As JoiningFailureDetails
    Dim result As JoiningFailureDetails
    result.resultMessage = message
    result.source = FailureDetails_Source_Host
    FailureDetails_makeFromHostRejectionMessage = result
End Function

' FailureDetails makeFromMultipleFailureDetails(Source source, const std::vector<FailureDetails>& details)
Private Function FailureDetails_makeFromMultipleFailureDetails(ByVal source As FailureDetails_Source, details() As JoiningFailureDetails, ByVal detailsCount As Long) As JoiningFailureDetails
    Dim result As JoiningFailureDetails, filteredDetails() As JoiningFailureDetails, filteredCount As Long
    Dim i As Long, j As Long, seenIdx As Long
    If detailsCount = 1 Then
        result = details(0)
        result.source = source
        FailureDetails_makeFromMultipleFailureDetails = result
        Exit Function
    End If

    ' filter duplicates (ignoring context)
    ' (std::unordered_map<WzString, size_t> seenResultMessageIdxMap: a search of the kept ones, case-sensitive)
    filteredCount = 0
    For i = 0 To detailsCount - 1
        seenIdx = -1
        For j = 0 To filteredCount - 1
            If StrComp(filteredDetails(j).resultMessage, details(i).resultMessage, vbBinaryCompare) = 0 Then
                seenIdx = j
                Exit For
            End If
        Next
        If seenIdx <> -1 Then
            ' already seen - clear the context of the first instance, and skip adding this one
            filteredDetails(seenIdx).context = ""
        Else
            ReDim Preserve filteredDetails(0 To filteredCount)
            filteredDetails(filteredCount) = details(i)
            filteredCount = filteredCount + 1
        End If
    Next

    For i = 0 To filteredCount - 1
        If Len(result.resultMessage) <> 0 Then
            result.resultMessage = result.resultMessage & vbLf
        End If
        result.resultMessage = result.resultMessage & FailureDetails_getResultMessage(filteredDetails(i))
    Next
    result.source = source
    FailureDetails_makeFromMultipleFailureDetails = result
End Function

Private Function FailureDetails_getResultMessage(fd As JoiningFailureDetails) As String
    If Len(fd.context) = 0 Then
        FailureDetails_getResultMessage = fd.resultMessage
        Exit Function
    End If
    FailureDetails_getResultMessage = "[" & fd.context & "] " & fd.resultMessage
End Function

' WzJoiningGameScreen_HandlerRoot(const W_FORMINIT *init) : W_CLICKFORM(init)
Public Sub WzJoiningGameScreen_HandlerRoot_construct_init(ByVal self As WIDGET, init As W_FORMINIT)
    Dim d As WzJoiningGameScreen_HandlerRoot
    ' (VB) WIDGET's constructor calls init's calcLayout, whose setGeometry calls geometryChanged: in C++ WIDGET's
    ' (the derived class is not constructed yet), here the kind's unless it is WIDGET's while WIDGET constructs
    self.kind = K_WIDGET
    W_CLICKFORM_construct_init self, init
    self.kind = K_WzJoiningGameScreen_HandlerRoot ' (VB) virtual calls from here on reach its overrides
    Set d = WzJoiningGameScreen_HandlerRoot_(self)
    d.backgroundColor = pal_RGBA(0, 0, 0, 80)
    d.currentJoiningState = JoiningState_AwaitingConnection
    d.weakSerial = WeakWidget_register(self)
End Sub

' ~WzJoiningGameScreen_HandlerRoot()
Public Sub WzJoiningGameScreen_HandlerRoot_destruct(ByVal self As WIDGET)
    ' close any in-progress connection attempt
    WzJoiningGameScreen_HandlerRoot_closeConnectionAttempt self
    WzJoiningGameScreen_HandlerRoot_(self).currentJoiningState = JoiningState_Failure
    WeakWidget_unregister WzJoiningGameScreen_HandlerRoot_(self).weakSerial
End Sub

' static optional<netlobby::IPVersion> knownAvailableConnectionTypesToIPVersion(const std::vector<netlobby::ConnectionType>*
' pKnownAvailableConnectionTypes) (has_known: the pointer is not nullptr; the result's has_value is the return)
Private Function knownAvailableConnectionTypesToIPVersion(ByVal has_known As Boolean, known() As netlobby_ConnectionType, ByVal knownCount As Long, out As netlobby_IPVersion) As Boolean
    Dim hasIPv4 As Boolean, hasIPv6 As Boolean, hasIPUnspecified As Boolean, i As Long
    If Not has_known Then
        If war_getLobbyDisableIPv6() Then
            ' IPv6 disabled, so just return IPv4
            out = netlobby_IPVersion_IPv4
            knownAvailableConnectionTypesToIPVersion = True
            Exit Function
        End If
        knownAvailableConnectionTypesToIPVersion = False
        Exit Function
    End If

    hasIPv4 = False
    hasIPv6 = False
    hasIPUnspecified = False
    For i = 0 To knownCount - 1
        If Not known(i).has_ipVersion Then
            hasIPUnspecified = True
        Else
            Select Case known(i).ipVersion
            Case netlobby_IPVersion_IPv4
                hasIPv4 = True
            Case netlobby_IPVersion_IPv6
                hasIPv6 = True
            End Select
        End If
    Next

    If hasIPUnspecified Then
        If war_getLobbyDisableIPv6() Then
            ' IPv6 disabled, so just return IPv4
            out = netlobby_IPVersion_IPv4
            knownAvailableConnectionTypesToIPVersion = True
            Exit Function
        End If
        knownAvailableConnectionTypesToIPVersion = False
        Exit Function
    End If
    If hasIPv4 And Not hasIPv6 Then
        out = netlobby_IPVersion_IPv4
        knownAvailableConnectionTypesToIPVersion = True
        Exit Function
    End If
    If (Not hasIPv4) And hasIPv6 Then
        out = netlobby_IPVersion_IPv6
        knownAvailableConnectionTypesToIPVersion = True
        Exit Function
    End If

    ' in case both IPv4 and IPv6 are available...
    If war_getLobbyDisableIPv6() Then
        ' IPv6 disabled, so just return IPv4
        out = netlobby_IPVersion_IPv4
        knownAvailableConnectionTypesToIPVersion = True
        Exit Function
    End If
    knownAvailableConnectionTypesToIPVersion = False
End Function

' static makeWithLobbyGame(...): the pKnownAvailableConnectionTypes pointer is has_known / known(0 .. knownCount - 1)
Private Function WzJoiningGameScreen_HandlerRoot_makeWithLobbyGame(ByVal lobbyAddress As String, ByVal gameId As String, _
        ByVal playerName As String, playerIdentity As EcKey, ByVal asSpectator As Boolean, expectedHostProps As ExpectedHostProperties, _
        ByVal has_known As Boolean, known() As netlobby_ConnectionType, ByVal knownCount As Long, _
        ByVal onSuccessFunc As WzFunc, ByVal onFailureFunc As WzFunc) As WIDGET
    Dim widget As WIDGET, has_ipVer As Boolean, ipVer As netlobby_IPVersion
    Set widget = WzJoiningGameScreen_HandlerRoot_makeInternal(playerName, playerIdentity, asSpectator, expectedHostProps, onSuccessFunc, onFailureFunc)
    has_ipVer = knownAvailableConnectionTypesToIPVersion(has_known, known, knownCount, ipVer)
    WzJoiningGameScreen_HandlerRoot_requestLobbyJoin widget, lobbyAddress, gameId, has_ipVer, ipVer

    Set WzJoiningGameScreen_HandlerRoot_makeWithLobbyGame = widget
End Function

' void requestLobbyJoin(std::string lobbyAddress, std::string gameId, optional<netlobby::IPVersion> specificIpVersion = nullopt,
' std::shared_ptr<netlobby::LobbyErrorResolutionData> lobbyErrorResolutionData = nullptr)
Private Sub WzJoiningGameScreen_HandlerRoot_requestLobbyJoin(ByVal self As WIDGET, ByVal lobbyAddress As String, ByVal gameId As String, _
        ByVal has_specificIpVersion As Boolean, ByVal specificIpVersion As netlobby_IPVersion, Optional ByVal lobbyErrorResolutionData As Object = Nothing)
    Dim d As WzJoiningGameScreen_HandlerRoot, debugStr As String, weakSelf As Long, playerIdentity As EcKey
    Set d = WzJoiningGameScreen_HandlerRoot_(self)
    d.currentJoiningState = JoiningState_FetchingLobbyJoinInfo

    ' Kick off fetching lobby join info & update progress display
    WzJoiningGameScreen_HandlerRoot_updateJoiningStatus self, "Gathering connection details"

    debugStr = "Gathering connection details (lobby gameId: " & gameId & ")"
    If has_specificIpVersion Then
        debugStr = debugStr & " (" & netlobby_to_string(specificIpVersion) & ")"
    Else
        debugStr = debugStr & " (ip4 + ip6)"
    End If
    debug_ 1144, LOG_INFO, "requestLobbyJoin", "%s", debugStr

    weakSelf = d.weakSerial
    playerIdentity = d.playerIdentity
    netlobby_RequestJoinDetails lobbyAddress, gameId, d.playerName, playerIdentity, d.asSpectator, _
        WzFunc_new(FN_joiningHandlerRoot_lobbyJoinResult, weakSelf, lobbyAddress, gameId, has_specificIpVersion, specificIpVersion), _
        has_specificIpVersion, specificIpVersion, lobbyErrorResolutionData
End Sub

' static const char* to_display_string(netlobby::RequestJoinDetailsResults::ConnectivityFailureType failureType)
Private Function connectivityFailureTypeToDisplayString(ByVal failureType As netlobby_ConnectivityFailureType) As String
    Select Case failureType
    Case netlobby_ConnectivityFailureType_ERR_COULD_NOT_CONNECT
        connectivityFailureTypeToDisplayString = "Could Not Connect"
    Case netlobby_ConnectivityFailureType_ERR_TIMEOUT
        connectivityFailureTypeToDisplayString = "Timeout"
    Case netlobby_ConnectivityFailureType_ERR_OTHER
        connectivityFailureTypeToDisplayString = "Error"
    Case Else
        connectivityFailureTypeToDisplayString = "Unknown error"
    End Select
End Function

' static optional<JoinConnectionDescription::JoinConnectionType> lobby_connection_method_to_join_connection_type(netlobby::ConnectionMethod method)
Private Function lobby_connection_method_to_join_connection_type(ByVal method As netlobby_ConnectionMethod, out As JoinConnectionType) As Boolean
    Select Case method
    Case netlobby_ConnectionMethod_TCP_DIRECT
        out = JoinConnectionType_TCP_DIRECT
        lobby_connection_method_to_join_connection_type = True
    Case netlobby_ConnectionMethod_GNS_DIRECT
        ' (WZ_GNS_NETWORK_BACKEND_ENABLED: not defined)
        lobby_connection_method_to_join_connection_type = False
    Case Else
        lobby_connection_method_to_join_connection_type = False   ' silence compiler warning
    End Select
End Function

' static optional<JoinConnectionDescription> ConvertToJoinConnectionDescription(const netlobby::ConnectionInfo& conn)
Private Function ConvertToJoinConnectionDescription(conn As netlobby_ConnectionInfo, out As JoinConnectionDescription) As Boolean
    Dim t As JoinConnectionType
    If Not (Len(conn.host) <> 0) Then ASSERT_FAILURE "!conn.host.empty()", "src/screens/joiningscreen.cpp:1192", "ConvertToJoinConnectionDescription", "Missing host": Exit Function

    If Not lobby_connection_method_to_join_connection_type(conn.typ.method, t) Then
        ' Unsupported connection type
        ConvertToJoinConnectionDescription = False
        Exit Function
    End If
    out = JoinConnectionDescription_make(t, conn.host, conn.port)
    ConvertToJoinConnectionDescription = True
End Function

' std::vector<InternalConnectionInfo> convertLobbyGameDetailsToInternalConnectionList(const netlobby::GameJoinDetails& joinDetails)
' (the result's count)
Private Function convertLobbyGameDetailsToInternalConnectionList(joinDetails As netlobby_GameJoinDetails, result() As InternalConnectionInfo) As Long
    Dim i As Long, n As Long, converted As JoinConnectionDescription
    Erase result
    If Not (joinDetails.connectionsCount = joinDetails.joinTokensCount) Then ASSERT_FAILURE "joinDetails.connections.size() == joinDetails.joinTokens.size()", "src/screens/joiningscreen.cpp:1205", "convertLobbyGameDetailsToInternalConnectionList", "Not the same number of connections (%zu) and join info (%zu)", joinDetails.connectionsCount, joinDetails.joinTokensCount: Exit Function

    n = 0
    For i = 0 To joinDetails.connectionsCount - 1
        If ConvertToJoinConnectionDescription(joinDetails.connections(i), converted) Then
            ReDim Preserve result(0 To n)
            result(n).conn = converted
            result(n).joinToken = joinDetails.joinTokens(i)
            n = n + 1
        End If
    Next

    convertLobbyGameDetailsToInternalConnectionList = n
End Function

' std::vector<InternalConnectionInfo> convertConnectionDescriptionsToInternalConnectionList(const std::vector<JoinConnectionDescription>& newConnectionList)
Private Function convertConnectionDescriptionsToInternalConnectionList(newConnectionList() As JoinConnectionDescription, ByVal newConnectionListCount As Long, result() As InternalConnectionInfo) As Long
    Dim i As Long
    Erase result
    For i = 0 To newConnectionListCount - 1
        ReDim Preserve result(0 To i)
        result(i).conn = newConnectionList(i)
        result(i).joinToken = ""
    Next

    convertConnectionDescriptionsToInternalConnectionList = newConnectionListCount
End Function

' (VB) the std::stable_sort comparator of setConnectionList
Private Function connectionListLess(a As InternalConnectionInfo, b As InternalConnectionInfo, ByVal bSortIPv6First As Boolean) As Boolean
    Dim a_isIPv6 As Boolean, b_isIPv6 As Boolean
    a_isIPv6 = InStr(a.conn.host, ":") <> 0     ' this is a very simplistic test - if the host contains ":" we treat it as IPv6
    b_isIPv6 = InStr(b.conn.host, ":") <> 0
    If bSortIPv6First Then
        connectionListLess = a_isIPv6 And Not b_isIPv6
    Else
        connectionListLess = (Not a_isIPv6) And b_isIPv6
    End If
End Function

Private Sub WzJoiningGameScreen_HandlerRoot_setConnectionList(ByVal self As WIDGET, newConnectionList() As InternalConnectionInfo, ByVal newConnectionListCount As Long)
    Dim bSortIPv6First As Boolean, i As Long, j As Long, key As InternalConnectionInfo
    ' sort the connection list by user prefs
    If newConnectionListCount > 1 Then
        ' sort the list, based on NETgetJoinPreferenceIPv6
        ' preserve the original relative order amongst each class of IPv4/IPv6 addresses
        bSortIPv6First = NETgetJoinPreferenceIPv6()
        ' std::stable_sort (an insertion sort: stable)
        For i = 1 To newConnectionListCount - 1
            key = newConnectionList(i)
            j = i - 1
            Do While j >= 0
                If Not connectionListLess(key, newConnectionList(j), bSortIPv6First) Then Exit Do
                newConnectionList(j + 1) = newConnectionList(j)
                j = j - 1
            Loop
            newConnectionList(j + 1) = key
        Next
    End If

    WzJoiningGameScreen_HandlerRoot_(self).setConnectionList newConnectionList, newConnectionListCount
End Sub

Private Sub WzJoiningGameScreen_HandlerRoot_handleLobbyErrorResolve(ByVal self As WIDGET, ByVal lobbyErrorResolveWidget As WIDGET)
    Dim d As WzJoiningGameScreen_HandlerRoot
    Set d = WzJoiningGameScreen_HandlerRoot_(self)
    If Not (Not lobbyErrorResolveWidget Is Nothing) Then ASSERT_FAILURE "lobbyErrorResolveWidget != nullptr", "src/screens/joiningscreen.cpp:1252", "handleLobbyErrorResolve", "Null lobbyErrorResolveWidget"

    ' Transition to NeedsLobbyErrorResolve state
    d.currentJoiningState = JoiningState_NeedsLobbyErrorResolve

    ' Display the widget
    WzJoiningGameScreen_HandlerRoot_updateJoiningStatus self, "The host has requested additional verification"
    WzJoiningStatusForm_displayResolvableJoinErrorWidget d.joiningProgressForm, lobbyErrorResolveWidget
    d.joiningProgressForm.callCalcLayout
End Sub

Private Sub WzJoiningGameScreen_HandlerRoot_handleLobbyRequestJoinResult(ByVal self As WIDGET, ByVal lobbyAddress As String, ByVal gameId As String, _
        ByVal has_specificIpVersion As Boolean, ByVal specificIpVersion As netlobby_IPVersion, result As netlobby_JoinResult)
    Dim results As netlobby_RequestJoinDetailsResults, convertedJoinDetails() As InternalConnectionInfo, convertedCount As Long
    Dim failureDetails() As JoiningFailureDetails, failureCount As Long, hasResolvableLobbyError As Boolean, i As Long, k As Long
    Dim errStr As String, weakSelf As Long, lobbyErrorResolveWidget As WIDGET, hasIPv4Error As Boolean
    If Not result.has_value Then
        WzJoiningGameScreen_HandlerRoot_handleFailure self, FailureDetails_makeFromNetLobbyError(False, netlobby_IPVersion_IPv4, result.error_)
        Exit Sub
    End If

    results = result.value

    convertedCount = convertLobbyGameDetailsToInternalConnectionList(results.joinDetails, convertedJoinDetails)

    failureCount = 0
    hasResolvableLobbyError = False
    For i = 0 To results.errorsCount - 1
        If Not results.errors(i).isLobbyError Then
            ' [&](const netlobby::RequestJoinDetailsResults::ConnectivityFailure& arg)
            ' TODO: Handle connectivityErrors (if no IPv6, but there *is* IPv4, for example, can avoid trying to request IPv6 for future joins this session?)
            If results.errors(i).first = netlobby_IPVersion_IPv6 Then
                hasIPv4Error = False
                For k = 0 To results.errorsCount - 1
                    If results.errors(k).first = netlobby_IPVersion_IPv4 Then hasIPv4Error = True
                Next
                If (Not hasIPv4Error) And results.joinDetails.connectionsCount <> 0 Then
                    ' There was an IPv6 connectivity issue, but IPv4 was fine
                    ' TODO: Suppress future lookups of IPv6 connectivity to avoid future delays
                End If
            End If
            errStr = "[" & netlobby_to_string(results.errors(i).first) & "] " & connectivityFailureTypeToDisplayString(results.errors(i).connectivityType) & ": " & results.errors(i).connectivityDetails
            ReDim Preserve failureDetails(0 To failureCount)
            failureDetails(failureCount) = FailureDetails_makeFromInternalError(errStr)
            failureCount = failureCount + 1
        Else
            ' [&](const netlobby::LobbyError& arg)
            debug_ 1294, LOG_NET, "handleLobbyRequestJoinResult", "%s: %s", results.errors(i).lobbyError.errCode, results.errors(i).lobbyError.errMessage
            ' Check for resolvable lobby errors
            hasResolvableLobbyError = hasResolvableLobbyError Or netlobby_LobbyError_isResolvableLobbyError(results.errors(i).lobbyError)
            ReDim Preserve failureDetails(0 To failureCount)
            failureDetails(failureCount) = FailureDetails_makeFromNetLobbyError(True, results.errors(i).first, results.errors(i).lobbyError)
            failureCount = failureCount + 1
        End If
    Next

    If hasResolvableLobbyError Then
        ' If there are *any* resolvable lobby error(s), prefer resolving that / those before attempting to join
        ' (since they might be preventing full retrieval of join details)

        ' Find first resolvable lobby error for which a resolve widget can be generated
        weakSelf = WzJoiningGameScreen_HandlerRoot_(self).weakSerial
        For i = 0 To results.errorsCount - 1
            If results.errors(i).isLobbyError Then          ' std::get<netlobby::LobbyError> (else std::bad_variant_access: continue)
                If lobbyErrorResolveWidget Is Nothing Then
                    Set lobbyErrorResolveWidget = netlobby_GetLobbyErrorResolveWidget(results.errors(i).lobbyError, _
                        WzFunc_new(FN_joiningHandlerRoot_lobbyErrorResolved, weakSelf, lobbyAddress, gameId, has_specificIpVersion, specificIpVersion))
                End If
            End If
        Next

        If Not lobbyErrorResolveWidget Is Nothing Then
            WzJoiningGameScreen_HandlerRoot_handleLobbyErrorResolve self, lobbyErrorResolveWidget
            Exit Sub
        Else
            debug_ 1339, LOG_LOBBY, "handleLobbyRequestJoinResult", "Have resolvable lobby error, but unable to create widget to resolve it - skipping"
        End If
    End If

    If convertedCount = 0 Then
        ' If no connections returned, presumably there were errors?
        If results.errorsCount <> 0 Then
            ' No connections, but there were error(s) - just display those
            WzJoiningGameScreen_HandlerRoot_handleFailure self, FailureDetails_makeFromMultipleFailureDetails(FailureDetails_Source_NetLobby, failureDetails, failureCount)
        Else
            ' No connections, but no errors fetching connections
            '
            ' Game may have disappeared from lobby, or stopped hosting,
            ' or this copy of WZ might not be compiled with support for all connection types.
            '
            ' In any case, treat as "Can't connect" state
            WzJoiningGameScreen_HandlerRoot_handleFailure self, FailureDetails_makeFromInternalError("No supported connection details")
        End If
        Exit Sub
    End If

    WzJoiningGameScreen_HandlerRoot_setConnectionList self, convertedJoinDetails, convertedCount
    WzJoiningGameScreen_HandlerRoot_startAttemptingConnections self
End Sub

' static make(const std::vector<JoinConnectionDescription>& connectionList, ...)
Private Function WzJoiningGameScreen_HandlerRoot_make(connectionList() As JoinConnectionDescription, ByVal connectionListCount As Long, _
        ByVal playerName As String, playerIdentity As EcKey, ByVal asSpectator As Boolean, expectedHostProps As ExpectedHostProperties, _
        ByVal onSuccessFunc As WzFunc, ByVal onFailureFunc As WzFunc) As WIDGET
    Dim widget As WIDGET, list() As InternalConnectionInfo, n As Long
    Set widget = WzJoiningGameScreen_HandlerRoot_makeInternal(playerName, playerIdentity, asSpectator, expectedHostProps, onSuccessFunc, onFailureFunc)
    n = convertConnectionDescriptionsToInternalConnectionList(connectionList, connectionListCount, list)
    WzJoiningGameScreen_HandlerRoot_setConnectionList widget, list, n

    WzJoiningGameScreen_HandlerRoot_startAttemptingConnections widget
    Set WzJoiningGameScreen_HandlerRoot_make = widget
End Function

Private Function WzJoiningGameScreen_HandlerRoot_makeInternal(ByVal playerName As String, playerIdentity As EcKey, ByVal asSpectator As Boolean, _
        expectedHostProps As ExpectedHostProperties, ByVal onSuccessFunc As WzFunc, ByVal onFailureFunc As WzFunc) As WIDGET
    Dim sInit As W_FORMINIT, widget As WIDGET, d As WzJoiningGameScreen_HandlerRoot
    W_FORMINIT_construct sInit
    sInit.id = 0
    sInit.style = WFORM_PLAIN Or WFORM_CLICKABLE
    sInit.x = 0
    sInit.y = 0
    sInit.width = screenWidth - 1
    sInit.height = screenHeight - 1
    Set sInit.calcLayout = WzFunc_new(FN_joiningHandlerRoot_formInitCalcLayout)

    ' (class make_shared_enabler)
    Set widget = WIDGET_alloc(K_WzJoiningGameScreen_HandlerRoot)
    WzJoiningGameScreen_HandlerRoot_construct_init widget, sInit
    Set d = WzJoiningGameScreen_HandlerRoot_(widget)
    d.playerName = playerName
    d.playerIdentity = playerIdentity
    d.asSpectator = asSpectator
    d.expectedHostProps = expectedHostProps
    Set d.onSuccessFunc = onSuccessFunc
    Set d.onFailureFunc = onFailureFunc

    WzJoiningGameScreen_HandlerRoot_initialize widget
    Set WzJoiningGameScreen_HandlerRoot_makeInternal = widget
End Function

Private Sub WzJoiningGameScreen_HandlerRoot_applyExpectedHostProps(ByVal self As WIDGET)
    Dim d As WzJoiningGameScreen_HandlerRoot, ehp As ExpectedHostProperties
    Set d = WzJoiningGameScreen_HandlerRoot_(self)
    ehp = d.expectedHostProps
    If ehp.has_gamePassword Then
        d.gamePassword = sstrcpyTo(ehp.gamePassword, password_string_size)
    End If
End Sub

Private Sub WzJoiningGameScreen_HandlerRoot_initialize(ByVal self As WIDGET)
    Dim d As WzJoiningGameScreen_HandlerRoot
    Set d = WzJoiningGameScreen_HandlerRoot_(self)
    d.timeStarted = widgSteadyNowMs()

    WzJoiningGameScreen_HandlerRoot_applyExpectedHostProps self

    ' (auto weakSelf = ...: unused here)

    ' Create progress display form
    Set d.joiningProgressForm = WzJoiningStatusForm_make()
    self.attach d.joiningProgressForm
    d.joiningProgressForm.setCalcLayout WzFunc_new(FN_joiningHandlerRoot_progressFormCalcLayout)

    WzJoiningGameScreen_HandlerRoot_recalcLayout self
End Sub

Private Sub WzJoiningGameScreen_HandlerRoot_startAttemptingConnections(ByVal self As WIDGET)
    Dim d As WzJoiningGameScreen_HandlerRoot
    Set d = WzJoiningGameScreen_HandlerRoot_(self)
    ' Kick-off the initial connection attempt
    If d.connectionListCount = 0 Then
        debug_ 1445, LOG_ERROR, "startAttemptingConnections", "No connections to try"
        WzJoiningGameScreen_HandlerRoot_handleFailure self, FailureDetails_makeFromInternalError("No connections available")
        Exit Sub
    End If

    If Not (d.currentJoiningState = JoiningState_AwaitingConnection Or d.currentJoiningState = JoiningState_NeedsPassword Or d.currentJoiningState = JoiningState_FetchingLobbyJoinInfo) Then ASSERT_FAILURE "currentJoiningState == JoiningState::AwaitingConnection || currentJoiningState == JoiningState::NeedsPassword || currentJoiningState == JoiningState::FetchingLobbyJoinInfo", "src/screens/joiningscreen.cpp:1450", "startAttemptingConnections", "Unexpected joining state: %d", d.currentJoiningState: Exit Sub
    d.currentJoiningState = JoiningState_AwaitingConnection

    WzJoiningGameScreen_HandlerRoot_attemptToOpenConnection self, 0
End Sub

Private Sub WzJoiningGameScreen_HandlerRoot_handleSuccess(ByVal self As WIDGET)
    Dim d As WzJoiningGameScreen_HandlerRoot, connDesc As InternalConnectionInfo
    Set d = WzJoiningGameScreen_HandlerRoot_(self)
    d.currentJoiningState = JoiningState_SuccessPendingClose
    WzJoiningStatusForm_displaySuccessStatus d.joiningProgressForm, "Synchronizing data with host ..."
    If Not (Not d.onSuccessFunc Is Nothing) Then ASSERT_FAILURE "onSuccessFunc != nullptr", "src/screens/joiningscreen.cpp:1460", "handleSuccess", "Missing success handler!": Exit Sub
    connDesc = d.connection(d.currentConnectionIdx)
    WzFunc_call d.onSuccessFunc, connDesc.conn.host, connDesc.conn.port, connDesc.conn.type_
End Sub

Private Sub WzJoiningGameScreen_HandlerRoot_promptForPassword(ByVal self As WIDGET)
    Dim d As WzJoiningGameScreen_HandlerRoot, weakSelf As Long
    Set d = WzJoiningGameScreen_HandlerRoot_(self)
    d.currentJoiningState = JoiningState_NeedsPassword

    weakSelf = d.weakSerial
    WzJoiningStatusForm_displayPasswordPrompt d.joiningProgressForm, WzFunc_new(FN_joiningHandlerRoot_passwordSubmitted, weakSelf)
    d.joiningProgressForm.callCalcLayout
End Sub

Private Sub WzJoiningGameScreen_HandlerRoot_updateJoiningStatus(ByVal self As WIDGET, ByVal statusDescription As String)
    debug_ 1485, LOG_NET, "updateJoiningStatus", "%s", statusDescription
    WzJoiningStatusForm_displayProgressStatus WzJoiningGameScreen_HandlerRoot_(self).joiningProgressForm, statusDescription
End Sub

Private Function joiningStateToString(ByVal s As JoiningState) As String
    Select Case s
    Case JoiningState_NeedsLobbyErrorResolve: joiningStateToString = "NeedsLobbyErrorResolve"
    Case JoiningState_FetchingLobbyJoinInfo: joiningStateToString = "FetchingJoinInfo"
    Case JoiningState_NeedsPassword: joiningStateToString = "NeedsPassword"
    Case JoiningState_AwaitingConnection: joiningStateToString = "AwaitingConnection"
    Case JoiningState_AwaitingInitialNetcodeHandshakeAck: joiningStateToString = "AwaitingInitialNetcodeHandshakeAck"
    Case JoiningState_ProcessingJoinMessages: joiningStateToString = "ProcessingJoinMessages"
    Case JoiningState_Failure: joiningStateToString = "Failure"
    Case JoiningState_SuccessPendingClose: joiningStateToString = "SuccessPendingClose"
    Case JoiningState_Success: joiningStateToString = "Success"
    Case Else: joiningStateToString = ""    ' silence compiler warning
    End Select
End Function

Private Function joiningStateToDisplayStr(ByVal s As JoiningState) As String
    Select Case s
    Case JoiningState_NeedsLobbyErrorResolve: joiningStateToDisplayStr = "NeedsLobbyErrorResolve"
    Case JoiningState_FetchingLobbyJoinInfo: joiningStateToDisplayStr = "FetchingJoinInfo"
    Case JoiningState_NeedsPassword: joiningStateToDisplayStr = "NeedsPassword"
    Case JoiningState_AwaitingConnection: joiningStateToDisplayStr = "PendingConnect"
    Case JoiningState_AwaitingInitialNetcodeHandshakeAck: joiningStateToDisplayStr = "NetcodeHandshake"
    Case JoiningState_ProcessingJoinMessages: joiningStateToDisplayStr = "ProcessingJoin"
    Case JoiningState_Failure: joiningStateToDisplayStr = "Failure"
    Case JoiningState_SuccessPendingClose: joiningStateToDisplayStr = "SuccessPendingClose"
    Case JoiningState_Success: joiningStateToDisplayStr = "Success"
    Case Else: joiningStateToDisplayStr = ""    ' silence compiler warning
    End Select
End Function

Private Function joiningStateToLocalizedStateFailDesc(ByVal s As JoiningState) As String
    Select Case s
    Case JoiningState_NeedsLobbyErrorResolve
        joiningStateToLocalizedStateFailDesc = "Waiting for lobby error resolution"
    Case JoiningState_FetchingLobbyJoinInfo
        joiningStateToLocalizedStateFailDesc = "Waiting for join info"
    Case JoiningState_NeedsPassword
        joiningStateToLocalizedStateFailDesc = "Waiting for correct join password"
    Case JoiningState_AwaitingConnection
        joiningStateToLocalizedStateFailDesc = "Attempting to connect"
    Case JoiningState_AwaitingInitialNetcodeHandshakeAck
        joiningStateToLocalizedStateFailDesc = "Establishing connection handshake"
    Case JoiningState_ProcessingJoinMessages
        joiningStateToLocalizedStateFailDesc = "Coordinating join with host"
    Case JoiningState_Failure
        joiningStateToLocalizedStateFailDesc = "Join attempt failed"
    Case JoiningState_SuccessPendingClose, JoiningState_Success
        joiningStateToLocalizedStateFailDesc = ""
    Case Else
        joiningStateToLocalizedStateFailDesc = ""   ' silence compiler warning
    End Select
End Function

Private Sub WzJoiningGameScreen_HandlerRoot_handleJoinTimeoutError(ByVal self As WIDGET)
    Dim d As WzJoiningGameScreen_HandlerRoot, timeoutErrorDetails As String, localizedJoinStateDesc As String
    Set d = WzJoiningGameScreen_HandlerRoot_(self)
    debug_ 1547, LOG_INFO, "handleJoinTimeoutError", "Failed to join with timeout, state: %s", joiningStateToString(d.currentJoiningState)

    timeoutErrorDetails = "Host did not respond before timeout"
    timeoutErrorDetails = timeoutErrorDetails & vbLf
    localizedJoinStateDesc = joiningStateToLocalizedStateFailDesc(d.currentJoiningState)
    If Len(localizedJoinStateDesc) <> 0 Then
        timeoutErrorDetails = timeoutErrorDetails & "Failed at: [" & joiningStateToDisplayStr(d.currentJoiningState) & "] - " & localizedJoinStateDesc
    Else
        timeoutErrorDetails = timeoutErrorDetails & "Failed at: [" & joiningStateToDisplayStr(d.currentJoiningState) & "]"
    End If

    d.currentJoiningState = JoiningState_Failure

    WzJoiningStatusForm_displayUnableToJoinError d.joiningProgressForm, timeoutErrorDetails
    d.joiningProgressForm.callCalcLayout

    If Not d.onFailureFunc Is Nothing Then
        WzFunc_call d.onFailureFunc
    End If

    ActivityManager_joinGameFailed
End Sub

Private Sub WzJoiningGameScreen_HandlerRoot_handleFailure(ByVal self As WIDGET, failureDetails As JoiningFailureDetails)
    Dim d As WzJoiningGameScreen_HandlerRoot, failureResultMessage As String
    Set d = WzJoiningGameScreen_HandlerRoot_(self)
    d.currentJoiningState = JoiningState_Failure

    failureResultMessage = FailureDetails_getResultMessage(failureDetails)
    debug_ 1579, LOG_INFO, "handleFailure", "Failed to join with error: %s", failureResultMessage

    Select Case failureDetails.source
    Case FailureDetails_Source_NetLobby
        WzJoiningStatusForm_displayNetLobbyError d.joiningProgressForm, failureResultMessage
    Case FailureDetails_Source_Host
        WzJoiningStatusForm_displayRejectionMessage d.joiningProgressForm, failureResultMessage
    Case Else
        WzJoiningStatusForm_displayUnableToJoinError d.joiningProgressForm, failureResultMessage
    End Select

    d.joiningProgressForm.callCalcLayout

    If Not d.onFailureFunc Is Nothing Then
        WzFunc_call d.onFailureFunc
    End If

    ActivityManager_joinGameFailed
End Sub

Private Sub WzJoiningGameScreen_HandlerRoot_recalcLayout(ByVal self As WIDGET)
    WzJoiningGameScreen_HandlerRoot_(self).joiningProgressForm.callCalcLayout
End Sub

Public Sub WzJoiningGameScreen_HandlerRoot_geometryChanged(ByVal self As WIDGET)
    WzJoiningGameScreen_HandlerRoot_recalcLayout self
End Sub

Public Sub WzJoiningGameScreen_HandlerRoot_clicked(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    Dim d As WzJoiningGameScreen_HandlerRoot
    Set d = WzJoiningGameScreen_HandlerRoot_(self)
    If Not d.onClickedFunc Is Nothing Then
        WzFunc_call d.onClickedFunc
    End If
End Sub

Public Sub WzJoiningGameScreen_HandlerRoot_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As WzJoiningGameScreen_HandlerRoot, x0 As Long, y0 As Long
    Set d = WzJoiningGameScreen_HandlerRoot_(self)
    If Not self.visible() Then
        Exit Sub ' skip if hidden
    End If

    If pielightIsTransparent(d.backgroundColor) Then
        Exit Sub
    End If

    x0 = self.x() + xOffset
    y0 = self.y() + yOffset

    ' draw background over everything
    pie_UniTransBoxFill x0, y0, x0 + self.width(), y0 + self.height(), d.backgroundColor
End Sub

Public Sub WzJoiningGameScreen_HandlerRoot_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    Dim d As WzJoiningGameScreen_HandlerRoot
    Set d = WzJoiningGameScreen_HandlerRoot_(self)
    If keyPressed(KEY_ESC) Then
        ' while this is displayed, clear the input buffer
        ' (ensuring that subsequent screens / the main screen do not get the input to handle)
        inputLoseFocus

        If Not d.onCancelPressed Is Nothing Then
            WzFunc_call d.onCancelPressed
        End If
        Exit Sub
    End If

    ' while this is displayed, clear the input buffer
    ' (ensuring that subsequent screens / the main screen do not get the input to handle)
    inputLoseFocus

    WzJoiningGameScreen_HandlerRoot_processJoining self
End Sub

' Process a connection result (to be called from non-main threads)
' (VB: JoiningOpenConnectionCallback calls it; the OpenConnectionResult object is the moved-in result)
Public Sub WzJoiningGameScreen_HandlerRoot_processOpenConnectionResultOnMainThread(ByVal self As WIDGET, ByVal connectionIdx As Long, ByVal result As OpenConnectionResult)
    Dim weakSelf As Long
    weakSelf = WzJoiningGameScreen_HandlerRoot_(self).weakSerial

    ' Work around the fact that we need to copy a move-only type into the lambda (which is stored as a std::function)
    wzAsyncExecOnMainThread WzFunc_new(FN_joiningHandlerRoot_resultOnMainThread, weakSelf, connectionIdx, result)
End Sub

' static ConnectionProviderType toConnectionProviderType(JoinConnectionDescription::JoinConnectionType ct)
Private Function toConnectionProviderType(ByVal ct As JoinConnectionType) As ConnectionProviderType
    Select Case ct
    Case JoinConnectionType_TCP_DIRECT
        toConnectionProviderType = ConnectionProviderType_TCP_DIRECT
        Exit Function
    ' (WZ_GNS_NETWORK_BACKEND_ENABLED: not defined)
    End Select
    Err.Raise 5, "toConnectionProviderType", "Invalid join connection type: " & ct  ' throw std::runtime_error (prevent GCC warning)
End Function

Private Sub WzJoiningGameScreen_HandlerRoot_tryNextConnectionOption(ByVal self As WIDGET, ByVal connectionIdx As Long, ByVal ec As Long, ByVal errorString As String)
    Dim d As WzJoiningGameScreen_HandlerRoot, sockErrorMsg As String, localizedError As String
    Set d = WzJoiningGameScreen_HandlerRoot_(self)
    If (connectionIdx + 1) < d.connectionListCount Then
        ' try the next connection
        WzJoiningGameScreen_HandlerRoot_closeConnectionAttempt self
        d.currentJoiningState = JoiningState_AwaitingConnection
        connectionIdx = connectionIdx + 1
        WzJoiningGameScreen_HandlerRoot_attemptToOpenConnection self, connectionIdx
    ElseIf net_errorEqualsErrc(ec, errc_timed_out) Then
        WzJoiningGameScreen_HandlerRoot_handleJoinTimeoutError self
    Else
        debug_ 1706, LOG_ERROR, "tryNextConnectionOption", "%s", errorString
        ' Done trying connections - all failed
        sockErrorMsg = net_errorMessage(ec)
        localizedError = "Failed to open connection: [" & CStr(net_errorValue(ec)) & "] " & sockErrorMsg
        WzJoiningGameScreen_HandlerRoot_handleFailure self, FailureDetails_makeFromInternalError(localizedError)
    End If
End Sub

' Process a connection result - **Only to be called on the main thread**
Private Sub WzJoiningGameScreen_HandlerRoot_processOpenConnectionResult(ByVal self As WIDGET, ByVal connectionIdx As Long, ByVal result As OpenConnectionResult)
    Dim d As WzJoiningGameScreen_HandlerRoot, buffer(0 To 7) As Byte, rawDummy As Long, ec As Long, writeErrMsg As String
    Dim connProvider As WzConnectionProvider, conn As InternalConnectionInfo
    Set d = WzJoiningGameScreen_HandlerRoot_(self)
    If Not (d.currentJoiningState = JoiningState_AwaitingConnection) Then ASSERT_FAILURE "currentJoiningState == JoiningState::AwaitingConnection", "src/screens/joiningscreen.cpp:1716", "processOpenConnectionResult", "Not awaiting connection? (Ignoring)": Exit Sub
    If Not (connectionIdx = d.currentConnectionIdx) Then ASSERT_FAILURE "connectionIdx == currentConnectionIdx", "src/screens/joiningscreen.cpp:1717", "processOpenConnectionResult", "Received result for connection %zu, but should be processing: %zu", connectionIdx, d.currentConnectionIdx: Exit Sub

    If result.hasError() Then
        WzJoiningGameScreen_HandlerRoot_tryNextConnectionOption self, connectionIdx, result.errorCode(), result.errorString()
        Exit Sub
    End If

    ' take ownership of the socket
    Set d.client_transient_socket = result.open_socket_release()

    If NETgetEnableTCPNoDelay() Then
        ' Disable use of Nagle Algorithm for the TCP socket (i.e. enable TCP_NODELAY option in case of TCP transport)
        d.client_transient_socket.useNagleAlgorithm False
    End If

    ' Send initial connection data: NETCODE_VERSION_MAJOR and NETCODE_VERSION_MINOR
    ' (pushu32: wz_htonl(value), the network byte order)
    pushu32 buffer, 0, NETGetMajorVersion()
    pushu32 buffer, 4, NETGetMinorVersion()

    d.client_transient_socket.writeAll buffer, 0, 8, rawDummy, ec
    If ec <> 0 Then
        writeErrMsg = net_errorMessage(ec)
        debug_ 1749, LOG_ERROR, "processOpenConnectionResult", "Couldn't send my version: %s", writeErrMsg
        WzJoiningGameScreen_HandlerRoot_closeConnectionAttempt self
        Exit Sub
    End If

    conn = d.connection(connectionIdx)
    Set connProvider = ConnectionProviderRegistry_Get(toConnectionProviderType(conn.conn.type_))
    Set d.tmp_joining_socket_set = connProvider.newConnectionPollGroup()
    If d.tmp_joining_socket_set Is Nothing Then
        debug_ 1758, LOG_ERROR, "processOpenConnectionResult", "Cannot create socket set - out of memory?"
        WzJoiningGameScreen_HandlerRoot_closeConnectionAttempt self
        Exit Sub
    End If
    debug_ 1762, LOG_NET, "processOpenConnectionResult", "Created socket_set %p", ObjPtr(d.tmp_joining_socket_set)

    ' `client_transient_socket` is used to talk to host machine
    d.tmp_joining_socket_set.add d.client_transient_socket

    ' Create temporary NETQUEUE
    d.tmpJoiningQUEUE = NETnetJoinTmpQueue()
    NETinitQueue d.tmpJoiningQUEUE
    d.has_tmpJoiningQueuePair = True

    ' transition to next internal state, handling responses in processJoining()
    d.currentJoiningState = JoiningState_AwaitingInitialNetcodeHandshakeAck
    d.startTime = wzGetTicks()
End Sub

' processOpenConnectionResult's pushu32 lambda: wz_htonl(value) at buffer(ofs)
Private Sub pushu32(buffer() As Byte, ByVal ofs As Long, ByVal value As Long)
    buffer(ofs) = SHRU(value, 24) And &HFF
    buffer(ofs + 1) = SHRU(value, 16) And &HFF
    buffer(ofs + 2) = SHRU(value, 8) And &HFF
    buffer(ofs + 3) = value And &HFF
End Sub

' processOpenConnectionResult's NETnetJoinTmpQueue lambda: queue = &tmpJoiningQueuePair (WzNet's pair: see above),
' isPair, index NET_HOST_ONLY, queueType QUEUE_TRANSIENT_JOIN, exclude NET_NO_EXCLUDE (a NETQUEUE is a Long:
' queueType * &H10000 + index * &H100 + exclude)
Private Function NETnetJoinTmpQueue() As Long
    NETnetJoinTmpQueue = QUEUE_TRANSIENT_JOIN * &H10000 + NET_HOST_ONLY * &H100& + NET_NO_EXCLUDE
End Function

Private Sub WzJoiningGameScreen_HandlerRoot_attemptToOpenConnection(ByVal self As WIDGET, ByVal connectionIdx As Long)
    Dim d As WzJoiningGameScreen_HandlerRoot, description As InternalConnectionInfo, weakSelf As Long, ct As ConnectionProviderType
    Dim connProvider As WzConnectionProvider, callback As JoiningOpenConnectionCallback
    Set d = WzJoiningGameScreen_HandlerRoot_(self)
    If Not (connectionIdx < d.connectionListCount) Then ASSERT_FAILURE "connectionIdx < connectionList.size()", "src/screens/joiningscreen.cpp:1789", "attemptToOpenConnection", "Invalid connectionIdx: %zu", connectionIdx: Exit Sub
    If Not (d.currentJoiningState = JoiningState_AwaitingConnection Or d.currentJoiningState = JoiningState_NeedsPassword) Then ASSERT_FAILURE "currentJoiningState == JoiningState::AwaitingConnection || currentJoiningState == JoiningState::NeedsPassword", "src/screens/joiningscreen.cpp:1790", "attemptToOpenConnection", "Unexpected joining state: %d", d.currentJoiningState: Exit Sub
    d.currentJoiningState = JoiningState_AwaitingConnection
    d.currentConnectionIdx = connectionIdx
    description = d.connection(connectionIdx)     ' JoinConnectionDescription& description = connectionList[connectionIdx].conn;
    If description.conn.port = 0 Then
        description.conn.port = NETgetGameserverPort() ' use default configured port
        d.setConnection connectionIdx, description
    End If
    weakSelf = d.weakSerial

    ' Reset any prior networking state before trying to connect
    NETshutdown

    ct = toConnectionProviderType(description.conn.type_)
    NETinit ct
    Set connProvider = ConnectionProviderRegistry_Get(ct)
    Set callback = New JoiningOpenConnectionCallback
    callback.init weakSelf, connectionIdx
    connProvider.openClientConnectionAsync description.conn.host, description.conn.port, CLIENT_OPEN_ASYNC_TIMEOUT, callback
    WzJoiningGameScreen_HandlerRoot_updateJoiningStatus self, "Establishing connection with host"
End Sub

Private Function WzJoiningGameScreen_HandlerRoot_joiningSocketNETsend(ByVal self As WIDGET) As Boolean
    Dim d As WzJoiningGameScreen_HandlerRoot, message As NetMessage, rawData() As Byte, rawLen As Long, msgType As Long
    Dim compressedRawLen As Long, ec As Long, writeErrMsg As String, flushDummy As Long
    Set d = WzJoiningGameScreen_HandlerRoot_(self)
    ' NetQueue *queue = &tmpJoiningQueuePair->send;
    Set message = NETtransientJoinSendQueue_getMessageForNet()
    rawData = message.rawData()
    rawLen = bytesCount(rawData)
    msgType = message.type_()

    compressedRawLen = 0
    d.client_transient_socket.writeAll rawData, 0, rawLen, compressedRawLen, ec
    NETtransientJoinSendQueue_popMessageForNet
    If ec = 0 Then
        ' success writing to socket
        debug_ 1835, LOG_NET, "joiningSocketNETsend", "Wrote initial message to socket to host"
    Else
        writeErrMsg = net_errorMessage(ec)
        ' Write error, most likely host disconnect.
        debug_ 1841, LOG_ERROR, "joiningSocketNETsend", "Failed to send message (type: %u, rawLen: %zu, compressedRawLen: %zu) to host: %s", msgType, rawLen, compressedRawLen, writeErrMsg
        WzJoiningGameScreen_HandlerRoot_joiningSocketNETsend = False
        Exit Function
    End If
    If Not (NETtransientJoinSendQueue_numMessagesForNet() = 0) Then ASSERT_FAILURE "queue->numMessagesForNet() == 0", "src/screens/joiningscreen.cpp:1844", "joiningSocketNETsend", "Queue not empty (%u messages remaining).", NETtransientJoinSendQueue_numMessagesForNet()
    ' Make sure the message was completely sent.
    If d.client_transient_socket.flush(flushDummy) <> 0 Then
        WzJoiningGameScreen_HandlerRoot_joiningSocketNETsend = False
        Exit Function
    End If
    WzJoiningGameScreen_HandlerRoot_joiningSocketNETsend = True
End Function

Private Sub WzJoiningGameScreen_HandlerRoot_closeConnectionAttempt(ByVal self As WIDGET)
    Dim d As WzJoiningGameScreen_HandlerRoot
    Set d = WzJoiningGameScreen_HandlerRoot_(self)
    If Not d.client_transient_socket Is Nothing Then
        If Not d.tmp_joining_socket_set Is Nothing Then
            d.tmp_joining_socket_set.remove d.client_transient_socket
        End If
        d.client_transient_socket.close_
        Set d.client_transient_socket = Nothing
    End If
    If Not d.tmp_joining_socket_set Is Nothing Then
        Set d.tmp_joining_socket_set = Nothing      ' delete tmp_joining_socket_set
    End If
    If d.has_tmpJoiningQueuePair Then
        NETtransientJoinQueuePair_delete            ' delete tmpJoiningQueuePair
        d.has_tmpJoiningQueuePair = False
    End If
    d.setInitialAckByte 0, 0                        ' initialAckBuffer[0] = '\0';
    d.usedInitialAckBuffer = 0
End Sub

' static std::vector<uint8_t> serializeConnectionDescription(const JoinConnectionDescription& connDesc) (the size)
Private Function serializeConnectionDescription(connDesc As JoinConnectionDescription, out() As Byte) As Long
    Dim connDescJson As nlohmann_json, connDescJsonStr As String
    Set connDescJson = JoinConnectionDescription_to_json(connDesc)
    connDescJsonStr = connDescJson.dump(-1, " ", False, error_handler_t_replace)
    serializeConnectionDescription = utf8_encode(connDescJsonStr, out)
End Function

Private Function WzJoiningGameScreen_HandlerRoot_verifyEncryptedHostResponse(ByVal self As WIDGET, encryptedHostChallengeResponse() As Byte) As Boolean
    Dim d As WzJoiningGameScreen_HandlerRoot, sk As SessionKeys, hostChallengeResponse() As Byte, responseLen As Long
    Dim hostIdentity As EcKey, challengeForHost() As Byte
    Set d = WzJoiningGameScreen_HandlerRoot_(self)
    ' Decrypt the encryptedHostChallengeResponse
    sk = d.connectionAuthSessionKeys
    If Not SessionKeys_decryptMessageFromOther(sk, encryptedHostChallengeResponse, 0, bytesCount(encryptedHostChallengeResponse), hostChallengeResponse, responseLen) Then
        debug_ 1891, LOG_ERROR, "verifyEncryptedHostResponse", "Invalid host challenge response data received!"
        WzJoiningGameScreen_HandlerRoot_verifyEncryptedHostResponse = False
        Exit Function
    End If
    d.connectionAuthSessionKeys = sk

    ' Verify the host identity challenge response
    hostIdentity = d.hostIdentity
    d.getChallengeForHost challengeForHost
    If Not EcKey_verify(hostIdentity, hostChallengeResponse, responseLen, challengeForHost, 0, bytesCount(challengeForHost)) Then
        debug_ 1898, LOG_ERROR, "verifyEncryptedHostResponse", "Unable to verify host challenge response!"
        WzJoiningGameScreen_HandlerRoot_verifyEncryptedHostResponse = False
        Exit Function
    End If

    WzJoiningGameScreen_HandlerRoot_verifyEncryptedHostResponse = True
End Function

' called from run()
Private Sub WzJoiningGameScreen_HandlerRoot_processJoining(ByVal self As WIDGET)
    Dim d As WzJoiningGameScreen_HandlerRoot, timePassed As Double, nReadable As Long, ec As Long, readCount As Long
    Dim rawDummy As Long, ackBytes(0 To 3) As Byte, result As Long, msgType As Long, r As MessageReader, w As MessageWriter
    Dim index As Long, hostPlayer As Long, blindModeVal As Long, encryptedHostChallengeResponse() As Byte
    Dim sock As IClientConnection, sockSet As IConnectionPollGroup, hostKey() As Byte, rejection As Long, reason As String
    Dim redirectInfo As KickRedirectInfo, ehp As ExpectedHostProperties, list() As InternalConnectionInfo, n As Long
    Dim challengeFromHost() As Byte, challengeSize As Long, hostPublicKey() As Byte, hostIdentity As EcKey
    Dim playerIdentity As EcKey, sk As SessionKeys, connectionDescriptionSerializedBytes() As Byte, connBytesSize As Long
    Dim challengeResponse() As Byte, challengeResponseSize As Long, identity() As Byte, playerType As Long
    Dim modListStr As String, challengeForHost() As Byte, joinChallengeAuthResponse() As Byte, jcarSize As Long
    Dim encryptedChallengeResponse() As Byte, encryptedSize As Long, joinToken As String, conn As InternalConnectionInfo
    Dim readBufferSlice() As Byte, i As Long, ackByte As Long
    Set d = WzJoiningGameScreen_HandlerRoot_(self)
    If d.currentJoiningState = JoiningState_Success Then
        Exit Sub
    End If
    If d.currentJoiningState = JoiningState_SuccessPendingClose Then
        ' close the joining screen after a minimum duration has passed from when it opened (to prevent a "flickering" effect)
        timePassed = widgSteadyNowMs() - d.timeStarted
        If timePassed >= minimumTimeBeforeAutoClose Then
            d.currentJoiningState = JoiningState_Success
            shutdownJoiningAttemptInternal self.screenPointer
        End If
        Exit Sub
    End If
    If d.currentJoiningState = JoiningState_NeedsLobbyErrorResolve Then
        ' waiting for user to resolve the error
        Exit Sub
    End If
    If d.currentJoiningState = JoiningState_FetchingLobbyJoinInfo Then
        ' waiting for lobby to return join info
        Exit Sub
    End If
    If d.currentJoiningState = JoiningState_NeedsPassword Then
        ' waiting for user to enter + submit password
        Exit Sub
    End If
    If d.currentJoiningState = JoiningState_AwaitingConnection Or d.currentJoiningState = JoiningState_Failure Then
        ' nothing much to do until background thread hands over a connection
        Exit Sub
    End If

    If d.client_transient_socket Is Nothing Then
        ' internal error - client_transient_socket should be valid
        WzJoiningGameScreen_HandlerRoot_closeConnectionAttempt self
        WzJoiningGameScreen_HandlerRoot_handleFailure self, FailureDetails_makeFromInternalError("An internal error occurred") ' internal error
        Exit Sub
    End If

    If ULT(HOST_RESPONSE_TIMEOUT, wzGetTicks() - d.startTime) Then  ' (unsigned)wzGetTicks() - startTime > HOST_RESPONSE_TIMEOUT
        ' exceeded timeout
        WzJoiningGameScreen_HandlerRoot_closeConnectionAttempt self
        WzJoiningGameScreen_HandlerRoot_tryNextConnectionOption self, d.currentConnectionIdx, make_network_error_code(ETIMEDOUT), "Timeout while waiting for host to respond"
        Exit Sub
    End If

    ' Handling JoiningStatus::AwaitingInitialNetcodeHandshakeAck
    If d.currentJoiningState = JoiningState_AwaitingInitialNetcodeHandshakeAck Then
        ' read in data, if we have it
        nReadable = d.tmp_joining_socket_set.checkConnectionsReadable(NET_READ_TIMEOUT, ec)
        If ec <> 0 Then nReadable = 0                   ' .value_or(0)
        If nReadable > 0 Then
            If Not d.client_transient_socket.readReady() Then
                Exit Sub ' wait for next check
            End If

            ' (VB) readNoInt into initialAckBuffer + usedInitialAckBuffer: read into ackBytes, then copied over
            readCount = d.client_transient_socket.readNoInt(ackBytes, 0, expectedInitialAckSize - d.usedInitialAckBuffer, rawDummy, ec)
            If ec = 0 Then
                For i = 0 To readCount - 1
                    d.setInitialAckByte d.usedInitialAckBuffer + i, ackBytes(i)
                Next
                d.usedInitialAckBuffer = d.usedInitialAckBuffer + readCount
            End If

            If d.usedInitialAckBuffer >= expectedInitialAckSize Then
                result = ERROR_CONNECTION
                ' memcpy(&result, initialAckBuffer, sizeof(result)); result = wz_ntohl(result);
                result = SHL(d.getInitialAckByte(0), 24) Or SHL(d.getInitialAckByte(1), 16) Or SHL(d.getInitialAckByte(2), 8) Or d.getInitialAckByte(3)
                If result <> ERROR_NOERROR Then
                    debug_ 1986, LOG_ERROR, "processJoining", "Received error %d", result

                    WzJoiningGameScreen_HandlerRoot_closeConnectionAttempt self

                    If result = ERROR_WRONGPASSWORD Then
                        ' We didn't send a password yet, so this is an invalid response - just treat as a generic failure
                        result = ERROR_CONNECTION
                    End If
                    If result = ERROR_INVALID Then
                        ' Not expected to be sent by hosts on join, map to generic failure
                        result = ERROR_CONNECTION
                    End If
                    WzJoiningGameScreen_HandlerRoot_handleFailure self, FailureDetails_makeFromLobbyError(result)
                    Exit Sub
                End If

                ' transition to net message mode (enable compression, wait for messages)
                d.client_transient_socket.enableCompression
                d.currentJoiningState = JoiningState_ProcessingJoinMessages
                ' permit fall-through to currentJoiningState == JoiningState::ProcessingJoinMessage case below
            End If
        End If
    End If

    ' Handling JoiningStatus::ProcessingJoinMessages
    If d.currentJoiningState = JoiningState_ProcessingJoinMessages Then
        ' read in data, if we have it
        nReadable = d.tmp_joining_socket_set.checkConnectionsReadable(NET_READ_TIMEOUT, ec)
        If ec <> 0 Then nReadable = 0                   ' .value_or(0)
        If nReadable > 0 Then
            If Not d.client_transient_socket.readReady() Then
                Exit Sub ' wait for next check
            End If

            readCount = d.client_transient_socket.readNoInt(readBuffer, 0, MaxMsgSize, rawDummy, ec)
            If ec <> 0 Then
                ' disconnect or programmer error
                If net_errorEqualsErrc(ec, errc_timed_out) Or net_errorEqualsErrc(ec, errc_connection_reset) Then
                    debug_ 2030, LOG_NET, "processJoining", "Client socket disconnected."
                Else
                    debug_ 2035, LOG_NET, "processJoining", "Client socket encountered error: %s", net_errorMessage(ec)
                End If
                NETlogEntry "Client socket disconnected (allowJoining)", SYNC_FLAG, d.startTime
                debug_ 2038, LOG_NET, "processJoining", "freeing temp socket %p (%d)", ObjPtr(d.client_transient_socket), 2038
                WzJoiningGameScreen_HandlerRoot_closeConnectionAttempt self
                WzJoiningGameScreen_HandlerRoot_handleFailure self, FailureDetails_makeFromInternalError("Invalid host - disconnected")
                Exit Sub
            Else
                NETinsertRawData d.tmpJoiningQUEUE, readBuffer, readCount
            End If
        End If

        If Not NETisMessageReady(d.tmpJoiningQUEUE) Then
            ' need to wait for a full and complete join message
            ' sanity check
            If ULT(MaxMsgSize * 8, NETincompleteMessageDataBuffered(d.tmpJoiningQUEUE)) Then   ' something definitely big enough to encompass the expected message(s) at this point
                ' host is sending data that doesn't appear to be a properly formatted message - cut it off
                WzJoiningGameScreen_HandlerRoot_closeConnectionAttempt self
                WzJoiningGameScreen_HandlerRoot_handleFailure self, FailureDetails_makeFromInternalError("Invalid host response")
                Exit Sub
            End If
            Exit Sub ' nothing to do until more data comes in
        End If

        msgType = NETgetMessage(d.tmpJoiningQUEUE).type_()

        If msgType = NET_ACCEPTED Then
            ' :)
            hostPlayer = MAX_CONNECTED_PLAYERS + 1 ' invalid host index
            blindModeVal = 0

            Set r = NETbeginDecode(d.tmpJoiningQUEUE, NET_ACCEPTED)
            ' Retrieve the player ID the game host arranged for us
            NETuint8_t r, index
            NETuint32_t r, hostPlayer ' and the host player idx
            NETuint8_t r, blindModeVal
            NETbytes r, encryptedHostChallengeResponse
            NETend r
            NETpop d.tmpJoiningQUEUE

            If Not ULT(hostPlayer, MAX_CONNECTED_PLAYERS) Then
                debug_ 2084, LOG_ERROR, "processJoining", "Bad host player number (%u) received from host!", NetPlay_hostPlayer
                WzJoiningGameScreen_HandlerRoot_closeConnectionAttempt self
                WzJoiningGameScreen_HandlerRoot_handleFailure self, FailureDetails_makeFromInternalError("Invalid host response")
                Exit Sub
            End If

            If index >= MAX_CONNECTED_PLAYERS Then
                debug_ 2092, LOG_ERROR, "processJoining", "Bad player number (%u) received from host!", index
                WzJoiningGameScreen_HandlerRoot_closeConnectionAttempt self
                WzJoiningGameScreen_HandlerRoot_handleFailure self, FailureDetails_makeFromInternalError("Invalid host response")
                Exit Sub
            End If

            ' Verify the host identity challenge response
            If Not WzJoiningGameScreen_HandlerRoot_verifyEncryptedHostResponse(self, encryptedHostChallengeResponse) Then
                WzJoiningGameScreen_HandlerRoot_closeConnectionAttempt self
                WzJoiningGameScreen_HandlerRoot_handleFailure self, FailureDetails_makeFromInternalError("Invalid host response")
                Exit Sub
            End If

            If blindModeVal > BLIND_MODE_MAX Then
                debug_ 2108, LOG_ERROR, "processJoining", "Bad blind mode (%u) received from host!", blindModeVal
                WzJoiningGameScreen_HandlerRoot_closeConnectionAttempt self
                WzJoiningGameScreen_HandlerRoot_handleFailure self, FailureDetails_makeFromInternalError("Invalid host response")
                Exit Sub
            End If

            game.blindMode = blindModeVal
            If game.blindMode <> BLIND_MODE_NONE Then
                ' currently permitted only if hostPlayer is a spectator
                If ULT(hostPlayer, MAX_PLAYER_SLOTS) Then
                    debug_ 2120, LOG_ERROR, "processJoining", "Bad blind mode (%u) received from host!", blindModeVal
                    WzJoiningGameScreen_HandlerRoot_closeConnectionAttempt self
                    WzJoiningGameScreen_HandlerRoot_handleFailure self, FailureDetails_makeFromInternalError("Invalid host response")
                    Exit Sub
                End If

                ' generate a fresh "blind identity"
                generateBlindIdentity
            End If

            ' On success, promote the temporary socket / socketset / queuepair to their permanent (stable) locations, owned by netplay
            ' (Function consumes the socket-related inputs)
            Set sock = d.client_transient_socket
            Set sockSet = d.tmp_joining_socket_set
            If Not NETpromoteJoinAttemptToEstablishedConnectionToHost(hostPlayer, index, d.playerName, d.tmpJoiningQUEUE, sock, sockSet) Then
                Set d.client_transient_socket = sock
                Set d.tmp_joining_socket_set = sockSet
                debug_ 2134, LOG_ERROR, "processJoining", "Failed to promote to established host connection"
                WzJoiningGameScreen_HandlerRoot_closeConnectionAttempt self
                WzJoiningGameScreen_HandlerRoot_handleFailure self, FailureDetails_makeFromInternalError("An internal error occurred")
                Exit Sub
            End If
            Set d.client_transient_socket = sock
            Set d.tmp_joining_socket_set = sockSet
            ' tmpJoiningQueuePair is now "owned" by NetPlay.hostPlayer's netQueue - do not delete it here!
            d.has_tmpJoiningQueuePair = False

            If (game.blindMode = BLIND_MODE_NONE) Or (Not ULT(NetPlay_hostPlayer, MAX_PLAYER_SLOTS)) Then
                hostIdentity = d.hostIdentity
                EcKey_toBytes hostIdentity, EcKey_Public, hostKey
                multiStatsSetVerifiedHostIdentityFromJoin hostKey
            End If

            WzJoiningGameScreen_HandlerRoot_handleSuccess self
            Exit Sub
        ElseIf msgType = NET_REJECTED Then
            rejection = 0
            reason = ""

            Set r = NETbeginDecode(d.tmpJoiningQUEUE, NET_REJECTED)
            NETuint8_t r, rejection
            NETstring r, reason, MAX_JOIN_REJECT_REASON
            If rejection = ERROR_REDIRECT Then
                NETbytes r, encryptedHostChallengeResponse
            End If
            NETend r
            NETpop d.tmpJoiningQUEUE

            debug_ 2166, LOG_NET, "processJoining", "NET_REJECTED received. Error code: %u", rejection

            WzJoiningGameScreen_HandlerRoot_closeConnectionAttempt self

            If rejection = ERROR_REDIRECT And d.numJoinRedirects = 0 Then
                d.numJoinRedirects = d.numJoinRedirects + 1

                ' Verify the host identity challenge response
                If Not WzJoiningGameScreen_HandlerRoot_verifyEncryptedHostResponse(self, encryptedHostChallengeResponse) Then
                    WzJoiningGameScreen_HandlerRoot_closeConnectionAttempt self
                    WzJoiningGameScreen_HandlerRoot_handleFailure self, FailureDetails_makeFromInternalError("Invalid host response")
                    Exit Sub
                End If

                conn = d.connection(d.currentConnectionIdx)
                If Not parseKickRedirectInfo(reason, conn.conn.host, redirectInfo) Then
                    WzJoiningGameScreen_HandlerRoot_closeConnectionAttempt self
                    WzJoiningGameScreen_HandlerRoot_handleFailure self, FailureDetails_makeFromInternalError("Unable to process redirect")
                    Exit Sub
                End If
                If redirectInfo.connListCount = 0 Then
                    ' No valid connection descriptions!
                    debug_ 2192, LOG_ERROR, "processJoining", "No valid connection descriptions in redirect: %s", reason
                    WzJoiningGameScreen_HandlerRoot_closeConnectionAttempt self
                    WzJoiningGameScreen_HandlerRoot_handleFailure self, FailureDetails_makeFromInternalError("Unable to process redirect")
                    Exit Sub
                End If

                ' Update expected host properties
                ehp = d.expectedHostProps
                hostIdentity = d.hostIdentity
                ehp.has_hostPublicKey = True
                EcKey_toBytes hostIdentity, EcKey_Public, ehp.hostPublicKey ' new join must match current hostPublicKey
                ehp.has_gamePassword = redirectInfo.has_gamePassword
                ehp.gamePassword = redirectInfo.gamePassword
                d.expectedHostProps = ehp

                WzJoiningGameScreen_HandlerRoot_closeConnectionAttempt self

                ' Attempt to join using the redirect connection details
                n = convertConnectionDescriptionsToInternalConnectionList(redirectInfo.connList, redirectInfo.connListCount, list)
                WzJoiningGameScreen_HandlerRoot_setConnectionList self, list, n
                d.asSpectator = redirectInfo.asSpectator
                WzJoiningGameScreen_HandlerRoot_applyExpectedHostProps self
                d.currentJoiningState = JoiningState_AwaitingConnection
                WzJoiningGameScreen_HandlerRoot_attemptToOpenConnection self, 0
                Exit Sub
            ElseIf rejection = ERROR_WRONGPASSWORD Then
                ehp = d.expectedHostProps
                If Not ehp.has_gamePassword Then
                    d.currentJoiningState = JoiningState_NeedsPassword
                    WzJoiningGameScreen_HandlerRoot_promptForPassword self
                Else
                    ' started join with expected game password, but the host rejected the password - treat as a failure
                    WzJoiningGameScreen_HandlerRoot_handleFailure self, FailureDetails_makeFromLobbyError(ERROR_WRONGDATA)
                End If
            Else
                ' size_t reasonStrLength = strnlen(reason, MAX_JOIN_REJECT_REASON);
                If Len(reason) > 0 Then
                    WzJoiningGameScreen_HandlerRoot_handleFailure self, FailureDetails_makeFromHostRejectionMessage(reason)
                Else
                    WzJoiningGameScreen_HandlerRoot_handleFailure self, FailureDetails_makeFromLobbyError(rejection)
                End If
            End If
            Exit Sub
        ElseIf msgType = NET_PING Then
            WzJoiningGameScreen_HandlerRoot_updateJoiningStatus self, "Requesting to join game"

            ' std::vector<uint8_t> challengeFromHost(NETgetJoinConnectionNETPINGChallengeFromHostSize(), 0); (replaced by NETbytes)
            ReDim challengeFromHost(0 To NETgetJoinConnectionNETPINGChallengeFromHostSize() - 1)
            Set r = NETbeginDecode(d.tmpJoiningQUEUE, NET_PING)
            NETbytes r, challengeFromHost, NETgetJoinConnectionNETPINGChallengeFromHostSize() * 4
            NETbytes r, hostPublicKey
            NETend r
            NETpop d.tmpJoiningQUEUE

            challengeSize = bytesCount(challengeFromHost)
            If challengeSize <> 0 And challengeSize < NETgetJoinConnectionNETPINGChallengeFromHostSize() Then
                ' Invalid challenge sent by host
                debug_ 2254, LOG_ERROR, "processJoining", "Invalid host challenge"
                ' Disconnect and treat as a failure
                WzJoiningGameScreen_HandlerRoot_closeConnectionAttempt self
                WzJoiningGameScreen_HandlerRoot_handleFailure self, FailureDetails_makeFromInternalError("Invalid host establishing data")
                Exit Sub
            End If

            ' Load host public key
            hostIdentity = d.hostIdentity
            If Not EcKey_fromBytes(hostIdentity, hostPublicKey, bytesCount(hostPublicKey), EcKey_Public) Then
                d.hostIdentity = hostIdentity
                debug_ 2264, LOG_ERROR, "processJoining", "Invalid host identity"
                ' Disconnect and treat as a failure
                WzJoiningGameScreen_HandlerRoot_closeConnectionAttempt self
                WzJoiningGameScreen_HandlerRoot_handleFailure self, FailureDetails_makeFromInternalError("Invalid host identity")
                Exit Sub
            End If
            d.hostIdentity = hostIdentity

            ehp = d.expectedHostProps
            If ehp.has_hostPublicKey Then
                ' Verify that the sent host public key matches the expected
                ' (Note: Actual verification of the key itself comes later with a challenge/response)
                If Not bytesEqual(hostPublicKey, ehp.hostPublicKey) Then
                    debug_ 2277, LOG_ERROR, "processJoining", "Unexpected host identity"
                    ' Disconnect and treat as a failure
                    WzJoiningGameScreen_HandlerRoot_closeConnectionAttempt self
                    WzJoiningGameScreen_HandlerRoot_handleFailure self, FailureDetails_makeFromInternalError("Invalid host identity")
                    Exit Sub
                End If
            End If

            ' try { connectionAuthSessionKeys = std::make_unique<SessionKeys>(playerIdentity, 1, hostIdentity, 0); }
            playerIdentity = d.playerIdentity
            If Not SessionKeys_init(sk, playerIdentity, 1, hostIdentity, 0) Then
                ' catch (const std::invalid_argument& e) (e.what(): SessionKeys_lastError)
                debug_ 2289, LOG_INFO, "processJoining", "Cannot create initial session key with host, error: %s", SessionKeys_lastError()
                ' Disconnect and treat as a failure
                WzJoiningGameScreen_HandlerRoot_closeConnectionAttempt self
                WzJoiningGameScreen_HandlerRoot_handleFailure self, FailureDetails_makeFromInternalError("Failed to establish session keys with host")
                Exit Sub
            End If
            d.connectionAuthSessionKeys = sk

            conn = d.connection(d.currentConnectionIdx)
            connBytesSize = serializeConnectionDescription(conn.conn, connectionDescriptionSerializedBytes)
            challengeResponseSize = 0
            If challengeSize <> 0 Then
                challengeResponseSize = EcKey_sign(playerIdentity, challengeFromHost, 0, challengeSize, challengeResponse)
            End If
            EcKey_toBytes playerIdentity, EcKey_Public, identity
            If Not d.asSpectator Then playerType = NET_JOIN_PLAYER Else playerType = NET_JOIN_SPECTATOR
            modListStr = getModList()

            ' generate a challenge for the host to sign
            ReDim challengeForHost(0 To NETgetJoinConnectionNETPINGChallengeFromClientSize() - 1)
            genSecRandomBytes challengeForHost, 0, NETgetJoinConnectionNETPINGChallengeFromClientSize()
            d.setChallengeForHost challengeForHost

            jcarSize = 0
            NETbytesOutputToVector challengeResponse, challengeResponseSize, joinChallengeAuthResponse, jcarSize
            NETbytesOutputToVector connectionDescriptionSerializedBytes, connBytesSize, joinChallengeAuthResponse, jcarSize
            NETbytesOutputToVector challengeForHost, bytesCount(challengeForHost), joinChallengeAuthResponse, jcarSize

            encryptedSize = SessionKeys_encryptMessageForOther(sk, joinChallengeAuthResponse, 0, jcarSize, encryptedChallengeResponse)
            d.connectionAuthSessionKeys = sk
            If encryptedSize = 0 Then
                debug_ 2318, LOG_ERROR, "processJoining", "Failed to encrypt response"
                ' Disconnect and treat as a failure
                WzJoiningGameScreen_HandlerRoot_closeConnectionAttempt self
                WzJoiningGameScreen_HandlerRoot_handleFailure self, FailureDetails_makeFromInternalError("Failure to generate valid response")
                Exit Sub
            End If

            Set w = NETbeginEncode(d.tmpJoiningQUEUE, NET_JOIN)
            NETstring w, d.playerName, jmin(StringSize, utf8Size(d.playerName) + 1)
            NETstring w, modListStr, jmin(modlist_string_size, utf8Size(modListStr) + 1)
            NETstring w, d.gamePassword, password_string_size
            NETuint8_t w, playerType
            NETbytes w, identity
            NETbytes w, encryptedChallengeResponse
            joinToken = ""
            If d.currentConnectionIdx < d.connectionListCount Then
                joinToken = conn.joinToken
            End If
            NETstring_std w, joinToken
            NETend w ' because of QUEUE_TRANSIENT_JOIN type, this won't trigger a NETsend() - we must write ourselves
            WzJoiningGameScreen_HandlerRoot_joiningSocketNETsend self
            ' and now we wait for the host to respond with a further message
            Exit Sub
        Else
            debug_ 2345, LOG_ERROR, "processJoining", "Unexpected %s.", messageTypeToString(msgType)
            NETpop d.tmpJoiningQUEUE
            ' Disconnect and treat as a failure
            WzJoiningGameScreen_HandlerRoot_closeConnectionAttempt self
            WzJoiningGameScreen_HandlerRoot_handleFailure self, FailureDetails_makeFromInternalError("Invalid host response")
            Exit Sub
        End If
    End If
End Sub

' MARK: - WzJoiningGameScreen

' static std::shared_ptr<WzJoiningGameScreen> makeWithLobbyGame(...)
Private Function WzJoiningGameScreen_makeWithLobbyGame(ByVal lobbyAddress As String, ByVal gameId As String, ByVal playerName As String, _
        playerIdentity As EcKey, ByVal asSpectator As Boolean, expectedHostProps As ExpectedHostProperties, _
        ByVal has_known As Boolean, known() As netlobby_ConnectionType, ByVal knownCount As Long, _
        ByVal onSuccessFunc As WzFunc, ByVal onFailureFunc As WzFunc) As W_SCREEN
    Dim newRootFrm As WIDGET, screen As W_SCREEN, psWeakHelpOverlayScreen As Long
    ' (class make_shared_enabler: public WzJoiningGameScreen {};)
    Set newRootFrm = WzJoiningGameScreen_HandlerRoot_makeWithLobbyGame(lobbyAddress, gameId, playerName, playerIdentity, asSpectator, expectedHostProps, has_known, known, knownCount, onSuccessFunc, onFailureFunc)
    Set screen = W_SCREEN_make(newRootFrm)  ' screen->initialize(newRootFrm);
    psWeakHelpOverlayScreen = ObjPtr(screen)
    Set WzJoiningGameScreen_HandlerRoot_(newRootFrm).onCancelPressed = WzFunc_new(FN_joiningScreen_cancelPressed, psWeakHelpOverlayScreen)

    Set WzJoiningGameScreen_makeWithLobbyGame = screen
End Function

' static std::shared_ptr<WzJoiningGameScreen> make(...)
Private Function WzJoiningGameScreen_make(connectionList() As JoinConnectionDescription, ByVal connectionListCount As Long, _
        ByVal playerName As String, playerIdentity As EcKey, ByVal asSpectator As Boolean, expectedHostProps As ExpectedHostProperties, _
        ByVal onSuccessFunc As WzFunc, ByVal onFailureFunc As WzFunc) As W_SCREEN
    Dim newRootFrm As WIDGET, screen As W_SCREEN, psWeakHelpOverlayScreen As Long
    ' (class make_shared_enabler: public WzJoiningGameScreen {};)
    Set newRootFrm = WzJoiningGameScreen_HandlerRoot_make(connectionList, connectionListCount, playerName, playerIdentity, asSpectator, expectedHostProps, onSuccessFunc, onFailureFunc)
    Set screen = W_SCREEN_make(newRootFrm)  ' screen->initialize(newRootFrm);
    psWeakHelpOverlayScreen = ObjPtr(screen)
    Set WzJoiningGameScreen_HandlerRoot_(newRootFrm).onCancelPressed = WzFunc_new(FN_joiningScreen_cancelPressed, psWeakHelpOverlayScreen)

    Set WzJoiningGameScreen_make = screen
End Function

' void WzJoiningGameScreen::closeScreen()
Private Sub WzJoiningGameScreen_closeScreen(ByVal self As W_SCREEN)
    shutdownJoiningAttemptInternal self
End Sub

' Internal helpers

' static void handleJoinSuccess(const JoinConnectionDescription& connection, const PLAYERSTATS& playerStats)
Private Sub handleJoinSuccess(ByVal connectionHost As String, ByVal connectionPort As Long, playerStats As PLAYERSTATS_)
    ingame.side = InGameSide_MULTIPLAYER_CLIENT

    ingame.localJoiningInProgress = True

    ' send initial messages of player data (stats, color request, etc)
    setMultiStats selectedPlayer, playerStats, False
    setMultiStats selectedPlayer, playerStats, True

    loadMultiOptionPrefValues sPlayer, selectedPlayer
    sendPlayerMultiOptPreferencesBuiltin

    If ULT(selectedPlayer, MAX_PLAYERS) And war_getMPcolour() >= 0 Then
        SendColourRequest selectedPlayer, war_getMPcolour()
    End If

    widgScheduleTask WzFunc_new(FN_joining_changeTitleUITask)

    ' NOTE: Do *NOT* close the overlay screen here - this will be handled automatically after a minimum duration has passed
End Sub

' static void handleJoinFailure()
Private Sub handleJoinFailure()
    ' NOTE: Do *NOT* close the overlay screen here - let the user dismiss it when they are done reading the failure message
End Sub

Public Sub shutdownJoiningAttemptInternal(ByVal expectedScreen As W_SCREEN)
    Dim strongJoiningAttemptScreen As W_SCREEN
    widgRemoveOverlayScreen expectedScreen
    Set strongJoiningAttemptScreen = psCurrentJoiningAttemptScreen
    If Not strongJoiningAttemptScreen Is Nothing Then
        If Not (strongJoiningAttemptScreen Is expectedScreen) Then ASSERT_FAILURE "strongJoiningAttemptScreen == expectedScreen", "src/screens/joiningscreen.cpp:2436", "shutdownJoiningAttemptInternal", "Current joining attempt screen does not match - ignoring": Exit Sub
        Set strongJoiningAttemptScreen = Nothing
        Set psCurrentJoiningAttemptScreen = Nothing
    End If
End Sub

Private Function prepareForJoiningAttemptInternal() As Boolean
    Dim currentGameMode As ActivitySink_GameMode
    currentGameMode = ActivityManager_getCurrentGameMode()
    If currentGameMode <> ActivitySink_GameMode_MENUS Then
        ' Can't join a game while already in a game
        debug_ 2448, LOG_ERROR, "prepareForJoiningAttemptInternal", "Can't join a game while already in a game / lobby. (Current mode: %s)", gameModeToString(currentGameMode)
        prepareForJoiningAttemptInternal = False
        Exit Function
    End If

    If ingame.localJoiningInProgress Then
        ' Can't join a game while already in a game
        debug_ 2455, LOG_ERROR, "prepareForJoiningAttemptInternal", "Can't join a game while already in a game / lobby."
        prepareForJoiningAttemptInternal = False
        Exit Function
    End If

    shutdownJoiningAttempt

    ' network communication preparation
    NetPlay_bComms = True ' use network = true

    prepareForJoiningAttemptInternal = True
End Function

Private Function startJoiningAttemptInternal(ByVal playerName As String, connection_list() As JoinConnectionDescription, ByVal connection_listCount As Long, _
        ByVal asSpectator As Boolean, expectedHostProps As ExpectedHostProperties) As Boolean
    Dim playerStats As PLAYERSTATS_, statsHandle As Long, screen As W_SCREEN
    If Not (connection_listCount <> 0) Then ASSERT_FAILURE "!connection_list.empty()", "src/screens/joiningscreen.cpp:2469", "startJoiningAttemptInternal", "Empty connection_list?": Exit Function

    If Not prepareForJoiningAttemptInternal() Then
        startJoiningAttemptInternal = False
        Exit Function
    End If

    loadMultiStats playerName, playerStats
    statsHandle = storeJoinPlayerStats(playerStats)

    Set screen = WzJoiningGameScreen_make(connection_list, connection_listCount, playerName, playerStats.identity, asSpectator, expectedHostProps, _
        WzFunc_new(FN_joining_onSuccess, statsHandle), _
        WzFunc_new(FN_joining_onFailure))
    ' Use widgScheduleTask to ensure we never modify the registered overlays while they are being enumerated
    widgScheduleTask WzFunc_new(FN_joining_registerOverlayTask, screen)
    Set psCurrentJoiningAttemptScreen = screen
    startJoiningAttemptInternal = True
End Function

Private Function startLobbyJoiningAttemptInternal(ByVal playerName As String, ByVal lobbyAddress As String, ByVal gameId As String, _
        ByVal asSpectator As Boolean, expectedHostProps As ExpectedHostProperties, _
        ByVal has_known As Boolean, known() As netlobby_ConnectionType, ByVal knownCount As Long) As Boolean
    Dim playerStats As PLAYERSTATS_, statsHandle As Long, screen As W_SCREEN, lobbyAddressCopy As String, gameIdCopy As String
    If Not (Len(gameId) <> 0) Then ASSERT_FAILURE "!gameId.empty()", "src/screens/joiningscreen.cpp:2500", "startLobbyJoiningAttemptInternal", "Empty gameId?": Exit Function

    If Not prepareForJoiningAttemptInternal() Then
        startLobbyJoiningAttemptInternal = False
        Exit Function
    End If

    loadMultiStats playerName, playerStats
    statsHandle = storeJoinPlayerStats(playerStats)

    lobbyAddressCopy = lobbyAddress
    gameIdCopy = gameId
    Set screen = WzJoiningGameScreen_makeWithLobbyGame(lobbyAddress, gameId, playerName, playerStats.identity, asSpectator, expectedHostProps, _
        has_known, known, knownCount, _
        WzFunc_new(FN_joining_lobbyOnSuccess, lobbyAddressCopy, gameIdCopy, statsHandle), _
        WzFunc_new(FN_joining_onFailure))
    ' Use widgScheduleTask to ensure we never modify the registered overlays while they are being enumerated
    widgScheduleTask WzFunc_new(FN_joining_registerOverlayTask, screen)
    Set psCurrentJoiningAttemptScreen = screen
    startLobbyJoiningAttemptInternal = True
End Function

' MARK: - Public API

' bool startLobbyJoiningAttempt(char* playerName, const std::string& lobbyAddress, const std::string& gameId, bool asSpectator = false,
' const ExpectedHostProperties& expectedHostProps = ExpectedHostProperties(), const std::vector<netlobby::ConnectionType>* pKnownAvailableConnectionTypes = nullptr)
' (VB: the default arguments here; startLobbyJoiningAttempt_props takes all)
Public Function startLobbyJoiningAttempt(ByVal playerName As String, ByVal lobbyAddress As String, ByVal gameId As String, Optional ByVal asSpectator As Boolean = False) As Boolean
    Dim ehp As ExpectedHostProperties, known() As netlobby_ConnectionType
    startLobbyJoiningAttempt = startLobbyJoiningAttempt_props(playerName, lobbyAddress, gameId, asSpectator, ehp, False, known, 0)
End Function

Public Function startLobbyJoiningAttempt_props(ByVal playerName As String, ByVal lobbyAddress As String, ByVal gameId As String, _
        ByVal asSpectator As Boolean, expectedHostProps As ExpectedHostProperties, _
        ByVal has_known As Boolean, known() As netlobby_ConnectionType, ByVal knownCount As Long) As Boolean
    resetJoinRedirectTracking
    startLobbyJoiningAttempt_props = startLobbyJoiningAttemptInternal(playerName, lobbyAddress, gameId, asSpectator, expectedHostProps, has_known, known, knownCount)
End Function

' bool startJoiningAttempt(char* playerName, std::vector<JoinConnectionDescription> connection_list, bool asSpectator = false,
' const ExpectedHostProperties& expectedHostProps = ExpectedHostProperties()) (VB: startJoiningAttempt_props takes the last)
Public Function startJoiningAttempt(ByVal playerName As String, connection_list() As JoinConnectionDescription, ByVal connection_listCount As Long, Optional ByVal asSpectator As Boolean = False) As Boolean
    Dim ehp As ExpectedHostProperties
    startJoiningAttempt = startJoiningAttempt_props(playerName, connection_list, connection_listCount, asSpectator, ehp)
End Function

Public Function startJoiningAttempt_props(ByVal playerName As String, connection_list() As JoinConnectionDescription, ByVal connection_listCount As Long, _
        ByVal asSpectator As Boolean, expectedHostProps As ExpectedHostProperties) As Boolean
    resetJoinRedirectTracking
    startJoiningAttempt_props = startJoiningAttemptInternal(playerName, connection_list, connection_listCount, asSpectator, expectedHostProps)
End Function

Public Sub shutdownJoiningAttempt()
    Dim strongJoiningAttemptScreen As W_SCREEN
    ' Closes the overlay screen, cancelling whatever join attempt is in progress (if one is in progress)
    Set strongJoiningAttemptScreen = psCurrentJoiningAttemptScreen
    If Not strongJoiningAttemptScreen Is Nothing Then
        widgRemoveOverlayScreen strongJoiningAttemptScreen
        Set strongJoiningAttemptScreen = Nothing
        Set psCurrentJoiningAttemptScreen = Nothing
    End If
End Sub

' bool startJoinRedirectAttempt(char* playerName, std::vector<JoinConnectionDescription> connection_list, bool asSpectator = false,
' const ExpectedHostProperties& expectedHostProps = ExpectedHostProperties())
Public Function startJoinRedirectAttempt(ByVal playerName As String, connection_list() As JoinConnectionDescription, ByVal connection_listCount As Long, _
        ByVal asSpectator As Boolean, expectedHostProps As ExpectedHostProperties) As Boolean
    If currentNumRedirectAttempts < 2147483647 Then     ' < std::numeric_limits<size_t>::max()
        currentNumRedirectAttempts = currentNumRedirectAttempts + 1
    End If
    ' The intent is that this allows a host match-making lobby (joined by a client) to pair players and redirect them to another hosted instance
    ' We do not want to allow repeated redirects
    If currentNumRedirectAttempts > 1 Then
        ' POSSIBLE FUTURE TODO: Could prompt before additional redirects (if we wanted to support such a thing)
        ' For now, deny it
        startJoinRedirectAttempt = False
    Else
        startJoinRedirectAttempt = startJoiningAttemptInternal(playerName, connection_list, connection_listCount, asSpectator, expectedHostProps)
    End If
End Function

Public Sub resetJoinRedirectTracking()
    ' Reset the internal counter that track redirects
    currentNumRedirectAttempts = 0
End Sub

Public Function createJoiningIndeterminateProgressWidget(ByVal fontID As iV_fonts) As WIDGET
    Dim progressIndicator As WIDGET
    Set progressIndicator = WzJoiningIndeterminateIndicatorWidget_new(fontID)
    progressIndicator.setGeometry 0, 0, progressIndicator.idealWidth(), progressIndicator.idealHeight()
    Set createJoiningIndeterminateProgressWidget = progressIndicator
End Function

'--------------------------------------------------------------------------
' the lambdas

Public Function joiningscreen_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim strongParent As WIDGET, strongSelf As WIDGET, widg As WIDGET, str_ As String, d As WzJoiningPasswordPrompt
    Dim but As WIDGET, passwordPrompt As WIDGET, jr As netlobby_JoinResult, psWidget As WIDGET, psParent As WIDGET
    Dim desiredWidth As Long, desiredHeight As Long, x0 As Long, y0 As Long, h As WzJoiningGameScreen_HandlerRoot
    Dim psOverlayScreen As W_SCREEN, ps As PLAYERSTATS_, gameIdMsg As String
    Select Case f.id
    Case FN_joiningPasswordPrompt_focusTask         ' [weakSelf]()
        ' give it back focus
        Set strongParent = WeakWidget_lock(f.cap(0))
        If Not (Not strongParent Is Nothing) Then ASSERT_FAILURE "strongParent != nullptr", "src/screens/joiningscreen.cpp:227", "submitPassword", "No parent?": Exit Function
        WzJoiningPasswordPrompt_givePasswordBoxFocus strongParent

    Case FN_joiningPasswordPrompt_onReturn          ' [weakSelf](W_EDITBOX& widg)
        Set widg = args(0)
        str_ = W_EDITBOX_getString(widg)
        Set strongParent = WeakWidget_lock(f.cap(0))
        If Not (Not strongParent Is Nothing) Then ASSERT_FAILURE "strongParent != nullptr", "src/screens/joiningscreen.cpp:272", "initialize", "No parent?": Exit Function
        WzJoiningPasswordPrompt_submitPassword strongParent, str_

    Case FN_joiningPasswordPrompt_onEditingStopped  ' [weakSelf](W_EDITBOX&)
        Set strongParent = WeakWidget_lock(f.cap(0))
        If Not (Not strongParent Is Nothing) Then ASSERT_FAILURE "strongParent != nullptr", "src/screens/joiningscreen.cpp:277", "initialize", "No parent?": Exit Function
        Set d = WzJoiningPasswordPrompt_(strongParent)
        If d.invalidPasswordPlaceholder Then
            W_EDITBOX_setPlaceholder d.passwordBox, "Enter password"
            W_EDITBOX_setPlaceholderTextColor d.passwordBox, False
            d.invalidPasswordPlaceholder = False
        End If

    Case FN_joiningPasswordPrompt_okClick           ' [weakSelf](W_BUTTON&)
        Set strongSelf = WeakWidget_lock(f.cap(0))
        If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/screens/joiningscreen.cpp:290", "initialize", "No parent?": Exit Function
        WzJoiningPasswordPrompt_submitPassword strongSelf, W_EDITBOX_getString(WzJoiningPasswordPrompt_(strongSelf).passwordBox)

    Case FN_joiningStatusForm_actionClick           ' [](W_BUTTON& but)
        Set but = args(0)
        shutdownJoiningAttemptInternal but.screenPointer

    Case FN_joiningStatusForm_wrappedSubmit         ' [weakSelf, submitFunc](WzString str)
        WzFunc_call f.cap(1), args(0)
        Set strongSelf = WeakWidget_lock(f.cap(0))
        If Not strongSelf Is Nothing Then
            Set passwordPrompt = WzJoiningStatusForm_(strongSelf).resolvableJoinErrorWidget
            If Not passwordPrompt Is Nothing Then
                If WIDGET_isKindOf(passwordPrompt, K_WzJoiningPasswordPrompt) Then
                    WzJoiningPasswordPrompt_setDisabled passwordPrompt, True
                End If
            End If
        End If

    Case FN_joiningHandlerRoot_lobbyJoinResult      ' [weakSelf, lobbyAddress, gameId, specificIpVersion](netlobby::JoinResult result)
        ' Handle request join result
        netlobby_JoinResult_take args(0), jr
        Set strongSelf = WeakWidget_lock(f.cap(0))
        If strongSelf Is Nothing Then
            ' background thread ultimately returned after the requester has gone away (join was cancelled?) - just return
            Exit Function
        End If
        WzJoiningGameScreen_HandlerRoot_handleLobbyRequestJoinResult strongSelf, f.cap(1), f.cap(2), f.cap(3), f.cap(4), jr

    Case FN_joiningHandlerRoot_lobbyErrorResolved   ' [weakSelf, lobbyAddressCopy, gameIdCopy, specificIpVersion](resolutionData)
        Set strongSelf = WeakWidget_lock(f.cap(0))
        If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/screens/joiningscreen.cpp:1321", "handleLobbyRequestJoinResult", "No parent form?": Exit Function
        WzJoiningGameScreen_HandlerRoot_requestLobbyJoin strongSelf, f.cap(1), f.cap(2), f.cap(3), f.cap(4), args(0)

    Case FN_joiningHandlerRoot_formInitCalcLayout   ' LAMBDA_CALCLAYOUT_SIMPLE: (WIDGET *psWidget)
        Set psWidget = args(0)
        psWidget.setGeometry 0, 0, screenWidth, screenHeight

    Case FN_joiningHandlerRoot_progressFormCalcLayout ' LAMBDA_CALCLAYOUT_SIMPLE: (WIDGET *psWidget)
        Set psWidget = args(0)
        ' auto psProgressForm = dynamic_cast<WzJoiningStatusForm*>(psWidget);
        Set psParent = psWidget.parent()
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/screens/joiningscreen.cpp:1429", "initialize", "No parent?": Exit Function
        desiredWidth = psWidget.idealWidth()
        desiredHeight = jmin(psWidget.idealHeight(), psParent.height())
        x0 = (psParent.width() - desiredWidth) \ 2
        y0 = (psParent.height() - WzJoiningStatusForm_maximumHeight(psWidget)) \ 2
        psWidget.setGeometry x0, jmax(y0, 0), desiredWidth, desiredHeight

    Case FN_joiningHandlerRoot_passwordSubmitted    ' [weakSelf](WzString newPassword)
        widgScheduleTask WzFunc_new(FN_joiningHandlerRoot_passwordTask, f.cap(0), args(0))

    Case FN_joiningHandlerRoot_passwordTask         ' [weakSelf, newPassword]()
        Set strongSelf = WeakWidget_lock(f.cap(0))
        If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/screens/joiningscreen.cpp:1473", "promptForPassword", "No parent form?": Exit Function
        Set h = WzJoiningGameScreen_HandlerRoot_(strongSelf)
        If Not (h.currentJoiningState = JoiningState_NeedsPassword) Then ASSERT_FAILURE "strongSelf->currentJoiningState == JoiningState::NeedsPassword", "src/screens/joiningscreen.cpp:1474", "promptForPassword", "Unexpected state": Exit Function
        h.gamePassword = sstrcpyTo(f.cap(1), password_string_size)
        ' re-attempt join with last used connection
        WzJoiningGameScreen_HandlerRoot_attemptToOpenConnection strongSelf, h.currentConnectionIdx

    Case FN_joiningHandlerRoot_resultOnMainThread   ' [weakSelf, connectionIdx, ppResult]() mutable
        Set strongSelf = WeakWidget_lock(f.cap(0))
        If strongSelf Is Nothing Then
            ' background thread ultimately returned after the requester has gone away - just return
            Exit Function
        End If
        WzJoiningGameScreen_HandlerRoot_processOpenConnectionResult strongSelf, f.cap(1), f.cap(2)

    Case FN_joiningScreen_cancelPressed             ' [psWeakHelpOverlayScreen]()
        Set psOverlayScreen = W_SCREEN_fromPtr(f.cap(0))
        If Not psOverlayScreen Is Nothing Then
            WzJoiningGameScreen_closeScreen psOverlayScreen
        End If

    Case FN_joining_onSuccess                       ' [playerStats](const JoinConnectionDescription& connection)
        ps = joinPlayerStatsStore(f.cap(0))
        handleJoinSuccess args(0), args(1), ps
        ActivityManager_joinGameSucceeded args(0), args(1)

    Case FN_joining_onFailure                       ' []()
        handleJoinFailure

    Case FN_joining_registerOverlayTask             ' [screen]()
        widgRegisterOverlayScreenOnTopOfScreen f.cap(0), psWScreen

    Case FN_joining_lobbyOnSuccess                  ' [lobbyAddressCopy, gameIdCopy, playerStats](const JoinConnectionDescription& connection)
        ps = joinPlayerStatsStore(f.cap(2))
        handleJoinSuccess args(0), args(1), ps
        ActivityManager_joinLobbyGameSucceeded f.cap(0), f.cap(1), args(0), args(1)

        ' used to schedule this after handleJoinSuccess's scheduled change of the TitleUI, to ensure that WzMultiplayerOptionsTitleUI::start() has been called first
        widgScheduleTask WzFunc_new(FN_joining_lobbyGameIdMsgTask, f.cap(1))

    Case FN_joining_lobbyGameIdMsgTask              ' [gameIdCopy]()
        gameIdMsg = "Lobby GameId"
        gameIdMsg = gameIdMsg & ": " ' build this manually to ensure that the format always separates the prefix text from the gameId with ": ",  without relying on the translation
        gameIdMsg = gameIdMsg & f.cap(0)
        addConsoleMessage gameIdMsg, DEFAULT_JUSTIFY, NOTIFY_MESSAGE

    Case FN_joining_changeTitleUITask               ' []()
        ' switch the TitleUI to the multiplayer options (lobby), which will take over handling messages from the host
        changeTitleUI WzMultiplayerOptionsTitleUI_new(wzTitleUICurrent)
    End Select
End Function
