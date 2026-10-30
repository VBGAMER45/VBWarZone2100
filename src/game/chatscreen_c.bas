Attribute VB_Name = "chatscreen_c"
' src/screens/chatscreen.cpp, chatscreen.h, whole: the in-game chat screen (an overlay screen with the chat edit box,
' the chat options button, the quick chat form). Its widget classes' fields are the data classes named as the
' classes. (VB) struct WzInGameChatScreen : public W_SCREEN is a W_SCREEN (W_SCREEN_make with its root form) whose
' OnCloseFunc this module keeps ("p<ObjPtr>" in chatScreens; isWzInGameChatScreen is the dynamic_pointer_cast);
' ~WzInGameChatScreen's setFocus(nullptr) is on the dying screen and has no effect left. psCurrentChatScreen is a
' strong reference: an address could outlive the screen; shutdownChatScreen drops it. The std::weak_ptr captures of
' widgets are their addresses (WIDGET_fromPtr), as in the other ported screens. createChatScreen keeps its
' stand-in's signature: a Sub (the caller, hci_c, does not use the returned screen). hci/quickchat.cpp and
' hci/chatoptions.cpp are stand-ins (createQuickChatForm, createChatOptionsForm: an empty WIDGET; createQuickChatForm
' takes no WzQuickChatMode); quick chat is offered only in skirmish / multiplayer, so a campaign never makes one.
' Translations (_()) are not ported: the English text is used.
' (VB) widget class WzInGameChatBoxButton : W_BUTTON
' (VB) widget class WzInGameChatBoxForm : IntFormAnimated
' (VB) widget class WzInGameChatScreen_CLICKFORM : W_CLICKFORM
Option Explicit

Private Const MAX_CONNECTED_PLAYERS As Long = MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS

' src/screens/chatscreen.h: enum class WzChatMode
Public Enum WzChatMode
    WzChatMode_Glob
    WzChatMode_Team
End Enum

' the lambdas (WzFunc: chatscreen_c_fn)
Public Const FN_chatBox_onReturn As Long = 14800                ' chatBox->setOnReturnHandler([](W_EDITBOX& widg) {...})
Public Const FN_chatBox_onReturnTask As Long = 14801            ' widgScheduleTask([enteredText, currChatMode, weakParent]() {...})
Public Const FN_chatBox_onEscape As Long = 14802                ' chatBox->setOnEscapeHandler([](W_EDITBOX& widg) {...})
Public Const FN_chatBox_onEscapeTask As Long = 14803            ' widgScheduleTask([weakParent]() {...})
Public Const FN_chatBox_onEditingStopped As Long = 14804        ' chatBox->setOnEditingStoppedHandler([](W_EDITBOX& widg) {...})
Public Const FN_chatBox_onEditingStoppedTask As Long = 14805    ' widgScheduleTask([weakParent]() {...})
Public Const FN_chatOptionsButton_onClick As Long = 14806       ' optionsButton->addOnClickHandler([weakSelf](W_BUTTON& widg) {...})
Public Const FN_chatOptionsButton_task As Long = 14807          ' widgScheduleTask([weakSelf]() {...})
Public Const FN_chatOptionsOverlay_clicked As Long = 14808      ' newRootFrm->onClickedFunc = [psWeakOptionsOverlayScreen, psWeakChatBoxWidget]() {...}
Public Const FN_chatOptions_onSettingsChange As Long = 14809    ' createChatOptionsForm(NetPlay.isHost, []() {...})
Public Const FN_chatOptionsForm_calcLayout As Long = 14810      ' chatOptionsForm->setCalcLayout([psWeakChatBoxWidget](WIDGET* psWidget) {...})
Public Const FN_chatClickForm_calcLayout As Long = 14811        ' WzInGameChatScreen_CLICKFORM::make, sInit.calcLayout
Public Const FN_chatClickForm_onTab As Long = 14812             ' [weakSelf](W_EDITBOX&) -> bool
Public Const FN_chatClickForm_onTabTask As Long = 14813         ' widgScheduleTask([weakSelf]() {...})
Public Const FN_chatConsoleBox_calcLayout As Long = 14814       ' consoleBox->setCalcLayout(LAMBDA_CALCLAYOUT_SIMPLE({...}))
Public Const FN_chatQuickChat_onSent As Long = 14815            ' createQuickChatForm(..., [weakSelf]() {...})
Public Const FN_chatQuickChat_onSentTask As Long = 14816        ' widgScheduleTask([weakSelf]() {...})
Public Const FN_chatQuickChatForm_calcLayout As Long = 14817    ' quickChatForm->setCalcLayout(LAMBDA_CALCLAYOUT_SIMPLE({...}))
Public Const FN_chatScreen_clicked As Long = 14818              ' WzInGameChatScreen::make, [psWeakHelpOverlayScreen]() {...}

' MARK: - Globals

Private psCurrentChatScreen As W_SCREEN  ' static std::weak_ptr<WzInGameChatScreen> psCurrentChatScreen;
Private chatScreens As Collection       ' (VB) the WzInGameChatScreens' OnCloseFunc: "p<ObjPtr>" -> WzFunc (or Nothing)

'--------------------------------------------------------------------------
' (VB) helpers

' a PIELIGHT's alpha (its byte.a: the top byte)
Private Function pielightAlpha(ByVal c As Long) As Long
    pielightAlpha = SHRU(c, 24)
End Function

Private Function pielightWithAlpha(ByVal c As Long, ByVal a As Long) As Long
    pielightWithAlpha = (c And &HFFFFFF) Or SHL(a And &HFF&, 24)
End Function

Private Function chatScreens_has(ByVal p As Long) As Boolean
    Dim v As Object
    If chatScreens Is Nothing Or p = 0 Then Exit Function
    On Error GoTo missing
    Set v = chatScreens.Item("p" & p)    ' (VB) Set: the items are WzFuncs or Nothing (a Let would ask for a default member)
    chatScreens_has = True
    Exit Function
missing:
    chatScreens_has = False
End Function

' std::dynamic_pointer_cast<WzInGameChatScreen>(screen) != nullptr
Private Function isWzInGameChatScreen(ByVal screen As W_SCREEN) As Boolean
    If screen Is Nothing Then Exit Function
    isWzInGameChatScreen = chatScreens_has(ObjPtr(screen))
End Function

' MARK: - WzInGameChatScreen definition

' void WzInGameChatScreen::closeScreen()
Private Sub WzInGameChatScreen_closeScreen(ByVal self As W_SCREEN)
    Dim onCloseFunc As WzFunc, p As Long
    p = ObjPtr(self)
    If chatScreens_has(p) Then Set onCloseFunc = chatScreens.Item("p" & p)
    shutdownChatScreen
    If Not onCloseFunc Is Nothing Then
        WzFunc_call onCloseFunc
    End If
End Sub

' MARK: - WzInGameChatBoxForm

' WzInGameChatBoxButton

Public Sub WzInGameChatBoxButton_construct(ByVal self As WIDGET)
    W_BUTTON_construct self
    self.kind = K_WzInGameChatBoxButton ' (VB) virtual calls from here on reach WzInGameChatBoxButton's overrides
End Sub

' std::make_shared<WzInGameChatBoxButton>()
Public Function WzInGameChatBoxButton_new() As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_WzInGameChatBoxButton)
    WzInGameChatBoxButton_construct widget
    Set WzInGameChatBoxButton_new = widget
End Function

Public Sub WzInGameChatBoxButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As WzInGameChatBoxButton, x0 As Long, y0 As Long, x1 As Long, y1 As Long, haveText As Boolean
    Dim isDown As Boolean, isDisabled As Boolean, isHighlight As Boolean, border_color As Long, fill_color As Long
    Dim fw As Long, fx As Long, fy As Long, textColor As Long
    Set d = WzInGameChatBoxButton_(self)
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    x1 = x0 + self.width()
    y1 = y0 + self.height()

    haveText = Len(W_BUTTON_(self).pText) <> 0

    isDown = (self.getState() And (WBUT_DOWN Or WBUT_LOCK Or WBUT_CLICKLOCK)) <> 0
    isDisabled = (self.getState() And WBUT_DISABLE) <> 0
    isHighlight = Not isDisabled And ((self.getState() And WBUT_HIGHLIGHT) <> 0)

    ' Display the button.
    If Not isDisabled Then border_color = pal_RGBA(100, 100, 255, 120) Else border_color = psPalette(WZCOL_FORM_DARK)
    If isDown Or isDisabled Then fill_color = pal_RGBA(0, 0, 45, 250) Else fill_color = pal_RGBA(0, 0, 67, 220)
    iV_ShadowBox x0, y0, x1, y1, 0, border_color, border_color, fill_color
    If isHighlight Then
        iV_Box x0 + 2, y0 + 2, x1 - 2, y1 - 2, psPalette(WZCOL_FORM_HILITE)
    End If

    If haveText Then
        d.text.setText W_BUTTON_(self).pText, W_BUTTON_(self).FontID
        fw = d.text.width()
        fx = x0 + (self.width() - fw) \ 2
        fy = y0 + (self.height() - d.text.lineSize()) \ 2 - d.text.aboveBase()
        textColor = psPalette(WZCOL_FORM_TEXT)
        If isDisabled Then
            textColor = pielightWithAlpha(textColor, pielightAlpha(textColor) \ 2)
        End If
        d.text.render fx, fy, textColor
    End If
End Sub

' WzInGameChatBoxForm

' WzInGameChatBoxForm(): IntFormAnimated(true) {}
Public Sub WzInGameChatBoxForm_construct(ByVal self As WIDGET)
    IntFormAnimated_construct self, True
    self.kind = K_WzInGameChatBoxForm ' (VB) virtual calls from here on reach WzInGameChatBoxForm's overrides
End Sub

' (onTabHandler: Nothing = an empty W_EDITBOX::OnTabHandler)
Public Function WzInGameChatBoxForm_make(ByVal initialChatMode As WzChatMode, ByVal onTabHandler As WzFunc) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_WzInGameChatBoxForm)
    WzInGameChatBoxForm_construct widget
    WzInGameChatBoxForm_initialize widget, initialChatMode, onTabHandler
    Set WzInGameChatBoxForm_make = widget
End Function

Public Sub WzInGameChatBoxForm_setChatBoxEnabledStatus(ByVal self As WIDGET, ByVal freeChatEnabled As Boolean, ByVal quickChatEnabled As Boolean)
    Dim d As WzInGameChatBoxForm
    Set d = WzInGameChatBoxForm_(self)
    If freeChatEnabled Then
        W_EDITBOX_setPlaceholder d.chatBox, ""
        d.chatBox.setTip ""
        W_EDITBOX_setPlaceholderTextColor d.chatBox, False
    Else
        If quickChatEnabled Then
            W_EDITBOX_setPlaceholder d.chatBox, "Use Quick Chat to chat with other players."
            W_EDITBOX_setPlaceholderTextColor d.chatBox, True, psPalette(WZCOL_TEXT_MEDIUM)
            d.chatBox.setTip "The host has disabled free chat. Please use Quick Chat."
        Else
            W_EDITBOX_setPlaceholder d.chatBox, "The host has disabled free chat."
            d.chatBox.setTip ""
            W_EDITBOX_setPlaceholderTextColor d.chatBox, True, psPalette(WZCOL_TEXT_MEDIUM)
        End If
        ' but don't disable the edit box, so cheat / chat commands still work
    End If
End Sub

Private Sub WzInGameChatBoxForm_closeParentScreen(ByVal self As WIDGET)
    Dim parentScreen As W_SCREEN
    Set parentScreen = self.screenPointer
    If Not isWzInGameChatScreen(parentScreen) Then Set parentScreen = Nothing
    If Not (Not parentScreen Is Nothing) Then ASSERT_FAILURE "parentScreen != nullptr", "src/screens/chatscreen.cpp:173", "closeParentScreen", "No screen?": Exit Sub
    WzInGameChatScreen_closeScreen parentScreen
End Sub

' WzInGameChatBoxForm::~WzInGameChatBoxForm()
Public Sub WzInGameChatBoxForm_destruct(ByVal self As WIDGET)
    Dim d As WzInGameChatBoxForm
    Set d = WzInGameChatBoxForm_(self)
    If Not d.optionsOverlayScreen Is Nothing Then
        widgRemoveOverlayScreen d.optionsOverlayScreen
        Set d.optionsOverlayScreen = Nothing
    End If
End Sub

Public Sub WzInGameChatBoxForm_focusLost(ByVal self As WIDGET)
    debug_ 199, LOG_INFO, "focusLost", "WzInGameChatBoxForm::focusLost"
End Sub

Public Sub WzInGameChatBoxForm_startChatBoxEditing(ByVal self As WIDGET)
    Dim d As WzInGameChatBoxForm, sContext As W_CONTEXT
    Set d = WzInGameChatBoxForm_(self)
    ' Auto-click it
    ' (W_CONTEXT sContext = W_CONTEXT::ZeroContext(): a new W_CONTEXT is zeroed)
    d.chatBox.clicked sContext
    Select Case d.chatMode
    Case WzChatMode_Glob
        W_EDITBOX_setBoxColours d.chatBox, psPalette(WZCOL_MENU_BORDER), psPalette(WZCOL_MENU_BORDER), psPalette(WZCOL_MENU_BACKGROUND)
        W_LABEL_setFontColour d.label, psPalette(WZCOL_TEXT_BRIGHT)
    Case WzChatMode_Team
        W_EDITBOX_setBoxColours d.chatBox, psPalette(WZCOL_YELLOW), psPalette(WZCOL_YELLOW), psPalette(WZCOL_MENU_BACKGROUND)
        W_LABEL_setFontColour d.label, psPalette(WZCOL_YELLOW)
    End Select
End Sub

Public Sub WzInGameChatBoxForm_endChatBoxEditing(ByVal self As WIDGET)
    Dim d As WzInGameChatBoxForm
    Set d = WzInGameChatBoxForm_(self)
    W_EDITBOX_stopEditing d.chatBox
    W_LABEL_setFontColour d.label, psPalette(WZCOL_TEXT_MEDIUM)
    W_EDITBOX_setBoxColours d.chatBox, psPalette(WZCOL_MENU_BORDER), psPalette(WZCOL_MENU_BORDER), psPalette(WZCOL_MENU_BACKGROUND)
End Sub

Public Sub WzInGameChatBoxForm_setChatMode(ByVal self As WIDGET, ByVal newMode As WzChatMode)
    Dim d As WzInGameChatBoxForm
    Set d = WzInGameChatBoxForm_(self)
    d.chatMode = newMode

    Select Case d.chatMode
    Case WzChatMode_Glob
        W_EDITBOX_setBoxColours d.chatBox, psPalette(WZCOL_MENU_BORDER), psPalette(WZCOL_MENU_BORDER), psPalette(WZCOL_MENU_BACKGROUND)
        W_LABEL_setFontColour d.label, psPalette(WZCOL_TEXT_BRIGHT)
        d.label.setString "Chat: All"
    Case WzChatMode_Team
        W_EDITBOX_setBoxColours d.chatBox, psPalette(WZCOL_YELLOW), psPalette(WZCOL_YELLOW), psPalette(WZCOL_MENU_BACKGROUND)
        W_LABEL_setFontColour d.label, psPalette(WZCOL_YELLOW)
        d.label.setString "Chat: Team"
    End Select
End Sub

'*
' * Parse what the player types in the chat box.
' *
' * Messages prefixed with "." are interpreted as messages to allies.
' * Messages prefixed with 1 or more digits (0-9) are interpreted as private messages to players.
' *
' * Examples:
' * - ".hi allies" sends "hi allies" to all allies.
' * - "123hi there" sends "hi there" to players 1, 2 and 3.
' * - ".123hi there" sends "hi there" to allies and players 1, 2 and 3.
' *
' (message.text: the char pointer moved on is the text without its first characters)
Private Sub parseChatMessageModifiers(message As InGameChatMessage)
    Dim c As Long
    If Left$(message.text, 1) = "." Then
        message.toAllies = True
        message.text = Mid$(message.text, 2)
    End If

    Do While Len(message.text) > 0   ' for each 0..9 numeric char encountered
        c = AscW(message.text)
        If c < 48 Or c > 57 Then Exit Do
        InGameChatMessage_addReceiverByPosition message, c - 48
        message.text = Mid$(message.text, 2)
    Loop
End Sub

Public Sub WzInGameChatBoxForm_geometryChanged(ByVal self As WIDGET)
    Dim d As WzInGameChatBoxForm, remainingWidth As Long, usableHeight As Long, optionsButtonHeight As Long
    Dim optionsButtonWidth As Long, optionsButtonY0 As Long, chatBoxPaddingX As Long, chatBoxWidth As Long, a As Long
    Set d = WzInGameChatBoxForm_(self)
    remainingWidth = self.width()

    usableHeight = self.height() - 4

    d.label.setGeometry 2, 2, 60, 16
    remainingWidth = remainingWidth - (d.label.x() + d.label.width())

    If Not d.optionsButton Is Nothing Then
        ' std::max({optionsButton->height(), usableHeight, iV_GetTextLineSize(optionsButton->FontID)})
        optionsButtonHeight = d.optionsButton.height()
        If usableHeight > optionsButtonHeight Then optionsButtonHeight = usableHeight
        a = iV_GetTextLineSize(W_BUTTON_(d.optionsButton).FontID)
        If a > optionsButtonHeight Then optionsButtonHeight = a
        optionsButtonWidth = d.optionsButton.idealWidth()
        If optionsButtonHeight > optionsButtonWidth Then optionsButtonWidth = optionsButtonHeight
        optionsButtonY0 = (self.height() - optionsButtonHeight) \ 2
        If optionsButtonY0 < 0 Then optionsButtonY0 = 0
        d.optionsButton.setGeometry self.width() - 2 - optionsButtonWidth, optionsButtonY0, optionsButtonWidth, optionsButtonHeight
        remainingWidth = remainingWidth - (optionsButtonWidth + 2 + 4)
    End If

    chatBoxPaddingX = 18
    chatBoxWidth = remainingWidth - chatBoxPaddingX
    If chatBoxWidth < 1 Then chatBoxWidth = 1
    d.chatBox.setGeometry d.label.x() + d.label.width() + chatBoxPaddingX, 2, chatBoxWidth, 16
End Sub

Private Sub WzInGameChatBoxForm_initialize(ByVal self As WIDGET, ByVal initialChatMode As WzChatMode, ByVal onTabHandler As WzFunc)
    Dim d As WzInGameChatBoxForm, minButtonWidthForText As Long, weakSelf As Long
    Set d = WzInGameChatBoxForm_(self)
    Set d.chatBox = W_EDITBOX_new()
    self.attach d.chatBox
    d.chatBox.id = CHAT_EDITBOX
    d.chatBox.setGeometry 80, 2, 320, 16

    W_EDITBOX_setOnReturnHandler d.chatBox, WzFunc_new(FN_chatBox_onReturn)

    W_EDITBOX_setOnEscapeHandler d.chatBox, WzFunc_new(FN_chatBox_onEscape)

    W_EDITBOX_setOnEditingStoppedHandler d.chatBox, WzFunc_new(FN_chatBox_onEditingStopped)

    If Not onTabHandler Is Nothing Then
        W_EDITBOX_setOnTabHandler d.chatBox, onTabHandler
    End If

    Set d.label = W_LABEL_new()
    self.attach d.label
    d.label.setGeometry 2, 2, 60, 16
    W_LABEL_setTextAlignment d.label, WLAB_ALIGNTOPLEFT

    If NetPlay_bComms Then
        Set d.optionsButton = WzInGameChatBoxButton_new()
        self.attach d.optionsButton
        d.optionsButton.setString ChrW$(&H2699) ' "U+2699": the gear
        d.optionsButton.setTip "Chat Options"
        W_BUTTON_(d.optionsButton).FontID = font_regular
        minButtonWidthForText = iV_GetTextWidth(W_BUTTON_(d.optionsButton).pText, W_BUTTON_(d.optionsButton).FontID)
        d.optionsButton.setGeometry 0, 0, minButtonWidthForText + 10, iV_GetTextLineSize(W_BUTTON_(d.optionsButton).FontID)
        weakSelf = ObjPtr(self)
        W_BUTTON_addOnClickHandler d.optionsButton, WzFunc_new(FN_chatOptionsButton_onClick, weakSelf)
    End If

    WzInGameChatBoxForm_setChatMode self, initialChatMode
End Sub

Private Sub WzInGameChatBoxForm_displayOptionsOverlay(ByVal self As WIDGET)
    Dim d As WzInGameChatBoxForm, lockedScreen As W_SCREEN, newRootFrm As WIDGET, psWeakOptionsOverlayScreen As Long
    Dim psWeakChatBoxWidget As Long, chatOptionsForm As WIDGET
    Set d = WzInGameChatBoxForm_(self)
    Set lockedScreen = self.screenPointer
    If Not (Not lockedScreen Is Nothing) Then ASSERT_FAILURE "lockedScreen != nullptr", "src/screens/chatscreen.cpp:415", "displayOptionsOverlay", "The WzInGameChatBoxForm does not have an associated screen pointer?"

    ' Initialize the options overlay screen
    Set d.optionsOverlayScreen = W_SCREEN_make()
    Set newRootFrm = W_FULLSCREENOVERLAY_CLICKFORM_make()
    psWeakOptionsOverlayScreen = ObjPtr(d.optionsOverlayScreen)
    psWeakChatBoxWidget = ObjPtr(self)
    Set W_FULLSCREENOVERLAY_CLICKFORM_(newRootFrm).onClickedFunc = WzFunc_new(FN_chatOptionsOverlay_clicked, psWeakOptionsOverlayScreen, psWeakChatBoxWidget)
    Set W_FULLSCREENOVERLAY_CLICKFORM_(newRootFrm).onCancelPressed = W_FULLSCREENOVERLAY_CLICKFORM_(newRootFrm).onClickedFunc
    d.optionsOverlayScreen.psForm.attach newRootFrm

    ' Add the Chat Options form to the overlay screen form
    Set chatOptionsForm = createChatOptionsForm(NetPlay_isHost, WzFunc_new(FN_chatOptions_onSettingsChange))
    newRootFrm.attach chatOptionsForm
    chatOptionsForm.setCalcLayout WzFunc_new(FN_chatOptionsForm_calcLayout, psWeakChatBoxWidget)

    widgRegisterOverlayScreenOnTopOfScreen d.optionsOverlayScreen, lockedScreen
End Sub

' MARK: - WzInGameChatScreen_CLICKFORM

' WzInGameChatScreen_CLICKFORM(W_FORMINIT const *init) : W_CLICKFORM(init) {}
Public Sub WzInGameChatScreen_CLICKFORM_construct_init(ByVal self As WIDGET, init As W_FORMINIT)
    W_CLICKFORM_construct_init self, init
    self.kind = K_WzInGameChatScreen_CLICKFORM ' (VB) virtual calls from here on reach WzInGameChatScreen_CLICKFORM's overrides
    WzInGameChatScreen_CLICKFORM_(self).backgroundColor = pal_RGBA(0, 0, 0, 80)
End Sub

' WzInGameChatScreen_CLICKFORM() : W_CLICKFORM() {}
Public Sub WzInGameChatScreen_CLICKFORM_construct(ByVal self As WIDGET)
    W_CLICKFORM_construct self
    self.kind = K_WzInGameChatScreen_CLICKFORM ' (VB) virtual calls from here on reach WzInGameChatScreen_CLICKFORM's overrides
    WzInGameChatScreen_CLICKFORM_(self).backgroundColor = pal_RGBA(0, 0, 0, 80)
End Sub

' (~WzInGameChatScreen_CLICKFORM(): empty)

Public Function WzInGameChatScreen_CLICKFORM_make(ByVal initialChatMode As WzChatMode, Optional ByVal formID As Long = 0) As WIDGET
    Dim sInit As W_FORMINIT, widget As WIDGET
    W_FORMINIT_construct sInit
    sInit.id = formID
    sInit.style = WFORM_PLAIN Or WFORM_CLICKABLE
    sInit.x = 0
    sInit.y = 0
    sInit.width = screenWidth - 1
    sInit.height = screenHeight - 1
    Set sInit.calcLayout = WzFunc_new(FN_chatClickForm_calcLayout)

    ' (class make_shared_enabler)
    Set widget = WIDGET_alloc(K_WzInGameChatScreen_CLICKFORM)
    WzInGameChatScreen_CLICKFORM_construct_init widget, sInit
    WzInGameChatScreen_CLICKFORM_initialize widget, initialChatMode
    Set WzInGameChatScreen_CLICKFORM_make = widget
End Function

Public Sub WzInGameChatScreen_CLICKFORM_geometryChanged(ByVal self As WIDGET)
    WzInGameChatScreen_CLICKFORM_recalcLayout self
End Sub

Public Sub WzInGameChatScreen_CLICKFORM_giveQuickChatFocus(ByVal self As WIDGET)
    Dim d As WzInGameChatScreen_CLICKFORM, context As W_CONTEXT
    Set d = WzInGameChatScreen_CLICKFORM_(self)
    If d.quickChatForm Is Nothing Then
        WzInGameChatScreen_CLICKFORM_giveChatBoxFocus self
        Exit Sub
    End If
    If Not d.consoleBox Is Nothing Then
        WzInGameChatBoxForm_endChatBoxEditing d.consoleBox
    End If
    ' Set the focus to the quickChatForm by simulating a click on it
    ' (W_CONTEXT context = W_CONTEXT::ZeroContext(): a new W_CONTEXT is zeroed)
    context.mx = d.quickChatForm.screenPosX()
    context.my = d.quickChatForm.screenPosY()
    d.quickChatForm.clicked context, WKEY_PRIMARY
    d.quickChatForm.released context, WKEY_PRIMARY
End Sub

Public Function WzInGameChatScreen_CLICKFORM_giveChatBoxFocus(ByVal self As WIDGET) As Boolean
    Dim d As WzInGameChatScreen_CLICKFORM
    Set d = WzInGameChatScreen_CLICKFORM_(self)
    If d.consoleBox Is Nothing Then
        WzInGameChatScreen_CLICKFORM_giveChatBoxFocus = False
        Exit Function
    End If

    WzInGameChatBoxForm_startChatBoxEditing d.consoleBox
    WzInGameChatScreen_CLICKFORM_giveChatBoxFocus = True
End Function

Private Sub WzInGameChatScreen_CLICKFORM_initialize(ByVal self As WIDGET, ByVal initialChatMode As WzChatMode)
    Dim d As WzInGameChatScreen_CLICKFORM, freeChatEnabled As Boolean, quickChatEnabled As Boolean, weakSelf As Long
    Set d = WzInGameChatScreen_CLICKFORM_(self)
    ' (VB) the C++'s || and && stop early: ingame.hostChatPermissions(selectedPlayer) is read only for a connected player
    If Not bMultiPlayer Then
        freeChatEnabled = True
    ElseIf Not ULT(selectedPlayer, MAX_CONNECTED_PLAYERS) Then
        freeChatEnabled = True
    Else
        freeChatEnabled = ingame.hostChatPermissions(selectedPlayer)
    End If
    ' for now, quick chat is only available in skirmish / multiplayer mode, for players
    quickChatEnabled = False
    If bMultiPlayer Then
        If ULT(selectedPlayer, MAX_PLAYERS) Then
            quickChatEnabled = Not NetPlay_players_isSpectator(selectedPlayer)
        End If
    End If

    weakSelf = ObjPtr(self)
    Set d.consoleBox = WzInGameChatBoxForm_make(initialChatMode, WzFunc_new(FN_chatClickForm_onTab, weakSelf))
    self.attach d.consoleBox
    d.consoleBox.id = CHAT_CONSOLEBOX
    d.consoleBox.setCalcLayout WzFunc_new(FN_chatConsoleBox_calcLayout)
    WzInGameChatBoxForm_setChatBoxEnabledStatus d.consoleBox, freeChatEnabled, quickChatEnabled

    ' for now, quick chat is only available in skirmish / multiplayer mode, for players
    If quickChatEnabled Then
        ' (createQuickChatForm(WzQuickChatContext::InGame, [weakSelf]() {...}, (initialChatMode == WzChatMode::Team) ?
        ' WzQuickChatMode::Team : WzQuickChatMode::Global): the stand-in takes no mode)
        Set d.quickChatForm = createQuickChatForm(WzQuickChatContext_InGame, WzFunc_new(FN_chatQuickChat_onSent, weakSelf))
        self.attach d.quickChatForm
        d.quickChatForm.setCalcLayout WzFunc_new(FN_chatQuickChatForm_calcLayout)
    End If

    WzInGameChatScreen_CLICKFORM_recalcLayout self
End Sub

Private Sub WzInGameChatScreen_CLICKFORM_recalcLayout(ByVal self As WIDGET)
    Dim d As WzInGameChatScreen_CLICKFORM
    Set d = WzInGameChatScreen_CLICKFORM_(self)
    If Not d.consoleBox Is Nothing Then
        d.consoleBox.callCalcLayout
    End If
End Sub

Private Sub WzInGameChatScreen_CLICKFORM_clearFocusAndEditing(ByVal self As WIDGET)
    Dim d As WzInGameChatScreen_CLICKFORM, lockedScreen As W_SCREEN
    Set d = WzInGameChatScreen_CLICKFORM_(self)
    Set lockedScreen = self.screenPointer
    If Not lockedScreen Is Nothing Then
        lockedScreen.setFocus Nothing
    End If
    If Not d.consoleBox Is Nothing Then
        WzInGameChatBoxForm_endChatBoxEditing d.consoleBox
    End If
End Sub

Private Sub WzInGameChatScreen_CLICKFORM_handleClickOnForm(ByVal self As WIDGET)
    Dim d As WzInGameChatScreen_CLICKFORM
    Set d = WzInGameChatScreen_CLICKFORM_(self)
    WzInGameChatScreen_CLICKFORM_clearFocusAndEditing self
    If Not d.onClickedFunc Is Nothing Then
        WzFunc_call d.onClickedFunc
    End If
End Sub

Public Sub WzInGameChatScreen_CLICKFORM_clicked(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    WzInGameChatScreen_CLICKFORM_handleClickOnForm self
End Sub

Public Sub WzInGameChatScreen_CLICKFORM_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As WzInGameChatScreen_CLICKFORM, x0 As Long, y0 As Long
    Set d = WzInGameChatScreen_CLICKFORM_(self)
    If Not self.visible() Then
        Exit Sub ' skip if hidden
    End If

    If pielightAlpha(d.backgroundColor) = 0 Then   ' backgroundColor.isTransparent()
        Exit Sub
    End If

    x0 = self.x() + xOffset
    y0 = self.y() + yOffset

    ' draw background over everything
    pie_UniTransBoxFill x0, y0, x0 + self.width(), y0 + self.height(), d.backgroundColor
End Sub

Public Sub WzInGameChatScreen_CLICKFORM_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    Dim d As WzInGameChatScreen_CLICKFORM, currFocusOnChatbox As Boolean, lockedScreen As W_SCREEN, widgetWithFocus As WIDGET
    Set d = WzInGameChatScreen_CLICKFORM_(self)
    If keyPressed(KEY_ESC) Then
        WzInGameChatScreen_CLICKFORM_clearFocusAndEditing self
        If Not d.onCancelPressed Is Nothing Then
            WzFunc_call d.onCancelPressed
        End If
    End If

    If keyPressed(KEY_TAB) Then
        ' no child widget consumed the TAB keypress - swap focus between quick chat and chat box
        currFocusOnChatbox = False
        Set lockedScreen = self.screenPointer
        If Not lockedScreen Is Nothing Then
            Set widgetWithFocus = lockedScreen.getWidgetWithFocus()
            If Not widgetWithFocus Is Nothing Then
                If Not d.consoleBox Is Nothing Then
                    If widgetWithFocus Is d.consoleBox Then
                        currFocusOnChatbox = True
                    ElseIf widgetWithFocus.hasAncestor(d.consoleBox) Then
                        currFocusOnChatbox = True
                    End If
                End If
            End If
        End If

        If currFocusOnChatbox Then
            WzInGameChatScreen_CLICKFORM_giveQuickChatFocus self
        Else
            WzInGameChatScreen_CLICKFORM_giveChatBoxFocus self
        End If
    End If

    ' while this is displayed, clear the input buffer
    ' (ensuring that subsequent screens / the main screen do not get the input to handle)
    inputLoseFocus
End Sub

' MARK: - WzGameStartOverlayScreen

' std::shared_ptr<WzInGameChatScreen> WzInGameChatScreen::make(const OnCloseFunc& _onCloseFunc, WzChatMode initialChatMode,
' bool startWithQuickChatFocused) (_onCloseFunc: Nothing = empty)
Private Function WzInGameChatScreen_make(ByVal onCloseFunc As WzFunc, ByVal initialChatMode As WzChatMode, ByVal startWithQuickChatFocused As Boolean) As W_SCREEN
    Dim newRootFrm As WIDGET, screen As W_SCREEN, psWeakHelpOverlayScreen As Long, chatBoxEnabled As Boolean
    ' (class make_shared_enabler: public WzInGameChatScreen {};)
    Set newRootFrm = WzInGameChatScreen_CLICKFORM_make(initialChatMode)
    Set screen = W_SCREEN_make(newRootFrm)  ' screen->initialize(newRootFrm);
    ' screen->onCloseFunc = _onCloseFunc;
    If chatScreens Is Nothing Then Set chatScreens = New Collection
    If chatScreens_has(ObjPtr(screen)) Then chatScreens.Remove "p" & ObjPtr(screen)  ' (a dead screen's address)
    chatScreens.Add onCloseFunc, "p" & ObjPtr(screen)
    psWeakHelpOverlayScreen = ObjPtr(screen)
    Set WzInGameChatScreen_CLICKFORM_(newRootFrm).onClickedFunc = WzFunc_new(FN_chatScreen_clicked, psWeakHelpOverlayScreen)
    Set WzInGameChatScreen_CLICKFORM_(newRootFrm).onCancelPressed = WzInGameChatScreen_CLICKFORM_(newRootFrm).onClickedFunc

    ' must select default element focus *after* adding the root form to the screen
    If startWithQuickChatFocused Then
        WzInGameChatScreen_CLICKFORM_giveQuickChatFocus newRootFrm
    Else
        ' (the C++'s || stops early)
        If Not bMultiPlayer Then
            chatBoxEnabled = True
        ElseIf Not ULT(selectedPlayer, MAX_CONNECTED_PLAYERS) Then
            chatBoxEnabled = True
        Else
            chatBoxEnabled = ingame.hostChatPermissions(selectedPlayer)
        End If
        If chatBoxEnabled Then
            WzInGameChatScreen_CLICKFORM_giveChatBoxFocus newRootFrm
        Else
            WzInGameChatScreen_CLICKFORM_giveQuickChatFocus newRootFrm
        End If
    End If

    Set WzInGameChatScreen_make = screen
End Function

' std::shared_ptr<W_SCREEN> createChatScreen(std::function<void ()> onCloseFunc, WzChatMode initialChatMode, bool
' startWithQuickChatFocused) (VB: a Sub, as its stand-in; the screen is psCurrentChatScreen)
Public Sub createChatScreen(ByVal onCloseFunc As WzFunc, ByVal initialChatMode As WzChatMode, ByVal startWithQuickChatFocused As Boolean)
    Dim screen As W_SCREEN
    Set screen = WzInGameChatScreen_make(onCloseFunc, initialChatMode, startWithQuickChatFocused)
    widgRegisterOverlayScreenOnTopOfScreen screen, psWScreen
    Set psCurrentChatScreen = screen
End Sub

Public Sub shutdownChatScreen()
    Dim strongGameStartScreen As W_SCREEN
    Set strongGameStartScreen = psCurrentChatScreen
    If Not strongGameStartScreen Is Nothing Then
        widgRemoveOverlayScreen strongGameStartScreen
        Set psCurrentChatScreen = Nothing
        If chatScreens_has(ObjPtr(strongGameStartScreen)) Then chatScreens.Remove "p" & ObjPtr(strongGameStartScreen)
    End If
End Sub

' the lambdas
Public Function chatscreen_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim widg As WIDGET, strongParent As WIDGET, enteredText As String, currChatMode As WzChatMode, weakParent As Long
    Dim message As InGameChatMessage, processedCheatCode As Boolean, skipSendingMessage As Boolean, parent As WIDGET
    Dim strongParentForm As WIDGET, psOverlayScreen As W_SCREEN, strongChatBoxWidget As WIDGET, psWidget As WIDGET
    Dim psChatBoxWidget As WIDGET, chatBoxScreenPosX As Long, chatBoxScreenWidth As Long, idealChatOptionsWidth As Long
    Dim maxWidth As Long, w As Long, x0 As Long, idealChatOptionsHeight As Long, y0 As Long, maxHeight As Long
    Dim lastIdealChatOptionsHeight As Long, desiredWidth As Long, parentScreen As W_SCREEN, quickChatWidth As Long
    Const MAX_QUICKCHAT_WIDTH As Long = 1200
    Select Case f.id
    Case FN_chatBox_onReturn                    ' [](W_EDITBOX& widg)
        Set widg = args(0)
        Set strongParent = widg.parent()
        If Not WIDGET_isKindOf(strongParent, K_WzInGameChatBoxForm) Then Set strongParent = Nothing
        If Not (Not strongParent Is Nothing) Then ASSERT_FAILURE "strongParent != nullptr", "src/screens/chatscreen.cpp:304", "initialize", "No parent?": Exit Function

        enteredText = widg.getString()
        currChatMode = WzInGameChatBoxForm_(strongParent).chatMode
        inputLoseFocus

        weakParent = ObjPtr(strongParent)
        widgScheduleTask WzFunc_new(FN_chatBox_onReturnTask, enteredText, currChatMode, weakParent)

    Case FN_chatBox_onReturnTask                ' [enteredText, currChatMode, weakParent]()
        ' must process the enteredText inside a scheduled task as side effects of attemptCheatCode may modify the overlay screens / widgets

        ' (const char* pStr = enteredText.toUtf8().c_str();)
        message = InGameChatMessage_new(selectedPlayer, f.cap(0))
        processedCheatCode = attemptCheatCode(message.text)       ' parse the message

        ' (the C++'s || stops early: hostChatPermissions is read only for a connected player)
        skipSendingMessage = False
        If bMultiPlayer Then
            If Not ULT(selectedPlayer, MAX_CONNECTED_PLAYERS) Then
                skipSendingMessage = True
            ElseIf Not ingame.hostChatPermissions(selectedPlayer) Then
                skipSendingMessage = True
            End If
        End If

        If Not skipSendingMessage Then
            Select Case f.cap(1)
            Case WzChatMode_Glob
                parseChatMessageModifiers message
            Case WzChatMode_Team
                message.toAllies = True
            End Select

            If Len(message.text) <> 0 Then
                InGameChatMessage_send message
            End If
        Else
            If Len(message.text) <> 0 And Not processedCheatCode And ULT(selectedPlayer, MAX_CONNECTED_PLAYERS) Then
                addConsoleMessage "Did not send message - free chat is disabled by host. Please use Quick Chat.", DEFAULT_JUSTIFY, NOTIFY_MESSAGE, False
            End If
        End If

        Set parent = WIDGET_fromPtr(f.cap(2))
        If Not (Not parent Is Nothing) Then ASSERT_FAILURE "parent != nullptr", "src/screens/chatscreen.cpp:347", "initialize", "No parent?": Exit Function
        WzInGameChatBoxForm_endChatBoxEditing parent
        WzInGameChatBoxForm_closeParentScreen parent

    Case FN_chatBox_onEscape                    ' [](W_EDITBOX& widg)
        Set widg = args(0)
        Set strongParent = widg.parent()
        If Not WIDGET_isKindOf(strongParent, K_WzInGameChatBoxForm) Then Set strongParent = Nothing
        If Not (Not strongParent Is Nothing) Then ASSERT_FAILURE "strongParent != nullptr", "src/screens/chatscreen.cpp:355", "initialize", "No parent?": Exit Function

        inputLoseFocus

        weakParent = ObjPtr(strongParent)
        widgScheduleTask WzFunc_new(FN_chatBox_onEscapeTask, weakParent)

    Case FN_chatBox_onEscapeTask                ' [weakParent]()
        Set parent = WIDGET_fromPtr(f.cap(0))
        If Not (Not parent Is Nothing) Then ASSERT_FAILURE "parent != nullptr", "src/screens/chatscreen.cpp:362", "initialize", "No parent?": Exit Function
        WzInGameChatBoxForm_closeParentScreen parent

    Case FN_chatBox_onEditingStopped            ' [](W_EDITBOX& widg)
        Set widg = args(0)
        Set strongParent = widg.parent()
        If Not WIDGET_isKindOf(strongParent, K_WzInGameChatBoxForm) Then Set strongParent = Nothing
        If Not (Not strongParent Is Nothing) Then ASSERT_FAILURE "strongParent != nullptr", "src/screens/chatscreen.cpp:369", "initialize", "No parent?": Exit Function

        weakParent = ObjPtr(strongParent)
        widgScheduleTask WzFunc_new(FN_chatBox_onEditingStoppedTask, weakParent)

    Case FN_chatBox_onEditingStoppedTask        ' [weakParent]()
        Set parent = WIDGET_fromPtr(f.cap(0))
        If Not (Not parent Is Nothing) Then ASSERT_FAILURE "parent != nullptr", "src/screens/chatscreen.cpp:374", "initialize", "No parent?": Exit Function
        WzInGameChatBoxForm_endChatBoxEditing parent

    Case FN_chatOptionsButton_onClick           ' [weakSelf](W_BUTTON& widg)
        widgScheduleTask WzFunc_new(FN_chatOptionsButton_task, f.cap(0))

    Case FN_chatOptionsButton_task              ' [weakSelf]()
        Set strongParentForm = WIDGET_fromPtr(f.cap(0))
        If Not (Not strongParentForm Is Nothing) Then ASSERT_FAILURE "strongParentForm != nullptr", "src/screens/chatscreen.cpp:402", "initialize", "No parent form": Exit Function
        ' Open the Chat Options overlay
        WzInGameChatBoxForm_displayOptionsOverlay strongParentForm

    Case FN_chatOptionsOverlay_clicked          ' [psWeakOptionsOverlayScreen, psWeakChatBoxWidget]()
        Set psOverlayScreen = W_SCREEN_fromPtr(f.cap(0))
        If Not psOverlayScreen Is Nothing Then
            widgRemoveOverlayScreen psOverlayScreen
        End If
        ' Destroy Options overlay / overlay screen
        Set strongChatBoxWidget = WIDGET_fromPtr(f.cap(1))
        If Not strongChatBoxWidget Is Nothing Then
            Set WzInGameChatBoxForm_(strongChatBoxWidget).optionsOverlayScreen = Nothing
        End If

    Case FN_chatOptions_onSettingsChange        ' []()
        ' FUTURE TODO: could update the visible chat screen UI if change impacted selectedPlayer?

    Case FN_chatOptionsForm_calcLayout          ' [psWeakChatBoxWidget](WIDGET* psWidget)
        Set psWidget = args(0)
        Set psChatBoxWidget = WIDGET_fromPtr(f.cap(0))
        If Not (Not psChatBoxWidget Is Nothing) Then ASSERT_FAILURE "psChatBoxWidget != nullptr", "src/screens/chatscreen.cpp:443", "displayOptionsOverlay", "ChatBoxWidget is null": Exit Function
        chatBoxScreenPosX = psChatBoxWidget.screenPosX()
        chatBoxScreenWidth = psChatBoxWidget.width()
        idealChatOptionsWidth = psWidget.idealWidth()
        maxWidth = screenWidth - 40
        If maxWidth < idealChatOptionsWidth Then w = maxWidth Else w = idealChatOptionsWidth
        x0 = chatBoxScreenPosX + (chatBoxScreenWidth - w) \ 2
        ' Try to position it below the chat box widget
        idealChatOptionsHeight = psWidget.idealHeight()
        y0 = psChatBoxWidget.screenPosY() + psChatBoxWidget.height() + 20
        If y0 < 20 Then y0 = 20
        maxHeight = screenHeight - y0 - 20
        If maxHeight < idealChatOptionsHeight Then psWidget.setGeometry x0, y0, w, maxHeight Else psWidget.setGeometry x0, y0, w, idealChatOptionsHeight
        ' idealHeight may have changed due to the width change, so check and set again...
        lastIdealChatOptionsHeight = idealChatOptionsHeight
        idealChatOptionsHeight = psWidget.idealHeight()
        If lastIdealChatOptionsHeight <> idealChatOptionsHeight Then
            If maxHeight < idealChatOptionsHeight Then psWidget.setGeometry x0, y0, w, maxHeight Else psWidget.setGeometry x0, y0, w, idealChatOptionsHeight
        End If

    Case FN_chatClickForm_calcLayout            ' LAMBDA_CALCLAYOUT_SIMPLE
        Set psWidget = args(0)
        psWidget.setGeometry 0, 0, screenWidth, screenHeight

    Case FN_chatClickForm_onTab                 ' [weakSelf](W_EDITBOX&) -> bool
        ' on tab handler
        widgScheduleTask WzFunc_new(FN_chatClickForm_onTabTask, f.cap(0))
        chatscreen_c_fn = True

    Case FN_chatClickForm_onTabTask             ' [weakSelf]()
        Set strongParentForm = WIDGET_fromPtr(f.cap(0))
        If Not (Not strongParentForm Is Nothing) Then ASSERT_FAILURE "strongParentForm != nullptr", "src/screens/chatscreen.cpp:576", "initialize", "No parent form": Exit Function
        WzInGameChatScreen_CLICKFORM_giveQuickChatFocus strongParentForm

    Case FN_chatConsoleBox_calcLayout           ' LAMBDA_CALCLAYOUT_SIMPLE
        Set psWidget = args(0)
        desiredWidth = CHAT_CONSOLEBOXW
        x0 = (screenWidth - desiredWidth) \ 2
        psWidget.setGeometry x0, CHAT_CONSOLEBOXY, desiredWidth, CHAT_CONSOLEBOXH

    Case FN_chatQuickChat_onSent                ' [weakSelf]()
        ' on quick-chat sent
        widgScheduleTask WzFunc_new(FN_chatQuickChat_onSentTask, f.cap(0))

    Case FN_chatQuickChat_onSentTask            ' [weakSelf]()
        ' click on the form so we ultimately close the chat screen
        Set strongParentForm = WIDGET_fromPtr(f.cap(0))
        If Not (Not strongParentForm Is Nothing) Then ASSERT_FAILURE "strongParentForm != nullptr", "src/screens/chatscreen.cpp:598", "initialize", "No parent form": Exit Function
        Set parentScreen = strongParentForm.screenPointer
        If Not isWzInGameChatScreen(parentScreen) Then Set parentScreen = Nothing
        If Not (Not parentScreen Is Nothing) Then ASSERT_FAILURE "parentScreen != nullptr", "src/screens/chatscreen.cpp:600", "initialize", "No screen?": Exit Function
        WzInGameChatScreen_closeScreen parentScreen

    Case FN_chatQuickChatForm_calcLayout        ' LAMBDA_CALCLAYOUT_SIMPLE
        Set psWidget = args(0)
        quickChatWidth = screenWidth - 40
        If MAX_QUICKCHAT_WIDTH < quickChatWidth Then quickChatWidth = MAX_QUICKCHAT_WIDTH
        x0 = (screenWidth - quickChatWidth) \ 2
        If x0 < 20 Then x0 = 20
        psWidget.setGeometry x0, CHAT_CONSOLEBOXY + CHAT_CONSOLEBOXH + 100, quickChatWidth, psWidget.idealHeight()

    Case FN_chatScreen_clicked                  ' [psWeakHelpOverlayScreen]()
        Set psOverlayScreen = W_SCREEN_fromPtr(f.cap(0))
        If Not psOverlayScreen Is Nothing Then
            WzInGameChatScreen_closeScreen psOverlayScreen
        End If
    End Select
End Function
