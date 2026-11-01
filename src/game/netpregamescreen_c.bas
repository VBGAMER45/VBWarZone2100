Attribute VB_Name = "netpregamescreen_c"
' src/screens/netpregamescreen.cpp, netpregamescreen.h, whole: the multiplayer pre-game screen (an overlay screen
' over the loading game: "Waiting for other players ...", the team planning with teamstrategy_c's view, the final
' countdown, the players' loading status). Its widget classes' fields are the data classes named as the classes
' (WzGameStartOverlayScreen_CLICKFORM's: WzGameStartOverlayCFData.cls; struct CachedConsoleMessage:
' CachedConsoleMessage.cls). (VB) struct WzGameStartOverlayScreen : public W_SCREEN is a W_SCREEN (W_SCREEN_make with
' its root form) whose OnCloseFunc this module keeps ("p<ObjPtr>" in gameStartScreens), as chatscreen_c does;
' psCurrentGameStartScreen is a strong reference that shutdownGameStartScreen drops. The std::weak_ptr captures of
' widgets and screens are their addresses (WIDGET_fromPtr, W_SCREEN_fromPtr). The std::chrono::steady_clock
' time_points are QueryPerformanceCounter counts. Translations (_()) are not ported: the English text is used.
' (VB) widget class WzLoadingPlayersStatusForm : W_FORM
' (VB) widget class WzPlayerStatusCheckboxButton : W_BUTTON
' (VB) widget class WzCountdownLabel : WIDGET
' (VB) widget class WzGameStartOverlayScreen_CLICKFORM : W_CLICKFORM
Option Explicit

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Private Const MAX_CONNECTED_PLAYERS As Long = MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS

' enum class PregameWaitingState
Private Const PregameWaitingState_Uninitialized As Long = 0
Private Const PregameWaitingState_WaitingForPlayersToJoin As Long = 1
Private Const PregameWaitingState_WaitingForTeamPlanning As Long = 2
Private Const PregameWaitingState_WaitingForFinalCountdown As Long = 3

' WzPlayerStatusCheckboxButton's enum class LoadingStatus
Private Const LoadingStatus_Loading As Long = 0
Private Const LoadingStatus_Ready As Long = 1
Private Const LoadingStatus_Disconnected As Long = 2

' the lambdas (WzFunc: netpregamescreen_c_fn)
Public Const FN_pregame_titleLabel_calcLayout As Long = 17020       ' result->titleLabel->setCalcLayout(LAMBDA_CALCLAYOUT_SIMPLE({...}))
Public Const FN_pregame_playersList_calcLayout As Long = 17021      ' result->playersList->setCalcLayout(LAMBDA_CALCLAYOUT_SIMPLE({...}))
Public Const FN_pregame_spectatorsList_calcLayout As Long = 17022   ' result->spectatorsList->setCalcLayout(LAMBDA_CALCLAYOUT_SIMPLE({...}))
Public Const FN_pregame_optionsButton_onClick As Long = 17023       ' button->addOnClickHandler([weakPlayerStatusButton, onClickFunc](W_BUTTON& button) {...})
Public Const FN_pregame_kickButton_onClick As Long = 17024          ' createOptionsButton(kickButtonStr, [weakPlayerStatusButton](W_BUTTON& button){...})
Public Const FN_pregame_optionsOverlay_clicked As Long = 17025      ' newRootFrm->onClickedFunc = [psWeakOptionsOverlayScreen, psWeakPlayerStatusButton]() {...}
Public Const FN_pregame_optionsPopOver_calcLayout As Long = 17026   ' optionsPopOver->setCalcLayout([weakParent](WIDGET *psWidget) {...})
Public Const FN_pregame_clickForm_calcLayout As Long = 17027        ' WzGameStartOverlayScreen_CLICKFORM::make, sInit.calcLayout
Public Const FN_pregame_consoleMessage As Long = 17028              ' std::make_shared<CONSOLE_MESSAGE_LISTENER>([&](ConsoleMessage const &message) -> void {...}): cap(0) the form's address

' MARK: - Globals

Private psCurrentGameStartScreen As W_SCREEN    ' static std::weak_ptr<WzGameStartOverlayScreen> psCurrentGameStartScreen;
Private gameStartScreens As Collection          ' (VB) the WzGameStartOverlayScreens' OnCloseFunc: "p<ObjPtr>" -> WzFunc (or Nothing)

'--------------------------------------------------------------------------
' (VB) helpers

Private Const PLAYER_LOADING_STATUS_ITEM_SPACING As Long = 10
Private Const IndicatorSize As Long = 10
Private Const IndicatorRightPadding As Long = 5
Private Const WZ_TEAMPLANNING_COUNTDOWN_MS As Long = 20000
Private Const WZ_FINAL_COUNTDOWN_MS As Long = 3000
Private Const INTERNAL_PADDING As Long = 15

' std::chrono::duration_cast<std::chrono::milliseconds>(b - a).count() of two QueryPerformanceCounter counts
Private Function steadyMsBetween(ByVal a As Currency, ByVal b As Currency) As Double
    Static freq As Currency
    If freq = 0 Then QueryPerformanceFrequency freq
    steadyMsBetween = Int((b - a) * 1000@ / freq)
End Function

' std::chrono::steady_clock::now()
Private Function steadyNow() As Currency
    QueryPerformanceCounter steadyNow
End Function

' a PIELIGHT's alpha (its byte.a: the top byte)
Private Function pielightAlpha(ByVal c As Long) As Long
    pielightAlpha = SHRU(c, 24)
End Function

Private Function pielightWithAlpha(ByVal c As Long, ByVal a As Long) As Long
    pielightWithAlpha = (c And &HFFFFFF) Or SHL(a And &HFF&, 24)
End Function

Private Function gameStartScreens_has(ByVal p As Long) As Boolean
    Dim v As Object
    If gameStartScreens Is Nothing Or p = 0 Then Exit Function
    On Error GoTo missing
    Set v = gameStartScreens.Item("p" & p)   ' (VB) Set: the items are WzFuncs or Nothing (a Let would ask for a default member)
    gameStartScreens_has = True
    Exit Function
missing:
    gameStartScreens_has = False
End Function

' std::dynamic_pointer_cast<WzGameStartOverlayScreen>(screen) != nullptr
Private Function isWzGameStartOverlayScreen(ByVal screen As W_SCREEN) As Boolean
    If screen Is Nothing Then Exit Function
    isWzGameStartOverlayScreen = gameStartScreens_has(ObjPtr(screen))
End Function

' (VB) the parent cast of the calcLayout lambdas: psWidget->parent() as a WzLoadingPlayersStatusForm (or Nothing)
Private Function loadingStatusFormParent(ByVal psWidget As WIDGET) As WIDGET
    Dim p As WIDGET
    Set p = psWidget.parent()
    If Not p Is Nothing Then
        If WIDGET_isKindOf(p, K_WzLoadingPlayersStatusForm) Then Set loadingStatusFormParent = p
    End If
End Function

' MARK: - WzGameStartOverlayScreen definition

' void WzGameStartOverlayScreen::closeScreen()
Private Sub WzGameStartOverlayScreen_closeScreen(ByVal self As W_SCREEN)
    Dim onCloseFunc As WzFunc, p As Long
    p = ObjPtr(self)
    If gameStartScreens_has(p) Then Set onCloseFunc = gameStartScreens.Item("p" & p)
    shutdownGameStartScreen
    WzFunc_call onCloseFunc
End Sub

' MARK: - WzLoadingPlayersStatusForm

' WzLoadingPlayersStatusForm() : W_FORM() {}
Public Sub WzLoadingPlayersStatusForm_construct(ByVal self As WIDGET)
    W_FORM_construct self
    self.kind = K_WzLoadingPlayersStatusForm ' (VB) virtual calls from here on reach WzLoadingPlayersStatusForm's overrides
End Sub

' (~WzLoadingPlayersStatusForm(): empty)


' WzPlayerStatusCheckboxButton


' WzPlayerStatusCheckboxButton(uint32_t playerIdx, std::shared_ptr<WzText> checkmarkText)
Public Sub WzPlayerStatusCheckboxButton_construct(ByVal self As WIDGET, ByVal playerIdx As Long, ByVal checkmarkText As WzText)
    W_BUTTON_construct self
    self.kind = K_WzPlayerStatusCheckboxButton ' (VB) virtual calls from here on reach WzPlayerStatusCheckboxButton's overrides
    WzPlayerStatusCheckboxButton_(self).player = playerIdx
    Set WzPlayerStatusCheckboxButton_(self).checkmarkText = checkmarkText
End Sub

' std::make_shared<WzPlayerStatusCheckboxButton>(player, checkmarkText)
Public Function WzPlayerStatusCheckboxButton_new(ByVal playerIdx As Long, ByVal checkmarkText As WzText) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_WzPlayerStatusCheckboxButton)
    WzPlayerStatusCheckboxButton_construct widget, playerIdx, checkmarkText
    Set WzPlayerStatusCheckboxButton_new = widget
End Function

Public Sub WzPlayerStatusCheckboxButton_destruct(ByVal self As WIDGET)
    WzPlayerStatusCheckboxButton_closeHostOptionsOverlay self
End Sub

Public Sub WzPlayerStatusCheckboxButton_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    Dim d As WzPlayerStatusCheckboxButton, newStatus As Long
    Set d = WzPlayerStatusCheckboxButton_(self)
    newStatus = LoadingStatus_Loading
    If isHumanPlayer(d.player) Then
        If ingame.JoiningInProgress(d.player) Then newStatus = LoadingStatus_Loading Else newStatus = LoadingStatus_Ready ' Human player, still joining or joined.
        If ingame.PendingDisconnect(d.player) Then
            newStatus = LoadingStatus_Disconnected
        End If
    ElseIf NetPlay_players_ai(d.player) >= 0 Then
        newStatus = LoadingStatus_Ready  ' AI player (automatically joined).
    Else
        newStatus = LoadingStatus_Disconnected
    End If
    If newStatus <> d.status Then
        Select Case newStatus
        Case LoadingStatus_Loading
            WzPlayerStatusCheckboxButton_enableProgressIndicator self
        Case LoadingStatus_Ready
            W_BUTTON_setProgressBorder_nullopt self
            WzPlayerStatusCheckboxButton_closeHostOptionsOverlay self
        Case LoadingStatus_Disconnected
            W_BUTTON_setProgressBorder_nullopt self
            WzPlayerStatusCheckboxButton_closeHostOptionsOverlay self
        End Select
        d.status = newStatus
    End If
End Sub

Public Sub WzPlayerStatusCheckboxButton_geometryChanged(ByVal self As WIDGET)
    If WzPlayerStatusCheckboxButton_(self).status = LoadingStatus_Loading Then
        WzPlayerStatusCheckboxButton_enableProgressIndicator self
    End If
End Sub

Public Sub WzPlayerStatusCheckboxButton_setString(ByVal self As WIDGET, ByVal string_ As String)
    W_BUTTON_setString self, string_
    WzPlayerStatusCheckboxButton_(self).wzText.setText string_, W_BUTTON_(self).FontID
End Sub

Public Sub WzPlayerStatusCheckboxButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As WzPlayerStatusCheckboxButton, x0 As Long, y0 As Long, isClickLock As Boolean, isDown As Boolean, isHighlight As Boolean
    Dim boxX0 As Long, boxX1 As Long, boxY0 As Long, borderColor As Long, indicator_x1 As Long, fx As Long, fy As Long
    Dim textColor As Long, checkmarkX0 As Long, checkmarkY0 As Long, state As Long
    Set d = WzPlayerStatusCheckboxButton_(self)
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset

    state = W_BUTTON_(self).state
    isClickLock = (state And WBUT_CLICKLOCK) <> 0
    isDown = (state And (WBUT_DOWN Or WBUT_LOCK Or WBUT_CLICKLOCK)) <> 0
    isHighlight = (state And WBUT_HIGHLIGHT) <> 0

    If NetPlay_isHost Then
        ' hosts can click on loading players to get options - show hover state
        If d.status = LoadingStatus_Loading And (isHighlight Or isDown) Then
            boxX0 = MAX_(x0 - 1, 0)
            boxX1 = boxX0 + self.width() + 2
            boxY0 = y0
            pie_UniTransBoxFill boxX0, boxY0, boxX1, boxY0 + self.height(), psPalette(WZCOL_MENU_BACKGROUND)
            If isDown Then
                pie_UniTransBoxFill boxX0, boxY0, boxX1, boxY0 + self.height(), pal_RGBA(0, 0, 0, 100)
            End If

            If isHighlight Or isClickLock Then borderColor = pal_RGBA(255, 255, 255, 180) Else borderColor = pal_RGBA(255, 255, 255, 75)
            iV_Box boxX0 - 1, boxY0 - 1, boxX1 + 1, boxY0 + self.height() + 1, borderColor
        End If
    End If

    d.wzText.setText W_BUTTON_(self).pText, W_BUTTON_(self).FontID

    indicator_x1 = IndicatorSize + IndicatorRightPadding

    fx = x0 + indicator_x1
    fy = y0 + (self.height() - d.wzText.lineSize()) \ 2 - d.wzText.aboveBase()

    textColor = psPalette(WZCOL_TEXT_MEDIUM)
    Select Case d.status
    Case LoadingStatus_Loading
        textColor = psPalette(WZCOL_TEXT_DARK)
    Case LoadingStatus_Ready
        textColor = psPalette(WZCOL_TEXT_BRIGHT)
    Case LoadingStatus_Disconnected
        textColor = psPalette(WZCOL_GREY)
    End Select

    If d.status = LoadingStatus_Ready Then
        checkmarkX0 = x0
        checkmarkY0 = y0 + (self.height() - d.checkmarkText.lineSize()) \ 2 - d.checkmarkText.aboveBase()
        d.checkmarkText.render checkmarkX0, checkmarkY0, pal_RGBA(0, 255, 0, 255)
    End If

    d.wzText.render fx, fy, textColor
End Sub

Public Function WzPlayerStatusCheckboxButton_idealWidth(ByVal self As WIDGET) As Long
    WzPlayerStatusCheckboxButton_idealWidth = WzPlayerStatusCheckboxButton_(self).wzText.width() + IndicatorSize + IndicatorRightPadding
End Function

Public Function WzPlayerStatusCheckboxButton_idealHeight(ByVal self As WIDGET) As Long
    WzPlayerStatusCheckboxButton_idealHeight = MAX_(iV_GetTextLineSize(W_BUTTON_(self).FontID), IndicatorSize)
End Function

' override default button click handling
Public Sub WzPlayerStatusCheckboxButton_released(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    Dim d As WzPlayerStatusCheckboxButton, clickAndReleaseOnThisButton As Boolean
    Set d = WzPlayerStatusCheckboxButton_(self)
    clickAndReleaseOnThisButton = ((W_BUTTON_(self).state And WBUT_DOWN) <> 0) ' relies on W_BUTTON handling to properly set WBUT_DOWN

    W_BUTTON_released self, psContext, key

    If Not clickAndReleaseOnThisButton Then
        Exit Sub ' do nothing
    End If

    If key = WKEY_PRIMARY Then
        If NetPlay_isHost And d.status = LoadingStatus_Loading Then
            If Not (d.player <> NetPlay_hostPlayer) Then ASSERT_FAILURE "player != NetPlay.hostPlayer", "src/screens/netpregamescreen.cpp:266", "released", "Shouldn't be possible for host player?"
            ' Display a pop-up for managing loading players
            WzPlayerStatusCheckboxButton_displayHostOptionsOverlay self, self
            self.setState WBUT_CLICKLOCK
        End If
    End If
End Sub

Private Sub WzPlayerStatusCheckboxButton_enableProgressIndicator(ByVal self As WIDGET)
    Dim yPadding As Long, inset As W_BUTTON_ProgressBorder_BorderInset, pb As W_BUTTON_ProgressBorder
    If self.height() > IndicatorSize Then yPadding = self.height() - IndicatorSize Else yPadding = 0
    inset = W_BUTTON_ProgressBorder_BorderInset_(0, yPadding \ 2, self.width() - IndicatorSize, yPadding \ 2)
    pb = W_BUTTON_ProgressBorder_indeterminate(inset)
    W_BUTTON_setProgressBorder self, pb, psPalette(WZCOL_TEXT_MEDIUM)
End Sub

' createHostOptionsPopoverForm's auto createOptionsButton = [weakPlayerStatusButton](const WzString& text, const
' std::function<void (W_BUTTON& button)>& onClickFunc) -> std::shared_ptr<W_BUTTON>
Private Function createOptionsButton(ByVal weakPlayerStatusButton As Long, ByVal text As String, ByVal onClickFunc As WzFunc) As WIDGET
    Dim button As WIDGET, minButtonWidthForText As Long
    Set button = W_BUTTON_new()
    button.setString text
    W_BUTTON_(button).FontID = font_regular
    Set button.displayFunction = WzFunc_new(FN_PopoverMenuButtonDisplayFunc)
    Set button.pUserData = New PopoverMenuButtonDisplayCache
    ' (setOnDelete deletes the cache: it goes with the widget)
    minButtonWidthForText = iV_GetTextWidth(text, W_BUTTON_(button).FontID)
    button.setGeometry 0, 0, minButtonWidthForText + 10, iV_GetTextLineSize(W_BUTTON_(button).FontID) + 2

    ' On click, perform the designated onClickHandler and close the popover form / overlay
    W_BUTTON_addOnClickHandler button, WzFunc_new(FN_pregame_optionsButton_onClick, weakPlayerStatusButton, onClickFunc)
    Set createOptionsButton = button
End Function

Private Function WzPlayerStatusCheckboxButton_createHostOptionsPopoverForm(ByVal self As WIDGET) As WIDGET
    Dim weakPlayerStatusButton As Long, buttons() As WIDGET, buttonsCount As Long, kickButtonStr As String
    Dim totalButtonHeight As Long, maxButtonWidth As Long, itemsList As WIDGET, listPadding As Padding, i As Long
    Const itemSpacing As Long = 4
    ' create all the buttons / option rows
    weakPlayerStatusButton = ObjPtr(self)

    ReDim buttons(0 To 0)
    If NetPlay_players_isSpectator(WzPlayerStatusCheckboxButton_(self).player) Then kickButtonStr = "Kick Spectator" Else kickButtonStr = "Kick Player"
    Set buttons(buttonsCount) = createOptionsButton(weakPlayerStatusButton, kickButtonStr, WzFunc_new(FN_pregame_kickButton_onClick, weakPlayerStatusButton))
    buttonsCount = buttonsCount + 1

    ' determine required height for all buttons
    totalButtonHeight = 0
    For i = 0 To buttonsCount - 1
        totalButtonHeight = totalButtonHeight + buttons(i).height()
    Next
    maxButtonWidth = buttons(0).width()     ' (*(std::max_element(...)))->width()
    For i = 1 To buttonsCount - 1
        If maxButtonWidth < buttons(i).width() Then maxButtonWidth = buttons(i).width()
    Next

    Set itemsList = ScrollableListWidget_make()
    ScrollableListWidget_setBackgroundColor itemsList, psPalette(WZCOL_MENU_BACKGROUND)
    ScrollableListWidget_setBorderColor itemsList, pal_RGBA(152, 151, 154, 255)
    listPadding.top = 3: listPadding.right_ = 4: listPadding.bottom = 3: listPadding.left_ = 4    ' Padding padding = {3, 4, 3, 4};
    ScrollableListWidget_setPadding itemsList, listPadding
    ScrollableListWidget_setItemSpacing itemsList, itemSpacing
    itemsList.setGeometry itemsList.x(), itemsList.y(), maxButtonWidth + listPadding.left_ + listPadding.right_, totalButtonHeight + ((buttonsCount - 1) * itemSpacing) + listPadding.top + listPadding.bottom
    For i = 0 To buttonsCount - 1
        ScrollableListWidget_addItem itemsList, buttons(i)
    Next

    Set WzPlayerStatusCheckboxButton_createHostOptionsPopoverForm = itemsList
End Function

Private Sub WzPlayerStatusCheckboxButton_displayHostOptionsOverlay(ByVal self As WIDGET, ByVal psParent As WIDGET)
    Dim d As WzPlayerStatusCheckboxButton, lockedScreen As W_SCREEN, newRootFrm As WIDGET, psWeakOptionsOverlayScreen As Long
    Dim psWeakPlayerStatusButton As Long, optionsPopOver As WIDGET, weakParent As Long
    Set d = WzPlayerStatusCheckboxButton_(self)
    Set lockedScreen = self.screenPointer
    If Not (Not lockedScreen Is Nothing) Then ASSERT_FAILURE "lockedScreen != nullptr", "src/screens/netpregamescreen.cpp:366", "displayHostOptionsOverlay", "No associated screen pointer?"

    ' Initialize the options overlay screen
    Set d.optionsOverlayScreen = W_SCREEN_make()
    Set newRootFrm = W_FULLSCREENOVERLAY_CLICKFORM_make()
    Set newRootFrm.displayFunction = WzFunc_new(FN_displayChildDropShadows)
    psWeakOptionsOverlayScreen = ObjPtr(d.optionsOverlayScreen)
    psWeakPlayerStatusButton = ObjPtr(self)
    Set W_FULLSCREENOVERLAY_CLICKFORM_(newRootFrm).onClickedFunc = WzFunc_new(FN_pregame_optionsOverlay_clicked, psWeakOptionsOverlayScreen, psWeakPlayerStatusButton)
    Set W_FULLSCREENOVERLAY_CLICKFORM_(newRootFrm).onCancelPressed = W_FULLSCREENOVERLAY_CLICKFORM_(newRootFrm).onClickedFunc
    d.optionsOverlayScreen.psForm.attach newRootFrm

    ' Create the pop-over form
    Set optionsPopOver = WzPlayerStatusCheckboxButton_createHostOptionsPopoverForm(self)
    newRootFrm.attach optionsPopOver

    ' Position the pop-over form
    weakParent = ObjPtr(psParent)
    optionsPopOver.setCalcLayout WzFunc_new(FN_pregame_optionsPopOver_calcLayout, weakParent)

    widgRegisterOverlayScreenOnTopOfScreen d.optionsOverlayScreen, lockedScreen
End Sub

Private Sub WzPlayerStatusCheckboxButton_closeHostOptionsOverlay(ByVal self As WIDGET)
    Dim d As WzPlayerStatusCheckboxButton
    Set d = WzPlayerStatusCheckboxButton_(self)
    If Not d.optionsOverlayScreen Is Nothing Then
        widgRemoveOverlayScreen d.optionsOverlayScreen
        Set d.optionsOverlayScreen = Nothing
        W_BUTTON_(self).state = W_BUTTON_(self).state And Not WBUT_CLICKLOCK
    End If
End Sub

Public Function WzLoadingPlayersStatusForm_make() As WIDGET
    Dim result As WIDGET, d As WzLoadingPlayersStatusForm, checkmarkText As WzText, playerIndexes() As Long
    Dim playerIndexesCount As Long, player As Long, pPlayerCheckbox As WIDGET, numSpectators As Long, i As Long, j As Long
    Dim t As Long
    Set result = WIDGET_alloc(K_WzLoadingPlayersStatusForm)
    WzLoadingPlayersStatusForm_construct result
    Set d = WzLoadingPlayersStatusForm_(result)

    ' Add title text
    Set d.titleLabel = W_LABEL_new()
    W_LABEL_setFont d.titleLabel, font_regular_bold, psPalette(WZCOL_FORM_TEXT)
    d.titleLabel.setString "Loading Status:"
    d.titleLabel.setGeometry 0, 0, W_LABEL_getMaxLineWidth(d.titleLabel), iV_GetTextLineSize(font_regular_bold)
    W_LABEL_setCacheNeverExpires d.titleLabel, True
    result.attach d.titleLabel
    d.titleLabel.setCalcLayout WzFunc_new(FN_pregame_titleLabel_calcLayout)

    ' add table for players
    Set d.playersList = ScrollableListWidget_make()
    result.attach d.playersList
    ScrollableListWidget_setBackgroundColor d.playersList, pal_RGBA(0, 0, 0, 25)
    ScrollableListWidget_setItemSpacing d.playersList, 1

    Set checkmarkText = New WzText
    checkmarkText.setText ChrW$(&H2713), font_regular   ' "\u2713" (a check mark)

    ' Add rows for all players
    ReDim playerIndexes(0 To MAX_PLAYERS - 1)
    For player = 0 To MIN_(game.maxPlayers, MAX_PLAYERS) - 1
        If isHumanPlayer(player) _
            Or NetPlay_players_difficulty(player) = AIDifficulty_HUMAN _
            Or NetPlay_players_ai(player) >= 0 Then
            ' is an active (connected) human player, was a human player (probably disconnected) or is an AI bot
            playerIndexes(playerIndexesCount) = player
            playerIndexesCount = playerIndexesCount + 1
        End If
    Next
    ' sort by player position (std::sort: an insertion sort)
    For i = 1 To playerIndexesCount - 1
        t = playerIndexes(i)
        j = i - 1
        Do While j >= 0
            If Not (NetPlay_players_position(t) < NetPlay_players_position(playerIndexes(j))) Then Exit Do
            playerIndexes(j + 1) = playerIndexes(j)
            j = j - 1
        Loop
        playerIndexes(j + 1) = t
    Next
    For i = 0 To playerIndexesCount - 1
        player = playerIndexes(i)
        Set pPlayerCheckbox = WzPlayerStatusCheckboxButton_new(player, checkmarkText)
        W_BUTTON_(pPlayerCheckbox).FontID = font_regular
        pPlayerCheckbox.setString getPlayerName(player)
        pPlayerCheckbox.setGeometry 0, 0, pPlayerCheckbox.idealWidth(), pPlayerCheckbox.idealHeight()
        ScrollableListWidget_addItem d.playersList, pPlayerCheckbox
    Next

    ' add table for spectators
    Set d.spectatorsList = ScrollableListWidget_make()
    ScrollableListWidget_setBackgroundColor d.spectatorsList, pal_RGBA(0, 0, 0, 25)
    ScrollableListWidget_setItemSpacing d.spectatorsList, 1

    ' Add rows for all spectators
    numSpectators = 0
    For player = MAX_PLAYERS To MAX_CONNECTED_PLAYERS - 1
        If isHumanPlayer(player) _
            Or NetPlay_players_difficulty(player) = AIDifficulty_HUMAN Then
            ' is an active (connected) human player or was a human player (probably disconnected)
            Set pPlayerCheckbox = WzPlayerStatusCheckboxButton_new(player, checkmarkText)
            W_BUTTON_(pPlayerCheckbox).FontID = font_regular
            pPlayerCheckbox.setString getPlayerName(player)
            pPlayerCheckbox.setGeometry 0, 0, pPlayerCheckbox.idealWidth(), pPlayerCheckbox.idealHeight()
            ScrollableListWidget_addItem d.spectatorsList, pPlayerCheckbox
            numSpectators = numSpectators + 1
        End If
    Next

    If numSpectators <> 0 Then
        result.attach d.spectatorsList
    Else
        Set d.spectatorsList = Nothing
    End If

    d.playersList.setCalcLayout WzFunc_new(FN_pregame_playersList_calcLayout)

    If Not d.spectatorsList Is Nothing Then
        d.spectatorsList.setCalcLayout WzFunc_new(FN_pregame_spectatorsList_calcLayout)
    End If

    Set WzLoadingPlayersStatusForm_make = result
End Function

Public Sub WzLoadingPlayersStatusForm_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    ' currently, no-op
End Sub

Public Sub WzLoadingPlayersStatusForm_geometryChanged(ByVal self As WIDGET)
    Dim d As WzLoadingPlayersStatusForm
    Set d = WzLoadingPlayersStatusForm_(self)
    If Not d.titleLabel Is Nothing Then
        d.titleLabel.callCalcLayout
    End If
    If Not d.playersList Is Nothing Then
        d.playersList.callCalcLayout
    End If
    If Not d.spectatorsList Is Nothing Then
        d.spectatorsList.callCalcLayout
    End If
End Sub

Public Function WzLoadingPlayersStatusForm_idealWidth(ByVal self As WIDGET) As Long
    Dim d As WzLoadingPlayersStatusForm, listsWidth As Long
    Set d = WzLoadingPlayersStatusForm_(self)
    listsWidth = d.playersList.idealWidth()
    If Not d.spectatorsList Is Nothing Then
        listsWidth = listsWidth + PLAYER_LOADING_STATUS_ITEM_SPACING + d.spectatorsList.idealWidth()
    End If
    WzLoadingPlayersStatusForm_idealWidth = MAX_(d.titleLabel.idealWidth(), listsWidth)
End Function

Public Function WzLoadingPlayersStatusForm_idealHeight(ByVal self As WIDGET) As Long
    Dim d As WzLoadingPlayersStatusForm, spectatorsHeight As Long
    Set d = WzLoadingPlayersStatusForm_(self)
    If Not d.spectatorsList Is Nothing Then spectatorsHeight = d.spectatorsList.idealHeight()
    WzLoadingPlayersStatusForm_idealHeight = d.titleLabel.y() + d.titleLabel.height() + PLAYER_LOADING_STATUS_ITEM_SPACING + MAX_(d.playersList.idealHeight(), spectatorsHeight)
End Function

Public Sub WzLoadingPlayersStatusForm_screenSizeDidChange(ByVal self As WIDGET, ByVal oldWidth As Long, ByVal oldHeight As Long, ByVal newWidth As Long, ByVal newHeight As Long)
    ' call default implementation (which ultimately propagates to all children
    W_FORM_screenSizeDidChange self, oldWidth, oldHeight, newWidth, newHeight
End Sub

' MARK: - WzCountdownLabel

Public Sub WzCountdownLabel_construct(ByVal self As WIDGET)
    WIDGET_construct self
    self.kind = K_WzCountdownLabel ' (VB) virtual calls from here on reach WzCountdownLabel's overrides
End Sub

' std::make_shared<WzCountdownLabel>()
Public Function WzCountdownLabel_new() As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_WzCountdownLabel)
    WzCountdownLabel_construct widget
    Set WzCountdownLabel_new = widget
End Function

Public Sub WzCountdownLabel_setCountdown(ByVal self As WIDGET, ByVal startTime_ As Currency, ByVal countdownSeconds_ As Double)
    Dim d As WzCountdownLabel
    Set d = WzCountdownLabel_(self)
    d.has_startTime = True
    d.startTime = startTime_
    d.countdownSeconds = countdownSeconds_
    d.maxWidth = iV_GetTextWidth(CStr(d.countdownSeconds), font_large) ' not accurate for all numbers, but hopefully in the ballpark...
    self.dirty = True
End Sub

Public Sub WzCountdownLabel_clearCountdown(ByVal self As WIDGET)
    WzCountdownLabel_(self).has_startTime = False
End Sub

Public Sub WzCountdownLabel_setTextAlignment(ByVal self As WIDGET, ByVal align As WzTextAlignment)
    self.style = self.style And Not (WLAB_ALIGNLEFT Or WLAB_ALIGNCENTRE Or WLAB_ALIGNRIGHT)
    self.style = self.style Or align
    self.dirty = True
End Sub

Public Sub WzCountdownLabel_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As WzCountdownLabel, currentTime As Currency, secondsLeft As Double, textTotalHeight As Long
    Dim textBoundingBoxOffsetX As Long, textBoundingBoxOffsetY As Long, fx As Long, fw As Long, fy As Single
    Dim textColour As Long, fractionOfWholeSecond As Double, a As Long
    Set d = WzCountdownLabel_(self)
    If Not d.has_startTime Then
        Exit Sub
    End If
    currentTime = steadyNow()
    secondsLeft = d.countdownSeconds - Fix(steadyMsBetween(d.startTime, currentTime) / 1000)  ' duration_cast<std::chrono::seconds>

    If secondsLeft = 0 Then
        Exit Sub
    End If

    d.displayedNumber.setText CStr(secondsLeft), d.FontID

    textTotalHeight = d.displayedNumber.lineSize()

    textBoundingBoxOffsetX = 0      ' Vector2i textBoundingBoxOffset(0, 0);
    textBoundingBoxOffsetY = 0
    If (self.style And WLAB_ALIGNTOPLEFT) <> 0 Then     ' Align top
        textBoundingBoxOffsetY = yOffset + self.y()
    ElseIf (self.style And WLAB_ALIGNBOTTOMLEFT) <> 0 Then  ' Align bottom
        textBoundingBoxOffsetY = yOffset + self.y() + (self.height() - textTotalHeight)
    Else
        textBoundingBoxOffsetY = yOffset + self.y() + (self.height() - textTotalHeight) \ 2
    End If

    fx = 0
    If (self.style And WLAB_ALIGNCENTRE) <> 0 Then
        fw = d.displayedNumber.width()
        fx = xOffset + self.x() + (self.width() - fw) \ 2
    ElseIf (self.style And WLAB_ALIGNRIGHT) <> 0 Then
        fw = d.displayedNumber.width()
        fx = xOffset + self.x() + self.width() - fw
    Else
        fx = xOffset + self.x()
    End If

    fy = CSng(textBoundingBoxOffsetY) - CSng(d.displayedNumber.aboveBase())

    textColour = d.fontColour
    If secondsLeft <= 3 Then
        fractionOfWholeSecond = steadyMsBetween(d.startTime, currentTime) - Int(steadyMsBetween(d.startTime, currentTime) / 1000) * 1000   ' % 1000
        If fractionOfWholeSecond <> 0 Then
            a = pielightAlpha(textColour)
            a = a - Fix(CSng(a) * MIN_F(CSng(fractionOfWholeSecond / 1000!), 1!))
            textColour = pielightWithAlpha(textColour, a)
        End If
    End If

    d.displayedNumber.render textBoundingBoxOffsetX + fx, fy, textColour
End Sub

' std::min<float>
Private Function MIN_F(ByVal a As Single, ByVal b As Single) As Single
    If b < a Then MIN_F = b Else MIN_F = a
End Function

Public Function WzCountdownLabel_idealWidth(ByVal self As WIDGET) As Long
    WzCountdownLabel_idealWidth = WzCountdownLabel_(self).maxWidth
End Function

Public Function WzCountdownLabel_idealHeight(ByVal self As WIDGET) As Long
    WzCountdownLabel_idealHeight = iV_GetTextLineSize(WzCountdownLabel_(self).FontID)
End Function

' MARK: - WzGameStartOverlayScreen_CLICKFORM



' WzGameStartOverlayScreen_CLICKFORM(W_FORMINIT const *init) : W_CLICKFORM(init) {}
Public Sub WzGameStartOverlayScreen_CLICKFORM_construct_init(ByVal self As WIDGET, init As W_FORMINIT)
    W_CLICKFORM_construct_init self, init
    self.kind = K_WzGameStartOverlayScreen_CLICKFORM ' (VB) virtual calls from here on reach WzGameStartOverlayScreen_CLICKFORM's overrides
End Sub

' WzGameStartOverlayScreen_CLICKFORM() : W_CLICKFORM() {}
Public Sub WzGameStartOverlayScreen_CLICKFORM_construct(ByVal self As WIDGET)
    W_CLICKFORM_construct self
    self.kind = K_WzGameStartOverlayScreen_CLICKFORM ' (VB) virtual calls from here on reach WzGameStartOverlayScreen_CLICKFORM's overrides
End Sub

Public Sub WzGameStartOverlayScreen_CLICKFORM_destruct(ByVal self As WIDGET)
    consoleRemoveMessageListener WzGameStartOverlayScreen_CLICKFORM_(self).handleConsoleMessage
End Sub

Public Function WzGameStartOverlayScreen_CLICKFORM_make(Optional ByVal formID As Long = 0) As WIDGET
    Dim sInit As W_FORMINIT, widget As WIDGET
    W_FORMINIT_construct sInit
    sInit.id = formID
    sInit.style = WFORM_PLAIN Or WFORM_CLICKABLE
    sInit.x = 0
    sInit.y = 0
    sInit.width = screenWidth - 1
    sInit.height = screenHeight - 1
    Set sInit.calcLayout = WzFunc_new(FN_pregame_clickForm_calcLayout)

    ' (class make_shared_enabler)
    Set widget = WIDGET_alloc(K_WzGameStartOverlayScreen_CLICKFORM)
    WzGameStartOverlayScreen_CLICKFORM_construct_init widget, sInit
    WzGameStartOverlayScreen_CLICKFORM_initialize widget
    Set WzGameStartOverlayScreen_CLICKFORM_make = widget
End Function

Private Sub WzGameStartOverlayScreen_CLICKFORM_initialize(ByVal self As WIDGET)
    Dim d As WzGameStartOverlayCFData
    Set d = WzGameStartOverlayScreen_CLICKFORM_(self)
    ' Create topLabel (bold)
    Set d.topLabel = W_LABEL_new()
    W_LABEL_setFont d.topLabel, font_medium_bold, psPalette(WZCOL_TEXT_BRIGHT)
    W_LABEL_setTextAlignment d.topLabel, WLAB_ALIGNCENTRE
    self.attach d.topLabel

    ' Create countdownLabel (bold)
    Set d.countdownLabel = WzCountdownLabel_new()
    WzCountdownLabel_setTextAlignment d.countdownLabel, WLAB_ALIGNRIGHT
    WzCountdownLabel_clearCountdown d.countdownLabel
    self.attach d.countdownLabel

    ' Create a console message listener
    Set d.handleConsoleMessage = WzFunc_new(FN_pregame_consoleMessage, ObjPtr(self))

    ' Determine if any players will have a team to share strategy with
    d.bHasTeamPlanningStage = gameHasTeamStrategyView()

    ' Create a team strategy view (if player is on a team) - may return null if there's no reason to display it
    Set d.teamStrategyView = createTeamStrategyView()
    If Not d.teamStrategyView Is Nothing Then
        ' Add side label
        Set d.strategySideTextLabel = addSideText_parent(self, 0, 0, 0, "TEAM STRATEGY")

        self.attach d.teamStrategyView
    End If

    ' Create a PlayerConnectionStatusForm
    Set d.playersLoadingStatus = WzLoadingPlayersStatusForm_make()
    self.attach d.playersLoadingStatus

    WzGameStartOverlayScreen_CLICKFORM_recalcLayout self

    consoleAddMessageListener d.handleConsoleMessage
End Sub

Private Sub WzGameStartOverlayScreen_CLICKFORM_recalcLayout(ByVal self As WIDGET)
    Dim d As WzGameStartOverlayCFData, maxChildWidth As Long, topLabelWidth As Long, countdownLabelWidth As Long
    Dim bottomOfLastElement As Long, bottomWidget As WIDGET, bottomWidgetWidth As Long, bottomWidgetHeight As Long
    Dim teamStrategyTopPadding As Long, teamStrategyViewWidth As Long, maxAvailableHeightForTeamStrategyView As Long
    Dim teamStrategyViewHeight As Long, teamStrategyX0 As Long, teamStrategyY0 As Long, strategySideTextX0 As Long
    Dim bottomVisibleHeight As Long
    Set d = WzGameStartOverlayScreen_CLICKFORM_(self)
    maxChildWidth = self.width() - (INTERNAL_PADDING * 2)

    topLabelWidth = maxChildWidth
    d.topLabel.setGeometry INTERNAL_PADDING, INTERNAL_PADDING + 8, topLabelWidth, W_LABEL_requiredHeight(d.topLabel)

    countdownLabelWidth = 30
    d.countdownLabel.setGeometry self.width() - (INTERNAL_PADDING + countdownLabelWidth), INTERNAL_PADDING, countdownLabelWidth, d.countdownLabel.idealHeight()

    bottomOfLastElement = d.topLabel.y() + d.topLabel.height()

    ' Bottom-aligned widget (player loading status)
    Set bottomWidget = d.playersLoadingStatus
    bottomWidgetWidth = bottomWidget.idealWidth()
    bottomWidgetHeight = bottomWidget.idealHeight()
    bottomWidget.setGeometry (self.width() - bottomWidgetWidth) \ 2, self.height() - (bottomWidgetHeight + INTERNAL_PADDING), bottomWidgetWidth, bottomWidgetHeight

    If Not d.teamStrategyView Is Nothing Then
        ' place the team strategy information in the middle, below the topLabel
        teamStrategyTopPadding = INTERNAL_PADDING * 2
        teamStrategyViewWidth = MIN_(maxChildWidth - W_LABEL_requiredHeight(d.strategySideTextLabel), d.teamStrategyView.idealWidth())
        If bottomWidget.visible() Then bottomVisibleHeight = bottomWidgetHeight Else bottomVisibleHeight = 0
        maxAvailableHeightForTeamStrategyView = self.height() - bottomVisibleHeight - (teamStrategyTopPadding * 2)
        teamStrategyViewHeight = MIN_(maxAvailableHeightForTeamStrategyView, d.teamStrategyView.idealHeight())
        teamStrategyX0 = (self.width() - teamStrategyViewWidth) \ 2
        teamStrategyY0 = bottomOfLastElement + teamStrategyTopPadding
        d.teamStrategyView.setGeometry teamStrategyX0, teamStrategyY0, teamStrategyViewWidth, teamStrategyViewHeight

        ' place the side text... to the side
        strategySideTextX0 = MAX_(teamStrategyX0 - W_LABEL_requiredHeight(d.strategySideTextLabel), 0)
        If strategySideTextX0 >= W_LABEL_requiredHeight(d.strategySideTextLabel) Then
            d.strategySideTextLabel.setGeometry strategySideTextX0, teamStrategyY0, d.strategySideTextLabel.width(), d.strategySideTextLabel.height()
            d.strategySideTextLabel.show
        Else
            d.strategySideTextLabel.hide
        End If
    End If
End Sub

Public Sub WzGameStartOverlayScreen_CLICKFORM_geometryChanged(ByVal self As WIDGET)
    ' (VB) WIDGET's constructor calls the calcLayout (sInit.calcLayout), whose setGeometry reaches this override:
    ' WIDGET_alloc gave the widget its final kind. C++'s constructor reaches WIDGET::geometryChanged there (a no-op):
    ' nothing to lay out before initialize
    If WzGameStartOverlayScreen_CLICKFORM_(self).topLabel Is Nothing Then Exit Sub
    WzGameStartOverlayScreen_CLICKFORM_recalcLayout self
End Sub

Public Sub WzGameStartOverlayScreen_CLICKFORM_clicked(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    Dim d As WzGameStartOverlayCFData
    Set d = WzGameStartOverlayScreen_CLICKFORM_(self)
    If Not d.onClickedFunc Is Nothing Then
        WzFunc_call d.onClickedFunc
    End If
End Sub

Private Sub pregamscreen_draw_console_text(ByVal display As WzText, ByVal colour As Long, ByVal x As Long, ByVal y As Long, ByVal justify As CONSOLE_TEXT_JUSTIFICATION, ByVal width As Long)
    Select Case justify
    Case LEFT_JUSTIFY
        ' do nothing
    Case RIGHT_JUSTIFY
        x = x + width - display.width()
    Case CENTRE_JUSTIFY
        x = x + (width - display.width()) \ 2
    End Select
    display.render x, y, colour
End Sub

Public Sub WzGameStartOverlayScreen_CLICKFORM_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As WzGameStartOverlayCFData, x0 As Long, y0 As Long, TextX0 As Long, TextYpos As Long, TextMaxWidth As Long
    Dim ActiveMessage As CachedConsoleMessage, textColor As Long
    Set d = WzGameStartOverlayScreen_CLICKFORM_(self)
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

    ' draw console messages, if present
    TextX0 = x0 + INTERNAL_PADDING
    TextYpos = y0 + d.topLabel.y() + d.topLabel.height() + INTERNAL_PADDING
    TextMaxWidth = self.width() - (INTERNAL_PADDING * 2)
    For Each ActiveMessage In d.currentConsoleMessages
        textColor = getConsoleTextColor(ActiveMessage.player)
        textColor = pielightWithAlpha(textColor, pielightAlpha(textColor) \ 2)
        pregamscreen_draw_console_text ActiveMessage.display, textColor, TextX0, _
                         TextYpos, ActiveMessage.JustifyType, TextMaxWidth
        TextYpos = TextYpos + ActiveMessage.display.lineSize()
    Next
End Sub

Public Sub WzGameStartOverlayScreen_CLICKFORM_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    Dim d As WzGameStartOverlayCFData
    Set d = WzGameStartOverlayScreen_CLICKFORM_(self)
    If keyPressed(KEY_ESC) Then
        If Not d.onCancelPressed Is Nothing Then
            WzFunc_call d.onCancelPressed
        End If
    End If
    inputLoseFocus   ' clear the input buffer.

    WzGameStartOverlayScreen_CLICKFORM_updateData self
End Sub

' transitionToNextState's auto transitionToFinalCountdown = [&]()
Private Sub WzGameStartOverlayScreen_CLICKFORM_transitionToFinalCountdown(ByVal self As WIDGET)
    Dim d As WzGameStartOverlayCFData
    Set d = WzGameStartOverlayScreen_CLICKFORM_(self)
    d.currentState = PregameWaitingState_WaitingForFinalCountdown
    d.topLabel.setString "Game will start in ..."
    d.currentCountdownStart = steadyNow()
    d.has_currentCountdownStart = True
    WzCountdownLabel_setCountdown d.countdownLabel, d.currentCountdownStart, WZ_FINAL_COUNTDOWN_MS \ 1000
    WzGameStartOverlayScreen_CLICKFORM_recalcLayout self
End Sub

Private Sub WzGameStartOverlayScreen_CLICKFORM_transitionToNextState(ByVal self As WIDGET)
    Dim d As WzGameStartOverlayCFData
    Set d = WzGameStartOverlayScreen_CLICKFORM_(self)
    Select Case d.currentState
    Case PregameWaitingState_Uninitialized
        d.currentState = PregameWaitingState_WaitingForPlayersToJoin
        d.topLabel.setString "Waiting for other players ..."
        WzGameStartOverlayScreen_CLICKFORM_recalcLayout self
    Case PregameWaitingState_WaitingForPlayersToJoin
        If d.bHasTeamPlanningStage Then
            d.currentState = PregameWaitingState_WaitingForTeamPlanning
            d.topLabel.setString "Team Planning"
            d.currentCountdownStart = steadyNow()
            d.has_currentCountdownStart = True
            WzCountdownLabel_setCountdown d.countdownLabel, d.currentCountdownStart, WZ_TEAMPLANNING_COUNTDOWN_MS \ 1000
            If Not d.teamStrategyView Is Nothing Then
                transformTeamStrategyViewMode d.teamStrategyView, True

                ' if needed, hide playersLoadingStatus so the team planning can use the full height
                If True Then
                    d.playersLoadingStatus.hide
                End If
            End If

            WzGameStartOverlayScreen_CLICKFORM_recalcLayout self
        Else
            WzGameStartOverlayScreen_CLICKFORM_transitionToFinalCountdown self
        End If
    Case PregameWaitingState_WaitingForTeamPlanning
        WzGameStartOverlayScreen_CLICKFORM_transitionToFinalCountdown self
    Case PregameWaitingState_WaitingForFinalCountdown
    End Select
End Sub

Private Sub WzGameStartOverlayScreen_CLICKFORM_updateData(ByVal self As WIDGET)
    Dim d As WzGameStartOverlayCFData, now_ As Currency, i As Long, m As CachedConsoleMessage, playersWaitingToJoin As Long
    Dim currentTime As Currency, strongScreen As W_SCREEN
    Set d = WzGameStartOverlayScreen_CLICKFORM_(self)
    ' Prune expired console messages
    now_ = steadyNow()
    For i = d.currentConsoleMessages.Count To 1 Step -1
        Set m = d.currentConsoleMessages(i)
        If steadyMsBetween(m.timeAdded, now_) > m.duration Then
            d.currentConsoleMessages.Remove i
        End If
    Next

    ' Handle game start state
    If d.currentState = PregameWaitingState_Uninitialized Then
        WzGameStartOverlayScreen_CLICKFORM_transitionToNextState self
    End If
    If d.currentState = PregameWaitingState_WaitingForPlayersToJoin Then
        playersWaitingToJoin = 0
        For i = 0 To MAX_CONNECTED_PLAYERS - 1
            If isHumanPlayer(i) Then
                If ingame.JoiningInProgress(i) Then
                    playersWaitingToJoin = playersWaitingToJoin + 1
                End If
            End If
        Next
        ' if all players have joined, transition to the next state
        If playersWaitingToJoin = 0 Then
            WzGameStartOverlayScreen_CLICKFORM_transitionToNextState self
        End If
    End If
    If d.currentState = PregameWaitingState_WaitingForTeamPlanning Then
        If d.has_currentCountdownStart Then
            currentTime = steadyNow()
            If steadyMsBetween(d.currentCountdownStart, currentTime) >= WZ_TEAMPLANNING_COUNTDOWN_MS Then
                WzGameStartOverlayScreen_CLICKFORM_transitionToNextState self
            End If
        Else
            d.currentCountdownStart = steadyNow()
            d.has_currentCountdownStart = True
        End If
    End If
    If d.currentState = PregameWaitingState_WaitingForFinalCountdown Then
        If d.has_currentCountdownStart Then
            currentTime = steadyNow()
            If steadyMsBetween(d.currentCountdownStart, currentTime) >= WZ_FINAL_COUNTDOWN_MS Then
                Set strongScreen = self.screenPointer
                If Not strongScreen Is Nothing Then
                    If Not isWzGameStartOverlayScreen(strongScreen) Then ASSERT_FAILURE "gameStartOverlayScreen != nullptr", "src/screens/netpregamescreen.cpp:1070", "updateData", "Invalid gameStartOverlayScreen"
                    WzGameStartOverlayScreen_closeScreen strongScreen
                Else
                    ASSERT_FAILURE "false", "src/screens/netpregamescreen.cpp:1075", "updateData", "No screen is currently associated with this root form?"
                End If
            End If
        Else
            d.currentCountdownStart = steadyNow()
            d.has_currentCountdownStart = True
        End If
    End If
End Sub

' MARK: - WzGameStartOverlayScreen

' std::shared_ptr<WzGameStartOverlayScreen> WzGameStartOverlayScreen::make(const OnCloseFunc& _onCloseFunc)
Private Function WzGameStartOverlayScreen_make(ByVal onCloseFunc As WzFunc) As W_SCREEN
    Dim newRootFrm As WIDGET, screen As W_SCREEN
    ' (class make_shared_enabler: public WzGameStartOverlayScreen {};)
    Set newRootFrm = WzGameStartOverlayScreen_CLICKFORM_make()
    Set screen = W_SCREEN_make(newRootFrm)  ' screen->initialize(newRootFrm);
    ' screen->onCloseFunc = _onCloseFunc;
    If gameStartScreens Is Nothing Then Set gameStartScreens = New Collection
    If gameStartScreens_has(ObjPtr(screen)) Then gameStartScreens.Remove "p" & ObjPtr(screen)  ' (a dead screen's address)
    gameStartScreens.Add onCloseFunc, "p" & ObjPtr(screen)

    Set WzGameStartOverlayScreen_make = screen
End Function

' std::shared_ptr<W_SCREEN> createGameStartScreen(std::function<void ()> onCloseFunc)
Public Function createGameStartScreen(ByVal onCloseFunc As WzFunc) As W_SCREEN
    Dim screen As W_SCREEN
    Set screen = WzGameStartOverlayScreen_make(onCloseFunc)
    widgRegisterOverlayScreenOnTopOfScreen screen, psWScreen
    Set psCurrentGameStartScreen = screen
    Set createGameStartScreen = screen
End Function

Public Sub shutdownGameStartScreen()
    Dim strongGameStartScreen As W_SCREEN
    Set strongGameStartScreen = psCurrentGameStartScreen
    If Not strongGameStartScreen Is Nothing Then
        widgRemoveOverlayScreen strongGameStartScreen
        Set psCurrentGameStartScreen = Nothing
        If gameStartScreens_has(ObjPtr(strongGameStartScreen)) Then gameStartScreens.Remove "p" & ObjPtr(strongGameStartScreen)
    End If
End Sub

' the lambdas
Public Function netpregamescreen_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim psWidget As WIDGET, psParent As WIDGET, pd As WzLoadingPlayersStatusForm, x0 As Long, y0 As Long, height As Long
    Dim strongParent As WIDGET, sd As WzPlayerStatusCheckboxButton, targetPlayer As Long, numHumanPlayers As Long, i As Long
    Dim psOverlayScreen As W_SCREEN, popOverX0 As Long, popOverY0 As Long, button As WIDGET, self As WIDGET
    Dim d As WzGameStartOverlayCFData, addedTime As Currency, messageDuration As Long, m As CachedConsoleMessage
    Dim duration As Long, sender As Long, text As String, justification As Long
    Select Case f.id
    Case FN_pregame_titleLabel_calcLayout       ' LAMBDA_CALCLAYOUT_SIMPLE (WIDGET *psWidget)
        Set psWidget = args(0)
        Set psParent = loadingStatusFormParent(psWidget)
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/screens/netpregamescreen.cpp:453", "make", "No parent": Exit Function
        psWidget.setGeometry 0, 0, psWidget.width(), psWidget.height()

    Case FN_pregame_playersList_calcLayout      ' LAMBDA_CALCLAYOUT_SIMPLE (WIDGET *psWidget)
        Set psWidget = args(0)
        Set psParent = loadingStatusFormParent(psWidget)
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/screens/netpregamescreen.cpp:523", "make", "No parent": Exit Function
        Set pd = WzLoadingPlayersStatusForm_(psParent)
        x0 = 0
        y0 = pd.titleLabel.y() + pd.titleLabel.height() + PLAYER_LOADING_STATUS_ITEM_SPACING
        height = psParent.height() - y0
        psWidget.setGeometry x0, y0, psWidget.idealWidth(), height

    Case FN_pregame_spectatorsList_calcLayout   ' LAMBDA_CALCLAYOUT_SIMPLE (WIDGET *psWidget)
        Set psWidget = args(0)
        Set psParent = loadingStatusFormParent(psWidget)
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/screens/netpregamescreen.cpp:534", "make", "No parent": Exit Function
        Set pd = WzLoadingPlayersStatusForm_(psParent)
        x0 = pd.playersList.x() + pd.playersList.width() + PLAYER_LOADING_STATUS_ITEM_SPACING
        y0 = pd.titleLabel.y() + pd.titleLabel.height() + PLAYER_LOADING_STATUS_ITEM_SPACING
        height = psParent.height() - y0
        psWidget.setGeometry x0, y0, psWidget.idealWidth(), height

    Case FN_pregame_optionsButton_onClick       ' [weakPlayerStatusButton, onClickFunc](W_BUTTON& button)
        Set button = args(0)
        Set strongParent = WIDGET_fromPtr(f.cap(0))
        If Not strongParent Is Nothing Then
            Set sd = WzPlayerStatusCheckboxButton_(strongParent)
            widgRemoveOverlayScreen sd.optionsOverlayScreen
            WzFunc_call f.cap(1), button
            Set sd.optionsOverlayScreen = Nothing
        End If

    Case FN_pregame_kickButton_onClick          ' [weakPlayerStatusButton](W_BUTTON& button)
        Set strongParent = WIDGET_fromPtr(f.cap(0))
        If Not strongParent Is Nothing Then
            Set sd = WzPlayerStatusCheckboxButton_(strongParent)
            targetPlayer = sd.player
            If isHumanPlayer(targetPlayer) And targetPlayer <> NetPlay_hostPlayer And sd.status = LoadingStatus_Loading Then
                numHumanPlayers = 0
                For i = 0 To MAX_PLAYERS - 1
                    If isHumanPlayer(i) Then
                        numHumanPlayers = numHumanPlayers + 1
                    End If
                Next
                If NetPlay_players_isSpectator(targetPlayer) Or (numHumanPlayers <= 2 And Not NetPlay_players_isSpectator(NetPlay_hostPlayer)) Then
                    kickPlayer targetPlayer, "The host has kicked you from the game.", ERROR_KICKED, False
                Else
                    startKickVote targetPlayer
                End If
            End If
        End If

    Case FN_pregame_optionsOverlay_clicked      ' [psWeakOptionsOverlayScreen, psWeakPlayerStatusButton]()
        Set psOverlayScreen = W_SCREEN_fromPtr(f.cap(0))
        If Not psOverlayScreen Is Nothing Then
            widgRemoveOverlayScreen psOverlayScreen
        End If
        ' Destroy Options overlay / overlay screen, reset button state
        Set strongParent = WIDGET_fromPtr(f.cap(1))
        If Not strongParent Is Nothing Then
            Set WzPlayerStatusCheckboxButton_(strongParent).optionsOverlayScreen = Nothing
            strongParent.setState 0
        End If

    Case FN_pregame_optionsPopOver_calcLayout   ' [weakParent](WIDGET *psWidget)
        Set psWidget = args(0)
        Set psParent = WIDGET_fromPtr(f.cap(0))
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/screens/netpregamescreen.cpp:397", "displayHostOptionsOverlay", "parent is null": Exit Function
        ' (Ideally, with its left aligned with the right side of the "parent" widget, but ensure full visibility on-screen)
        popOverX0 = psParent.screenPosX() + psParent.width() + 2
        If popOverX0 + psWidget.width() > screenWidth Then
            popOverX0 = screenWidth - psWidget.width()
        End If
        ' (Ideally, with its top aligned with the top of the "parent" widget, but ensure full visibility on-screen)
        popOverY0 = psParent.screenPosY() - 4
        If popOverY0 + psWidget.height() > screenHeight Then
            popOverY0 = screenHeight - psWidget.height()
        End If
        psWidget.move popOverX0, popOverY0

    Case FN_pregame_clickForm_calcLayout        ' LAMBDA_CALCLAYOUT_SIMPLE (WIDGET *psWidget)
        Set psWidget = args(0)
        psWidget.setGeometry 0, 0, screenWidth, screenHeight

    Case FN_pregame_consoleMessage              ' [&](ConsoleMessage const &message) (sender, text, justification, team, duration)
        Set self = WIDGET_fromPtr(f.cap(0))
        If self Is Nothing Then Exit Function
        Set d = WzGameStartOverlayScreen_CLICKFORM_(self)
        sender = args(0)
        text = args(1)
        justification = args(2)
        duration = args(4)
        addedTime = steadyNow()
        If duration > 0 Then messageDuration = duration Else messageDuration = 5000
        messageDuration = MIN_(messageDuration, 5000)   ' std::min<uint32_t>
        ' currentConsoleMessages.emplace_back(message.text, font_regular, addedTime, std::chrono::milliseconds(messageDuration), message.justification, message.sender)
        Set m = New CachedConsoleMessage
        m.display.setText text, font_regular
        m.timeAdded = addedTime
        m.duration = messageDuration
        m.JustifyType = justification
        m.player = sender
        d.currentConsoleMessages.Add m
        If d.currentConsoleMessages.Count > 5 Then
            d.currentConsoleMessages.Remove 1
        End If
    End Select
End Function
