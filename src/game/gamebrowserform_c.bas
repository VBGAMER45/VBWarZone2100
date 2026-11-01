Attribute VB_Name = "gamebrowserform_c"
' src/titleui/widgets/gamebrowserform.cpp, gamebrowserform.h, whole: the game browser (Multiplayer > Join Game): the
' lobby's games list, its filters, the "GameID or IP:port" box with its join menu (Last IP) and the lobby status
' overlay. Its widget classes' fields are the data classes named as the classes. The weak pointers that a scheduled
' task, a popover's close handler or wzAsyncExecOnMainThread can outlive are widgweak_vb serials. netlobby (curl)
' is not ported: its stand-ins answer as when the lobby cannot be reached ("Failed to connect to lobby server");
' a netlobby::ListResult reaches the lambdas as a store handle. A std::vector<std::shared_ptr<WIDGET>> is a
' Collection; a WzSize is its two Longs. if constexpr (fnv1a_hash(GetDefaultLobbyAddress()) ==
' LOB_DEFAULT_BEHAVIOR_CHECKSUM) is computed when it runs. Translations (_()) are not ported: the English text.
' (VB) widget class LobbyStatusOverlayWidget : WIDGET
' (VB) widget class LobbyGameIconWidget : WIDGET
' (VB) widget class LobbyGamePlayersWidget : WIDGET
' (VB) widget class LobbyGameNameWidget : WIDGET
' (VB) widget class LobbyGameMapWidget : WIDGET
' (VB) widget class LobbyGameModsWidget : WIDGET
' (VB) widget class LobbyActionWidgets : WIDGET
' (VB) widget class ConnectToForm : W_BUTTON
' (VB) widget class LobbyBrowser : WIDGET
Option Explicit

Private Const LOB_DEFAULT_BEHAVIOR_CHECKSUM As Double = 2462613339#  ' static constexpr uint32_t

Private Const verticalBetweenItemsPadding As Long = 20  ' LobbyStatusOverlayWidget's const int

Private Const WZLOBBY_WIDG_PADDING_Y As Long = 4
Private Const WZLOBBY_WIDG_PADDING_X As Long = 0
Private Const WZLOBBY_WIDG_LINE_SPACING As Long = 2

' ConnectToForm's consts
Private Const minInputBoxHeight As Long = 16
Private Const inputBoxHorizontalPadding As Long = 6
Private Const verticalPadding As Long = 5
Private Const horizontalPadding As Long = 8
Private Const maxButtonTextAreaWidth As Long = 100
Private Const imageDimensions As Long = 12
Private Const horizontalPaddingBetweenButtonSplit As Long = 6

Private Const maxConnectToInputLength As Long = 256     ' constexpr int

Private Const connectToFormHorizontalPadding As Long = 15  ' LobbyBrowser's const int

Private Const WZLOBBY_TABLE_COL_PADDING_X As Long = 5
Private Const WZLOBBY_TABLE_ROW_SPACING As Long = 2

Private Const ALLIANCE_TYPE_MAX As Long = ALLIANCES_UNSHARED       ' multiplaydefs.h
Private Const BLIND_MODE_MAX As Long = BLIND_MODE_BLIND_GAME_SIMPLE_LOBBY

' the lambdas (WzFunc: gamebrowserform_c_fn)
Public Const FN_lobbyStatusOverlay_resetClick As Long = 16100        ' LobbyStatusOverlayWidget::initialize: resetLobbyAddressButton->addOnClickHandler([weakSelf](W_BUTTON& but) {...})
Public Const FN_connectToForm_joinAsPlayer As Long = 16101           ' buildJoinOptionsMenu: [weakSelf](WIDGET_KEY) { ...submitInputBox(false); }
Public Const FN_connectToForm_joinAsSpectator As Long = 16102        ' buildJoinOptionsMenu: [weakSelf](WIDGET_KEY) { ...submitInputBox(true); }
Public Const FN_connectToForm_lastIPJoinAsPlayer As Long = 16103     ' buildJoinOptionsMenu: [weakSelf, lastIPServerName](WIDGET_KEY) { ...submit(..., false); }
Public Const FN_connectToForm_lastIPJoinAsSpectator As Long = 16104  ' buildJoinOptionsMenu: [weakSelf, lastIPServerName](WIDGET_KEY) { ...submit(..., true); }
Public Const FN_connectToForm_clearLastIP As Long = 16105            ' buildJoinOptionsMenu: [](WIDGET_KEY) { war_setLastIpServerConnect(""); }
Public Const FN_connectToForm_menuClosed As Long = 16106             ' displayJoinOptionsMenu: [weakSelf]() { // close handler }
Public Const FN_displayConnectToEditBox As Long = 16107              ' void displayConnectToEditBox(WIDGET *psWidget, UDWORD xOffset, UDWORD yOffset)
Public Const FN_connectToForm_pasteTransform As Long = 16108         ' initialize: inputBox->setOnPasteTransformFunc([](const WzString& pastedString) -> WzString {...})
Public Const FN_connectToForm_onReturn As Long = 16109               ' initialize: inputBox->setOnReturnHandler([weakSelf](W_EDITBOX& widg) {...})
Public Const FN_connectToForm_click As Long = 16110                  ' initialize: addOnClickHandler([weakSelf](W_BUTTON&) {...})
Public Const FN_connectToForm_simulateClickTask As Long = 16111      ' clicked: widgScheduleTask([weakSelf]() {...})
Public Const FN_addClickableJoinMenuItem_click As Long = 16112       ' addClickableJoinMenuItem: [onClick](WzClickableOptionsChoiceWidget&, WIDGET_KEY key) {...}
Public Const FN_lobbyBrowser_listResult As Long = 16113              ' triggerAsyncGameListFetch: [weakSelf](netlobby::ListResult results) {...}
Public Const FN_lobbyBrowser_listResultMainThread As Long = 16114    ' triggerAsyncGameListFetch: wzAsyncExecOnMainThread([weakSelf, results = std::move(results)]() mutable {...})
Public Const FN_lobbyBrowser_rowClick As Long = 16115                ' populateTableFromGameList: row->addOnClickHandler([weakSelf, idx](W_BUTTON&) {...})
Public Const FN_lobbyBrowser_spectate As Long = 16116                ' createLobbyGameRowColumnWidgets: onSpectateFunc = [weakSelf, idx]() {...}
Public Const FN_lobbyActionWidgets_watchClick As Long = 16117        ' LobbyActionWidgets::initialize: watchButton->addOnClickHandler([onSpectateHandler](W_BUTTON&) {...})
Public Const FN_lobbyBrowser_backClick As Long = 16118               ' initialize: backButton->addOnClickHandler([onBackButtonFunc](W_BUTTON& button) {...})
Public Const FN_lobbyBrowser_backTask As Long = 16119                ' initialize: widgScheduleTask([onBackButtonFunc]() {...})
Public Const FN_lobbyBrowser_connectSubmit As Long = 16120           ' initialize: ConnectToForm::make([](WzString inputStr, bool asSpectator) {...})
Public Const FN_lobbyBrowser_filtersClick As Long = 16121            ' initialize: filtersButton->addOnClickHandler([weakSelf](W_BUTTON& but) {...})
Public Const FN_lobbyBrowser_refreshClick As Long = 16122            ' initialize: refreshButton->addOnClickHandler([weakSelf](W_BUTTON& but) {...})
Public Const FN_lobbyBrowser_resetLobbyAddress As Long = 16123       ' initialize: LobbyStatusOverlayWidget::make([weakSelf]() {...})
Public Const FN_addPopoverMenuOptionsCheckbox_click As Long = 16124  ' addPopoverMenuOptionsCheckbox: pCheckbox->addOnClickHandler([onClickFunc](W_BUTTON& button){...})
Public Const FN_lobbyBrowser_filterIncompatible As Long = 16125      ' createFiltersPopoverForm: [weakSelf](WzAdvCheckbox& button){...}
Public Const FN_lobbyBrowser_filterEmpty As Long = 16126             ' createFiltersPopoverForm: [weakSelf](WzAdvCheckbox& button){...}
Public Const FN_lobbyBrowser_filterModded As Long = 16127            ' createFiltersPopoverForm: [weakSelf](WzAdvCheckbox& button){...}
Public Const FN_lobbyBrowser_filterIPv6Only As Long = 16128          ' createFiltersPopoverForm: [weakSelf](WzAdvCheckbox& button){...}
Public Const FN_lobbyBrowser_filtersMenuClosed As Long = 16129       ' displayFiltersOverlay: [psWeakSelf](){...}

'--------------------------------------------------------------------------
' (VB) helpers

Private Function gmax(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then gmax = a Else gmax = b
End Function

Private Function gmin(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then gmin = a Else gmin = b
End Function

Private Function pielightAlpha(ByVal c As Long) As Long
    pielightAlpha = SHRU(c, 24)
End Function

Private Function pielightWithAlpha(ByVal c As Long, ByVal a As Long) As Long
    pielightWithAlpha = (c And &HFFFFFF) Or SHL(a And &HFF&, 24)
End Function

' (VB) FrontImages == nullptr (init_c's Variant: Empty, or a loader's Nothing)
Private Function isFrontImagesNull() As Boolean
    If IsObject(FrontImages) Then
        isFrontImagesNull = (FrontImages Is Nothing)
    Else
        isFrontImagesNull = IsEmpty(FrontImages)
    End If
End Function

' string_ext.h: constexpr uint32_t fnv1a_hash(std::string_view sv) (the result as an unsigned value)
Private Function fnv1a_hash(ByVal sv As String) As Double
    Dim b() As Byte, n As Long, i As Long, hash As Long
    n = utf8_encode(sv, b)
    hash = &H811C9DC5
    For i = 0 To n - 1
        hash = hash Xor b(i)
        hash = MUL32(hash, &H1000193)
    Next
    fnv1a_hash = U32(hash)
End Function

' std::isspace (the classic locale)
Private Function isSpaceChar(ByVal c As String) As Boolean
    isSpaceChar = (c = " " Or c = vbTab Or c = vbLf Or c = vbVerticalTab Or c = vbFormFeed Or c = vbCr)
End Function

' MARK: - Lobby Status Overlay Widget

Public Sub LobbyStatusOverlayWidget_construct(ByVal self As WIDGET)
    WIDGET_construct self
    self.kind = K_LobbyStatusOverlayWidget ' (VB) virtual calls from here on reach LobbyStatusOverlayWidget's overrides
    LobbyStatusOverlayWidget_(self).weakSerial = WeakWidget_register(self)
End Sub

Public Sub LobbyStatusOverlayWidget_destruct(ByVal self As WIDGET)
    WeakWidget_unregister LobbyStatusOverlayWidget_(self).weakSerial
End Sub

' static std::shared_ptr<LobbyStatusOverlayWidget> make(const OnResetLobbyAddressClickFunc& onResetLobbyAddressClickFunc)
Public Function LobbyStatusOverlayWidget_make(ByVal onResetLobbyAddressClickFunc As WzFunc) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_LobbyStatusOverlayWidget)
    LobbyStatusOverlayWidget_construct widget
    LobbyStatusOverlayWidget_initialize widget, onResetLobbyAddressClickFunc
    Set LobbyStatusOverlayWidget_make = widget
End Function

Private Sub LobbyStatusOverlayWidget_initialize(ByVal self As WIDGET, ByVal onResetLobbyAddressClickFunc_ As WzFunc)
    Dim d As LobbyStatusOverlayWidget, weakSelf As Long
    Set d = LobbyStatusOverlayWidget_(self)
    Set d.onResetLobbyAddressClickFunc = onResetLobbyAddressClickFunc_

    Set d.textLabel = W_LABEL_new()
    W_LABEL_setFont d.textLabel, font_medium_bold, psPalette(WZCOL_TEXT_MEDIUM)
    W_LABEL_setTextAlignment d.textLabel, WLAB_ALIGNTOP
    W_LABEL_setCanTruncate d.textLabel, True
    d.textLabel.setTransparentToClicks True
    self.attach d.textLabel

    Set d.indeterminateIndicator = createJoiningIndeterminateProgressWidget(font_medium_bold)
    d.indeterminateIndicator.show False
    self.attach d.indeterminateIndicator

    Set d.resetLobbyAddressButton = WzFrontendImageButton_make(True, IMAGE_ARROW_UNDO)
    d.resetLobbyAddressButton.setString "Reset Lobby Address to Default"
    WzFrontendImageButton_setPadding d.resetLobbyAddressButton, 6, 5
    WzFrontendImageButton_setBackgroundColor d.resetLobbyAddressButton, True, psPalette(WZCOL_TRANSPARENT_BOX)
    d.resetLobbyAddressButton.setGeometry 0, 0, d.resetLobbyAddressButton.idealWidth(), d.resetLobbyAddressButton.idealHeight()
    weakSelf = d.weakSerial
    W_BUTTON_addOnClickHandler d.resetLobbyAddressButton, WzFunc_new(FN_lobbyStatusOverlay_resetClick, weakSelf)
    d.resetLobbyAddressButton.show False
    self.attach d.resetLobbyAddressButton
End Sub

Public Function LobbyStatusOverlayWidget_idealWidth(ByVal self As WIDGET) As Long
    Dim d As LobbyStatusOverlayWidget
    Set d = LobbyStatusOverlayWidget_(self)
    LobbyStatusOverlayWidget_idealWidth = gmax(d.textLabel.idealWidth(), d.indeterminateIndicator.idealWidth())
End Function

Public Function LobbyStatusOverlayWidget_idealHeight(ByVal self As WIDGET) As Long
    Dim d As LobbyStatusOverlayWidget, result As Long
    Set d = LobbyStatusOverlayWidget_(self)
    result = 0
    If Len(d.textLabel.getString()) <> 0 Then
        result = result + iV_GetTextLineSize(font_medium_bold)
    End If
    If d.indeterminateIndicator.visible() Then
        If result > 0 Then
            result = result + verticalBetweenItemsPadding
        End If
        result = result + d.indeterminateIndicator.idealHeight()
    End If
    If d.resetLobbyAddressButton.visible() Then
        If result > 0 Then
            result = result + verticalBetweenItemsPadding
        End If
        result = result + d.resetLobbyAddressButton.idealHeight()
    End If
    LobbyStatusOverlayWidget_idealHeight = result
End Function

Private Sub LobbyStatusOverlayWidget_recalcLayout(ByVal self As WIDGET)
    Dim d As LobbyStatusOverlayWidget, w As Long, nextWidgY0 As Long, labelWidth As Long, labelX0 As Long
    Dim indicatorX0 As Long, resetButtonWidth As Long, resetButtonX0 As Long
    Set d = LobbyStatusOverlayWidget_(self)
    w = self.width()

    nextWidgY0 = 0
    If Len(d.textLabel.getString()) <> 0 Then
        labelWidth = gmin(w, d.textLabel.idealWidth())
        labelX0 = (w - labelWidth) \ 2
        d.textLabel.setGeometry labelX0, nextWidgY0, labelWidth, d.textLabel.idealHeight()
        nextWidgY0 = d.textLabel.height() + verticalBetweenItemsPadding
    End If

    If d.indeterminateIndicator.visible() Then
        indicatorX0 = (w - d.indeterminateIndicator.idealWidth()) \ 2
        d.indeterminateIndicator.setGeometry indicatorX0, nextWidgY0, d.indeterminateIndicator.idealWidth(), d.indeterminateIndicator.idealHeight()
        nextWidgY0 = d.indeterminateIndicator.y() + d.indeterminateIndicator.height() + verticalBetweenItemsPadding
    End If

    If d.resetLobbyAddressButton.visible() Then
        resetButtonWidth = gmin(d.resetLobbyAddressButton.idealWidth(), w)
        resetButtonX0 = (w - resetButtonWidth) \ 2
        d.resetLobbyAddressButton.setGeometry resetButtonX0, nextWidgY0, resetButtonWidth, d.resetLobbyAddressButton.idealHeight()
    End If
End Sub

Public Sub LobbyStatusOverlayWidget_geometryChanged(ByVal self As WIDGET)
    LobbyStatusOverlayWidget_recalcLayout self
End Sub

Public Sub LobbyStatusOverlayWidget_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As LobbyStatusOverlayWidget, x0 As Long, y0 As Long
    Dim indicatorX0 As Long, indicatorY0 As Long, indicatorX1 As Long, indicatorY1 As Long
    Set d = LobbyStatusOverlayWidget_(self)
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset

    If d.indeterminateIndicator.visible() Then
        indicatorX0 = x0 + d.indeterminateIndicator.x()
        indicatorY0 = y0 + d.indeterminateIndicator.y()
        indicatorX1 = indicatorX0 + d.indeterminateIndicator.width()
        indicatorY1 = indicatorY0 + d.indeterminateIndicator.height()

        pie_UniTransBoxFill gmax(indicatorX0 - 20, 0), gmax(indicatorY0 - 10, 0), gmin(indicatorX1 + 20, screenWidth), gmin(indicatorY1 + 20, screenHeight), psPalette(WZCOL_TRANSPARENT_BOX)
    End If
End Sub

Public Sub LobbyStatusOverlayWidget_showIndeterminateIndicator(ByVal self As WIDGET, ByVal show As Boolean)
    Dim d As LobbyStatusOverlayWidget
    Set d = LobbyStatusOverlayWidget_(self)
    If d.indeterminateIndicator.visible() = show Then
        Exit Sub
    End If
    d.indeterminateIndicator.show show
    LobbyStatusOverlayWidget_recalcLayout self
End Sub

Public Sub LobbyStatusOverlayWidget_setString(ByVal self As WIDGET, ByVal string_ As String)
    LobbyStatusOverlayWidget_(self).textLabel.setString string_
    LobbyStatusOverlayWidget_recalcLayout self
End Sub

Public Sub LobbyStatusOverlayWidget_showResetLobbyAddressButton(ByVal self As WIDGET, ByVal show As Boolean)
    LobbyStatusOverlayWidget_(self).resetLobbyAddressButton.show show
    LobbyStatusOverlayWidget_recalcLayout self
End Sub

' MARK: - Lobby Game Column Widgets

' LobbyGameIconWidget() { }
Public Sub LobbyGameIconWidget_construct(ByVal self As WIDGET)
    Dim d As LobbyGameIconWidget
    WIDGET_construct self
    self.kind = K_LobbyGameIconWidget ' (VB) virtual calls from here on reach LobbyGameIconWidget's overrides
    Set d = LobbyGameIconWidget_(self)
    d.has_frontendImgID = False         ' = nullopt
    d.has_customImgColor = False        ' = nullopt
    d.missingImage = False
    d.imageWidth = 16                   ' WzSize imageDimensions = WzSize(16,16);
    d.imageHeight = 16
End Sub

' static std::shared_ptr<LobbyGameIconWidget> make(optional<UWORD> frontendImgId)
Public Function LobbyGameIconWidget_make(ByVal has_frontendImgId As Boolean, ByVal frontendImgId As Long) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_LobbyGameIconWidget)
    LobbyGameIconWidget_construct widget
    LobbyGameIconWidget_initialize widget, has_frontendImgId, frontendImgId
    Set LobbyGameIconWidget_make = widget
End Function

Private Sub LobbyGameIconWidget_initialize(ByVal self As WIDGET, ByVal has_frontendImgId As Boolean, ByVal frontendImgId As Long)
    LobbyGameIconWidget_setImage self, has_frontendImgId, frontendImgId
    self.setGeometry 0, 0, self.idealWidth(), self.idealHeight()
End Sub

Public Function LobbyGameIconWidget_idealWidth(ByVal self As WIDGET) As Long
    LobbyGameIconWidget_idealWidth = LobbyGameIconWidget_(self).imageWidth + (WZLOBBY_WIDG_PADDING_X * 2)
End Function

Public Function LobbyGameIconWidget_idealHeight(ByVal self As WIDGET) As Long
    LobbyGameIconWidget_idealHeight = LobbyGameIconWidget_(self).imageHeight + (WZLOBBY_WIDG_PADDING_Y * 2)
End Function

Public Sub LobbyGameIconWidget_setImage(ByVal self As WIDGET, ByVal has_newFrontendImgID As Boolean, ByVal newFrontendImgID As Long)
    Dim d As LobbyGameIconWidget
    Set d = LobbyGameIconWidget_(self)
    d.missingImage = False

    ' validate newFrontendImgID value
    If has_newFrontendImgID Then
        If newFrontendImgID >= IMAGEFILE_numImages(FrontImages) Then
            ' out of bounds value (likely a mod has an incomplete / older frontend.img file)
            debug_ 285, LOG_ERROR, "setImage", "frontend.img missing expected image id ""%u"" (num loaded: %zu) - remove any mods", newFrontendImgID, IMAGEFILE_numImages(FrontImages)
            has_newFrontendImgID = False
            d.missingImage = True
        End If
    End If

    d.has_frontendImgID = has_newFrontendImgID
    d.frontendImgID = newFrontendImgID
End Sub

Public Sub LobbyGameIconWidget_setCustomImageColor(ByVal self As WIDGET, ByVal has_color As Boolean, ByVal color As Long)
    Dim d As LobbyGameIconWidget
    Set d = LobbyGameIconWidget_(self)
    d.has_customImgColor = has_color
    d.customImgColor = color
End Sub

Public Sub LobbyGameIconWidget_setImageSize(ByVal self As WIDGET, ByVal newImageWidth As Long, ByVal newImageHeight As Long)
    Dim d As LobbyGameIconWidget
    Set d = LobbyGameIconWidget_(self)
    d.imageWidth = newImageWidth
    d.imageHeight = newImageHeight
End Sub

Public Sub LobbyGameIconWidget_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As LobbyGameIconWidget, x0 As Long, y0 As Long, w As Long, imgPosX0 As Long, imgPosY0 As Long, imgColor As Long
    Set d = LobbyGameIconWidget_(self)
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    w = self.width()

    ' Display the image horizontally centered
    imgPosX0 = x0 + (w - d.imageWidth) \ 2
    ' Display the image top-aligned
    imgPosY0 = y0 + WZLOBBY_WIDG_PADDING_Y

    If d.has_frontendImgID Then
        If d.has_customImgColor Then imgColor = d.customImgColor Else imgColor = psPalette(WZCOL_TEXT_MEDIUM)
        iV_DrawImageFileAnisotropicTint FrontImages, d.frontendImgID, imgPosX0, imgPosY0, d.imageWidth, d.imageHeight, imgColor
    ElseIf d.missingImage Then
        pie_UniTransBoxFill imgPosX0, imgPosY0, imgPosX0 + d.imageWidth, imgPosY0 + d.imageHeight, psPalette(WZCOL_RED)
    End If
End Sub

Public Sub LobbyGameIconWidget_setTip(ByVal self As WIDGET, ByVal string_ As String)
    LobbyGameIconWidget_(self).tipStr = string_
End Sub

Public Function LobbyGameIconWidget_getTip(ByVal self As WIDGET) As String
    LobbyGameIconWidget_getTip = LobbyGameIconWidget_(self).tipStr
End Function

' LobbyGamePlayersWidget() { }
Public Sub LobbyGamePlayersWidget_construct(ByVal self As WIDGET)
    WIDGET_construct self
    self.kind = K_LobbyGamePlayersWidget ' (VB) virtual calls from here on reach LobbyGamePlayersWidget's overrides
End Sub

' static std::shared_ptr<LobbyGamePlayersWidget> make(uint8_t currentPlayers, uint8_t maxPlayers)
Public Function LobbyGamePlayersWidget_make(ByVal currentPlayers As Long, ByVal maxPlayers As Long) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_LobbyGamePlayersWidget)
    LobbyGamePlayersWidget_construct widget
    LobbyGamePlayersWidget_initialize widget, currentPlayers, maxPlayers
    Set LobbyGamePlayersWidget_make = widget
End Function

Private Sub LobbyGamePlayersWidget_initialize(ByVal self As WIDGET, ByVal currentPlayers As Long, ByVal maxPlayers As Long)
    Dim d As LobbyGamePlayersWidget
    Set d = LobbyGamePlayersWidget_(self)
    Set d.playersCount = W_LABEL_new()
    W_LABEL_setFont d.playersCount, font_regular, psPalette(WZCOL_FORM_LIGHT)
    W_LABEL_setTextAlignment d.playersCount, WLAB_ALIGNTOP
    W_LABEL_setCanTruncate d.playersCount, True
    d.playersCount.setTransparentToClicks True
    self.attach d.playersCount

    LobbyGamePlayersWidget_updateValue self, currentPlayers, maxPlayers

    self.setGeometry 0, 0, self.idealWidth(), self.idealHeight()
End Sub

Public Sub LobbyGamePlayersWidget_geometryChanged(ByVal self As WIDGET)
    Dim w As Long, h As Long
    w = self.width()
    h = self.height()
    LobbyGamePlayersWidget_(self).playersCount.setGeometry WZLOBBY_WIDG_PADDING_X, WZLOBBY_WIDG_PADDING_Y, w - (WZLOBBY_WIDG_PADDING_X * 2), h - (WZLOBBY_WIDG_PADDING_Y * 2)
End Sub

Public Sub LobbyGamePlayersWidget_updateValue(ByVal self As WIDGET, ByVal currentPlayers As Long, ByVal maxPlayers As Long)
    Dim d As LobbyGamePlayersWidget, str_ As String
    Set d = LobbyGamePlayersWidget_(self)
    str_ = CStr(currentPlayers) & "/" & CStr(maxPlayers)
    d.playersCount.setString str_
    If currentPlayers = maxPlayers Then
        W_LABEL_setFontColour d.playersCount, psPalette(WZCOL_TEXT_MEDIUM)
    Else
        W_LABEL_setFontColour d.playersCount, psPalette(WZCOL_FORM_LIGHT)
    End If
End Sub

Public Function LobbyGamePlayersWidget_idealWidth(ByVal self As WIDGET) As Long
    LobbyGamePlayersWidget_idealWidth = LobbyGamePlayersWidget_(self).playersCount.idealWidth() + (WZLOBBY_WIDG_PADDING_X * 2)
End Function

Public Function LobbyGamePlayersWidget_idealHeight(ByVal self As WIDGET) As Long
    Dim result As Long
    result = (WZLOBBY_WIDG_PADDING_Y * 2)
    result = result + LobbyGamePlayersWidget_(self).playersCount.idealHeight()
    LobbyGamePlayersWidget_idealHeight = result
End Function

' LobbyGameNameWidget() { }
Public Sub LobbyGameNameWidget_construct(ByVal self As WIDGET)
    WIDGET_construct self
    self.kind = K_LobbyGameNameWidget ' (VB) virtual calls from here on reach LobbyGameNameWidget's overrides
End Sub

' static std::shared_ptr<LobbyGameNameWidget> make(const WzString& roomName, const WzString& hostName, const std::string& hostPublicIdentity)
Public Function LobbyGameNameWidget_make(ByVal roomName As String, ByVal hostName As String, ByVal hostPublicIdentity As String) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_LobbyGameNameWidget)
    LobbyGameNameWidget_construct widget
    LobbyGameNameWidget_initialize widget, roomName, hostName, hostPublicIdentity
    Set LobbyGameNameWidget_make = widget
End Function

Private Sub LobbyGameNameWidget_initialize(ByVal self As WIDGET, ByVal roomName As String, ByVal hostName As String, ByVal hostPublicIdentity As String)
    Dim d As LobbyGameNameWidget, tipStr As String
    Set d = LobbyGameNameWidget_(self)
    Set d.roomNameWidg = W_LABEL_new()
    W_LABEL_setFont d.roomNameWidg, font_regular, psPalette(WZCOL_TEXT_BRIGHT)
    W_LABEL_setCanTruncate d.roomNameWidg, True
    d.roomNameWidg.setTransparentToClicks True
    self.attach d.roomNameWidg

    Set d.hostNameWidg = W_LABEL_new()
    W_LABEL_setFont d.hostNameWidg, font_regular, psPalette(WZCOL_TEXT_MEDIUM)
    W_LABEL_setCanTruncate d.hostNameWidg, True
    d.hostNameWidg.setTransparentToClicks True
    self.attach d.hostNameWidg
    If Len(hostPublicIdentity) <> 0 Then
        tipStr = "Host Name: " & hostName
        tipStr = tipStr & vbLf
        tipStr = tipStr & "ID: " & hostPublicIdentity
        d.hostNameWidg.setTip tipStr
    End If

    LobbyGameNameWidget_updateValue self, roomName, hostName
    self.setGeometry 0, 0, self.idealWidth(), self.idealHeight()
End Sub

Public Sub LobbyGameNameWidget_geometryChanged(ByVal self As WIDGET)
    Dim d As LobbyGameNameWidget, w As Long, nextLineY0 As Long
    Set d = LobbyGameNameWidget_(self)
    w = self.width()
    d.roomNameWidg.setGeometry WZLOBBY_WIDG_PADDING_X, WZLOBBY_WIDG_PADDING_Y, w - (WZLOBBY_WIDG_PADDING_X * 2), d.roomNameWidg.idealHeight()
    nextLineY0 = d.roomNameWidg.y() + d.roomNameWidg.height() + WZLOBBY_WIDG_LINE_SPACING
    d.hostNameWidg.setGeometry WZLOBBY_WIDG_PADDING_X, nextLineY0, w - (WZLOBBY_WIDG_PADDING_X * 2), d.hostNameWidg.idealHeight()
End Sub

Public Function LobbyGameNameWidget_idealWidth(ByVal self As WIDGET) As Long
    Dim d As LobbyGameNameWidget
    Set d = LobbyGameNameWidget_(self)
    LobbyGameNameWidget_idealWidth = gmax(d.roomNameWidg.idealWidth(), d.hostNameWidg.idealWidth()) + (WZLOBBY_WIDG_PADDING_X * 2)
End Function

Public Function LobbyGameNameWidget_idealHeight(ByVal self As WIDGET) As Long
    Dim d As LobbyGameNameWidget, result As Long
    Set d = LobbyGameNameWidget_(self)
    result = (WZLOBBY_WIDG_PADDING_Y * 2)
    result = result + d.roomNameWidg.idealHeight() + WZLOBBY_WIDG_LINE_SPACING + d.hostNameWidg.idealHeight()
    LobbyGameNameWidget_idealHeight = result
End Function

Public Sub LobbyGameNameWidget_updateValue(ByVal self As WIDGET, ByVal roomName As String, ByVal hostName As String)
    Dim d As LobbyGameNameWidget
    Set d = LobbyGameNameWidget_(self)
    d.roomNameWidg.setString roomName
    d.hostNameWidg.setString hostName
End Sub

' LobbyGameMapWidget() { }
Public Sub LobbyGameMapWidget_construct(ByVal self As WIDGET)
    WIDGET_construct self
    self.kind = K_LobbyGameMapWidget ' (VB) virtual calls from here on reach LobbyGameMapWidget's overrides
    LobbyGameMapWidget_(self).blindMode_ = BLIND_MODE_NONE
End Sub

' static std::shared_ptr<LobbyGameMapWidget> make(const WzString& mapName, optional<AllianceType> alliancesMode, BLIND_MODE blindMode)
Public Function LobbyGameMapWidget_make(ByVal mapName As String, ByVal has_alliancesMode As Boolean, ByVal alliancesMode As Long, ByVal blindMode As Long) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_LobbyGameMapWidget)
    LobbyGameMapWidget_construct widget
    LobbyGameMapWidget_initialize widget, mapName, has_alliancesMode, alliancesMode, blindMode
    Set LobbyGameMapWidget_make = widget
End Function

Private Sub LobbyGameMapWidget_initialize(ByVal self As WIDGET, ByVal mapName As String, ByVal has_alliancesMode As Boolean, ByVal alliancesMode As Long, ByVal blindMode As Long)
    Dim d As LobbyGameMapWidget
    Set d = LobbyGameMapWidget_(self)
    Set d.mapNameWidg = W_LABEL_new()
    W_LABEL_setFont d.mapNameWidg, font_regular, psPalette(WZCOL_TEXT_MEDIUM)
    W_LABEL_setCanTruncate d.mapNameWidg, True
    d.mapNameWidg.setTransparentToClicks True
    self.attach d.mapNameWidg

    Set d.modeDetailsWidg = W_LABEL_new()
    W_LABEL_setFont d.modeDetailsWidg, font_small, psPalette(WZCOL_TEXT_MEDIUM)
    W_LABEL_setCanTruncate d.modeDetailsWidg, True
    d.modeDetailsWidg.setTransparentToClicks True
    self.attach d.modeDetailsWidg

    LobbyGameMapWidget_updateValue self, mapName, has_alliancesMode, alliancesMode, blindMode
    self.setGeometry 0, 0, self.idealWidth(), self.idealHeight()
End Sub

Public Sub LobbyGameMapWidget_geometryChanged(ByVal self As WIDGET)
    Dim d As LobbyGameMapWidget, w As Long, nextLineY0 As Long
    Set d = LobbyGameMapWidget_(self)
    w = self.width()
    d.mapNameWidg.setGeometry WZLOBBY_WIDG_PADDING_X, WZLOBBY_WIDG_PADDING_Y, w - (WZLOBBY_WIDG_PADDING_X * 2), d.mapNameWidg.idealHeight()
    nextLineY0 = d.mapNameWidg.y() + d.mapNameWidg.height() + WZLOBBY_WIDG_LINE_SPACING
    d.modeDetailsWidg.setGeometry WZLOBBY_WIDG_PADDING_X, nextLineY0, w - (WZLOBBY_WIDG_PADDING_X * 2), d.modeDetailsWidg.idealHeight()
End Sub

Public Function LobbyGameMapWidget_idealWidth(ByVal self As WIDGET) As Long
    Dim d As LobbyGameMapWidget
    Set d = LobbyGameMapWidget_(self)
    LobbyGameMapWidget_idealWidth = gmax(d.mapNameWidg.idealWidth(), d.modeDetailsWidg.idealWidth()) + (WZLOBBY_WIDG_PADDING_X * 2)
End Function

Public Function LobbyGameMapWidget_idealHeight(ByVal self As WIDGET) As Long
    Dim d As LobbyGameMapWidget, result As Long
    Set d = LobbyGameMapWidget_(self)
    result = (WZLOBBY_WIDG_PADDING_Y * 2)
    result = result + d.mapNameWidg.idealHeight() + WZLOBBY_WIDG_LINE_SPACING + d.modeDetailsWidg.idealHeight()
    LobbyGameMapWidget_idealHeight = result
End Function

' static const char* allianceTypeToLobbyGameDisplayString(optional<AllianceType> v)
Private Function allianceTypeToLobbyGameDisplayString(ByVal has_v As Boolean, ByVal v As Long) As String
    If Not has_v Then
        allianceTypeToLobbyGameDisplayString = ""
        Exit Function
    End If
    Select Case v
    Case NO_ALLIANCES
        ' TRANSLATORS: "Free-for-all" acronym / abbreviation
        allianceTypeToLobbyGameDisplayString = "FFA"
        Exit Function
    Case ALLIANCES_
        allianceTypeToLobbyGameDisplayString = "Alliances Allowed"
        Exit Function
    Case ALLIANCES_TEAMS
        allianceTypeToLobbyGameDisplayString = "Teams"
        Exit Function
    Case ALLIANCES_UNSHARED
        allianceTypeToLobbyGameDisplayString = "Teams (No Shared Res)"
        Exit Function
    End Select
    allianceTypeToLobbyGameDisplayString = ""
End Function

Public Sub LobbyGameMapWidget_updateValue(ByVal self As WIDGET, ByVal mapName As String, ByVal has_alliancesMode As Boolean, ByVal alliancesMode As Long, ByVal blindMode As Long)
    Dim d As LobbyGameMapWidget, modeDetailsStr As String
    Set d = LobbyGameMapWidget_(self)
    d.blindMode_ = blindMode

    If isBlindSimpleLobby(blindMode) Then
        ' Display "Matchmaking Room" label instead of the map
        W_LABEL_setFont d.mapNameWidg, font_regular_bold
        d.mapNameWidg.setString "Matchmaking"
    Else
        W_LABEL_setFont d.mapNameWidg, font_regular
        d.mapNameWidg.setString mapName
    End If

    modeDetailsStr = allianceTypeToLobbyGameDisplayString(has_alliancesMode, alliancesMode)
    If blindMode <> BLIND_MODE_NONE And blindMode <> BLIND_MODE_BLIND_LOBBY_SIMPLE_LOBBY Then
        If Len(modeDetailsStr) <> 0 Then
            modeDetailsStr = modeDetailsStr & ", "
        End If
        Select Case blindMode
        Case BLIND_MODE_BLIND_LOBBY
            modeDetailsStr = modeDetailsStr & "Blind Lobby"
        Case BLIND_MODE_BLIND_GAME, BLIND_MODE_BLIND_GAME_SIMPLE_LOBBY
            modeDetailsStr = modeDetailsStr & "Blind Game"
        Case Else
        End Select
    End If
    d.modeDetailsWidg.setString modeDetailsStr
End Sub

' LobbyGameModsWidget() { }
Public Sub LobbyGameModsWidget_construct(ByVal self As WIDGET)
    WIDGET_construct self
    self.kind = K_LobbyGameModsWidget ' (VB) virtual calls from here on reach LobbyGameModsWidget's overrides
End Sub

' static std::shared_ptr<LobbyGameModsWidget> make(const std::vector<netlobby::ModDetails>& modsList, bool noVtol, bool noTanks, bool noCyborgs)
Public Function LobbyGameModsWidget_make(modsList() As netlobby_ModDetails, ByVal modsListCount As Long, ByVal noVtol As Boolean, ByVal noTanks As Boolean, ByVal noCyborgs As Boolean) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_LobbyGameModsWidget)
    LobbyGameModsWidget_construct widget
    LobbyGameModsWidget_initialize widget, modsList, modsListCount, noVtol, noTanks, noCyborgs
    Set LobbyGameModsWidget_make = widget
End Function

Private Sub LobbyGameModsWidget_initialize(ByVal self As WIDGET, modsList() As netlobby_ModDetails, ByVal modsListCount As Long, ByVal noVtol As Boolean, ByVal noTanks As Boolean, ByVal noCyborgs As Boolean)
    Dim d As LobbyGameModsWidget
    Set d = LobbyGameModsWidget_(self)
    Set d.modsListWidg = W_LABEL_new()
    W_LABEL_setFont d.modsListWidg, font_small, pal_RGBA(255, 50, 50, 255)
    W_LABEL_setCanTruncate d.modsListWidg, True
    d.modsListWidg.setTransparentToClicks True
    self.attach d.modsListWidg

    Set d.limitsListWidg = W_LABEL_new()
    W_LABEL_setFont d.limitsListWidg, font_small, psPalette(WZCOL_TEXT_MEDIUM)
    W_LABEL_setCanTruncate d.limitsListWidg, True
    d.limitsListWidg.setTransparentToClicks True
    self.attach d.limitsListWidg

    LobbyGameModsWidget_updateValue self, modsList, modsListCount, noVtol, noTanks, noCyborgs
    self.setGeometry 0, 0, self.idealWidth(), self.idealHeight()
End Sub

Public Sub LobbyGameModsWidget_geometryChanged(ByVal self As WIDGET)
    Dim d As LobbyGameModsWidget, w As Long, nextLineY0 As Long, limitWidgX0 As Long, widg As WIDGET, i As Long
    Set d = LobbyGameModsWidget_(self)
    w = self.width()
    d.modsListWidg.setGeometry WZLOBBY_WIDG_PADDING_X, WZLOBBY_WIDG_PADDING_Y + 2, w - (WZLOBBY_WIDG_PADDING_X * 2), d.modsListWidg.idealHeight()

    nextLineY0 = d.modsListWidg.y() + d.modsListWidg.height() + WZLOBBY_WIDG_LINE_SPACING + 1
    d.limitsListWidg.setGeometry WZLOBBY_WIDG_PADDING_X, nextLineY0, w - (WZLOBBY_WIDG_PADDING_X * 2), d.limitsListWidg.idealHeight()

    limitWidgX0 = 0
    For i = 1 To d.limitWidgets.Count
        Set widg = d.limitWidgets(i)
        widg.setGeometry limitWidgX0, nextLineY0, widg.width(), widg.height()
        limitWidgX0 = widg.x() + widg.width()
    Next
End Sub

Public Function LobbyGameModsWidget_idealWidth(ByVal self As WIDGET) As Long
    Dim d As LobbyGameModsWidget
    Set d = LobbyGameModsWidget_(self)
    LobbyGameModsWidget_idealWidth = gmax(d.modsListWidg.idealWidth(), d.limitsListWidg.idealWidth()) + (WZLOBBY_WIDG_PADDING_X * 2)
End Function

Public Function LobbyGameModsWidget_idealHeight(ByVal self As WIDGET) As Long
    Dim d As LobbyGameModsWidget, result As Long
    Set d = LobbyGameModsWidget_(self)
    result = (WZLOBBY_WIDG_PADDING_Y * 2) + 2
    result = result + d.modsListWidg.idealHeight() + (WZLOBBY_WIDG_LINE_SPACING + 1) + d.limitsListWidg.idealHeight()
    LobbyGameModsWidget_idealHeight = result
End Function

Public Sub LobbyGameModsWidget_updateValue(ByVal self As WIDGET, modsList() As netlobby_ModDetails, ByVal modsListCount As Long, ByVal noVtol As Boolean, ByVal noTanks As Boolean, ByVal noCyborgs As Boolean)
    Dim d As LobbyGameModsWidget, modsListStr As String, limitsStr As String, i As Long
    Set d = LobbyGameModsWidget_(self)
    For i = 0 To modsListCount - 1
        If Len(modsListStr) <> 0 Then
            modsListStr = modsListStr & ", "
        End If
        modsListStr = modsListStr & modsList(i).name
    Next
    d.modsListWidg.setString modsListStr

    If noTanks Then
        limitsStr = limitsStr & "No Tanks"
    End If
    If noVtol Then
        If Len(limitsStr) <> 0 Then
            limitsStr = limitsStr & " | "
        End If
        limitsStr = limitsStr & "No VTOL"
    End If
    If noCyborgs Then
        If Len(limitsStr) <> 0 Then
            limitsStr = limitsStr & " | "
        End If
        limitsStr = limitsStr & "No Cyborgs"
    End If
    d.limitsListWidg.setString limitsStr

    For i = 1 To d.limitWidgets.Count
        self.detach d.limitWidgets(i)
    Next
    Set d.limitWidgets = New Collection
End Sub

' LobbyActionWidgets() { }
Public Sub LobbyActionWidgets_construct(ByVal self As WIDGET)
    WIDGET_construct self
    self.kind = K_LobbyActionWidgets ' (VB) virtual calls from here on reach LobbyActionWidgets' overrides
End Sub

' static std::shared_ptr<LobbyActionWidgets> make(const OnSpectateHandlerFunc& onSpectateHandler) (Nothing: nullptr)
Public Function LobbyActionWidgets_make(ByVal onSpectateHandler As WzFunc) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_LobbyActionWidgets)
    LobbyActionWidgets_construct widget
    LobbyActionWidgets_initialize widget, onSpectateHandler
    Set LobbyActionWidgets_make = widget
End Function

Private Sub LobbyActionWidgets_initialize(ByVal self As WIDGET, ByVal onSpectateHandler As WzFunc)
    Dim d As LobbyActionWidgets
    Set d = LobbyActionWidgets_(self)
    Set d.watchButton = WzFrontendImageButton_make(True, IMAGE_EYE)
    d.watchButton.setTip "Spectate Game"
    WzFrontendImageButton_setBorderDrawMode d.watchButton, BorderDrawMode_Highlighted
    d.watchButton.setGeometry 0, 0, d.watchButton.idealWidth(), d.watchButton.idealHeight()
    self.attach d.watchButton
    If Not onSpectateHandler Is Nothing Then
        W_BUTTON_addOnClickHandler d.watchButton, WzFunc_new(FN_lobbyActionWidgets_watchClick, onSpectateHandler)
    Else
        d.watchButton.hide
    End If
End Sub

Public Sub LobbyActionWidgets_geometryChanged(ByVal self As WIDGET)
    Dim d As LobbyActionWidgets, w As Long, watchButtonX0 As Long, watchButtonY0 As Long
    Set d = LobbyActionWidgets_(self)
    w = self.width()

    watchButtonX0 = w - d.watchButton.width()
    watchButtonY0 = WZLOBBY_WIDG_PADDING_Y
    d.watchButton.setGeometry watchButtonX0, watchButtonY0, d.watchButton.width(), d.watchButton.height()
End Sub

Public Function LobbyActionWidgets_idealWidth(ByVal self As WIDGET) As Long
    Dim result As Long
    result = (WZLOBBY_WIDG_PADDING_X * 2)
    result = result + LobbyActionWidgets_(self).watchButton.idealWidth()
    LobbyActionWidgets_idealWidth = result
End Function

Public Function LobbyActionWidgets_idealHeight(ByVal self As WIDGET) As Long
    LobbyActionWidgets_idealHeight = LobbyActionWidgets_(self).watchButton.idealHeight() + (WZLOBBY_WIDG_PADDING_Y * 2)
End Function

' MARK: - ConnectToForm

' ConnectToForm() { }
Public Sub ConnectToForm_construct(ByVal self As WIDGET)
    Dim d As ConnectToForm
    W_BUTTON_construct self
    self.kind = K_ConnectToForm ' (VB) virtual calls from here on reach ConnectToForm's overrides
    Set d = ConnectToForm_(self)
    d.has_frontendImgID = False         ' = nullopt
    d.has_lastFrameMouseIsOverRowOrChildren = False ' = nullopt
    d.weakSerial = WeakWidget_register(self)
End Sub

' ConnectToForm::~ConnectToForm()
Public Sub ConnectToForm_destruct(ByVal self As WIDGET)
    Dim d As ConnectToForm
    Set d = ConnectToForm_(self)
    If Not d.optionsPopover Is Nothing Then
        PopoverWidget_close d.optionsPopover
    End If
    WeakWidget_unregister d.weakSerial
End Sub

' static std::shared_ptr<ConnectToForm> make(const OnSubmitFunc& onSubmitFunc)
Public Function ConnectToForm_make(ByVal onSubmitFunc As WzFunc) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_ConnectToForm)
    ConnectToForm_construct widget
    ConnectToForm_initialize widget, onSubmitFunc
    Set ConnectToForm_make = widget
End Function

' static std::shared_ptr<WzClickableOptionsChoiceWidget> addClickableJoinMenuItem(const std::shared_ptr<PopoverMenuWidget>& menu,
' const WzString& text, const std::function<void (WIDGET_KEY)>& onClick, bool disabled = false, iV_fonts fontId = font_regular)
Private Function addClickableJoinMenuItem(ByVal menu As WIDGET, ByVal text As String, ByVal onClick As WzFunc, _
        Optional ByVal disabled As Boolean = False, Optional ByVal fontId As iV_fonts = font_regular) As WIDGET
    Dim result As WIDGET
    Set result = WzClickableOptionsChoiceWidget_make(fontId)
    result.setString text
    WzOptionsChoiceWidget_setTextAlignment result, WLAB_ALIGNRIGHT
    If Not onClick Is Nothing Then
        WzClickableOptionsChoiceWidget_addOnClickHandler result, WzFunc_new(FN_addClickableJoinMenuItem_click, onClick)
    End If
    WzOptionsChoiceWidget_setDisabled result, disabled
    result.setGeometry 0, 0, result.idealWidth(), result.idealHeight()
    PopoverMenuWidget_addMenuItem menu, result, True
    Set addClickableJoinMenuItem = result
End Function

' addSpacerWidget = [&popoverMenu](int height)
Private Sub addSpacerWidget(ByVal popoverMenu As WIDGET, ByVal height As Long)
    Dim spacerWidget As WIDGET
    Set spacerWidget = WIDGET_new()
    spacerWidget.setGeometry 0, 0, 10, height
    PopoverMenuWidget_addMenuItem popoverMenu, spacerWidget
End Sub

Private Function ConnectToForm_buildJoinOptionsMenu(ByVal self As WIDGET) As WIDGET
    Dim d As ConnectToForm, weakSelf As Long, popoverMenu As WIDGET, lastIPServerName As String
    Dim inputItemsDisabled As Boolean, displayStandardInputJoinItems As Boolean, idealMenuHeight As Long, menuHeight As Long
    Set d = ConnectToForm_(self)
    weakSelf = d.weakSerial

    Set popoverMenu = PopoverMenuWidget_make()

    lastIPServerName = war_getLastIpServerConnect()
    inputItemsDisabled = (Len(W_EDITBOX_getString(d.inputBox)) = 0)
    displayStandardInputJoinItems = ((Not inputItemsDisabled) Or Len(lastIPServerName) = 0)

    If displayStandardInputJoinItems Then
        WzOptionsChoiceWidget_setDisabled addClickableJoinMenuItem(popoverMenu, "Join as Player", WzFunc_new(FN_connectToForm_joinAsPlayer, weakSelf)), inputItemsDisabled
        WzOptionsChoiceWidget_setDisabled addClickableJoinMenuItem(popoverMenu, "Join as Spectator", WzFunc_new(FN_connectToForm_joinAsSpectator, weakSelf)), inputItemsDisabled
    End If

    If Len(lastIPServerName) <> 0 Then
        If displayStandardInputJoinItems Then
            addSpacerWidget popoverMenu, 10
        End If

        ' [Last IP: %s]
        addClickableJoinMenuItem popoverMenu, "Last IP: " & lastIPServerName, Nothing, True

        addClickableJoinMenuItem popoverMenu, "Join as Player", WzFunc_new(FN_connectToForm_lastIPJoinAsPlayer, weakSelf, lastIPServerName)

        addClickableJoinMenuItem popoverMenu, "Join as Spectator", WzFunc_new(FN_connectToForm_lastIPJoinAsSpectator, weakSelf, lastIPServerName)

        addSpacerWidget popoverMenu, 2

        addClickableJoinMenuItem popoverMenu, "Clear Last IP", WzFunc_new(FN_connectToForm_clearLastIP), False, font_small
    End If

    idealMenuHeight = popoverMenu.idealHeight()
    menuHeight = idealMenuHeight
    If menuHeight > screenHeight Then
        menuHeight = screenHeight
    End If
    popoverMenu.setGeometry popoverMenu.x(), popoverMenu.y(), popoverMenu.idealWidth(), menuHeight

    Set ConnectToForm_buildJoinOptionsMenu = popoverMenu
End Function

Private Sub ConnectToForm_displayJoinOptionsMenu(ByVal self As WIDGET, ByVal psParent As WIDGET)
    Dim d As ConnectToForm, popoverMenu As WIDGET, weakSelf As Long
    Set d = ConnectToForm_(self)
    Set popoverMenu = ConnectToForm_buildJoinOptionsMenu(self)

    weakSelf = d.weakSerial
    self.setState WBUT_CLICKLOCK

    If Not d.optionsPopover Is Nothing Then
        PopoverWidget_close d.optionsPopover
    End If
    Set d.optionsPopover = PopoverMenuWidget_openMenu(popoverMenu, psParent, PopoverWidget_Alignment_RightOfParent, 0, 1, WzFunc_new(FN_connectToForm_menuClosed, weakSelf))
End Sub

' static WzString removeLobbyGameIdPrefixPattern(const WzString& input)
Private Function removeLobbyGameIdPrefixPattern(ByVal input_ As String) As String
    Dim stdStrInput As String, colonPos As Long, afterColon As Long, firstNonSpace As Long
    stdStrInput = input_

    ' Find the first colon
    colonPos = InStr(stdStrInput, ":") - 1     ' (-1: std::string::npos)

    ' Must be at least one character before the colon
    If colonPos = -1 Or colonPos = 0 Then
        removeLobbyGameIdPrefixPattern = input_
        Exit Function
    End If

    ' Check for 1 or more characters of whitespace after the colon
    afterColon = colonPos + 1
    firstNonSpace = afterColon

    Do While firstNonSpace < Len(stdStrInput)
        If Not isSpaceChar(Mid$(stdStrInput, firstNonSpace + 1, 1)) Then Exit Do
        firstNonSpace = firstNonSpace + 1
    Loop

    ' Must have at least one whitespace character after the colon
    ' (This is to exclude detection of ip:port combinations)
    If firstNonSpace = afterColon Then
        removeLobbyGameIdPrefixPattern = input_
        Exit Function
    End If

    ' Only return the remaining substring if it is non-empty
    If firstNonSpace < Len(stdStrInput) Then
        removeLobbyGameIdPrefixPattern = Mid$(stdStrInput, firstNonSpace + 1)
        Exit Function
    End If

    ' Otherwise, return original
    removeLobbyGameIdPrefixPattern = input_
End Function

Private Sub ConnectToForm_initialize(ByVal self As WIDGET, ByVal onSubmitFunc_ As WzFunc)
    Dim d As ConnectToForm, weakSelf As Long
    Set d = ConnectToForm_(self)
    Set d.onSubmitFunc = onSubmitFunc_

    If Not isFrontImagesNull() Then
        d.has_frontendImgID = True
        d.frontendImgID = IMAGE_CARET_DOWN_FILL
    End If

    d.inputBoxBackground = psPalette(WZCOL_TRANSPARENT_BOX)
    d.inputBoxBackground = pielightWithAlpha(d.inputBoxBackground, pielightAlpha(d.inputBoxBackground) \ 3)

    d.inputBoxBackgroundHighlight = psPalette(WZCOL_TRANSPARENT_BOX)
    d.inputBoxBackgroundHighlight = pielightWithAlpha(d.inputBoxBackgroundHighlight, pielightAlpha(d.inputBoxBackgroundHighlight) \ 2)

    weakSelf = d.weakSerial

    Set d.inputBox = W_EDITBOX_new()
    self.attach d.inputBox
    d.inputBox.setGeometry horizontalPadding, 0, 280, minInputBoxHeight
    Set W_EDITBOX_(d.inputBox).pBoxDisplay = WzFunc_new(FN_displayConnectToEditBox)
    W_EDITBOX_setMaxStringSize d.inputBox, maxConnectToInputLength
    W_EDITBOX_setPlaceholder d.inputBox, "GameID or IP:port"
    W_EDITBOX_setPlaceholderTextColor d.inputBox, True, psPalette(WZCOL_TEXT_MEDIUM)

    W_EDITBOX_setOnPasteTransformFunc d.inputBox, WzFunc_new(FN_connectToForm_pasteTransform)

    W_EDITBOX_setOnReturnHandler d.inputBox, WzFunc_new(FN_connectToForm_onReturn, weakSelf)

    d.joinTextStr = "Join"
    d.wzText.setText d.joinTextStr, font_regular

    W_BUTTON_addOnClickHandler self, WzFunc_new(FN_connectToForm_click, weakSelf)
End Sub

Public Sub ConnectToForm_geometryChanged(ByVal self As WIDGET)
    ConnectToForm_recalcLayout self
End Sub

Public Sub ConnectToForm_clicked(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    Dim clickLocationX As Long, weakSelf As Long
    clickLocationX = psContext.mx - self.x()

    ' If click is over button area, forward to W_BUTTON impl
    If clickLocationX >= (self.width() - ConnectToForm_buttonAreaWidth(self)) Then
        W_BUTTON_clicked self, psContext, key
    ElseIf clickLocationX <= inputBoxHorizontalPadding Then
        ' padding area for left of input box - simulate click
        weakSelf = ConnectToForm_(self).weakSerial
        widgScheduleTask WzFunc_new(FN_connectToForm_simulateClickTask, weakSelf)
    End If
End Sub

Public Function ConnectToForm_findMouseTargetRecursive(ByVal self As WIDGET, psContext As W_CONTEXT, ByVal key As WIDGET_KEY, ByVal wasPressed As Boolean) As WIDGET
    Dim d As ConnectToForm, mouseOverWidget As WIDGET
    Set d = ConnectToForm_(self)
    Set mouseOverWidget = WIDGET_findMouseTargetRecursive(self, psContext, key, wasPressed)
    If Not mouseOverWidget Is Nothing Then
        ' if findMouseTargetRecursive was called for this row, it means the mouse is over it
        ' (see WIDGET::findMouseTargetRecursive)
        d.has_lastFrameMouseIsOverRowOrChildren = True
        d.lastFrameMouseIsOverRowOrChildren = frameGetFrameNumber()
    Else
        d.has_lastFrameMouseIsOverRowOrChildren = False
    End If
    Set ConnectToForm_findMouseTargetRecursive = mouseOverWidget
End Function

Private Function ConnectToForm_isMouseOverFormOrChildren(ByVal self As WIDGET) As Boolean
    Dim d As ConnectToForm
    Set d = ConnectToForm_(self)
    If Not d.has_lastFrameMouseIsOverRowOrChildren Then
        ConnectToForm_isMouseOverFormOrChildren = False
        Exit Function
    End If
    ConnectToForm_isMouseOverFormOrChildren = (d.lastFrameMouseIsOverRowOrChildren = frameGetFrameNumber())
End Function

Public Sub ConnectToForm_highlight(ByVal self As WIDGET, psContext As W_CONTEXT)
    Dim mouseXLoc As Long
    mouseXLoc = psContext.mx - self.x()
    If mouseXLoc <= inputBoxHorizontalPadding Then
        ' silently ignoring button highlight when over this padding area that "belongs" to the inputBox
        Exit Sub
    End If
    W_BUTTON_highlight self, psContext
End Sub

Public Sub ConnectToForm_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As ConnectToForm, x0 As Long, y0 As Long, w As Long, state As Long
    Dim isDown As Boolean, isDisabled As Boolean, isMouseOverButtonArea As Boolean, isInputEditing As Boolean
    Dim isHighlight As Boolean, isMouseOverInputBoxArea As Boolean, buttonAreaW As Long
    Dim boxX0 As Long, boxY0 As Long, boxX1 As Long, boxY1 As Long, fillColor As Long, borderColor As Long, lineColor As Long
    Dim textColor As Long, splitButtonImageAreaX0 As Long, imgPosX0 As Long, imgPosY0 As Long, imgColor As Long
    Dim textX0 As Long, textY0 As Long, maxTextDisplayableWidth As Long, maxDisplayableMainTextWidth As Long, isTruncated As Boolean
    Set d = ConnectToForm_(self)
    x0 = xOffset + self.x()
    y0 = yOffset + self.y()
    w = self.width()
    state = self.getState()

    isDown = (state And (WBUT_DOWN Or WBUT_LOCK Or WBUT_CLICKLOCK)) <> 0
    isDisabled = (state And WBUT_DISABLE) <> 0
    isMouseOverButtonArea = ((state And WBUT_HIGHLIGHT) <> 0)
    isInputEditing = W_EDITBOX_isEditing(d.inputBox)
    isHighlight = ConnectToForm_isMouseOverFormOrChildren(self) Or isInputEditing
    isMouseOverInputBoxArea = ConnectToForm_isMouseOverFormOrChildren(self) And Not isMouseOverButtonArea

    buttonAreaW = ConnectToForm_buttonAreaWidth(self)

    ' Draw background for input box area
    boxX0 = x0
    boxY0 = y0
    boxX1 = boxX0 + (self.width() - buttonAreaW)
    boxY1 = boxY0 + self.height()
    If isInputEditing Then
        fillColor = psPalette(WZCOL_TRANSPARENT_BOX)
    ElseIf isHighlight Then
        fillColor = d.inputBoxBackgroundHighlight
    Else
        fillColor = d.inputBoxBackground
    End If
    pie_UniTransBoxFill boxX0, boxY0, boxX1, boxY1, fillColor

    ' Draw background for button
    boxX0 = x0 + (w - buttonAreaW)
    boxY0 = y0
    boxX1 = boxX0 + buttonAreaW
    boxY1 = boxY0 + self.height()
    pie_UniTransBoxFill boxX0, boxY0, boxX1, boxY1, psPalette(WZCOL_TRANSPARENT_BOX)

    ' Draw entire outer border
    borderColor = psPalette(WZCOL_TEXT_MEDIUM)
    If isInputEditing Or isMouseOverInputBoxArea Then
        borderColor = psPalette(WZCOL_TEXT_BRIGHT)
    End If

    If isDisabled Then
        borderColor = pielightWithAlpha(borderColor, pielightAlpha(borderColor) \ 2)
    End If
    boxX0 = x0
    boxY0 = y0
    boxX1 = boxX0 + self.width()
    boxY1 = boxY0 + self.height()
    iV_Box boxX0 + 1, boxY0 + 1, boxX1, boxY1, borderColor

    ' Draw Button area outer border
    If Not isInputEditing And Not isMouseOverInputBoxArea Then
        If isDown Then
            borderColor = psPalette(WZCOL_TEXT_DARK)
        ElseIf isMouseOverButtonArea Then
            borderColor = psPalette(WZCOL_TEXT_BRIGHT)
        Else
            borderColor = psPalette(WZCOL_TEXT_MEDIUM)
        End If
        If isDisabled Then
            borderColor = pielightWithAlpha(borderColor, pielightAlpha(borderColor) \ 2)
        End If
        boxX0 = x0 + (w - buttonAreaW)
        boxY0 = y0
        boxX1 = boxX0 + buttonAreaW
        boxY1 = boxY0 + self.height()
        iV_Box boxX0 + 1, boxY0 + 1, boxX1, boxY1, borderColor
    Else
        ' outer border handled by the "entire" case - just draw the dividing line
        boxX0 = x0 + (w - buttonAreaW)
        boxY0 = y0
        boxY1 = boxY0 + self.height()
        If isMouseOverButtonArea Then lineColor = psPalette(WZCOL_TEXT_BRIGHT) Else lineColor = psPalette(WZCOL_TEXT_MEDIUM)
        iV_Line boxX0 + 1, boxY0, boxX0 + 1, boxY1, lineColor
    End If

    If isMouseOverButtonArea Then textColor = psPalette(WZCOL_TEXT_BRIGHT) Else textColor = psPalette(WZCOL_TEXT_MEDIUM)
    If isDisabled Then
        textColor = pielightWithAlpha(textColor, pielightAlpha(textColor) \ 2)
    End If

    ' Draw image portion of button
    splitButtonImageAreaX0 = x0 + (w - horizontalPadding - imageDimensions - (horizontalPaddingBetweenButtonSplit \ 2))
    ' Display the image centered in the image split area
    imgPosX0 = splitButtonImageAreaX0 + (horizontalPaddingBetweenButtonSplit \ 2)
    imgPosY0 = y0 + (self.height() - imageDimensions) \ 2

    If d.has_frontendImgID Then
        imgColor = textColor
        iV_DrawImageFileAnisotropicTint FrontImages, d.frontendImgID, imgPosX0, imgPosY0, imageDimensions, imageDimensions, imgColor
    End If

    ' Draw text portion of button
    ' Draw the main text
    textX0 = gmax(x0 + horizontalPadding, splitButtonImageAreaX0 - d.wzText.width())
    textY0 = ToInt(y0 + (self.height() - d.wzText.lineSize()) \ 2 - CSng(d.wzText.aboveBase()))

    maxTextDisplayableWidth = splitButtonImageAreaX0 - textX0
    maxDisplayableMainTextWidth = maxTextDisplayableWidth
    isTruncated = maxDisplayableMainTextWidth < d.wzText.width()
    If isTruncated Then
        maxDisplayableMainTextWidth = maxDisplayableMainTextWidth - (iV_GetEllipsisWidth(d.wzText.getFontID()) + 2)
    End If
    d.wzText.render textX0, textY0, textColor, 0!, maxDisplayableMainTextWidth
    If isTruncated Then
        ' Render ellipsis
        iV_DrawEllipsis d.wzText.getFontID(), textX0 + maxDisplayableMainTextWidth + 2, textY0, textColor
    End If
End Sub

Public Function ConnectToForm_idealWidth(ByVal self As WIDGET) As Long
    ConnectToForm_idealWidth = 360
End Function

Public Function ConnectToForm_idealHeight(ByVal self As WIDGET) As Long
    ConnectToForm_idealHeight = (verticalPadding * 2) + gmax(iV_GetTextLineSize(font_regular), minInputBoxHeight)
End Function

Private Function ConnectToForm_buttonAreaWidth(ByVal self As WIDGET) As Long
    ConnectToForm_buttonAreaWidth = horizontalPadding + gmin(ConnectToForm_(self).wzText.width(), maxButtonTextAreaWidth) + horizontalPaddingBetweenButtonSplit + imageDimensions + horizontalPadding
End Function

Private Sub ConnectToForm_recalcLayout(ByVal self As WIDGET)
    Dim w As Long, h As Long, inputBoxWidth As Long
    w = self.width()
    h = self.height()

    ' input box takes up the leftmost portion, minus the space reserved for the button area
    inputBoxWidth = w - ConnectToForm_buttonAreaWidth(self) - (inputBoxHorizontalPadding * 2)
    If inputBoxWidth > 0 Then
        ConnectToForm_(self).inputBox.setGeometry inputBoxHorizontalPadding, 0, inputBoxWidth, h
    End If
End Sub

Private Sub ConnectToForm_submit(ByVal self As WIDGET, ByVal inputStr As String, ByVal asSpectator As Boolean)
    Dim d As ConnectToForm
    Set d = ConnectToForm_(self)
    If Not (Not d.onSubmitFunc Is Nothing) Then ASSERT_FAILURE "onSubmitFunc != nullptr", "src/titleui/widgets/gamebrowserform.cpp:1266", "submit", "Missing onSubmitFunc": Exit Sub
    WzFunc_call d.onSubmitFunc, inputStr, asSpectator
End Sub

Private Sub ConnectToForm_submitInputBox(ByVal self As WIDGET, ByVal asSpectator As Boolean)
    Dim d As ConnectToForm
    Set d = ConnectToForm_(self)
    ConnectToForm_submit self, W_EDITBOX_getString(d.inputBox), asSpectator
    W_EDITBOX_setString d.inputBox, ""
End Sub

' MARK: - LobbyBrowser Form

' LobbyBrowser() { }
Public Sub LobbyBrowser_construct(ByVal self As WIDGET)
    Dim d As LobbyBrowser
    WIDGET_construct self
    self.kind = K_LobbyBrowser ' (VB) virtual calls from here on reach LobbyBrowser's overrides
    Set d = LobbyBrowser_(self)
    d.minDesiredTableWidth = 0
    d.filterIncompatible = True
    d.filterEmpty = False
    d.filterModded = False
    d.filterIPv6Only = False
    d.hasPendingFetchRequest = False
    d.weakSerial = WeakWidget_register(self)
End Sub

' LobbyBrowser::~LobbyBrowser()
Public Sub LobbyBrowser_destruct(ByVal self As WIDGET)
    Dim d As LobbyBrowser
    Set d = LobbyBrowser_(self)
    If Not d.currentPopoverMenu Is Nothing Then
        PopoverMenuWidget_closeMenu d.currentPopoverMenu
    End If
    WeakWidget_unregister d.weakSerial
End Sub

' static std::shared_ptr<LobbyBrowser> make(const std::function<void()>& onBackButtonFunc)
Public Function LobbyBrowser_make(ByVal onBackButtonFunc As WzFunc) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_LobbyBrowser)
    LobbyBrowser_construct widget
    LobbyBrowser_initialize widget, onBackButtonFunc
    Set LobbyBrowser_make = widget
End Function

Public Sub LobbyBrowser_triggerRefresh(ByVal self As WIDGET)
    LobbyBrowser_triggerAsyncGameListFetch self
End Sub

Private Sub LobbyBrowser_triggerAsyncGameListFetch(ByVal self As WIDGET)
    Dim d As LobbyBrowser, weakSelf As Long, dispatchedListRequest As Boolean
    Set d = LobbyBrowser_(self)
    If d.hasPendingFetchRequest Then Exit Sub

    If NET_getLobbyDisabled() Then
        d.lobbyStatusOverlayWidg.setString "There appears to be a game update available!"
        LobbyBrowser_recalcLobbyStatusOverlayLayout self
        d.lobbyStatusOverlayWidg.show
        Exit Sub
    End If

    d.refreshButton.setState WBUT_DISABLE
    LobbyStatusOverlayWidget_showIndeterminateIndicator d.lobbyStatusOverlayWidg, True
    LobbyStatusOverlayWidget_showResetLobbyAddressButton d.lobbyStatusOverlayWidg, False
    d.lobbyStatusOverlayWidg.setString ""
    LobbyBrowser_recalcLobbyStatusOverlayLayout self
    d.lobbyStatusOverlayWidg.show True

    weakSelf = d.weakSerial
    dispatchedListRequest = netlobby_EnumerateGames(NETgetLobbyserverAddress(), WzFunc_new(FN_lobbyBrowser_listResult, weakSelf), True, 200)

    If Not dispatchedListRequest Then
        ' Restore widgets config
        d.refreshButton.setState 0
        LobbyStatusOverlayWidget_showIndeterminateIndicator d.lobbyStatusOverlayWidg, False

        debug_ 1387, LOG_INFO, "triggerAsyncGameListFetch", "Failed to dispatch lobby list request"

        ' Display error if no results
        If d.currentResultsCount = 0 Then
            d.lobbyStatusOverlayWidg.setString "Error dispatching request"
            d.lobbyStatusOverlayWidg.show
        End If
        LobbyBrowser_recalcLobbyStatusOverlayLayout self
        Exit Sub
    End If

    d.hasPendingFetchRequest = True
End Sub

Private Sub LobbyBrowser_joinLobbyGame(ByVal self As WIDGET, ByVal idx As Long, ByVal asSpectator As Boolean)
    Dim d As LobbyBrowser, listing As netlobby_GameListing, expectedHostProps As ExpectedHostProperties
    Set d = LobbyBrowser_(self)
    If Not (idx < d.currentResultsCount) Then ASSERT_FAILURE "idx < currentResults.size()", "src/titleui/widgets/gamebrowserform.cpp:1404", "joinLobbyGame", "Invalid idx: %zu", idx: Exit Sub

    clearActiveConsole

    ingame.localOptionsReceived = False                             ' note, we are awaiting options
    listing = d.currentResult(idx)
    game.name_ = sstrcpy_cut(listing.details.name, 128)             ' store name

    expectedHostProps.has_hostPublicKey = True
    base64Decode listing.details.host.publicIdentity, expectedHostProps.hostPublicKey

    startLobbyJoiningAttempt_props sPlayer, NETgetLobbyserverAddress(), listing.gameId, asSpectator, expectedHostProps, _
        True, listing.availableConnectionTypes, listing.availableConnectionTypesCount
End Sub

Private Sub LobbyBrowser_setMotd(ByVal self As WIDGET, ByVal motd As String)
    Dim d As LobbyBrowser, lobbyMOTDParagraph As WIDGET
    Set d = LobbyBrowser_(self)
    Set lobbyMOTDParagraph = Paragraph_new()
    lobbyMOTDParagraph.setGeometry 0, 0, ScrollableListWidget_calculateListViewWidth(d.lobbyStatusMessageContainer), ScrollableListWidget_calculateListViewHeight(d.lobbyStatusMessageContainer)
    Paragraph_setFont lobbyMOTDParagraph, font_regular
    Paragraph_setFontColour lobbyMOTDParagraph, psPalette(WZCOL_TEXT_MEDIUM)
    lobbyMOTDParagraph.setTransparentToMouse True
    Paragraph_addText lobbyMOTDParagraph, motd
    ScrollableListWidget_addItem d.lobbyStatusMessageContainer, lobbyMOTDParagraph
End Sub

Private Sub LobbyBrowser_processAsyncGameListFetchResults(ByVal self As WIDGET, results As netlobby_ListResult)
    Dim d As LobbyBrowser
    Set d = LobbyBrowser_(self)
    d.hasPendingFetchRequest = False

    d.refreshButton.setState 0
    LobbyStatusOverlayWidget_showIndeterminateIndicator d.lobbyStatusOverlayWidg, False

    If Not results.has_value Then
        If results.error_.errCode = "RATE_LIMITED" And d.currentResultsCount <> 0 Then
            ' Have prior results, hit temporary rate limit - silently ignore
            Exit Sub
        End If

        d.lobbyStatusOverlayWidg.setString "Failed to connect to lobby server"
        If fnv1a_hash(netlobby_GetDefaultLobbyAddress()) = LOB_DEFAULT_BEHAVIOR_CHECKSUM Then
            If NETgetLobbyserverAddress() <> netlobby_GetDefaultLobbyAddress() Then
                LobbyStatusOverlayWidget_showResetLobbyAddressButton d.lobbyStatusOverlayWidg, True
            End If
        End If
        LobbyBrowser_recalcLobbyStatusOverlayLayout self

        ScrollableTableWidget_clearRows d.table
        ScrollableListWidget_clear d.lobbyStatusMessageContainer
        d.clearCurrentResults
        Exit Sub
    End If

    ScrollableTableWidget_clearRows d.table
    ScrollableListWidget_clear d.lobbyStatusMessageContainer

    d.setCurrentResults results.value.games, results.value.gamesCount
    LobbyBrowser_populateTableFromGameList self, True

    ' Display the lobby MOTD
    If Len(results.value.lobbyMOTD) <> 0 Then
        LobbyBrowser_setMotd self, results.value.lobbyMOTD
    End If
End Sub

' static inline optional<AllianceType> alliancesModeFromLobbyGameDetails(uint8_t alliancesMode) (has_value: the result)
Private Function alliancesModeFromLobbyGameDetails(ByVal alliancesMode As Long, out As Long) As Boolean
    If alliancesMode <= ALLIANCE_TYPE_MAX Then
        out = alliancesMode
        alliancesModeFromLobbyGameDetails = True
        Exit Function
    End If
    alliancesModeFromLobbyGameDetails = False
End Function

' static inline optional<BLIND_MODE> blindModeFromLobbyGameDetails(uint8_t blindMode) (has_value: the result)
Private Function blindModeFromLobbyGameDetails(ByVal blindMode As Long, out As Long) As Boolean
    If blindMode <= BLIND_MODE_MAX Then
        out = blindMode
        blindModeFromLobbyGameDetails = True
        Exit Function
    End If
    blindModeFromLobbyGameDetails = False
End Function

' blindModeFromLobbyGameDetails(blindMode).value_or(BLIND_MODE::NONE)
Private Function blindModeOrNone(ByVal blindMode As Long) As Long
    Dim v As Long
    If blindModeFromLobbyGameDetails(blindMode, v) Then blindModeOrNone = v Else blindModeOrNone = BLIND_MODE_NONE
End Function

' static bool isIPv6Only(const std::vector<netlobby::ConnectionType>& availableConnectionTypes)
Private Function isIPv6Only(availableConnectionTypes() As netlobby_ConnectionType, ByVal count As Long) As Boolean
    Dim i As Long, ipv As netlobby_IPVersion
    For i = 0 To count - 1
        ' connType.ipVersion.value_or(netlobby::IPVersion::IPv4) == netlobby::IPVersion::IPv4
        If availableConnectionTypes(i).has_ipVersion Then ipv = availableConnectionTypes(i).ipVersion Else ipv = netlobby_IPVersion_IPv4
        If ipv = netlobby_IPVersion_IPv4 Then
            isIPv6Only = False
            Exit Function
        End If
    Next
    isIPv6Only = True
End Function

Private Sub LobbyBrowser_populateTableFromGameList(ByVal self As WIDGET, Optional ByVal force As Boolean = False)
    Dim d As LobbyBrowser, currentMaxColumnWidths() As Long, numColumns As Long, weakSelf As Long
    Dim disableIPv6Listings As Boolean, foundGreaterVersion As Boolean, idx As Long, listing As netlobby_GameListing
    Dim isCompatibleGame As Boolean, isIPv6OnlyGame As Boolean, columnWidgets As Collection, rowHeight As Long, i As Long
    Dim row As WIDGET, isEnabledListing As Boolean, blindSimpleLobby As Boolean, tooltip As String, omittedStr As String
    Dim statusStr As String, colWidg As WIDGET, newVersionAvailable As Boolean
    Set d = LobbyBrowser_(self)
    If Not force And d.currentResultsCount = 0 Then
        ' do nothing, so we don't clear any lobby error status message (ex. when changing filters)
        Exit Sub
    End If

    ScrollableTableWidget_clearRows d.table

    numColumns = ScrollableTableWidget_getNumColumns(d.table)
    If numColumns > 0 Then ReDim currentMaxColumnWidths(0 To numColumns - 1)

    weakSelf = d.weakSerial

    disableIPv6Listings = war_getLobbyDisableIPv6()
    foundGreaterVersion = False

    ' Populate table with the results
    For idx = 0 To d.currentResultsCount - 1
        listing = d.currentResult(idx)
        ' const auto& gameInfo = listing.details;

        foundGreaterVersion = foundGreaterVersion Or NETisGreaterVersion(listing.details.netcodeVer.major, listing.details.netcodeVer.minor)

        isCompatibleGame = NETisCorrectVersion(listing.details.netcodeVer.major, listing.details.netcodeVer.minor)
        If d.filterIncompatible And Not isCompatibleGame Then
            GoTo continue_
        End If
        If d.filterEmpty And listing.details.players.current = 0 And Not isBlindSimpleLobby(blindModeOrNone(listing.details.blindMode)) Then
            GoTo continue_
        End If
        If d.filterModded And listing.details.modsCount <> 0 Then
            GoTo continue_
        End If
        isIPv6OnlyGame = isIPv6Only(listing.availableConnectionTypes, listing.availableConnectionTypesCount)
        If d.filterIPv6Only And isIPv6OnlyGame Then
            GoTo continue_
        End If

        Set columnWidgets = LobbyBrowser_createLobbyGameRowColumnWidgets(self, idx, listing)
        rowHeight = 0
        For i = 0 To columnWidgets.Count - 1
            Set colWidg = columnWidgets(i + 1)
            currentMaxColumnWidths(i) = gmax(currentMaxColumnWidths(i), gmax(colWidg.idealWidth(), 0))
            rowHeight = gmax(rowHeight, colWidg.idealHeight())
        Next

        Set row = TableRow_make(columnWidgets, rowHeight)
        TableRow_setDisabledColor row, pal_RGBA(0, 0, 0, 100)
        isEnabledListing = isCompatibleGame And (Not disableIPv6Listings Or Not isIPv6OnlyGame)
        If isEnabledListing Then
            blindSimpleLobby = isBlindSimpleLobby(blindModeOrNone(listing.details.blindMode))
            TableRow_setHighlightsOnMouseOver row, True
            TableRow_setHighlightColor row, pal_RGBA(5, 29, 245, 200)
            If blindSimpleLobby Then
                TableRow_setBackgroundColor row, True, pal_RGBA(0, 20, 130, 50)
            Else
                TableRow_setBackgroundColor row, True, pal_RGBA(0, 0, 0, 50)
            End If
        Else
            If isCompatibleGame And disableIPv6Listings And isIPv6OnlyGame Then
                TableRow_setDisabled row, True
                TableRow_setDisabledColor row, pal_RGBA(0, 0, 0, 120)
                tooltip = "This game is IPv6-only, and you have the ""Enable IPv6"" option disabled in Options > Defaults."
                row.setTip tooltip
            Else
                TableRow_setDisabled row, True
                TableRow_setDisabledColor row, pal_RGBA(0, 0, 0, 120)
                tooltip = "Your version of Warzone is incompatible with this game."
                tooltip = tooltip & vbLf
                tooltip = tooltip & "Host Version: " & listing.details.versionStr
                row.setTip tooltip
            End If
        End If
        W_BUTTON_addOnClickHandler row, WzFunc_new(FN_lobbyBrowser_rowClick, weakSelf, idx)
        ScrollableTableWidget_addRow d.table, row
continue_:
    Next

    If ScrollableTableWidget_getNumRows(d.table) = 0 Then
        If d.currentResultsCount > 50 Then
            omittedStr = "(50+ omitted)"
        Else
            omittedStr = "(" & CStr(d.currentResultsCount) & " omitted)"
        End If

        If Not d.filterEmpty And Not d.filterModded Then
            d.lobbyStatusOverlayWidg.setString "No games for your version"
        Else
            statusStr = "No games matching filters"
            statusStr = statusStr & " " & omittedStr
            d.lobbyStatusOverlayWidg.setString statusStr
        End If

        ' getVersionCheckNewVersionAvailable().value_or(false)
        If Not getVersionCheckNewVersionAvailable(newVersionAvailable) Then newVersionAvailable = False
        If NET_getLobbyDisabled() Or newVersionAvailable Or foundGreaterVersion Then
            d.lobbyStatusOverlayWidg.setString "There appears to be a game update available!"
        End If

        d.lobbyStatusOverlayWidg.show
    Else
        d.lobbyStatusOverlayWidg.hide
    End If

    LobbyBrowser_recalcLobbyStatusOverlayLayout self
End Sub

' std::vector<std::shared_ptr<WIDGET>> createLobbyGameRowColumnWidgets(size_t idx, const netlobby::GameListing& listing)
Private Function LobbyBrowser_createLobbyGameRowColumnWidgets(ByVal self As WIDGET, ByVal idx As Long, listing As netlobby_GameListing) As Collection
    Dim columnWidgets As Collection, privateGameWidg As WIDGET, playersWidg As WIDGET, frontendImgID As Long
    Dim nameWidg As WIDGET, mapWidg As WIDGET, modsWidg As WIDGET, actionsWidg As WIDGET, weakSelf As Long
    Dim onSpectateFunc As WzFunc, has_alliancesMode As Boolean, alliancesMode As Long, mods() As netlobby_ModDetails
    Set columnWidgets = New Collection
    ' const auto& gameInfo = listing.details;

    If listing.details.isPrivate Then
        Set privateGameWidg = LobbyGameIconWidget_make(True, IMAGE_LOCK)
    Else
        Set privateGameWidg = LobbyGameIconWidget_make(False, 0)
    End If
    privateGameWidg.setTransparentToClicks True
    columnWidgets.Add privateGameWidg

    If Not isBlindSimpleLobby(blindModeOrNone(listing.details.blindMode)) Then
        Set playersWidg = LobbyGamePlayersWidget_make(listing.details.players.current, listing.details.players.max)
        playersWidg.setTransparentToClicks True
        columnWidgets.Add playersWidg
    Else
        If listing.details.players.current > 0 Then frontendImgID = IMAGE_PERSON_FILL Else frontendImgID = IMAGE_PERSON
        Set playersWidg = LobbyGameIconWidget_make(True, frontendImgID)
        If listing.details.players.current > 0 Then
            playersWidg.setTip CStr(listing.details.players.current) & " players queued"
        Else
            playersWidg.setTip "Matchmaking Room: Open"
        End If
        LobbyGameIconWidget_setCustomImageColor playersWidg, True, psPalette(WZCOL_TEXT_BRIGHT)
        playersWidg.setTransparentToClicks True
        columnWidgets.Add playersWidg
    End If

    Set nameWidg = LobbyGameNameWidget_make(listing.details.name, listing.details.host.name, listing.details.host.publicIdentity)
    nameWidg.setTransparentToClicks True
    columnWidgets.Add nameWidg

    has_alliancesMode = alliancesModeFromLobbyGameDetails(listing.details.alliancesMode, alliancesMode)
    Set mapWidg = LobbyGameMapWidget_make(listing.details.map.name, has_alliancesMode, alliancesMode, blindModeOrNone(listing.details.blindMode))
    mapWidg.setTransparentToClicks True
    columnWidgets.Add mapWidg

    mods = listing.details.mods
    Set modsWidg = LobbyGameModsWidget_make(mods, listing.details.modsCount, _
        (listing.details.limits And netlobby_StructureLimits_NO_VTOLS) <> 0, _
        (listing.details.limits And netlobby_StructureLimits_NO_TANKS) <> 0, _
        (listing.details.limits And netlobby_StructureLimits_NO_BORGS) <> 0)
    modsWidg.setTransparentToClicks True
    columnWidgets.Add modsWidg

    weakSelf = LobbyBrowser_(self).weakSerial

    Set onSpectateFunc = Nothing
    If netlobby_PlayerCounts_availableSlots(listing.details.spectators) > 0 Then
        Set onSpectateFunc = WzFunc_new(FN_lobbyBrowser_spectate, weakSelf, idx)
    End If

    Set actionsWidg = LobbyActionWidgets_make(onSpectateFunc)
    columnWidgets.Add actionsWidg

    Set LobbyBrowser_createLobbyGameRowColumnWidgets = columnWidgets
End Function

Private Sub LobbyBrowser_initialize(ByVal self As WIDGET, ByVal onBackButtonFunc As WzFunc)
    Dim d As LobbyBrowser, weakSelf As Long, pad As Padding
    Set d = LobbyBrowser_(self)
    weakSelf = d.weakSerial

    d.filterIPv6Only = war_getLobbyFilterIPv6Only()

    Set d.backButton = LobbyBrowser_makeBackButton(self)
    If Not onBackButtonFunc Is Nothing Then
        W_BUTTON_addOnClickHandler d.backButton, WzFunc_new(FN_lobbyBrowser_backClick, onBackButtonFunc)
    End If
    self.attach d.backButton

    Set d.connectToForm = ConnectToForm_make(WzFunc_new(FN_lobbyBrowser_connectSubmit))
    d.connectToForm.setGeometry 0, 0, d.connectToForm.idealWidth(), d.connectToForm.idealHeight()
    self.attach d.connectToForm

    Set d.filtersButton = WzFrontendImageButton_make(True, IMAGE_GEAR)
    d.filtersButton.setString "Filters"
    WzFrontendImageButton_setPadding d.filtersButton, 7, 5
    WzFrontendImageButton_setBackgroundColor d.filtersButton, True, psPalette(WZCOL_TRANSPARENT_BOX)
    d.filtersButton.setGeometry 0, 0, d.filtersButton.idealWidth(), d.filtersButton.idealHeight()
    W_BUTTON_addOnClickHandler d.filtersButton, WzFunc_new(FN_lobbyBrowser_filtersClick, weakSelf)
    self.attach d.filtersButton

    Set d.refreshButton = WzFrontendImageButton_make(True, IMAGE_ARROW_RELOAD)
    d.refreshButton.setString "Refresh"
    WzFrontendImageButton_setPadding d.refreshButton, 6, 5
    WzFrontendImageButton_setBackgroundColor d.refreshButton, True, psPalette(WZCOL_TRANSPARENT_BOX)
    d.refreshButton.setGeometry 0, 0, d.refreshButton.idealWidth(), d.refreshButton.idealHeight()
    W_BUTTON_addOnClickHandler d.refreshButton, WzFunc_new(FN_lobbyBrowser_refreshClick, weakSelf)
    self.attach d.refreshButton

    ' Create scrolling table
    Set d.table = LobbyBrowser_createGamesList(self)
    self.attach d.table

    ' Create progress widget
    Set d.lobbyStatusOverlayWidg = LobbyStatusOverlayWidget_make(WzFunc_new(FN_lobbyBrowser_resetLobbyAddress, weakSelf))
    self.attach d.lobbyStatusOverlayWidg, ChildZPos_Front

    Set d.lobbyStatusMessageContainer = ScrollableListWidget_make()
    pad.top = 15
    pad.right_ = 15
    pad.bottom = 15
    pad.left_ = 15
    ScrollableListWidget_setPadding d.lobbyStatusMessageContainer, pad
    self.attach d.lobbyStatusMessageContainer
    d.lobbyStatusMessageContainer.setGeometry 0, 0, 100, (iV_GetTextLineSize(font_regular) * 3) + 30

    If Not NET_getLobbyDisabled() Then
        ' Start up background thread to gather lobby listing
        LobbyBrowser_triggerAsyncGameListFetch self
    Else
        d.lobbyStatusOverlayWidg.setString "There appears to be a game update available!"
        d.lobbyStatusOverlayWidg.show
        LobbyBrowser_recalcLobbyStatusOverlayLayout self

        displayLobbyDisabledNotification
    End If
End Sub

Private Sub LobbyBrowser_recalcLobbyStatusOverlayLayout(ByVal self As WIDGET)
    Dim d As LobbyBrowser, w As Long, tableHeight As Long, lobbyStatusOverlayHeight As Long, lobbyStatusOverlayY0 As Long
    Set d = LobbyBrowser_(self)
    w = self.width()

    tableHeight = d.table.height()

    lobbyStatusOverlayHeight = d.lobbyStatusOverlayWidg.idealHeight()
    lobbyStatusOverlayY0 = (tableHeight - lobbyStatusOverlayHeight) \ 2
    d.lobbyStatusOverlayWidg.setGeometry 0, lobbyStatusOverlayY0, w, lobbyStatusOverlayHeight
End Sub

Public Sub LobbyBrowser_geometryChanged(ByVal self As WIDGET)
    Dim d As LobbyBrowser, w As Long, h As Long, topButtonsY0 As Long, backButtonX0 As Long, buttonMaxHeight As Long
    Dim refreshButtonX0 As Long, filtersButtonX0 As Long, connectWidgetX0 As Long, maxAvailableConnectToFormWidget As Long
    Dim maxTopAreaY1 As Long, tablePosY0 As Long, tableHeight As Long, lobbyStatusContainerY0 As Long, widths() As Long
    Set d = LobbyBrowser_(self)
    w = self.width()
    h = self.height()

    topButtonsY0 = 10

    ' left-aligned buttons

    backButtonX0 = 10
    d.backButton.setGeometry backButtonX0, topButtonsY0, d.backButton.width(), d.backButton.height()

    ' right-aligned buttons

    buttonMaxHeight = gmax(d.refreshButton.idealHeight(), d.filtersButton.idealHeight())

    refreshButtonX0 = w - 10 - d.refreshButton.idealWidth()
    d.refreshButton.setGeometry refreshButtonX0, topButtonsY0, d.refreshButton.idealWidth(), d.refreshButton.idealHeight()

    filtersButtonX0 = d.refreshButton.x() - 5 - d.filtersButton.idealWidth()
    d.filtersButton.setGeometry filtersButtonX0, topButtonsY0, d.filtersButton.idealWidth(), d.filtersButton.idealHeight()

    ' connect-to form

    connectWidgetX0 = d.backButton.x() + d.backButton.width() + connectToFormHorizontalPadding
    maxAvailableConnectToFormWidget = d.filtersButton.x() - connectWidgetX0 - connectToFormHorizontalPadding
    d.connectToForm.setGeometry connectWidgetX0, topButtonsY0, gmin(d.connectToForm.idealWidth(), maxAvailableConnectToFormWidget), buttonMaxHeight

    ' top area height

    maxTopAreaY1 = gmax(gmax(d.backButton.y() + d.backButton.height(), d.connectToForm.y() + d.connectToForm.height()), _
        gmax(d.refreshButton.y() + d.refreshButton.height(), d.filtersButton.y() + d.filtersButton.height()))

    ' table

    tablePosY0 = maxTopAreaY1 + 10
    tableHeight = h - tablePosY0 - d.lobbyStatusMessageContainer.height()
    d.table.setGeometry 0, tablePosY0, w, tableHeight

    LobbyBrowser_recalcLobbyStatusOverlayLayout self ' after modifying table layout + height

    ' recalc column widths
    d.getMinimumColumnWidths widths
    ScrollableTableWidget_setMinimumColumnWidths d.table, widths, d.minimumColumnWidthsCount

    lobbyStatusContainerY0 = h - d.lobbyStatusMessageContainer.height()
    d.lobbyStatusMessageContainer.setGeometry 0, lobbyStatusContainerY0, w, d.lobbyStatusMessageContainer.height()
End Sub

Public Sub LobbyBrowser_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As LobbyBrowser, x0 As Long, y0 As Long, tableX0 As Long, tableY0 As Long, tableX1 As Long, tableY1 As Long
    Dim lobbyStatusMessageContainerY0 As Long, lobbyStatusMessageContainerY1 As Long
    Set d = LobbyBrowser_(self)
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset

    ' Draw background behind table
    tableX0 = x0 + d.table.x()
    tableY0 = y0 + d.table.y()
    tableX1 = tableX0 + d.table.width()
    tableY1 = tableY0 + d.table.height()
    pie_UniTransBoxFill tableX0, tableY0, tableX1, tableY1, pal_RGBA(150, 150, 150, 50)
    pie_UniTransBoxFill tableX0, tableY0, tableX1, tableY1, pal_RGBA(0, 0, 0, 70)
    pie_UniTransBoxFill tableX0, tableY0, tableX1, tableY1, psPalette(WZCOL_TRANSPARENT_BOX)

    ' Draw darker area beneath where the lobby status message container is
    lobbyStatusMessageContainerY0 = d.lobbyStatusMessageContainer.y()
    lobbyStatusMessageContainerY1 = lobbyStatusMessageContainerY0 + d.lobbyStatusMessageContainer.height()
    pie_UniTransBoxFill x0, y0 + lobbyStatusMessageContainerY0, x0 + self.width(), y0 + lobbyStatusMessageContainerY1, psPalette(WZCOL_TRANSPARENT_BOX)
    pie_UniTransBoxFill x0, y0 + lobbyStatusMessageContainerY0, x0 + self.width(), y0 + lobbyStatusMessageContainerY1, pal_RGBA(0, 0, 0, 128)
End Sub

Public Function LobbyBrowser_idealWidth(ByVal self As WIDGET) As Long
    LobbyBrowser_idealWidth = gmax(750, LobbyBrowser_(self).minDesiredTableWidth)
End Function

Public Function LobbyBrowser_idealHeight(ByVal self As WIDGET) As Long
    LobbyBrowser_idealHeight = 700
End Function

Private Function LobbyBrowser_makeBackButton(ByVal self As WIDGET) As WIDGET
    Dim psBackButton As WIDGET
    Set psBackButton = WzMultiButton_new()
    psBackButton.setGeometry 0, 0, 30, 29
    psBackButton.setTip "Return"                    ' P_("menu", "Return")
    WzMultiButton_(psBackButton).imNormal = AtlasImage_(FrontImages, IMAGE_RETURN)
    WzMultiButton_(psBackButton).imDown = AtlasImage_(FrontImages, IMAGE_RETURN_HI)
    WzMultiButton_(psBackButton).doHighlight = IMAGE_RETURN_HI
    WzMultiButton_(psBackButton).tc = MAX_PLAYERS
    WzMultiButton_(psBackButton).alpha = 255
    Set LobbyBrowser_makeBackButton = psBackButton
End Function

Private Function LobbyBrowser_createColumnHeaderLabel(ByVal self As WIDGET, ByVal text As String, ByVal align As WzTextAlignment, ByVal minLeftPadding As Long) As WIDGET
    Dim label As WIDGET
    Set label = W_LABEL_new()
    W_LABEL_setTextAlignment label, align
    W_LABEL_setFontColour label, psPalette(WZCOL_TEXT_MEDIUM)
    label.setString text
    label.setGeometry 0, 0, W_LABEL_getMaxLineWidth(label), label.idealHeight()
    W_LABEL_setCacheNeverExpires label, True
    W_LABEL_setCanTruncate label, True
    Set LobbyBrowser_createColumnHeaderLabel = Margin_wrap(0, 0, 0, minLeftPadding, label)
End Function

Private Function LobbyBrowser_createGamesList(ByVal self As WIDGET) As WIDGET
    Dim d As LobbyBrowser, columns(0 To 5) As TableColumn, maxIdealColumnHeight As Long, i As Long
    Dim widths(0 To 5) As Long, total As Long
    Set d = LobbyBrowser_(self)
    ' Create column headers for games listing
    Set columns(0).columnWidget = LobbyBrowser_createColumnHeaderLabel(self, "", WLAB_ALIGNLEFT, 0)
    columns(0).resizeBehavior = TableColumn_ResizeBehavior_FIXED_WIDTH
    Set columns(1).columnWidget = LobbyBrowser_createColumnHeaderLabel(self, "", WLAB_ALIGNCENTRE, 0)
    columns(1).resizeBehavior = TableColumn_ResizeBehavior_FIXED_WIDTH
    Set columns(2).columnWidget = LobbyBrowser_createColumnHeaderLabel(self, "Name", WLAB_ALIGNLEFT, 0)
    columns(2).resizeBehavior = TableColumn_ResizeBehavior_RESIZABLE_AUTOEXPAND
    Set columns(3).columnWidget = LobbyBrowser_createColumnHeaderLabel(self, "Details", WLAB_ALIGNLEFT, 0)
    columns(3).resizeBehavior = TableColumn_ResizeBehavior_RESIZABLE_AUTOEXPAND
    Set columns(4).columnWidget = LobbyBrowser_createColumnHeaderLabel(self, "Mods", WLAB_ALIGNLEFT, 0)
    columns(4).resizeBehavior = TableColumn_ResizeBehavior_RESIZABLE_AUTOEXPAND
    Set columns(5).columnWidget = LobbyBrowser_createColumnHeaderLabel(self, "", WLAB_ALIGNLEFT, 0)
    columns(5).resizeBehavior = TableColumn_ResizeBehavior_FIXED_WIDTH

    maxIdealColumnHeight = iV_GetTextLineSize(font_regular) + 14
    For i = 0 To 5
        maxIdealColumnHeight = gmax(maxIdealColumnHeight, columns(i).columnWidget.idealHeight())
    Next
    widths(0) = 16 + (WZLOBBY_WIDG_PADDING_X * 2)
    widths(1) = 40
    widths(2) = 150
    widths(3) = 100
    widths(4) = 60
    widths(5) = 30
    d.setMinimumColumnWidths widths, 6
    total = 0
    For i = 0 To 5
        total = total + widths(i)               ' std::accumulate
    Next
    d.minDesiredTableWidth = total
    d.minDesiredTableWidth = d.minDesiredTableWidth + (WZLOBBY_TABLE_COL_PADDING_X * 2) + (WZLOBBY_TABLE_COL_PADDING_X * 2 * d.minimumColumnWidthsCount)

    ' Create table
    Set d.table = ScrollableTableWidget_make(columns, 6, maxIdealColumnHeight)
    ScrollableTableWidget_setColumnPadding d.table, Vector2i_(WZLOBBY_TABLE_COL_PADDING_X, 0)
    ScrollableTableWidget_setItemSpacing d.table, WZLOBBY_TABLE_ROW_SPACING
    ScrollableTableWidget_setBackgroundColor d.table, pal_RGBA(0, 0, 0, 0)
    ScrollableTableWidget_setMinimumColumnWidths d.table, widths, 6
    ScrollableTableWidget_setDrawColumnLines d.table, False

    Set LobbyBrowser_createGamesList = d.table
End Function

' static std::shared_ptr<WzAdvCheckbox> addPopoverMenuOptionsCheckbox(const std::shared_ptr<PopoverMenuWidget>& popoverMenu, const WzString& text,
' bool isChecked, bool isDisabled, const std::function<void (WzAdvCheckbox& button)>& onClickFunc)
Private Function addPopoverMenuOptionsCheckbox(ByVal popoverMenu As WIDGET, ByVal text As String, ByVal isChecked As Boolean, ByVal isDisabled As Boolean, ByVal onClickFunc As WzFunc) As WIDGET
    Dim pCheckbox As WIDGET
    Set pCheckbox = WzAdvCheckbox_make(text, "")
    W_BUTTON_(pCheckbox).FontID = font_regular
    WzAdvCheckbox_setOuterVerticalPadding pCheckbox, 4
    WzAdvCheckbox_setInnerHorizontalPadding pCheckbox, 5
    WzAdvCheckbox_setIsChecked pCheckbox, isChecked
    If isDisabled Then
        pCheckbox.setState WBUT_DISABLE
    End If
    pCheckbox.setGeometry 0, 0, pCheckbox.idealWidth(), pCheckbox.idealHeight()
    If Not onClickFunc Is Nothing Then
        W_BUTTON_addOnClickHandler pCheckbox, WzFunc_new(FN_addPopoverMenuOptionsCheckbox_click, onClickFunc)
    End If
    PopoverMenuWidget_addMenuItem popoverMenu, pCheckbox, False
    Set addPopoverMenuOptionsCheckbox = pCheckbox
End Function

Private Function LobbyBrowser_createFiltersPopoverForm(ByVal self As WIDGET) As WIDGET
    Dim d As LobbyBrowser, popoverMenu As WIDGET, weakSelf As Long, idealMenuHeight As Long, menuHeight As Long
    Set d = LobbyBrowser_(self)
    Set popoverMenu = PopoverMenuWidget_make()

    ' addOptionsCheckbox = [&popoverMenu](...) { return addPopoverMenuOptionsCheckbox(popoverMenu, ...); }

    weakSelf = d.weakSerial

    addPopoverMenuOptionsCheckbox popoverMenu, "Hide Incompatible Games", d.filterIncompatible, False, WzFunc_new(FN_lobbyBrowser_filterIncompatible, weakSelf)
    addPopoverMenuOptionsCheckbox popoverMenu, "Hide Empty Games", d.filterEmpty, False, WzFunc_new(FN_lobbyBrowser_filterEmpty, weakSelf)
    addPopoverMenuOptionsCheckbox popoverMenu, "Hide Modded Games", d.filterModded, False, WzFunc_new(FN_lobbyBrowser_filterModded, weakSelf)
    addPopoverMenuOptionsCheckbox popoverMenu, "Hide IPv6-Only", d.filterIPv6Only, False, WzFunc_new(FN_lobbyBrowser_filterIPv6Only, weakSelf)

    idealMenuHeight = popoverMenu.idealHeight()
    menuHeight = idealMenuHeight
    If menuHeight > screenHeight Then
        menuHeight = screenHeight
    End If
    popoverMenu.setGeometry popoverMenu.x(), popoverMenu.y(), popoverMenu.idealWidth(), menuHeight

    Set LobbyBrowser_createFiltersPopoverForm = popoverMenu
End Function

Private Sub LobbyBrowser_displayFiltersOverlay(ByVal self As WIDGET, ByVal psParent As WIDGET)
    Dim d As LobbyBrowser, psWeakSelf As Long
    Set d = LobbyBrowser_(self)
    If Not d.currentPopoverMenu Is Nothing Then
        PopoverMenuWidget_closeMenu d.currentPopoverMenu
    End If

    d.filtersButton.setState WBUT_CLICKLOCK
    Set d.currentPopoverMenu = LobbyBrowser_createFiltersPopoverForm(self)
    psWeakSelf = d.weakSerial
    PopoverMenuWidget_openMenu d.currentPopoverMenu, psParent, PopoverWidget_Alignment_LeftOfParent, 0, 1, WzFunc_new(FN_lobbyBrowser_filtersMenuClosed, psWeakSelf)
End Sub

' MARK: - Public methods

' std::shared_ptr<WIDGET> makeLobbyBrowser(const std::function<void()>& onBackButtonFunc)
Public Function makeLobbyBrowser(ByVal onBackButtonFunc As WzFunc) As WIDGET
    Set makeLobbyBrowser = LobbyBrowser_make(onBackButtonFunc)
End Function

' bool refreshLobbyBrowser(const std::shared_ptr<WIDGET>& lobbyBrowserWidg)
Public Function refreshLobbyBrowser(ByVal lobbyBrowserWidg As WIDGET) As Boolean
    Dim lobbyBrowser As WIDGET
    If Not lobbyBrowserWidg Is Nothing Then
        If WIDGET_isKindOf(lobbyBrowserWidg, K_LobbyBrowser) Then Set lobbyBrowser = lobbyBrowserWidg
    End If
    If Not (Not lobbyBrowser Is Nothing) Then ASSERT_FAILURE "lobbyBrowser != nullptr", "src/titleui/widgets/gamebrowserform.cpp:2051", "refreshLobbyBrowser", "Not a LobbyBrowser": Exit Function
    LobbyBrowser_triggerRefresh lobbyBrowser
    refreshLobbyBrowser = True
End Function

'--------------------------------------------------------------------------
' the lambdas

Public Function gamebrowserform_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim strongSelf As WIDGET, strongParent As WIDGET, strongConnectToForm As WIDGET, d As LobbyStatusOverlayWidget
    Dim result As String, withRemovedPrefix As String, context As W_CONTEXT, but As WIDGET, button As WIDGET
    Dim checkBoxButton As WIDGET, results As netlobby_ListResult, inputStr As String, asSpectator As Boolean
    Dim outcome As JoinGameFromConnectionStrOutcome, newValue As Boolean
    Select Case f.id
    Case FN_lobbyStatusOverlay_resetClick           ' [weakSelf](W_BUTTON& but)
        Set strongSelf = WeakWidget_lock(f.cap(0))
        If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/titleui/widgets/gamebrowserform.cpp:116", "initialize", "No parent?": Exit Function
        Set d = LobbyStatusOverlayWidget_(strongSelf)
        If Not (Not d.onResetLobbyAddressClickFunc Is Nothing) Then ASSERT_FAILURE "strongSelf->onResetLobbyAddressClickFunc != nullptr", "src/titleui/widgets/gamebrowserform.cpp:117", "initialize", "Func is null?": Exit Function
        WzFunc_call d.onResetLobbyAddressClickFunc

    Case FN_connectToForm_joinAsPlayer              ' [weakSelf](WIDGET_KEY)
        Set strongSelf = WeakWidget_lock(f.cap(0))
        If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/titleui/widgets/gamebrowserform.cpp:876", "buildJoinOptionsMenu", "Null widget?": Exit Function
        ConnectToForm_submitInputBox strongSelf, False

    Case FN_connectToForm_joinAsSpectator           ' [weakSelf](WIDGET_KEY)
        Set strongSelf = WeakWidget_lock(f.cap(0))
        If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/titleui/widgets/gamebrowserform.cpp:881", "buildJoinOptionsMenu", "Null widget?": Exit Function
        ConnectToForm_submitInputBox strongSelf, True

    Case FN_connectToForm_lastIPJoinAsPlayer        ' [weakSelf, lastIPServerName](WIDGET_KEY)
        Set strongSelf = WeakWidget_lock(f.cap(0))
        If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/titleui/widgets/gamebrowserform.cpp:898", "buildJoinOptionsMenu", "Null widget?": Exit Function
        ConnectToForm_submit strongSelf, f.cap(1), False

    Case FN_connectToForm_lastIPJoinAsSpectator     ' [weakSelf, lastIPServerName](WIDGET_KEY)
        Set strongSelf = WeakWidget_lock(f.cap(0))
        If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/titleui/widgets/gamebrowserform.cpp:904", "buildJoinOptionsMenu", "Null widget?": Exit Function
        ConnectToForm_submit strongSelf, f.cap(1), True

    Case FN_connectToForm_clearLastIP               ' [](WIDGET_KEY)
        war_setLastIpServerConnect ""

    Case FN_connectToForm_menuClosed                ' [weakSelf]() close handler
        Set strongConnectToForm = WeakWidget_lock(f.cap(0))
        If Not strongConnectToForm Is Nothing Then
            strongConnectToForm.setState 0 ' clear clicklock state
        End If

    Case FN_displayConnectToEditBox                 ' (WIDGET *psWidget, UDWORD xOffset, UDWORD yOffset)
        ' do not draw any background

    Case FN_connectToForm_pasteTransform            ' [](const WzString& pastedString) -> WzString
        result = args(0)

        ' Replace any \r, \n chars with a space
        result = Replace(result, vbCr, " ")
        result = Replace(result, vbLf, " ")

        ' Trim any whitespace
        result = WzString_trimmed(result)

        ' Check if text begins with "Lobby GameId: ", the standard prefix used when outputting the GameId on host / join
        ' NOTE: The translated version of this prefix is output - someone might copy it to another who uses a different language
        ' So use a heuristic: Check for "<prefix>: <possible gameid>" (key being the ":" followed by a space), then remove that prefix if found
        ' (This should avoid interfering with ip:port or hostname:port strings, which fortunately should not have a space in them)
        withRemovedPrefix = WzString_trimmed(removeLobbyGameIdPrefixPattern(result))
        If Len(withRemovedPrefix) <> 0 Then
            result = withRemovedPrefix
        End If

        gamebrowserform_c_fn = result

    Case FN_connectToForm_onReturn                  ' [weakSelf](W_EDITBOX& widg)
        Set strongParent = WeakWidget_lock(f.cap(0))
        If Not (Not strongParent Is Nothing) Then ASSERT_FAILURE "strongParent != nullptr", "src/titleui/widgets/gamebrowserform.cpp:1042", "initialize", "No parent?": Exit Function
        ConnectToForm_submitInputBox strongParent, False

    Case FN_connectToForm_click                     ' [weakSelf](W_BUTTON&)
        Set strongConnectToForm = WeakWidget_lock(f.cap(0))
        If Not (Not strongConnectToForm Is Nothing) Then ASSERT_FAILURE "strongConnectToForm != nullptr", "src/titleui/widgets/gamebrowserform.cpp:1051", "initialize", "Null form?": Exit Function
        ConnectToForm_displayJoinOptionsMenu strongConnectToForm, strongConnectToForm

    Case FN_connectToForm_simulateClickTask         ' [weakSelf]()
        Set strongSelf = WeakWidget_lock(f.cap(0))
        If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/titleui/widgets/gamebrowserform.cpp:1076", "clicked", "Widget gone?": Exit Function
        ' W_CONTEXT context = W_CONTEXT::ZeroContext(); (a new W_CONTEXT is zeroed)
        W_EDITBOX_simulateClick ConnectToForm_(strongSelf).inputBox, context

    Case FN_addClickableJoinMenuItem_click          ' [onClick](WzClickableOptionsChoiceWidget&, WIDGET_KEY key)
        WzFunc_call f.cap(0), args(1)

    Case FN_lobbyBrowser_listResult                 ' [weakSelf](netlobby::ListResult results)
        wzAsyncExecOnMainThread WzFunc_new(FN_lobbyBrowser_listResultMainThread, f.cap(0), args(0))

    Case FN_lobbyBrowser_listResultMainThread       ' [weakSelf, results = std::move(results)]() mutable
        netlobby_ListResult_take f.cap(1), results
        Set strongSelf = WeakWidget_lock(f.cap(0))
        If strongSelf Is Nothing Then
            debug_ 1372, LOG_WZ, "triggerAsyncGameListFetch", "Widget gone"
            ' Parent widget is gone
            ' (Can happen if the user closes this form before the async task returns results)
            Exit Function
        End If
        LobbyBrowser_processAsyncGameListFetchResults strongSelf, results

    Case FN_lobbyBrowser_rowClick                   ' [weakSelf, idx](W_BUTTON&)
        Set strongSelf = WeakWidget_lock(f.cap(0))
        If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/titleui/widgets/gamebrowserform.cpp:1581", "populateTableFromGameList", "Parent no longer exists?": Exit Function
        LobbyBrowser_joinLobbyGame strongSelf, f.cap(1), False

    Case FN_lobbyBrowser_spectate                   ' [weakSelf, idx]()
        Set strongSelf = WeakWidget_lock(f.cap(0))
        If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/titleui/widgets/gamebrowserform.cpp:1689", "createLobbyGameRowColumnWidgets", "Parent no longer exists?": Exit Function
        LobbyBrowser_joinLobbyGame strongSelf, f.cap(1), True

    Case FN_lobbyActionWidgets_watchClick           ' [onSpectateHandler](W_BUTTON&)
        WzFunc_call f.cap(0)

    Case FN_lobbyBrowser_backClick                  ' [onBackButtonFunc](W_BUTTON& button)
        widgScheduleTask WzFunc_new(FN_lobbyBrowser_backTask, f.cap(0))

    Case FN_lobbyBrowser_backTask                   ' [onBackButtonFunc]()
        WzFunc_call f.cap(0)

    Case FN_lobbyBrowser_connectSubmit              ' [](WzString inputStr, bool asSpectator)
        inputStr = args(0)
        asSpectator = args(1)
        If Len(inputStr) = 0 Then
            Exit Function
        End If
        If joinGameFromConnectionStr(inputStr, outcome, asSpectator) Then
            If outcome = JoinGameFromConnectionStrOutcome_Attempting_IP_Join Then
                ' Persist the input IP address
                war_setLastIpServerConnect inputStr
            End If
        End If

    Case FN_lobbyBrowser_filtersClick               ' [weakSelf](W_BUTTON& but)
        Set but = args(0)
        Set strongSelf = WeakWidget_lock(f.cap(0))
        If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/titleui/widgets/gamebrowserform.cpp:1747", "initialize", "No parent?": Exit Function
        LobbyBrowser_displayFiltersOverlay strongSelf, but

    Case FN_lobbyBrowser_refreshClick               ' [weakSelf](W_BUTTON& but)
        Set strongSelf = WeakWidget_lock(f.cap(0))
        If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/titleui/widgets/gamebrowserform.cpp:1759", "initialize", "No parent?": Exit Function
        LobbyBrowser_triggerAsyncGameListFetch strongSelf

    Case FN_lobbyBrowser_resetLobbyAddress          ' [weakSelf]()
        ' on reset lobby address button click
        NETsetLobbyserverAddress netlobby_GetDefaultLobbyAddress()
        ' trigger lobby re-fetch
        Set strongSelf = WeakWidget_lock(f.cap(0))
        If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/titleui/widgets/gamebrowserform.cpp:1774", "initialize", "No parent?": Exit Function
        LobbyBrowser_triggerAsyncGameListFetch strongSelf

    Case FN_addPopoverMenuOptionsCheckbox_click     ' [onClickFunc](W_BUTTON& button)
        Set button = args(0)
        If WIDGET_isKindOf(button, K_WzAdvCheckbox) Then Set checkBoxButton = button
        If Not (Not checkBoxButton Is Nothing) Then ASSERT_FAILURE "checkBoxButton != nullptr", "src/titleui/widgets/gamebrowserform.cpp:1963", "addPopoverMenuOptionsCheckbox", "checkBoxButton is null": Exit Function
        WzFunc_call f.cap(0), checkBoxButton

    Case FN_lobbyBrowser_filterIncompatible         ' [weakSelf](WzAdvCheckbox& button)
        Set strongSelf = WeakWidget_lock(f.cap(0))
        If Not strongSelf Is Nothing Then
            LobbyBrowser_(strongSelf).filterIncompatible = WzAdvCheckbox_isChecked(args(0))
            LobbyBrowser_populateTableFromGameList strongSelf
        End If

    Case FN_lobbyBrowser_filterEmpty                ' [weakSelf](WzAdvCheckbox& button)
        Set strongSelf = WeakWidget_lock(f.cap(0))
        If Not strongSelf Is Nothing Then
            LobbyBrowser_(strongSelf).filterEmpty = WzAdvCheckbox_isChecked(args(0))
            LobbyBrowser_populateTableFromGameList strongSelf
        End If

    Case FN_lobbyBrowser_filterModded               ' [weakSelf](WzAdvCheckbox& button)
        Set strongSelf = WeakWidget_lock(f.cap(0))
        If Not strongSelf Is Nothing Then
            LobbyBrowser_(strongSelf).filterModded = WzAdvCheckbox_isChecked(args(0))
            LobbyBrowser_populateTableFromGameList strongSelf
        End If

    Case FN_lobbyBrowser_filterIPv6Only             ' [weakSelf](WzAdvCheckbox& button)
        Set strongSelf = WeakWidget_lock(f.cap(0))
        If Not strongSelf Is Nothing Then
            newValue = WzAdvCheckbox_isChecked(args(0))
            LobbyBrowser_(strongSelf).filterIPv6Only = newValue
            war_setLobbyFilterIPv6Only newValue
            LobbyBrowser_populateTableFromGameList strongSelf
        End If

    Case FN_lobbyBrowser_filtersMenuClosed          ' [psWeakSelf]()
        Set strongSelf = WeakWidget_lock(f.cap(0))
        If Not strongSelf Is Nothing Then
            LobbyBrowser_(strongSelf).filtersButton.setState 0
        End If
    End Select
End Function
