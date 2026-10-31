Attribute VB_Name = "chatoptions_c"
' src/hci/chatoptions.cpp, chatoptions.h, whole: the chat options form (the host's free chat permissions, the local
' mutes; the lobby's and the in-game chat's options overlay). Its widget classes' fields are the data classes named as
' the classes. (VB) The pure virtual methods of WzPlayerChatMuteButtonBase and WzChatMuteColButtonBase (updateData,
' getMuteStatus, setMuteStatus) have empty bodies that no concrete class reaches (gen_widgets needs the root's
' procedure). optional<uint32_t> playerIdx is has_playerIdx and the value; optional<PIELIGHT> muteIconColorOverride
' is has_muteIconColorOverride and the value; the std::weak_ptr<WzChatOptionsForm> parentForm and the forms'
' weakSelf are the form's widgweak_vb serial. A std::vector<std::shared_ptr<WIDGET>> (TableRow::make's columns) is a
' Collection; std::vector<TableColumn> and std::vector<size_t> are arrays and their counts (table_c's convention);
' getMaxTableColumnDataWidths' std::pair is the total returned and the widths ByRef. std::sort of the players by
' position is quickchat_c's UInt32Vector_sortByPlayerPosition (MSVC's insertion sort). Translations (_()) are not
' ported: the English text is used.
' (VB) widget class WzChatOptionsForm : W_CLICKFORM
' (VB) widget class WzChatOptionsButton : W_BUTTON
' (VB) widget class WzPlayerChatMuteButtonBase : W_BUTTON
' (VB) virtual WzPlayerChatMuteButtonBase.updateData
' (VB) virtual WzPlayerChatMuteButtonBase.getMuteStatus
' (VB) virtual WzPlayerChatMuteButtonBase.setMuteStatus
' (VB) virtual WzPlayerChatMuteButtonBase.clickPrimary
' (VB) virtual WzPlayerChatMuteButtonBase.clickSecondary
' (VB) widget class WzHostFreeChatPlayerMuteButton : WzPlayerChatMuteButtonBase
' (VB) widget class WzPlayerMuteStatusLabel : WIDGET
' (VB) widget class WzChatMuteColButtonBase : W_BUTTON
' (VB) virtual WzChatMuteColButtonBase.getMuteStatus
' (VB) virtual WzChatMuteColButtonBase.setMuteStatus
' (VB) virtual WzChatMuteColButtonBase.clickPrimary
' (VB) virtual WzChatMuteColButtonBase.clickSecondary
' (VB) widget class WzLocalChatMuteColButton : WzChatMuteColButtonBase
' (VB) widget class WzGlobalChatMuteColButton : WzChatMuteColButtonBase
Option Explicit

Private Const MAX_CONNECTED_PLAYERS As Long = MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS

' MARK: - Constants

Private Const WZCOL_CHATOPTIONS_BUTTON_FILL_COLOR As Long = &HDC6E0019         ' pal_RGBA(25, 0, 110, 220)
Private Const WZCOL_CHATOPTIONS_BUTTON_FILL_COLOR_DARK As Long = &HFA46000A    ' pal_RGBA(10, 0, 70, 250)
Private Const WZCOL_CHATOPTIONS_BUTTON_BORDER_LIGHT As Long = &HFFFF6464       ' pal_RGBA(100, 100, 255, 255)
Private Const WZCOL_CHATOPTIONS_BUTTON_BORDER_HIGHLIGHT As Long = &HE1FFFFFF   ' pal_RGBA(255, 255, 255, 225)

Private Const WZCOL_CHATOPTIONS_HOST_GLOBAL_SET_COLOR As Long = &HFFFF6464     ' pal_RGBA(100, 100, 255, 255)

Private Const ChatOptionsButtonHorizontalPadding As Long = 5
Private Const CHATOPT_FORM_EXTERNAL_PADDING As Long = 20
Private Const CHATOPT_FORM_INTERNAL_PADDING As Long = 10
Private Const CHATOPT_PLAYER_BUTTON_EXTERNAL_PADDING_X As Long = 5
Private Const CHATOPT_PLAYER_BUTTON_INTERNAL_PADDING_X As Long = 6
Private Const CHATOPT_PLAYER_BUTTON_HEIGHT As Long = 22

Private Const CHATOPT_COLUMN_HEADER_PADDING_X As Long = 5
Private Const CHATOPT_COLUMN_HEADER_PADDING_Y As Long = 5
Private Const CHATOPT_COLUMN_PADDING_X As Long = 2

Private Const playerNameColIdx As Long = 0
Private Const MAX_PLAYER_NAME_COLUMN_WIDTH As Long = 200

' the lambdas (WzFunc: chatoptions_c_fn)
Public Const FN_chatOpt_updateButton_onClick As Long = 15100   ' updateButton->addOnClickHandler([](W_BUTTON& button) {...})
Public Const FN_chatOpt_updateButton_task As Long = 15101      ' widgScheduleTask([psParent](){...})
Public Const FN_chatOpt_enableAll_onClick As Long = 15102      ' freechatEnableAllButton->addOnClickHandler([weakSelf](W_BUTTON& widg) {...})
Public Const FN_chatOpt_disableAll_onClick As Long = 15103     ' freechatDisableAllButton->addOnClickHandler([weakSelf](W_BUTTON& widg) {...})

'--------------------------------------------------------------------------
' (VB) helpers

' a PIELIGHT's alpha (its byte.a: the top byte)
Private Function pielightAlpha(ByVal c As Long) As Long
    pielightAlpha = SHRU(c, 24)
End Function

Private Function pielightWithAlpha(ByVal c As Long, ByVal a As Long) As Long
    pielightWithAlpha = (c And &HFFFFFF) Or SHL(a And &HFF&, 24)
End Function

' MARK: - WzChatOptionsButton

Public Sub WzChatOptionsButton_construct(ByVal self As WIDGET)
    W_BUTTON_construct self
    self.kind = K_WzChatOptionsButton ' (VB) virtual calls from here on reach WzChatOptionsButton's overrides
    WzChatOptionsButton_(self).lastWidgetWidth = 0
    WzChatOptionsButton_(self).isTruncated = False
End Sub

Public Function WzChatOptionsButton_make(ByVal messageText As String, Optional ByVal disabled As Boolean = False) As WIDGET
    Dim result As WIDGET
    Set result = WIDGET_alloc(K_WzChatOptionsButton)
    WzChatOptionsButton_construct result
    W_BUTTON_(result).pText = messageText
    Set WzChatOptionsButton_make = result
End Function

Public Function WzChatOptionsButton_getButtonTextWidth(ByVal self As WIDGET) As Long
    WzChatOptionsButton_getButtonTextWidth = iV_GetTextWidth(W_BUTTON_(self).pText, W_BUTTON_(self).FontID)
End Function

' Display the text, shrunk to fit, with a faint border around it (that gets bright when highlighted)
Public Sub WzChatOptionsButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As WzChatOptionsButton, b As W_BUTTON, pParent As WIDGET, xPos As Long, yPos As Long, h As Long, w As Long
    Dim isDown As Boolean, isDisabled As Boolean, isHighlight As Boolean, light_border As Long, regular_border As Long
    Dim fill_color As Long, border As Long, msgColour As Long, fontID As iV_fonts, availableButtonTextWidth As Long
    Dim result As iV_fonts, textX0 As Long, textY0 As Long, maxTextDisplayableWidth As Long
    Set d = WzChatOptionsButton_(self)
    Set b = W_BUTTON_(self)
    Set pParent = self.parent()
    If Not (Not pParent Is Nothing) Then ASSERT_FAILURE "pParent != nullptr", "src/hci/chatoptions.cpp:128", "display", "Keymap buttons should have a parent container!": Exit Sub

    xPos = xOffset + self.x()
    yPos = yOffset + self.y()
    h = self.height()
    w = self.width()

    isDown = (self.getState() And (WBUT_DOWN Or WBUT_LOCK Or WBUT_CLICKLOCK)) <> 0
    isDisabled = (b.state And WBUT_DISABLE) <> 0
    isHighlight = (b.state And WBUT_HIGHLIGHT) <> 0

    ' Draw background
    light_border = WZCOL_CHATOPTIONS_BUTTON_BORDER_HIGHLIGHT
    If Not isDisabled Then regular_border = WZCOL_CHATOPTIONS_BUTTON_BORDER_LIGHT Else regular_border = psPalette(WZCOL_FORM_DARK)
    If isDown Or isDisabled Then fill_color = WZCOL_CHATOPTIONS_BUTTON_FILL_COLOR_DARK Else fill_color = WZCOL_CHATOPTIONS_BUTTON_FILL_COLOR
    If isDown Or isHighlight Then border = light_border Else border = regular_border
    iV_ShadowBox xPos, yPos, xPos + w, yPos + h, 0, border, border, fill_color

    ' Select label text and color
    If isDown Then msgColour = psPalette(WZCOL_TEXT_BRIGHT) Else msgColour = psPalette(WZCOL_FORM_TEXT)
    If isDisabled Then
        msgColour = psPalette(WZCOL_FORM_DISABLE)
    End If

    fontID = d.wzMessageText.getFontID()
    If fontID = font_count Or d.lastWidgetWidth <> self.width() Or d.wzMessageText.getText() <> b.pText Then
        fontID = b.FontID
    End If
    d.wzMessageText.setText b.pText, fontID

    availableButtonTextWidth = w - (ChatOptionsButtonHorizontalPadding * 2)
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

    textX0 = MAX_(xPos + ChatOptionsButtonHorizontalPadding, xPos + ((w - d.wzMessageText.width()) \ 2))
    ' (static_cast<int>(yPos + (h - lineSize()) / 2 - float(aboveBase())): integers throughout)
    textY0 = yPos + (h - d.wzMessageText.lineSize()) \ 2 - d.wzMessageText.aboveBase()

    maxTextDisplayableWidth = w - (ChatOptionsButtonHorizontalPadding * 2)
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

Public Function WzChatOptionsButton_getTip(ByVal self As WIDGET) As String
    Dim d As WzChatOptionsButton
    Set d = WzChatOptionsButton_(self)
    If Len(W_BUTTON_(self).pTip) = 0 And Not d.isTruncated Then
        WzChatOptionsButton_getTip = ""
        Exit Function
    End If

    If Not d.isTruncated Then
        WzChatOptionsButton_getTip = W_BUTTON_(self).pTip
        Exit Function
    End If

    WzChatOptionsButton_getTip = W_BUTTON_(self).pText
End Function

Public Sub WzChatOptionsButton_simulateClick(ByVal self As WIDGET)
    Dim context As W_CONTEXT    ' W_CONTEXT::ZeroContext()
    context.mx = self.screenPosX()
    context.my = self.screenPosY()

    self.clicked context, WKEY_PRIMARY
    self.released context, WKEY_PRIMARY
End Sub

' MARK: - WzPlayerChatMuteButtonBase

Public Sub WzPlayerChatMuteButtonBase_construct(ByVal self As WIDGET)
    W_BUTTON_construct self
    self.kind = K_WzPlayerChatMuteButtonBase ' (VB) virtual calls from here on reach WzPlayerChatMuteButtonBase's overrides
    WzPlayerChatMuteButtonBase_(self).has_muteIconColorOverride = False
    WzPlayerChatMuteButtonBase_(self).isSelectedPlayer = False
End Sub

' virtual void updateData() = 0;
Public Sub WzPlayerChatMuteButtonBase_updateData(ByVal self As WIDGET)
    ' (VB) pure virtual: never reached (WzHostFreeChatPlayerMuteButton overrides it)
End Sub

' virtual bool getMuteStatus() const = 0;
Public Function WzPlayerChatMuteButtonBase_getMuteStatus(ByVal self As WIDGET) As Boolean
    ' (VB) pure virtual: never reached (WzHostFreeChatPlayerMuteButton overrides it)
End Function

' virtual void setMuteStatus(bool muted) = 0;
Public Sub WzPlayerChatMuteButtonBase_setMuteStatus(ByVal self As WIDGET, ByVal muted As Boolean)
    ' (VB) pure virtual: never reached (WzHostFreeChatPlayerMuteButton overrides it)
End Sub

' special click-handling functions - subclasses can override
Public Sub WzPlayerChatMuteButtonBase_clickPrimary(ByVal self As WIDGET)
    WzPlayerChatMuteButtonBase_toggle self
End Sub

Public Function WzPlayerChatMuteButtonBase_clickSecondary(ByVal self As WIDGET, ByVal synthesizedFromHold As Boolean) As Boolean
    WzPlayerChatMuteButtonBase_clickSecondary = False
End Function

Public Sub WzPlayerChatMuteButtonBase_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As WzPlayerChatMuteButtonBase, b As W_BUTTON, x0 As Long, y0 As Long, x1 As Long, y1 As Long
    Dim isDown As Boolean, isDisabled As Boolean, isHighlight As Boolean, light_border As Long, regular_border As Long
    Dim fill_color As Long, border As Long, muteStatus As Boolean, fx As Long, fy As Long, maxTextDisplayableWidth As Long
    Dim isTruncated As Boolean, textX0 As Long, textY0 As Long, textColor As Long, ImageID As Long, imageAlpha As Long
    Dim skipVolumeIconDraw As Boolean
    Set d = WzPlayerChatMuteButtonBase_(self)
    Set b = W_BUTTON_(self)
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    x1 = x0 + self.width()
    y1 = y0 + self.height()

    isDown = Not d.isSelectedPlayer And ((self.getState() And (WBUT_DOWN Or WBUT_LOCK Or WBUT_CLICKLOCK)) <> 0)
    isDisabled = (self.getState() And WBUT_DISABLE) <> 0
    isHighlight = Not d.isSelectedPlayer And Not isDisabled And ((self.getState() And WBUT_HIGHLIGHT) <> 0)

    ' Display the button.
    light_border = WZCOL_CHATOPTIONS_BUTTON_BORDER_HIGHLIGHT
    If Not isDisabled Then regular_border = WZCOL_CHATOPTIONS_BUTTON_BORDER_LIGHT Else regular_border = psPalette(WZCOL_FORM_DARK)
    If d.isSelectedPlayer Then
        regular_border = WZCOL_CHATOPTIONS_BUTTON_FILL_COLOR
    End If
    If isDown Or isDisabled Then fill_color = WZCOL_CHATOPTIONS_BUTTON_FILL_COLOR_DARK Else fill_color = WZCOL_CHATOPTIONS_BUTTON_FILL_COLOR
    If isDown Or isHighlight Then border = light_border Else border = regular_border
    iV_ShadowBox x0, y0, x1, y1, 0, border, border, fill_color

    If Len(b.pText) = 0 Then
        Exit Sub
    End If

    muteStatus = vt_WzPlayerChatMuteButtonBase_getMuteStatus(self)

    d.wzText.setText b.pText, b.FontID

    fx = CHATOPT_PLAYER_BUTTON_EXTERNAL_PADDING_X + 16 + CHATOPT_PLAYER_BUTTON_INTERNAL_PADDING_X
    fy = (self.height() - d.wzText.lineSize()) \ 2 - d.wzText.aboveBase()
    maxTextDisplayableWidth = self.width() - fx
    isTruncated = maxTextDisplayableWidth < d.wzText.width()
    If isTruncated Then
        maxTextDisplayableWidth = maxTextDisplayableWidth - (iV_GetEllipsisWidth(d.wzText.getFontID()) + 2)
    End If
    textX0 = x0 + fx
    textY0 = y0 + fy

    If isDown Then textColor = psPalette(WZCOL_TEXT_BRIGHT) Else textColor = psPalette(WZCOL_FORM_TEXT)
    If isDisabled Then
        textColor = pielightWithAlpha(textColor, pielightAlpha(textColor) \ 2)
    ElseIf muteStatus Then
        textColor = psPalette(WZCOL_TEXT_MEDIUM)
    End If
    d.wzText.render textX0, textY0, textColor, 0!, maxTextDisplayableWidth
    If isTruncated Then
        ' Render ellipsis
        iV_DrawEllipsis d.wzText.getFontID(), textX0 + maxTextDisplayableWidth + 2, textY0, textColor
    End If

    ImageID = IMAGE_INTFAC_VOLUME_UP
    If Not isDisabled Then imageAlpha = 255 Else imageAlpha = 125
    skipVolumeIconDraw = False
    If Not muteStatus And Not isDisabled Then
        ImageID = IMAGE_INTFAC_VOLUME_UP
        If d.isSelectedPlayer Then
            skipVolumeIconDraw = True
        End If
    Else
        ImageID = IMAGE_INTFAC_VOLUME_MUTE
    End If
    If Not skipVolumeIconDraw Then
        If Not d.has_muteIconColorOverride Then
            iV_DrawImageFileAnisotropic IntImages, ImageID, x0 + CHATOPT_PLAYER_BUTTON_EXTERNAL_PADDING_X, y0 + 2, 16, 16, 0, imageAlpha
        Else
            iV_DrawImageTint IntImages, ImageID, x0 + CHATOPT_PLAYER_BUTTON_EXTERNAL_PADDING_X, y0 + 2, d.muteIconColorOverride, True, 16, 16
        End If
    End If
End Sub

Public Function WzPlayerChatMuteButtonBase_idealWidth(ByVal self As WIDGET) As Long
    WzPlayerChatMuteButtonBase_idealWidth = iV_GetTextWidth(W_BUTTON_(self).pText, W_BUTTON_(self).FontID) + (CHATOPT_PLAYER_BUTTON_EXTERNAL_PADDING_X * 2) + CHATOPT_PLAYER_BUTTON_INTERNAL_PADDING_X + 16
End Function

Public Sub WzPlayerChatMuteButtonBase_toggle(ByVal self As WIDGET)
    vt_WzPlayerChatMuteButtonBase_setMuteStatus self, Not vt_WzPlayerChatMuteButtonBase_getMuteStatus(self)
End Sub

Public Sub WzPlayerChatMuteButtonBase_updateStringFromPlayerName(ByVal self As WIDGET, ByVal playerIdx As Long)
    Dim d As WzPlayerChatMuteButtonBase
    Set d = WzPlayerChatMuteButtonBase_(self)
    ' keep player name updated
    If Not isHumanPlayer(playerIdx) Then
        self.setString ""
        d.isSelectedPlayer = False
        Exit Sub
    End If
    self.setString getPlayerName(playerIdx)
    d.isSelectedPlayer = (playerIdx = selectedPlayer)
End Sub

' void setMuteIconColorOverride(optional<PIELIGHT> _muteIconColorOverride)
Public Sub WzPlayerChatMuteButtonBase_setMuteIconColorOverride(ByVal self As WIDGET, ByVal has_muteIconColorOverride As Boolean, Optional ByVal muteIconColorOverride_ As Long = 0)
    WzPlayerChatMuteButtonBase_(self).has_muteIconColorOverride = has_muteIconColorOverride
    WzPlayerChatMuteButtonBase_(self).muteIconColorOverride = muteIconColorOverride_
End Sub

' optional<uint32_t> getPlayerIdxFromPosition(uint32_t playerPosition): has_value(), the value in result
Public Function WzPlayerChatMuteButtonBase_getPlayerIdxFromPosition(ByVal self As WIDGET, ByVal playerPosition As Long, result As Long) As Boolean
    Dim idx As Long
    For idx = 0 To NetPlay_players_size() - 1
        If NetPlay_players_position(idx) = playerPosition Then
            result = idx
            WzPlayerChatMuteButtonBase_getPlayerIdxFromPosition = True
            Exit Function
        End If
    Next
    WzPlayerChatMuteButtonBase_getPlayerIdxFromPosition = False
End Function

' override default button click handling
Public Sub WzPlayerChatMuteButtonBase_released(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    Dim clickAndReleaseOnThisButton As Boolean
    clickAndReleaseOnThisButton = ((W_BUTTON_(self).state And WBUT_DOWN) <> 0) ' relies on W_BUTTON handling to properly set WBUT_DOWN

    W_BUTTON_released self, psContext, key

    If Not clickAndReleaseOnThisButton Then
        Exit Sub ' do nothing
    End If

    If key = WKEY_PRIMARY Then
        vt_WzPlayerChatMuteButtonBase_clickPrimary self
    ElseIf key = WKEY_SECONDARY Then
        vt_WzPlayerChatMuteButtonBase_clickSecondary self, False
    End If
End Sub

' MARK: - WzHostFreeChatPlayerMuteButton

Public Sub WzHostFreeChatPlayerMuteButton_construct(ByVal self As WIDGET)
    WzPlayerChatMuteButtonBase_construct self
    self.kind = K_WzHostFreeChatPlayerMuteButton ' (VB) virtual calls from here on reach WzHostFreeChatPlayerMuteButton's overrides
    WzHostFreeChatPlayerMuteButton_(self).has_playerIdx = False
End Sub

' static std::shared_ptr<WzHostFreeChatPlayerMuteButton> make(optional<uint32_t> playerIdx,
' std::shared_ptr<WzChatOptionsForm> parentForm)
Public Function WzHostFreeChatPlayerMuteButton_make(ByVal has_playerIdx As Boolean, ByVal playerIdx As Long, ByVal parentForm As WIDGET) As WIDGET
    Dim button As WIDGET, d As WzHostFreeChatPlayerMuteButton, valueOr0 As Long
    If has_playerIdx Then valueOr0 = playerIdx Else valueOr0 = 0
    If Not ULT(valueOr0, MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "playerIdx.value_or(0) < MAX_CONNECTED_PLAYERS", "src/hci/chatoptions.cpp:398", "make", "Invalid playerPosition": Exit Function
    Set button = WIDGET_alloc(K_WzHostFreeChatPlayerMuteButton)
    WzHostFreeChatPlayerMuteButton_construct button
    Set d = WzHostFreeChatPlayerMuteButton_(button)
    d.has_playerIdx = has_playerIdx
    d.playerIdx = playerIdx
    If Not parentForm Is Nothing Then d.parentForm = WzChatOptionsForm_(parentForm).weakSerial
    W_BUTTON_(button).FontID = font_regular
    vt_WzPlayerChatMuteButtonBase_updateData button
    button.setGeometry 0, 0, button.idealWidth(), CHATOPT_PLAYER_BUTTON_HEIGHT
    Set WzHostFreeChatPlayerMuteButton_make = button
End Function

Public Sub WzHostFreeChatPlayerMuteButton_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    Dim d As WzHostFreeChatPlayerMuteButton, isDisabled As Boolean, isHuman As Boolean
    Set d = WzHostFreeChatPlayerMuteButton_(self)
    If d.has_playerIdx Then
        ' keep player name updated
        WzPlayerChatMuteButtonBase_updateStringFromPlayerName self, d.playerIdx
        isDisabled = (self.getState() And WBUT_DISABLE) <> 0
        isHuman = isHumanPlayer(d.playerIdx)
        If isHuman = isDisabled Then
            If isHuman Then self.setState 0 Else self.setState WBUT_DISABLE
        End If
    End If
End Sub

Public Sub WzHostFreeChatPlayerMuteButton_updateData(ByVal self As WIDGET)
    Dim d As WzHostFreeChatPlayerMuteButton
    Set d = WzHostFreeChatPlayerMuteButton_(self)
    If d.has_playerIdx Then
        WzPlayerChatMuteButtonBase_updateStringFromPlayerName self, d.playerIdx
    Else
        self.setString "New Joins"
    End If
End Sub

Public Function WzHostFreeChatPlayerMuteButton_getMuteStatus(ByVal self As WIDGET) As Boolean
    Dim d As WzHostFreeChatPlayerMuteButton
    Set d = WzHostFreeChatPlayerMuteButton_(self)
    If d.has_playerIdx Then
        WzHostFreeChatPlayerMuteButton_getMuteStatus = Not ingame.hostChatPermissions(d.playerIdx)
    Else
        WzHostFreeChatPlayerMuteButton_getMuteStatus = Not NETgetDefaultMPHostFreeChatPreference()
    End If
End Function

Public Sub WzHostFreeChatPlayerMuteButton_setMuteStatus(ByVal self As WIDGET, ByVal muted As Boolean)
    Dim d As WzHostFreeChatPlayerMuteButton, strongParent As WIDGET
    Set d = WzHostFreeChatPlayerMuteButton_(self)
    If Not d.has_playerIdx Then
        NETsetDefaultMPHostFreeChatPreference Not muted
        ' do not update any existing players - this just sets the default for new joins
        Exit Sub
    End If
    If Not isHumanPlayer(d.playerIdx) Then
        Exit Sub
    End If
    ingame.hostChatPermissions(d.playerIdx) = Not muted
    If NetPlay_isHost And NetPlay_bComms Then
        sendHostConfig
    End If
    Set strongParent = WeakWidget_lock(d.parentForm)
    If Not strongParent Is Nothing Then
        WzChatOptionsForm_callOnSettingsChanged strongParent
    End If
End Sub

' MARK: - WzPlayerMuteStatusLabel

Public Sub WzPlayerMuteStatusLabel_construct(ByVal self As WIDGET)
    Dim d As WzPlayerMuteStatusLabel
    WIDGET_construct self
    self.kind = K_WzPlayerMuteStatusLabel ' (VB) virtual calls from here on reach WzPlayerMuteStatusLabel's overrides
    Set d = WzPlayerMuteStatusLabel_(self)
    d.playerIdx = 0
    d.FontID = font_regular
    d.has_muteIconColorOverride = False
    d.isSelectedPlayer = False
    d.isDisabled = False
End Sub

Public Function WzPlayerMuteStatusLabel_make(ByVal playerIdx As Long) As WIDGET
    Dim label As WIDGET, d As WzPlayerMuteStatusLabel
    If Not ULT(playerIdx, MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "playerIdx < MAX_CONNECTED_PLAYERS", "src/hci/chatoptions.cpp:482", "make", "Invalid playerIdx": Exit Function
    Set label = WIDGET_alloc(K_WzPlayerMuteStatusLabel)
    WzPlayerMuteStatusLabel_construct label
    Set d = WzPlayerMuteStatusLabel_(label)
    d.playerIdx = playerIdx
    d.FontID = font_regular
    WzPlayerMuteStatusLabel_updateData label
    label.setGeometry 0, 0, label.idealWidth(), CHATOPT_PLAYER_BUTTON_HEIGHT
    Set WzPlayerMuteStatusLabel_make = label
End Function

Public Sub WzPlayerMuteStatusLabel_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    Dim d As WzPlayerMuteStatusLabel, isHuman As Boolean
    Set d = WzPlayerMuteStatusLabel_(self)
    ' keep player name updated
    WzPlayerMuteStatusLabel_updateStringFromPlayerName self

    isHuman = isHumanPlayer(d.playerIdx)
    If isHuman = d.isDisabled Then
        If Not isHuman Then
            ' no human player for this slot anymore - just disable it
            d.isDisabled = True
        Else
            ' enable it
            d.isDisabled = False
        End If
    End If
End Sub

Public Sub WzPlayerMuteStatusLabel_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As WzPlayerMuteStatusLabel, x0 As Long, y0 As Long, muteStatus As Boolean, fx As Long, fy As Long
    Dim maxTextDisplayableWidth As Long, isTruncated As Boolean, textX0 As Long, textY0 As Long, textColor As Long
    Dim ImageID As Long, imageAlpha As Long, skipVolumeIconDraw As Boolean, iconY0 As Long
    Const iconSize As Long = 16
    Set d = WzPlayerMuteStatusLabel_(self)
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset

    If Len(d.pText) = 0 Then
        Exit Sub
    End If

    muteStatus = WzPlayerMuteStatusLabel_getMuteStatus(self)

    d.wzText.setText d.pText, d.FontID

    fx = CHATOPT_PLAYER_BUTTON_EXTERNAL_PADDING_X + 16 + CHATOPT_PLAYER_BUTTON_INTERNAL_PADDING_X
    fy = (self.height() - d.wzText.lineSize()) \ 2 - d.wzText.aboveBase()
    maxTextDisplayableWidth = self.width() - fx
    isTruncated = maxTextDisplayableWidth < d.wzText.width()
    If isTruncated Then
        maxTextDisplayableWidth = maxTextDisplayableWidth - (iV_GetEllipsisWidth(d.wzText.getFontID()) + 2)
    End If
    textX0 = x0 + fx
    textY0 = y0 + fy

    textColor = psPalette(WZCOL_FORM_TEXT)
    If d.isDisabled Then
        textColor = pielightWithAlpha(textColor, pielightAlpha(textColor) \ 2)
    ElseIf muteStatus Then
        textColor = psPalette(WZCOL_TEXT_MEDIUM)
    End If
    d.wzText.render textX0, textY0, textColor, 0!, maxTextDisplayableWidth
    If isTruncated Then
        ' Render ellipsis
        iV_DrawEllipsis d.wzText.getFontID(), textX0 + maxTextDisplayableWidth + 2, textY0, textColor
    End If

    ImageID = IMAGE_INTFAC_VOLUME_UP
    If Not d.isDisabled Then imageAlpha = 255 Else imageAlpha = 125
    skipVolumeIconDraw = False
    If Not muteStatus And Not d.isDisabled Then
        ImageID = IMAGE_INTFAC_VOLUME_UP
        If d.isSelectedPlayer Then
            skipVolumeIconDraw = True
        End If
    Else
        ImageID = IMAGE_INTFAC_VOLUME_MUTE
    End If
    If Not skipVolumeIconDraw Then
        iconY0 = y0 + (self.height() - iconSize) \ 2
        If Not d.has_muteIconColorOverride Then
            iV_DrawImageFileAnisotropic IntImages, ImageID, x0 + CHATOPT_PLAYER_BUTTON_EXTERNAL_PADDING_X, iconY0, iconSize, iconSize, 0, imageAlpha
        Else
            iV_DrawImageTint IntImages, ImageID, x0 + CHATOPT_PLAYER_BUTTON_EXTERNAL_PADDING_X, iconY0, d.muteIconColorOverride, True, iconSize, iconSize
        End If
    End If
End Sub

Public Function WzPlayerMuteStatusLabel_idealWidth(ByVal self As WIDGET) As Long
    Dim d As WzPlayerMuteStatusLabel
    Set d = WzPlayerMuteStatusLabel_(self)
    WzPlayerMuteStatusLabel_idealWidth = iV_GetTextWidth(d.pText, d.FontID) + (CHATOPT_PLAYER_BUTTON_EXTERNAL_PADDING_X * 2) + CHATOPT_PLAYER_BUTTON_INTERNAL_PADDING_X + 16
End Function

Public Function WzPlayerMuteStatusLabel_requiredHeight(ByVal self As WIDGET) As Long
    Dim d As WzPlayerMuteStatusLabel
    Set d = WzPlayerMuteStatusLabel_(self)
    If Len(d.pText) = 0 Then
        WzPlayerMuteStatusLabel_requiredHeight = 0
        Exit Function
    End If

    WzPlayerMuteStatusLabel_requiredHeight = iV_GetTextLineSize(d.FontID)
End Function

Public Sub WzPlayerMuteStatusLabel_updateData(ByVal self As WIDGET)
    WzPlayerMuteStatusLabel_updateStringFromPlayerName self
End Sub

Public Function WzPlayerMuteStatusLabel_getMuteStatus(ByVal self As WIDGET) As Boolean
    ' display *applied* mute setting (including global and local settings)
    WzPlayerMuteStatusLabel_getMuteStatus = isPlayerMuted(WzPlayerMuteStatusLabel_(self).playerIdx)
End Function

Private Sub WzPlayerMuteStatusLabel_updateStringFromPlayerName(ByVal self As WIDGET)
    Dim d As WzPlayerMuteStatusLabel
    Set d = WzPlayerMuteStatusLabel_(self)
    ' keep player name updated
    If Not isHumanPlayer(d.playerIdx) Then
        d.pText = ""
        d.isSelectedPlayer = False
        Exit Sub
    End If
    d.pText = getPlayerName(d.playerIdx)
    d.isSelectedPlayer = (d.playerIdx = selectedPlayer)
End Sub

' MARK: - WzChatMuteColButtonBase

Public Sub WzChatMuteColButtonBase_construct(ByVal self As WIDGET)
    W_BUTTON_construct self
    self.kind = K_WzChatMuteColButtonBase ' (VB) virtual calls from here on reach WzChatMuteColButtonBase's overrides
    WzChatMuteColButtonBase_(self).has_muteIconColorOverride = False
    WzChatMuteColButtonBase_(self).hideAllowedIcon = False
End Sub

' virtual bool getMuteStatus() const = 0;
Public Function WzChatMuteColButtonBase_getMuteStatus(ByVal self As WIDGET) As Boolean
    ' (VB) pure virtual: never reached (WzLocalChatMuteColButton and WzGlobalChatMuteColButton override it)
End Function

' virtual void setMuteStatus(bool muted) = 0;
Public Sub WzChatMuteColButtonBase_setMuteStatus(ByVal self As WIDGET, ByVal muted As Boolean)
    ' (VB) pure virtual: never reached (WzLocalChatMuteColButton and WzGlobalChatMuteColButton override it)
End Sub

' special click-handling functions - subclasses can override
Public Sub WzChatMuteColButtonBase_clickPrimary(ByVal self As WIDGET)
    WzChatMuteColButtonBase_toggle self
End Sub

Public Function WzChatMuteColButtonBase_clickSecondary(ByVal self As WIDGET, ByVal synthesizedFromHold As Boolean) As Boolean
    WzChatMuteColButtonBase_clickSecondary = False
End Function

Public Sub WzChatMuteColButtonBase_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As WzChatMuteColButtonBase, x0 As Long, y0 As Long, isDown As Boolean, isDisabled As Boolean
    Dim isHighlight As Boolean, displayedWidth As Long, boxX0 As Long, boxX1 As Long, y1 As Long, light_border As Long
    Dim regular_border As Long, fill_color As Long, border As Long, muteStatus As Boolean, ImageID As Long
    Dim imageAlpha As Long, skipVolumeIconDraw As Boolean, iconX0 As Long, iconY0 As Long
    Const iconSize As Long = 16
    Set d = WzChatMuteColButtonBase_(self)
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset

    isDown = ((self.getState() And (WBUT_DOWN Or WBUT_LOCK Or WBUT_CLICKLOCK)) <> 0)
    isDisabled = (self.getState() And WBUT_DISABLE) <> 0
    isHighlight = Not isDisabled And ((self.getState() And WBUT_HIGHLIGHT) <> 0)

    If Not isDisabled Then
        ' Display the background and border
        displayedWidth = MIN_((CHATOPT_PLAYER_BUTTON_EXTERNAL_PADDING_X * 2) + 16, self.width())
        boxX0 = x0 + ((self.width() - displayedWidth) \ 2)
        boxX1 = boxX0 + displayedWidth
        y1 = y0 + self.height()
        light_border = WZCOL_CHATOPTIONS_BUTTON_BORDER_HIGHLIGHT
        If Not isDisabled Then regular_border = WZCOL_CHATOPTIONS_BUTTON_BORDER_LIGHT Else regular_border = psPalette(WZCOL_FORM_DARK)
        If isDown Or isDisabled Then fill_color = WZCOL_CHATOPTIONS_BUTTON_FILL_COLOR_DARK Else fill_color = WZCOL_CHATOPTIONS_BUTTON_FILL_COLOR
        If isDown Or isHighlight Then border = light_border Else border = regular_border
        iV_ShadowBox boxX0, y0, boxX1, y1, 0, border, border, fill_color
    End If

    muteStatus = vt_WzChatMuteColButtonBase_getMuteStatus(self)

    ImageID = IMAGE_INTFAC_VOLUME_UP
    If Not isDisabled Then imageAlpha = 255 Else imageAlpha = 125
    skipVolumeIconDraw = False
    If Not muteStatus Then
        ImageID = IMAGE_INTFAC_VOLUME_UP
        If d.hideAllowedIcon Then
            skipVolumeIconDraw = True
        End If
    Else
        ImageID = IMAGE_INTFAC_VOLUME_MUTE
    End If
    If Not skipVolumeIconDraw Then
        iconX0 = x0 + (self.width() - iconSize) \ 2
        iconY0 = y0 + (self.height() - iconSize) \ 2
        If Not d.has_muteIconColorOverride Then
            iV_DrawImageFileAnisotropic IntImages, ImageID, iconX0, iconY0, iconSize, iconSize, 0, imageAlpha
        Else
            iV_DrawImageTint IntImages, ImageID, iconX0, iconY0, d.muteIconColorOverride, True, iconSize, iconSize
        End If
    End If
End Sub

Public Function WzChatMuteColButtonBase_idealWidth(ByVal self As WIDGET) As Long
    WzChatMuteColButtonBase_idealWidth = WzChatMuteColButtonBase_getDefaultIdealWidth()
End Function

' static int32_t getDefaultIdealWidth()
Public Function WzChatMuteColButtonBase_getDefaultIdealWidth() As Long
    WzChatMuteColButtonBase_getDefaultIdealWidth = (CHATOPT_PLAYER_BUTTON_EXTERNAL_PADDING_X * 2) + 16
End Function

Public Sub WzChatMuteColButtonBase_toggle(ByVal self As WIDGET)
    vt_WzChatMuteColButtonBase_setMuteStatus self, Not vt_WzChatMuteColButtonBase_getMuteStatus(self)
End Sub

Public Sub WzChatMuteColButtonBase_setHideAllowedIcon(ByVal self As WIDGET, ByVal hide As Boolean)
    WzChatMuteColButtonBase_(self).hideAllowedIcon = hide
End Sub

' void setMuteIconColorOverride(optional<PIELIGHT> _muteIconColorOverride)
Public Sub WzChatMuteColButtonBase_setMuteIconColorOverride(ByVal self As WIDGET, ByVal has_muteIconColorOverride As Boolean, Optional ByVal muteIconColorOverride_ As Long = 0)
    WzChatMuteColButtonBase_(self).has_muteIconColorOverride = has_muteIconColorOverride
    WzChatMuteColButtonBase_(self).muteIconColorOverride = muteIconColorOverride_
End Sub

' override default button click handling
Public Sub WzChatMuteColButtonBase_released(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    Dim clickAndReleaseOnThisButton As Boolean
    clickAndReleaseOnThisButton = ((W_BUTTON_(self).state And WBUT_DOWN) <> 0) ' relies on W_BUTTON handling to properly set WBUT_DOWN

    W_BUTTON_released self, psContext, key

    If Not clickAndReleaseOnThisButton Then
        Exit Sub ' do nothing
    End If

    If key = WKEY_PRIMARY Then
        vt_WzChatMuteColButtonBase_clickPrimary self
    ElseIf key = WKEY_SECONDARY Then
        vt_WzChatMuteColButtonBase_clickSecondary self, False
    End If
End Sub

' MARK: - WzLocalChatMuteColButton

Public Sub WzLocalChatMuteColButton_construct(ByVal self As WIDGET)
    WzChatMuteColButtonBase_construct self
    self.kind = K_WzLocalChatMuteColButton ' (VB) virtual calls from here on reach WzLocalChatMuteColButton's overrides
    WzLocalChatMuteColButton_(self).playerIdx = 0
End Sub

Public Function WzLocalChatMuteColButton_make(ByVal playerIdx As Long) As WIDGET
    Dim button As WIDGET
    If Not ULT(playerIdx, MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "playerIdx < MAX_CONNECTED_PLAYERS", "src/hci/chatoptions.cpp:762", "make", "Invalid playerPosition": Exit Function
    Set button = WIDGET_alloc(K_WzLocalChatMuteColButton)
    WzLocalChatMuteColButton_construct button
    WzLocalChatMuteColButton_(button).playerIdx = playerIdx
    If playerIdx = selectedPlayer Then
        WzChatMuteColButtonBase_setHideAllowedIcon button, True
    End If
    button.setGeometry 0, 0, button.idealWidth(), CHATOPT_PLAYER_BUTTON_HEIGHT
    Set WzLocalChatMuteColButton_make = button
End Function

Public Sub WzLocalChatMuteColButton_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    Dim playerIdx As Long, isDisabled As Boolean, isHuman As Boolean, shouldBeDisabled As Boolean
    playerIdx = WzLocalChatMuteColButton_(self).playerIdx
    isDisabled = (self.getState() And WBUT_DISABLE) <> 0
    isHuman = isHumanPlayer(playerIdx)
    shouldBeDisabled = (playerIdx = selectedPlayer) Or Not isHuman
    If shouldBeDisabled <> isDisabled Then
        If shouldBeDisabled Then self.setState WBUT_DISABLE Else self.setState 0
    End If
End Sub

Public Function WzLocalChatMuteColButton_getMuteStatus(ByVal self As WIDGET) As Boolean
    WzLocalChatMuteColButton_getMuteStatus = ingame.muteChat(WzLocalChatMuteColButton_(self).playerIdx)
End Function

Public Sub WzLocalChatMuteColButton_setMuteStatus(ByVal self As WIDGET, ByVal muted As Boolean)
    Dim playerIdx As Long
    playerIdx = WzLocalChatMuteColButton_(self).playerIdx
    If Not isHumanPlayer(playerIdx) Then
        Exit Sub
    End If
    If playerIdx = selectedPlayer Then
        Exit Sub
    End If
    setPlayerMuted playerIdx, muted
End Sub

' MARK: - WzGlobalChatMuteColButton

Public Sub WzGlobalChatMuteColButton_construct(ByVal self As WIDGET)
    WzChatMuteColButtonBase_construct self
    self.kind = K_WzGlobalChatMuteColButton ' (VB) virtual calls from here on reach WzGlobalChatMuteColButton's overrides
    WzGlobalChatMuteColButton_(self).has_playerIdx = False
    WzGlobalChatMuteColButton_(self).readOnly = False
End Sub

' static std::shared_ptr<WzGlobalChatMuteColButton> make(optional<uint32_t> playerIdx,
' std::shared_ptr<WzChatOptionsForm> parentForm, bool readOnly)
Public Function WzGlobalChatMuteColButton_make(ByVal has_playerIdx As Boolean, ByVal playerIdx As Long, ByVal parentForm As WIDGET, ByVal readOnly As Boolean) As WIDGET
    Dim button As WIDGET, d As WzGlobalChatMuteColButton, valueOr0 As Long
    If has_playerIdx Then valueOr0 = playerIdx Else valueOr0 = 0
    If Not ULT(valueOr0, MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "playerIdx.value_or(0) < MAX_CONNECTED_PLAYERS", "src/hci/chatoptions.cpp:814", "make", "Invalid playerPosition": Exit Function
    Set button = WIDGET_alloc(K_WzGlobalChatMuteColButton)
    WzGlobalChatMuteColButton_construct button
    Set d = WzGlobalChatMuteColButton_(button)
    d.has_playerIdx = has_playerIdx
    d.playerIdx = playerIdx
    If Not parentForm Is Nothing Then d.parentForm = WzChatOptionsForm_(parentForm).weakSerial
    d.readOnly = readOnly
    WzChatMuteColButtonBase_setHideAllowedIcon button, True
    WzGlobalChatMuteColButton_updateButtonStatus button
    button.setGeometry 0, 0, button.idealWidth(), CHATOPT_PLAYER_BUTTON_HEIGHT
    Set WzGlobalChatMuteColButton_make = button
End Function

Public Sub WzGlobalChatMuteColButton_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    WzGlobalChatMuteColButton_updateButtonStatus self
End Sub

Public Function WzGlobalChatMuteColButton_getMuteStatus(ByVal self As WIDGET) As Boolean
    Dim d As WzGlobalChatMuteColButton
    Set d = WzGlobalChatMuteColButton_(self)
    If d.has_playerIdx Then
        WzGlobalChatMuteColButton_getMuteStatus = Not ingame.hostChatPermissions(d.playerIdx)
    Else
        WzGlobalChatMuteColButton_getMuteStatus = Not NETgetDefaultMPHostFreeChatPreference()
    End If
End Function

Public Sub WzGlobalChatMuteColButton_setMuteStatus(ByVal self As WIDGET, ByVal muted As Boolean)
    Dim d As WzGlobalChatMuteColButton, strongParent As WIDGET
    Set d = WzGlobalChatMuteColButton_(self)
    If d.readOnly Then
        Exit Sub
    End If
    If Not d.has_playerIdx Then
        NETsetDefaultMPHostFreeChatPreference Not muted
        ' do not update any existing players - this just sets the default for new joins
        Exit Sub
    End If
    If Not isHumanPlayer(d.playerIdx) Then
        Exit Sub
    End If
    ingame.hostChatPermissions(d.playerIdx) = Not muted
    If NetPlay_isHost And NetPlay_bComms Then
        sendHostConfig
    End If
    Set strongParent = WeakWidget_lock(d.parentForm)
    If Not strongParent Is Nothing Then
        WzChatOptionsForm_callOnSettingsChanged strongParent
    End If
End Sub

Private Sub WzGlobalChatMuteColButton_updateButtonStatus(ByVal self As WIDGET)
    Dim d As WzGlobalChatMuteColButton, isDisabled As Boolean, isHuman As Boolean, shouldBeDisabled As Boolean
    Set d = WzGlobalChatMuteColButton_(self)
    If d.has_playerIdx Then
        isDisabled = (self.getState() And WBUT_DISABLE) <> 0
        isHuman = isHumanPlayer(d.playerIdx)
        shouldBeDisabled = d.readOnly Or Not isHuman
        If shouldBeDisabled <> isDisabled Then
            If shouldBeDisabled Then self.setState WBUT_DISABLE Else self.setState 0
        End If
    End If
End Sub

' MARK: - WzChatOptionsform

' WzChatOptionsForm(): W_CLICKFORM() {}
Public Sub WzChatOptionsForm_construct(ByVal self As WIDGET)
    W_CLICKFORM_construct self
    self.kind = K_WzChatOptionsForm ' (VB) virtual calls from here on reach WzChatOptionsForm's overrides
    WzChatOptionsForm_(self).showingHostingOptions = False
    WzChatOptionsForm_(self).weakSerial = WeakWidget_register(self)   ' (VB) its std::weak_ptrs
End Sub

Public Sub WzChatOptionsForm_destruct(ByVal self As WIDGET)
    WeakWidget_unregister WzChatOptionsForm_(self).weakSerial
End Sub

Public Function WzChatOptionsForm_make(ByVal showHostOptions As Boolean, ByVal onSettingsChange As WzFunc) As WIDGET
    Dim result As WIDGET
    Set result = WIDGET_alloc(K_WzChatOptionsForm)
    WzChatOptionsForm_construct result
    Set WzChatOptionsForm_(result).onSettingsChange = onSettingsChange
    WzChatOptionsForm_initialize result, showHostOptions
    Set WzChatOptionsForm_make = result
End Function

Public Sub WzChatOptionsForm_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x0 As Long, y0 As Long, x1 As Long, y1 As Long, showHighlighted As Boolean, borderColor As Long
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    x1 = x0 + self.width()
    y1 = y0 + self.height()

    showHighlighted = True
    If Not showHighlighted Then borderColor = pal_RGBA(0, 0, 0, 255) Else borderColor = pal_RGBA(255, 255, 255, 100)

    ' draw background + border
    pie_UniTransBoxFill x0, y0, x1, y1, pal_RGBA(0, 0, 0, 150)
    iV_TransBoxFill x0, y0, x1, y1
    iV_Box x0, y0, x1, y1, borderColor
End Sub

Private Function WzChatOptionsForm_requiredHeightForTextAndButtons(ByVal self As WIDGET) As Long
    Dim d As WzChatOptionsForm, result As Long
    Set d = WzChatOptionsForm_(self)
    result = 0
    If d.showingHostingOptions Then
        result = result + CHATOPT_FORM_EXTERNAL_PADDING + W_LABEL_requiredHeight(d.hostChatOptionsLabel)
        result = result + CHATOPT_FORM_INTERNAL_PADDING + W_LABEL_requiredHeight(d.globalOptionsLabel)
        result = result + CHATOPT_FORM_INTERNAL_PADDING + MAX_(MAX_(d.freechatEnableAllButton.height(), d.freechatDisableAllButton.height()), d.joiningClientsButton.height())
        result = result + CHATOPT_FORM_INTERNAL_PADDING
    Else
        result = result + (CHATOPT_FORM_EXTERNAL_PADDING - CHATOPT_FORM_INTERNAL_PADDING)
    End If

    result = result + CHATOPT_FORM_INTERNAL_PADDING + W_LABEL_requiredHeight(d.chatOptionsLabel)
    result = result + CHATOPT_FORM_INTERNAL_PADDING + W_LABEL_requiredHeight(d.chatOptionsDescriptionLabel)
    result = result + CHATOPT_FORM_INTERNAL_PADDING + W_LABEL_requiredHeight(d.chatOptionsHostDisabledMsgLabel)
    ' do not include players table
    result = result + CHATOPT_FORM_EXTERNAL_PADDING
    WzChatOptionsForm_requiredHeightForTextAndButtons = result
End Function

Public Function WzChatOptionsForm_idealHeight(ByVal self As WIDGET) As Long
    Dim result As Long
    result = WzChatOptionsForm_requiredHeightForTextAndButtons(self)
    result = result + CHATOPT_FORM_INTERNAL_PADDING + WzChatOptionsForm_(self).playersTable.idealHeight()
    WzChatOptionsForm_idealHeight = result
End Function

Public Function WzChatOptionsForm_idealWidth(ByVal self As WIDGET) As Long
    Dim d As WzChatOptionsForm, result As Long, maxColumnDataWidths() As Long, maxColumnDataWidthsCount As Long
    Dim totalNeededColumnWidth As Long, idealTableWidth As Long
    Set d = WzChatOptionsForm_(self)
    result = 0

    If d.showingHostingOptions Then
        result = MAX_(result, (CHATOPT_FORM_EXTERNAL_PADDING * 2) + d.globalOptionsLabel.idealWidth())
        If Not d.joiningClientsButton Is Nothing And Not d.freechatEnableAllButton Is Nothing And Not d.freechatDisableAllButton Is Nothing Then
            result = MAX_(result, (CHATOPT_FORM_EXTERNAL_PADDING + d.joiningClientsButton.width() + CHATOPT_FORM_INTERNAL_PADDING + d.freechatEnableAllButton.width() + CHATOPT_FORM_INTERNAL_PADDING + d.freechatDisableAllButton.width() + CHATOPT_FORM_EXTERNAL_PADDING))
        End If
    End If
    result = MAX_(result, (CHATOPT_FORM_EXTERNAL_PADDING * 2) + d.chatOptionsLabel.idealWidth())
    result = MAX_(result, (CHATOPT_FORM_EXTERNAL_PADDING * 2) + W_LABEL_getMaxLineWidth(d.chatOptionsDescriptionLabel))
    result = MAX_(result, (CHATOPT_FORM_EXTERNAL_PADDING * 2) + W_LABEL_getMaxLineWidth(d.chatOptionsHostDisabledMsgLabel))

    totalNeededColumnWidth = WzChatOptionsForm_getMaxTableColumnDataWidths(self, maxColumnDataWidths, maxColumnDataWidthsCount)
    idealTableWidth = ScrollableTableWidget_getTableWidthNeededForTotalColumnWidth(d.playersTable, ScrollableTableWidget_getNumColumns(d.playersTable), totalNeededColumnWidth)
    result = MAX_(result, (CHATOPT_FORM_EXTERNAL_PADDING * 2) + idealTableWidth)

    WzChatOptionsForm_idealWidth = result
End Function

Private Function hostHasMutedAnyPlayers() As Boolean
    Dim i As Long
    For i = 0 To MAX_CONNECTED_PLAYERS - 1
        If Not isHumanPlayer(i) Then
            GoTo continue_
        End If
        If Not ingame.hostChatPermissions(i) Then
            hostHasMutedAnyPlayers = True
            Exit Function
        End If
continue_:
    Next
    hostHasMutedAnyPlayers = False
End Function

Public Sub WzChatOptionsForm_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    If hostHasMutedAnyPlayers() Then
        WzChatOptionsForm_(self).chatOptionsHostDisabledMsgLabel.show
    Else
        WzChatOptionsForm_(self).chatOptionsHostDisabledMsgLabel.hide
    End If
End Sub

Public Sub WzChatOptionsForm_geometryChanged(ByVal self As WIDGET)
    Dim d As WzChatOptionsForm, w As Long, h As Long, maxPlayersTableWidth As Long, totalAvailableHeightForPlayersTable As Long
    Dim currentY0 As Long, lastButtonX1 As Long, maxColumnDataWidths() As Long, maxColumnDataWidthsCount As Long
    Dim totalNeededColumnWidth As Long, idealTableWidth As Long, playersTableWidth As Long
    Set d = WzChatOptionsForm_(self)
    w = self.width()
    h = self.height()

    If w = 0 Or h = 0 Then
        Exit Sub
    End If

    maxPlayersTableWidth = MAX_((w - (CHATOPT_FORM_EXTERNAL_PADDING * 2)), 0)
    totalAvailableHeightForPlayersTable = h - WzChatOptionsForm_requiredHeightForTextAndButtons(self) - CHATOPT_FORM_INTERNAL_PADDING

    If totalAvailableHeightForPlayersTable < 0 Then
        ' not enough height to display the lists at all?
        Exit Sub
    End If

    currentY0 = 0

    d.updateButton.setGeometry w - (CHATOPT_FORM_EXTERNAL_PADDING + d.updateButton.width()), CHATOPT_FORM_EXTERNAL_PADDING, d.updateButton.width(), d.updateButton.height()

    If d.showingHostingOptions Then
        currentY0 = d.globalOptionsLabel.y() + d.globalOptionsLabel.height()

        lastButtonX1 = 0
        If GetGameMode() <> GS_NORMAL Then ' if not yet in a game
            d.joiningClientsButton.setGeometry CHATOPT_FORM_EXTERNAL_PADDING, currentY0 + CHATOPT_FORM_INTERNAL_PADDING, d.joiningClientsButton.width(), d.joiningClientsButton.height()
            lastButtonX1 = d.joiningClientsButton.x() + d.joiningClientsButton.width()
        Else
            ' in game, hide this button
            d.joiningClientsButton.hide
            lastButtonX1 = (CHATOPT_FORM_EXTERNAL_PADDING - CHATOPT_FORM_INTERNAL_PADDING)
        End If

        d.freechatEnableAllButton.setGeometry lastButtonX1 + CHATOPT_FORM_INTERNAL_PADDING, currentY0 + CHATOPT_FORM_INTERNAL_PADDING, d.freechatEnableAllButton.width(), d.freechatEnableAllButton.height()
        lastButtonX1 = d.freechatEnableAllButton.x() + d.freechatEnableAllButton.width()

        d.freechatDisableAllButton.setGeometry lastButtonX1 + CHATOPT_FORM_INTERNAL_PADDING, currentY0 + CHATOPT_FORM_INTERNAL_PADDING, d.freechatDisableAllButton.width(), d.freechatDisableAllButton.height()

        currentY0 = MAX_(d.globalOptionsLabel.y() + d.globalOptionsLabel.height(), d.freechatDisableAllButton.y() + d.freechatDisableAllButton.height()) + CHATOPT_FORM_INTERNAL_PADDING
    Else
        currentY0 = (CHATOPT_FORM_EXTERNAL_PADDING - CHATOPT_FORM_INTERNAL_PADDING)
    End If

    d.chatOptionsLabel.setGeometry CHATOPT_FORM_EXTERNAL_PADDING, currentY0 + CHATOPT_FORM_INTERNAL_PADDING, W_LABEL_getMaxLineWidth(d.chatOptionsLabel), W_LABEL_requiredHeight(d.chatOptionsLabel)

    currentY0 = d.chatOptionsLabel.y() + d.chatOptionsLabel.height()

    d.chatOptionsDescriptionLabel.setGeometry CHATOPT_FORM_EXTERNAL_PADDING, currentY0 + CHATOPT_FORM_INTERNAL_PADDING, W_LABEL_getMaxLineWidth(d.chatOptionsDescriptionLabel), W_LABEL_requiredHeight(d.chatOptionsDescriptionLabel)
    currentY0 = d.chatOptionsDescriptionLabel.y() + d.chatOptionsDescriptionLabel.height()

    totalNeededColumnWidth = WzChatOptionsForm_getMaxTableColumnDataWidths(self, maxColumnDataWidths, maxColumnDataWidthsCount)
    idealTableWidth = ScrollableTableWidget_getTableWidthNeededForTotalColumnWidth(d.playersTable, ScrollableTableWidget_getNumColumns(d.playersTable), totalNeededColumnWidth)

    playersTableWidth = MIN_(maxPlayersTableWidth, idealTableWidth)
    d.playersTable.setGeometry CHATOPT_FORM_EXTERNAL_PADDING, currentY0 + CHATOPT_FORM_INTERNAL_PADDING, playersTableWidth, totalAvailableHeightForPlayersTable

    currentY0 = d.playersTable.y() + d.playersTable.height()

    d.chatOptionsHostDisabledMsgLabel.setGeometry CHATOPT_FORM_EXTERNAL_PADDING, currentY0 + CHATOPT_FORM_INTERNAL_PADDING, W_LABEL_getMaxLineWidth(d.chatOptionsHostDisabledMsgLabel), W_LABEL_requiredHeight(d.chatOptionsHostDisabledMsgLabel)
End Sub

' std::pair<std::vector<size_t>, size_t> getMaxTableColumnDataWidths(): (VB) .second returned, .first in
' maxColumnDataWidths(0 .. maxColumnDataWidthsCount - 1)
Private Function WzChatOptionsForm_getMaxTableColumnDataWidths(ByVal self As WIDGET, maxColumnDataWidths() As Long, maxColumnDataWidthsCount As Long) As Long
    Dim d As WzChatOptionsForm, totalNeededColumnWidth As Long, minimumColumnWidths() As Long, minimumColumnWidthsCount As Long
    Dim colIdx As Long, maxIdealContentWidth As Long, maxIdealContentWidth_sizet As Long
    Set d = WzChatOptionsForm_(self)
    totalNeededColumnWidth = 0
    maxColumnDataWidthsCount = 0
    Erase maxColumnDataWidths
    minimumColumnWidthsCount = ScrollableTableWidget_getMinimumColumnWidths(d.playersTable, minimumColumnWidths)
    If Not (minimumColumnWidthsCount = ScrollableTableWidget_getNumColumns(d.playersTable)) Then ASSERT_FAILURE "minimumColumnWidths.size() == playersTable->getNumColumns()", "src/hci/chatoptions.cpp:1078", "getMaxTableColumnDataWidths", "Number of minimum column widths does not match number of columns!"
    For colIdx = 0 To ScrollableTableWidget_getNumColumns(d.playersTable) - 1
        maxIdealContentWidth = ScrollableTableWidget_getColumnMaxContentIdealWidth(d.playersTable, colIdx)
        ' std::max<int32_t>({maxIdealContentWidth, 0, static_cast<int32_t>(minimumColumnWidths.at(colIdx))})
        maxIdealContentWidth = MAX_(MAX_(maxIdealContentWidth, 0), minimumColumnWidths(colIdx))
        If colIdx = playerNameColIdx Then
            ' restrict player name column to a fixed maximum width
            maxIdealContentWidth = MIN_(maxIdealContentWidth, MAX_PLAYER_NAME_COLUMN_WIDTH)
        End If
        maxIdealContentWidth_sizet = maxIdealContentWidth
        ReDim Preserve maxColumnDataWidths(0 To maxColumnDataWidthsCount)
        maxColumnDataWidths(maxColumnDataWidthsCount) = maxIdealContentWidth_sizet
        maxColumnDataWidthsCount = maxColumnDataWidthsCount + 1
        totalNeededColumnWidth = totalNeededColumnWidth + maxIdealContentWidth_sizet
    Next
    WzChatOptionsForm_getMaxTableColumnDataWidths = totalNeededColumnWidth
End Function

Private Function WzChatOptionsForm_newPlayerMuteRow(ByVal self As WIDGET, ByVal playerIdx As Long, ByVal allowGlobalConfig As Boolean) As WIDGET
    Dim columnWidgets As Collection, sharedSelf As WIDGET, playerNameLabel As WIDGET, globalMuteButton As WIDGET
    Dim localMuteButton As WIDGET, rowHeight As Long, i As Long, col As WIDGET, row As WIDGET
    If Not ULT(playerIdx, MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "playerIdx < MAX_CONNECTED_PLAYERS", "src/hci/chatoptions.cpp:1097", "newPlayerMuteRow", "Invalid playerIdx": Exit Function

    Set columnWidgets = New Collection   ' std::vector<std::shared_ptr<WIDGET>> columnWidgets;

    Set sharedSelf = self

    ' Player Name widget
    Set playerNameLabel = WzPlayerMuteStatusLabel_make(playerIdx)
    If Not (Not playerNameLabel Is Nothing) Then ASSERT_FAILURE "playerNameLabel != nullptr", "src/hci/chatoptions.cpp:1105", "newPlayerMuteRow", "Failed to make widget": Exit Function
    playerNameLabel.setGeometry 0, 0, playerNameLabel.idealWidth(), WzPlayerMuteStatusLabel_requiredHeight(playerNameLabel)
    playerNameLabel.setTransparentToMouse True
    columnWidgets.Add playerNameLabel

    ' Global Mute column
    Set globalMuteButton = WzGlobalChatMuteColButton_make(True, playerIdx, sharedSelf, Not allowGlobalConfig)
    If Not (Not globalMuteButton Is Nothing) Then ASSERT_FAILURE "globalMuteButton != nullptr", "src/hci/chatoptions.cpp:1112", "newPlayerMuteRow", "Failed to make widget": Exit Function
    columnWidgets.Add globalMuteButton

    ' Local Mute column
    Set localMuteButton = WzLocalChatMuteColButton_make(playerIdx)
    If Not (Not localMuteButton Is Nothing) Then ASSERT_FAILURE "localMuteButton != nullptr", "src/hci/chatoptions.cpp:1117", "newPlayerMuteRow", "Failed to make widget": Exit Function
    If playerIdx = selectedPlayer Then
        localMuteButton.setState WBUT_DISABLE
    End If
    columnWidgets.Add localMuteButton

    rowHeight = 0
    For i = 1 To columnWidgets.Count
        Set col = columnWidgets(i)
        rowHeight = MAX_(rowHeight, col.idealHeight())
    Next

    playerNameLabel.setGeometry 0, 0, playerNameLabel.idealWidth(), rowHeight

    Set row = TableRow_make(columnWidgets, rowHeight)
    TableRow_setDisabledColor row, psPalette(WZCOL_FORM_DARK)
    TableRow_setHighlightsOnMouseOver row, True

    Set WzChatOptionsForm_newPlayerMuteRow = row
End Function

Private Sub WzChatOptionsForm_refreshData(ByVal self As WIDGET)
    Dim d As WzChatOptionsForm, playerIndexes As UInt32Vector, spectatorIndexes As UInt32Vector, player As Long, i As Long
    Set d = WzChatOptionsForm_(self)
    ' Add rows for all players
    For player = 0 To MIN_(game.maxPlayers, MAX_PLAYERS) - 1
        If isHumanPlayer(player) Then ' is an active (connected) human
            UInt32Vector_push_back playerIndexes, player
        End If
    Next
    ' sort by player position
    UInt32Vector_sortByPlayerPosition playerIndexes

    ' Rows for all spectators
    For player = MAX_PLAYERS To MAX_CONNECTED_PLAYERS - 1
        If isHumanPlayer(player) And NetPlay_players_isSpectator(player) Then ' is an active (connected) human, and a spectator
            UInt32Vector_push_back spectatorIndexes, player
        End If
    Next

    ScrollableTableWidget_clearRows d.playersTable
    ' Add players to playersTable
    For i = 0 To playerIndexes.n - 1
        ScrollableTableWidget_addRow d.playersTable, WzChatOptionsForm_newPlayerMuteRow(self, playerIndexes.v(i), d.showingHostingOptions)
    Next
    ' Add spectators to playersTable
    For i = 0 To spectatorIndexes.n - 1
        ScrollableTableWidget_addRow d.playersTable, WzChatOptionsForm_newPlayerMuteRow(self, spectatorIndexes.v(i), d.showingHostingOptions)
    Next

    self.geometryChanged
End Sub

Private Function createColHeaderLabel(ByVal text As String) As WIDGET   ' (, int minLeftPadding)
    Dim label As WIDGET
    Set label = W_LABEL_new()
    W_LABEL_setTextAlignment label, WLAB_ALIGNCENTRE
    label.setString text
    label.setGeometry 0, 0, W_LABEL_getMaxLineWidth(label), 0
    W_LABEL_setCacheNeverExpires label, True
    W_LABEL_setCanTruncate label, True
    Set createColHeaderLabel = Margin_wrap(CHATOPT_COLUMN_HEADER_PADDING_Y, CHATOPT_COLUMN_HEADER_PADDING_X, CHATOPT_COLUMN_HEADER_PADDING_Y, CHATOPT_COLUMN_HEADER_PADDING_X, label)
End Function

Private Sub WzChatOptionsForm_initialize(ByVal self As WIDGET, ByVal showHostOptions As Boolean)
    Dim d As WzChatOptionsForm, currentY0 As Long, sharedSelf As WIDGET, weakSelf As Long, columns() As TableColumn
    Dim columnsCount As Long, maxConfigOptColumnWidth As Long, i As Long, maxIdealColumnHeight As Long
    Dim minimumColumnWidths() As Long, minimumColumnWidthsCount As Long, playersTableColumnHeaderHeight As Long
    Set d = WzChatOptionsForm_(self)
    d.showingHostingOptions = showHostOptions

    currentY0 = 0

    ' top-right controls
    Set d.updateButton = WzChatOptionsButton_make(ChrW$(&H21BB), False) ' "(U+21BB: the clockwise open circle arrow)"
    d.updateButton.setGeometry d.updateButton.x(), d.updateButton.y(), MAX_(WzChatOptionsButton_getButtonTextWidth(d.updateButton) + CHATOPT_FORM_INTERNAL_PADDING, CHATOPT_PLAYER_BUTTON_HEIGHT), CHATOPT_PLAYER_BUTTON_HEIGHT
    d.updateButton.setTip "Refresh"
    self.attach d.updateButton
    W_BUTTON_addOnClickHandler d.updateButton, WzFunc_new(FN_chatOpt_updateButton_onClick)

    Set sharedSelf = self
    weakSelf = d.weakSerial

    If showHostOptions Then
        Set d.hostChatOptionsLabel = W_LABEL_new()
        W_LABEL_setFont d.hostChatOptionsLabel, font_medium_bold, psPalette(WZCOL_TEXT_BRIGHT)
        d.hostChatOptionsLabel.setString "Host Chat Options"
        self.attach d.hostChatOptionsLabel
        d.hostChatOptionsLabel.setGeometry CHATOPT_FORM_EXTERNAL_PADDING, currentY0 + CHATOPT_FORM_EXTERNAL_PADDING, W_LABEL_getMaxLineWidth(d.hostChatOptionsLabel), W_LABEL_requiredHeight(d.hostChatOptionsLabel)
        d.hostChatOptionsLabel.setTransparentToMouse True

        currentY0 = d.hostChatOptionsLabel.y() + d.hostChatOptionsLabel.height()

        Set d.globalOptionsLabel = W_LABEL_new()
        W_LABEL_setFont d.globalOptionsLabel, font_regular_bold, psPalette(WZCOL_TEXT_MEDIUM)
        d.globalOptionsLabel.setString "Global Options:"
        self.attach d.globalOptionsLabel
        d.globalOptionsLabel.setGeometry CHATOPT_FORM_EXTERNAL_PADDING, currentY0 + CHATOPT_FORM_INTERNAL_PADDING, W_LABEL_getMaxLineWidth(d.globalOptionsLabel), W_LABEL_requiredHeight(d.globalOptionsLabel)
        d.globalOptionsLabel.setTransparentToMouse True

        ' Add "Enable All" and "Disable All" buttons
        Set d.freechatEnableAllButton = WzChatOptionsButton_make("Enable All", False)
        self.attach d.freechatEnableAllButton
        W_BUTTON_addOnClickHandler d.freechatEnableAllButton, WzFunc_new(FN_chatOpt_enableAll_onClick, weakSelf)
        d.freechatEnableAllButton.setGeometry CHATOPT_FORM_EXTERNAL_PADDING, currentY0 + CHATOPT_FORM_INTERNAL_PADDING, WzChatOptionsButton_getButtonTextWidth(d.freechatEnableAllButton) + CHATOPT_FORM_INTERNAL_PADDING, CHATOPT_PLAYER_BUTTON_HEIGHT
        Set d.freechatDisableAllButton = WzChatOptionsButton_make("Disable All", False)
        self.attach d.freechatDisableAllButton
        W_BUTTON_addOnClickHandler d.freechatDisableAllButton, WzFunc_new(FN_chatOpt_disableAll_onClick, weakSelf)
        d.freechatDisableAllButton.setGeometry CHATOPT_FORM_EXTERNAL_PADDING, currentY0 + CHATOPT_FORM_INTERNAL_PADDING, WzChatOptionsButton_getButtonTextWidth(d.freechatDisableAllButton) + CHATOPT_FORM_INTERNAL_PADDING, CHATOPT_PLAYER_BUTTON_HEIGHT

        ' Add "New Joins" button
        Set d.joiningClientsButton = WzHostFreeChatPlayerMuteButton_make(False, 0, sharedSelf)
        self.attach d.joiningClientsButton
        d.joiningClientsButton.setGeometry CHATOPT_FORM_EXTERNAL_PADDING, currentY0 + CHATOPT_FORM_INTERNAL_PADDING, d.joiningClientsButton.width(), d.joiningClientsButton.height()

        currentY0 = MAX_(d.globalOptionsLabel.y() + d.globalOptionsLabel.height(), d.freechatDisableAllButton.y() + d.freechatDisableAllButton.height()) + CHATOPT_FORM_INTERNAL_PADDING
    End If

    Set d.chatOptionsLabel = W_LABEL_new()
    W_LABEL_setFont d.chatOptionsLabel, font_medium_bold, psPalette(WZCOL_TEXT_BRIGHT)
    d.chatOptionsLabel.setString "Chat Options"
    self.attach d.chatOptionsLabel
    d.chatOptionsLabel.setGeometry CHATOPT_FORM_EXTERNAL_PADDING, currentY0 + CHATOPT_FORM_INTERNAL_PADDING, W_LABEL_getMaxLineWidth(d.chatOptionsLabel), W_LABEL_requiredHeight(d.chatOptionsLabel)
    d.chatOptionsLabel.setTransparentToMouse True
    currentY0 = d.chatOptionsLabel.y() + d.chatOptionsLabel.height()

    Set d.chatOptionsDescriptionLabel = W_LABEL_new()
    W_LABEL_setFont d.chatOptionsDescriptionLabel, font_regular, psPalette(WZCOL_TEXT_MEDIUM)
    d.chatOptionsDescriptionLabel.setString "Mute or configure player free chat."
    W_LABEL_setCanTruncate d.chatOptionsDescriptionLabel, True
    self.attach d.chatOptionsDescriptionLabel
    d.chatOptionsDescriptionLabel.setGeometry CHATOPT_FORM_EXTERNAL_PADDING, currentY0 + CHATOPT_FORM_INTERNAL_PADDING, W_LABEL_getMaxLineWidth(d.chatOptionsDescriptionLabel), W_LABEL_requiredHeight(d.chatOptionsDescriptionLabel)
    d.chatOptionsDescriptionLabel.setTransparentToMouse True
    currentY0 = d.chatOptionsDescriptionLabel.y() + d.chatOptionsDescriptionLabel.height()

    Set d.chatOptionsHostDisabledMsgLabel = W_LABEL_new()
    W_LABEL_setFont d.chatOptionsHostDisabledMsgLabel, font_regular, WZCOL_CHATOPTIONS_HOST_GLOBAL_SET_COLOR
    d.chatOptionsHostDisabledMsgLabel.setString "The host has globally muted 1+ players."
    W_LABEL_setCanTruncate d.chatOptionsHostDisabledMsgLabel, True
    self.attach d.chatOptionsHostDisabledMsgLabel
    d.chatOptionsHostDisabledMsgLabel.setGeometry CHATOPT_FORM_EXTERNAL_PADDING, currentY0 + CHATOPT_FORM_INTERNAL_PADDING, W_LABEL_getMaxLineWidth(d.chatOptionsHostDisabledMsgLabel), W_LABEL_requiredHeight(d.chatOptionsHostDisabledMsgLabel)
    d.chatOptionsHostDisabledMsgLabel.setTransparentToMouse True
    d.chatOptionsHostDisabledMsgLabel.hide

    ' Table of players

    ' Create column headers for players table
    ' (std::vector<TableColumn> columns; three push_backs)
    ReDim columns(0 To 2)
    columns(0) = TableColumn_(createColHeaderLabel(" "), TableColumn_ResizeBehavior_FIXED_WIDTH)
    columns(1) = TableColumn_(createColHeaderLabel("Global"), TableColumn_ResizeBehavior_FIXED_WIDTH)
    columns(2) = TableColumn_(createColHeaderLabel("Local"), TableColumn_ResizeBehavior_FIXED_WIDTH)
    columnsCount = 3

    maxConfigOptColumnWidth = WzChatMuteColButtonBase_getDefaultIdealWidth()
    For i = 1 To columnsCount - 1
        maxConfigOptColumnWidth = MAX_(maxConfigOptColumnWidth, columns(i).columnWidget.idealWidth())
    Next

    maxIdealColumnHeight = 0
    ' std::vector<size_t> minimumColumnWidths;
    ReDim minimumColumnWidths(0 To columnsCount - 1)
    minimumColumnWidths(0) = 200 ' fixed minimum width for player names
    minimumColumnWidthsCount = 1
    For i = 1 To columnsCount - 1
        ' std::max<int>({column.columnWidget->idealWidth(), 0, maxConfigOptColumnWidth})
        minimumColumnWidths(minimumColumnWidthsCount) = MAX_(MAX_(columns(i).columnWidget.idealWidth(), 0), maxConfigOptColumnWidth)
        minimumColumnWidthsCount = minimumColumnWidthsCount + 1
        maxIdealColumnHeight = MAX_(maxIdealColumnHeight, columns(i).columnWidget.idealHeight())
    Next

    playersTableColumnHeaderHeight = maxIdealColumnHeight
    Set d.playersTable = ScrollableTableWidget_make(columns, columnsCount, playersTableColumnHeaderHeight)
    ScrollableTableWidget_setColumnPadding d.playersTable, Vector2i_(CHATOPT_COLUMN_PADDING_X, 0)
    self.attach d.playersTable
    ScrollableTableWidget_setBackgroundColor d.playersTable, pal_RGBA(0, 0, 0, 0)
    ScrollableTableWidget_setMinimumColumnWidths d.playersTable, minimumColumnWidths, minimumColumnWidthsCount
    ScrollableTableWidget_setDrawColumnLines d.playersTable, True
    ScrollableTableWidget_setItemSpacing d.playersTable, 2

    WzChatOptionsForm_refreshData self
End Sub

Public Sub WzChatOptionsForm_callOnSettingsChanged(ByVal self As WIDGET)
    Dim d As WzChatOptionsForm
    Set d = WzChatOptionsForm_(self)
    If Not d.onSettingsChange Is Nothing Then
        WzFunc_call d.onSettingsChange
    End If
End Sub

Private Sub WzChatOptionsForm_changeFreeChatForAll(ByVal self As WIDGET, ByVal enabled As Boolean)
    Dim player As Long
    NETsetDefaultMPHostFreeChatPreference enabled
    If NetPlay_isHost And NetPlay_bComms Then
        ' update all existing slots
        For player = 0 To MAX_CONNECTED_PLAYERS - 1
            ingame.hostChatPermissions(player) = NETgetDefaultMPHostFreeChatPreference()
        Next
        sendHostConfig
        WzChatOptionsForm_callOnSettingsChanged self
    End If
End Sub

' MARK: - Public functions

Public Function createChatOptionsForm(ByVal showHostOptions As Boolean, ByVal onSettingsChange As WzFunc) As WIDGET
    Dim result As WIDGET
    Set result = WzChatOptionsForm_make(showHostOptions, onSettingsChange)
    Set createChatOptionsForm = result
End Function

'--------------------------------------------------------------------------
' the lambdas

Public Function chatoptions_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim button As WIDGET, psParent As WIDGET, strongSelf As WIDGET
    Select Case f.id
    Case FN_chatOpt_updateButton_onClick        ' [](W_BUTTON& button)
        Set button = args(0)
        Set psParent = button.parent()
        If Not WIDGET_isKindOf(psParent, K_WzChatOptionsForm) Then Set psParent = Nothing
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/hci/chatoptions.cpp:1205", "initialize", "No parent": Exit Function
        widgScheduleTask WzFunc_new(FN_chatOpt_updateButton_task, psParent)

    Case FN_chatOpt_updateButton_task           ' [psParent]()
        Set psParent = f.cap(0)
        WzChatOptionsForm_refreshData psParent
        psParent.callCalcLayout

    Case FN_chatOpt_enableAll_onClick           ' [weakSelf](W_BUTTON& widg)
        Set strongSelf = WeakWidget_lock(f.cap(0))
        If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/hci/chatoptions.cpp:1238", "initialize", "Null form?": Exit Function
        WzChatOptionsForm_changeFreeChatForAll strongSelf, True

    Case FN_chatOpt_disableAll_onClick          ' [weakSelf](W_BUTTON& widg)
        Set strongSelf = WeakWidget_lock(f.cap(0))
        If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/hci/chatoptions.cpp:1246", "initialize", "Null form?": Exit Function
        WzChatOptionsForm_changeFreeChatForAll strongSelf, False
    End Select
End Function
