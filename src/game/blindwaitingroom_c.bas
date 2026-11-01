Attribute VB_Name = "blindwaitingroom_c"
' src/titleui/widgets/blindwaitingroom.cpp, blindwaitingroom.h, whole: the "waiting room" a client of a blind
' (simple) lobby sees in place of the players list: the title banner with its info button, the host, the player's
' own row, the status and the count of players ready. Its widget classes' fields are the data classes named as the
' classes. std::chrono::steady_clock is widgSteadyNowMs (a time_point's default, the epoch, is 0). The weak_ptr the
' info button's lambda captures is the banner's address (WIDGET_fromPtr): the button dies with it. Translations
' (_()) are not ported: the English text is used.
' (VB) widget class WzRoomTitleBanner : WIDGET
' (VB) widget class WzBlindRoomHostInfo : WIDGET
' (VB) widget class WzBlindRoomPlayerInfo : WIDGET
' (VB) widget class WzBlindRoomStatusInfo : WIDGET
' (VB) widget class WzBlindPlayersReadyInfo : WIDGET
' (VB) widget class WzPlayerBlindWaitingRoom : WIDGET
Option Explicit

' the lambdas (WzFunc: blindwaitingroom_c_fn)
Public Const FN_roomTitleLabel_calcLayout As Long = 18860    ' WzRoomTitleBanner::initialize's titleLabel
Public Const FN_roomInfoButton_calcLayout As Long = 18861    ' WzRoomTitleBanner::initialize's infoButton
Public Const FN_roomInfoButton_onClick As Long = 18862       ' [weakSelf](W_BUTTON&): cap(0) the banner's address
Public Const FN_waitingRoomInfoClick As Long = 18863         ' WzPlayerBlindWaitingRoom::initialize's InfoClickHandler
Public Const FN_printBlindHelp_task As Long = 18864          ' widgScheduleTask([]() { printBlindModeHelpMessagesToConsole(); })
Public Const FN_blindRoom_calcLayout As Long = 18865         ' WzPlayerBlindWaitingRoom::initialize's setCalcLayout

' ////////////////////////////////////////////////////////////////////////////
' Blind Waiting Room

' WzRoomTitleBanner(): WIDGET()
Public Sub WzRoomTitleBanner_construct(ByVal self As WIDGET)
    Dim d As WzRoomTitleBanner
    WIDGET_construct self
    self.kind = K_WzRoomTitleBanner ' (VB) virtual calls from here on reach WzRoomTitleBanner's overrides
    Set d = WzRoomTitleBanner_(self)
    d.topPadding = 10
    d.bottomPadding = 10
    d.internalHorizontalPadding = 10
End Sub

Public Function WzRoomTitleBanner_make(ByVal titleText As String, ByVal infoClickHandler As WzFunc) As WIDGET
    Dim widget As WIDGET
    ' (class make_shared_enabler)
    Set widget = WIDGET_alloc(K_WzRoomTitleBanner)
    WzRoomTitleBanner_construct widget

    WzRoomTitleBanner_initialize widget, titleText, infoClickHandler
    Set WzRoomTitleBanner_make = widget
End Function

Public Function WzRoomTitleBanner_idealHeight(ByVal self As WIDGET) As Long
    Dim d As WzRoomTitleBanner, titleHeight As Long, infoHeight As Long
    Set d = WzRoomTitleBanner_(self)
    ' the height for one row of text
    If Not d.infoButton Is Nothing Then infoHeight = d.infoButton.idealHeight()
    titleHeight = MAX_(d.titleLabel.idealHeight(), infoHeight)
    WzRoomTitleBanner_idealHeight = d.topPadding + d.bottomPadding + titleHeight
End Function

Private Sub WzRoomTitleBanner_initialize(ByVal self As WIDGET, ByVal titleText As String, ByVal onInfoButtonClick As WzFunc)
    Dim d As WzRoomTitleBanner
    Set d = WzRoomTitleBanner_(self)
    Set d.infoClickHandler = onInfoButtonClick

    ' add the titleLabel
    Set d.titleLabel = W_LABEL_new()
    W_LABEL_setFont d.titleLabel, font_regular, psPalette(WZCOL_TEXT_BRIGHT)
    d.titleLabel.setString titleText
    d.titleLabel.setGeometry 0, 0, W_LABEL_getMaxLineWidth(d.titleLabel), d.titleLabel.idealHeight()
    W_LABEL_setCanTruncate d.titleLabel, True
    d.titleLabel.setTransparentToMouse True
    self.attach d.titleLabel
    d.titleLabel.setCalcLayout WzFunc_new(FN_roomTitleLabel_calcLayout)

    If Not d.infoClickHandler Is Nothing Then
        Set d.infoButton = WzInfoButton_make()
        WzFrontendImageButton_setImageDimensions d.infoButton, 16
        self.attach d.infoButton
        d.infoButton.setCalcLayout WzFunc_new(FN_roomInfoButton_calcLayout)
        W_BUTTON_addOnClickHandler d.infoButton, WzFunc_new(FN_roomInfoButton_onClick, ObjPtr(self))   ' [weakSelf]
    End If
End Sub

Public Sub WzRoomTitleBanner_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x0 As Long, y0 As Long, w As Long, h As Long, highlight As Boolean, boxBorder As Long, boxBackground As Long
    x0 = xOffset + self.x()
    y0 = yOffset + self.y()
    w = self.width()
    h = self.height()
    highlight = False

    ' draw box
    boxBorder = psPalette(WZCOL_MENU_BORDER)
    If highlight Then
        boxBorder = pal_RGBA(255, 255, 255, 255)
    End If
    boxBackground = psPalette(WZCOL_MENU_BORDER)
    pie_BoxFill x0, y0, x0 + w, y0 + h, boxBorder
    pie_BoxFill x0 + 1, y0 + 1, x0 + w - 1, y0 + h - 1, boxBackground
End Sub

Public Sub WzRoomTitleBanner_geometryChanged(ByVal self As WIDGET)
    Dim d As WzRoomTitleBanner
    Set d = WzRoomTitleBanner_(self)
    If d.titleLabel Is Nothing Then Exit Sub    ' (VB) inside WIDGET's constructor
    d.titleLabel.callCalcLayout
    If Not d.infoButton Is Nothing Then
        d.infoButton.callCalcLayout
    End If
End Sub

' MARK: - WzBlindRoomHostInfo

' WzBlindRoomHostInfo(): WIDGET()
Public Sub WzBlindRoomHostInfo_construct(ByVal self As WIDGET)
    Dim d As WzBlindRoomHostInfo
    WIDGET_construct self
    self.kind = K_WzBlindRoomHostInfo ' (VB) virtual calls from here on reach WzBlindRoomHostInfo's overrides
    Set d = WzBlindRoomHostInfo_(self)
    d.topPadding = 4
    d.bottomPadding = 4
    d.internalHorizontalPadding = 10
End Sub

Public Function WzBlindRoomHostInfo_make() As WIDGET
    Dim widget As WIDGET
    ' (class make_shared_enabler)
    Set widget = WIDGET_alloc(K_WzBlindRoomHostInfo)
    WzBlindRoomHostInfo_construct widget
    WzBlindRoomHostInfo_initialize widget
    Set WzBlindRoomHostInfo_make = widget
End Function

Public Function WzBlindRoomHostInfo_idealHeight(ByVal self As WIDGET) As Long
    Dim d As WzBlindRoomHostInfo, titleHeight As Long
    Set d = WzBlindRoomHostInfo_(self)
    ' the height for one row of text
    titleHeight = MAX_(d.hostLabel.idealHeight(), d.hostName.idealHeight())
    WzBlindRoomHostInfo_idealHeight = d.topPadding + d.bottomPadding + titleHeight
End Function

Private Sub WzBlindRoomHostInfo_initialize(ByVal self As WIDGET)
    Dim d As WzBlindRoomHostInfo, hostTipStr As String, identity As EcKey, keyBytes() As Byte, n As Long
    Dim b64Identity As String
    Set d = WzBlindRoomHostInfo_(self)
    Set d.hostLabel = W_LABEL_new()
    W_LABEL_setFont d.hostLabel, font_regular, psPalette(WZCOL_TEXT_DARK)
    d.hostLabel.setString "Host:"
    d.hostLabel.setGeometry 0, 0, W_LABEL_getMaxLineWidth(d.hostLabel), d.hostLabel.idealHeight()
    W_LABEL_setCanTruncate d.hostLabel, False
    d.hostLabel.setTransparentToMouse True
    self.attach d.hostLabel

    Set d.hostName = W_LABEL_new()
    W_LABEL_setFont d.hostName, font_regular, psPalette(WZCOL_TEXT_MEDIUM)
    d.hostName.setString NetPlay_players_name(NetPlay_hostPlayer)
    d.hostName.setGeometry 0, 0, W_LABEL_getMaxLineWidth(d.hostName), d.hostName.idealHeight()
    W_LABEL_setCanTruncate d.hostName, True
    self.attach d.hostName

    If ULT(NetPlay_hostPlayer, MAX_PLAYER_SLOTS) Then
        hostTipStr = "NOTE: Host is also a player."
    Else
        identity = getOutputPlayerIdentity(NetPlay_hostPlayer)
        If Not EcKey_empty(identity) Then
            If Len(hostTipStr) <> 0 Then
                hostTipStr = hostTipStr & vbLf
            End If
            n = EcKey_toBytes(identity, EcKey_Public, keyBytes)
            b64Identity = base64Encode(keyBytes, n)
            hostTipStr = hostTipStr & "Player ID: "
            hostTipStr = hostTipStr & vbLf
            If Len(b64Identity) = 0 Then hostTipStr = hostTipStr & "(none)" Else hostTipStr = hostTipStr & b64Identity
        End If
    End If
    d.hostName.setTip hostTipStr
End Sub

Public Sub WzBlindRoomHostInfo_geometryChanged(ByVal self As WIDGET)
    Dim d As WzBlindRoomHostInfo, w As Long, h As Long, hostNameX0 As Long, hostNameWidth As Long
    Set d = WzBlindRoomHostInfo_(self)
    If d.hostName Is Nothing Then Exit Sub      ' (VB) inside WIDGET's constructor
    w = self.width()
    h = self.height()
    d.hostLabel.setGeometry 0, 0, d.hostLabel.width(), h

    hostNameX0 = d.hostLabel.width() + d.internalHorizontalPadding
    hostNameWidth = w - hostNameX0
    d.hostName.setGeometry hostNameX0, 0, hostNameWidth, h
End Sub

Public Sub WzBlindRoomHostInfo_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    WzBlindRoomHostInfo_(self).hostName.setString NetPlay_players_name(NetPlay_hostPlayer)
End Sub

' MARK: - WzBlindRoomPlayerInfo

' WzBlindRoomPlayerInfo(): WIDGET()
Public Sub WzBlindRoomPlayerInfo_construct(ByVal self As WIDGET)
    Dim d As WzBlindRoomPlayerInfo
    WIDGET_construct self
    self.kind = K_WzBlindRoomPlayerInfo ' (VB) virtual calls from here on reach WzBlindRoomPlayerInfo's overrides
    Set d = WzBlindRoomPlayerInfo_(self)
    d.topPadding = 4
    d.bottomPadding = 4
    d.internalVerticalPadding = 10
End Sub

Public Function WzBlindRoomPlayerInfo_make(ByVal parent As WzMultiplayerOptionsTitleUI) As WIDGET
    Dim widget As WIDGET
    ' (class make_shared_enabler)
    Set widget = WIDGET_alloc(K_WzBlindRoomPlayerInfo)
    WzBlindRoomPlayerInfo_construct widget
    WzBlindRoomPlayerInfo_initialize widget, parent
    Set WzBlindRoomPlayerInfo_make = widget
End Function

Public Function WzBlindRoomPlayerInfo_idealHeight(ByVal self As WIDGET) As Long
    Dim d As WzBlindRoomPlayerInfo, result As Long
    Set d = WzBlindRoomPlayerInfo_(self)
    result = d.topPadding + d.bottomPadding
    result = result + d.playerLabel.idealHeight()
    result = result + d.internalVerticalPadding + MULTIOP_PLAYERHEIGHT
    WzBlindRoomPlayerInfo_idealHeight = result
End Function

Private Sub WzBlindRoomPlayerInfo_initialize(ByVal self As WIDGET, ByVal parent As WzMultiplayerOptionsTitleUI)
    Dim d As WzBlindRoomPlayerInfo, textBottom As Long
    Set d = WzBlindRoomPlayerInfo_(self)
    Set d.playerLabel = W_LABEL_new()
    W_LABEL_setFont d.playerLabel, font_regular, psPalette(WZCOL_TEXT_DARK)
    d.playerLabel.setString "You:"
    d.playerLabel.setGeometry 0, 0, W_LABEL_getMaxLineWidth(d.playerLabel), d.playerLabel.idealHeight()
    W_LABEL_setCanTruncate d.playerLabel, False
    d.playerLabel.setTransparentToMouse True
    self.attach d.playerLabel
    textBottom = d.playerLabel.y() + d.playerLabel.height()

    Set d.playerRow = WzPlayerRow_make(selectedPlayer, parent)
    self.attach d.playerRow
    d.playerRow.setGeometry 0, textBottom + d.internalVerticalPadding, MULTIOP_PLAYERWIDTH, MULTIOP_PLAYERHEIGHT

    WzBlindRoomPlayerInfo_geometryChanged self
End Sub

Public Sub WzBlindRoomPlayerInfo_geometryChanged(ByVal self As WIDGET)
    Dim d As WzBlindRoomPlayerInfo, w As Long, textBottom As Long, playerWidgetWidth As Long
    Set d = WzBlindRoomPlayerInfo_(self)
    If d.playerRow Is Nothing Then Exit Sub     ' (VB) inside WIDGET's constructor, or before initialize made the row
    w = self.width()

    d.playerLabel.setGeometry 0, 0, d.playerLabel.width(), d.playerLabel.idealHeight()
    textBottom = d.playerLabel.y() + d.playerLabel.height()

    playerWidgetWidth = MIN_(MULTIOP_PLAYERWIDTH, w)
    If playerWidgetWidth > 0 Then
        d.playerRow.setGeometry 0, textBottom + d.internalVerticalPadding, playerWidgetWidth, MULTIOP_PLAYERHEIGHT
    End If
End Sub

' (run: a no-op, as WIDGET's)

' MARK: - WzBlindRoomStatusInfo

' WzBlindRoomStatusInfo(): WIDGET()
Public Sub WzBlindRoomStatusInfo_construct(ByVal self As WIDGET)
    Dim d As WzBlindRoomStatusInfo
    WIDGET_construct self
    self.kind = K_WzBlindRoomStatusInfo ' (VB) virtual calls from here on reach WzBlindRoomStatusInfo's overrides
    Set d = WzBlindRoomStatusInfo_(self)
    d.topPadding = 4
    d.bottomPadding = 4
    d.internalVerticalPadding = 5
    d.minUpdateFreq = 500
End Sub

Public Function WzBlindRoomStatusInfo_make() As WIDGET
    Dim widget As WIDGET
    ' (class make_shared_enabler)
    Set widget = WIDGET_alloc(K_WzBlindRoomStatusInfo)
    WzBlindRoomStatusInfo_construct widget
    WzBlindRoomStatusInfo_initialize widget
    Set WzBlindRoomStatusInfo_make = widget
End Function

Public Function WzBlindRoomStatusInfo_idealHeight(ByVal self As WIDGET) As Long
    Dim d As WzBlindRoomStatusInfo, result As Long
    Set d = WzBlindRoomStatusInfo_(self)
    ' the height for one row of text
    result = d.topPadding + d.bottomPadding
    result = result + d.statusLabel.idealHeight()
    result = result + d.internalVerticalPadding
    result = result + d.statusDetails.idealHeight()
    WzBlindRoomStatusInfo_idealHeight = result
End Function

Private Sub WzBlindRoomStatusInfo_initialize(ByVal self As WIDGET)
    Dim d As WzBlindRoomStatusInfo
    Set d = WzBlindRoomStatusInfo_(self)
    Set d.statusLabel = W_LABEL_new()
    W_LABEL_setFont d.statusLabel, font_regular, psPalette(WZCOL_TEXT_DARK)
    d.statusLabel.setString "Status:"
    d.statusLabel.setGeometry 0, 0, W_LABEL_getMaxLineWidth(d.statusLabel), d.statusLabel.idealHeight()
    W_LABEL_setCanTruncate d.statusLabel, False
    d.statusLabel.setTransparentToMouse True
    self.attach d.statusLabel

    Set d.statusDetails = W_LABEL_new()
    W_LABEL_setFont d.statusDetails, font_regular, psPalette(WZCOL_TEXT_MEDIUM)
    d.statusDetails.setString NetPlay_players_name(NetPlay_hostPlayer)
    d.statusDetails.setGeometry 0, 0, W_LABEL_getMaxLineWidth(d.statusDetails), d.statusDetails.idealHeight()
    W_LABEL_setCanTruncate d.statusDetails, True
    self.attach d.statusDetails

    Set d.indeterminateIndicator = createJoiningIndeterminateProgressWidget(font_regular_bold)
    self.attach d.indeterminateIndicator
End Sub

Public Sub WzBlindRoomStatusInfo_geometryChanged(ByVal self As WIDGET)
    Dim d As WzBlindRoomStatusInfo, w As Long, textBottom As Long, statusDetailsWidth As Long
    Set d = WzBlindRoomStatusInfo_(self)
    If d.indeterminateIndicator Is Nothing Then Exit Sub    ' (VB) inside WIDGET's constructor, or before initialize
    w = self.width()
    d.statusLabel.setGeometry 0, 0, d.statusLabel.width(), d.statusLabel.idealHeight()
    textBottom = d.statusLabel.y() + d.statusLabel.height()

    statusDetailsWidth = w - d.indeterminateIndicator.idealWidth() - 5
    d.statusDetails.setGeometry 0, textBottom + d.internalVerticalPadding, statusDetailsWidth, d.statusDetails.idealHeight()

    d.indeterminateIndicator.setGeometry statusDetailsWidth + 5, textBottom + d.internalVerticalPadding, d.indeterminateIndicator.idealWidth(), d.indeterminateIndicator.idealHeight()
End Sub

Public Sub WzBlindRoomStatusInfo_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    Dim d As WzBlindRoomStatusInfo, now As Double, numNotReadyPlayers As Long, numReadyPlayers As Long
    Dim numNotReadySpectators As Long, hostIsReady As Boolean, selfIsReady As Boolean, playersAllOnSameTeam As Boolean
    Dim player As Long
    Set d = WzBlindRoomStatusInfo_(self)
    now = widgSteadyNowMs()
    If Int(now - d.lastStatusUpdate) < d.minUpdateFreq Then   ' duration_cast<milliseconds>(now - lastStatusUpdate) < minUpdateFreq
        Exit Sub
    End If

    numNotReadyPlayers = 0
    numReadyPlayers = 0
    numNotReadySpectators = 0
    hostIsReady = False
    selfIsReady = NetPlay_players_ready(realSelectedPlayer)
    playersAllOnSameTeam = (allPlayersOnSameTeam(-1) <> -1)

    For player = 0 To NetPlay_players_size() - 1
        ' check if this human player is ready, ignore AIs
        If NetPlay_players_allocated(player) Then
            If NetPlay_players_ready(player) Then
                If player < game.maxPlayers Then
                    numReadyPlayers = numReadyPlayers + 1
                End If
                If player = NetPlay_hostPlayer Then
                    hostIsReady = True
                End If
            Else
                If player < game.maxPlayers Then
                    numNotReadyPlayers = numNotReadyPlayers + 1
                Else
                    numNotReadySpectators = numNotReadySpectators + 1
                End If
            End If
        ElseIf NetPlay_players_ai(player) >= 0 And player < game.maxPlayers Then
            numReadyPlayers = numReadyPlayers + 1
        End If
    Next

    ' Determine status string
    If playersAllOnSameTeam Then
        If numNotReadyPlayers + numReadyPlayers > 1 Then
            d.statusStr = "Waiting for host to configure teams"
        Else
            d.statusStr = "Waiting for players"
        End If
    ElseIf Not selfIsReady Then
        If NETgetDownloadProgress(realSelectedPlayer) <> 100 Then
            d.statusStr = "Downloading map / mods"
        Else
            d.statusStr = "Waiting for you to check ready"
        End If
    ElseIf numReadyPlayers > 0 And numNotReadyPlayers > 0 Then
        d.statusStr = "Waiting for players"
    ElseIf Not hostIsReady And NetPlay_players_isSpectator(NetPlay_hostPlayer) Then
        d.statusStr = "Waiting for host to start game"
    ElseIf numNotReadySpectators > 0 Then
        d.statusStr = "Waiting for spectators"
    Else
        d.statusStr = "Waiting for players"
    End If

    d.statusDetails.setString d.statusStr

    d.lastStatusUpdate = now
End Sub

' MARK: - WzBlindPlayersReadyInfo

' WzBlindPlayersReadyInfo(): WIDGET()
Public Sub WzBlindPlayersReadyInfo_construct(ByVal self As WIDGET)
    Dim d As WzBlindPlayersReadyInfo
    WIDGET_construct self
    self.kind = K_WzBlindPlayersReadyInfo ' (VB) virtual calls from here on reach WzBlindPlayersReadyInfo's overrides
    Set d = WzBlindPlayersReadyInfo_(self)
    d.topPadding = 4
    d.bottomPadding = 4
    d.internalHorizontalPadding = 10
    d.minUpdateFreq = 500
End Sub

Public Function WzBlindPlayersReadyInfo_make() As WIDGET
    Dim widget As WIDGET
    ' (class make_shared_enabler)
    Set widget = WIDGET_alloc(K_WzBlindPlayersReadyInfo)
    WzBlindPlayersReadyInfo_construct widget
    WzBlindPlayersReadyInfo_initialize widget
    Set WzBlindPlayersReadyInfo_make = widget
End Function

Public Function WzBlindPlayersReadyInfo_idealHeight(ByVal self As WIDGET) As Long
    Dim d As WzBlindPlayersReadyInfo, titleHeight As Long
    Set d = WzBlindPlayersReadyInfo_(self)
    ' the height for one row of text
    titleHeight = MAX_(d.readyCountLabel.idealHeight(), d.readyCountValue.idealHeight())
    WzBlindPlayersReadyInfo_idealHeight = d.topPadding + d.bottomPadding + titleHeight
End Function

Private Sub WzBlindPlayersReadyInfo_initialize(ByVal self As WIDGET)
    Dim d As WzBlindPlayersReadyInfo
    Set d = WzBlindPlayersReadyInfo_(self)
    Set d.readyCountLabel = W_LABEL_new()
    W_LABEL_setFont d.readyCountLabel, font_regular, psPalette(WZCOL_TEXT_DARK)
    d.readyCountLabel.setString "Players ready & waiting for match:"
    d.readyCountLabel.setGeometry 0, 0, W_LABEL_getMaxLineWidth(d.readyCountLabel), d.readyCountLabel.idealHeight()
    W_LABEL_setCanTruncate d.readyCountLabel, False
    d.readyCountLabel.setTransparentToMouse True
    self.attach d.readyCountLabel

    Set d.readyCountValue = W_LABEL_new()
    W_LABEL_setFont d.readyCountValue, font_regular, psPalette(WZCOL_TEXT_DARK)
    d.readyCountValue.setString "0"
    d.readyCountValue.setGeometry 0, 0, W_LABEL_getMaxLineWidth(d.readyCountValue), d.readyCountValue.idealHeight()
    W_LABEL_setCanTruncate d.readyCountValue, True
    d.readyCountValue.setTransparentToMouse True
    self.attach d.readyCountValue
End Sub

Public Sub WzBlindPlayersReadyInfo_geometryChanged(ByVal self As WIDGET)
    Dim d As WzBlindPlayersReadyInfo, w As Long, h As Long, readyCountValueWidth As Long, readyCountValueX0 As Long
    Dim labelWidth As Long
    Set d = WzBlindPlayersReadyInfo_(self)
    If d.readyCountValue Is Nothing Then Exit Sub   ' (VB) inside WIDGET's constructor, or before initialize
    w = self.width()
    h = self.height()

    ' ensure the value is fully visible
    readyCountValueWidth = W_LABEL_getMaxLineWidth(d.readyCountValue)
    readyCountValueX0 = w - readyCountValueWidth
    d.readyCountValue.setGeometry readyCountValueX0, 0, readyCountValueWidth, h

    ' then use the rest of the space for the label
    labelWidth = MAX_(w - readyCountValueWidth - d.internalHorizontalPadding, 0)
    d.readyCountLabel.setGeometry 0, 0, labelWidth, h
End Sub

Public Sub WzBlindPlayersReadyInfo_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    Dim d As WzBlindPlayersReadyInfo, now As Double, numNotReadyPlayers As Long, numReadyPlayers As Long
    Dim numReadyAIBots As Long, numNotReadySpectators As Long, numReadySpectators As Long, hostIsReady As Boolean
    Dim player As Long, readyStr As String
    Set d = WzBlindPlayersReadyInfo_(self)
    now = widgSteadyNowMs()
    If Int(now - d.lastStatusUpdate) < d.minUpdateFreq Then   ' duration_cast<milliseconds>(now - lastStatusUpdate) < minUpdateFreq
        Exit Sub
    End If

    For player = 0 To NetPlay_players_size() - 1
        ' check if this human player is ready, ignore AIs
        If NetPlay_players_allocated(player) Then
            If NetPlay_players_ready(player) Then
                If player < game.maxPlayers Then
                    numReadyPlayers = numReadyPlayers + 1
                Else
                    numReadySpectators = numReadySpectators + 1
                End If
                If player = NetPlay_hostPlayer Then
                    hostIsReady = True
                End If
            Else
                If player < game.maxPlayers Then
                    numNotReadyPlayers = numNotReadyPlayers + 1
                Else
                    numNotReadySpectators = numNotReadySpectators + 1
                End If
            End If
        ElseIf NetPlay_players_ai(player) >= 0 And player < game.maxPlayers Then
            numReadyAIBots = numReadyAIBots + 1
        End If
    Next

    d.readyCountValue.setString CStr(numReadyPlayers)

    d.tooltip = va("Players: (%u ready, %u not ready)", numReadyPlayers, numNotReadyPlayers)
    If numReadyAIBots > 0 Then
        d.tooltip = d.tooltip & vbLf
        d.tooltip = d.tooltip & va("AI Bots: %u", numReadyAIBots)
    End If
    If numReadySpectators > 0 Or numNotReadySpectators > 0 Then
        d.tooltip = d.tooltip & vbLf
        d.tooltip = d.tooltip & va("Spectators: (%u ready, %u not ready)", numReadySpectators, numNotReadySpectators)
    End If
    d.tooltip = d.tooltip & vbLf
    If hostIsReady Then readyStr = "ready" Else readyStr = "not ready"
    d.tooltip = d.tooltip & va("Host: %s", readyStr)
    ' (C++ never sets lastStatusUpdate here: the counts are taken every run)
End Sub

Public Function WzBlindPlayersReadyInfo_getTip(ByVal self As WIDGET) As String
    WzBlindPlayersReadyInfo_getTip = WzBlindPlayersReadyInfo_(self).tooltip
End Function

' MARK: - WzPlayerBlindWaitingRoom

' WzPlayerBlindWaitingRoom(const std::shared_ptr<WzMultiplayerOptionsTitleUI>& titleUI): WIDGET(), weakTitleUI(titleUI)
Public Sub WzPlayerBlindWaitingRoom_construct(ByVal self As WIDGET)
    Dim d As WzPlayerBlindWaitingRoom
    WIDGET_construct self
    self.kind = K_WzPlayerBlindWaitingRoom ' (VB) virtual calls from here on reach WzPlayerBlindWaitingRoom's overrides
    Set d = WzPlayerBlindWaitingRoom_(self)
    d.titleBannerBottomPadding = 10
    d.outerWidgetPaddingX = 15
    d.betweenWidgetPaddingY = 10
    d.bottomPaddingY = 15
End Sub

Public Function WzPlayerBlindWaitingRoom_make(ByVal titleUI As WzMultiplayerOptionsTitleUI) As WIDGET
    Dim widget As WIDGET
    ' (class make_shared_enabler)
    Set widget = WIDGET_alloc(K_WzPlayerBlindWaitingRoom)
    WzPlayerBlindWaitingRoom_construct widget
    WzPlayerBlindWaitingRoom_initialize widget, titleUI
    Set WzPlayerBlindWaitingRoom_make = widget
End Function

Private Sub WzPlayerBlindWaitingRoom_initialize(ByVal self As WIDGET, ByVal titleUI As WzMultiplayerOptionsTitleUI)
    Dim d As WzPlayerBlindWaitingRoom
    Set d = WzPlayerBlindWaitingRoom_(self)
    ' Add the "Waiting Room" info bar at top
    Set d.titleBanner = WzRoomTitleBanner_make("Waiting Room", WzFunc_new(FN_waitingRoomInfoClick))
    self.attach d.titleBanner

    ' Add host details
    Set d.hostInfo = WzBlindRoomHostInfo_make()
    self.attach d.hostInfo

    ' Add the current player details
    Set d.playerInfo = WzBlindRoomPlayerInfo_make(titleUI)
    self.attach d.playerInfo

    ' Add Status field)
    Set d.statusInfo = WzBlindRoomStatusInfo_make()
    self.attach d.statusInfo

    ' Add "People ready & waiting for match: <number>"
    Set d.playersReadyInfo = WzBlindPlayersReadyInfo_make()
    self.attach d.playersReadyInfo

    self.setCalcLayout WzFunc_new(FN_blindRoom_calcLayout)
End Sub

Public Sub WzPlayerBlindWaitingRoom_geometryChanged(ByVal self As WIDGET)
    WzPlayerBlindWaitingRoom_recalculateInternalLayout self
End Sub

Private Function WzPlayerBlindWaitingRoom_getPlayerRowY0(ByVal self As WIDGET) As Long
    Dim d As WzPlayerBlindWaitingRoom
    Set d = WzPlayerBlindWaitingRoom_(self)
    If Not (Not d.playerInfo Is Nothing) Then ASSERT_FAILURE "playerInfo != nullptr", "src/titleui/widgets/blindwaitingroom.cpp:725", "getPlayerRowY0", "Invalid widget?": WzPlayerBlindWaitingRoom_getPlayerRowY0 = 0: Exit Function
    WzPlayerBlindWaitingRoom_getPlayerRowY0 = d.playerInfo.y()
End Function

Public Sub WzPlayerBlindWaitingRoom_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x0 As Long, y0 As Long, w As Long, h As Long
    x0 = xOffset + self.x()
    y0 = yOffset + self.y()
    w = self.width()
    h = self.height()

    pie_BoxFill x0, y0, x0 + w, y0 + h, psPalette(WZCOL_MENU_BACKGROUND)
    iV_Box x0, y0, x0 + w, y0 + h, psPalette(WZCOL_MENU_BORDER)
End Sub

Private Sub WzPlayerBlindWaitingRoom_recalculateInternalLayout(ByVal self As WIDGET)
    Dim d As WzPlayerBlindWaitingRoom, w As Long, h As Long, nextWidgetY0 As Long, widgetX0 As Long
    Dim usableWidgetWidth As Long, bottomWidgetY0 As Long
    Set d = WzPlayerBlindWaitingRoom_(self)
    If d.playersReadyInfo Is Nothing Then Exit Sub  ' (VB) inside WIDGET's constructor, or before initialize
    w = self.width()
    h = self.height()

    d.titleBanner.setGeometry 0, 0, w, d.titleBanner.idealHeight()
    nextWidgetY0 = d.titleBanner.y() + d.titleBanner.height() + d.titleBannerBottomPadding

    widgetX0 = d.outerWidgetPaddingX
    usableWidgetWidth = w - (d.outerWidgetPaddingX * 2)
    d.hostInfo.setGeometry widgetX0, nextWidgetY0, usableWidgetWidth, d.hostInfo.idealHeight()
    nextWidgetY0 = d.hostInfo.y() + d.hostInfo.height() + d.betweenWidgetPaddingY

    d.playerInfo.setGeometry widgetX0, nextWidgetY0, usableWidgetWidth, d.playerInfo.idealHeight()
    nextWidgetY0 = d.playerInfo.y() + d.playerInfo.height() + d.betweenWidgetPaddingY

    d.statusInfo.setGeometry widgetX0, nextWidgetY0, usableWidgetWidth, d.statusInfo.idealHeight()
    nextWidgetY0 = d.statusInfo.y() + d.statusInfo.height() + d.betweenWidgetPaddingY

    ' aligned bottom-up
    bottomWidgetY0 = h - d.playersReadyInfo.idealHeight() - d.bottomPaddingY
    d.playersReadyInfo.setGeometry widgetX0, bottomWidgetY0, usableWidgetWidth, d.playersReadyInfo.idealHeight()
End Sub

' MARK: - Public methods

Public Function makeWzPlayerBlindWaitingRoom(ByVal titleUI As WzMultiplayerOptionsTitleUI) As WIDGET
    Set makeWzPlayerBlindWaitingRoom = WzPlayerBlindWaitingRoom_make(titleUI)
End Function

Public Function getWzPlayerBlindWaitingRoomPlayerRowY0(ByVal blindWaitingRoomWidget As WIDGET) As Long
    Dim isRoom As Boolean
    If Not blindWaitingRoomWidget Is Nothing Then isRoom = WIDGET_isKindOf(blindWaitingRoomWidget, K_WzPlayerBlindWaitingRoom)
    If Not isRoom Then ASSERT_FAILURE "pBlindWaitingRoomWidget != nullptr", "src/titleui/widgets/blindwaitingroom.cpp:795", "getWzPlayerBlindWaitingRoomPlayerRowY0", "Invalid widget": getWzPlayerBlindWaitingRoomPlayerRowY0 = 0: Exit Function
    getWzPlayerBlindWaitingRoomPlayerRowY0 = WzPlayerBlindWaitingRoom_getPlayerRowY0(blindWaitingRoomWidget)
End Function

' the lambdas
Public Function blindwaitingroom_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim psWidget As WIDGET, psParent As WIDGET, pd As WzRoomTitleBanner, x0 As Long, w As Long, h As Long
    Dim strongSelf As WIDGET
    Select Case f.id
    Case FN_roomTitleLabel_calcLayout               ' LAMBDA_CALCLAYOUT_SIMPLE (WIDGET *psWidget)
        Set psWidget = args(0)
        Set psParent = psWidget.parent()
        If Not psParent Is Nothing Then
            If Not WIDGET_isKindOf(psParent, K_WzRoomTitleBanner) Then Set psParent = Nothing
        End If
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/titleui/widgets/blindwaitingroom.cpp:102", "initialize", "No parent?": Exit Function
        Set pd = WzRoomTitleBanner_(psParent)
        x0 = pd.internalHorizontalPadding
        w = psParent.width() - (pd.internalHorizontalPadding * 3) - 16
        h = psParent.height() - (pd.topPadding + pd.bottomPadding)
        psWidget.setGeometry x0, pd.topPadding, w, h

    Case FN_roomInfoButton_calcLayout               ' LAMBDA_CALCLAYOUT_SIMPLE (WIDGET *psWidget)
        Set psWidget = args(0)
        Set psParent = psWidget.parent()
        If Not psParent Is Nothing Then
            If Not WIDGET_isKindOf(psParent, K_WzRoomTitleBanner) Then Set psParent = Nothing
        End If
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/titleui/widgets/blindwaitingroom.cpp:116", "initialize", "No parent?": Exit Function
        Set pd = WzRoomTitleBanner_(psParent)
        w = 16 + pd.internalHorizontalPadding
        x0 = psParent.width() - w
        h = psParent.height() - (pd.topPadding + pd.bottomPadding)
        psWidget.setGeometry x0, pd.topPadding, w, h

    Case FN_roomInfoButton_onClick                  ' [weakSelf](W_BUTTON&)
        Set strongSelf = WIDGET_fromPtr(f.cap(0))   ' weakSelf.lock()
        If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/titleui/widgets/blindwaitingroom.cpp:125", "initialize", "No parent?": Exit Function
        If Not WzRoomTitleBanner_(strongSelf).infoClickHandler Is Nothing Then
            WzFunc_call WzRoomTitleBanner_(strongSelf).infoClickHandler
        End If

    Case FN_waitingRoomInfoClick                    ' []()
        widgScheduleTask WzFunc_new(FN_printBlindHelp_task)

    Case FN_printBlindHelp_task                     ' []()
        printBlindModeHelpMessagesToConsole

    Case FN_blindRoom_calcLayout                    ' LAMBDA_CALCLAYOUT_SIMPLE (WIDGET *psWidget)
        Set psWidget = args(0)
        If Not WIDGET_isKindOf(psWidget, K_WzPlayerBlindWaitingRoom) Then ASSERT_FAILURE "psPlayerBlindWaitingRoom.get() != nullptr", "src/titleui/widgets/blindwaitingroom.cpp:713", "initialize", "Wrong type of psWidget??": Exit Function
        WzPlayerBlindWaitingRoom_recalculateInternalLayout psWidget
    End Select
End Function
