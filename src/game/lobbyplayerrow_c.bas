Attribute VB_Name = "lobbyplayerrow_c"
' src/titleui/widgets/lobbyplayerrow.cpp, lobbyplayerrow.h: the lobby's player
' row (WzPlayerRow: team, colour, player info and ready / difficulty boxes)
' and its display functions. The widget object model is docs/CONVENTIONS.md's
' "Widgets": the classes' fields are their data classes (same names), the
' lambdas and display functions WzFuncs run by lobbyplayerrow_c_fn; a
' std::weak_ptr is an address (ObjPtr). struct DisplayPlayerCache is the
' class DisplayPlayerCache.
' (VB) widget class ReadyBoxContainerWidget : W_FORM
' (VB) widget class WzPlayerRow : WIDGET
Option Explicit

Private Const MAX_CONNECTED_PLAYERS As Long = MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

' the lambdas and display functions
Public Const FN_displayColour As Long = 6200
Public Const FN_displayTeamChooser As Long = 6201
Public Const FN_displayPlayer As Long = 6202
Public Const FN_PR_TEAM_CLICK As Long = 6203            ' teamButton->addOnClickHandler([playerIdx, titleUI](W_BUTTON& button) {...})
Public Const FN_PR_TEAM_TASK As Long = 6204             ' widgScheduleTask([strongTitleUI, playerIdx] {...})
Public Const FN_PR_COLOUR_CLICK As Long = 6205
Public Const FN_PR_COLOUR_TASK As Long = 6206
Public Const FN_PR_PLAYERINFO_ONDELETE As Long = 6207   ' playerInfo->setOnDelete([](WIDGET *psWidget) {...})
Public Const FN_PR_PLAYERINFO_CALCLAYOUT As Long = 6208
Public Const FN_PR_PLAYERINFO_CLICK As Long = 6209
Public Const FN_PR_POSITION_TASK As Long = 6210         ' widgScheduleTask([strongTitleUI, player] { openPositionChooser })
Public Const FN_PR_DISTRIBUTE_TASK As Long = 6211       ' widgScheduleTask([strongTitleUI] { updatePlayers; resetReadyStatus })
Public Const FN_PR_AI_TASK As Long = 6212               ' widgScheduleTask([strongTitleUI, player] { openAiChooser })
Public Const FN_PR_READYCONTAINER_CALCLAYOUT As Long = 6213
Public Const FN_PR_DIFFICULTY_CLICK As Long = 6214      ' difficultyChooserButton->addOnClickHandler([player, weakTitleUi](W_BUTTON&) {...})
Public Const FN_PR_DIFFICULTY_TASK As Long = 6215
Public Const FN_PR_READY_CLICK As Long = 6216           ' readyButton->addOnClickHandler([player, weakTitleUi](W_BUTTON& button) {...})
Public Const FN_PR_READY_TASK As Long = 6217            ' widgScheduleTask([weakTitleUi, player, pButton] {...})

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

Private Sub drawBlueBox_Spectator(ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long)
    Dim backgroundClr As Long
    ' Match drawBlueBox behavior
    x = x - 1
    y = y - 1
    w = w + 2
    h = h + 2
    backgroundClr = psPalette(WZCOL_FORM_DARK)
    backgroundClr = pielightWithA(backgroundClr, 200)
    pie_BoxFill x, y, x + w, y + h, psPalette(WZCOL_MENU_BORDER)
    pie_UniTransBoxFill x + 1, y + 1, x + w - 1, y + h - 1, backgroundClr
End Sub

Private Sub drawBlueBox_SpectatorOnly(ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long)
    Dim backgroundClr As Long
    ' Match drawBlueBox behavior
    x = x - 1
    y = y - 1
    w = w + 2
    h = h + 2
    backgroundClr = psPalette(WZCOL_FORM_DARK)
    backgroundClr = pielightWithA(backgroundClr, 175)
'   PIELIGHT borderClr = pal_RGBA(30, 30, 30, 120);
'   pie_UniTransBoxFill(x, y, x + w, y + h, borderClr);
'   pie_UniTransBoxFill(x + 1, y + 1, x + w - 1, y + h - 1, backgroundClr);
    pie_UniTransBoxFill x, y, x + w, y + h, backgroundClr
End Sub

Private Sub drawBoxForPlayerInfoSegment(ByVal playerIdx As Long, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long)
    If Not (ULT(playerIdx, NetPlay_players_size())) Then ASSERT_FAILURE "playerIdx < NetPlay.players.size()", "src/titleui/widgets/lobbyplayerrow.cpp:89", "drawBoxForPlayerInfoSegment", "Invalid playerIdx: %zu", playerIdx
    If Not ULT(playerIdx, NetPlay_players_size()) Then
        drawBlueBox x, y, w, h
    ElseIf Not NetPlay_players_isSpectator(playerIdx) Then
        drawBlueBox x, y, w, h
    Else
        If Not isSpectatorOnlySlot(playerIdx) Then
            drawBlueBox_Spectator x, y, w, h
        Else
            drawBlueBox_SpectatorOnly x, y, w, h
        End If
    End If
End Sub

' MARK: - ReadyBoxContainerWidget

Public Sub ReadyBoxContainerWidget_construct(ByVal self As WIDGET)
    W_FORM_construct self
    self.kind = K_ReadyBoxContainerWidget ' (VB) virtual calls from here on reach ReadyBoxContainerWidget's overrides
    ReadyBoxContainerWidget_(self).showAnimatedHighlight = False
End Sub

Public Function ReadyBoxContainerWidget_new() As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_ReadyBoxContainerWidget)
    ReadyBoxContainerWidget_construct w
    Set ReadyBoxContainerWidget_new = w
End Function

Public Sub ReadyBoxContainerWidget_setShowAnimatedHighlight(ByVal self As WIDGET, ByVal val_ As Boolean)
    ReadyBoxContainerWidget_(self).showAnimatedHighlight = val_
End Sub

Public Sub ReadyBoxContainerWidget_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Const scale_ As Long = 2500
    Dim x0 As Long, y0 As Long, w As Long, h As Long, playerIdx As Long, f As Long, mix As Long
    x0 = xOffset + self.x()
    y0 = yOffset + self.y()
    w = self.width()
    h = self.height()
    playerIdx = self.UserData

    drawBoxForPlayerInfoSegment playerIdx, x0, y0, w, h

    If ReadyBoxContainerWidget_(self).showAnimatedHighlight Then
        f = realTime Mod scale_
        mix = pal_RGBA(3, 15, 252, 255)
        mix = pielightWithA(mix, 155 + iSinR(65536 * f \ scale_, 75))

        pie_UniTransBoxFill x0 + 2, y0 + 1, x0 + w - 1, y0 + h - 1, mix
    End If
End Sub

' MARK: - display functions

Private Function factionIcon(ByVal faction_ As Long) As Long
    Select Case faction_
    Case 0: factionIcon = IMAGE_FACTION_NORMAL
    Case 1: factionIcon = IMAGE_FACTION_NEXUS
    Case 2: factionIcon = IMAGE_FACTION_COLLECTIVE
    Case Else: factionIcon = IMAGE_NO   ' / what??
    End Select
End Function

Private Sub displayColour(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x As Long, y As Long, j As Long, player As Long, faction_ As Long
    x = xOffset + psWidget.x()
    y = yOffset + psWidget.y()
    j = psWidget.UserData

    drawBoxForPlayerInfoSegment j, x, y, psWidget.width(), psWidget.height()
    If Not NetPlay_players_fileSendInProgress(j) And (isHumanPlayer(j) Or NetPlay_players_difficulty(j) <> AIDifficulty_DISABLED) And Not NetPlay_players_isSpectator(j) Then
        player = getPlayerColour(j)
        ' (STATIC_ASSERT(MAX_PLAYERS <= 16))
        iV_DrawImageTc_ids FrontImages, IMAGE_PLAYERN, IMAGE_PLAYERN_TC, x + 3, y + 9, pal_GetTeamColour(player)
        faction_ = NetPlay_players_faction(j)
        iV_DrawImageFileAnisotropic FrontImages, factionIcon(faction_), x, y, 11, 9
    End If
End Sub

Public Sub displayTeamChooser(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x As Long, y As Long, i As Long
    x = xOffset + psWidget.x()
    y = yOffset + psWidget.y()
    i = psWidget.UserData

    If Not (ULT(i, MAX_CONNECTED_PLAYERS)) Then ASSERT_FAILURE "i < MAX_CONNECTED_PLAYERS", "src/titleui/widgets/lobbyplayerrow.cpp:172", "displayTeamChooser", "Index (%u) out of bounds", i: Exit Sub

    drawBoxForPlayerInfoSegment i, x, y, psWidget.width(), psWidget.height()

    If Not NetPlay_players_isSpectator(i) Then
        If Not isBlindSimpleLobby(game.blindMode) Or NetPlay_isHost Then
            If alliancesSetTeamsBeforeGame(game.alliance) Then
                If Not (NetPlay_players_team(i) >= 0 And NetPlay_players_team(i) < MAX_PLAYERS) Then ASSERT_FAILURE "NetPlay.players[i].team >= 0 && NetPlay.players[i].team < MAX_PLAYERS", "src/titleui/widgets/lobbyplayerrow.cpp:182", "displayTeamChooser", "Team index out of bounds": Exit Sub
                iV_DrawImage FrontImages, IMAGE_TEAM0 + NetPlay_players_team(i), x + 1, y + 8
            Else
                ' TODO: Maybe display something else here to signify "no team, FFA"
            End If
        End If
    Else
        iV_DrawImage FrontImages, IMAGE_SPECTATOR, x + 1, y + 8
    End If
End Sub

' void displaySpectatorAddButton(WIDGET *psWidget, UDWORD xOffset, UDWORD yOffset): (VB) not used, and named
' for its file: multiint.cpp draws its add-spectator button with its own static displaySpectatorAddButton
' (multiint_lobby_c's, the same body)
Private Sub displaySpectatorAddButton_lobbyplayerrow(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x As Long, y As Long
    x = xOffset + psWidget.x()
    y = yOffset + psWidget.y()

    drawBlueBox_SpectatorOnly x, y, psWidget.width(), psWidget.height()
    iV_DrawImage FrontImages, IMAGE_SPECTATOR, x + 2, y + 1
End Sub

' static bool isKnownPlayer(std::map<std::string, EcKey::Key> const &knownPlayers, std::string const &name, EcKey const &key)
' (knownPlayers: a Scripting.Dictionary of name -> the key's bytes; Nothing is the empty map)
Private Function isKnownPlayer(ByVal knownPlayers As Object, ByVal name_ As String, key As EcKey) As Boolean
    Dim keyBytes() As Byte, n As Long, known As Variant, i As Long
    If EcKey_empty(key) Then
        isKnownPlayer = False
        Exit Function
    End If
    If knownPlayers Is Nothing Then Exit Function
    If Not knownPlayers.Exists(name_) Then Exit Function
    known = knownPlayers.Item(name_)
    n = EcKey_toBytes_public(key, keyBytes)
    If n <> UBound(known) - LBound(known) + 1 Then Exit Function
    For i = 0 To n - 1
        If keyBytes(i) <> known(LBound(known) + i) Then Exit Function
    Next
    isKnownPlayer = True
End Function

Private Sub displayAltNameBox(ByVal x As Long, ByVal y As Long, ByVal psWidget As WIDGET, ByVal cache As DisplayPlayerCache, ByVal isHighlight As Boolean)
    Dim altNameBoxWidth As Long, altNameBoxHeight As Long, altNameBoxX0 As Long, altNameBoxColor As Long, altNameTextY0 As Long
    Dim altNameTextColor As Long, factor As Single
    altNameBoxWidth = cache.wzAltNameText.width() + 4
    altNameBoxHeight = cache.wzAltNameText.lineSize() + 2
    altNameBoxX0 = (x + psWidget.width()) - altNameBoxWidth
    altNameBoxColor = psPalette(WZCOL_MENU_BORDER)
    If isHighlight Then factor = 0.3! Else factor = 0.75!
    altNameBoxColor = pielightWithA(altNameBoxColor, Fix(CSng(pielightA(altNameBoxColor)) * factor))
    pie_UniTransBoxFill altNameBoxX0, y, altNameBoxX0 + altNameBoxWidth, y + altNameBoxHeight, altNameBoxColor

    altNameTextY0 = y + (altNameBoxHeight - cache.wzAltNameText.lineSize()) \ 2 - cache.wzAltNameText.aboveBase()
    altNameTextColor = psPalette(WZCOL_TEXT_MEDIUM)
    If isHighlight Then
        altNameTextColor = pielightWithA(altNameTextColor, Fix(CSng(pielightA(altNameTextColor)) * 0.3!))
    End If
    cache.wzAltNameText.render altNameBoxX0 + 2, altNameTextY0, altNameTextColor
End Sub

Public Sub displayPlayer(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Const nameX As Long = 32
    Dim cache As DisplayPlayerCache, aidata_size As Long, x As Long, y As Long, j As Long, isHighlight As Boolean
    Dim downloadProgress As Long, progressString As String, stat As PLAYERSTATS_, name_ As String
    Dim hasAltName As Boolean, pAltName As String, serverPlayers As Object, colour As Long, nameFont As Long
    Dim subText As String, buf As String, pingQualifierStr As String, dummy As Boolean, star(0 To 2) As Long, medal As Long
    Dim altName As String, maxAltNameWidth As Long, fontID As Long, H As Long, textY As Long
    Dim starImgs(0 To 3) As Long, medalImgs(0 To 3) As Long, aitext As String, textColor As Long

    ' Any widget using displayPlayer must have its pUserData initialized to a (DisplayPlayerCache*)
    ' (assert(psWidget->pUserData != nullptr): the C assert, off in a release build)
    Set cache = psWidget.pUserData

    aidata_size = getAIData_size()  ' const auto& aidata = getAIData();

    x = xOffset + psWidget.x()
    y = yOffset + psWidget.y()
    j = psWidget.UserData
    isHighlight = (psWidget.getState() And WBUT_HIGHLIGHT) <> 0

    downloadProgress = NETgetDownloadProgress(j)

    drawBoxForPlayerInfoSegment j, x, y, psWidget.width(), psWidget.height()
    If downloadProgress <> 100 Then
        ' (char progressString[MAX_STR_LENGTH])
        If j <> selectedPlayer Then
            progressString = "Sending Map: " & CStr(downloadProgress) & "% "
        Else
            progressString = "Map: " & CStr(downloadProgress) & "% downloaded"
        End If
        cache.wzMainText.setText progressString, font_regular
        cache.wzMainText.render x + 5, y + 22, psPalette(WZCOL_FORM_TEXT)
        cache.fullMainText = progressString
        Exit Sub
    ElseIf ingame.localOptionsReceived And NetPlay_players_allocated(j) Then    ' only draw if real player!
        stat = getMultiStats(j)

        name_ = getPlayerName(j)
        If game.blindMode <> BLIND_MODE_NONE Then
            ' Special override: Show our own name if in spectators
            If j = selectedPlayer And j > MAX_PLAYER_SLOTS Then
                name_ = NetPlay_players_name(j)
            End If
        End If
        hasAltName = False      ' const char* pAltName = nullptr;
        If NetPlay_players_isAdmin(j) Then
            hasAltName = True
            pAltName = "Admin"
        End If

        Set serverPlayers = Nothing     ' TODO Fill this with players known to the server (needs implementing on the server, too). Currently useless.

        nameFont = font_regular
        If j = selectedPlayer Then
            colour = psPalette(WZCOL_TEXT_BRIGHT)
'           nameFont = font_regular_bold;
        ElseIf ingame.PingTimes(j) >= PING_LIMIT Then
            colour = psPalette(WZCOL_FORM_PLAYER_NOPING)
        ElseIf isKnownPlayer(serverPlayers, name_, stat.identity) Or game.blindMode <> BLIND_MODE_NONE Then
            colour = psPalette(WZCOL_FORM_PLAYER_KNOWN_BY_SERVER)
        ElseIf isLocallyKnownPlayer(name_, stat.identity) Then
            colour = psPalette(WZCOL_FORM_PLAYER_KNOWN)
        Else
            colour = psPalette(WZCOL_FORM_PLAYER_UNKNOWN)
        End If

        ' name
        If cache.fullMainText <> name_ Then
            If iV_GetTextWidth(name_, font_regular) > psWidget.width() - nameX Then
                Do While Len(name_) <> 0 And iV_GetTextWidth(name_ & "...", font_regular) > psWidget.width() - nameX
                    name_ = Left$(name_, Len(name_) - 1)  ' Clip name.
                Loop
                name_ = name_ & "..."
            End If
            cache.wzMainText.setText name_, nameFont
            cache.fullMainText = name_
        End If

        subText = ""

        If j = NetPlay_hostPlayer And NetPlay_bComms Then
            subText = subText & "HOST"
        ElseIf (j < MAX_PLAYERS) And (game.blindMode <> BLIND_MODE_NONE) And NetPlay_isHost And (NetPlay_hostPlayer >= MAX_PLAYER_SLOTS) Then
            subText = subText & """"
            subText = subText & getPlayerGenericName(j)
            subText = subText & """"
        End If

        If NetPlay_bComms And j <> selectedPlayer Then
            ' (char buf[250])

            ' show ping time
            If Len(subText) = 0 Then buf = "Ping: " Else buf = ", Ping: "
            subText = subText & buf
            If ingame.PingTimes(j) < PING_LIMIT Then
                If game.blindMode = BLIND_MODE_NONE _
                    Or (NetPlay_isHost And (NetPlay_hostPlayer >= MAX_PLAYER_SLOTS)) _
                    Or (j = NetPlay_hostPlayer) Then
                    ' show actual ping time
                    buf = Format$(ingame.PingTimes(j), "000")  ' "%03d"
                Else
                    ' show non-exact ping in blind mode
                    pingQualifierStr = "Ok"
                    If ingame.PingTimes(j) > 350 Then
                        pingQualifierStr = "++"
                    ElseIf ingame.PingTimes(j) > 1000 Then
                        pingQualifierStr = "+++"
                    End If
                    buf = pingQualifierStr
                End If
            Else
                buf = ChrW$(&H221E)     ' Player has ping of somewhat questionable quality.
            End If
            subText = subText & buf
        End If

        dummy = stat.played < 5
        ' (uint8_t star[3] = {0, 0, 0}; uint8_t medal = 0;)

        ' star 1 total droid kills
        If stat.totalKills > 600 Then
            star(0) = 1
        ElseIf stat.totalKills > 300 Then
            star(0) = 2
        ElseIf stat.totalKills > 150 Then
            star(0) = 3
        Else
            star(0) = 0
        End If

        ' star 2 games played
        If stat.played > 200 Then
            star(1) = 1
        ElseIf stat.played > 100 Then
            star(1) = 2
        ElseIf stat.played > 50 Then
            star(1) = 3
        Else
            star(1) = 0
        End If

        ' star 3 games won.
        If stat.wins > 80 Then
            star(2) = 1
        ElseIf stat.wins > 40 Then
            star(2) = 2
        ElseIf stat.wins > 10 Then
            star(2) = 3
        Else
            star(2) = 0
        End If

        ' medals.
        If stat.wins >= 24 And stat.wins > 8 * stat.losses Then
            medal = 1
        ElseIf stat.wins >= 12 And stat.wins > 4 * stat.losses Then
            medal = 2
        ElseIf stat.wins >= 6 And stat.wins > 2 * stat.losses Then
            medal = 3
        Else
            medal = 0
        End If

        If hasAltName Then
            If cache.fullAltNameText <> pAltName Then
                altName = pAltName
                maxAltNameWidth = Fix(CSng(psWidget.width() - nameX) * 0.65!)
                fontID = font_small
                cache.wzAltNameText.setText altName, fontID
                cache.fullAltNameText = altName
                If cache.wzAltNameText.width() > maxAltNameWidth Then
                    Do While Len(altName) <> 0 And (iV_GetTextWidth(altName, cache.wzAltNameText.getFontID()) + iV_GetEllipsisWidth(cache.wzAltNameText.getFontID())) > maxAltNameWidth
                        altName = Left$(altName, Len(altName) - 1)  ' Clip alt name.
                    Loop
                    altName = altName & ChrW$(&H2026)
                    cache.wzAltNameText.setText altName, fontID
                End If
            End If
        End If

        If hasAltName And isHighlight Then
            ' display first, behind everything
            displayAltNameBox x, y, psWidget, cache, isHighlight
        End If

        H = 5
        If Len(subText) <> 0 Then textY = y + 22 - H Else textY = y + 22
        cache.wzMainText.render x + nameX, textY, colour
        If Len(subText) <> 0 Then
            cache.wzSubText.setText subText, font_small
            cache.wzSubText.render x + nameX, y + 28, psPalette(WZCOL_TEXT_MEDIUM)
        End If

        If (game.blindMode <> BLIND_MODE_NONE) And (Not NetPlay_isHost Or NetPlay_hostPlayer < MAX_PLAYER_SLOTS) Then
            iV_DrawImage FrontImages, IMAGE_WEE_GUY, x + 4, y + 13
        ElseIf dummy Then
            iV_DrawImage FrontImages, IMAGE_MEDAL_DUMMY, x + 4, y + 13
        Else
            starImgs(0) = 0: starImgs(1) = IMAGE_MULTIRANK1: starImgs(2) = IMAGE_MULTIRANK2: starImgs(3) = IMAGE_MULTIRANK3
            If 1 <= star(0) And star(0) < 4 Then
                iV_DrawImage FrontImages, starImgs(star(0)), x + 4, y + 3
            End If
            If 1 <= star(1) And star(1) < 4 Then
                iV_DrawImage FrontImages, starImgs(star(1)), x + 4, y + 13
            End If
            If 1 <= star(2) And star(2) < 4 Then
                iV_DrawImage FrontImages, starImgs(star(2)), x + 4, y + 23
            End If
            medalImgs(0) = 0: medalImgs(1) = IMAGE_MEDAL_GOLD: medalImgs(2) = IMAGE_MEDAL_SILVER: medalImgs(3) = IMAGE_MEDAL_BRONZE
            If 1 <= medal And medal < 4 Then
                iV_DrawImage FrontImages, medalImgs(medal), x + 16, y + 11
            End If
        End If

        If hasAltName And Not isHighlight Then
            ' display last, over top of everything
            displayAltNameBox x, y, psWidget, cache, isHighlight
        End If
    Else    ' AI
        ' (char aitext[100])

        If NetPlay_players_ai(j) >= 0 Then
            iV_DrawImage FrontImages, IMAGE_PLAYER_PC, x, y + 11
        End If

        ' challenges may use custom AIs that are not in aidata and set to 127
        If Not challengeActive Then
            If Not (NetPlay_players_ai(j) < aidata_size) Then ASSERT_FAILURE "NetPlay.players[j].ai < (int)aidata.size()", "src/titleui/widgets/lobbyplayerrow.cpp:468", "displayPlayer", "Uh-oh, AI index out of bounds": Exit Sub
        End If

        textColor = psPalette(WZCOL_FORM_TEXT)
        Select Case NetPlay_players_ai(j)
        Case AI_OPEN
            If Not NetPlay_players_isSpectator(j) Then
                aitext = "Open"
            Else
                aitext = "Spectator"
                textColor = pielightWithA(textColor, 220)
            End If
        Case AI_CLOSED: aitext = "Closed"
        Case Else: aitext = getPlayerName(j)
        End Select
        cache.wzMainText.setText aitext, font_regular
        cache.wzMainText.render x + nameX, y + 22, textColor
        cache.fullMainText = aitext
    End If
End Sub

Private Function canChooseTeamFor(ByVal i As Long) As Boolean
    If isBlindSimpleLobby(game.blindMode) And Not NetPlay_isHost Then
        canChooseTeamFor = False
        Exit Function
    End If
    canChooseTeamFor = (i = selectedPlayer Or isHostOrAdmin())
End Function

' ////////////////////////////////////////////////////////////////////////////
' Lobby player row

' WzPlayerRow(uint32_t playerIdx, const std::shared_ptr<WzMultiplayerOptionsTitleUI>& parent)
Public Sub WzPlayerRow_construct(ByVal self As WIDGET, ByVal playerIdx As Long, ByVal parent As WzMultiplayerOptionsTitleUI)
    WIDGET_construct self
    self.kind = K_WzPlayerRow ' (VB) virtual calls from here on reach WzPlayerRow's overrides
    WzPlayerRow_(self).parentTitleUI = ObjPtr(parent)
    WzPlayerRow_(self).playerIdx = playerIdx
End Sub

' static std::shared_ptr<WzPlayerRow> make(uint32_t playerIdx, const std::shared_ptr<WzMultiplayerOptionsTitleUI>& parent)
Public Function WzPlayerRow_make(ByVal playerIdx As Long, ByVal parent As WzMultiplayerOptionsTitleUI) As WIDGET
    Dim widget As WIDGET, d As WzPlayerRow, titleUI As Long
    Set widget = WIDGET_alloc(K_WzPlayerRow)
    WzPlayerRow_construct widget, playerIdx, parent
    Set d = WzPlayerRow_(widget)

    titleUI = ObjPtr(parent)    ' std::weak_ptr<WzMultiplayerOptionsTitleUI> titleUI(parent);

    ' add team button (also displays the spectator "eye" for spectators)
    Set d.teamButton = W_BUTTON_new()
    d.teamButton.setGeometry 0, 0, MULTIOP_TEAMSWIDTH, MULTIOP_TEAMSHEIGHT
    d.teamButton.UserData = playerIdx
    Set d.teamButton.displayFunction = WzFunc_new(FN_displayTeamChooser)
    widget.attach d.teamButton
    W_BUTTON_addOnClickHandler d.teamButton, WzFunc_new(FN_PR_TEAM_CLICK, playerIdx, titleUI)

    ' add player colour
    Set d.colorButton = W_BUTTON_new()
    d.colorButton.setGeometry MULTIOP_TEAMSWIDTH, 0, MULTIOP_COLOUR_WIDTH, MULTIOP_PLAYERHEIGHT
    d.colorButton.UserData = playerIdx
    Set d.colorButton.displayFunction = WzFunc_new(FN_displayColour)
    widget.attach d.colorButton
    W_BUTTON_addOnClickHandler d.colorButton, WzFunc_new(FN_PR_COLOUR_CLICK, playerIdx, titleUI)

    ' add ready button
    WzPlayerRow_updateReadyButton widget

    ' add player info box (takes up the rest of the space in the middle)
    Set d.playerInfo = W_BUTTON_new()
    d.playerInfo.style = d.playerInfo.style Or WBUT_SECONDARY
    d.playerInfo.UserData = playerIdx
    Set d.playerInfo.displayFunction = WzFunc_new(FN_displayPlayer)
    Set d.playerInfo.pUserData = New DisplayPlayerCache
    d.playerInfo.setOnDelete WzFunc_new(FN_PR_PLAYERINFO_ONDELETE)
    widget.attach d.playerInfo
    d.playerInfo.setCalcLayout WzFunc_new(FN_PR_PLAYERINFO_CALCLAYOUT)
    W_BUTTON_addOnClickHandler d.playerInfo, WzFunc_new(FN_PR_PLAYERINFO_CLICK, playerIdx, titleUI)

    ' update tooltips and such
    WzPlayerRow_updateState widget

    Set WzPlayerRow_make = widget
End Function

Public Sub WzPlayerRow_geometryChanged(ByVal self As WIDGET)
    Dim d As WzPlayerRow
    Set d = WzPlayerRow_(self)
    If Not d.readyButtonContainer Is Nothing Then
        d.readyButtonContainer.callCalcLayout
    End If
    If Not d.playerInfo Is Nothing Then
        d.playerInfo.callCalcLayout
    End If
End Sub

Public Sub WzPlayerRow_updateState(ByVal self As WIDGET)
    Dim d As WzPlayerRow, aidata_size As Long, locked As MultiplayOptionsLocked, playerIdx As Long
    Dim trueMultiplayerMode As Boolean, playerInfoTooltip As String, identity As EcKey, b64Identity As String
    Dim keyBytes() As Byte, n As Long
    Set d = WzPlayerRow_(self)
    aidata_size = getAIData_size()  ' const auto& aidata = getAIData();
    locked = getLockedOptions()
    playerIdx = d.playerIdx

    ' update team button tooltip
    If NetPlay_players_isSpectator(playerIdx) Then
        d.teamButton.setTip "Spectator"
    ElseIf canChooseTeamFor(playerIdx) And Not locked.teams Then
        d.teamButton.setTip "Choose Team"
    ElseIf locked.teams Then
        d.teamButton.setTip "Teams locked"
    Else
        d.teamButton.setTip vbNullString
    End If

    ' hide team button if needed
    trueMultiplayerMode = (bMultiPlayer And NetPlay_bComms) Or (Not NetPlay_isHost And ingame.localJoiningInProgress)
    If Not alliancesSetTeamsBeforeGame(game.alliance) And Not NetPlay_players_isSpectator(playerIdx) And Not trueMultiplayerMode Then
        d.teamButton.hide
    Else
        d.teamButton.show
    End If

    ' update color tooltip
    If (selectedPlayer = playerIdx Or NetPlay_isHost) And (Not NetPlay_players_isSpectator(playerIdx)) Then
        d.colorButton.setTip "Click to change player colour"
    Else
        d.colorButton.setTip vbNullString
    End If

    ' update player info box tooltip
    If (selectedPlayer = playerIdx Or NetPlay_isHost) _
        And NetPlay_players_allocated(playerIdx) _
        And Not locked.position _
        And Not isSpectatorOnlySlot(playerIdx) _
        And (Not isBlindSimpleLobby(game.blindMode) Or NetPlay_isHost) Then
        playerInfoTooltip = "Click to change player position"
    ElseIf Not NetPlay_players_allocated(playerIdx) Then
        If NetPlay_isHost And Not locked.ai Then
            If Not isSpectatorOnlySlot(playerIdx) Then
                playerInfoTooltip = "Click to change AI, right click to distribute choice"
            Else
                playerInfoTooltip = "Click to close spectator slot"
            End If
        ElseIf NetPlay_players_ai(playerIdx) >= 0 Then
            If NetPlay_isHost Or (Not trueMultiplayerMode) Or (game.blindMode = BLIND_MODE_NONE) Or challengeActive Then
                ' show AI description. Useful for challenges.
                If ULT(NetPlay_players_ai(playerIdx), aidata_size) Then
                    playerInfoTooltip = getAIData_tip(NetPlay_players_ai(playerIdx))
                End If
            End If
        End If
    End If
    If NetPlay_players_allocated(playerIdx) And (game.blindMode = BLIND_MODE_NONE Or (NetPlay_isHost And NetPlay_hostPlayer >= MAX_PLAYER_SLOTS)) Then
        identity = getOutputPlayerIdentity(playerIdx)
        If Not EcKey_empty(identity) Then
            If Len(playerInfoTooltip) <> 0 Then
                playerInfoTooltip = playerInfoTooltip & vbLf
            End If
            n = EcKey_toBytes_public(identity, keyBytes)
            b64Identity = base64Encode(keyBytes, n)
            playerInfoTooltip = playerInfoTooltip & "Player ID: "
            playerInfoTooltip = playerInfoTooltip & vbLf
            If Len(b64Identity) = 0 Then playerInfoTooltip = playerInfoTooltip & "(none)" Else playerInfoTooltip = playerInfoTooltip & b64Identity
        End If
    End If
    d.playerInfo.setTip playerInfoTooltip

    ' update ready button
    WzPlayerRow_updateReadyButton self
End Sub

Private Function createReadyButtonContainerForm(ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long) As WIDGET
    Dim form As WIDGET
    Set form = ReadyBoxContainerWidget_new()
    form.setGeometry x, y, w, h
    form.style = WFORM_PLAIN
    Set createReadyButtonContainerForm = form
End Function

' updateReadyButton's auto deleteExistingReadyButton = [this]() {...}
Private Sub WzPlayerRow_deleteExistingReadyButton(ByVal self As WIDGET)
    Dim d As WzPlayerRow
    Set d = WzPlayerRow_(self)
    If Not d.readyButton Is Nothing Then
        widgDelete_widget d.readyButton
        Set d.readyButton = Nothing
    End If
    If Not d.readyTextLabel Is Nothing Then
        widgDelete_widget d.readyTextLabel
        Set d.readyTextLabel = Nothing
    End If
    If WIDGET_isKindOf(d.readyButtonContainer, K_ReadyBoxContainerWidget) Then
        ReadyBoxContainerWidget_setShowAnimatedHighlight d.readyButtonContainer, False
    End If
End Sub

' updateReadyButton's auto deleteExistingDifficultyButton = [this]() {...}
Private Sub WzPlayerRow_deleteExistingDifficultyButton(ByVal self As WIDGET)
    Dim d As WzPlayerRow
    Set d = WzPlayerRow_(self)
    If Not d.difficultyChooserButton Is Nothing Then
        widgDelete_widget d.difficultyChooserButton
        Set d.difficultyChooserButton = Nothing
    End If
End Sub

Private Sub WzPlayerRow_updateReadyButton(ByVal self As WIDGET)
    Dim d As WzPlayerRow, aidata_size As Long, locked As MultiplayOptionsLocked, playerIdx As Long
    Dim playerDifficulty As Long, icon As Long, tooltip As String, difficultyTip As String, freshDifficultyButton As Boolean
    Dim player As Long, weakTitleUi As Long, tipres As String, alpha As Long
    Dim isMe As Boolean, meIdx As Long, isReady As Long, toolTips(0 To 1, 0 To 2) As String, images(0 To 1, 0 To 2) As Long
    Dim greyedOutReady As Boolean, freshReadyButton As Boolean, currentState As Long
    Set d = WzPlayerRow_(self)
    aidata_size = getAIData_size()  ' const auto& aidata = getAIData();
    locked = getLockedOptions()
    playerIdx = d.playerIdx

    If d.readyButtonContainer Is Nothing Then
        ' add form to hold 'ready' botton
        Set d.readyButtonContainer = createReadyButtonContainerForm(MULTIOP_PLAYERWIDTH - MULTIOP_READY_WIDTH, 0, _
                    MULTIOP_READY_WIDTH, MULTIOP_READY_HEIGHT)
        d.readyButtonContainer.UserData = playerIdx
        self.attach d.readyButtonContainer
        d.readyButtonContainer.setCalcLayout WzFunc_new(FN_PR_READYCONTAINER_CALCLAYOUT)
    End If

    If Not NetPlay_players_allocated(playerIdx) And NetPlay_players_ai(playerIdx) >= 0 Then
        ' Add AI difficulty chooser in place of normal "ready" button
        WzPlayerRow_deleteExistingReadyButton self
        playerDifficulty = ToInt8(NetPlay_players_difficulty(playerIdx))
        icon = difficultyIcon(playerDifficulty)
        ' (char tooltip[128 + 255]: left as it was, empty here, for a difficulty out of range)
        If playerDifficulty >= AIDifficulty_EASY And playerDifficulty < AIDifficulty_INSANE + 1 Then
            tooltip = getDifficultyListStr(playerDifficulty)
            If ULT(NetPlay_players_ai(playerIdx), aidata_size) Then
                difficultyTip = getAIData_difficultyTip(NetPlay_players_ai(playerIdx), playerDifficulty)
                If difficultyTip <> "" Then
                    tooltip = tooltip & vbLf
                    tooltip = tooltip & difficultyTip
                End If
            End If
        End If
        freshDifficultyButton = (d.difficultyChooserButton Is Nothing)
        If NetPlay_isHost And Not locked.difficulty Then tipres = "Click to change difficulty" Else tipres = tooltip
        If NetPlay_isHost And Not locked.difficulty Then alpha = 255 Else alpha = 125
        Set d.difficultyChooserButton = addMultiBut(d.readyButtonContainer, MULTIOP_DIFFICULTY_INIT_START + playerIdx, 6, 4, MULTIOP_READY_WIDTH, MULTIOP_READY_HEIGHT, _
                    tipres, icon, icon, icon, MAX_PLAYERS, alpha)
        player = playerIdx
        weakTitleUi = d.parentTitleUI
        If freshDifficultyButton Then
            W_BUTTON_addOnClickHandler d.difficultyChooserButton, WzFunc_new(FN_PR_DIFFICULTY_CLICK, player, weakTitleUi)
        End If
        Exit Sub
    ElseIf Not NetPlay_players_allocated(playerIdx) Then
        ' closed or open - remove ready / difficulty button
        WzPlayerRow_deleteExistingReadyButton self
        WzPlayerRow_deleteExistingDifficultyButton self
        Exit Sub
    End If

    If Not multiplayPlayersCanCheckReady() Then
        ' remove ready / difficulty button
        WzPlayerRow_deleteExistingReadyButton self
        WzPlayerRow_deleteExistingDifficultyButton self
        Exit Sub
    End If

    WzPlayerRow_deleteExistingDifficultyButton self

    isMe = playerIdx = selectedPlayer
    If isMe Then meIdx = 1 Else meIdx = 0
    If NETgetDownloadProgress(playerIdx) <> 100 Then
        isReady = 2
    ElseIf NetPlay_players_ready(playerIdx) Then
        isReady = 1
    Else
        isReady = 0
    End If
    toolTips(0, 0) = "Waiting for player": toolTips(0, 1) = "Player is ready": toolTips(0, 2) = "Player is downloading"
    toolTips(1, 0) = "Click when ready": toolTips(1, 1) = "Waiting for other players": toolTips(1, 2) = "Waiting for download"
    images(0, 0) = IMAGE_CHECK_OFF: images(0, 1) = IMAGE_CHECK_ON: images(0, 2) = IMAGE_CHECK_DOWNLOAD
    images(1, 0) = IMAGE_CHECK_OFF_HI: images(1, 1) = IMAGE_CHECK_ON_HI: images(1, 2) = IMAGE_CHECK_DOWNLOAD_HI

    ' draw 'ready' button
    greyedOutReady = (NetPlay_players_isSpectator(playerIdx) And NetPlay_players_ready(playerIdx)) Or (playerIdx <> selectedPlayer)
    freshReadyButton = (d.readyButton Is Nothing)
    If Not greyedOutReady Then alpha = 255 Else alpha = 125
    Set d.readyButton = addMultiBut(d.readyButtonContainer, MULTIOP_READY_START + playerIdx, 3, 10, MULTIOP_READY_WIDTH, MULTIOP_READY_HEIGHT, _
                toolTips(meIdx, isReady), images(0, isReady), images(0, isReady), images(meIdx, isReady), MAX_PLAYERS, alpha)
    If WIDGET_isKindOf(d.readyButtonContainer, K_ReadyBoxContainerWidget) Then
        ReadyBoxContainerWidget_setShowAnimatedHighlight d.readyButtonContainer, (isMe And isReady = 0 And Not greyedOutReady And multiplayPlayersShouldCheckReady())
    End If
    If Not (Not d.readyButton Is Nothing) Then ASSERT_FAILURE "readyButton != nullptr", "src/titleui/widgets/lobbyplayerrow.cpp:897", "updateReadyButton", "Failed to create ready button": Exit Sub
    W_BUTTON_(d.readyButton).minClickInterval = GAME_TICKS_PER_SEC
    W_BUTTON_unlock d.readyButton
    If greyedOutReady And Not NetPlay_isHost Then
        If WIDGET_isKindOf(d.readyButton, K_WzMultiButton) Then
            WzMultiButton_(d.readyButton).downStateMask = WBUT_DOWN Or WBUT_CLICKLOCK
        End If
        currentState = d.readyButton.getState()
        d.readyButton.setState currentState Or WBUT_LOCK
    End If
    If freshReadyButton Then
        ' must add onclick handler
        player = playerIdx
        weakTitleUi = d.parentTitleUI
        W_BUTTON_addOnClickHandler d.readyButton, WzFunc_new(FN_PR_READY_CLICK, player, weakTitleUi)
    End If

    If d.readyTextLabel Is Nothing Then
        Set d.readyTextLabel = W_LABEL_new()
        d.readyButtonContainer.attach d.readyTextLabel
        d.readyTextLabel.id = MULTIOP_READY_START + MAX_CONNECTED_PLAYERS + playerIdx
    End If
    d.readyTextLabel.setGeometry 0, 0, MULTIOP_READY_WIDTH, 17
    W_LABEL_setTextAlignment d.readyTextLabel, WLAB_ALIGNBOTTOM
    If isMe Then
        W_LABEL_setFont d.readyTextLabel, font_small, psPalette(WZCOL_TEXT_BRIGHT)
    Else
        W_LABEL_setFont d.readyTextLabel, font_small, psPalette(WZCOL_TEXT_MEDIUM)
    End If
    d.readyTextLabel.setString "READY?"
End Sub

' MARK: - the lambdas and display functions

Public Function lobbyplayerrow_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim strongTitleUI As WzMultiplayerOptionsTitleUI, locked As MultiplayOptionsLocked, playerIdx As Long, player As Long
    Dim button As WIDGET, psWidget As WIDGET, psParent As WIDGET, pButton As WIDGET, targets() As Boolean
    Dim identity As EcKey, b64Identity As String, keyBytes() As Byte, n As Long, i As Long
    Dim previousPlayersShouldCheckReadyValue As Boolean, x0 As Long, width_ As Long, msg As String
    Select Case f.id
    Case FN_displayColour
        displayColour args(0), args(1), args(2)
    Case FN_displayTeamChooser
        displayTeamChooser args(0), args(1), args(2)
    Case FN_displayPlayer
        displayPlayer args(0), args(1), args(2)

    Case FN_PR_TEAM_CLICK                       ' [playerIdx, titleUI](W_BUTTON& button)
        playerIdx = f.cap(0)
        Set strongTitleUI = titleUIFromPtr(f.cap(1))
        If Not (Not strongTitleUI Is Nothing) Then ASSERT_FAILURE "strongTitleUI != nullptr", "src/titleui/widgets/lobbyplayerrow.cpp:530", "make", "Title UI is gone?": Exit Function
        locked = getLockedOptions()
        If (Not locked.teams Or Not locked.spectators) Then  ' Clicked on a team chooser
            If canChooseTeamFor(playerIdx) Then
                widgScheduleTask WzFunc_new(FN_PR_TEAM_TASK, strongTitleUI, playerIdx)
            End If
        End If

    Case FN_PR_TEAM_TASK                        ' [strongTitleUI, playerIdx]
        Set strongTitleUI = f.cap(0)
        strongTitleUI.openTeamChooser f.cap(1)

    Case FN_PR_COLOUR_CLICK                     ' [playerIdx, titleUI](W_BUTTON& button)
        playerIdx = f.cap(0)
        Set strongTitleUI = titleUIFromPtr(f.cap(1))
        If Not (Not strongTitleUI Is Nothing) Then ASSERT_FAILURE "strongTitleUI != nullptr", "src/titleui/widgets/lobbyplayerrow.cpp:551", "make", "Title UI is gone?": Exit Function
        If playerIdx = selectedPlayer Or isHostOrAdmin() Then
            If Not NetPlay_players_isSpectator(playerIdx) Then ' not a spectator
                widgScheduleTask WzFunc_new(FN_PR_COLOUR_TASK, strongTitleUI, playerIdx)
            End If
        End If

    Case FN_PR_COLOUR_TASK                      ' [strongTitleUI, playerIdx]
        Set strongTitleUI = f.cap(0)
        strongTitleUI.openColourChooser f.cap(1)

    Case FN_PR_PLAYERINFO_ONDELETE              ' [](WIDGET *psWidget)
        Set psWidget = args(0)
        ' (assert(psWidget->pUserData != nullptr): the C assert, off in a release build)
        Set psWidget.pUserData = Nothing        ' delete static_cast<DisplayPlayerCache *>(psWidget->pUserData);

    Case FN_PR_PLAYERINFO_CALCLAYOUT            ' [](WIDGET *psWidget)
        Set psWidget = args(0)
        Set psParent = psWidget.parent()
        If Not WIDGET_isKindOf(psParent, K_WzPlayerRow) Then Set psParent = Nothing
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/titleui/widgets/lobbyplayerrow.cpp:580", "make", "Null parent": Exit Function
        x0 = MULTIOP_TEAMSWIDTH + MULTIOP_COLOUR_WIDTH
        width_ = WzPlayerRow_(psParent).readyButtonContainer.x() - x0
        psWidget.setGeometry x0, 0, width_, psParent.height()

    Case FN_PR_PLAYERINFO_CLICK                 ' [playerIdx, titleUI](W_BUTTON& button)
        Set button = args(0)
        playerIdx = f.cap(0)
        Set strongTitleUI = titleUIFromPtr(f.cap(1))
        If Not (Not strongTitleUI Is Nothing) Then ASSERT_FAILURE "strongTitleUI != nullptr", "src/titleui/widgets/lobbyplayerrow.cpp:587", "make", "Title UI is gone?": Exit Function
        locked = getLockedOptions()
        If W_BUTTON_getOnClickButtonPressed(button) = WKEY_SECONDARY And NetPlay_players_allocated(playerIdx) Then
            ' right-clicking on human players copies the player id to the clipboard
            If game.blindMode = BLIND_MODE_NONE Or (NetPlay_isHost And NetPlay_hostPlayer >= MAX_PLAYER_SLOTS) Then
                identity = getOutputPlayerIdentity(playerIdx)
                If Not EcKey_empty(identity) Then
                    n = EcKey_toBytes_public(identity, keyBytes)
                    b64Identity = base64Encode(keyBytes, n)
                    wzSetClipboardText b64Identity
                End If
            End If
            Exit Function
        End If
        If playerIdx = selectedPlayer Or isHostOrAdmin() Then
            player = playerIdx
            ' host/admin can move any player, clients can request to move themselves if there are available slots
            If ((player = selectedPlayer And validPlayerIdxTargetsForPlayerPositionMove(player, targets) > 0) Or _
                    (NetPlay_players_allocated(player) And isHostOrAdmin())) _
                And Not locked.position _
                And player < MAX_PLAYERS _
                And Not isSpectatorOnlySlot(player) _
                And (Not isBlindSimpleLobby(game.blindMode) Or NetPlay_isHost) Then
                widgScheduleTask WzFunc_new(FN_PR_POSITION_TASK, strongTitleUI, player)
            ElseIf Not NetPlay_players_allocated(player) And Not locked.ai And NetPlay_isHost Then
                If W_BUTTON_getOnClickButtonPressed(button) = WKEY_SECONDARY And player < MAX_PLAYERS Then
                    ' Right clicking distributes selected AI's type and difficulty to all other AIs
                    previousPlayersShouldCheckReadyValue = multiplayPlayersShouldCheckReady()
                    For i = 0 To MAX_PLAYERS - 1
                        ' Don't change open/closed slots or humans or spectator-only slots
                        If NetPlay_players_ai(i) >= 0 And i <> player And Not isHumanPlayer(i) And Not isSpectatorOnlySlot(i) Then
                            NetPlay_players_ai(i) = NetPlay_players_ai(player)
                            NetPlay_players_isSpectator(i) = NetPlay_players_isSpectator(player)
                            NetPlay_players_difficulty(i) = NetPlay_players_difficulty(player)
                            setPlayerName i, getAIName(player)
                            NETBroadcastPlayerInfo i
                        End If
                    Next
                    If ingame.localJoiningInProgress Then  ' Only if game hasn't actually started yet.
                        handlePossiblePlayersShouldCheckReadyChange previousPlayersShouldCheckReadyValue
                    End If
                    widgScheduleTask WzFunc_new(FN_PR_DISTRIBUTE_TASK, strongTitleUI)
                Else
                    widgScheduleTask WzFunc_new(FN_PR_AI_TASK, strongTitleUI, player)
                End If
            End If
        End If

    Case FN_PR_POSITION_TASK                    ' [strongTitleUI, player]
        Set strongTitleUI = f.cap(0)
        strongTitleUI.openPositionChooser f.cap(1)

    Case FN_PR_DISTRIBUTE_TASK                  ' [strongTitleUI]
        Set strongTitleUI = f.cap(0)
        strongTitleUI.updatePlayers
        resetReadyStatus False, isBlindSimpleLobby(game.blindMode)

    Case FN_PR_AI_TASK                          ' [strongTitleUI, player]
        Set strongTitleUI = f.cap(0)
        strongTitleUI.openAiChooser f.cap(1)

    Case FN_PR_READYCONTAINER_CALCLAYOUT        ' [](WIDGET *psWidget)
        Set psWidget = args(0)
        Set psParent = psWidget.parent()
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/titleui/widgets/lobbyplayerrow.cpp:795", "updateReadyButton", "Null parent": Exit Function
        psWidget.setGeometry psParent.width() - MULTIOP_READY_WIDTH, 0, psWidget.width(), psWidget.height()

    Case FN_PR_DIFFICULTY_CLICK                 ' [player, weakTitleUi](W_BUTTON&)
        player = f.cap(0)
        Set strongTitleUI = titleUIFromPtr(f.cap(1))
        If Not (Not strongTitleUI Is Nothing) Then ASSERT_FAILURE "strongTitleUI != nullptr", "src/titleui/widgets/lobbyplayerrow.cpp:853", "updateReadyButton", "Title UI is gone?": Exit Function
        locked = getLockedOptions()
        If Not locked.difficulty And NetPlay_isHost Then
            widgScheduleTask WzFunc_new(FN_PR_DIFFICULTY_TASK, strongTitleUI, player)
        End If

    Case FN_PR_DIFFICULTY_TASK                  ' [strongTitleUI, player]
        Set strongTitleUI = f.cap(0)
        strongTitleUI.openDifficultyChooser f.cap(1)

    Case FN_PR_READY_CLICK                      ' [player, weakTitleUi](W_BUTTON& button)
        Set pButton = args(0)                   ' button.shared_from_this()
        widgScheduleTask WzFunc_new(FN_PR_READY_TASK, f.cap(1), f.cap(0), pButton)

    Case FN_PR_READY_TASK                       ' [weakTitleUi, player, pButton]
        player = f.cap(1)
        Set pButton = f.cap(2)
        Set strongTitleUI = titleUIFromPtr(f.cap(0))
        If Not (Not strongTitleUI Is Nothing) Then ASSERT_FAILURE "strongTitleUI != nullptr", "src/titleui/widgets/lobbyplayerrow.cpp:919", "updateReadyButton", "Title UI is gone?": Exit Function

        If player = selectedPlayer _
            And (Not NetPlay_players_isSpectator(player) Or Not NetPlay_players_ready(player) Or NetPlay_isHost) Then ' spectators can never toggle off "ready"
            ' Lock the "ready" button (until the request is processed)
            pButton.setState WBUT_LOCK

            sendReadyRequest selectedPlayer, Not NetPlay_players_ready(player)

            ' if hosting try to start the game if everyone is ready
            If NetPlay_isHost And multiplayPlayersReady() Then
                startMultiplayerGame
                ' reset flag in case people dropped/quit on join screen
                NETsetPlayerConnectionStatus CONNECTIONSTATUS_NORMAL, NET_ALL_PLAYERS
            End If
        End If

        If NetPlay_isHost And Not alliancesSetTeamsBeforeGame(game.alliance) Then
            If mouseDown(MOUSE_RMB) And player <> NetPlay_hostPlayer Then ' both buttons....
                msg = "The host has kicked " & getPlayerName(player, True) & " from the game!"
                sendRoomSystemMessage msg
                kickPlayer player, "The host has kicked you from the game.", ERROR_KICKED, False
                resetReadyStatus True, shouldSkipReadyResetOnPlayerJoinLeaveEvent()     'reset and send notification to all clients
            End If
        End If
    End Select
End Function
