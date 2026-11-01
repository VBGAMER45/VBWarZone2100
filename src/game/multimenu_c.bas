Attribute VB_Name = "multimenu_c"
' src/multimenu.cpp, multimenu.h: the request lists of the lobby (the map
' chooser with its player-count buttons and search box, the player name
' chooser: addMultiRequest, runMultiRequester, closeMultiRequester,
' displayRequestOption, displayNumPlayersBut, stringRelevance) and their
' globals (psRScreen, MultiMenuUp, current_numplayers, multiRequestUp).
' The in-game multiplayer menu (cluster S): GetPlayerTextColor, the
' alliance / channel / mute displays, the widget classes MultiMenuDroidView,
' MultiMenuGrid, WzMultiMenuTabs and WzMultiWidget (their fields:
' MultiMenuDroidView.cls, MultiMenuGrid.cls, WzMultiWidget.cls),
' intCreateMultiMenuForm, intAddMultiMenu, intCloseMultiMenu and
' intCloseMultiMenuNoAnim.
' (VB) widget class MultiMenuDroidView : WIDGET
' (VB) widget class MultiMenuGrid : GridLayout
' (VB) widget class WzMultiMenuTabs : MultichoiceWidget
' (VB) widget class WzMultiWidget : IntFormAnimated
' A button's DisplayRequestOptionData is DisplayRequestOptionData.cls, a
' DisplayNumPlayersButCache its WzText.
Option Explicit

' ////////////////////////////////////////////////////////////////////////////
' defines

Public psRScreen As W_SCREEN            ' std::shared_ptr<W_SCREEN> = nullptr: requester stuff.

Public MultiMenuUp As Boolean           ' = false
Private current_context As Long         ' UDWORD = 0
Public current_numplayers As Long       ' UDWORD = 4 (multimenu_c_init)
Private current_searchString As String

' multimenu.h
Public Const MULTIMENU As Long = 10600

Public Const MULTIMENU_FORM As Long = MULTIMENU

Private Const MAX_CONNECTED_PLAYERS As Long = MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS

Private Const MULTIMENU_FORM_Y As Long = 25

Private Const MULTIMENU_CLOSE As Long = MULTIMENU + 1
Private Const MULTIMENU_PLAYER As Long = MULTIMENU + 2

Private Const MULTIMENU_ALLIANCE_BASE As Long = MULTIMENU_PLAYER + MAX_PLAYERS
Private Const MULTIMENU_GIFT_RAD As Long = MULTIMENU_ALLIANCE_BASE + MAX_PLAYERS
Private Const MULTIMENU_GIFT_RES As Long = MULTIMENU_GIFT_RAD + MAX_PLAYERS
Private Const MULTIMENU_GIFT_DRO As Long = MULTIMENU_GIFT_RES + MAX_PLAYERS
Private Const MULTIMENU_GIFT_POW As Long = MULTIMENU_GIFT_DRO + MAX_PLAYERS
Private Const MULTIMENU_CHANNEL As Long = MULTIMENU_GIFT_POW + MAX_PLAYERS

' (class MultiMenuGrid's const int32_t CELL_HORIZONTAL_PADDING = 8)
Private Const CELL_HORIZONTAL_PADDING As Long = 8

Private Const MULTIMENUFORM_INTERNAL_PADDING As Long = 10
Private Const MULTIMENUFORM_PANEL_TABS_HEIGHT As Long = 20

' requester stuff.
Private Const M_REQUEST_CLOSE As Long = MULTIMENU + 49
Private Const M_REQUEST As Long = MULTIMENU + 50

Private Const M_REQUEST_AP As Long = MULTIMENU + 70
Private Const M_REQUEST_2P As Long = MULTIMENU + 71
Private Const M_REQUEST_3P As Long = MULTIMENU + 72
Private Const M_REQUEST_4P As Long = MULTIMENU + 73
Private Const M_REQUEST_5P As Long = MULTIMENU + 74
Private Const M_REQUEST_6P As Long = MULTIMENU + 75
Private Const M_REQUEST_7P As Long = MULTIMENU + 76
Private Const M_REQUEST_8P As Long = MULTIMENU + 77
Private Const M_REQUEST_9P As Long = MULTIMENU + 78
Private Const M_REQUEST_10P As Long = MULTIMENU + 79
Private Const M_REQUEST_11P As Long = MULTIMENU + 80
Private Const M_REQUEST_12P As Long = MULTIMENU + 81
Private Const M_REQUEST_13P As Long = MULTIMENU + 82
Private Const M_REQUEST_14P As Long = MULTIMENU + 83
Private Const M_REQUEST_15P As Long = MULTIMENU + 84
Private Const M_REQUEST_16P As Long = MULTIMENU + 85
' static const unsigned M_REQUEST_NP[] = {M_REQUEST_2P, ..., M_REQUEST_16P}: M_REQUEST_NP_at(i)

Private Const M_REQUEST_BUT As Long = MULTIMENU + 100        ' allow loads of buttons.
Private Const M_REQUEST_BUTM As Long = MULTIMENU + 1100

Private Const MULTIMENU_PLAYER_MUTE_START As Long = M_REQUEST_BUTM + 1

Private Const M_REQUEST_X As Long = MULTIOP_PLAYERSX
Private Const M_REQUEST_Y As Long = MULTIOP_PLAYERSY
Private Const M_REQUEST_W As Long = MULTIOP_PLAYERSW
Private Const M_REQUEST_H As Long = MULTIOP_PLAYERSH

Private Const R_BUT_W As Long = 120 '105
Private Const R_BUT_H As Long = 30

Private Const HOVER_PREVIEW_TIME As Long = 300

' hci.h

Public multiRequestUp As Boolean        ' = false: multimenu is up.
Private hoverPreviewId As Long          ' unsigned
Private giftsUp(0 To MAX_PLAYERS - 1) As Boolean ' = {true}: gift buttons for player are up. (multimenu_c_init)

' the FN ids of this module's lambdas and function pointers (WzFunc)
Public Const FN_displayRequestOption As Long = 5100
Public Const FN_displayNumPlayersBut As Long = 5101
Public Const FN_requestForm_calcLayout As Long = 5102         ' LAMBDA_CALCLAYOUT_SIMPLE
Public Const FN_searchBox_onEditingStopped As Long = 5103     ' [searchDir, fileExtension, mode, numPlayers](W_EDITBOX& widg)
Public Const FN_initPUserData_DisplayNumPlayersButCache As Long = 5104 ' []() -> void * { return new DisplayNumPlayersButCache(); }
Public Const FN_multimenu_intDisplayImageHilight As Long = 5105 ' (VB) intdisplay.cpp's intDisplayImageHilight (cluster G): f_F2's stand-in
' the in-game menu's (cluster S, fork S1: 18000-18019)
Public Const FN_displayAllianceState As Long = 18000
Public Const FN_displayChannelState As Long = 18001
Public Const FN_displayMuteState As Long = 18002
Public Const FN_multimenu_channelOnClick As Long = 18003         ' [player](W_BUTTON&)
Public Const FN_multimenu_allianceOnClick As Long = 18004        ' [player](W_BUTTON& widg)
Public Const FN_multimenu_giftOnClick As Long = 18005            ' the four gift buttons' [player](W_BUTTON&) (cap(0): the gift)
Public Const FN_multimenu_muteOnClick As Long = 18006            ' [](W_BUTTON& button)
Public Const FN_multiWidget_closeOnClick As Long = 18007         ' [closeButtonHandler](W_BUTTON&)
Public Const FN_multiWidget_closeCalcLayout As Long = 18008      ' [](WIDGET *psButton)
Public Const FN_multiWidget_panelSwitcher_onChoose As Long = 18009 ' [](MultibuttonWidget& widget, int newValue)
Public Const FN_multiWidget_panelSwitcher_task As Long = 18010   ' widgScheduleTask([newValue, weakParent](){...})
Public Const FN_multiWidget_panelSwitcher_calcLayout As Long = 18011 ' LAMBDA_CALCLAYOUT_SIMPLE
Public Const FN_intAddMultiMenu_closeHandler As Long = 18012     ' []() { widgScheduleTask(...); }
Public Const FN_intAddMultiMenu_closeTask As Long = 18013        ' []() { intCloseMultiMenu(); }
Public Const FN_intAddMultiMenu_calcLayout As Long = 18014       ' [](WIDGET *form)

' (VB) the C++ initialisers of this module's globals
Public Sub multimenu_c_init()
    current_numplayers = 4
    giftsUp(0) = True
End Sub

Private Function M_REQUEST_NP_at(ByVal i As Long) As Long
    M_REQUEST_NP_at = M_REQUEST_2P + i
End Function

' ////////////////////////////////////////////////////////////////////////////
' Map / force / name load save stuff.

' Sets the player's text color depending on if alliance formed, or if dead.
' \param mode  the specified alliance
' \param player the specified player
Private Function GetPlayerTextColor(ByVal mode As Long, ByVal player As Long) As Long
    ' override color if they are dead...
    If Not ULT(player, MAX_PLAYERS) Then
        GetPlayerTextColor = psPalette(WZCOL_GREY)          ' dead text color
    ElseIf apsDroidLists(player).size = 0 And apsStructLists(player).size = 0 Then
        GetPlayerTextColor = psPalette(WZCOL_GREY)          ' dead text color
    ' the colors were chosen to match the FRIEND/FOE radar map colors.
    ElseIf mode = ALLIANCE_FORMED Then
        GetPlayerTextColor = psPalette(WZCOL_YELLOW)        ' Human alliance text color
    ElseIf isHumanPlayer(player) Then                       ' Human player, no alliance
        GetPlayerTextColor = psPalette(WZCOL_TEXT_BRIGHT)   ' Normal text color
    Else
        GetPlayerTextColor = psPalette(WZCOL_RED)           ' Enemy color
    End If
End Function

' ////////////////////////////////////////////////////////////////////////////
' ////////////////////////////////////////////////////////////////////////////

Public Sub displayRequestOption(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim data As DisplayRequestOptionData, mapData As LEVEL_DATASET, x As Long, y As Long, butString As String
    Dim colour As Long, fullButString As String, hash As Sha256, count As Long, cachedHash As Sha256
    Static lastHashTime As Long         ' static uint32_t lastHashTime = 0;
    ' Any widget using displayRequestOption must have its pUserData initialized to a (DisplayRequestOptionData*)
    ' (assert(psWidget->pUserData != nullptr))
    Set data = psWidget.pUserData

    Set mapData = data.pMapData

    x = xOffset + psWidget.x()
    y = yOffset + psWidget.y()

    butString = sstrcpy_cut(W_BUTTON_(psWidget).pTip, 255)

    drawBlueBox x, y, psWidget.width(), psWidget.height()

    If Not mapData Is Nothing Then
        If CheckForMod(mapData.realFileName) Then colour = psPalette(WZCOL_RED) Else colour = psPalette(WZCOL_TEXT_BRIGHT)
    Else
        colour = psPalette(WZCOL_TEXT_BRIGHT)
    End If

    If Not data.canUseCachedText(butString, psWidget.width()) Then
        fullButString = butString
        Do While iV_GetTextWidth(butString, font_regular) > psWidget.width() - 10
            butString = Left$(butString, Len(butString) - 1)
        Loop
        data.setCachedText butString, font_regular, fullButString, psWidget.width()
    End If

    data.renderCachedText x + 6, y + 12, colour  'draw text

    If Not mapData Is Nothing Then
        ' Display map hash, so we can see the difference between identically named maps.
        hash = mapData.realFileHash  ' levGetFileHash can be slightly expensive.
        If lastHashTime <> realTime And Sha256_isZero(hash) Then
            hash = levGetFileHash(mapData)
            If Not Sha256_isZero(hash) Then
                lastHashTime = realTime  ' We just calculated a hash. Don't calculate any more hashes this frame.
            End If
        End If
        If Not Sha256_isZero(hash) Then
            cachedHash = data.hash
            If Not Sha256_equals(hash, cachedHash) Then
                butString = sstrcpy_cut(Sha256_toString(hash), 255)
                Do While iV_GetTextWidth(butString, font_small) > psWidget.width() - 10 - (8 + mapData.players * 6)
                    butString = Left$(butString, Len(butString) - 1)
                Loop

                ' Update the cached hash text
                data.hash = hash
                data.wzHashText.setText butString, font_small
            End If
            data.wzHashText.render x + 6 + 8 + mapData.players * 6, y + 26, psPalette(WZCOL_TEXT_DARK)
        End If

        ' if map, then draw no. of players.
        For count = 0 To mapData.players - 1
            iV_DrawImage FrontImages, IMAGE_WEE_GUY, x + 6 * count + 6, y + 16
        Next
        If CheckForRandom(mapData.realFileName, mapData.apDataFiles(0)) Then
            iV_DrawImage FrontImages, IMAGE_WEE_DIE, x + 80 + 6, y + 15
        End If
    End If
End Sub

' ////////////////////////////////////////////////////////////////////////////
' ////////////////////////////////////////////////////////////////////////////

Private Sub displayNumPlayersBut(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim cache As WzText, x As Long, y As Long, buffer As String, colour As Long
    ' Any widget using displayNumPlayersBut must have its pUserData initialized to a (DisplayNumPlayersButCache*)
    ' (its wzText: the pUserData itself)
    Set cache = psWidget.pUserData

    x = xOffset + psWidget.x()
    y = yOffset + psWidget.y()

    drawBlueBox x, y, psWidget.width(), psWidget.height()

    If psWidget.UserData = current_numplayers Then
        colour = psPalette(WZCOL_TEXT_BRIGHT)
    Else
        colour = psPalette(WZCOL_TEXT_MEDIUM)
    End If
    If psWidget.UserData = 0 Then
        buffer = " *"
    Else
        buffer = va("%iP", psWidget.UserData)
        buffer = Left$(buffer, 2)  ' Truncate 'P' if 2 digits, since there isn't room.
    End If
    cache.setText buffer, font_regular
    cache.render x + 2, y + 12, colour
End Sub

' (VB) WzString::normalized(NormalizationForm_KD) is not applied (the strings are compared as they are); a
' codepoint's caseFolded() is its LCase
Private Function stringRelevance(ByVal string_ As String, ByVal search As String) As Long
    Dim str_ As String, sea As String, strDim As Long, seaDim As Long, scores() As Long, sum As Long
    Dim iStr As Long, iSea As Long, score As Long, a As String, b As String, strSize As Long, seaSize As Long
    str_ = string_
    sea = search
    strSize = Len(str_)
    seaSize = Len(sea)
    strDim = strSize + 1
    seaDim = seaSize + 1

    If strDim > 10000 Or seaDim > 10000 Or strDim * seaDim > 100000 Then
        stringRelevance = 0
        Exit Function
    End If

    ReDim scores(0 To strDim * seaDim - 1)
    For sum = 0 To strSize + seaSize
        For iStr = MAX_(0, sum - seaSize) To MIN_(strSize, sum - 0)
            iSea = sum - iStr
            score = 0
            If iStr > 0 And iSea > 0 Then
                score = (scores(iStr - 1 + (iSea - 1) * strDim) + 1) Or 1
                a = Mid$(str_, iStr, 1)
                b = Mid$(sea, iSea, 1)
                If a = b Then
                    score = score + 100
                ElseIf LCase$(a) = LCase$(b) Then
                    score = score + 80
                End If
            End If
            If iStr > 0 Then
                score = MAX_(score, scores(iStr - 1 + iSea * strDim) And Not 1)
            End If
            If iSea > 0 Then
                score = MAX_(score, scores(iStr + (iSea - 1) * strDim) And Not 1)
            End If
            scores(iStr + iSea * strDim) = score
        Next
    Next

    stringRelevance = scores(strSize + seaSize * strDim)
End Function

Public Sub multiMenuScreenSizeDidChange(ByVal oldWidth As Long, ByVal oldHeight As Long, ByVal newWidth As Long, ByVal newHeight As Long)
    If psRScreen Is Nothing Then Exit Sub
    psRScreen.screenSizeDidChange oldWidth, oldHeight, newWidth, newHeight
End Sub

'* Searches in the given search directory for files ending with the
'*  given extension. Then will create a window with buttons for each
'*  found file.
'*  \param searchDir the directory to search in
'*  \param fileExtension the extension files should end with, if the
'*         extension has a dot (.) then this dot _must_ be present as
'*         the first char in this parameter
'*  \param mode (purpose unknown)
'*  \param numPlayers (purpose unknown)
Public Sub addMultiRequest(ByVal searchDir As String, ByVal fileExtension As String, ByVal mode As Long, ByVal numPlayers As Long, Optional ByVal searchString As String = "")
    Dim extensionLength As Long, buttonsX As Long, backdrop As WIDGET, requestForm As WIDGET, sButInit As W_BUTINIT
    Dim requestList As WIDGET, nextButtonId As Long, files() As String, nFiles As Long, f As Long, currFile As String
    Dim fileNameLength As Long, withoutExtension As String, withoutTechlevel As String, button As WIDGET
    Dim searchBox As WIDGET, levels() As LEVEL_DATASET, nLevels As Long, i As Long, j As Long
    Dim buttons() As WIDGET, relevances() As Long, nButtons As Long, data As DisplayRequestOptionData
    Dim mapNumPlayers As Long, ttip As String, emptyInit As W_BUTINIT, keyRel As Long, keyBut As WIDGET
    extensionLength = Len(fileExtension)
    If mode = MULTIOP_MAP Then buttonsX = 22 Else buttonsX = 17

    current_context = mode
    If mode = MULTIOP_MAP Then
        ' only save these when they select MAP button
        current_numplayers = numPlayers
        current_searchString = searchString
    End If

    Set psRScreen = W_SCREEN_make() '/< move this to intinit or somewhere like that.. (close too.)

    Set backdrop = addBackdrop(psRScreen)
    backdrop.setCalcLayout WzFunc_new(FN_backdrop_calcLayout)

    ' add a form to place the tabbed form on
    Set requestForm = IntFormAnimated_new()
    backdrop.attach requestForm
    requestForm.id = M_REQUEST
    requestForm.setCalcLayout WzFunc_new(FN_requestForm_calcLayout)

    ' Add the close button.
    sButInit.formID = M_REQUEST
    sButInit.id = M_REQUEST_CLOSE
    sButInit.x = M_REQUEST_W - CLOSE_WIDTH - 3
    sButInit.y = 0
    sButInit.width = CLOSE_WIDTH
    sButInit.height = CLOSE_HEIGHT
    sButInit.pTip = "Close"
    Set sButInit.pDisplay = WzFunc_new(FN_multimenu_intDisplayImageHilight)
    sButInit.UserData = PACKDWORD_TRI(0, IMAGE_CLOSEHILIGHT, IMAGE_CLOSE)
    widgAddButton psRScreen, sButInit

    ' Add the button list.
    Set requestList = IntListTabWidget_make()
    requestForm.attach requestList
    ListTabWidget_setChildSize requestList, R_BUT_W, R_BUT_H
    ListTabWidget_setChildSpacing requestList, 4, 4
    requestList.setGeometry 2 + buttonsX, 2, sButInit.x - buttonsX - 8, M_REQUEST_H - 4

    ' Put the buttons on it.
    nextButtonId = M_REQUEST_BUT
    ' (WZ_PHYSFS_enumerateFiles(searchDir, [&](const char *currFile) -> bool {...}): the files, then the lambda's body)
    If WZ_PHYSFS_enumerateFiles(searchDir, files, nFiles) Then
        For f = 0 To nFiles - 1
            currFile = files(f)
            fileNameLength = Len(currFile)

            ' Check to see if this file matches the given extension
            If fileNameLength <= extensionLength Or Right$(currFile, extensionLength) <> fileExtension Then
                GoTo continue_ 'continue
            End If

            withoutExtension = Left$(currFile, fileNameLength - extensionLength)
            withoutTechlevel = mapNameWithoutTechlevel(withoutExtension)

            ' Set the tip and add the button
            Set button = W_BUTTON_new()
            requestList.attach button
            button.id = nextButtonId
            button.setTip withoutTechlevel
            button.setString withoutTechlevel
            Set button.displayFunction = WzFunc_new(FN_displayRequestOption)
            Set button.pUserData = DisplayRequestOptionData_new(Nothing)
            ' (setOnDelete deletes the DisplayRequestOptionData: it goes with the widget)
            ListTabWidget_addWidgetToLayout requestList, button

            ' Update the init struct for the next button
            nextButtonId = nextButtonId + 1
continue_:
        Next
    End If

    multiRequestUp = True
    hoverPreviewId = 0

    If mode = MULTIOP_MAP Then
        ' Add the search edit box
        Set searchBox = W_EDITBOX_new()
        requestForm.attach searchBox
        searchBox.setGeometry 3, requestForm.height() - MULTIOP_SEARCHBOXH - 3, requestForm.width() - 6, MULTIOP_SEARCHBOXH
        W_EDITBOX_setBoxColours searchBox, psPalette(WZCOL_MENU_BORDER), psPalette(WZCOL_MENU_BORDER), psPalette(WZCOL_MENU_BACKGROUND)
        W_EDITBOX_setPlaceholder searchBox, "Search for map"
        searchBox.setString current_searchString

        W_EDITBOX_setOnEditingStoppedHandler searchBox, WzFunc_new(FN_searchBox_onEditingStopped, searchDir, fileExtension, mode, numPlayers)

        nLevels = enumerateMultiMaps(game.techLevel, numPlayers, levels)
        ' (std::vector<std::pair<int, std::shared_ptr<W_BUTTON>>> buttons)
        If nLevels > 0 Then
            ReDim buttons(0 To nLevels - 1)
            ReDim relevances(0 To nLevels - 1)
        End If

        For i = 0 To nLevels - 1
            withoutTechlevel = mapNameWithoutTechlevel(levels(i).pName)
            ' add number of players to string.
            Set button = W_BUTTON_new()
            requestList.attach button
            button.id = nextButtonId
            button.setTip withoutTechlevel
            button.setString withoutTechlevel
            Set button.displayFunction = WzFunc_new(FN_displayRequestOption)
            Set button.pUserData = DisplayRequestOptionData_new(levels(i))
            ' (setOnDelete deletes the DisplayRequestOptionData: it goes with the widget)
            relevances(nButtons) = stringRelevance(levels(i).pName, searchString)
            Set buttons(nButtons) = button
            nButtons = nButtons + 1

            nextButtonId = nextButtonId + 1
        Next
        ' std::stable_sort(buttons, [](a, b) { return a.first > b.first; }): an insertion sort
        For i = 1 To nButtons - 1
            keyRel = relevances(i)
            Set keyBut = buttons(i)
            j = i - 1
            Do While j >= 0
                If Not (keyRel > relevances(j)) Then Exit Do
                relevances(j + 1) = relevances(j)
                Set buttons(j + 1) = buttons(j)
                j = j - 1
            Loop
            relevances(j + 1) = keyRel
            Set buttons(j + 1) = keyBut
        Next
        For i = 0 To nButtons - 1
            ListTabWidget_addWidgetToLayout requestList, buttons(i)
        Next

        ' if it's map select then add the cam style buttons.
        sButInit = emptyInit
        sButInit.formID = M_REQUEST
        sButInit.x = 3
        sButInit.width = 17
        sButInit.height = 17

        sButInit.id = M_REQUEST_AP
        sButInit.y = 17
        sButInit.UserData = 0
        sButInit.pTip = "Any number of players"
        Set sButInit.pDisplay = WzFunc_new(FN_displayNumPlayersBut)
        Set sButInit.initPUserDataFunc = WzFunc_new(FN_initPUserData_DisplayNumPlayersButCache)
        ' (sButInit.onDelete deletes the DisplayNumPlayersButCache: its WzText goes with the widget)
        widgAddButton psRScreen, sButInit

        ' STATIC_ASSERT(MAX_PLAYERS_IN_GUI <= ARRAY_SIZE(M_REQUEST_NP) + 1);
        For mapNumPlayers = 2 To MAX_PLAYERS_IN_GUI
            sButInit.id = M_REQUEST_NP_at(mapNumPlayers - 2)
            sButInit.y = sButInit.y + 22
            sButInit.UserData = mapNumPlayers
            If mapNumPlayers = 1 Then ttip = va("%d player", mapNumPlayers) Else ttip = va("%d players", mapNumPlayers) ' ngettext
            sButInit.pTip = ttip
            widgAddButton psRScreen, sButInit
        Next
    End If
End Sub

Public Sub closeMultiRequester()
    multiRequestUp = False
    resetReadyStatus False
    Set psRScreen = Nothing
End Sub

' mode, chosen, chosenValue, isHoverPreview: the pointers' targets (UDWORD *mode, WzString *chosen,
' LEVEL_DATASET **chosenValue, bool *isHoverPreview)
Public Function runMultiRequester(ByVal id As Long, mode As Long, chosen As String, chosenValue As LEVEL_DATASET, isHoverPreview As Boolean) As Boolean
    Static hoverId As Long              ' static unsigned hoverId = 0;
    Static hoverStartTime As Long       ' static unsigned hoverStartTime = 0;
    Dim hoverPreview As Boolean, pButton As WIDGET, pData As DisplayRequestOptionData, numPlayers As Long

    isHoverPreview = False

    If (id = M_REQUEST_CLOSE) Or CancelPressed() Then         ' user hit close box || hit the cancel key
        closeMultiRequester
        mode = 0
        runMultiRequester = True
        Exit Function
    End If

    hoverPreview = False
    If id = 0 And current_context = MULTIOP_MAP Then
        id = widgGetMouseOver(psRScreen)
        If id <> hoverId Then
            hoverId = id
            hoverStartTime = wzGetTicks() + HOVER_PREVIEW_TIME
        End If
        If id = hoverPreviewId Or hoverStartTime > wzGetTicks() Then
            id = 0  ' Don't re-render preview nor render preview before HOVER_PREVIEW_TIME.
        End If
        hoverPreview = True
    End If
    If id >= M_REQUEST_BUT And id <= M_REQUEST_BUTM Then  ' chose a file.
        Set pButton = widgGetFromID(psRScreen, id)
        chosen = W_BUTTON_(pButton).pText

        If IsObject(pButton.pUserData) Then Set pData = pButton.pUserData
        If Not (Not pData Is Nothing) Then ASSERT_FAILURE "pData != nullptr", "src/multimenu.cpp:608", "runMultiRequester", "Unable to get map data pointer for: %s", chosen: Exit Function
        If current_context = MULTIOP_MAP Then
            If Not (Not pData.pMapData Is Nothing) Then ASSERT_FAILURE "pData->pMapData != nullptr", "src/multimenu.cpp:611", "runMultiRequester", "Unable to get map data: %s", chosen: Exit Function
            Set chosenValue = pData.pMapData
        Else
            Set chosenValue = Nothing
        End If
        mode = current_context
        isHoverPreview = hoverPreview
        hoverPreviewId = id
        If Not hoverPreview Then
            closeMultiRequester
        End If

        runMultiRequester = True
        Exit Function
    End If
    If hoverPreview Then
        id = 0
    End If

    Select Case id
    Case M_REQUEST_AP
        closeMultiRequester
        addMultiRequest MultiCustomMapsPath, ".wrf", MULTIOP_MAP, 0, current_searchString
    Case Else
        For numPlayers = 2 To MAX_PLAYERS_IN_GUI
            If id = M_REQUEST_NP_at(numPlayers - 2) Then
                closeMultiRequester
                addMultiRequest MultiCustomMapsPath, ".wrf", MULTIOP_MAP, numPlayers, current_searchString
                Exit For
            End If
        Next
    End Select

    runMultiRequester = False
End Function

' (VB) DisplayRequestOptionData(LEVEL_DATASET *pMapData = nullptr)
Private Function DisplayRequestOptionData_new(ByVal pMapData As LEVEL_DATASET) As DisplayRequestOptionData
    Dim d As DisplayRequestOptionData
    Set d = New DisplayRequestOptionData
    Set d.pMapData = pMapData
    Set DisplayRequestOptionData_new = d
End Function


' ////////////////////////////////////////////////////////////////////////////
' Multiplayer in game menu stuff

' ////////////////////////////////////////////////////////////////////////////
' alliance display funcs

Private Sub displayAllianceState(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim a As Long, b As Long, c As Long, player As Long
    If Not ULT(selectedPlayer, MAX_PLAYERS) Then Exit Sub

    player = psWidget.UserData
    Select Case alliances(selectedPlayer, player)
    Case ALLIANCE_BROKEN
        a = 0
        b = IMAGE_MULTI_NOAL_HI
        c = IMAGE_MULTI_NOAL                    ' replace with real gfx
    Case ALLIANCE_FORMED
        a = 0
        b = IMAGE_MULTI_AL_HI
        c = IMAGE_MULTI_AL                      ' replace with real gfx
    Case ALLIANCE_REQUESTED, ALLIANCE_INVITATION
        a = 0
        b = IMAGE_MULTI_OFFAL_HI
        c = IMAGE_MULTI_OFFAL                   ' replace with real gfx
    Case Else
        a = 0
        b = IMAGE_MULTI_NOAL_HI
        c = IMAGE_MULTI_NOAL
    End Select

    psWidget.UserData = PACKDWORD_TRI(a, b, c)
    intDisplayImageHilight psWidget, xOffset, yOffset
    psWidget.UserData = player
End Sub

Private Sub displayChannelState(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim player As Long
    player = psWidget.UserData
    If Not ULT(player, MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "player < MAX_CONNECTED_PLAYERS", "src/multimenu.cpp:700", "displayChannelState", "invalid player: %u", player: Exit Sub

    If openchannels(player) Then
        psWidget.UserData = PACKDWORD_TRI(0, IMAGE_MULTI_CHAN, IMAGE_MULTI_CHAN)
    Else
        psWidget.UserData = PACKDWORD_TRI(0, IMAGE_MULTI_NOCHAN, IMAGE_MULTI_NOCHAN)
    End If
    intDisplayImageHilight psWidget, xOffset, yOffset
    psWidget.UserData = player
End Sub

Private Sub displayMuteState(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim player As Long, ImageID As Long, x As Long, y As Long, isHighlighted As Boolean, isDown As Boolean, colour As Long
    player = psWidget.UserData
    If Not ULT(player, MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "player < MAX_CONNECTED_PLAYERS", "src/multimenu.cpp:717", "displayMuteState", "invalid player: %u", player: Exit Sub

    ImageID = IMAGE_INTFAC_VOLUME_UP
    If Not ingame.muteChat(player) Then
        ImageID = IMAGE_INTFAC_VOLUME_UP
    Else
        ImageID = IMAGE_INTFAC_VOLUME_MUTE
    End If
    x = xOffset + psWidget.x()
    y = yOffset + psWidget.y()
    isHighlighted = (psWidget.getState() And WBUT_HIGHLIGHT) <> 0
    isDown = (psWidget.getState() And WBUT_DOWN) <> 0
    If isHighlighted Or isDown Then
        If isDown Then colour = psPalette(WZCOL_TRANSPARENT_BOX) Else colour = pal_RGBA(255, 255, 255, 65)
        pie_UniTransBoxFill x - 2, y - 2, x + 16 + 4, y + 16 + 4, colour
    End If
    iV_DrawImageFileAnisotropic IntImages, ImageID, x + 1, y + 1, 16, 16
End Sub


' ////////////////////////////////////////////////////////////////////////////

' MARK: - MultiMenuDroidView

' MultiMenuDroidView(uint32_t player): player(player)
Public Sub MultiMenuDroidView_construct(ByVal self As WIDGET, ByVal player As Long)
    WIDGET_construct self
    self.kind = K_MultiMenuDroidView ' (VB) virtual calls from here on reach MultiMenuDroidView's overrides
    MultiMenuDroidView_(self).player = player
    self.setGeometry 0, 0, 60, 32
End Sub

' std::make_shared<MultiMenuDroidView>(player)
Private Function MultiMenuDroidView_new(ByVal player As Long) As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_MultiMenuDroidView)
    MultiMenuDroidView_construct w, player
    Set MultiMenuDroidView_new = w
End Function

Public Sub MultiMenuDroidView_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim player As Long, it As Long, displayDroid As Long, centerX As Long, y0 As Long, rotationV As Vector3i, position As Vector3i
    player = MultiMenuDroidView_(self).player
    ' a droid of theirs.
    ' (DroidList* displayDroidList = (player < MAX_PLAYERS) ? &apsDroidLists[player] : nullptr)
    If Not ULT(player, MAX_PLAYERS) Then
        Exit Sub
    End If
    ' (std::find_if: the first droid visibleForLocalDisplay, 0 = end())
    it = apsDroidLists(player).head
    Do While it <> 0
        If BASE_OBJECT_visibleForLocalDisplay(lnObj(it)) <> 0 Then Exit Do
        it = lnNext(it)
    Loop

    centerX = xOffset + self.x() + self.width() \ 2
    y0 = yOffset + self.y()
    If it <> 0 Then
        displayDroid = lnObj(it)
        pie_SetGeometricOffset centerX, y0 + self.height() * 3 \ 4
        rotationV = Vector3i_(-15, 45, 0)
        position = Vector3i_(0, 0, BUTTON_DEPTH)  ' Scale them.
        If aDroid(displayDroid).droidType = DROID_SUPERTRANSPORTER Then
            position.z = 7850
        ElseIf aDroid(displayDroid).droidType = DROID_TRANSPORTER Then
            position.z = 4100
        End If

        displayComponentButtonObject displayDroid, rotationV, position, 100
    ElseIf ULT(player, MAX_PLAYERS) And apsDroidLists(player).size <> 0 Then
        ' Show that they have droids, but not which droids, since we can't see them.
        iV_DrawImageTc_ids IntImages, IMAGE_GENERIC_TANK, IMAGE_GENERIC_TANK_TC, _
            centerX - iV_GetImageWidth(IntImages, IMAGE_GENERIC_TANK) \ 2, _
            y0 + self.height() \ 2 - iV_GetImageHeight(IntImages, IMAGE_GENERIC_TANK) \ 2, _
            pal_GetTeamColour(getPlayerColour(player))
    End If
End Sub

' MARK: - MultiMenuGrid

Public Sub MultiMenuGrid_construct(ByVal self As WIDGET)
    GridLayout_construct self
    self.kind = K_MultiMenuGrid ' (VB) virtual calls from here on reach MultiMenuGrid's overrides
End Sub

' static std::shared_ptr<MultiMenuGrid> make()
Private Function MultiMenuGrid_make() As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_MultiMenuGrid)
    MultiMenuGrid_construct widget
    MultiMenuGrid_initialize widget
    Set MultiMenuGrid_make = widget
End Function

Private Sub MultiMenuGrid_initialize(ByVal self As WIDGET)
    Dim d As MultiMenuGrid, lastRight As Long, player As Long
    Set d = MultiMenuGrid_(self)
    ' auto margin = Margin(10, CELL_HORIZONTAL_PADDING); (10, 8, 10, 8)
    ' auto lastMargin = margin; lastMargin.right = CLOSE_WIDTH + 5;
    lastRight = CLOSE_WIDTH + 5
    Set d.timerLabel = MultiMenuGrid_makeLabel("")
    W_LABEL_setCanTruncate d.timerLabel, False
    Set d.powerLabel = MultiMenuGrid_makeLabel("Power")
    Set d.pingLabel = MultiMenuGrid_makeLabel("Ping")
    Set d.structsLabel = MultiMenuGrid_makeLabel("Structs")
    GridLayout_place self, grid_allocation_slot_(0), grid_allocation_slot_(0), Margin_wrap(10, CELL_HORIZONTAL_PADDING, 10, CELL_HORIZONTAL_PADDING, Alignment_wrap(Alignment_center(), d.timerLabel))
    GridLayout_place self, grid_allocation_slot_(1), grid_allocation_slot_(0), Margin_wrap(10, CELL_HORIZONTAL_PADDING, 10, CELL_HORIZONTAL_PADDING, MultiMenuGrid_makeLabel("Alliances"))
    GridLayout_place self, grid_allocation_slot_(2), grid_allocation_slot_(0), Resize_wrap(Resize_width(100), Margin_wrap(10, CELL_HORIZONTAL_PADDING, 10, CELL_HORIZONTAL_PADDING, MultiMenuGrid_makeLabel("Score")))
    GridLayout_place self, grid_allocation_slot_(3), grid_allocation_slot_(0), Resize_wrap(Resize_width(50), Margin_wrap(10, CELL_HORIZONTAL_PADDING, 10, CELL_HORIZONTAL_PADDING, MultiMenuGrid_makeLabel("Kills")))
    GridLayout_place self, grid_allocation_slot_(4), grid_allocation_slot_(0), Resize_wrap(Resize_width(50), Margin_wrap(10, CELL_HORIZONTAL_PADDING, 10, CELL_HORIZONTAL_PADDING, MultiMenuGrid_makeLabel("Units")))
    GridLayout_place self, grid_allocation_slot_(5), grid_allocation_slot_(0), Resize_wrap(Resize_width(50), Margin_wrap(10, lastRight, 10, CELL_HORIZONTAL_PADDING, d.powerLabel))
    GridLayout_place self, grid_allocation_slot_(5), grid_allocation_slot_(0), Resize_wrap(Resize_width(50), Margin_wrap(10, lastRight, 10, CELL_HORIZONTAL_PADDING, d.pingLabel))
    GridLayout_place self, grid_allocation_slot_(5), grid_allocation_slot_(0), Resize_wrap(Resize_width(50), Margin_wrap(10, lastRight, 10, CELL_HORIZONTAL_PADDING, d.structsLabel))
    GridLayout_place self, grid_allocation_slot_(6), grid_allocation_slot_(0), Resize_wrap(Resize_width(16), Margin_wrap(10, lastRight, 10, CELL_HORIZONTAL_PADDING, MultiMenuGrid_makeLabel("")))

    For player = 0 To MAX_CONNECTED_PLAYERS - 1
        If isHumanPlayer(player) Or (game.typ = LEVEL_TYPE_SKIRMISH And player < game.maxPlayers And NetPlay_players_difficulty(player) <> AIDifficulty_DISABLED) Then
            MultiMenuGrid_addPlayerWidgets self, player, 1 + NetPlay_players_position(player)
        End If
    Next
End Sub

Private Function MultiMenuGrid_makeLabel(ByVal text As String) As WIDGET
    Dim label As WIDGET
    Set label = W_LABEL_new()
    label.setString text
    W_LABEL_setCanTruncate label, True
    W_LABEL_setFontColour label, psPalette(WZCOL_TEXT_BRIGHT)
    Set MultiMenuGrid_makeLabel = label
End Function

Private Function MultiMenuGrid_wrapGift(ByVal gift As WIDGET) As WIDGET
    Set MultiMenuGrid_wrapGift = Alignment_wrap(Alignment_(Alignment_Vertical_Center, Alignment_Horizontal_Center), Margin_wrap(0, 1, 0, 0, gift))
End Function

Private Sub MultiMenuGrid_addPlayerWidgets(ByVal self As WIDGET, ByVal player As Long, ByVal row As Long)
    Dim d As MultiMenuGrid, nameGrid As WIDGET, nameLabel As WIDGET, alliancesGrid As WIDGET, selectedPlayerIsSpectator As Boolean
    Dim sButInit As W_BUTINIT, button As WIDGET, scoreLabel As WIDGET, killsLabel As WIDGET, unitsLabel As WIDGET
    Dim lastColumnLabel As WIDGET, muteWidget As WIDGET
    Set d = MultiMenuGrid_(self)
    Set nameGrid = GridLayout_new()
    GridLayout_place nameGrid, grid_allocation_slot_(0, 1, False), grid_allocation_slot_(0), MultiMenuDroidView_new(player)
    ' char name[128]
    Set nameLabel = MultiMenuGrid_makeLabel(sstrcpy_cut(va("%d: %s", NetPlay_players_position(player), getPlayerName(player)), 127))
    GridLayout_place nameGrid, grid_allocation_slot_(1), grid_allocation_slot_(0), nameLabel

    Set alliancesGrid = GridLayout_new()

    selectedPlayerIsSpectator = NetPlay_players_isSpectator(selectedPlayer)

    ' add channel opener.
    ' - not configurable for self (obviously)
    ' - players shouldn't configure channel to spectators
    If (player <> selectedPlayer) And (selectedPlayerIsSpectator Or Not NetPlay_players_isSpectator(player)) Then
        W_BUTINIT_construct sButInit
        sButInit.width = 35
        sButInit.height = 24
        sButInit.id = MULTIMENU_CHANNEL + player
        sButInit.pTip = "Channel"
        Set sButInit.pDisplay = WzFunc_new(FN_displayChannelState)
        sButInit.UserData = player

        Set button = W_BUTTON_new_init(sButInit)        ' channelButton
        W_BUTTON_addOnClickHandler button, WzFunc_new(FN_multimenu_channelOnClick, player)
        GridLayout_place alliancesGrid, grid_allocation_slot_(0), grid_allocation_slot_(0), Margin_wrap(0, 1, 0, 0, button)
    End If

    If alliancesCanGiveAnything(game.alliance) And player <> selectedPlayer And ULT(player, MAX_PLAYERS) And Not NetPlay_players_isSpectator(player) Then
        W_BUTINIT_construct sButInit
        'alliance
        sButInit.id = MULTIMENU_ALLIANCE_BASE + player
        sButInit.pTip = "Toggle Alliance State"
        Set sButInit.pDisplay = WzFunc_new(FN_displayAllianceState)
        sButInit.UserData = player
        sButInit.width = 35
        sButInit.height = 24

        'can't break alliances in 'Locked Teams' mode
        If Not alliancesFixed(game.alliance) Then
            Set button = W_BUTTON_new_init(sButInit)    ' allianceButton
            W_BUTTON_addOnClickHandler button, WzFunc_new(FN_multimenu_allianceOnClick, player)
            GridLayout_place alliancesGrid, grid_allocation_slot_(1), grid_allocation_slot_(0), Margin_wrap(0, 1, 0, 0, button)
        End If

        Set sButInit.pDisplay = WzFunc_new(FN_intDisplayImageHilight)

        ' add the gift buttons.
        sButInit.width = 31
        sButInit.height = 21

        If alliancesCanGiveResearchAndRadar(game.alliance) Then
            sButInit.id = MULTIMENU_GIFT_RAD + player
            sButInit.pTip = "Give Visibility Report"
            sButInit.UserData = PACKDWORD_TRI(0, IMAGE_MULTI_VIS_HI, IMAGE_MULTI_VIS)
            Set button = W_BUTTON_new_init(sButInit)    ' visGiftButton
            W_BUTTON_addOnClickHandler button, WzFunc_new(FN_multimenu_giftOnClick, RADAR_GIFT, player)
            GridLayout_place alliancesGrid, grid_allocation_slot_(2), grid_allocation_slot_(0), MultiMenuGrid_wrapGift(button)

            sButInit.id = MULTIMENU_GIFT_RES + player
            sButInit.pTip = "Leak Technology Documents"
            sButInit.UserData = PACKDWORD_TRI(0, IMAGE_MULTI_TEK_HI, IMAGE_MULTI_TEK)
            Set button = W_BUTTON_new_init(sButInit)    ' resGiftButton
            W_BUTTON_addOnClickHandler button, WzFunc_new(FN_multimenu_giftOnClick, RESEARCH_GIFT, player)
            GridLayout_place alliancesGrid, grid_allocation_slot_(3), grid_allocation_slot_(0), MultiMenuGrid_wrapGift(button)
        End If

        sButInit.id = MULTIMENU_GIFT_DRO + player
        sButInit.pTip = "Hand Over Selected Units"
        sButInit.UserData = PACKDWORD_TRI(0, IMAGE_MULTI_DRO_HI, IMAGE_MULTI_DRO)
        Set button = W_BUTTON_new_init(sButInit)        ' droidGiftButton
        W_BUTTON_addOnClickHandler button, WzFunc_new(FN_multimenu_giftOnClick, DROID_GIFT, player)
        GridLayout_place alliancesGrid, grid_allocation_slot_(4), grid_allocation_slot_(0), MultiMenuGrid_wrapGift(button)

        sButInit.id = MULTIMENU_GIFT_POW + player
        sButInit.pTip = "Give Power To Player"
        sButInit.UserData = PACKDWORD_TRI(0, IMAGE_MULTI_POW_HI, IMAGE_MULTI_POW)
        Set button = W_BUTTON_new_init(sButInit)        ' pwrGiftButton
        W_BUTTON_addOnClickHandler button, WzFunc_new(FN_multimenu_giftOnClick, POWER_GIFT, player)
        GridLayout_place alliancesGrid, grid_allocation_slot_(5), grid_allocation_slot_(0), MultiMenuGrid_wrapGift(button)

        giftsUp(player) = True                          ' note buttons are up!
    End If

    Set scoreLabel = MultiMenuGrid_makeLabel("")
    W_LABEL_setTextAlignment scoreLabel, WLAB_ALIGNRIGHT
    Set killsLabel = MultiMenuGrid_makeLabel("")
    W_LABEL_setTextAlignment killsLabel, WLAB_ALIGNRIGHT
    Set unitsLabel = MultiMenuGrid_makeLabel("")
    W_LABEL_setTextAlignment unitsLabel, WLAB_ALIGNRIGHT
    Set lastColumnLabel = MultiMenuGrid_makeLabel("")
    W_LABEL_setTextAlignment lastColumnLabel, WLAB_ALIGNRIGHT
    ' playersWidgets.push_back({player, nameLabel, scoreLabel, killsLabel, unitsLabel, lastColumnLabel})
    d.playersWidgets.Add Array(player, nameLabel, scoreLabel, killsLabel, unitsLabel, lastColumnLabel)

    ' chat mute button
    If player <> selectedPlayer Then
        W_BUTINIT_construct sButInit
        sButInit.id = MULTIMENU_PLAYER_MUTE_START + player
        sButInit.pTip = "Toggle Chat Mute"
        Set sButInit.pDisplay = WzFunc_new(FN_displayMuteState)
        sButInit.UserData = player
        sButInit.width = 18
        sButInit.height = 18
        Set button = W_BUTTON_new_init(sButInit)        ' muteButton
        W_BUTTON_addOnClickHandler button, WzFunc_new(FN_multimenu_muteOnClick)
        Set muteWidget = button
    Else
        Set muteWidget = WIDGET_new()
        muteWidget.setGeometry 0, 0, 0, 0
    End If

    ' Margin margin(0, CELL_HORIZONTAL_PADDING); (0, 8, 0, 8)
    GridLayout_place self, grid_allocation_slot_(0), grid_allocation_slot_(row), Margin_wrap(0, CELL_HORIZONTAL_PADDING, 0, 0, nameGrid)
    GridLayout_place self, grid_allocation_slot_(1), grid_allocation_slot_(row), Margin_wrap(0, CELL_HORIZONTAL_PADDING, 0, CELL_HORIZONTAL_PADDING, Alignment_wrap(Alignment_center(), alliancesGrid))
    GridLayout_place self, grid_allocation_slot_(2), grid_allocation_slot_(row), Margin_wrap(0, CELL_HORIZONTAL_PADDING, 0, CELL_HORIZONTAL_PADDING, scoreLabel)
    GridLayout_place self, grid_allocation_slot_(3), grid_allocation_slot_(row), Margin_wrap(0, CELL_HORIZONTAL_PADDING, 0, CELL_HORIZONTAL_PADDING, killsLabel)
    GridLayout_place self, grid_allocation_slot_(4), grid_allocation_slot_(row), Margin_wrap(0, CELL_HORIZONTAL_PADDING, 0, CELL_HORIZONTAL_PADDING, unitsLabel)
    GridLayout_place self, grid_allocation_slot_(5), grid_allocation_slot_(row), Margin_wrap(0, CELL_HORIZONTAL_PADDING, 0, CELL_HORIZONTAL_PADDING, lastColumnLabel)
    GridLayout_place self, grid_allocation_slot_(6), grid_allocation_slot_(row), Margin_wrap(0, CELL_HORIZONTAL_PADDING, 0, CELL_HORIZONTAL_PADDING, Alignment_wrap(Alignment_center(), muteWidget))
End Sub

Private Sub MultiMenuGrid_updatePlayersWidgets(ByVal self As WIDGET)
    Dim d As MultiMenuGrid, pw As Variant, player As Long, color As Long, isHuman As Boolean, isAlly As Boolean
    Dim isSelectedPlayer As Boolean, scoreString As String, killsString As String, unitsString As String, lastString As String
    Dim lockedScreen As W_SCREEN, num As Long, label As WIDGET, i As Long
    Set d = MultiMenuGrid_(self)
    For Each pw In d.playersWidgets
        player = pw(0)
        If ULT(selectedPlayer, MAX_PLAYERS) And ULT(player, MAX_PLAYERS) Then
            color = GetPlayerTextColor(alliances(selectedPlayer, player), player)
            isAlly = aiCheckAlliances(selectedPlayer, player)
        Else
            color = GetPlayerTextColor(0, player)
            isAlly = False
        End If
        isHuman = isHumanPlayer(player)
        isSelectedPlayer = player = selectedPlayer

        For i = 1 To 5                  ' name, score, kills, units, lastColumn ->setFontColour(color)
            Set label = pw(i)
            W_LABEL_setFontColour label, color
        Next

        If Not ULT(player, MAX_PLAYERS) Then
            GoTo continue_
        End If

        ' (char scoreString[128], killsString[20], unitsString[20] = {0}, lastString[20] = {0})
        unitsString = ""
        lastString = ""

        ' Let's use the real score for MP games
        If Cheated And NetPlay_bComms Then
            scoreString = "(cheated)"
        Else
            scoreString = CStr(getMultiStats(player).recentScore)
        End If
        killsString = CStr(getMultiStats(player).recentKills)

        'only show player's and allies' unit counts, and nobody elses.
        If isAlly Or debugMappingsAllowed() Then
            unitsString = CStr(getNumDroids(player) + getNumTransporterDroids(player))
        End If

        ' Display player power instead of number of played games
        ' and number of units instead of ping when in debug mode
        If debugMappingsAllowed() Then  'Won't pass this when in both release and multiplayer modes
            ' Player power
            lastString = CStr(getPower(player))
        ElseIf runningMultiplayer() Then
            ' ping
            If Not isSelectedPlayer And isHuman Then
                If ingame.PingTimes(player) < PING_LIMIT Then
                    lastString = CStr(ingame.PingTimes(player))
                Else
                    lastString = ChrW$(&H221E)  ' U+221E, the infinity sign
                End If
            End If
        Else
            ' Structures
            If isAlly Or debugMappingsAllowed() Then
                ' NOTE, This tallys up *all* the structures you have. Test out via 'start with no base'.
                num = apsStructLists(player).size
                lastString = CStr(num)
            End If
        End If

        Set label = pw(2): label.setString scoreString
        Set label = pw(3): label.setString killsString
        Set label = pw(4): label.setString unitsString
        Set label = pw(5): label.setString lastString

        Set lockedScreen = self.screenPointer

        If isHuman Or (game.typ = LEVEL_TYPE_SKIRMISH And player < game.maxPlayers) Then
            ' alliance
            ' manage buttons by showing or hiding them. gifts only in campaign,
            If alliancesCanGiveAnything(game.alliance) Then
                If isAlly And Not isSelectedPlayer And Not giftsUp(player) Then
                    If alliancesCanGiveResearchAndRadar(game.alliance) Then
                        widgReveal lockedScreen, MULTIMENU_GIFT_RAD + player
                        widgReveal lockedScreen, MULTIMENU_GIFT_RES + player
                    End If
                    widgReveal lockedScreen, MULTIMENU_GIFT_DRO + player
                    widgReveal lockedScreen, MULTIMENU_GIFT_POW + player
                    giftsUp(player) = True
                ElseIf Not isAlly And Not isSelectedPlayer And giftsUp(player) Then
                    If alliancesCanGiveResearchAndRadar(game.alliance) Then
                        widgHide lockedScreen, MULTIMENU_GIFT_RAD + player
                        widgHide lockedScreen, MULTIMENU_GIFT_RES + player
                    End If
                    widgHide lockedScreen, MULTIMENU_GIFT_DRO + player
                    widgHide lockedScreen, MULTIMENU_GIFT_POW + player
                    giftsUp(player) = False
                End If
            End If
        End If

        ' clean up widgets if player leaves while menu is up.
        If Not isHuman And Not (game.typ = LEVEL_TYPE_SKIRMISH And player < game.maxPlayers) Then
            If Not widgGetFromID(lockedScreen, MULTIMENU_CHANNEL + player) Is Nothing Then
                widgDelete lockedScreen, MULTIMENU_CHANNEL + player
            End If

            If Not widgGetFromID(lockedScreen, MULTIMENU_ALLIANCE_BASE + player) Is Nothing Then
                widgDelete lockedScreen, MULTIMENU_ALLIANCE_BASE + player
                widgDelete lockedScreen, MULTIMENU_GIFT_RAD + player
                widgDelete lockedScreen, MULTIMENU_GIFT_RES + player
                widgDelete lockedScreen, MULTIMENU_GIFT_DRO + player
                widgDelete lockedScreen, MULTIMENU_GIFT_POW + player
                giftsUp(player) = False
            End If
        End If
continue_:
    Next
End Sub

Public Sub MultiMenuGrid_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x0 As Long, y0 As Long, columns As grid_offsets, rows As grid_offsets, lines(0 To 6) As Vector4i, i As Long
    x0 = xOffset + self.x()
    y0 = yOffset + self.y()

    columns = GridLayout_getColumnOffsets(self)
    rows = GridLayout_getRowOffsets(self)
    lines(0) = glm_ivec4(x0, y0 + grid_offsets_get(rows, 1), x0 + self.width(), y0 + grid_offsets_get(rows, 1))
    For i = 1 To 6
        lines(i) = glm_ivec4(x0 + grid_offsets_get(columns, i), y0, x0 + grid_offsets_get(columns, i), y0 + self.height())
    Next

    iV_Lines VarPtr(lines(0)), 7, psPalette(WZCOL_BLACK)

    ' (#ifdef DEBUG: the traffic statistics under the grid; the release build has none)
End Sub

Public Sub MultiMenuGrid_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    Dim d As MultiMenuGrid
    Set d = MultiMenuGrid_(self)
    d.timerLabel.setString getAsciiTime(gameTime)
    MultiMenuGrid_updatePlayersWidgets self

    d.powerLabel.hide
    d.pingLabel.hide
    d.structsLabel.hide
    If debugMappingsAllowed() Then
        d.powerLabel.show
    ElseIf runningMultiplayer() Then
        d.pingLabel.show
    Else
        d.structsLabel.show
    End If
End Sub

' MARK: - WzMultiMenuTabs

' WzMultiMenuTabs(int value = -1) : MultichoiceWidget(value) { }
Public Sub WzMultiMenuTabs_construct(ByVal self As WIDGET, Optional ByVal value As Long = -1)
    MultichoiceWidget_construct self, value
    self.kind = K_WzMultiMenuTabs ' (VB) virtual calls from here on reach WzMultiMenuTabs's overrides
End Sub

' std::make_shared<WzMultiMenuTabs>(value)
Private Function WzMultiMenuTabs_new(Optional ByVal value As Long = -1) As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_WzMultiMenuTabs)
    WzMultiMenuTabs_construct w, value
    Set WzMultiMenuTabs_new = w
End Function

Public Sub WzMultiMenuTabs_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
End Sub

' MARK: - WzMultiWidget

' WzMultiWidget(bool openAnimate = true) : IntFormAnimated(openAnimate)
Public Sub WzMultiWidget_construct(ByVal self As WIDGET, Optional ByVal openAnimate As Boolean = True)
    IntFormAnimated_construct self, openAnimate
    self.kind = K_WzMultiWidget ' (VB) virtual calls from here on reach WzMultiWidget's overrides
    WzMultiWidget_(self).weakSerial = WeakWidget_register(self)    ' (VB) its std::weak_ptrs
End Sub

Public Sub WzMultiWidget_destruct(ByVal self As WIDGET)
    WeakWidget_unregister WzMultiWidget_(self).weakSerial
End Sub

' static std::shared_ptr<WzMultiWidget> make(bool openAnimate, std::function<void ()> closeButtonHandler)
Private Function WzMultiWidget_make(ByVal openAnimate As Boolean, ByVal closeButtonHandler As WzFunc) As WIDGET
    Dim result As WIDGET
    Set result = WIDGET_alloc(K_WzMultiWidget)
    WzMultiWidget_construct result, openAnimate
    WzMultiWidget_initialize result, closeButtonHandler
    Set WzMultiWidget_make = result
End Function

Public Sub WzMultiWidget_geometryChanged(ByVal self As WIDGET)
    Dim d As WzMultiWidget, panel As WIDGET, i As Long, panelY0 As Long, panelHeight As Long
    Set d = WzMultiWidget_(self)
    If Not d.closeButton Is Nothing Then
        d.closeButton.callCalcLayout
    End If
    If Not d.panelSwitcher Is Nothing Then
        d.panelSwitcher.callCalcLayout
    End If
    For i = 1 To d.panels.Count
        Set panel = d.panels(i)
        panelY0 = 0
        If Not d.panelSwitcher Is Nothing Then
            panelY0 = panelY0 + d.panelSwitcher.y() + d.panelSwitcher.height() + MULTIMENUFORM_INTERNAL_PADDING
        End If
        panelHeight = MIN_(panel.idealHeight(), self.height() - panelY0)
        panel.setGeometry 0, panelY0, self.width(), panelHeight
    Next
End Sub

Public Function WzMultiWidget_idealWidth(ByVal self As WIDGET) As Long
    Dim d As WzMultiWidget, panel As WIDGET, i As Long, maxPanelIdealWidth As Long
    Set d = WzMultiWidget_(self)
    maxPanelIdealWidth = 0
    For i = 1 To d.panels.Count
        Set panel = d.panels(i)
        maxPanelIdealWidth = MAX_(maxPanelIdealWidth, panel.idealWidth())
    Next
    WzMultiWidget_idealWidth = maxPanelIdealWidth
End Function

Public Function WzMultiWidget_idealHeight(ByVal self As WIDGET) As Long
    Dim d As WzMultiWidget, panel As WIDGET, i As Long, result As Long, maxPanelIdealHeight As Long
    Set d = WzMultiWidget_(self)
    result = 0
    If Not d.panelSwitcher Is Nothing Then
        result = result + MULTIMENUFORM_INTERNAL_PADDING + d.panelSwitcher.idealHeight() + MULTIMENUFORM_INTERNAL_PADDING
    End If
    maxPanelIdealHeight = 0
    For i = 1 To d.panels.Count
        Set panel = d.panels(i)
        maxPanelIdealHeight = MAX_(maxPanelIdealHeight, panel.idealHeight())
    Next
    result = result + maxPanelIdealHeight + MULTIMENUFORM_INTERNAL_PADDING
    WzMultiWidget_idealHeight = result
End Function

Private Sub WzMultiWidget_switchAttachedPanel(ByVal self As WIDGET, ByVal newPanel As WIDGET)
    Dim d As WzMultiWidget
    Set d = WzMultiWidget_(self)
    If Not d.currentlyAttachedPanel Is Nothing Then
        self.detach d.currentlyAttachedPanel
    End If
    Set d.currentlyAttachedPanel = newPanel
    self.attach newPanel, ChildZPos_Back
End Sub

Private Sub WzMultiWidget_initialize(ByVal self As WIDGET, ByVal closeButtonHandler As WzFunc)
    Dim d As WzMultiWidget, sButInit As W_BUTINIT, allowAIsForTeamStrategy As Boolean, teamStrategyView As WIDGET
    Set d = WzMultiWidget_(self)
    If Not closeButtonHandler Is Nothing Then
        ' Add the close button.
        W_BUTINIT_construct sButInit
        sButInit.id = MULTIMENU_CLOSE
        sButInit.pTip = "Close"
        Set sButInit.pDisplay = WzFunc_new(FN_intDisplayImageHilight)
        sButInit.UserData = PACKDWORD_TRI(0, IMAGE_CLOSEHILIGHT, IMAGE_CLOSE)
        Set d.closeButton = W_BUTTON_new_init(sButInit)
        W_BUTTON_addOnClickHandler d.closeButton, WzFunc_new(FN_multiWidget_closeOnClick, closeButtonHandler)
        self.attach d.closeButton
        d.closeButton.setCalcLayout WzFunc_new(FN_multiWidget_closeCalcLayout)
    End If

    Set d.multiMenuGrid = MultiMenuGrid_make()
    d.panels.Add d.multiMenuGrid

    allowAIsForTeamStrategy = Not NetPlay_bComms
    If gameHasTeamStrategyView(allowAIsForTeamStrategy) Then
        Set teamStrategyView = createTeamStrategyView(allowAIsForTeamStrategy)
        If Not teamStrategyView Is Nothing Then
            transformTeamStrategyViewMode teamStrategyView, True
            teamStrategyViewSetBackgroundColor teamStrategyView, pal_RGBA(25, 0, 110, 180)
            Set d.teamStrategyViewWrapped = Margin_wrap(0, 10, 0, 10, teamStrategyView)
            d.panels.Add d.teamStrategyViewWrapped
        End If
    End If

    If Not d.teamStrategyViewWrapped Is Nothing Then
        ' add tabs
        Set d.panelSwitcher = WzMultiMenuTabs_new(0)
        self.attach d.panelSwitcher
        MultibuttonWidget_setButtonAlignment d.panelSwitcher, MultibuttonWidget_ButtonAlignment_CENTER_ALIGN
        MultibuttonWidget_addButton d.panelSwitcher, 0, WzPanelTabButton_make("Players View")
        MultibuttonWidget_addButton d.panelSwitcher, 1, WzPanelTabButton_make("Team Strategy")
        MultibuttonWidget_choose d.panelSwitcher, 0
        MultibuttonWidget_addOnChooseHandler d.panelSwitcher, WzFunc_new(FN_multiWidget_panelSwitcher_onChoose)
        d.panelSwitcher.setCalcLayout WzFunc_new(FN_multiWidget_panelSwitcher_calcLayout)
    End If

    WzMultiWidget_switchAttachedPanel self, d.multiMenuGrid
End Sub

' std::shared_ptr<IntFormAnimated> intCreateMultiMenuForm(std::function<void ()> closeButtonHandler = nullptr)
Public Function intCreateMultiMenuForm(Optional ByVal closeButtonHandler As WzFunc = Nothing) As WIDGET
    Dim form_ As WIDGET
    Set form_ = WzMultiWidget_make(True, closeButtonHandler)
    Set intCreateMultiMenuForm = form_
End Function

Public Function intAddMultiMenu() As Boolean
    Dim form_ As WIDGET
    'check for already open.
    If Not widgGetFromID(psWScreen, MULTIMENU_FORM) Is Nothing Then
        intCloseMultiMenu
        intAddMultiMenu = True
        Exit Function
    End If

    If intMode <> INT_INTELMAP Then
        intResetScreen False
    End If

    Set form_ = intCreateMultiMenuForm(WzFunc_new(FN_intAddMultiMenu_closeHandler))
    form_.id = MULTIMENU_FORM
    form_.setCalcLayout WzFunc_new(FN_intAddMultiMenu_calcLayout)

    psWScreen.psForm.attach form_

    intShowPowerBar                     ' add power bar

    If intMode <> INT_INTELMAP Then
        intMode = INT_MULTIMENU         ' change interface mode.
    End If
    MultiMenuUp = True
    intAddMultiMenu = True
End Function

' ////////////////////////////////////////////////////////////////////////////
Public Sub intCloseMultiMenuNoAnim()
    widgDelete psWScreen, MULTIMENU_FORM
    If Not MultiMenuUp Then
        Exit Sub
    End If
    MultiMenuUp = False
    If intMode <> INT_INTELMAP Then
        intMode = INT_NORMAL
    End If
End Sub


' ////////////////////////////////////////////////////////////////////////////
Public Function intCloseMultiMenu() As Boolean
    Dim form_ As WIDGET
    If Not MultiMenuUp Then
        intCloseMultiMenu = True
        Exit Function
    End If

    ' Start the window close animation.
    Set form_ = widgGetFromID(psWScreen, MULTIMENU_FORM)     ' (IntFormAnimated *)
    If Not form_ Is Nothing Then
        IntFormAnimated_closeAnimateDelete form_
        MultiMenuUp = False
    End If

    If intMode <> INT_INTELMAP Then
        intMode = INT_NORMAL
    End If
    intCloseMultiMenu = True
End Function


'--------------------------------------------------------------------------
' (VB) the lambdas (docs/CONVENTIONS.md, "Widgets")
Public Function multimenu_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim psWidget As WIDGET, value As String, i As Long, psParent As WIDGET, strongParent As WIDGET, weakParent As Long
    Dim x As Long, w As Long, h As Long
    Select Case f.id
    Case FN_displayRequestOption
        displayRequestOption args(0), args(1), args(2)
    Case FN_displayNumPlayersBut
        displayNumPlayersBut args(0), args(1), args(2)
    Case FN_multimenu_intDisplayImageHilight
        intDisplayImageHilight args(0), args(1), args(2)

    Case FN_requestForm_calcLayout
        Set psWidget = args(0)
        psWidget.setGeometry M_REQUEST_X, M_REQUEST_Y, M_REQUEST_W, M_REQUEST_H

    Case FN_searchBox_onEditingStopped               ' [searchDir, fileExtension, mode, numPlayers](W_EDITBOX& widg)
        Set psWidget = args(0)
        value = psWidget.getString()
        If value = current_searchString Then
            Exit Function
        End If
        closeMultiRequester
        addMultiRequest f.cap(0), f.cap(1), f.cap(2), f.cap(3), value

    Case FN_initPUserData_DisplayNumPlayersButCache  ' []() -> void * { return new DisplayNumPlayersButCache(); }
        Set multimenu_c_fn = WzText_new("", font_regular)

    ' the in-game menu's
    Case FN_displayAllianceState
        displayAllianceState args(0), args(1), args(2)
    Case FN_displayChannelState
        displayChannelState args(0), args(1), args(2)
    Case FN_displayMuteState
        displayMuteState args(0), args(1), args(2)

    Case FN_multimenu_channelOnClick                ' [player](W_BUTTON&)
        i = f.cap(0)                    ' UBYTE i = (UBYTE)player;
        openchannels(i) = Not openchannels(i)

        If mouseDown(MOUSE_RMB) And NetPlay_isHost Then ' both buttons....
            ' Allow the host to kick the AI only in a MP game, or if they activated cheats in a skirmish game
            If (NetPlay_bComms Or Cheated) And (NetPlay_players_allocated(i) Or (Not NetPlay_players_allocated(i) And NetPlay_players_ai(i) <> AI_OPEN)) Then
                inputLoseFocus
                startKickVote i
                Exit Function
            End If
        End If

    Case FN_multimenu_allianceOnClick               ' [player](W_BUTTON& widg)
        i = f.cap(0)                    ' UBYTE i = (UBYTE)(player);
        ' (VB) a spectator's click (selectedPlayer >= MAX_PLAYERS) would read past alliances[] in C++
        If Not ULT(selectedPlayer, MAX_PLAYERS) Then Exit Function

        Select Case alliances(selectedPlayer, i)
        Case ALLIANCE_BROKEN
            requestAlliance selectedPlayer, i, True, True           ' request an alliance
        Case ALLIANCE_INVITATION
            formAlliance selectedPlayer, i, True, True, True        ' form an alliance
        Case ALLIANCE_REQUESTED
            breakAlliance selectedPlayer, i, True, True             ' break an alliance
        Case ALLIANCE_FORMED
            breakAlliance selectedPlayer, i, True, True             ' break an alliance
        Case Else
        End Select

    Case FN_multimenu_giftOnClick                   ' [player](W_BUTTON&) { sendGift(<the button's gift>, player); }
        sendGift f.cap(0), f.cap(1)

    Case FN_multimenu_muteOnClick                   ' [](W_BUTTON& button)
        Set psWidget = args(0)
        i = psWidget.UserData           ' auto playerIdx = button.UserData;
        If Not ULT(i, MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "playerIdx < MAX_CONNECTED_PLAYERS", "src/multimenu.cpp:1034", "addPlayerWidgets", "Invalid playerIdx: %u", i: Exit Function
        setPlayerMuted i, Not ingame.muteChat(i)

    Case FN_multiWidget_closeOnClick                ' [closeButtonHandler](W_BUTTON&) { closeButtonHandler(); }
        WzFunc_call f.cap(0)

    Case FN_multiWidget_closeCalcLayout             ' [](WIDGET *psButton)
        Set psWidget = args(0)
        Set psParent = psWidget.parent()
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/multimenu.cpp:1386", "initialize", "No parent": Exit Function
        psWidget.setGeometry psParent.width() - CLOSE_WIDTH, 0, CLOSE_WIDTH, CLOSE_HEIGHT

    Case FN_multiWidget_panelSwitcher_onChoose      ' [](MultibuttonWidget& widget, int newValue)
        Set psWidget = args(0)
        ' Switch actively-displayed panel
        ' (auto weakParent = std::weak_ptr<WIDGET>(widget.parent()): the form's serial, 0 for another parent, which the
        ' task's dynamic_pointer_cast would refuse)
        Set psParent = psWidget.parent()
        weakParent = 0
        If WIDGET_isKindOf(psParent, K_WzMultiWidget) Then weakParent = WzMultiWidget_(psParent).weakSerial
        widgScheduleTask WzFunc_new(FN_multiWidget_panelSwitcher_task, args(1), weakParent)

    Case FN_multiWidget_panelSwitcher_task          ' [newValue, weakParent]()
        Set strongParent = WeakWidget_lock(f.cap(1))
        If Not (Not strongParent Is Nothing) Then ASSERT_FAILURE "strongParent != nullptr", "src/multimenu.cpp:1421", "initialize", "Parent doesn't exist?": Exit Function
        Set psWidget = WzMultiWidget_(strongParent).panels(CLng(f.cap(0)) + 1)     ' strongParent->panels[newValue]
        WzMultiWidget_switchAttachedPanel strongParent, psWidget

    Case FN_multiWidget_panelSwitcher_calcLayout    ' LAMBDA_CALCLAYOUT_SIMPLE
        Set psWidget = args(0)
        Set psParent = psWidget.parent()
        If Not WIDGET_isKindOf(psParent, K_WzMultiWidget) Then Set psParent = Nothing
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/multimenu.cpp:1427", "initialize", "No parent": Exit Function
        w = psWidget.idealWidth()       ' panelSwitcherWidth
        x = (psParent.width() - w) \ 2  ' panelSwitcherX0
        psWidget.setGeometry x, MULTIMENUFORM_INTERNAL_PADDING, w, MULTIMENUFORM_PANEL_TABS_HEIGHT

    Case FN_intAddMultiMenu_closeHandler            ' []() { widgScheduleTask([]() { intCloseMultiMenu(); }); }
        widgScheduleTask WzFunc_new(FN_intAddMultiMenu_closeTask)

    Case FN_intAddMultiMenu_closeTask               ' []() { intCloseMultiMenu(); }
        intCloseMultiMenu

    Case FN_intAddMultiMenu_calcLayout              ' [](WIDGET *form)
        Set psWidget = args(0)
        w = MIN_(screenWidth - 20, psWidget.idealWidth())
        h = psWidget.idealHeight()
        psWidget.setGeometry (screenWidth - w) \ 2, MULTIMENU_FORM_Y, w, h
    End Select
End Function
