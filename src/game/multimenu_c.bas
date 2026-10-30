Attribute VB_Name = "multimenu_c"
' src/multimenu.cpp, multimenu.h: the request lists of the lobby (the map
' chooser with its player-count buttons and search box, the player name
' chooser: addMultiRequest, runMultiRequester, closeMultiRequester,
' displayRequestOption, displayNumPlayersBut, stringRelevance) and their
' globals (psRScreen, MultiMenuUp, current_numplayers, multiRequestUp).
' Not ported yet: the in-game multiplayer menu (GetPlayerTextColor, the
' alliance / channel / mute displays, MultiMenuDroidView, MultiMenuGrid,
' WzMultiMenuTabs, WzMultiWidget, intCreateMultiMenuForm, intAddMultiMenu,
' intCloseMultiMenu, intCloseMultiMenuNoAnim): cluster G, stand-ins.
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

' #define MULTIMENU_FORM_Y 25 ... (the in-game menu's ids: cluster G)

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

' the FN ids of this module's lambdas and function pointers (WzFunc)
Public Const FN_displayRequestOption As Long = 5100
Public Const FN_displayNumPlayersBut As Long = 5101
Public Const FN_requestForm_calcLayout As Long = 5102         ' LAMBDA_CALCLAYOUT_SIMPLE
Public Const FN_searchBox_onEditingStopped As Long = 5103     ' [searchDir, fileExtension, mode, numPlayers](W_EDITBOX& widg)
Public Const FN_initPUserData_DisplayNumPlayersButCache As Long = 5104 ' []() -> void * { return new DisplayNumPlayersButCache(); }
Public Const FN_multimenu_intDisplayImageHilight As Long = 5105 ' (VB) intdisplay.cpp's intDisplayImageHilight (cluster G): f_F2's stand-in

' (VB) the C++ initialisers of this module's globals
Public Sub multimenu_c_init()
    current_numplayers = 4
End Sub

Private Function M_REQUEST_NP_at(ByVal i As Long) As Long
    M_REQUEST_NP_at = M_REQUEST_2P + i
End Function

' ////////////////////////////////////////////////////////////////////////////
' Map / force / name load save stuff.

' (static PIELIGHT GetPlayerTextColor(int mode, UDWORD player): the in-game menu's, cluster G)

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
' Multiplayer in game menu stuff (cluster G)

'--------------------------------------------------------------------------
' (VB) the lambdas (docs/CONVENTIONS.md, "Widgets")
Public Function multimenu_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim psWidget As WIDGET, value As String
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
    End Select
End Function
