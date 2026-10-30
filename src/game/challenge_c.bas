Attribute VB_Name = "challenge_c"
' src/challenge.cpp, challenge.h: the challenges' menu and their high scores. The slots' pUserData is a
' DisplayLoadSlotData (the C++'s DisplayLoadSlotData with its DisplayLoadSlotCache).
'
' Run challenges dialog.
Option Explicit

Public Const CHALLENGE_SCORES As String = "scores.json"

Private Const totalslots As Long = 36           ' challenge slots
Private Const slotsInColumn As Long = 12        ' # of slots in a column
Private Const totalslotspace As Long = 256      ' max chars for slot strings.

Private Const CHALLENGE_W As Long = 610
Private Const CHALLENGE_H As Long = 215

Private Const CHALLENGE_HGAP As Long = 9
Private Const CHALLENGE_VGAP As Long = 9
Private Const CHALLENGE_BANNER_DEPTH As Long = 40   ' top banner which displays either load or save

Private Const CHALLENGE_ENTRY_W As Long = ((CHALLENGE_W \ 3) - (3 * CHALLENGE_HGAP))
Private Const CHALLENGE_ENTRY_H As Long = (CHALLENGE_H - (5 * CHALLENGE_VGAP) - (CHALLENGE_BANNER_DEPTH + CHALLENGE_VGAP)) \ 5

Private Const ID_LOADSAVE As Long = 21000
Private Const CHALLENGE_FORM As Long = ID_LOADSAVE + 1      ' back form.
Private Const CHALLENGE_CANCEL As Long = ID_LOADSAVE + 2    ' cancel but.
Private Const CHALLENGE_LABEL As Long = ID_LOADSAVE + 3     ' load/save
Private Const CHALLENGE_BANNER As Long = ID_LOADSAVE + 4    ' banner.

Private Const CHALLENGE_ENTRY_START As Long = ID_LOADSAVE + 10                  ' each of the buttons.
Private Const CHALLENGE_ENTRY_END As Long = ID_LOADSAVE + 10 + totalslots       ' must have unique ID hmm -Q

' the lambdas (WzFunc: challenge_c_fn)
Public Const FN_challenge_displayLoadBanner As Long = 14000     ' (widget, xOffset, yOffset)
Public Const FN_challenge_displayLoadSlot As Long = 14001       ' (widget, xOffset, yOffset)
Public Const FN_challengeForm_calcLayout As Long = 14002        ' (widget)
Public Const FN_newDisplayLoadSlotData As Long = 14003          ' []() -> void * { return new DisplayLoadSlotData(); }
Public Const FN_deleteDisplayLoadSlotData As Long = 14004       ' (widget): the slots' onDelete

Private psRequestScreen As W_SCREEN     ' static std::shared_ptr<W_SCREEN> psRequestScreen = nullptr; // Widget screen for requester

Public challengesUp As Boolean          ' True when interface is up and should be run.
Public challengeActive As Boolean       ' Whether we are running a challenge
Public challengeName As String          ' std::string
Public challengeFileName As String      ' WzString

' addChallenges' static char sSlotCaps[totalslots][totalslotspace], sSlotTips, sSlotFile
Private sSlotCaps(0 To totalslots - 1) As String
Private sSlotTips(0 To totalslots - 1) As String
Private sSlotFile(0 To totalslots - 1) As String

' #define CHALLENGE_X D_W + 16
Private Function CHALLENGE_X() As Long
    CHALLENGE_X = D_W() + 16
End Function

' #define CHALLENGE_Y D_H + 5
Private Function CHALLENGE_Y() As Long
    CHALLENGE_Y = D_H() + 5
End Function

Private Sub displayLoadBanner(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim col As Long, x As Long, y As Long
    col = psPalette(WZCOL_GREEN)
    x = xOffset + psWidget.x()
    y = yOffset + psWidget.y()

    pie_BoxFill x, y, x + psWidget.width(), y + psWidget.height(), col
    pie_BoxFill x + 2, y + 2, x + psWidget.width() - 2, y + psWidget.height() - 2, psPalette(WZCOL_MENU_BACKGROUND)
End Sub

' const char* currentChallengeName(): (VB) vbNullString = nullptr
Public Function currentChallengeName() As String
    If challengeActive Then
        currentChallengeName = challengeName
        Exit Function
    End If
    currentChallengeName = vbNullString
End Function

Public Sub updateChallenge(ByVal gameWon As Boolean)
    Dim fullName As String, seconds As Long, newtime As Long, victory As Boolean, scores As WzConfig
    Dim fName As String, pos As Long, sPath As String
    fullName = challengeFileName
    seconds = 0
    newtime = FromU32(Fix(U32(gameTime - mission_startTime) / GAME_TICKS_PER_SEC))
    victory = False
    Set scores = WzConfig_new(CHALLENGE_SCORES, WzConfig_ReadAndWrite)
    If Not (Len(fullName) > 0) Then ASSERT_FAILURE "fullName.length() > 0", "src/challenge.cpp:107", "updateChallenge", "Empty challengeFileName": Exit Sub

    pos = InStr(1, fullName, "/", vbBinaryCompare)
    If pos <> 0 Then
        fName = Mid$(fullName, pos + 1)
    Else
        fName = fullName ' Likely if challengeFileName no longer has the path included in it.
    End If
    If Len(fName) = 0 Then
        debug_ 121, LOG_ERROR, "updateChallenge", "Bad path to challenge file (%s)", fName
        Exit Sub
    End If
    sPath = fName
    ' remove .json
    If Right$(sPath, 5) = ".json" Then
        sPath = Left$(sPath, Len(sPath) - 5)
    End If
    scores.beginGroup sPath
    victory = scores.value("victory", json_variant_new(False)).toBool()
    seconds = scores.value("seconds", json_variant_new(0)).toInt()

    ' Update score if we have a victory and best recorded was a loss,
    ' or both were losses but time is higher, or both were victories
    ' but time is lower.
    If (Not victory And gameWon) _
       Or (Not gameWon And Not victory And newtime > seconds) _
       Or (gameWon And victory And newtime < seconds) Then
        scores.setValue "seconds", json_value(newtime)
        scores.setValue "victory", json_value(gameWon)
        scores.setValue "player", json_value(NetPlay_players_name(selectedPlayer))
    End If
    scores.endGroup
End Sub

' ////////////////////////////////////////////////////////////////////////////

Private Sub displayLoadSlot(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim data As DisplayLoadSlotData, x As Long, y As Long, butString As String
    ' Any widget using displayLoadSlot must have its pUserData initialized to a (DisplayLoadSlotData*)
    ' (assert(psWidget->pUserData != nullptr): the C assert, off in a release build)
    Set data = psWidget.pUserData

    x = xOffset + psWidget.x()
    y = yOffset + psWidget.y()
    ' (char butString[64])

    drawBlueBox x, y, psWidget.width(), psWidget.height()   ' draw box

    If Len(W_BUTTON_(psWidget).pText) <> 0 Then
        butString = Left$(W_BUTTON_(psWidget).pText, 63)    ' sstrcpy(butString, ...)

        If data.cache_fullText <> butString Then
            ' Update cache
            Do While iV_GetTextWidth(butString, font_regular) > psWidget.width()
                butString = Left$(butString, Len(butString) - 1)
            Loop
            data.cache_wzText.setText butString, font_regular
            data.cache_fullText = butString
        End If

        data.cache_wzText.render x + 4, y + 17, psPalette(WZCOL_FORM_TEXT)
    End If
End Sub

Public Sub challengesScreenSizeDidChange(ByVal oldWidth As Long, ByVal oldHeight As Long, ByVal newWidth As Long, ByVal newHeight As Long)
    If psRequestScreen Is Nothing Then Exit Sub
    psRequestScreen.screenSizeDidChange oldWidth, oldHeight, newWidth, newHeight
End Sub

'****************************************************************************************
' Challenge menu
'*****************************************************************************************
Public Function addChallenges() As Boolean
    Dim sPath As String, sSearchPath As String, slotCount As Long, parent As WIDGET, challengeForm As WIDGET
    Dim sFormInit As W_FORMINIT, sLabInit As W_LABINIT, sButInit As W_BUTINIT
    Dim files() As String, nFiles As Long, k As Long, i As String
    Dim button As WIDGET, name_ As String, map As String, difficulty As String, highscore As String, description As String
    Dim victory As Boolean, seconds As Long, scores As WzConfig, challenge As WzConfig, psTimeText As String, tmp As tm

    sSearchPath = "challenges"

    Set psRequestScreen = W_SCREEN_make() ' init the screen

    Set parent = psRequestScreen.psForm

    ' add a form to place the tabbed form on
    Set challengeForm = IntFormAnimated_new()

    challengeForm.id = CHALLENGE_FORM
    parent.attach challengeForm
    challengeForm.setCalcLayout WzFunc_new(FN_challengeForm_calcLayout)

    ' Add Banner
    W_FORMINIT_construct sFormInit
    sFormInit.formID = CHALLENGE_FORM
    sFormInit.id = CHALLENGE_BANNER
    sFormInit.style = WFORM_PLAIN
    sFormInit.x = CHALLENGE_HGAP
    sFormInit.y = CHALLENGE_VGAP
    sFormInit.width = CHALLENGE_W - (2 * CHALLENGE_HGAP)
    sFormInit.height = CHALLENGE_BANNER_DEPTH
    Set sFormInit.pDisplay = WzFunc_new(FN_challenge_displayLoadBanner)
    sFormInit.UserData = 0
    widgAddForm psRequestScreen, sFormInit

    ' Add Banner Label
    W_LABINIT_construct sLabInit
    sLabInit.formID = CHALLENGE_BANNER
    sLabInit.FontID = font_large
    sLabInit.id = CHALLENGE_LABEL
    sLabInit.style = WLAB_ALIGNCENTRE
    sLabInit.x = 0
    sLabInit.y = 0
    sLabInit.width = CHALLENGE_W - (2 * CHALLENGE_HGAP)     ' CHALLENGE_W;
    sLabInit.height = CHALLENGE_BANNER_DEPTH                ' This looks right -Q
    sLabInit.pText = "Challenge"
    widgAddLabel psRequestScreen, sLabInit

    ' add cancel.
    W_BUTINIT_construct sButInit
    sButInit.formID = CHALLENGE_BANNER
    sButInit.x = 8
    sButInit.y = 8
    sButInit.width = iV_GetImageWidth(IntImages, IMAGE_NRUTER)
    sButInit.height = iV_GetImageHeight(IntImages, IMAGE_NRUTER)
    sButInit.UserData = PACKDWORD_TRI(0, IMAGE_NRUTER, IMAGE_NRUTER)

    sButInit.id = CHALLENGE_CANCEL
    sButInit.pTip = "Close"
    Set sButInit.pDisplay = WzFunc_new(FN_intDisplayImageHilight)
    widgAddButton psRequestScreen, sButInit

    ' add slots
    W_BUTINIT_construct sButInit                ' sButInit = W_BUTINIT();
    sButInit.formID = CHALLENGE_FORM
    sButInit.width = CHALLENGE_ENTRY_W
    sButInit.height = CHALLENGE_ENTRY_H
    Set sButInit.pDisplay = WzFunc_new(FN_challenge_displayLoadSlot)
    Set sButInit.initPUserDataFunc = WzFunc_new(FN_newDisplayLoadSlotData)
    Set sButInit.onDelete = WzFunc_new(FN_deleteDisplayLoadSlotData)

    For slotCount = 0 To totalslots - 1
        sButInit.id = slotCount + CHALLENGE_ENTRY_START

        If slotCount < slotsInColumn Then
            sButInit.x = 22 + CHALLENGE_HGAP
            sButInit.y = ToInt16((CHALLENGE_BANNER_DEPTH + (2 * CHALLENGE_VGAP)) + ( _
                                  slotCount * (CHALLENGE_VGAP + CHALLENGE_ENTRY_H)))
        ElseIf slotCount >= slotsInColumn And (slotCount < (slotsInColumn * 2)) Then
            sButInit.x = 22 + (2 * CHALLENGE_HGAP + CHALLENGE_ENTRY_W)
            sButInit.y = ToInt16((CHALLENGE_BANNER_DEPTH + (2 * CHALLENGE_VGAP)) + ( _
                                  (slotCount Mod slotsInColumn) * (CHALLENGE_VGAP + CHALLENGE_ENTRY_H)))
        Else
            sButInit.x = 22 + (3 * CHALLENGE_HGAP + (2 * CHALLENGE_ENTRY_W))
            sButInit.y = ToInt16((CHALLENGE_BANNER_DEPTH + (2 * CHALLENGE_VGAP)) + ( _
                                  (slotCount Mod slotsInColumn) * (CHALLENGE_VGAP + CHALLENGE_ENTRY_H)))
        End If
        widgAddButton psRequestScreen, sButInit
    Next

    ' fill slots.
    slotCount = 0

    sPath = sSearchPath
    sPath = sPath & "/*.json"

    debug_ 313, LOG_SAVE, "addChallenges", "Searching ""%s"" for challenges", sPath

    ' add challenges to buttons
    ' (WZ_PHYSFS_enumerateFiles(sSearchPath, [&](const char *i) -> bool {...}): the files, then the lambda's body)
    If WZ_PHYSFS_enumerateFiles(sSearchPath, files, nFiles) Then
        For k = 0 To nFiles - 1
            i = files(k)
            ' See if this filename contains the extension we're looking for
            If InStr(1, i, ".json", vbBinaryCompare) = 0 Then
                ' If it doesn't, move on to the next filename
                GoTo continue_                  ' return true; // continue;
            End If

            ' First grab any high score associated with this challenge
            sPath = Left$(i, Len(i) - 5)        ' sstrcpy(sPath, i); sPath[strlen(sPath) - 5] = '\0';	// remove .json
            highscore = "no score"
            Set scores = WzConfig_new(CHALLENGE_SCORES, WzConfig_ReadOnly)
            scores.beginGroup sPath
            name_ = scores.value("player", json_variant_new("NO NAME")).toWzString()
            victory = scores.value("victory", json_variant_new(False)).toBool()
            seconds = scores.value("seconds", json_variant_new(-1)).toInt()
            If seconds > 0 Then
                ' (char psTimeText[sizeof("HH:MM:SS")] = { 0 };)
                tmp = getUtcTime(seconds)
                psTimeText = strftime("%H:%M:%S", tmp)
                If victory Then
                    highscore = psTimeText & " by " & name_ & " (" & "Victory" & ")"
                Else
                    highscore = psTimeText & " by " & name_ & " (" & "Survived" & ")"
                End If
            End If
            scores.endGroup
            Set scores = Nothing
            sPath = sSearchPath & "/" & i       ' ssprintf(sPath, "%s/%s", sSearchPath, i);
            Set challenge = WzConfig_new(sPath, WzConfig_ReadOnlyAndRequired)
            If Not (challenge.contains("challenge")) Then ASSERT_FAILURE "challenge.contains(""challenge"")", "src/challenge.cpp:348", "addChallenges", "Invalid challenge file %s - no challenge section!", sPath
            challenge.beginGroup "challenge"
            If Not (challenge.contains("name")) Then ASSERT_FAILURE "challenge.contains(""name"")", "src/challenge.cpp:350", "addChallenges", "Invalid challenge file %s - no name", sPath
            name_ = challenge.value("name", json_variant_new("BAD NAME")).toWzString()
            If Not (challenge.contains("map")) Then ASSERT_FAILURE "challenge.contains(""map"")", "src/challenge.cpp:352", "addChallenges", "Invalid challenge file %s - no map", sPath
            map = challenge.value("map", json_variant_new("BAD MAP")).toWzString()
            difficulty = challenge.value("difficulty", json_variant_new("BAD DIFFICULTY")).toWzString()
            description = map & ", " & difficulty & ", " & highscore & "." & vbLf & challenge.value("description", json_variant_new("")).toWzString()

            Set button = widgGetFromID(psRequestScreen, CHALLENGE_ENTRY_START + slotCount)

            debug_ 359, LOG_SAVE, "addChallenges", "We found [%s]", i

            ' Set the button-text
            sSlotCaps(slotCount) = Left$(name_, totalslotspace - 1)         ' store it!
            sSlotTips(slotCount) = Left$(description, totalslotspace - 1)  ' store it, too!
            sSlotFile(slotCount) = Left$(sPath, totalslotspace - 1)        ' store filename

            ' Add button
            W_BUTTON_(button).pTip = sSlotTips(slotCount)
            W_BUTTON_(button).pText = sSlotCaps(slotCount)
            ' (assert(button->pUserData != nullptr): the C assert, off in a release build)
            button.pUserData.filename = sSlotFile(slotCount)
            slotCount = slotCount + 1   ' go to next button...
            If slotCount = totalslots Then
                Exit For                ' return false; // break;
            End If
            challenge.endGroup
            Set challenge = Nothing
continue_:
        Next
    End If

    challengesUp = True

    addChallenges = True
End Function

' ////////////////////////////////////////////////////////////////////////////
Public Function closeChallenges() As Boolean
    Set psRequestScreen = Nothing
    ' need to "eat" up the return key so it don't pass back to game.
    inputLoseFocus
    challengesUp = False
    closeChallenges = True
End Function

' ////////////////////////////////////////////////////////////////////////////
' Returns true if cancel pressed or a valid game slot was selected.
' if when returning true strlen(sRequestResult) != 0 then a valid game
' slot was selected otherwise cancel was selected..
Public Function runChallenges() As Boolean
    Dim triggers() As WidgetTrigger, nTriggers As Long, t As Long, id As Long, psWidget As WIDGET, data As DisplayLoadSlotData
    nTriggers = widgRunScreen(psRequestScreen, triggers)
    For t = 0 To nTriggers - 1
        id = triggers(t).widget.id

        sRequestResult = ""  ' set returned filename to null;

        ' cancel this operation...
        If id = CHALLENGE_CANCEL Then
            GoTo failure
        ElseIf CancelPressed() Then             ' (||: CancelPressed only when not the cancel button)
            GoTo failure
        End If

        ' clicked a load entry
        If id >= CHALLENGE_ENTRY_START And id <= CHALLENGE_ENTRY_END Then
            Set psWidget = widgGetFromID(psRequestScreen, id)
            ' (assert(psWidget != nullptr): the C assert, off in a release build)
            If Not (Len(W_BUTTON_(psWidget).pText) = 0) Then
                Set data = psWidget.pUserData
                ' (assert(data != nullptr); assert(data->filename != nullptr): the C asserts)
                sRequestResult = data.filename
                challengeFileName = sRequestResult
                challengeName = W_BUTTON_(psWidget).pText
            Else
                GoTo failure  ' clicked on an empty box
            End If
            GoTo success
        End If
    Next

    runChallenges = False
    Exit Function

' failed and/or cancelled..
failure:
    closeChallenges
    challengeActive = False
    runChallenges = False
    Exit Function

' success on load.
success:
    closeChallenges
    challengeActive = True
    ingame.side = InGameSide_HOST_OR_SINGLEPLAYER
    changeTitleUI WzMultiplayerOptionsTitleUI_new(wzTitleUICurrent)
    runChallenges = True
End Function

' ////////////////////////////////////////////////////////////////////////////
' should be done when drawing the other widgets.
Public Function displayChallenges() As Boolean
    widgDisplayScreen psRequestScreen    ' display widgets.
    displayChallenges = True
End Function

' the lambdas
Public Function challenge_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim psWidget As WIDGET
    Select Case f.id
    Case FN_challenge_displayLoadBanner
        displayLoadBanner args(0), args(1), args(2)
    Case FN_challenge_displayLoadSlot
        displayLoadSlot args(0), args(1), args(2)
    Case FN_challengeForm_calcLayout
        Set psWidget = args(0)
        psWidget.setGeometry CHALLENGE_X(), CHALLENGE_Y(), CHALLENGE_W, (slotsInColumn * CHALLENGE_ENTRY_H + CHALLENGE_HGAP * slotsInColumn) + CHALLENGE_BANNER_DEPTH + 20
    Case FN_newDisplayLoadSlotData
        Set challenge_c_fn = New DisplayLoadSlotData
    Case FN_deleteDisplayLoadSlotData
        Set psWidget = args(0)
        ' (assert(psWidget->pUserData != nullptr): the C assert)
        psWidget.pUserData = Empty          ' delete static_cast<DisplayLoadSlotData *>(psWidget->pUserData); psWidget->pUserData = nullptr;
    End Select
End Function
