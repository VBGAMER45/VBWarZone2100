Attribute VB_Name = "ingameop_c"
' src/ingameop.cpp, ingameop.h: the in-game options menu (resume, guide,
' options, load, save, quit and its confirmation, the host's quit warning) and
' the popup when the host has left, whole. The text buttons are frontend_c's
' displayTextOption with a DisplayTextOptionCache (frontend_c's onDelete
' lambda drops it); the forms are intdisplay_c's IntFormAnimated. Load and
' save go to loadsave.cpp's addLoadSave (stand-in: cluster I), the options to
' ingameopscreen.cpp's showInGameOptionsScreen with optionsforms_c's
' createOptionsBrowser. The macros that use D_W / D_H (hci.h: the screen's
' size) are functions.
Option Explicit

' ////////////////////////////////////////////////////////////////////////////
' defines

' position info for window.

' game options

' initial options
Public Const INTINGAMEOP_W As Long = 150
Public Const INTINGAMEOP_H As Long = 148
Public Const INTINGAMEOP_HS As Long = 112

Public Const INTINGAMEOP2_W As Long = 350
Public Const INTINGAMEOP2_H As Long = 160

Public Const INTINGAMEOPLINE_H As Long = 17
Public Const INTINGAMEOPMARGIN_H As Long = 10

' Host quit confirmation.
Public Const INTINGAMEOP3_W As Long = 150
Public Const INTINGAMEOP3_H As Long = 65

'Singleplayer quit confirmation
Public Const INTINGAMEOP4_W As Long = 400

Public Const PAUSEMESSAGE_YOFFSET As Long = 0

' button sizes.
Public Const INTINGAMEOP_OP_W As Long = INTINGAMEOP_W - 10
Public Const INTINGAMEOP4_OP_W As Long = INTINGAMEOP4_W - 15
Public Const INTINGAMEOP_SW_W As Long = INTINGAMEOP2_W - 15
Public Const INTINGAMEOP_OP_H As Long = 15

Public Enum INTINGAMEOP_IDS
    INTINGAMEOP = 10500
    INTINGAMEPOPUP
    INTINGAMEOP_HOSTQUIT
    INTINGAMEOP_CONFIRM_QUIT
    INTINGAMEOP_QUIT                    ' < The all important quit button
    INTINGAMEOP_RESUME
    INTINGAMEOP_GO_BACK
    INTINGAMEOP_LOAD_MISSION
    INTINGAMEOP_LOAD_SKIRMISH
    INTINGAMEOP_SAVE_MISSION
    INTINGAMEOP_SAVE_SKIRMISH
    INTINGAMEOP_OPTIONS
    INTINGAMEOP_PAUSELABEL              ' < The paused message
    INTINGAMEOP_CENTRESCREEN
    INTINGAMEOP_REPLAY
    INTINGAMEOP_POPUP_QUIT
    INTINGAMEOP_POPUP_MSG1
    INTINGAMEOP_POPUP_MSG2
    INTINGAMEOP_OPENGAMEGUIDE
End Enum

' positions within option boxes.
Public Const INTINGAMEOP_1_X As Long = 5
Public Const INTINGAMEOP_2_X As Long = 10
Public Const INTINGAMEOP_MID As Long = 160
Public Const INTINGAMEOP_1_Y As Long = 20
Public Const INTINGAMEOP_2_Y As Long = 40

Public Const OPALIGN As Long = WBUT_PLAIN Or WBUT_TXTCENTRE

Public hostQuitConfirmation As Boolean  ' = true (ingameop_c_init)

Public InGameOpUp As Boolean            ' = false
Public isInGamePopupUp As Boolean       ' = false
Private isInGameConfirmQuitUp As Boolean ' = false

' (VB) the lambdas
Public Const FN_addHostQuitOptions_inGameOpLayout As Long = 11010
Public Const FN_addHostQuitOptions_inGamePopupLayout As Long = 11011
Public Const FN_intAddInGameOptions_layout As Long = 11012     ' [lines]
Public Const FN_startInGameConfirmQuit_layout As Long = 11013  ' [row]
Public Const FN_intAddInGamePopup_layout As Long = 11014
Public Const FN_newDisplayTextOptionCache As Long = 11015      ' []() -> void * { return new DisplayTextOptionCache(); } (also mission_c's)
Public Const FN_startIGOptionsMenu_onClose As Long = 11016
Public Const FN_showGuideScreen_noopOnClose As Long = 11017

' C++'s static initialisers
Public Sub ingameop_c_init()
    hostQuitConfirmation = True
End Sub

' (VB) ingameop.h's macros that depend on the screen's size
Private Function INTINGAMEOP_X() As Long
    INTINGAMEOP_X = ((320 - (INTINGAMEOP_W \ 2)) + D_W)
End Function

Private Function INTINGAMEOP3_X() As Long
    INTINGAMEOP3_X = ((320 - (INTINGAMEOP3_W \ 2)) + D_W)
End Function

Private Function INTINGAMEOP3_Y() As Long
    INTINGAMEOP3_Y = ((240 - (INTINGAMEOP3_H \ 2)) + D_H)
End Function

Private Function INTINGAMEOP4_X() As Long
    INTINGAMEOP4_X = ((320 - (INTINGAMEOP4_W \ 2)) + D_W)
End Function

Private Function INTINGAMEOPAUTO_H(ByVal nlines As Long) As Long
    INTINGAMEOPAUTO_H = ((nlines) * INTINGAMEOPLINE_H + INTINGAMEOPMARGIN_H * 2)
End Function

Private Function INTINGAMEOPAUTO_Y(ByVal nlines As Long) As Long
    INTINGAMEOPAUTO_Y = ((240 - (INTINGAMEOPAUTO_H(nlines) \ 2)) + D_H)
End Function

Private Function INTINGAMEOPAUTO_Y_LINE(ByVal line_ As Long) As Long
    INTINGAMEOPAUTO_Y_LINE = (((line_) - 1) * INTINGAMEOPLINE_H + INTINGAMEOPMARGIN_H)
End Function

' ////////////////////////////////////////////////////////////////////////////
' functions

' ////////////////////////////////////////////////////////////////////////////

Private Function addIGTextButton(ByVal id As Long, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal string_ As String, ByVal Style As Long) As Boolean
    Dim sButInit As W_BUTINIT
    W_BUTINIT_construct sButInit

    'resume
    sButInit.formID = INTINGAMEOP
    sButInit.id = id
    sButInit.style = Style


    sButInit.x = x
    sButInit.y = y
    sButInit.width = width
    sButInit.height = INTINGAMEOP_OP_H

    Set sButInit.pDisplay = WzFunc_new(FN_displayTextOption)
    sButInit.pText = string_
    Set sButInit.pUserData = New DisplayTextOptionCache
    Set sButInit.onDelete = WzFunc_new(FN_deleteDisplayTextOptionCache)
    widgAddButton psWScreen, sButInit

    addIGTextButton = True
End Function

Private Function addHostQuitOptions() As Boolean
    Dim parent As WIDGET, inGameOp As WIDGET, inGamePopup As WIDGET, label As WIDGET

    ' get rid of the old stuff.
    widgDelete psWScreen, INTINGAMEPOPUP
    widgDelete psWScreen, INTINGAMEOP

    Set parent = psWScreen.psForm

    ' add form
    Set inGameOp = IntFormAnimated_new()
    parent.attach inGameOp
    inGameOp.id = INTINGAMEOP
    inGameOp.setCalcLayout WzFunc_new(FN_addHostQuitOptions_inGameOpLayout)

    addIGTextButton INTINGAMEOP_RESUME, INTINGAMEOP_1_X, INTINGAMEOP_1_Y, INTINGAMEOP_OP_W, "Resume Game", OPALIGN

    Set inGamePopup = IntFormAnimated_new()
    parent.attach inGamePopup
    inGamePopup.id = INTINGAMEPOPUP
    inGamePopup.setCalcLayout WzFunc_new(FN_addHostQuitOptions_inGamePopupLayout)

    Set label = W_LABEL_new()
    inGamePopup.attach label
    label.setGeometry 0, 0, inGamePopup.width(), inGamePopup.height()
    label.setString "WARNING: You're the host. If you quit, the game ends for everyone!"
    W_LABEL_setTextAlignment label, WLAB_ALIGNCENTRE
    W_LABEL_setFont label, font_medium, psPalette(WZCOL_YELLOW)

    addIGTextButton INTINGAMEOP_QUIT, INTINGAMEOP_1_X, INTINGAMEOP_2_Y, INTINGAMEOP_OP_W, "Host Quit", OPALIGN

    addHostQuitOptions = True
End Function

' ////////////////////////////////////////////////////////////////////////////

Private Function intAddInGameOptions_() As Boolean    ' static bool _intAddInGameOptions()
    Dim parent As WIDGET, ingameOp As WIDGET, row As Long, lines As Long, id As Long
    audio_StopAll

    'clear out any mission widgets - timers etc that may be on the screen
    clearMissionWidgets

    setWidgetsStatus True

    'if already open, then close!
    If Not widgGetFromID(psWScreen, INTINGAMEOP) Is Nothing Then
        intCloseInGameOptions False, True
        intAddInGameOptions_ = True
        Exit Function
    End If

    setReticulesEnabled False

    ' Pause the game.
    If Not gamePaused() Then
        kf_TogglePauseMode
    End If

    Set parent = psWScreen.psForm

    ' add form
    Set ingameOp = IntFormAnimated_new()
    parent.attach ingameOp
    ingameOp.id = INTINGAMEOP

    row = 1
    ' add 'resume'
    addIGTextButton INTINGAMEOP_RESUME, INTINGAMEOP_1_X, INTINGAMEOPAUTO_Y_LINE(row), INTINGAMEOP_OP_W, "Resume Game", OPALIGN
    row = row + 1

    If hasGameGuideTopics() Then
        ' add "View Guide"
        addIGTextButton INTINGAMEOP_OPENGAMEGUIDE, INTINGAMEOP_1_X, INTINGAMEOPAUTO_Y_LINE(row), INTINGAMEOP_OP_W, "View Guide", OPALIGN
        row = row + 1
    End If

    ' add 'options'
    addIGTextButton INTINGAMEOP_OPTIONS, INTINGAMEOP_1_X, INTINGAMEOPAUTO_Y_LINE(row), INTINGAMEOP_OP_W, "Options", OPALIGN
    row = row + 1

    If Not ((bMultiPlayer And NetPlay_bComms) Or bInTutorial Or NETisReplay()) Then
        If bMultiPlayer Then
            ' add 'load'
            addIGTextButton INTINGAMEOP_LOAD_SKIRMISH, INTINGAMEOP_1_X, INTINGAMEOPAUTO_Y_LINE(row), INTINGAMEOP_OP_W, "Load Game", OPALIGN
            row = row + 1
            ' add 'save'
            addIGTextButton INTINGAMEOP_SAVE_SKIRMISH, INTINGAMEOP_1_X, INTINGAMEOPAUTO_Y_LINE(row), INTINGAMEOP_OP_W, "Save Game", OPALIGN
            row = row + 1
        Else
            ' add 'load'
            addIGTextButton INTINGAMEOP_LOAD_MISSION, INTINGAMEOP_1_X, INTINGAMEOPAUTO_Y_LINE(row), INTINGAMEOP_OP_W, "Load Game", OPALIGN
            row = row + 1
            If Not getCamTweakOption_AutosavesOnly() Then
                ' add 'save'
                addIGTextButton INTINGAMEOP_SAVE_MISSION, INTINGAMEOP_1_X, INTINGAMEOPAUTO_Y_LINE(row), INTINGAMEOP_OP_W, "Save Game", OPALIGN
                row = row + 1
            End If
        End If
    End If

    ' add 'quit' text
    If NetPlay_isHost And bMultiPlayer And NetPlay_bComms Then
        If hostQuitConfirmation Then id = INTINGAMEOP_HOSTQUIT Else id = INTINGAMEOP_QUIT
        addIGTextButton id, INTINGAMEOP_1_X, INTINGAMEOPAUTO_Y_LINE(row), INTINGAMEOP_OP_W, "Host Quit", OPALIGN
    Else
        addIGTextButton INTINGAMEOP_CONFIRM_QUIT, INTINGAMEOP_1_X, INTINGAMEOPAUTO_Y_LINE(row), INTINGAMEOP_OP_W, "Quit", OPALIGN
    End If

    ' calculate layout
    lines = row
    ingameOp.setCalcLayout WzFunc_new(FN_intAddInGameOptions_layout, lines)

    intMode = INT_INGAMEOP              ' change interface mode.
    InGameOpUp = True                   ' inform interface.

    wzSetCursor CURSOR_DEFAULT

    intAddInGameOptions_ = True
End Function

Public Function intAddInGameOptions() As Boolean
    intAddInGameOptions = intAddInGameOptions_()
End Function

Private Function startInGameConfirmQuit() As Boolean
    Dim parent As WIDGET, ingameOp As WIDGET, row As Long, label As WIDGET
    widgDelete psWScreen, INTINGAMEOP   ' get rid of the old stuff.
    Set parent = psWScreen.psForm

    ' add form
    Set ingameOp = IntFormAnimated_new()
    parent.attach ingameOp
    ingameOp.id = INTINGAMEOP

    row = 1

    Set label = W_LABEL_new()
    ingameOp.attach label
    label.setGeometry INTINGAMEOP_2_X, INTINGAMEOPAUTO_Y_LINE(row), INTINGAMEOP4_OP_W, INTINGAMEOP_OP_H
    If bMultiPlayer And (NetPlay_bComms Or NETisReplay()) Then
        label.setString "Warning: Are you sure?" 'Do not mention saving in real multiplayer matches
    Else
        label.setString "Warning: Are you sure? Any unsaved progress will be lost."
    End If
    W_LABEL_setTextAlignment label, WLAB_ALIGNCENTRE
    W_LABEL_setFont label, font_medium, psPalette(WZCOL_YELLOW)

    row = row + 1

    ' add quit confirmation text
    addIGTextButton INTINGAMEOP_QUIT, INTINGAMEOP_2_X, INTINGAMEOPAUTO_Y_LINE(row), INTINGAMEOP4_OP_W, "Confirm", OPALIGN
    row = row + 1
    addIGTextButton INTINGAMEOP_GO_BACK, INTINGAMEOP_2_X, INTINGAMEOPAUTO_Y_LINE(row), INTINGAMEOP4_OP_W, "Back", OPALIGN

    ingameOp.setCalcLayout WzFunc_new(FN_startInGameConfirmQuit_layout, row)

    startInGameConfirmQuit = True
End Function

Private Function runInGameConfirmQuit(ByVal id As Long) As Boolean
    Select Case id
    Case INTINGAMEOP_QUIT
        runInGameConfirmQuit = True
        Exit Function

    Case INTINGAMEOP_GO_BACK
        intReopenMenuWithoutUnPausing

    Case Else
    End Select

    runInGameConfirmQuit = False
End Function

'
' Quick hack to throw up a ingame 'popup' for when the host drops connection.
'
Public Sub intAddInGamePopup()
    If vbDisplayOff Then Exit Sub  ' (VB) the vbtest runs: statsref's stand-in
    Dim parent As WIDGET, ingamePopup As WIDGET, sButInit As W_BUTINIT
    'clear out any mission widgets - timers etc that may be on the screen
    clearMissionWidgets
    setWidgetsStatus True
    intResetScreen False

    If isInGamePopupUp Then
        Exit Sub
    End If

    intCloseInGameOptions False, False  'clear out option-like menus

    audio_StopAll

    If Not gamePaused() Then
        kf_TogglePauseMode              ' Pause the game.
    End If

    Set parent = psWScreen.psForm

    Set ingamePopup = IntFormAnimated_new()
    parent.attach ingamePopup
    ingamePopup.id = INTINGAMEPOPUP
    ingamePopup.setCalcLayout WzFunc_new(FN_intAddInGamePopup_layout)

    ' add the text "buttons" now
    W_BUTINIT_construct sButInit

    sButInit.formID = INTINGAMEPOPUP
    sButInit.style = OPALIGN
    sButInit.width = 600
    sButInit.FontID = font_large
    sButInit.x = 0
    sButInit.height = 10
    Set sButInit.pDisplay = WzFunc_new(FN_displayTextOption)
    Set sButInit.initPUserDataFunc = WzFunc_new(FN_newDisplayTextOptionCache)
    Set sButInit.onDelete = WzFunc_new(FN_deleteDisplayTextOptionCache)

    sButInit.id = INTINGAMEOP_POPUP_MSG2
    sButInit.y = 20
    sButInit.pText = "Host has quit the game!"

    widgAddButton psWScreen, sButInit

    sButInit.id = INTINGAMEOP_POPUP_MSG1
    sButInit.y = 60
    sButInit.pText = "The game can't continue without the host."

    widgAddButton psWScreen, sButInit

    sButInit.id = INTINGAMEOP_POPUP_QUIT
    sButInit.y = 124
    sButInit.pText = "-->  QUIT  <--"

    widgAddButton psWScreen, sButInit

    intMode = INT_POPUPMSG              ' change interface mode.
    isInGamePopupUp = True

    shutdownGameStartScreen
End Sub

' ////////////////////////////////////////////////////////////////////////////

Private Sub ProcessOptionFinished()
    intMode = INT_NORMAL

    If gamePaused() Then
        kf_TogglePauseMode
    End If

    setReticulesEnabled True
End Sub

Public Sub intCloseInGameOptionsNoAnim()
    If NetPlay_isHost Then
        widgDelete psWScreen, INTINGAMEPOPUP
    End If
    widgDelete psWScreen, INTINGAMEOP
    InGameOpUp = False

    ProcessOptionFinished
    resetMissionWidgets
End Sub

' (VB) dynamic_cast<IntFormAnimated *>(widget)
Private Function asIntFormAnimated(ByVal widget As WIDGET) As WIDGET
    If Not widget Is Nothing Then
        If WIDGET_isKindOf(widget, K_IntFormAnimated) Then Set asIntFormAnimated = widget
    End If
End Function

' ////////////////////////////////////////////////////////////////////////////
Public Function intCloseInGameOptions(ByVal bPutUpLoadSave As Boolean, ByVal bResetMissionWidgets As Boolean) As Boolean
    Dim widg As WIDGET, form As WIDGET
    isInGameConfirmQuitUp = False

    If NetPlay_isHost Then
        widgDelete psWScreen, INTINGAMEPOPUP
    End If

    If bPutUpLoadSave Then
        Set widg = widgGetFromID(psWScreen, INTINGAMEOP)
        If Not widg Is Nothing Then
            widgDelete psWScreen, INTINGAMEOP
        End If

        InGameOpUp = False
    End If
    If Not bPutUpLoadSave Then
        ' close the form.
        ' Start the window close animation.
        Set form = asIntFormAnimated(widgGetFromID(psWScreen, INTINGAMEPOPUP))
        If Not form Is Nothing Then     ' FIXME: we hijack this routine for the popup close.
            IntFormAnimated_closeAnimateDelete form
            isInGamePopupUp = False
        End If
        Set form = asIntFormAnimated(widgGetFromID(psWScreen, INTINGAMEOP_POPUP_QUIT))
        If Not form Is Nothing Then
            IntFormAnimated_closeAnimateDelete form
            isInGamePopupUp = False
        End If

        Set form = asIntFormAnimated(widgGetFromID(psWScreen, INTINGAMEOP))
        If Not form Is Nothing Then
            IntFormAnimated_closeAnimateDelete form
            InGameOpUp = False
        End If
        Set form = asIntFormAnimated(widgGetFromID(psWScreen, INTINGAMEOP_QUIT))
        If Not form Is Nothing Then
            IntFormAnimated_closeAnimateDelete form
            InGameOpUp = False
        End If
    End If

    ProcessOptionFinished

    'don't add the widgets if the load/save screen is put up or exiting to front end
    If bResetMissionWidgets Then
        'put any widgets back on for the missions
        resetMissionWidgets
    End If

    ' the setting for group menu display may have been modified
    intShowGroupSelectionMenu
    intCloseInGameOptions = True
End Function

Public Sub intReopenMenuWithoutUnPausing()
    isInGameConfirmQuitUp = False

    If NetPlay_isHost Then
        widgDelete psWScreen, INTINGAMEPOPUP
    End If
    widgDelete psWScreen, INTINGAMEOP
    intAddInGameOptions
End Sub

Private Function startIGOptionsMenu() As Boolean
    showInGameOptionsScreen psWScreen, createOptionsBrowser(True), WzFunc_new(FN_startIGOptionsMenu_onClose)
    startIGOptionsMenu = True
End Function

' ////////////////////////////////////////////////////////////////////////////
' process clicks made by user.
Public Sub intProcessInGameOptions(ByVal id As Long)
    If vbDisplayOff Then Exit Sub  ' (VB) the vbtest runs: statsref's stand-in
    If isInGameConfirmQuitUp Then
        If runInGameConfirmQuit(id) Then
            intCloseInGameOptions True, True
        End If
        Exit Sub
    End If

    Select Case id
    ' NORMAL KEYS
    Case INTINGAMEOP_HOSTQUIT           'quit was pressed
        addHostQuitOptions

    Case INTINGAMEOP_CONFIRM_QUIT
        startInGameConfirmQuit
        isInGameConfirmQuitUp = True

    Case INTINGAMEOP_POPUP_QUIT, INTINGAMEOP_QUIT ' quit was confirmed.
        intCloseInGameOptions False, False

    Case INTINGAMEOP_OPENGAMEGUIDE
        intCloseInGameOptions False, False
        showGuideScreen WzFunc_new(FN_showGuideScreen_noopOnClose), True

    Case INTINGAMEOP_OPTIONS            'game options  was pressed
        intCloseInGameOptions False, False
        startIGOptionsMenu
    Case INTINGAMEOP_RESUME             'resume was pressed.
        intCloseInGameOptions False, True
    Case INTINGAMEOP_GO_BACK
        intReopenMenuWithoutUnPausing


    ' (the commented out INTINGAMEOP_REPLAY case)
    Case INTINGAMEOP_LOAD_MISSION
        intCloseInGameOptions True, False
        addLoadSave LOAD_INGAME_MISSION, "Load Campaign Saved Game" ' change mode when loadsave returns
    Case INTINGAMEOP_LOAD_SKIRMISH
        intCloseInGameOptions True, False
        addLoadSave LOAD_INGAME_SKIRMISH, "Load Skirmish Saved Game" ' change mode when loadsave returns
    Case INTINGAMEOP_SAVE_MISSION
        intCloseInGameOptions True, False
        addLoadSave SAVE_INGAME_MISSION, "Save Campaign Game"
    Case INTINGAMEOP_SAVE_SKIRMISH
        intCloseInGameOptions True, False
        addLoadSave SAVE_INGAME_SKIRMISH, "Save Skirmish Game"

    Case Else
    End Select
End Sub

'--------------------------------------------------------------------------
' (VB) the lambdas

Public Function ingameop_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim psWidget As WIDGET, inGameOp As WIDGET
    Select Case f.id
    Case FN_addHostQuitOptions_inGameOpLayout
        Set psWidget = args(0)
        psWidget.setGeometry INTINGAMEOP3_X, INTINGAMEOP3_Y, INTINGAMEOP3_W, INTINGAMEOP3_H

    Case FN_addHostQuitOptions_inGamePopupLayout
        Set psWidget = args(0)
        ' (assert(psWScreen != nullptr), assert(inGameOp != nullptr): the C asserts, off in a release build)
        Set inGameOp = widgGetFromID(psWScreen, INTINGAMEOP)
        psWidget.setGeometry (pie_GetVideoBufferWidth() - 600) \ 2, inGameOp.y() - 26 - 20, 600, 26

    Case FN_intAddInGameOptions_layout  ' [lines](WIDGET *psWidget)
        Set psWidget = args(0)
        psWidget.setGeometry INTINGAMEOP_X, INTINGAMEOPAUTO_Y(f.cap(0)), INTINGAMEOP_W, INTINGAMEOPAUTO_H(f.cap(0))

    Case FN_startInGameConfirmQuit_layout ' [row](WIDGET* psWidget)
        Set psWidget = args(0)
        psWidget.setGeometry INTINGAMEOP4_X, INTINGAMEOPAUTO_Y(f.cap(0)), INTINGAMEOP4_W, INTINGAMEOPAUTO_H(f.cap(0))

    Case FN_intAddInGamePopup_layout
        Set psWidget = args(0)
        psWidget.setGeometry 20 + D_W, (240 - 160 \ 2) + D_H, 600, 160

    Case FN_newDisplayTextOptionCache
        Set ingameop_c_fn = New DisplayTextOptionCache

    Case FN_startIGOptionsMenu_onClose
        ' the setting for group menu display may have been modified
        intShowGroupSelectionMenu

    Case FN_showGuideScreen_noopOnClose
        ' no=op on close func
    End Select
End Function
