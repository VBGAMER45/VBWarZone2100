Attribute VB_Name = "frontend_c"
' src/frontend.cpp, frontend.h: the title menus (title, tutorial, single
' player, multiplayer), the front end's widget helpers (addBackdrop,
' addTopForm, addBottomForm, addText, addSideText, addTextButton,
' addSmallTextButton, addFESlider) and their display functions, SPinit,
' changeTitleMode, the flags image file, the display scale prompt. The
' widgets are the object model's (CONVENTIONS, "Widgets"); the lambdas and
' display functions are this module's WzFuncs (FN_ 3000-3099). aLevelName,
' saveGameName, sRequestResult and MAX_LEVEL_NAME_SIZE are still in
' stubs_game. tMode's SINGLE is SINGLE_ (VB's type keyword). Translations
' are not ported (_() and P_() give the English text).
Option Explicit

' frontend.h: enum tMode
Public Enum tMode
    TITLE                               ' 0 intro mode
    SINGLE_                             ' 1 single player menu
    MULTI                               ' 2 multiplayer menu
    TUTORIAL                            ' 3  tutorial/fastplay
    TITLE_UNUSED                        ' 4
    FORCESELECT                         ' 5 MULTIPLAYER, Force design screen
    STARTGAME                           ' 6 Fire up the game
    SHOWINTRO                           ' 7 reshow the intro
    QUIT                                ' 8 leaving game
    LOADSAVEGAME                        ' 9 loading a save game
    CAMPAIGNS                           ' 10 campaign selector
End Enum

' frontend.h: struct DisplayTextOptionCache's enum class OverflowBehavior (the cache is DisplayTextOptionCache.cls)
Public Enum OverflowBehavior
    OverflowBehavior_None
    OverflowBehavior_ShrinkFont
End Enum

' frontend.h: the front end's layout
Public Const FRONTEND_TOPFORMX As Long = 50
Public Const FRONTEND_TOPFORMY As Long = 10
Public Const FRONTEND_TOPFORMW As Long = 540
Public Const FRONTEND_TOPFORMH As Long = 150
Public Const FRONTEND_TOPFORM_WIDEX As Long = 28
Public Const FRONTEND_TOPFORM_WIDEY As Long = 10
Public Const FRONTEND_TOPFORM_WIDEW As Long = 588
Public Const FRONTEND_TOPFORM_WIDEH As Long = 150
Public Const FRONTEND_BOTFORMX As Long = FRONTEND_TOPFORMX
Public Const FRONTEND_BOTFORMY As Long = 170
Public Const FRONTEND_BOTFORMW As Long = FRONTEND_TOPFORMW
Public Const FRONTEND_BOTFORMH As Long = 305             ' keep Y+H < 480 (minimum display height)
Public Const FRONTEND_BOTFORM_WIDEX As Long = 30
Public Const FRONTEND_BOTFORM_WIDEY As Long = FRONTEND_BOTFORMY
Public Const FRONTEND_BOTFORM_WIDEW As Long = 580
Public Const FRONTEND_BOTFORM_WIDEH As Long = FRONTEND_BOTFORMH ' keep Y+H < 480 (minimum display height)
Public Const FRONTEND_BUTWIDTH As Long = FRONTEND_BOTFORMW - 40 ' text button sizes.
Public Const FRONTEND_BUTHEIGHT As Long = 35
Public Const FRONTEND_BUTHEIGHT_LIST_SPACER As Long = (FRONTEND_BUTHEIGHT + 15)
Public Const FRONTEND_BUTWIDTH_WIDE As Long = FRONTEND_BOTFORM_WIDEW - 40 ' text button sizes.
Public Const FRONTEND_BUTHEIGHT_WIDE As Long = FRONTEND_BUTHEIGHT
Public Const FRONTEND_POS1X As Long = 20                 ' button positions
Public Const FRONTEND_POS1Y As Long = (0 * FRONTEND_BUTHEIGHT)
Public Const FRONTEND_POS1M As Long = 340
Public Const FRONTEND_POS2X As Long = 20
Public Const FRONTEND_POS2Y As Long = (1 * FRONTEND_BUTHEIGHT)
Public Const FRONTEND_POS2M As Long = 340
Public Const FRONTEND_POS3X As Long = 20
Public Const FRONTEND_POS3Y As Long = (2 * FRONTEND_BUTHEIGHT)
Public Const FRONTEND_POS3M As Long = 340
Public Const FRONTEND_POS4X As Long = 20
Public Const FRONTEND_POS4Y As Long = (3 * FRONTEND_BUTHEIGHT)
Public Const FRONTEND_POS4M As Long = 340
Public Const FRONTEND_POS5X As Long = 20
Public Const FRONTEND_POS5Y As Long = (4 * FRONTEND_BUTHEIGHT)
Public Const FRONTEND_POS5M As Long = 340
Public Const FRONTEND_POS6X As Long = 20
Public Const FRONTEND_POS6Y As Long = (5 * FRONTEND_BUTHEIGHT)
Public Const FRONTEND_POS6M As Long = 340
Public Const FRONTEND_POS7X As Long = 20
Public Const FRONTEND_POS7Y As Long = (6 * FRONTEND_BUTHEIGHT)
Public Const FRONTEND_POS7M As Long = 340
Public Const FRONTEND_POS8X As Long = 20
Public Const FRONTEND_POS8Y As Long = (7 * FRONTEND_BUTHEIGHT)
Public Const FRONTEND_POS8M As Long = 340
Public Const FRONTEND_POS9X As Long = -30                ' special case for our hyperlink
Public Const FRONTEND_POS9Y As Long = (8 * FRONTEND_BUTHEIGHT)
Public Const FRONTEND_SIDEX As Long = 24
Public Const FRONTEND_SIDEY As Long = FRONTEND_BOTFORMY

' frontend.h: the front end's widget ids
Public Enum FRONTEND_ID
    FRONTEND_BACKDROP = 20000
    FRONTEND_TOPFORM
    FRONTEND_BOTFORM
    FRONTEND_LOGO
    FRONTEND_SIDETEXT                   ' sideways text
    FRONTEND_MULTILINE_SIDETEXT         ' sideways text
    FRONTEND_SIDETEXT1                  ' sideways text
    FRONTEND_SIDETEXT2                  ' sideways text
    FRONTEND_SIDETEXT3                  ' sideways text
    FRONTEND_SIDETEXT4                  ' sideways text
    FRONTEND_PASSWORDFORM
    FRONTEND_HYPERLINK
    FRONTEND_UPGRDLINK
    FRONTEND_DONATELINK
    FRONTEND_CHATLINK
    ' begin menu
    FRONTEND_SINGLEPLAYER = 20100       ' title screen
    FRONTEND_MULTIPLAYER
    FRONTEND_TUTORIAL
    FRONTEND_PLAYINTRO
    FRONTEND_OPTIONS
    FRONTEND_QUIT
    FRONTEND_FASTPLAY                   'tutorial menu option
    FRONTEND_CONTINUE
    FRONTEND_NEWGAME = 20200            ' single player (menu)
    FRONTEND_LOADGAME_MISSION
    FRONTEND_LOADGAME_SKIRMISH
    FRONTEND_REPLAY
    FRONTEND_SKIRMISH
    FRONTEND_CHALLENGES
    FRONTEND_HOST = 20300               'multiplayer menu options
    FRONTEND_JOIN
    FE_P0                               ' player 0 buton
    FE_P1                               ' player 1 buton
    FE_P2                               ' player 2 buton
    FE_P3                               ' player 3 buton
    FE_P4                               ' player 4 buton
    FE_P5                               ' player 5 buton
    FE_P6                               ' player 6 buton
    FE_P7                               ' player 7 buton
    FE_MP_PR                            ' Multiplayer player random button
    FE_MP_PMAX = FE_MP_PR + MAX_PLAYERS_IN_GUI  ' Multiplayer player blah button
    FRONTEND_CAMPAIGN_1
    FRONTEND_CAMPAIGN_2
    FRONTEND_CAMPAIGN_3
    FRONTEND_CAMPAIGN_4
    FRONTEND_CAMPAIGN_5
    FRONTEND_CAMPAIGN_6
    FRONTEND_KEYMAP = 26000             ' Keymap menu
    FRONTEND_MUSICMANAGER = 27000       ' Music manager menu
    FRONTEND_NOGAMESAVAILABLE = 31666   ' Used when no games are available in lobby
End Enum

' (VB) this module's lambdas and function pointers (WzFunc ids, frontend_c_fn)
Public Const FN_backdropCalcLayout As Long = 3000
Public Const FN_backdropOnDelete As Long = 3001
Public Const FN_displayTitleBitmap As Long = 3002
Public Const FN_topFormCalcLayoutWide As Long = 3003
Public Const FN_topFormCalcLayout As Long = 3004
Public Const FN_botFormCalcLayoutWide As Long = 3005
Public Const FN_botFormCalcLayout As Long = 3006
Public Const FN_displayLogo As Long = 3007
Public Const FN_displayBigSlider As Long = 3008
Public Const FN_displayTextAt270 As Long = 3009
Public Const FN_displayTextOption As Long = 3010
Public Const FN_deleteDisplayTextOptionCache As Long = 3011  ' the onDelete lambdas of addSideText, makeTextButton, addSmallTextButton
Public Const FN_intUpdateQuantitySlider As Long = 3012       ' intdisplay.cpp's (intdisplay_c): makeFESlider's pCallback
Public Const FN_missingVideosAction As Long = 3013
Public Const FN_displayScalePromptTagIs As Long = 3014
Public Const FN_displayScaleIncreaseAction As Long = 3015
Public Const FN_displayScalePromptDismissed As Long = 3016

Private Const TUTORIAL_LEVEL As String = "TUTORIAL3"
Private Const TRANSLATION_URL As String = "https://translate.wz2100.net"

' hci.h (cluster G)

Private Const VIDEO_TAG As String = "videoMissing"

Private Const DISPLAY_SCALE_PROMPT_TAG_PREFIX As String = "displayscale::prompt::"
Private Const DISPLAY_SCALE_PROMPT_INCREASE_TAG As String = DISPLAY_SCALE_PROMPT_TAG_PREFIX & "increase"
Private Const WZ_SUGGESTED_MAX_LOGICAL_SCREEN_SIZE_DIMENSION As Long = 1440

' ////////////////////////////////////////////////////////////////////////////
' Global variables

Public bLimiterLoaded As Boolean        ' = false

Private pFlagsImages As Long            ' static std::shared_ptr<IMAGEFILE> pFlagsImages (a WzGfx image file handle, 0 = nullptr)

Private lastProcessedDisplayScale As Long ' unsigned = 0
Private lastProcessedScreenWidth As Long
Private lastProcessedScreenHeight As Long

' ////////////////////////////////////////////////////////////////////////////
' Helper functions

' Returns true if escape key pressed.
'
Public Function CancelPressed() As Boolean
    Dim cancel As Boolean
    cancel = keyPressed(KEY_ESC)
    If cancel Then
        inputLoseFocus   ' clear the input buffer.
    End If
    CancelPressed = cancel
End Function

' Cycle through options as in program seq(1) from coreutils
' The T cast is to cycle through enums. (VB: T is Long)
Private Function seqCycle(ByVal value As Long, ByVal min As Long, ByVal inc As Long, ByVal max As Long) As Long
    If widgGetButtonKey_DEPRECATED(psWScreen) = WKEY_PRIMARY Then
        If value < max Then seqCycle = value + inc Else seqCycle = min   ' Cycle forwards.
    Else
        If min < value Then seqCycle = value - inc Else seqCycle = max   ' Cycle backwards.
    End If
End Function

' Cycle through options, which are powers of two, such as [128, 256, 512, 1024, 2048].
Private Function pow2Cycle(ByVal value As Long, ByVal min As Long, ByVal max As Long) As Long
    If widgGetButtonKey_DEPRECATED(psWScreen) = WKEY_PRIMARY Then
        If value < max Then
            If value > 1 Then pow2Cycle = value * 2 Else pow2Cycle = 2   ' Cycle forwards.
        Else
            pow2Cycle = min
        End If
    Else
        If min < value Then
            If value \ 2 > 1 Then pow2Cycle = value \ 2 Else pow2Cycle = 0   ' Cycle backwards.
        Else
            pow2Cycle = max
        End If
    End If
End Function

Private Sub moveToParentRightEdge(ByVal widget As WIDGET, ByVal right_ As Long)
    Dim parent As WIDGET
    Set parent = widget.parent()
    If Not parent Is Nothing Then
        widget.move parent.width() - (widget.width() + right_), widget.y()
    End If
End Sub

' ////////////////////////////////////////////////////////////////////////////
' Title Screen
Private Sub runUpgrdHyperlink()
    Dim link As String, version As String, versionStr As String, i As Long, ch As String
    link = "https://wz2100.net/versioncheck/?ver="
    version = version_getVersionString()
    versionStr = ""
    For i = 1 To Len(version)
        ch = Mid$(version, i, 1)
        If ch = " " Then versionStr = versionStr & "_" Else versionStr = versionStr & ch
    Next
    link = link & urlEncode(versionStr)
    openURLInBrowser link
End Sub

Private Sub runHyperlink()
    openURLInBrowser "https://wz2100.net/"
End Sub

Private Sub rundonatelink()
    openURLInBrowser "http://donations.wz2100.net/"
End Sub

Private Sub runchatlink()
    openURLInBrowser "https://wz2100.net/webchat/"
End Sub

Public Sub notifyAboutMissingVideos()
    Dim notification As WZ_Notification
    If Not hasNotificationsWithTag(VIDEO_TAG) Then
        notification.duration = 0
        notification.contentTitle = "Campaign videos are missing"
        notification.contentText = "See our FAQ on how to install videos"
        notification.tag = VIDEO_TAG
        notification.largeIcon = "images/notifications/exclamation_triangle.png"    ' WZ_Notification_Image(...)
        notification.action_title = "Open wz2100.net"                               ' WZ_Notification_Action(title, [](const WZ_Notification&) {...})
        Set notification.action_onAction = WzFunc_new(FN_missingVideosAction)
        notification.displayOptions_identifier = "campaignVideoNotification"        ' WZ_Notification_Display_Options::makeIgnorable(..., 2)
        notification.displayOptions_numTimesSeenBeforeDoNotShowAgainOption = 2

        addNotification notification, GAME_TICKS_PER_SEC                          ' WZ_Notification_Trigger(GAME_TICKS_PER_SEC)
    End If
End Sub

Private Sub closeMissingVideosNotification()
    cancelOrDismissNotificationsWithTag VIDEO_TAG
End Sub

Public Sub startTitleMenu()
    Dim pRightAlignedButton As WIDGET
    intRemoveReticule

    addBackdrop
    addTopForm False
    addBottomForm

    addTextButton FRONTEND_SINGLEPLAYER, FRONTEND_POS2X, FRONTEND_POS2Y, "Single Player", WBUT_TXTCENTRE
    addTextButton FRONTEND_MULTIPLAYER, FRONTEND_POS3X, FRONTEND_POS3Y, "Multi Player", WBUT_TXTCENTRE
    addTextButton FRONTEND_TUTORIAL, FRONTEND_POS4X, FRONTEND_POS4Y, "Tutorial", WBUT_TXTCENTRE
    addTextButton FRONTEND_OPTIONS, FRONTEND_POS5X, FRONTEND_POS5Y, "Options", WBUT_TXTCENTRE

    ' check whether video sequences are installed
    If seq_hasVideos() Then
        addTextButton FRONTEND_PLAYINTRO, FRONTEND_POS6X, FRONTEND_POS6Y, "View Intro", WBUT_TXTCENTRE
    Else
        addTextButton FRONTEND_PLAYINTRO, FRONTEND_POS6X, FRONTEND_POS6Y, "View Intro", WBUT_TXTCENTRE Or WBUT_DISABLE
        widgSetTip psWScreen, FRONTEND_PLAYINTRO, "Videos are missing, download them from http://wz2100.net"

        notifyAboutMissingVideos
    End If

    If findLastSave() Then
        addTextButton FRONTEND_CONTINUE, FRONTEND_POS7X, FRONTEND_POS7Y, "Continue Last Save", WBUT_TXTCENTRE
    Else
        addTextButton FRONTEND_CONTINUE, FRONTEND_POS7X, FRONTEND_POS7Y, "Continue Last Save", WBUT_TXTCENTRE Or WBUT_DISABLE
        widgSetTip psWScreen, FRONTEND_CONTINUE, "No last save available"
    End If
    addTextButton FRONTEND_QUIT, FRONTEND_POS8X, FRONTEND_POS8Y, "Quit Game", WBUT_TXTCENTRE
    addSideText FRONTEND_SIDETEXT, FRONTEND_SIDEX, FRONTEND_SIDEY, "MAIN MENU"

    addSmallTextButton FRONTEND_HYPERLINK, FRONTEND_POS9X, FRONTEND_POS9Y, "Official site: http://wz2100.net/", 0
    widgSetTip psWScreen, FRONTEND_HYPERLINK, "Come visit the forums and all Warzone 2100 news! Click this link."
    Set pRightAlignedButton = addSmallTextButton(FRONTEND_DONATELINK, FRONTEND_POS9X + 360, FRONTEND_POS9Y, "Donate: http://donations.wz2100.net/", 0)
    moveToParentRightEdge pRightAlignedButton, 1
    widgSetTip psWScreen, FRONTEND_DONATELINK, "Help support the project with our server costs, Click this link."
    Set pRightAlignedButton = addSmallTextButton(FRONTEND_CHATLINK, FRONTEND_POS9X + 360, 0, "Chat with players on Discord or IRC", 0)
    moveToParentRightEdge pRightAlignedButton, 6
    widgSetTip psWScreen, FRONTEND_CHATLINK, "Connect to Discord or Freenode through webchat by clicking this link."
    addMultiBut_screen psWScreen, FRONTEND_BOTFORM, FRONTEND_UPGRDLINK, 7, 7, MULTIOP_BUTW, MULTIOP_BUTH, "Check for a newer version", IMAGE_GAMEVERSION, IMAGE_GAMEVERSION_HI, 1 ' (VB) hi: true (1)
End Sub

Public Sub runContinue()
    If lastSaveMP Then SPinit LEVEL_TYPE_SKIRMISH Else SPinit LEVEL_TYPE_CAMPAIGN
    saveGameName = Left$(lastSavePath_toPath_GAM(), 255)  ' sstrcpy(saveGameName, lastSavePath.toPath(SaveGamePath_t::Extension::GAM).c_str())
    bMultiPlayer = lastSaveMP
End Sub

' (VB) WidgetTriggers const &triggers = widgRunScreen(psScreen); triggers.empty() ? 0 : triggers.front().widget->id
' Just use first click here, since the next click could be on another menu.
Private Function firstTriggerId(ByVal psScreen As W_SCREEN) As Long
    Dim triggers() As WidgetTrigger, n As Long
    n = widgRunScreen(psScreen, triggers)
    If n = 0 Then firstTriggerId = 0 Else firstTriggerId = triggers(0).widget.id
End Function

Public Function runTitleMenu() As Boolean
    Dim id As Long
    id = firstTriggerId(psWScreen) ' Just use first click here, since the next click could be on another menu.

    Select Case id
    Case FRONTEND_QUIT
        changeTitleMode QUIT
    Case FRONTEND_MULTIPLAYER
        changeTitleMode MULTI
    Case FRONTEND_SINGLEPLAYER
        changeTitleMode SINGLE_
    Case FRONTEND_OPTIONS
        changeTitleUI WzOptionsTitleUI_new(wzTitleUICurrent)
    Case FRONTEND_TUTORIAL
        changeTitleMode TUTORIAL
    Case FRONTEND_PLAYINTRO
        changeTitleMode SHOWINTRO
    Case FRONTEND_HYPERLINK
        runHyperlink
    Case FRONTEND_UPGRDLINK
        runUpgrdHyperlink
    Case FRONTEND_DONATELINK
        rundonatelink
    Case FRONTEND_CHATLINK
        runchatlink
    Case FRONTEND_CONTINUE
        runContinue
        changeTitleMode LOADSAVEGAME

    Case Else
    End Select

    widgDisplayScreen psWScreen ' show the widgets currently running

    runTitleMenu = True
End Function


' ////////////////////////////////////////////////////////////////////////////
' Tutorial Menu
Public Sub startTutorialMenu()
    addBackdrop
    addTopForm False
    addBottomForm


    addTextButton FRONTEND_TUTORIAL, FRONTEND_POS3X, FRONTEND_POS3Y, "Tutorial", WBUT_TXTCENTRE
    addTextButton FRONTEND_FASTPLAY, FRONTEND_POS4X, FRONTEND_POS4Y, "Fast Play", WBUT_TXTCENTRE
    addSideText FRONTEND_SIDETEXT, FRONTEND_SIDEX, FRONTEND_SIDEY, "TUTORIALS"
    ' TRANSLATORS: "Return", in this context, means "return to previous screen/menu"
    addMultiBut_screen psWScreen, FRONTEND_BOTFORM, FRONTEND_QUIT, 10, 10, 30, 29, "Return", IMAGE_RETURN, IMAGE_RETURN_HI, IMAGE_RETURN_HI
End Sub

Public Function runTutorialMenu() As Boolean
    Dim id As Long
    id = firstTriggerId(psWScreen) ' Just use first click here, since the next click could be on another menu.

    Select Case id
    Case FRONTEND_TUTORIAL
        SPinit LEVEL_TYPE_CAMPAIGN
        aLevelName = Left$(TUTORIAL_LEVEL, MAX_LEVEL_NAME_SIZE)
        setGroupButtonEnabled False ' hack to disable the groups UI for the tutorial
        changeTitleMode STARTGAME

    Case FRONTEND_FASTPLAY
        SPinit LEVEL_TYPE_CAMPAIGN
        aLevelName = "FASTPLAY"
        changeTitleMode STARTGAME

    Case FRONTEND_QUIT
        changeTitleMode TITLE

    Case Else
    End Select

    ' If close button pressed then return from this menu.
    If CancelPressed() Then
        changeTitleMode TITLE
    End If

    widgDisplayScreen psWScreen                      ' show the widgets currently running

    runTutorialMenu = True
End Function


' ////////////////////////////////////////////////////////////////////////////
' Single Player Menu
Public Sub startSinglePlayerMenu()
    challengeActive = False
    addBackdrop
    addTopForm False
    addBottomForm

    addTextButton FRONTEND_NEWGAME, FRONTEND_POS2X, FRONTEND_POS2Y, "New Campaign", WBUT_TXTCENTRE
    addTextButton FRONTEND_SKIRMISH, FRONTEND_POS3X, FRONTEND_POS3Y, "Start Skirmish Game", WBUT_TXTCENTRE
    addTextButton FRONTEND_CHALLENGES, FRONTEND_POS4X, FRONTEND_POS4Y, "Challenges", WBUT_TXTCENTRE
    addTextButton FRONTEND_LOADGAME_MISSION, FRONTEND_POS5X, FRONTEND_POS5Y, "Load Campaign Game", WBUT_TXTCENTRE
    addTextButton FRONTEND_LOADGAME_SKIRMISH, FRONTEND_POS6X, FRONTEND_POS6Y, "Load Skirmish Game", WBUT_TXTCENTRE
    addTextButton FRONTEND_REPLAY, FRONTEND_POS7X, FRONTEND_POS7Y, "View Skirmish Replay", WBUT_TXTCENTRE

    addSideText FRONTEND_SIDETEXT, FRONTEND_SIDEX, FRONTEND_SIDEY, "SINGLE PLAYER"
    addMultiBut_screen psWScreen, FRONTEND_BOTFORM, FRONTEND_QUIT, 10, 10, 30, 29, "Return", IMAGE_RETURN, IMAGE_RETURN_HI, IMAGE_RETURN_HI
    ' show this only when the video sequences are not installed
    If Not seq_hasVideos() Then
        addSmallTextButton FRONTEND_HYPERLINK, FRONTEND_POS9X, FRONTEND_POS9Y, "Campaign videos are missing! Get them from http://wz2100.net", 0
    End If
End Sub

Private Sub loadOK()
    If Len(sRequestResult) > 0 Then
        saveGameName = Left$(sRequestResult, 255)
        changeTitleMode LOADSAVEGAME
    End If
End Sub

Public Sub SPinit(ByVal gameType As Long)
    Dim playercolor As Long             ' uint8_t

    NetPlay_bComms = False
    bMultiPlayer = False
    bMultiMessages = False
    game.typ = gameType
    NET_InitPlayers
    NetPlay_players_allocated(0) = True
    NetPlay_players_autoGame(0) = False
    NetPlay_players_difficulty(0) = AIDifficulty_HUMAN
    game.maxPlayers = MAX_PLAYERS - 1  ' Currently, 0 - 10 for a total of MAX_PLAYERS
    ' make sure we have a valid color choice for our SP game. Valid values are 0, 4-7
    playercolor = war_GetSPcolor() And &HFF&

    If playercolor >= 1 And playercolor <= 3 Then
        playercolor = 0    ' default is green
    End If
    setPlayerColour 0, playercolor
    Sha256_setZero game.hash   ' must reset this to zero
    builtInMap = True
    useTerrainOverrides = True
End Sub

Public Function runSinglePlayerMenu() As Boolean
    Dim id As Long
    If bLoadSaveUp Then
        If runLoadSave(False) Then ' check for file name.
            loadOK
        End If
    ElseIf challengesUp Then
        runChallenges
    Else
        id = firstTriggerId(psWScreen) ' Just use first click here, since the next click could be on another menu.

        Select Case id
        Case FRONTEND_NEWGAME
            ActivityManager_navigateToMenu "Campaign"
            changeTitleUI WzCampaignSelectorTitleUI_new(wzTitleUICurrent)

        Case FRONTEND_LOADGAME_MISSION
            SPinit LEVEL_TYPE_CAMPAIGN
            addLoadSave LOAD_FRONTEND_MISSION, "Load Campaign Saved Game"   ' change mode when loadsave returns

        Case FRONTEND_LOADGAME_SKIRMISH
            SPinit LEVEL_TYPE_SKIRMISH
            bMultiPlayer = True
            addLoadSave LOAD_FRONTEND_SKIRMISH, "Load Skirmish Saved Game"  ' change mode when loadsave returns

        Case FRONTEND_REPLAY
            SPinit LEVEL_TYPE_SKIRMISH
            bMultiPlayer = True
            addLoadSave LOADREPLAY_FRONTEND_SKIRMISH, "Load Skirmish Replay"    ' change mode when loadsave returns

        Case FRONTEND_SKIRMISH
            SPinit LEVEL_TYPE_SKIRMISH
            game.map = Left$(DEFAULTSKIRMISHMAP, 127)
            game.hash = levGetMapNameHash(game.map)
            game.maxPlayers = DEFAULTSKIRMISHMAPMAXPLAYERS

            ingame.side = InGameSide_HOST_OR_SINGLEPLAYER
            changeTitleUI WzMultiplayerOptionsTitleUI_new(wzTitleUICurrent)

        Case FRONTEND_QUIT
            changeTitleMode TITLE

        Case FRONTEND_CHALLENGES
            SPinit LEVEL_TYPE_SKIRMISH
            addChallenges

        Case FRONTEND_HYPERLINK
            runHyperlink

        Case Else
        End Select

        If CancelPressed() Then
            changeTitleMode TITLE
        End If
    End If

    If Not bLoadSaveUp And Not challengesUp Then                     ' if save/load screen is up
        widgDisplayScreen psWScreen                      ' show the widgets currently running
    End If
    If bLoadSaveUp Then                             ' if save/load screen is up
        displayLoadSave
    ElseIf challengesUp Then
        displayChallenges
    End If

    runSinglePlayerMenu = True
End Function


' ////////////////////////////////////////////////////////////////////////////
' Multi Player Menu
Public Sub startMultiPlayerMenu()
    Dim buf As String
    closeMissingVideosNotification

    addBackdrop
    addTopForm False
    addBottomForm

    addSideText FRONTEND_SIDETEXT, FRONTEND_SIDEX, FRONTEND_SIDEY, "MULTI PLAYER"

    addTextButton FRONTEND_HOST, FRONTEND_POS2X, FRONTEND_POS2Y, "Host Game", WBUT_TXTCENTRE
    addTextButton FRONTEND_JOIN, FRONTEND_POS3X, FRONTEND_POS3Y, "Join Game", WBUT_TXTCENTRE
    addTextButton FRONTEND_REPLAY, FRONTEND_POS7X, FRONTEND_POS7Y, "View Replay", WBUT_TXTCENTRE

    addMultiBut_screen psWScreen, FRONTEND_BOTFORM, FRONTEND_QUIT, 10, 10, 30, 29, "Return", IMAGE_RETURN, IMAGE_RETURN_HI, IMAGE_RETURN_HI

    ' This isn't really a hyperlink for now... perhaps link to the wiki ?
    buf = Left$(va("TCP port %d must be opened in your firewall or router to host games!", NETgetGameserverPort()), 511)
    addSmallTextButton FRONTEND_HYPERLINK, FRONTEND_POS9X, FRONTEND_POS9Y, buf, 0
End Sub

Public Function runMultiPlayerMenu() As Boolean
    Dim id As Long
    If bLoadSaveUp Then
        If runLoadSave(False) Then ' check for file name.
            loadOK
        End If
    Else
        id = firstTriggerId(psWScreen) ' Just use first click here, since the next click could be on another menu.

        Select Case id
        Case FRONTEND_HOST
            ' First of all, make sure we've reset any prior networking state
            NETshutdown
            ' don't pretend we are running a network game. Really do it!
            NetPlay_bComms = True ' use network = true
            ingame.side = InGameSide_HOST_OR_SINGLEPLAYER
            bMultiPlayer = True
            bMultiMessages = True
            NETinit war_getHostConnectionProvider()
            NETinitPortMapping
            game.typ = LEVEL_TYPE_SKIRMISH      ' needed?
            changeTitleUI WzMultiplayerOptionsTitleUI_new(wzTitleUICurrent)
        Case FRONTEND_JOIN
            ' Don't call `NETinit()` just yet.
            ' It will be called automatically during join attempts.
            changeTitleUI WzGameBrowserTitleUI_new(wzTitleUICurrent)
        Case FRONTEND_REPLAY
            SPinit LEVEL_TYPE_SKIRMISH
            bMultiPlayer = True
            game.maxPlayers = DEFAULTSKIRMISHMAPMAXPLAYERS
            addLoadSave LOADREPLAY_FRONTEND_MULTI, "Load Multiplayer Replay"   ' change mode when loadsave returns

        Case FRONTEND_QUIT
            changeTitleMode TITLE
        Case Else
        End Select

        If CancelPressed() Then
            changeTitleMode TITLE
        End If
    End If

    If Not bLoadSaveUp Then
        widgDisplayScreen psWScreen      ' show the widgets currently running
    ElseIf bLoadSaveUp Then                  ' if save/load screen is up
        displayLoadSave
    End If

    runMultiPlayerMenu = True
End Function

' std::vector<unsigned int>: the scales in displayScales(0 ..), their number returned
Public Function availableDisplayScalesSorted(displayScales() As Long) As Long
    Dim n As Long, i As Long, j As Long, t As Long
    n = wzAvailableDisplayScales(displayScales)
    ' std::sort
    For i = 1 To n - 1
        t = displayScales(i)
        j = i - 1
        Do While j >= 0
            If displayScales(j) <= t Then Exit Do
            displayScales(j + 1) = displayScales(j)
            j = j - 1
        Loop
        displayScales(j + 1) = t
    Next
    availableDisplayScalesSorted = n
End Function

' std::shared_ptr<IMAGEFILE>: the image file's handle (0 = nullptr)
Public Function getFlagsImages() As Long
    Dim errorMessage As String, mods() As WzMods_LoadedMod
    If pFlagsImages = 0 Then
        pFlagsImages = iV_LoadImageFile("images/flags.img")
        If pFlagsImages = 0 Then
            errorMessage = va("Unable to load: %s.", "flags.img")
            If getLoadedMods(mods) > 0 Then
                errorMessage = errorMessage & " "
                errorMessage = errorMessage & "Please remove all incompatible mods."
            End If
            debug_ 639, LOG_FATAL, "getFlagsImages", "%s", errorMessage
        End If
    End If
    getFlagsImages = pFlagsImages
End Function

' ////////////////////////////////////////////////////////////////////////////
' drawing functions

' show a background picture (currently used for version and mods labels)
Private Sub displayTitleBitmap(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim modListText As String, cache As TitleBitmapCache
    modListText = ""

    If Len(getModList()) > 0 Then
        modListText = Left$(modListText & "Mod: ", MAX_STR_LENGTH - 1)
        modListText = Left$(modListText & getModList(), MAX_STR_LENGTH - 1)
    End If

    If pie_GetVideoBufferWidth() <= 0 Or pie_GetVideoBufferHeight() <= 0 Then
        ' don't bother drawing when the dimensions are <= 0,0
        Exit Sub
    End If

    ' (assert(psWidget->pUserData != nullptr): the C assert, off in a release build)
    Set cache = psWidget.pUserData

    cache.formattedVersionString.setText version_getFormattedVersionString(), font_regular
    cache.modListText.setText modListText, font_regular
    cache.gfxBackend.setText gfx_api_context_getFormattedRendererInfoString(), font_small

    cache.formattedVersionString.render pie_GetVideoBufferWidth() - 9, pie_GetVideoBufferHeight() - 14, psPalette(WZCOL_GREY), 270!

    If Len(getModList()) > 0 Then
        cache.modListText.render 9, 14, psPalette(WZCOL_GREY)
    End If

    cache.gfxBackend.render 9, pie_GetVideoBufferHeight() - 10, psPalette(WZCOL_GREY)

    cache.formattedVersionString.render pie_GetVideoBufferWidth() - 10, pie_GetVideoBufferHeight() - 15, psPalette(WZCOL_TEXT_BRIGHT), 270!

    If Len(getModList()) > 0 Then
        cache.modListText.render 10, 15, psPalette(WZCOL_TEXT_BRIGHT)
    End If

    cache.gfxBackend.render 10, pie_GetVideoBufferHeight() - 11, psPalette(WZCOL_TEXT_BRIGHT)
End Sub

' ////////////////////////////////////////////////////////////////////////////
' show warzone logo
Private Sub displayLogo(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    iV_DrawImage2 "image_fe_logo.png", xOffset + psWidget.x(), yOffset + psWidget.y(), psWidget.width(), psWidget.height()
End Sub


' ////////////////////////////////////////////////////////////////////////////
' show, well have a guess..
Private Sub displayBigSlider(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim Slider As WIDGET, x As Long, y As Long, sx As Long, stops As Long
    Set Slider = psWidget               ' W_SLIDER *Slider = (W_SLIDER *)psWidget;
    x = xOffset + psWidget.x()
    y = yOffset + psWidget.y()

    iV_DrawImage IntImages, IMAGE_SLIDER_BIG, x + STAT_SLD_OX, y + STAT_SLD_OY           ' draw bdrop

    stops = W_SLIDER_(Slider).numStops
    If stops < 1 Then stops = 1
    sx = (Slider.width() - 3 - W_SLIDER_(Slider).barSize) * W_SLIDER_(Slider).pos \ stops  ' determine pos.
    iV_DrawImage IntImages, IMAGE_SLIDER_BIGBUT, x + 3 + sx, y + 3                             'draw amount
End Sub

' ////////////////////////////////////////////////////////////////////////////
' show text written on its side.
Private Sub displayTextAt270(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim fx As Long, fy As Long, psLab As WIDGET, cache As DisplayTextOptionCache

    Set psLab = psWidget                ' (W_LABEL *)psWidget

    ' Any widget using displayTextAt270 must have its pUserData initialized to a (DisplayTextOptionCache*)
    ' (assert(psWidget->pUserData != nullptr): the C assert, off in a release build)
    Set cache = psWidget.pUserData

    ' TODO: Only works for single-line (not "formatted text") labels
    cache.wzText.setText psLab.getString(), font_large

    fx = xOffset + psWidget.x()
    fy = yOffset + psWidget.y() + cache.wzText.width()

    cache.wzText.render fx + 2, fy + 2, psPalette(WZCOL_GREY), 270!
    cache.wzText.render fx, fy, psPalette(WZCOL_TEXT_BRIGHT), 270!
End Sub

' ////////////////////////////////////////////////////////////////////////////
' show a text option.
Public Sub displayTextOption(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim fx As Long, fy As Long, fw As Long, psBut As WIDGET, hilight As Boolean, greyOut As Boolean
    Dim cache As DisplayTextOptionCache, fontID As Long, result As iV_fonts, colour As Long
    Set psBut = psWidget                ' (W_BUTTON *)psWidget
    hilight = False
    greyOut = (psWidget.UserData <> 0) Or ((psBut.getState() And WBUT_DISABLE) <> 0) ' if option is unavailable.

    ' Any widget using displayTextOption must have its pUserData initialized to a (DisplayTextOptionCache*)
    ' (assert(psWidget->pUserData != nullptr): the C assert, off in a release build)
    Set cache = psWidget.pUserData

    fontID = cache.wzText.getFontID()
    If fontID = font_count Or cache.lastWidgetWidth <> psBut.width() Or cache.wzText.getText() <> W_BUTTON_(psBut).pText Then
        fontID = W_BUTTON_(psBut).FontID
    End If
    cache.wzText.setText W_BUTTON_(psBut).pText, fontID

    If cache.wzText.width() > psBut.width() Then
        ' text would exceed the bounds of the button area
        Select Case cache.overflowBehavior_
        Case OverflowBehavior_ShrinkFont
            Do
                If fontID = font_small Or fontID = font_bar Or fontID = font_regular Or fontID = font_regular_bold Then
                    Exit Do
                End If
                If Not iV_ShrinkFont(fontID, result) Then
                    Exit Do
                End If
                fontID = result
                cache.wzText.setText W_BUTTON_(psBut).pText, fontID
            Loop While cache.wzText.width() > psBut.width()
        Case Else
        End Select
    End If
    cache.lastWidgetWidth = psBut.width()

    If psBut.isMouseOverWidget() Then ' if mouse is over text then hilight.
        hilight = True
    End If

    fw = cache.wzText.width()
    fy = yOffset + psWidget.y() + (psWidget.height() - cache.wzText.lineSize()) \ 2 - cache.wzText.aboveBase()

    If (psWidget.style And WBUT_TXTCENTRE) <> 0 Then                         'check for centering, calculate offset.
        fx = xOffset + psWidget.x() + ((psWidget.width() - fw) \ 2)
    Else
        fx = xOffset + psWidget.x()
    End If

    If greyOut Then                                                     ' unavailable
        colour = psPalette(WZCOL_TEXT_DARK)
    Else                                                            ' available
        If hilight Then                                                 ' hilight
            colour = psPalette(WZCOL_TEXT_BRIGHT)
        ElseIf psWidget.id = FRONTEND_HYPERLINK Or psWidget.id = FRONTEND_DONATELINK Or psWidget.id = FRONTEND_CHATLINK Then ' special case for our hyperlink
            colour = psPalette(WZCOL_YELLOW)
        Else                                                        ' don't highlight
            colour = psPalette(WZCOL_TEXT_MEDIUM)
        End If
    End If

    cache.wzText.render fx, fy, colour
End Sub


' ////////////////////////////////////////////////////////////////////////////
' ////////////////////////////////////////////////////////////////////////////
' ////////////////////////////////////////////////////////////////////////////
' common widgets.

' W_FORM *addBackdrop() and addBackdrop(const std::shared_ptr<W_SCREEN> &screen): screen missing is the
' first (psWScreen); Nothing is nullptr
Public Function addBackdrop(Optional ByVal screen As Variant) As WIDGET
    Dim scr As W_SCREEN, sFormInit As W_FORMINIT
    If IsMissing(screen) Then Set scr = psWScreen Else Set scr = screen
    If Not (Not scr Is Nothing) Then ASSERT_FAILURE "screen != nullptr", "src/frontend.cpp:849", "addBackdrop", "Invalid screen pointer": Exit Function
    ' W_FORMINIT sFormInit;                              // Backdrop (the default constructor: all zero)
    sFormInit.formID = 0
    sFormInit.id = FRONTEND_BACKDROP
    sFormInit.style = WFORM_PLAIN
    Set sFormInit.calcLayout = WzFunc_new(FN_backdropCalcLayout)
    Set sFormInit.pDisplay = WzFunc_new(FN_displayTitleBitmap)
    Set sFormInit.pUserData = New TitleBitmapCache
    Set sFormInit.onDelete = WzFunc_new(FN_backdropOnDelete)

    Set addBackdrop = widgAddForm(scr, sFormInit)
End Function

' ////////////////////////////////////////////////////////////////////////////
Public Sub addTopForm(ByVal wide As Boolean)
    Dim parent As WIDGET, topForm As WIDGET, sFormInit As W_FORMINIT
    Dim imgW As Long, imgH As Long, dstW As Long, dstH As Long
    Set parent = widgGetFromID(psWScreen, FRONTEND_BACKDROP)
    If Not (Not parent Is Nothing) Then ASSERT_FAILURE "parent != nullptr", "src/frontend.cpp:875", "addTopForm", "Unable to find FRONTEND_BACKDROP?"

    Set topForm = IntFormTransparent_new()
    parent.attach topForm
    topForm.id = FRONTEND_TOPFORM
    If wide Then
        topForm.setCalcLayout WzFunc_new(FN_topFormCalcLayoutWide)
    Else
        topForm.setCalcLayout WzFunc_new(FN_topFormCalcLayout)
    End If

    ' W_FORMINIT sFormInit; (the default constructor: all zero)
    sFormInit.formID = FRONTEND_TOPFORM
    sFormInit.id = FRONTEND_LOGO
    imgW = iV_GetImageWidth(FrontImages, IMAGE_FE_LOGO)
    imgH = iV_GetImageHeight(FrontImages, IMAGE_FE_LOGO)
    dstW = topForm.width()
    dstH = topForm.height()
    If imgW * dstH < imgH * dstW Then ' Want to set aspect ratio dstW/dstH = imgW/imgH.
        dstW = imgW * dstH \ imgH ' Too wide.
    ElseIf imgW * dstH > imgH * dstW Then
        dstH = imgH * dstW \ imgW ' Too high.
    End If
    sFormInit.x = (topForm.width() - dstW) \ 2
    sFormInit.y = (topForm.height() - dstH) \ 2
    sFormInit.width = dstW
    sFormInit.height = dstH
    Set sFormInit.pDisplay = WzFunc_new(FN_displayLogo)
    widgAddForm psWScreen, sFormInit
End Sub

' ////////////////////////////////////////////////////////////////////////////
Public Sub addBottomForm(Optional ByVal wide As Boolean = False)
    Dim parent As WIDGET, botForm As WIDGET
    Set parent = widgGetFromID(psWScreen, FRONTEND_BACKDROP)

    Set botForm = IntFormAnimated_new()
    parent.attach botForm
    botForm.id = FRONTEND_BOTFORM

    If wide Then
        botForm.setCalcLayout WzFunc_new(FN_botFormCalcLayoutWide)
    Else
        botForm.setCalcLayout WzFunc_new(FN_botFormCalcLayout)
    End If
End Sub

' ////////////////////////////////////////////////////////////////////////////
Public Sub addText(ByVal id As Long, ByVal PosX As Long, ByVal PosY As Long, ByVal txt As String, ByVal formID As Long)
    Dim parent As WIDGET, label As WIDGET
    Set parent = widgGetFromID(psWScreen, formID)
    If Not (Not parent Is Nothing) Then ASSERT_FAILURE "parent != nullptr", "src/frontend.cpp:943", "addText", "Unable to find formID: %u", formID

    Set label = W_LABEL_new()
    parent.attach label
    label.id = id
    label.setGeometry PosX, PosY, MULTIOP_READY_WIDTH, FRONTEND_BUTHEIGHT
    W_LABEL_setTextAlignment label, WLAB_ALIGNCENTRE
    W_LABEL_setFont label, font_small, psPalette(WZCOL_TEXT_BRIGHT)
    label.setString txt
End Sub

' ////////////////////////////////////////////////////////////////////////////
' std::shared_ptr<W_LABEL> addSideText(WIDGET* psParent, UDWORD id, UDWORD PosX, UDWORD PosY, const char *txt)
Public Function addSideText_parent(ByVal psParent As WIDGET, ByVal id As Long, ByVal PosX As Long, ByVal PosY As Long, ByVal txt As String) As WIDGET
    Dim sLabInit As W_LABINIT, psLabel As WIDGET
    If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/frontend.cpp:957", "addSideText", "Invalid parent": Exit Function

    ' W_LABINIT sLabInit; (the default constructor: all zero, font_regular)

    sLabInit.formID = 0
    sLabInit.id = id
    sLabInit.x = PosX
    sLabInit.y = PosY
    sLabInit.width = 30
    sLabInit.height = FRONTEND_BOTFORMH

    sLabInit.FontID = font_large

    Set sLabInit.pDisplay = WzFunc_new(FN_displayTextAt270)
    sLabInit.pText = txt
    Set sLabInit.pUserData = New DisplayTextOptionCache
    Set sLabInit.onDelete = WzFunc_new(FN_deleteDisplayTextOptionCache)

    Set psLabel = W_LABEL_new_init(sLabInit)
    psLabel.setTransparentToClicks True
    psLabel.setTransparentToMouse True
    psParent.attach psLabel
    Set addSideText_parent = psLabel
End Function

' W_LABEL *addSideText(const std::shared_ptr<W_SCREEN> &psScreen, UDWORD formId, UDWORD id, UDWORD PosX, UDWORD PosY, const char *txt)
Public Function addSideText_screen(ByVal psScreen As W_SCREEN, ByVal formId As Long, ByVal id As Long, ByVal PosX As Long, ByVal PosY As Long, ByVal txt As String) As WIDGET
    Dim psParent As WIDGET
    If Not (Not psScreen Is Nothing) Then ASSERT_FAILURE "psScreen != nullptr", "src/frontend.cpp:988", "addSideText", "Invalid screen pointer": Exit Function

    Set psParent = widgGetFromID(psScreen, formId)
    Set addSideText_screen = addSideText_parent(psParent, id, PosX, PosY, txt)
End Function

' W_LABEL *addSideText(UDWORD id, UDWORD PosX, UDWORD PosY, const char *txt)
Public Function addSideText(ByVal id As Long, ByVal PosX As Long, ByVal PosY As Long, ByVal txt As String) As WIDGET
    Set addSideText = addSideText_screen(psWScreen, FRONTEND_BACKDROP, id, PosX, PosY, txt)
End Function

' ////////////////////////////////////////////////////////////////////////////
' optional<unsigned int> minimumWidth: has_minimumWidth and minimumWidth (nullopt by default)
Private Function makeTextButton(ByVal id As Long, ByVal txt As String, ByVal style As Long, Optional ByVal has_minimumWidth As Boolean = False, Optional ByVal minimumWidth As Long = 0) As WIDGET
    Dim sButInit As W_BUTINIT, textWidth As Long, pUserData As DisplayTextOptionCache, button As WIDGET, minW As Long

    ' W_BUTINIT sButInit; (the default constructor: all zero, font_regular)
    sButInit.formID = FRONTEND_BOTFORM
    sButInit.id = id

    ' Align
    sButInit.FontID = font_large
    If (style And WBUT_TXTCENTRE) = 0 Then
        textWidth = iV_GetTextWidth(txt, sButInit.FontID)
        If has_minimumWidth Then minW = minimumWidth Else minW = 0
        If ULT(minW, textWidth) Then sButInit.width = textWidth Else sButInit.width = minW
    Else
        sButInit.style = sButInit.style Or WBUT_TXTCENTRE
    End If

    ' Enable right clicks
    If (style And WBUT_SECONDARY) <> 0 Then
        sButInit.style = sButInit.style Or WBUT_SECONDARY
    End If

    sButInit.UserData = (style And WBUT_DISABLE) ' store disable state
    Set pUserData = New DisplayTextOptionCache
    pUserData.overflowBehavior_ = OverflowBehavior_ShrinkFont
    Set sButInit.pUserData = pUserData
    Set sButInit.onDelete = WzFunc_new(FN_deleteDisplayTextOptionCache)

    sButInit.height = FRONTEND_BUTHEIGHT
    Set sButInit.pDisplay = WzFunc_new(FN_displayTextOption)
    sButInit.pText = txt

    Set button = W_BUTTON_new_init(sButInit)
    ' Disable button
    If (style And WBUT_DISABLE) <> 0 Then
        button.setState WBUT_DISABLE
    End If

    Set makeTextButton = button
End Function

Public Function addTextButton(ByVal id As Long, ByVal PosX As Long, ByVal PosY As Long, ByVal txt As String, ByVal style As Long) As WIDGET
    Dim button As WIDGET, parent As WIDGET
    Set button = makeTextButton(id, txt, style)
    If (style And WBUT_TXTCENTRE) <> 0 Then
        button.setGeometry PosX, PosY, FRONTEND_BUTWIDTH, button.height()
    Else
        button.move PosX + 35, PosY
    End If

    Set parent = widgGetFromID(psWScreen, FRONTEND_BOTFORM)
    If Not (Not parent Is Nothing) Then ASSERT_FAILURE "parent != nullptr", "src/frontend.cpp:1062", "addTextButton", "Unable to find FRONTEND_BOTFORM?"
    parent.attach button
    Set addTextButton = button
End Function

Public Function addSmallTextButton(ByVal id As Long, ByVal PosX As Long, ByVal PosY As Long, ByVal txt As String, ByVal style As Long) As WIDGET
    Dim sButInit As W_BUTINIT, pButton As WIDGET

    ' W_BUTINIT sButInit; (the default constructor: all zero, font_regular)
    sButInit.formID = FRONTEND_BOTFORM
    sButInit.id = id
    sButInit.x = PosX
    sButInit.y = PosY

    ' Align
    If (style And WBUT_TXTCENTRE) = 0 Then
        sButInit.width = (iV_GetTextWidth(txt, font_small) + 4) And &HFFFF& ' (uint16_t)
        sButInit.x = sButInit.x + 35
    Else
        sButInit.style = sButInit.style Or WBUT_TXTCENTRE
        sButInit.width = FRONTEND_BUTWIDTH
    End If

    ' Enable right clicks
    If (style And WBUT_SECONDARY) <> 0 Then
        sButInit.style = sButInit.style Or WBUT_SECONDARY
    End If

    sButInit.UserData = (style And WBUT_DISABLE) ' store disable state
    Set sButInit.pUserData = New DisplayTextOptionCache
    Set sButInit.onDelete = WzFunc_new(FN_deleteDisplayTextOptionCache)

    sButInit.height = FRONTEND_BUTHEIGHT
    Set sButInit.pDisplay = WzFunc_new(FN_displayTextOption)
    sButInit.FontID = font_small
    sButInit.pText = txt
    Set pButton = widgAddButton(psWScreen, sButInit)

    ' Disable button
    If (style And WBUT_DISABLE) <> 0 Then
        widgSetButtonState psWScreen, id, WBUT_DISABLE
    End If
    Set addSmallTextButton = pButton
End Function

' ////////////////////////////////////////////////////////////////////////////
Private Function makeFESlider(ByVal id As Long, ByVal parent As Long, ByVal stops As Long, ByVal pos As Long) As WIDGET
    Dim sSldInit As W_SLDINIT
    W_SLDINIT_construct sSldInit        ' W_SLDINIT sSldInit;
    sSldInit.formID = parent
    sSldInit.id = id
    sSldInit.width = iV_GetImageWidth(IntImages, IMAGE_SLIDER_BIG)
    sSldInit.height = iV_GetImageHeight(IntImages, IMAGE_SLIDER_BIG)
    sSldInit.numStops = stops And &HFF&     ' (UBYTE)
    sSldInit.barSize = iV_GetImageHeight(IntImages, IMAGE_SLIDER_BIG)
    sSldInit.pos = pos And &HFF&            ' (UBYTE)
    Set sSldInit.pDisplay = WzFunc_new(FN_displayBigSlider)
    Set sSldInit.pCallback = WzFunc_new(FN_intUpdateQuantitySlider)

    Set makeFESlider = W_SLIDER_new_init(sSldInit)
End Function

Public Function addFESlider(ByVal id As Long, ByVal parent As Long, ByVal x As Long, ByVal y As Long, ByVal stops As Long, ByVal pos As Long) As WIDGET
    Dim slider As WIDGET
    Set slider = makeFESlider(id, parent, stops, pos)
    slider.move x, y
    widgGetFromID(psWScreen, parent).attach slider
    Set addFESlider = slider
End Function

' ////////////////////////////////////////////////////////////////////////////
' Change Mode
Public Sub changeTitleMode(ByVal mode As tMode)
    changeTitleUI WzOldTitleUI_new(mode)
End Sub

' ////////////////////////////////////////////////////////////////////////////
' Handling Screen Size Changes

' Tell the frontend when the screen has been resized
Public Sub frontendScreenSizeDidChange(ByVal oldWidth As Long, ByVal oldHeight As Long, ByVal newWidth As Long, ByVal newHeight As Long)
    Dim current As WzTitleUI, currentDisplayScale As Long, suggestedDisplayScale As Long
    Dim notification As WZ_Notification, contextText As String
    ' NOTE:
    ' By setting the appropriate calcLayout functions on all interface elements,
    ' they should automatically recalculate their layout on screen resize.
    If Not wzTitleUICurrent Is Nothing Then
        Set current = wzTitleUICurrent
        current.screenSizeDidChange oldWidth, oldHeight, newWidth, newHeight
    End If

    ' Determine if there should be a prompt to increase display scaling
    currentDisplayScale = wzGetCurrentDisplayScale()
    If (newWidth >= oldWidth And newHeight >= oldHeight) _
       And (lastProcessedScreenWidth <> newWidth Or lastProcessedScreenHeight <> newHeight) _
       And (lastProcessedDisplayScale = 0 Or lastProcessedDisplayScale = currentDisplayScale) _
       And (newWidth >= WZ_SUGGESTED_MAX_LOGICAL_SCREEN_SIZE_DIMENSION Or newHeight >= WZ_SUGGESTED_MAX_LOGICAL_SCREEN_SIZE_DIMENSION) Then
        ' filter out duplicate events; filter out duplicate events & display-scale-only changes
        suggestedDisplayScale = wzGetSuggestedDisplayScaleForCurrentWindowSize(WZ_SUGGESTED_MAX_LOGICAL_SCREEN_SIZE_DIMENSION)
        If ULT(currentDisplayScale, suggestedDisplayScale) _
           And Not hasNotificationsWithTag(DISPLAY_SCALE_PROMPT_INCREASE_TAG, NotificationScope_DISPLAYED_ONLY) Then
            cancelOrDismissNotificationIfTag WzFunc_new(FN_displayScalePromptTagIs)

            notification.duration = 0
            notification.contentTitle = "Increase Game Display Scale?"
            contextText = vbLf
            contextText = contextText & "With your current resolution & display scale settings, the game's user interface may appear small, and the game perspective may appear distorted."
            contextText = contextText & vbLf & vbLf
            contextText = contextText & "You can fix this by increasing the game's Display Scale setting."
            contextText = contextText & vbLf & vbLf
            contextText = contextText & va("Would you like to increase the game's Display Scale to: %u%%?", suggestedDisplayScale)
            notification.contentText = contextText
            notification.action_title = "Increase Display Scale"
            Set notification.action_onAction = WzFunc_new(FN_displayScaleIncreaseAction)
            Set notification.onDismissed = WzFunc_new(FN_displayScalePromptDismissed)
            notification.tag = DISPLAY_SCALE_PROMPT_INCREASE_TAG
            notification.displayOptions_identifier = "displayscale::prompt_increase"   ' makeIgnorable(..., 4)
            notification.displayOptions_numTimesSeenBeforeDoNotShowAgainOption = 4

            addNotification notification, GAME_TICKS_PER_SEC * 2
        End If
    End If
    lastProcessedDisplayScale = currentDisplayScale
    lastProcessedScreenWidth = newWidth
    lastProcessedScreenHeight = newHeight
End Sub

' the "Increase Display Scale" action of frontendScreenSizeDidChange's notification
Private Sub displayScaleIncreaseAction()
    Dim desiredDisplayScale As Long, priorDisplayScale As Long, completed_notification As WZ_Notification
    ' Determine maximum display scale for current window size
    desiredDisplayScale = wzGetSuggestedDisplayScaleForCurrentWindowSize(WZ_SUGGESTED_MAX_LOGICAL_SCREEN_SIZE_DIMENSION)
    If desiredDisplayScale = wzGetCurrentDisplayScale() Then
        ' nothing to do now
        Exit Sub
    End If

    ' Store the new display scale
    priorDisplayScale = war_GetDisplayScale()
    war_SetDisplayScale desiredDisplayScale

    If wzChangeDisplayScale(desiredDisplayScale) Then
        completed_notification.duration = 6 * GAME_TICKS_PER_SEC
        completed_notification.contentTitle = va("Display Scale Increased to: %u%%", desiredDisplayScale)
        completed_notification.contentText = "You can adjust the Display Scale at any time in the Video Options menu."
        completed_notification.tag = DISPLAY_SCALE_PROMPT_TAG_PREFIX & "completed"
        addNotification completed_notification, 0     ' WZ_Notification_Trigger::Immediate()
    Else
        ' failed to change display scale - restore the prior setting
        war_SetDisplayScale priorDisplayScale
    End If
End Sub

' the onDismissed of frontendScreenSizeDidChange's notification
Private Sub displayScalePromptDismissed(ByVal reason As Long)
    Dim dismissed_notification As WZ_Notification
    If reason <> WZ_Notification_Dismissal_Reason_USER_DISMISSED Then Exit Sub
    dismissed_notification.duration = 0
    dismissed_notification.contentTitle = "Tip: Adjusting Display Scale"
    dismissed_notification.contentText = "You can adjust the Display Scale at any time in the Video Options menu."
    dismissed_notification.tag = DISPLAY_SCALE_PROMPT_TAG_PREFIX & "dismissed_info"
    dismissed_notification.displayOptions_identifier = "displayscale::tip"     ' WZ_Notification_Display_Options::makeOneTime(...)
    dismissed_notification.displayOptions_oneTime = True
    addNotification dismissed_notification, 0         ' WZ_Notification_Trigger::Immediate()
End Sub

' To be called from frontendShutdown(), which for some reason is in init.cpp...
Public Sub frontendIsShuttingDown()
    ' std::weak_ptr<IMAGEFILE> pFlagsImagesWeak(pFlagsImages); pFlagsImages.reset();
    ' (VB) the handle has one owner, so the IMAGEFILE goes with it and the ASSERT(pFlagsImagesWeak.expired(), ...)
    ' holds: its users call getFlagsImages() each time and keep no copy
    If pFlagsImages <> 0 Then
        iV_FreeImageFile pFlagsImages
        pFlagsImages = 0
    End If
End Sub

'--------------------------------------------------------------------------
' (VB) this module's lambdas and function pointers (WzFunc ids FN_ above): args are the C++ call's
' arguments (a widget's display function: the widget, xOffset, yOffset; calcLayout and onDelete: the
' widget)
Public Function frontend_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim psWidget As WIDGET, ctx As W_CONTEXT
    Select Case f.id
    Case FN_backdropCalcLayout
        Set psWidget = args(0)
        psWidget.setGeometry (pie_GetVideoBufferWidth() - HIDDEN_FRONTEND_WIDTH) \ 2, _
                             (pie_GetVideoBufferHeight() - HIDDEN_FRONTEND_HEIGHT) \ 2, _
                             HIDDEN_FRONTEND_WIDTH - 1, _
                             HIDDEN_FRONTEND_HEIGHT - 1
    Case FN_backdropOnDelete, FN_deleteDisplayTextOptionCache
        ' (assert(psWidget->pUserData != nullptr): the C assert, off in a release build)
        Set psWidget = args(0)
        psWidget.pUserData = Empty      ' delete ((TitleBitmapCache *)psWidget->pUserData); psWidget->pUserData = nullptr;
    Case FN_displayTitleBitmap
        displayTitleBitmap args(0), args(1), args(2)
    Case FN_topFormCalcLayoutWide
        Set psWidget = args(0)
        psWidget.setGeometry FRONTEND_TOPFORM_WIDEX, FRONTEND_TOPFORM_WIDEY, FRONTEND_TOPFORM_WIDEW, FRONTEND_TOPFORM_WIDEH
    Case FN_topFormCalcLayout
        Set psWidget = args(0)
        psWidget.setGeometry FRONTEND_TOPFORMX, FRONTEND_TOPFORMY, FRONTEND_TOPFORMW, FRONTEND_TOPFORMH
    Case FN_botFormCalcLayoutWide
        Set psWidget = args(0)
        psWidget.setGeometry FRONTEND_BOTFORM_WIDEX, FRONTEND_BOTFORM_WIDEY, FRONTEND_BOTFORM_WIDEW, FRONTEND_BOTFORM_WIDEH
    Case FN_botFormCalcLayout
        Set psWidget = args(0)
        psWidget.setGeometry FRONTEND_BOTFORMX, FRONTEND_BOTFORMY, FRONTEND_BOTFORMW, FRONTEND_BOTFORMH
    Case FN_displayLogo
        displayLogo args(0), args(1), args(2)
    Case FN_displayBigSlider
        displayBigSlider args(0), args(1), args(2)
    Case FN_displayTextAt270
        displayTextAt270 args(0), args(1), args(2)
    Case FN_displayTextOption
        displayTextOption args(0), args(1), args(2)
    Case FN_intUpdateQuantitySlider
        ' WIDGET_CALLBACK (WIDGET *psWidget, W_CONTEXT *psContext): the context as its four fields
        If UBound(args) >= 4 Then ctx = W_CONTEXT_(args(1), args(2), args(3), args(4))
        intUpdateQuantitySlider args(0), ctx
    Case FN_missingVideosAction
        runHyperlink
    Case FN_displayScalePromptTagIs
        ' [](const std::string& tag) { return (tag.rfind(DISPLAY_SCALE_PROMPT_TAG_PREFIX, 0) == 0); }
        frontend_c_fn = (Left$(args(0), Len(DISPLAY_SCALE_PROMPT_TAG_PREFIX)) = DISPLAY_SCALE_PROMPT_TAG_PREFIX)
    Case FN_displayScaleIncreaseAction
        displayScaleIncreaseAction
    Case FN_displayScalePromptDismissed
        ' (const WZ_Notification&, WZ_Notification_Dismissal_Reason reason)
        displayScalePromptDismissed args(1)
    End Select
End Function
