Attribute VB_Name = "wrappers_c"
' src/wrappers.cpp, wrappers.h (ported in part): the host launch and
' headless settings (setHostLaunch, getHostLaunch, resetHostLaunch,
' setHeadlessGameMode, headlessGameMode, setHostLaunchStartNotReady,
' getHostLaunchStartNotReady), frontendInitVars, titleLoop (its host
' launch opens the multiplayer options screen, WzMultiplayerOptionsTitleUI),
' displayGameOver, the win and lose flags (testPlayerHasLost/Won,
' setPlayerHasLost/Won), the script's win/lose video and the loading
' screen (initLoadingScreen, closeLoadingScreen, loadingScreenCallback).
Option Explicit

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

' lib/netplay/nettypes.h
Private Const MAX_CONNECTED_PLAYERS As Long = MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS

' wrappers.h: enum TITLECODE
Public Enum TITLECODE
    TITLECODE_CONTINUE
    TITLECODE_STARTGAME
    TITLECODE_QUITGAME
    TITLECODE_SHOWINTRO
    TITLECODE_SAVEGAMELOAD
End Enum

'used to set the scriptWinLoseVideo variable
Public Const PLAY_NONE As Long = 0
Public Const PLAY_WIN As Long = 1
Public Const PLAY_LOSE As Long = 2

' wrappers.h: enum class HostLaunch
Public Const HostLaunch_Normal As Long = 0
Public Const HostLaunch_Host As Long = 1
Public Const HostLaunch_Skirmish As Long = 2
Public Const HostLaunch_Autohost As Long = 3
Public Const HostLaunch_LoadReplay As Long = 4

Private firstcall As Boolean
Private bPlayerHasLost As Boolean
Private bPlayerHasWon As Boolean
Private scriptWinLoseVideo As Long      ' UBYTE (PLAY_NONE = 0)

Private hostlaunch As Long              ' HostLaunch: used to detect if we are hosting a game via command line option. (Normal = 0)
Private bHeadlessAutoGameModeCLIOption As Boolean
Private bActualHeadlessAutoGameMode As Boolean
Private bHostLaunchStartNotReady As Boolean

Private Type STAR
    xPos As Long
    speed As Long
    colour As Long                      ' PIELIGHT
End Type

Private lastTick As Long                ' uint32_t (= 0)
Private barLeftX As Long, barLeftY As Long, barRightX As Long, barRightY As Long, boxWidth As Long, boxHeight As Long, starsNum As Long, starHeight As Long
Private stars() As STAR                 ' (VB) starsAllocated: stars != nullptr
Private starsAllocated As Boolean

' (VB) resSetLoadCallback's function numbers (FrameworkGame's resLoadCallback calls them)
Public Const fn_loadingScreenCallback As Long = 1

Private Function recalculateEffectiveHeadlessValue() As Boolean
    If hostlaunch = HostLaunch_Skirmish Or hostlaunch = HostLaunch_Autohost Or hostlaunch = HostLaunch_LoadReplay Or autogame_enabled() Then
        ' only support headless mode if hostlaunch is --skirmish or --autogame
        recalculateEffectiveHeadlessValue = bHeadlessAutoGameModeCLIOption
        Exit Function
    End If
    recalculateEffectiveHeadlessValue = False
End Function

Public Sub setHostLaunch(ByVal value As Long)
    hostlaunch = value
    bActualHeadlessAutoGameMode = recalculateEffectiveHeadlessValue()
End Sub

Public Sub resetHostLaunch()
    setHostLaunch HostLaunch_Normal
    setHostLaunchStartNotReady False
End Sub

Public Function getHostLaunch() As Long
    getHostLaunch = hostlaunch
End Function

Public Sub setHeadlessGameMode(ByVal enabled As Boolean)
    bHeadlessAutoGameModeCLIOption = enabled
    bActualHeadlessAutoGameMode = recalculateEffectiveHeadlessValue()
End Sub

Public Function headlessGameMode() As Boolean
    headlessGameMode = bActualHeadlessAutoGameMode
End Function

Public Sub setHostLaunchStartNotReady(ByVal value As Boolean)
    bHostLaunchStartNotReady = value
End Sub

Public Function getHostLaunchStartNotReady() As Boolean
    If bHostLaunchStartNotReady And headlessGameMode() And Not wz_command_interface_enabled() Then
        debug_ 167, LOG_ERROR, "getHostLaunchStartNotReady", "--autohost-not-ready specified while in headless mode without --enablecmdinterface specified. No way to start the host. Ignoring."
        bHostLaunchStartNotReady = False
    End If
    getHostLaunchStartNotReady = bHostLaunchStartNotReady
End Function


' //////////////////////////////////////////////////////////////////
' Initialise frontend globals and statics.
'
Public Function frontendInitVars() As Boolean
    firstcall = True

    frontendInitVars = True
End Function

' ///////////////// /////////////////////////////////////////////////
' Main Front end game loop.
Public Function titleLoop() As TITLECODE
    Dim RetCode As TITLECODE, current As WzTitleUI
    RetCode = TITLECODE_CONTINUE

    pie_SetFogStatus False
    If Not headlessGameMode() Then
        If screen_RestartBackDrop() Then
            ' changed value - draw the backdrop
            ' otherwise, pie_ScreenFrameRenderBegin handles drawing it
            screen_Display
        End If
    End If
    wzShowMouse True

    ' When we first init the game, firstcall is true.
    If firstcall Then
        firstcall = False
        ' First check to see if --host was given as a command line option, if not,
        ' then check --join and if neither, run the normal game menu.
        If hostlaunch <> HostLaunch_Normal Then
            If hostlaunch = HostLaunch_Skirmish Then
                SPinit LEVEL_TYPE_SKIRMISH
            Else ' single player
                NetPlay_bComms = True ' use network = true
                bMultiMessages = True
                NETinit war_getHostConnectionProvider()
                NETinitPortMapping
            End If
            bMultiPlayer = True
            ingame.side = InGameSide_HOST_OR_SINGLEPLAYER
            game.typ = LEVEL_TYPE_SKIRMISH
            ' Ensure the game has a place to return to
            changeTitleMode TITLE
            changeTitleUI WzMultiplayerOptionsTitleUI_new(wzTitleUICurrent)
        ElseIf Len(iptoconnect) > 0 Then
            NetPlay_bComms = True ' use network = true
            ' Ensure the joinGame has a place to return to
            changeTitleMode TITLE
            ' Don't call `NETinit()` just yet.
            ' It will be automatically called by `joinGame()` upon connection attempt
            ' with the correct connection provider type corresponding to the connection string.
            joinGameFromIPOrHostnameConnectionStr iptoconnect, cliConnectAsSpectator
        ElseIf Len(cli_lobby_game_to_connect_str()) > 0 Then
            NetPlay_bComms = True ' use network = true
            ' Ensure the joinGame has a place to return to
            changeTitleMode TITLE
            ' Don't call `NETinit()` just yet.
            ' It will be automatically called upon connection attempt
            ' with the correct connection provider type corresponding to discovered connection info.
            joinLobbyGame NETgetLobbyserverAddress(), cli_lobby_game_to_connect_str(), cliConnectAsSpectator
        Else
            changeTitleMode TITLE          ' normal game, run main title screen.
        End If
        ' Using software cursors (when on) for these menus due to a bug in SDL's SDL_ShowCursor()
        wzSetCursor CURSOR_DEFAULT
    End If

    If Not wzTitleUICurrent Is Nothing Then
        ' Creates a pointer, so if... when, the UI changes during a run, this does not disappear
        Set current = wzTitleUICurrent
        RetCode = current.run()
    End If
    drawtest_frame                      ' (VB) --vbdrawtest

    If (RetCode = TITLECODE_SAVEGAMELOAD) Or (RetCode = TITLECODE_STARTGAME) Then
        titleLoop = RetCode ' don't flip
        Exit Function
    End If
    NETflush  ' Send any pending network data.

    audio_Update

    pie_SetFogStatus False

    If (keyDown(KEY_LALT) Or keyDown(KEY_RALT)) And keyPressed(KEY_RETURN) Then
        war_setWindowMode wzAltEnterToggleFullscreen()
    End If
    titleLoop = RetCode
End Function

Private Function newStar() As STAR
    Dim s As STAR
    s.xPos = rand() Mod barRightX
    s.speed = Int(((rand() Mod 30) + 6) * pie_GetVideoBufferWidth() / 640#)
    s.colour = pal_SetBrightness(150 + (rand() Mod 100))
    newStar = s
End Function

Private Sub setupLoadingScreen()
    Dim i As Long, w As Long, h As Long, offset As Long
    w = pie_GetVideoBufferWidth()
    h = pie_GetVideoBufferHeight()

    boxHeight = Int(h / 40#)
    offset = boxHeight
    boxWidth = w - 2 * offset

    barRightX = w - offset
    barRightY = h - offset

    barLeftX = barRightX - boxWidth
    barLeftY = barRightY - boxHeight

    If boxHeight > 1 Then starsNum = boxWidth \ boxHeight Else starsNum = boxWidth \ 1
    starHeight = Int(2# * h / 640#)

    If Not starsAllocated Then
        If starsNum <= 0 Then
            starsNum = 0
            Exit Sub
        End If
        ReDim stars(0 To starsNum - 1)
        starsAllocated = True
    End If

    For i = 0 To starsNum - 1
        stars(i) = newStar()
    Next
End Sub

' (VB) resSetLoadCallback(loadingScreenCallback): FrameworkGame's resLoadCallback calls it by its number
Public Sub loadingScreenCallback()
    Dim loadingbar_background As Long, currTick As Long, i As Long
    Dim topX As Long, topY As Long, botX As Long, botY As Long
    loadingbar_background = psPalette(WZCOL_LOADING_BAR_BACKGROUND)
    currTick = wzGetTicks()

    If U32(currTick) - U32(lastTick) < 50 Then
        Exit Sub
    End If

    lastTick = currTick

    ' Draw the black rectangle at the bottom, with a two pixel border
    pie_UniTransBoxFill barLeftX - 2, barLeftY - 2, barRightX + 2, barRightY + 2, loadingbar_background

    For i = 1 To starsNum - 1
        stars(i).xPos = stars(i).xPos + stars(i).speed
        If barLeftX + stars(i).xPos >= barRightX Then
            stars(i) = newStar()
            stars(i).xPos = 1
        End If
        topX = barLeftX + stars(i).xPos
        topY = barLeftY + i * (boxHeight - starHeight) \ starsNum
        botX = topX + stars(i).speed
        If barRightX < botX Then botX = barRightX
        botY = topY + starHeight

        pie_UniTransBoxFill topX, topY, botX, botY, stars(i).colour
    Next

    pie_ScreenFrameRenderEnd
    pie_ScreenFrameRenderBegin

    audio_Update

    wzPumpEventsWhileLoading
End Sub

' fill buffers with the static screen
Public Sub initLoadingScreen(ByVal drawbdrop As Boolean)
    pie_ScreenFrameRenderBegin ' start a frame *if one isn't yet started*
    setupLoadingScreen
    wzShowMouse False
    pie_SetFogStatus False

    ' setup the callback....
    resSetLoadCallback fn_loadingScreenCallback

    If drawbdrop And Not headlessGameMode() Then
        If Not screen_GetBackDrop() Then
            pie_LoadBackDrop SCREEN_RANDOMBDROP
        End If
        screen_RestartBackDrop
    Else
        screen_StopBackDrop
    End If
End Sub

' shut down the loading screen
Public Sub closeLoadingScreen()
    If starsAllocated Then
        Erase stars
        starsAllocated = False
    End If
    resSetLoadCallback 0
End Sub


'//////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////
' Gameover screen.

Public Function displayGameOver(ByVal bDidit As Boolean, ByVal showBackDrop As Boolean) As Boolean
    Dim isFirstCallForThisGame As Boolean, isGameFullyOver As Boolean, idx As Long, st As PLAYERSTATS_, freq As Currency
    isFirstCallForThisGame = Not testPlayerHasLost() And Not testPlayerHasWon()
    If bMultiPlayer Then
        ' This is a bit of a hack and partially relies upon the logic in endconditions.js
        isGameFullyOver = NetPlay_players_isSpectator(selectedPlayer) Or bDidit   ' gameOverMessage is only called for spectators when the game fully ends; can only win when the game actually ends :)
        If isGameFullyOver And Not ingame.has_endTime Then
            ingame.has_endTime = True
            QueryPerformanceCounter ingame.endTime
            QueryPerformanceFrequency freq
            debug_ 399, LOG_INFO, "displayGameOver", "Game ended (duration: %lld)", Fix(CDbl(ingame.endTime - ingame.startTime) / CDbl(freq))

            ' If in blind mode, send data on who the players were
            If game.blindMode <> BLIND_MODE_NONE Then
                If NetPlay_isHost Then
                    ' Send updated player info (which will include real names, now that game has ended) to all players
                    NETSendAllPlayerInfoTo NET_ALL_PLAYERS

                    ' Send the verified player identity from initial join for each player (now that game has ended)
                    For idx = 0 To MAX_CONNECTED_PLAYERS - 1
                        sendMultiStatsHostVerifiedIdentities idx
                    Next
                End If

                ' Note: Replay player info updating occurs as part of NETreplaySaveStop
            End If
        End If
    End If
    If bDidit Then
        setPlayerHasWon True
        multiplayerWinSequence True
        If bMultiPlayer Then
            updateMultiStatsWins
        End If
    Else
        setPlayerHasLost True
        If bMultiPlayer And isFirstCallForThisGame Then ' make sure we only accumulate one loss (even if this is called more than once, for example when losing initially, and then when the game fully ends)
            updateMultiStatsLoses
        End If
    End If
    If bMultiPlayer And isFirstCallForThisGame Then
        updateMultiStatsGames ' update games played.

        st = getMultiStats(selectedPlayer)
        saveMultiStats getPlayerName(selectedPlayer), getPlayerName(selectedPlayer), st
    End If

    'clear out any mission widgets - timers etc that may be on the screen
    clearMissionWidgets

    If bMultiPlayer And NetPlay_players_isSpectator(selectedPlayer) Then
        ' Special message for spectators to inform them that the game is fully over
        addConsoleMessage "GAME OVER", CENTRE_JUSTIFY, SYSTEM_MESSAGE, False, MAX_CONSOLE_MESSAGE_DURATION
        addConsoleMessage "The battle is over - you can leave the room.", CENTRE_JUSTIFY, SYSTEM_MESSAGE, False, MAX_CONSOLE_MESSAGE_DURATION
        ' TODO: Display this in a form with a "Quit to Main Menu" button?, or adapt intAddMissionResult to have a separate display for spectators?
    Else
        intAddMissionResult bDidit, True, showBackDrop
    End If

    displayGameOver = True
End Function


'//////////////////////////////////////////////////////////////////////////////
Public Function testPlayerHasLost() As Boolean
    testPlayerHasLost = (bPlayerHasLost)
End Function

Public Sub setPlayerHasLost(ByVal val_ As Boolean)
    bPlayerHasLost = val_
End Sub


'//////////////////////////////////////////////////////////////////////////////
Public Function testPlayerHasWon() As Boolean
    testPlayerHasWon = (bPlayerHasWon)
End Function

Public Sub setPlayerHasWon(ByVal val_ As Boolean)
    bPlayerHasWon = val_
End Sub

'access functions for scriptWinLoseVideo - used to indicate when the script is playing the win/lose video
Public Sub setScriptWinLoseVideo(ByVal val_ As Long)
    scriptWinLoseVideo = val_
End Sub

Public Function getScriptWinLoseVideo() As Long
    getScriptWinLoseVideo = scriptWinLoseVideo
End Function
