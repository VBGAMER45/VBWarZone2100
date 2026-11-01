Attribute VB_Name = "main_c"
' src/main.cpp (ported in part). realmain runs the parts whose subsystems
' exist: the debug log (stderr, then the log file in the config dir), the
' command line, PhysicsFS, the config dir and the user folders, the config
' file, the screen, the main loop and the shutdown. Ported: the globals,
' getDefaultLogFilePath, getPlatformPrefDir (the non-portable, non-packaged
' branch), getWzPlatformPrefDir, initialize_ConfigDir, initialize_PhysicsFS,
' make_dir, realmain (in part: no crash handling, shaders or screen
' setup), the title and game loops (startTitleLoop, stopTitleLoop,
' setCDAudioForCurrentGameMode, startGameLoop, stopGameLoop,
' initSaveGameLoad, runGameLoop, runTitleLoop), mainLoop, mainShutdown (in
' part), GetGameMode, SetGameMode. scanDataDirs, the command line's mods
' check; lib/ivis_opengl/gfx_api.h's inline to_swap_mode.
' Missing: crash handling, check_Physfs, portable mode, the update
' check (asyncGetCompatCheckResults), WzInfoManager, the favourite
' structures file, the command interface and the old logs' cleanup.
' optional<TITLECODE> queuedTilecodeChange is the value and
' hasQueuedTilecodeChange; executeFnAndProcessScriptQueuedRemovals([]{ f(); })
' is f followed by wzapi_processScriptQueuedObjectRemovals (as in loop_c).
Option Explicit

Private Const EXIT_FAILURE As Long = 1
Private Const BACKEND As String = "Win32"
Private Const PATH_MAX As Long = 260            ' MAX_PATH on Windows
Private Const WZ_DATADIR As String = "./data"   ' the Windows build's (CMAKE_INSTALL_DATAROOTDIR "." + "/data"), relative

' main.h: enum GS_GAMEMODE
Public Const GS_TITLE_SCREEN As Long = 0
Public Const GS_NORMAL As Long = 1
Public Const GS_SAVEGAMELOAD As Long = 2

Private Enum FOCUS_STATE
    FOCUS_OUT       ' Window does not have the focus
    FOCUS_IN        ' Window has got the focus
End Enum

Private Const MAX_CONNECTED_PLAYERS As Long = MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS

Public customDebugfile As Boolean       ' Default false: user has NOT specified where to store the stdout/err file.

Public datadir As String                ' Global that src/clparse.c:ParseCommandLine can write to, so it can override the default datadir on runtime. Needs to be empty on startup for ParseCommandLine to work!
Public configdir As String              ' specifies custom USER directory. Same rules apply as datadir above.
Public rulesettag As String             ' char[40]

' flag to indicate when initialisation is complete
Public Const SAVEGAME_CAM As String = "savegames/campaign"
Public Const SAVEGAME_CAM_AUTO As String = "savegames/campaign/auto"
Public Const SAVEGAME_SKI As String = "savegames/skirmish"
Public Const SAVEGAME_SKI_AUTO As String = "savegames/skirmish/auto"

Public gameInitialised As Boolean
Public SaveGamePath As String
Public ReplayPath As String
Public ScreenDumpPath As String
Public MultiCustomMapsPath As String
Public MultiPlayersPath As String
Public FavoriteStructuresPath As String
Private forcedAutosaveTime As Long      ' uint32_t
' Start game in title mode:
Private gameStatus As Long              ' GS_GAMEMODE (GS_TITLE_SCREEN = 0)
' Status of the gameloop
Private gameLoopStatus As GAMECODE      ' (GAMECODE_CONTINUE = 0)
Private focusState As FOCUS_STATE       ' (FOCUS_IN)

Private queuedTilecodeChange As TITLECODE
Private hasQueuedTilecodeChange As Boolean ' (nullopt)
Private staticsReady As Boolean

Private utfargc As Long
Private utfargv() As String
Private defaultLogFileName As String

' the C++ initialisers of the statics
Private Sub InitStatics()
    If staticsReady Then Exit Sub
    staticsReady = True
    focusState = FOCUS_IN
End Sub

' C: getDefaultLogFilePath (platformDirSeparator): logs\WZlog-MMDD_HHMMSS.txt, generated once per run
Private Function getDefaultLogFilePath(ByVal platformDirSeparator As String) As String
    Dim newtime As tm
    If Len(defaultLogFileName) = 0 Then ' only generate this once per run, so multiple callers get the same value
        newtime = getLocalTime(time_())   ' Convert time to struct
        defaultLogFileName = va("WZlog-%02d%02d_%02d%02d%02d.txt", _
            newtime.tm_mon + 1, newtime.tm_mday, newtime.tm_hour, newtime.tm_min, newtime.tm_sec)
    End If
    ' log name is logs/(or \)WZlog-MMDD_HHMMSS.txt
    getDefaultLogFilePath = "logs" & platformDirSeparator & defaultLogFileName
End Function

Private Function getPlatformPrefDir(ByVal org As String, ByVal app As String) As String
    Dim prefsDir As String
    ' (portable mode and packaged Windows apps: not ported)
    prefsDir = PHYSFS_getPrefDir(org, app)
    If StrPtr(prefsDir) = 0 Then
        debug_ 449, LOG_FATAL, "getPlatformPrefDir", "Failed to obtain prefs directory: %s", WZ_PHYSFS_getLastError()
        exit_ 1
    End If
    getPlatformPrefDir = prefsDir & PHYSFS_getDirSeparator()
End Function

Private Function getWzPlatformPrefDir() As String
    getWzPlatformPrefDir = getPlatformPrefDir("Warzone 2100 Project", version_getVersionedAppDirFolderName())
End Function

Private Sub initialize_ConfigDir()
    Dim configDir_ As String

    If Len(configdir) = 0 Then
        configDir_ = getWzPlatformPrefDir()
    Else
        configDir_ = configdir

        ' Make sure that we have a directory separator at the end of the string
        If Right$(configDir_, Len(PHYSFS_getDirSeparator())) <> PHYSFS_getDirSeparator() Then
            configDir_ = configDir_ & PHYSFS_getDirSeparator()
        End If

        debug_ 486, LOG_WZ, "initialize_ConfigDir", "Using custom configuration directory: %s", configDir_
    End If

    If PHYSFS_setWriteDir(configDir_) = 0 Then ' Workaround for PhysFS not creating the writedir as expected.
        debug_ 491, LOG_FATAL, "initialize_ConfigDir", "Error setting write directory to ""%s"": %s", _
            configDir_, WZ_PHYSFS_getLastError()
        exit_ 1
    End If

    ' (OverrideRPTDirectory: the crash handler is not ported)

    ' Config dir first so we always see what we write
    PHYSFS_mount PHYSFS_getWriteDir(), vbNullString, PHYSFS_PREPEND

    ' Do not follow symlinks *inside* search paths / archives
    PHYSFS_permitSymbolicLinks 0

    debug_ 509, LOG_WZ, "initialize_ConfigDir", "Write dir: %s", PHYSFS_getWriteDir()
    debug_ 510, LOG_WZ, "initialize_ConfigDir", "Base dir: %s", PHYSFS_getBaseDir()
End Sub

'!
' \brief Adds default data dirs
'
' Priority:
' Lower loads first. Current:
' --datadir > User's home dir > source tree data > AutoPackage > BaseDir > DEFAULT_DATADIR
'
' Only --datadir and home dir are always examined. Others only if data still not found.
'
' We need ParseCommandLine, before we can add any mods...
'
' \sa rebuildSearchPath
Private Sub scanDataDirs()
    Dim tmpstr As String, prefix As String, dirSeparator As String, wzDataDir As String, pCurrWriteDir As String

    ' Check PREFIX-based paths

    ' Find out which PREFIX we are in...
    prefix = getWZInstallPrefix()
    dirSeparator = PHYSFS_getDirSeparator()

    If PHYSFS_exists("gamedesc.lev") = 0 Then
        ' Data in source tree (<prefix>/data/)
        tmpstr = prefix & dirSeparator & "data" & dirSeparator
        registerSearchPath tmpstr, 3
        rebuildSearchPath mod_multiplay, True

        If PHYSFS_exists("gamedesc.lev") = 0 Then
            ' Guessed fallback default datadir on Unix
            wzDataDir = WZ_DATADIR
            If Len(wzDataDir) > 0 Then
                ' Treat WZ_DATADIR as a relative path - append to the install PREFIX
                tmpstr = prefix & dirSeparator & wzDataDir
                registerSearchPath tmpstr, 4
                rebuildSearchPath mod_multiplay, True
            End If

            If PHYSFS_exists("gamedesc.lev") = 0 Then
                ' Relocation for AutoPackage (<prefix>/share/warzone2100/)
                tmpstr = prefix & dirSeparator & "share" & dirSeparator & "warzone2100" & dirSeparator
                registerSearchPath tmpstr, 5
                rebuildSearchPath mod_multiplay, True
            End If
        End If
    End If

    ' Commandline supplied datadir
    If Len(datadir) <> 0 Then
        registerSearchPath datadir, 1
    End If

    ' User's home dir
    pCurrWriteDir = PHYSFS_getWriteDir()
    If StrPtr(pCurrWriteDir) <> 0 Then
        registerSearchPath pCurrWriteDir, 2
    Else
        debug_ 680, LOG_FATAL, "scanDataDirs", "No write dir set?"
    End If
    rebuildSearchPath mod_multiplay, True

    '* Debugging and sanity checks *

    printSearchPath

    If PHYSFS_exists("gamedesc.lev") <> 0 Then
        debug_ 690, LOG_WZ, "scanDataDirs", "gamedesc.lev found at %s", PHYSFS_getRealDir("gamedesc.lev") & ""
    Else
        debug_ 694, LOG_FATAL, "scanDataDirs", "Could not find game data. Aborting."
        exit_ 1
    End If
End Sub

' gfx_api.h
Private Function to_swap_mode(ByVal value As Long) As swap_interval_mode
    Select Case value
    Case swap_interval_mode_immediate
        to_swap_mode = swap_interval_mode_immediate
        Exit Function
    Case swap_interval_mode_vsync
        to_swap_mode = swap_interval_mode_vsync
        Exit Function
    Case swap_interval_mode_adaptive_vsync
        to_swap_mode = swap_interval_mode_adaptive_vsync
        Exit Function
    Case Else
        debug_ 1352, LOG_WARNING, "to_swap_mode", "Invalid vsync value (%d); defaulting to vsync ON", value
    End Select
    to_swap_mode = swap_interval_mode_vsync
End Function

Private Sub make_dir(dest As String, ByVal dirname As String, ByVal subdir As String)
    dest = dirname
    If StrPtr(subdir) <> 0 Then
        dest = dest & "/"
        dest = dest & subdir
    End If
    If Right$(dest, 1) <> "/" Then
        dest = dest & "/"
    End If
    If PHYSFS_mkdir(dest) = 0 Then
        debug_ 722, LOG_FATAL, "make_dir", "Unable to create directory ""%s"" in write dir ""%s""!", _
            dest, PHYSFS_getWriteDir()
        exit_ EXIT_FAILURE
    End If
End Sub

' Initialize the PhysicsFS library.
Private Sub initialize_PhysicsFS(ByVal argv_0 As String)
    Dim result As Long
    result = PHYSFS_init(argv_0)

    If result = 0 Then
        debug_ 523, LOG_FATAL, "initialize_PhysicsFS", "There was a problem trying to init Physfs.  Error was %s", WZ_PHYSFS_getLastError()
        exit_ -1
    End If
End Sub

' (VB) PLAYER::resetAll() of NetPlay.players[i], through NetApi (IPtextAddress is not kept)
Private Sub PLAYER_resetAll(ByVal i As Long)
    NetPlay_players_name(i) = ""
    NetPlay_players_position(i) = -1
    NetPlay_players_colour(i) = 0
    NetPlay_players_allocated(i) = False
    NetPlay_players_heartattacktime(i) = 0
    NetPlay_players_heartbeat(i) = False
    NetPlay_players_team(i) = -1
    NetPlay_players_ready(i) = False
    NetPlay_players_ai(i) = 0
    NetPlay_players_difficulty(i) = AIDifficulty_DISABLED
    NetPlay_players_autoGame(i) = False
    NetPlay_players_faction(i) = FACTION_NORMAL
    NetPlay_players_isSpectator(i) = False
    NetPlay_players_isAdmin(i) = False
End Sub

'
' Preparations before entering the title (mainmenu) loop
' Would start the timer in an event based mainloop
'
Private Sub startTitleLoop(Optional ByVal onInitialStartup As Boolean = False)
    SetGameMode GS_TITLE_SCREEN

    If Not onInitialStartup Then
        initLoadingScreen True
    End If
    If Not frontendInitialise("wrf/frontend.wrf") Then
        debug_ 743, LOG_FATAL, "startTitleLoop", "Shutting down after failure"
        exit_ EXIT_FAILURE
    End If

    closeLoadingScreen ' always ensure the loading screen is closed
End Sub


'
' Shutdown/cleanup after the title (mainmenu) loop
' Would stop the timer
'
Private Sub stopTitleLoop()
    If Not frontendShutdown() Then
        debug_ 759, LOG_FATAL, "stopTitleLoop", "Shutting down after failure"
        exit_ EXIT_FAILURE
    End If
End Sub

Private Sub setCDAudioForCurrentGameMode()
    Dim gameMode As ActivitySink_GameMode
    gameMode = ActivityManager_getCurrentGameMode()
    Select Case gameMode
    Case ActivitySink_GameMode_CAMPAIGN
        cdAudio_SetGameMode MusicGameMode_CAMPAIGN
    Case ActivitySink_GameMode_CHALLENGE
        cdAudio_SetGameMode MusicGameMode_CHALLENGE
    Case ActivitySink_GameMode_SKIRMISH
        cdAudio_SetGameMode MusicGameMode_SKIRMISH
    Case ActivitySink_GameMode_MULTIPLAYER
        cdAudio_SetGameMode MusicGameMode_MULTIPLAYER
    Case Else
        debug_ 782, LOG_ERROR, "setCDAudioForCurrentGameMode", "Unhandled started game mode for cd audio: %u", gameMode
    End Select
End Sub

'
' Preparations before entering the game loop
' Would start the timer in an event based mainloop
'
Private Function startGameLoop() As Boolean
    Dim hashStr As String, currentGameMode As ActivitySink_GameMode, replayOptions As WZGameReplayOptionsHandler
    Dim replayFilename As String, subdir As String
    If runningMultiplayer() Then
        ' Adjust connected timeouts for clients (which is typically larger than the value of timeout when
        ' waiting in a lobby room).
        NETadjustConnectedTimeoutForClients
    End If

    SetGameMode GS_NORMAL
    initLoadingScreen True

    ActivityManager_startingGame

    ' set up CD audio for game mode
    setCDAudioForCurrentGameMode

    ' Not sure what aLevelName is, in relation to game.map. But need to use aLevelName here, to be able to start the right map for campaign, and need game.hash, to start the right non-campaign map, if there are multiple identically named maps.
    If Not levLoadData(aLevelName, game.hash, True, vbNullString, GTYPE_SCENARIO_START) Then
        If Not Sha256_isZero(game.hash) Then hashStr = Sha256_toString(game.hash) Else hashStr = ""
        debug_ 811, LOG_ERROR, "startGameLoop", "Failed to load level data / map: %s %s", aLevelName, hashStr
        If bMultiPlayer Then
            multiGameShutdown
        End If
        levReleaseAll
        closeLoadingScreen
        cdAudio_SetGameMode MusicGameMode_MENUS
        stopGameLoop
        pie_LoadBackDrop SCREEN_RANDOMBDROP
        startTitleLoop ' Restart into titleloop
        gameLoopStatus = GAMECODE_CONTINUE
        startGameLoop = False
        Exit Function
    End If

    screen_StopBackDrop
    closeLoadingScreen

    ' Trap the cursor if cursor snapping is enabled
    If shouldTrapCursor() Then
        wzGrabMouse
    End If

    ' Disable resizable windows if it's a multiplayer game
    If runningMultiplayer() Then
        ' This is because the main loop gets frozen while the window resizing / edge dragging occurs
        ' which effectively pauses the game, and pausing is otherwise disabled in multiplayer games.
        ' FIXME: Figure out a workaround?
        wzSetWindowIsResizable False
    End If

    ' set a flag for the trigger/event system to indicate initialisation is complete
    gameInitialised = True

    If challengeActive Then
        addMissionTimerInterface
    End If
    triggerEvent TRIGGER_START_LEVEL
    wzapi_processScriptQueuedObjectRemovals
    screen_disableMapPreview

    ' Call once again to update starting counts (if modified by TRIGGER_START_LEVEL event)
    countUpdate False

    If Not bMultiPlayer Then
        If getCamTweakOption_AutosavesOnly() Then
            forcedAutosaveTime = gameTime + 1000 'Really just to prevent Intel videos messages from not getting saved if run immediately.
        End If
    End If

    GameStoryLogger_logStartGame

    currentGameMode = ActivityManager_getCurrentGameMode()
    Select Case currentGameMode
    Case ActivitySink_GameMode_MENUS
        ' should not happen
    Case ActivitySink_GameMode_CAMPAIGN
        ' replays not currently supported
    Case ActivitySink_GameMode_CHALLENGE, ActivitySink_GameMode_SKIRMISH, ActivitySink_GameMode_MULTIPLAYER
        ' start saving a replay
        If Not war_getDisableReplayRecording() Then
            Set replayOptions = New WZGameReplayOptionsHandler
            If currentGameMode = ActivitySink_GameMode_MULTIPLAYER Then subdir = "multiplay" Else subdir = "skirmish"
            replayFilename = NETreplaySaveStart(subdir, replayOptions, war_getMaxReplaysSaved(), (currentGameMode = ActivitySink_GameMode_MULTIPLAYER))
            If Len(replayFilename) > 0 Then
                wz_command_interface_output "WZEVENT: replaySaveStarted: %s\n", replayFilename
            End If
        End If
    Case Else
        debug_ 889, LOG_INFO, "startGameLoop", "Unhandled case: %u", currentGameMode
    End Select

    setMaxFastForwardTicks True, WZ_DEFAULT_MAX_FASTFORWARD_TICKS, True ' default value / spectator "catch-up" behavior
    If NETisReplay() Then
        If Not headlessGameMode() And Not autogame_enabled() Then
            ' for replays, ensure we don't start off fast-forwarding
            setMaxFastForwardTicks True, 0, True
        Else
            ' when loading replays in headless / autogame mode, set to fast-forward
            setMaxFastForwardTicks True, 10, False
        End If
    End If

    startGameLoop = True
End Function


'
' Shutdown/cleanup after the game loop
' Would stop the timer
'
Private Sub stopGameLoop()
    Dim replayOptions As WZGameReplayOptionsHandler, i As Long
    clearInfoMessages ' clear CONPRINTF messages before each new game/mission

    Set replayOptions = New WZGameReplayOptionsHandler
    NETreplaySaveStop replayOptions
    NETshutdownReplay

    If gameLoopStatus <> GAMECODE_NEWLEVEL Then
        clearBlueprints
        initLoadingScreen True ' returning to f.e. do a loader.render not active
        If gameLoopStatus <> GAMECODE_LOADGAME Then
            game.modHashCount = 0 ' must clear this before calling levReleaseAll so that when search paths are reloaded we don't reload mods downloaded for the last game (or loaded for the last replay)
            Erase game.modHashes
            If Not levReleaseAll() Then
                debug_ 932, LOG_ERROR, "stopGameLoop", "levReleaseAll failed!"
            End If
            For i = 0 To NetPlay_players_size() - 1
                PLAYER_resetAll i
            Next
            NetPlay_players_resize MAX_CONNECTED_PLAYERS
            ' (NetPlay.scriptSetPlayerDataStrings[i].clear(): the scripts' player data, not kept before cluster C)
        End If
        closeLoadingScreen
        reloadMPConfig
    End If

    ' Disable cursor trapping
    If shouldTrapCursor() Then
        wzReleaseMouse
    End If

    ' Re-enable resizable windows
    If wzGetCurrentWindowMode() = WINDOW_MODE_windowed Then
        ' FIXME: This is required because of the disabling in startGameLoop()
        wzSetWindowIsResizable True
    End If

    GameStoryLogger_reset

    gameInitialised = False
    forcedAutosaveTime = 0
End Sub


'
' Load a savegame and start into the game loop
' Game data should be initialised afterwards, so that startGameLoop is not necessary anymore.
'
Private Function initSaveGameLoad() As Boolean
    ' NOTE: always setGameMode correctly before *any* loading routines!
    SetGameMode GS_NORMAL
    initLoadingScreen True

    ' load up a save game
    If Not loadGameInit(GameLoadDetails_makeUserSaveGameLoad(saveGameName)) Then
        ' FIXME: we really should throw up a error window, but we can't (easily) so I won't.
        debug_ 982, LOG_ERROR, "initSaveGameLoad", "Trying to load Game %s failed!", saveGameName
        debug_ 983, LOG_POPUP, "initSaveGameLoad", "Failed to load a save game! It is either corrupted or a unsupported format.\n\nRestarting main menu."
        ' FIXME: If we bomb out on a in game load, then we would crash if we don't do the next two calls
        ' Doesn't seem to be a way to tell where we are in game loop to determine if/when we should do the two calls.
        gameLoopStatus = GAMECODE_FASTEXIT ' clear out all old data
        stopGameLoop
        startTitleLoop ' Restart into titleloop
        SetGameMode GS_TITLE_SCREEN
        initSaveGameLoad = False
        Exit Function
    End If

    ActivityManager_startingSavedGame

    ' set up CD audio for game mode
    ' (must come after savegame is loaded so that proper game mode can be determined)
    setCDAudioForCurrentGameMode

    screen_StopBackDrop
    closeLoadingScreen

    ' Trap the cursor if cursor snapping is enabled
    If shouldTrapCursor() Then
        wzGrabMouse
    End If
    If challengeActive Then
        addMissionTimerInterface
    End If

    ' set a flag for the trigger/event system to indicate initialisation is complete
    gameInitialised = True

    initSaveGameLoad = True
End Function


'
' Run the code inside the gameloop
'
Private Sub runGameLoop()
    ' Run the second half of a queued gameloopstatus change
    ' (This is needed so that the main loop completes once between initializing the loading screen and actually loading)
    Select Case gameLoopStatus
    Case GAMECODE_QUITGAME
        debug_ 1029, LOG_MAIN, "runGameLoop", "GAMECODE_QUITGAME"
        ActivityManager_quitGame Cheated  ' (and collectEndGameStatsData(): the end of game statistics, cluster J)
        cdAudio_SetGameMode MusicGameMode_MENUS
        stopGameLoop
        startTitleLoop ' Restart into titleloop
        gameLoopStatus = GAMECODE_CONTINUE
        Exit Sub
    Case GAMECODE_LOADGAME
        debug_ 1037, LOG_MAIN, "runGameLoop", "GAMECODE_LOADGAME"
        stopGameLoop
        initSaveGameLoad ' Restart and load a savegame
        gameLoopStatus = GAMECODE_CONTINUE
        Exit Sub
    Case GAMECODE_NEWLEVEL
        debug_ 1043, LOG_MAIN, "runGameLoop", "GAMECODE_NEWLEVEL"
        stopGameLoop
        If startGameLoop() Then ' Restart gameloop
            gameLoopStatus = GAMECODE_CONTINUE
        Else
            debug_ 1051, LOG_POPUP, "runGameLoop", "Failed to load level data or map. Exiting to main menu."
        End If
        Exit Sub
    Case Else
        ' ignore other values, and proceed with gameLoop
    End Select

    gameLoopStatus = gameLoop()
    Select Case gameLoopStatus
    Case GAMECODE_CONTINUE, GAMECODE_PLAYVIDEO
    Case GAMECODE_QUITGAME
        debug_ 1066, LOG_MAIN, "runGameLoop", "GAMECODE_QUITGAME"
        initLoadingScreen True
    Case GAMECODE_LOADGAME
        debug_ 1070, LOG_MAIN, "runGameLoop", "GAMECODE_LOADGAME"
        initLoadingScreen True
    Case GAMECODE_NEWLEVEL
        debug_ 1074, LOG_MAIN, "runGameLoop", "GAMECODE_NEWLEVEL"
        initLoadingScreen True
    ' Never thrown:
    Case GAMECODE_FASTEXIT, GAMECODE_RESTARTGAME
    Case Else
        debug_ 1082, LOG_ERROR, "runGameLoop", "Unknown code returned by gameLoop"
    End Select
End Sub


'
' Run the code inside the titleloop
'
Private Sub runTitleLoop()
    Dim toProcess As TITLECODE
    If hasQueuedTilecodeChange Then
        ' Run the second half of a queued titlecode change
        ' (This is needed so that the main loop completes once between initializing the loading screen and actually loading)
        toProcess = queuedTilecodeChange
        hasQueuedTilecodeChange = False
        Select Case toProcess
        Case TITLECODE_SAVEGAMELOAD
            debug_ 1105, LOG_MAIN, "runTitleLoop", "TITLECODE_SAVEGAMELOAD"
            ' Restart into gameloop and load a savegame, ONLY on a good savegame load!
            stopTitleLoop
            If Not initSaveGameLoad() Then
                ' we had a error loading savegame (corrupt?), so go back to title screen?
                stopGameLoop
                startTitleLoop
                changeTitleMode TITLE
            End If
            closeLoadingScreen
            Exit Sub
        Case TITLECODE_STARTGAME
            debug_ 1119, LOG_MAIN, "runTitleLoop", "TITLECODE_STARTGAME"
            stopTitleLoop
            If startGameLoop() Then ' Restart into gameloop
                closeLoadingScreen
            Else
                debug_ 1127, LOG_POPUP, "runTitleLoop", "Failed to load level data or map. Exiting to main menu."
            End If
            Exit Sub
        Case Else
            ' ignore unexpected value
        End Select
    End If

    Select Case titleLoop()
    Case TITLECODE_CONTINUE
    Case TITLECODE_QUITGAME
        debug_ 1141, LOG_MAIN, "runTitleLoop", "TITLECODE_QUITGAME"
        stopTitleLoop
        wzQuit 0
    Case TITLECODE_SAVEGAMELOAD
        debug_ 1147, LOG_MAIN, "runTitleLoop", "TITLECODE_SAVEGAMELOAD"
        initLoadingScreen True
        queuedTilecodeChange = TITLECODE_SAVEGAMELOAD
        hasQueuedTilecodeChange = True
    Case TITLECODE_STARTGAME
        debug_ 1153, LOG_MAIN, "runTitleLoop", "TITLECODE_STARTGAME"
        initLoadingScreen True
        queuedTilecodeChange = TITLECODE_STARTGAME
        hasQueuedTilecodeChange = True
    Case TITLECODE_SHOWINTRO
        debug_ 1158, LOG_MAIN, "runTitleLoop", "TITLECODE_SHOWINTRO"
        seq_ClearSeqList
        seq_AddSeqToList "titles.ogg", vbNullString, vbNullString, False
        seq_AddSeqToList "devastation.ogg", vbNullString, "devastation.txa", False
        seq_StartNextFullScreenVideo
    Case Else
        debug_ 1165, LOG_ERROR, "runTitleLoop", "Unknown code returned by titleLoop"
    End Select
End Sub

'
' The mainloop.
' Fetches events, executes appropriate code
'
Public Sub mainLoop()
    urlRequestPoll                      ' (VB) urlrequest's transfer thread: the URL requests' progress and callbacks
    wzAsyncExecOnMainThread_runPending  ' (VB) the SDL backend's WZ_SDL_MAINTHREAD_EXEC events, run before mainLoop
    frameUpdate ' General housekeeping
    vbscript_frame                       ' (VB) --vbscript: the UI test script's next command

    pie_ScreenFrameRenderBegin

    ' Screenshot key is now available globally
    If keyPressed(KEY_F10) Then
        kf_ScreenDump
        inputLoseFocus      ' remove it from input stream
    End If

    wzSetCursor CURSOR_DEFAULT ' if cursor isn't set by anything in the mainLoop, it should revert to default.

    If NetPlay_bComms Or focusState = FOCUS_IN Or Not war_GetPauseOnFocusLoss() Then
        If loop_GetVideoStatus() Then
            videoLoop ' Display the video if necessary
            pie_ScreenFrameRenderEnd
        Else
            Select Case GetGameMode()
            Case GS_NORMAL ' Run the gameloop code
                runGameLoop
                ' gameLoop handles pie_ScreenFrameRenderEnd()
            Case GS_TITLE_SCREEN ' Run the titleloop code
                runTitleLoop
                pie_ScreenFrameRenderEnd
            Case Else
            End Select
        End If
        realTimeUpdate ' Update realTime.
    End If

    If (forcedAutosaveTime <> 0) And ULT(forcedAutosaveTime, gameTime) Then
        forcedAutosaveTime = 0
        autoSave True
    End If

    wzApplyCursor
    runNotifications
    wz_command_interface_process_queued_status_output
End Sub

Public Sub mainShutdown()
    ActivityManager_preSystemShutdown

    Select Case GetGameMode()
    Case GS_NORMAL
        ' if running a game while quitting, stop the game loop
        ' (currently required for some cleanup) (should modelShutdown() be added to systemShutdown?)
        stopGameLoop
    Case GS_TITLE_SCREEN
        ' if showing the title / menus while quitting, stop the title loop
        ' (currently required for some cleanup)
        stopTitleLoop
    Case Else
    End Select
    saveConfig
    ' (writeFavoriteStructsFile, wzCmdInterfaceShutdown, cleanupOldLogFiles: later clusters)
    ' NOTE: urlRequestShutdown is called inside systemShutdown, as it must happen after certain other calls
    systemShutdown
    ActivityManager_shutdown
End Sub

Public Function realmain(ByVal argc As Long, argv() As String) As Long
    Dim debug_filename As String, exitCode As Long, earlyCommandLineParsingResult As ParseCLIEarlyResult
    Dim modtocheck As String, i As Long, globalModsPath As String, campaignModsPath As String, multiplayModsPath As String
    Dim metaData_filesize As Double, metaData_modtime As Double, metaData_filetype As Long, metaData_readonly As Long  ' PHYSFS_Stat metaData
    Dim horizScaleFactor As Long, vertScaleFactor As Long, fogConfigOption As Boolean
    InitStatics
    utfargc = argc
    utfargv = argv

    debug_init
    debug_register_callback CB_debug_callback_stderr, CB_NULL, CB_NULL, Empty

    ' *****
    ' NOTE: Try *NOT* to use debug() output routines without some other method of informing the user.
    ' All this output is sent to /dev/nul at this point on some platforms!
    ' *****
    If Not ParseCommandLineDebugFlags(utfargc, utfargv) Then
        realmain = EXIT_FAILURE
        Exit Function
    End If

    '** Initialize PhysicsFS **
    initialize_PhysicsFS utfargv(0)

    '** Initialize translations **
    '** NOTE: Should occur before any use of gettext / libintl translation routines. **
    initI18n

    wzMain ' init Qt integration first
    PlatformSetGame New PlatformGame       ' (VB) lib/sdl's calls into the game
    frameSetGame New FrameworkGame         ' (VB) lib/framework's calls into the game
    NetSetGame New NetGame                 ' (VB) lib/netplay's calls into the game
    SoundSetGame New SoundGame             ' (VB) lib/sound's calls into the game

    '** Initialize URL Request library **
    urlRequestInit

    ' find early boot info
    earlyCommandLineParsingResult = ParseCommandLineEarly(utfargc, utfargv)
    If earlyCommandLineParsingResult <> ParseCLIEarlyResult_OK_CONTINUE Then
        PlatformSetGame Nothing                ' (VB)
        frameSetGame Nothing                   ' (VB)
        NetSetGame Nothing                     ' (VB)
        SoundSetGame Nothing                   ' (VB)
        If earlyCommandLineParsingResult = ParseCLIEarlyResult_HANDLED_QUIT_EARLY_COMMAND Then realmain = 0 Else realmain = EXIT_FAILURE
        Exit Function
    End If

    ' Initialize the write/config directory for PhysicsFS.
    ' This needs to be done __after__ the early commandline parsing,
    ' because the user might tell us to use an alternative configuration
    ' directory.
    initialize_ConfigDir

    '** Initialize directory structure **

    PHYSFS_mkdir "autohost"    ' autohost games launched with --autohost=game

    PHYSFS_mkdir "challenges"  ' custom challenges

    PHYSFS_mkdir "logs"        ' netplay, mingw crash reports & WZ logs

    make_dir MultiCustomMapsPath, "maps", vbNullString ' needed to prevent crashes when getting map

    PHYSFS_mkdir version_getVersionedModsFolderPath("autoload")   ' mods that are automatically loaded
    PHYSFS_mkdir "mods/campaign"   ' campaign only mods - no longer versioned because of new campaign mod packaging and selector
    PHYSFS_mkdir "mods/downloads"  ' mod download directory - NOT currently versioned
    PHYSFS_mkdir version_getVersionedModsFolderPath("global")     ' global mods activated with --mod=example.wz
    PHYSFS_mkdir version_getVersionedModsFolderPath("multiplay")  ' multiplay only mods activated with --mod_mp=example.wz
    PHYSFS_mkdir version_getVersionedModsFolderPath("music")      ' music mods that are automatically loaded

    make_dir MultiPlayersPath, "multiplay", "players" ' player profiles

    PHYSFS_mkdir "music"           ' custom music overriding default music and music mods
    PHYSFS_mkdir "music/albums"    ' where custom music albums will go

    make_dir SaveGamePath, "savegames", vbNullString  ' save games
    PHYSFS_mkdir SAVEGAME_CAM      ' campaign save games
    PHYSFS_mkdir SAVEGAME_CAM_AUTO ' campaign autosave games
    PHYSFS_mkdir SAVEGAME_SKI      ' skirmish save games
    PHYSFS_mkdir SAVEGAME_SKI_AUTO ' skirmish autosave games

    make_dir ReplayPath, "replay", vbNullString  ' replays
    PHYSFS_mkdir "replay/skirmish"
    PHYSFS_mkdir "replay/multiplay"

    make_dir ScreenDumpPath, "screenshots", vbNullString   ' for screenshots

    PHYSFS_mkdir "tests"           ' test games launched with --skirmish=game

    PHYSFS_mkdir "userdata"        ' per-mod data user generated data
    PHYSFS_mkdir "userdata/campaign"   ' contains campaign templates
    PHYSFS_mkdir "userdata/mp"     ' contains multiplayer templates
    rulesettag = ""                ' stores tag property of ruleset.json files

    If Not customDebugfile Then
        ' there was no custom debug file specified  (--debug-file=blah)
        ' so we use our write directory to store our logs.
        debug_filename = PHYSFS_getWriteDir() & getDefaultLogFilePath(PHYSFS_getDirSeparator())
        debug_register_callback CB_debug_callback_file, CB_debug_callback_file_init, CB_debug_callback_file_exit, debug_filename ' note: by the time this function returns, all use of debug_filename has completed

        debug_ 1886, LOG_WZ, "realmain", "Using %s debug file", debug_filename
    End If

    ' Initialize random number generators
    srand FromU32(time_())
    sound_srand FromU32(time_())    ' (VB) C's rand() has one state per process: WzSound's (playlist.cpp's) is seeded too
    gameSRand rand()

    ' NOTE: it is now safe to use debug() calls to make sure output gets captured.
    debug_ 1895, LOG_WZ, "realmain", "Warzone 2100 - %s", version_getFormattedVersionString(False)
    debug_ 1896, LOG_WZ, "realmain", "Using language: %s", getLanguage()
    debug_ 1897, LOG_WZ, "realmain", "Backend: %s", BACKEND
    urlRequestOutputDebugInfo

    ' Put in the writedir root
    gInputManager_setKeyMapJsonPath "keymap.json"
    FavoriteStructuresPath = "favoriteStructures.json"

    ' initialise all the command line states
    war_SetDefaultStates

    debug_ 1916, LOG_MAIN, "realmain", "initializing"

    loadConfig

    ' parse the command line
    If Not ParseCommandLine(utfargc, utfargv) Then
        realmain = EXIT_FAILURE
        Exit Function
    End If

    ' Save new (commandline) settings
    saveConfig

    ' Print out some initial information if in headless mode
    If headlessGameMode() Then
        fprintf STDOUT_FILE, "--------------------------------------------------------------------------------------\n"
        fprintf STDOUT_FILE, " * Warzone 2100 - Headless Mode\n"
        fprintf STDOUT_FILE, " * %s\n", version_getFormattedVersionString(False)
        If to_swap_mode(war_GetVsync()) = swap_interval_mode_immediate Then
            fprintf STDOUT_FILE, " * NOTE: VSYNC IS DISABLED - CPU USAGE MAY BE UNBOUNDED\n"
        End If
        fprintf STDOUT_FILE, "--------------------------------------------------------------------------------------\n"
        fflush STDOUT_FILE
    End If

    ' Find out where to find the data
    scanDataDirs

    ' Now we check the mods to see if they exist or not (specified on the command line)
    ' FIX ME: I know this is a bit hackish, but better than nothing for now?
    ' check whether given global mods are regular files
    globalModsPath = version_getVersionedModsFolderPath("global")
    i = 0
    Do While i < global_mods.size
        modtocheck = globalModsPath & "/" & global_mods.items(i)
        PHYSFS_stat modtocheck, metaData_filesize, metaData_modtime, metaData_filetype, metaData_readonly
        If metaData_filetype <> PHYSFS_FILETYPE_REGULAR Then
            debug_ 1961, LOG_ERROR, "realmain", "The global mod ""%s"" you have specified doesn't exist!", global_mods.items(i)
            StringVector_erase global_mods, i
            rebuildSearchPath mod_multiplay, True
        Else
            wz_info 1967, "realmain", "global mod ""%s"" is enabled", global_mods.items(i)
            i = i + 1
        End If
    Loop
    ' check whether given campaign mods are regular files
    campaignModsPath = version_getVersionedModsFolderPath("campaign")
    i = 0
    Do While i < campaign_mods.size
        modtocheck = campaignModsPath & "/" & campaign_mods.items(i)
        PHYSFS_stat modtocheck, metaData_filesize, metaData_modtime, metaData_filetype, metaData_readonly
        If metaData_filetype <> PHYSFS_FILETYPE_REGULAR Then
            debug_ 1979, LOG_ERROR, "realmain", "The campaign mod ""%s"" you have specified doesn't exist!", campaign_mods.items(i)
            StringVector_erase campaign_mods, i
            rebuildSearchPath mod_campaign, True
        Else
            wz_info 1985, "realmain", "campaign mod ""%s"" is enabled", campaign_mods.items(i)
            i = i + 1
        End If
    Loop
    ' check whether given multiplay mods are regular files
    multiplayModsPath = version_getVersionedModsFolderPath("multiplay")
    i = 0
    Do While i < multiplay_mods.size
        modtocheck = multiplayModsPath & "/" & multiplay_mods.items(i)
        PHYSFS_stat modtocheck, metaData_filesize, metaData_modtime, metaData_filetype, metaData_readonly
        If metaData_filetype <> PHYSFS_FILETYPE_REGULAR Then
            debug_ 1997, LOG_ERROR, "realmain", "The multiplay mod ""%s"" you have specified doesn't exist!", multiplay_mods.items(i)
            StringVector_erase multiplay_mods, i
            rebuildSearchPath mod_multiplay, True
        Else
            wz_info 2003, "realmain", "multiplay mod ""%s"" is enabled", multiplay_mods.items(i)
            i = i + 1
        End If
    Loop

    ' (VB) --vbtest=<file>: the test commands of tools\run_tests.py, before the window opens
    If Len(vbTestCommandFile) > 0 Then
        realmain = vbtest_run(vbTestCommandFile)
        PlatformSetGame Nothing
        frameSetGame Nothing
        NetSetGame Nothing
        SoundSetGame Nothing
        Exit Function
    End If

    ' (optional<video_backend> gfxbackend: none in headless mode)
    If Not wzMainScreenSetup(Not headlessGameMode(), war_getGfxBackend(), war_getAntialiasing(), war_getWindowMode(), war_GetVsync(), war_getLODDistanceBiasPercentage(), war_getShadowMapResolution()) Then
        saveConfig ' ensure any setting changes are persisted on failure
        realmain = EXIT_FAILURE
        Exit Function
    End If

    debug_ 2024, LOG_WZ, "realmain", "Warzone 2100 - %s", version_getFormattedVersionString(False)
    debug_ 2025, LOG_WZ, "realmain", "Using language: %s", getLanguage()
    debug_ 2026, LOG_WZ, "realmain", "Backend: %s", BACKEND

    wzGetGameToRendererScaleFactorInt horizScaleFactor, vertScaleFactor
    debug_ 2039, LOG_WZ, "realmain", "Game to Renderer Scale Factor (w x h): %u%% x %u%%", horizScaleFactor, vertScaleFactor

    debug_ 2041, LOG_MAIN, "realmain", "Final initialization"
    If Not frameInitialise() Then
        realmain = EXIT_FAILURE
        Exit Function
    End If
    If Not pie_LoadShaders(war_getShadowFilterSize(), war_getPointLightPerPixelLighting() And getTerrainShaderQuality() = NORMAL_MAPPING) Then
        realmain = EXIT_FAILURE
        Exit Function
    End If
    If Not screenInitialise() Then
        realmain = EXIT_FAILURE
        Exit Function
    End If

    fogConfigOption = pie_GetFogEnabled()
    pie_SetFogStatus False

    pal_Init

    pie_LoadBackDrop SCREEN_RANDOMBDROP
    pie_SetFogStatus False

    pie_ScreenFrameRenderEnd

    If Not systemInitialise(horizScaleFactor, vertScaleFactor) Then
        realmain = EXIT_FAILURE
        Exit Function
    End If

    ' The systemInitialise() above sets fog to disabled as default so this must be after it
    If fogConfigOption Then
        pie_EnableFog True
    End If

    pie_InitLighting

    'set all the pause states to false
    setAllPauseStates False

    ' Copy this info to be used by the crash handler for the dump file
    addDumpInfo va("Using Backend: %s", BACKEND)
    addDumpInfo va("Using language: %s", getLanguageName())

    ' Do the game mode specific initialisation.
    Select Case GetGameMode()
    Case GS_TITLE_SCREEN
        ' The usual case (unless command-line flags specify otherwise): Load into the title menu
        startTitleLoop True
    Case GS_SAVEGAMELOAD
        If headlessGameMode() Then
            fprintf STDOUT_FILE, "Loading savegame ...\n"
        End If
        initSaveGameLoad
    Case GS_NORMAL
        If Not startGameLoop() Then
            ' Attempted to load straight into a game from the command-line, but starting the game loop (loading the map, etc) failed
            ' Treat this as a failure and queue an exit
            wzQuit EXIT_FAILURE
        End If
    Case Else
        debug_ 2110, LOG_ERROR, "realmain", "Weirdy game status, I'm afraid!!"
    End Select

    ' (asyncGetCompatCheckResults(mainProcessCompatCheckResults), WzInfoManager::initialize(): not ported)

    debug_ 2123, LOG_MAIN, "realmain", "Entering main loop"
    wzMainEventLoop True                   ' mainShutdown afterwards
    exitCode = wzGetQuitExitCode()
    wzShutdown
    PlatformSetGame Nothing                ' (VB)
    frameSetGame Nothing                   ' (VB)
    NetSetGame Nothing                     ' (VB)
    SoundSetGame Nothing                   ' (VB)
    debug_ 2138, LOG_MAIN, "realmain", "Completed shutting down Warzone 2100"
    realmain = exitCode
End Function

' Get the mode the game is currently in
Public Function GetGameMode() As Long
    GetGameMode = gameStatus
End Function

' Set the current mode
Public Sub SetGameMode(ByVal status As Long)
    gameStatus = status
End Sub

