Attribute VB_Name = "loop_c"
' src/loop.cpp, loop.h: so far the game state's update (gameStateUpdate,
' made public for the tests and the headless loop: the C++ is static),
' gameLoop (its render part is renderLoop, a stand-in until the display
' clusters), countUpdate, the droid counts (getNumDroids, adjustDroidCount,
' ...), the pause states and the fast forward's settings. The videos and
' the mission loop states come with the display clusters. The video playback (videoLoop, the
' video mode: cluster I) is at the end.
' executeFnAndProcessScriptQueuedRemovals([]{ f(); }) is the call to f
' followed by wzapi_processScriptQueuedObjectRemovals (its assert that no
' removals are left over cannot fire: no script queues one yet).
' The Rationals renderFraction (2/5) and updateFraction (3/5) are their
' numerators and the floors of their products with 500.
Option Explicit

Public Enum GAMECODE
    GAMECODE_CONTINUE
    GAMECODE_RESTARTGAME
    GAMECODE_QUITGAME
    GAMECODE_PLAYVIDEO
    GAMECODE_NEWLEVEL
    GAMECODE_FASTEXIT
    GAMECODE_LOADGAME
End Enum

Public Const WZ_DEFAULT_MAX_FASTFORWARD_TICKS As Long = 1

'holds which pause is valid at any one time
Private Type PAUSE_STATE
    gameUpdatePause As Boolean
    audioPause As Boolean
    scriptPause As Boolean
    scrollPause As Boolean
    consolePause As Boolean
End Type

Private paused As Boolean
Private video As Boolean                    ' static bool video = false;

' Time data for skipping videos.
Private Type VIDEO_TIME_SKIP_STATE
    begin As Double                         ' std::chrono::steady_clock::time_point (widgSteadyNowMs)
    canSkip As Boolean
End Type
Private pauseState As PAUSE_STATE
Private videoTimeSkipState As VIDEO_TIME_SKIP_STATE
Private Const sequenceMinSkipTime As Long = 800 ' std::chrono::milliseconds sequenceMinSkipTime = std::chrono::milliseconds(800);

Private videoMode As Long                   ' static SDWORD videoMode = 0;

Private numDroids(0 To MAX_PLAYERS - 1) As Long             ' unsigned
Private numMissionDroids(0 To MAX_PLAYERS - 1) As Long
Private numTransporterDroids(0 To MAX_PLAYERS - 1) As Long
Private numCommandDroids(0 To MAX_PLAYERS - 1) As Long
Private numConstructorDroids(0 To MAX_PLAYERS - 1) As Long

Private maxFastForwardTicks As Long         ' size_t (loop_c_init: WZ_DEFAULT_MAX_FASTFORWARD_TICKS)
Private fastForwardTicksFixedToNormalTickRate As Boolean ' can be set to false to "catch-up" as quickly as possible (but this may result in more jerky behavior)

Private renderBudget As Long                ' Scaled time spent rendering minus scaled time spent updating.
Private Const renderFraction_n As Long = 2  ' renderFraction(2, 5): Minimum fraction of time spent rendering.
Private Const updateFraction_n As Long = 3  ' updateFraction = Rational(1) - renderFraction
Private Const renderFraction_500_floor As Long = 200    ' (renderFraction * 500).floor()
Private Const updateFraction_500_neg_floor As Long = -300   ' (-updateFraction * 500).floor()

' gameLoop's statics
Private lastFlushTime As Long               ' uint32_t
Private numForcedUpdatesLastCall As Long    ' size_t
Private previousUpdateWasRender As Boolean

' loop.cpp: the pie and polygon counts of the frame (pie_GetResetCounts)
Public loopPieCount As Long
Public loopPolyCount As Long

' the C++ initialisers
Public Sub loop_c_init()
    nextMissionType = LEVEL_TYPE_LDS_NONE
    prevMissionType = LEVEL_TYPE_LDS_NONE
    maxFastForwardTicks = WZ_DEFAULT_MAX_FASTFORWARD_TICKS
    fastForwardTicksFixedToNormalTickRate = True
End Sub

Private Function renderLoop() As GAMECODE
    Dim skipDrawing As Boolean, intRetVal As Long, quitting As Boolean, i As Long, it As Long, widgval As Long
    Dim triggers() As WidgetTrigger, nTriggers As Long
    Dim msgbuffer As String
    If bMultiPlayer And Not NetPlay_isHostAlive And NetPlay_bComms And Not NetPlay_isHost Then
        intAddInGamePopup
    End If

    skipDrawing = Not gfx_api_context_shouldDraw()

    audio_Update

    wzShowMouse True

    intRetVal = INT_NONE
    If Not paused Then
        ' Always refresh the widgets' backing stores if needed, even if we don't process clicks below
        intDoScreenRefresh

        ' Run the in game interface and see if it grabbed any mouse clicks
        If Not getRotActive() And getWidgetsStatus() And dragBox3D.status <> DRAG_DRAGGING And wallDrag.status <> DRAG_DRAGGING Then
            intRetVal = intRunWidgets()
            screen_FlipIfBackDropTransition

            ' Send droid orders, if any. (Should do between intRunWidgets() calls, to avoid droid orders getting mixed up, in the case of multiple orders given while the game freezes due to net lag.)
            sendQueuedDroidInfo
        End If

        'don't process the object lists if paused or about to quit to the front end
        If Not gameUpdatePaused() And intRetVal <> INT_QUIT Then
            If dragBox3D.status <> DRAG_DRAGGING And wallDrag.status <> DRAG_DRAGGING _
               And (intRetVal = INT_INTERCEPT Or isMouseOverRadar()) Then
                ' Using software cursors (when on) for these menus due to a bug in SDL's SDL_ShowCursor()
                wzSetCursor CURSOR_DEFAULT
            End If

            ' (#ifdef DEBUG: checkFactoryFlags, the duplicate delivery points: release build)

            'handles callbacks for positioning of DP's
            process3DBuilding
            processDeliveryRepos

            'ajl. get the incoming netgame messages and process them.
            ' FIXME Previous comment is deprecated. multiPlayerLoop does some other weird stuff, but not that anymore.
            If bMultiPlayer Then
                multiPlayerLoop
            End If

            For i = 0 To MAX_PLAYERS - 1
                it = apsDroidLists(i).head
                Do While it <> 0
                    ' Don't copy the next pointer - if droids somehow get destroyed in the graphics rendering loop, who cares if we crash.
                    calcDroidIllumination lnObj(it)
                    it = lnNext(it)
                Loop
            Next
        End If

        If Not consolePaused() Then
            ' Process all the console messages
            updateConsoleMessages
        End If
        If Not scrollPaused() And dragBox3D.status <> DRAG_DRAGGING And intMode <> INT_INGAMEOP Then
            displayRenderLoop
        End If
    Else  ' paused
        ' Using software cursors (when on) for these menus due to a bug in SDL's SDL_ShowCursor()
        wzSetCursor CURSOR_DEFAULT

        If Not scrollPaused() And dragBox3D.status <> DRAG_DRAGGING Then
            displayRenderLoop
        End If

        If Not bLoadSaveUp Then
            nTriggers = widgRunScreen(psWScreen, triggers)      ' always run the screen, so overlays can process input
            ' (VB) unsigned widgval = triggers.empty() ? 0 : triggers.front().widget->id;
            If nTriggers = 0 Then widgval = 0 Else widgval = triggers(0).widget.id

            If InGameOpUp Or isInGamePopupUp Or intHelpOverlayIsUp() Then   ' ingame options menu up, run it!
                ' Just use first click here, since the next click could be on another menu.
                intProcessInGameOptions widgval
                If widgval = INTINGAMEOP_QUIT Or widgval = INTINGAMEOP_POPUP_QUIT Then
                    If gamePaused() Then
                        kf_TogglePauseMode
                    End If
                    intRetVal = INT_QUIT
                End If
            End If
        ElseIf runLoadSave(True) And Len(sRequestResult) > 0 Then
            debug_ 262, LOG_NEVER, "renderLoop", "Returned %s", sRequestResult
            If bRequestLoad Then
                loopMissionState = LMS_LOADGAME
                saveGameName = sRequestResult
            Else
                msgbuffer = ""

                If saveInMissionRes() Then
                    If saveGame(sRequestResult, GTYPE_SAVE_START) Then
                        msgbuffer = "GAME SAVED: " & savegameWithoutExtension(sRequestResult)
                        addConsoleMessage msgbuffer, LEFT_JUSTIFY, NOTIFY_MESSAGE
                    Else
                        ASSERT_FAILURE "false", "src/loop.cpp:282", "renderLoop", "Mission Results: saveGame Failed"
                        msgbuffer = "Could not save game!"
                        addConsoleMessage msgbuffer, LEFT_JUSTIFY, NOTIFY_MESSAGE
                        deleteSaveGame_classic sRequestResult
                    End If
                ElseIf bMultiPlayer Or saveMidMission() Then
                    If saveGame(sRequestResult, GTYPE_SAVE_MIDMISSION) Then 'mid mission from [esc] menu
                        msgbuffer = "GAME SAVED: " & savegameWithoutExtension(sRequestResult)
                        addConsoleMessage msgbuffer, LEFT_JUSTIFY, NOTIFY_MESSAGE
                    Else
                        ASSERT_FAILURE "!""saveGame(sRequestResult, GTYPE_SAVE_MIDMISSION) failed""", "src/loop.cpp:298", "renderLoop", "Mid Mission: saveGame Failed"
                        msgbuffer = "Could not save game!"
                        addConsoleMessage msgbuffer, LEFT_JUSTIFY, NOTIFY_MESSAGE
                        deleteSaveGame_classic sRequestResult
                    End If
                Else
                    ASSERT_FAILURE "false", "src/loop.cpp:306", "renderLoop", "Attempt to save game with incorrect load/save mode"
                End If
            End If
        End If
    End If

    ' Check for quit
    quitting = False
    If intRetVal = INT_QUIT Then
        If Not loop_GetVideoStatus() Then
            'quitting from the game to the front end
            'so get a new backdrop
            quitting = True

            pie_LoadBackDrop SCREEN_RANDOMBDROP
        End If
    End If
    If Not loop_GetVideoStatus() And Not quitting And Not headlessGameMode() And Not skipDrawing Then
        If Not gameUpdatePaused() Then
            processInput

            'no key clicks or in Intelligence Screen
            If Not isMouseOverRadar() And Not isDraggingInGameNotification() And Not isMouseClickDownOnScreenOverlayChild() _
               And intRetVal = INT_NONE And Not InGameOpUp And Not isInGamePopupUp Then
                processMouseClickInput
            End If
            displayWorld
        End If
        wzPerfBegin PERF_GUI, "User interface"
        ' Display the in game interface
        pie_SetFogStatus False

        If getWidgetsStatus() Then
            intDisplayWidgets
        End If
        pie_SetFogStatus True
        wzPerfEnd PERF_GUI
    End If

    pie_GetResetCounts loopPieCount, loopPolyCount

    ' deal with the mission state
    Select Case loopMissionState
    Case LMS_CLEAROBJECTS
        missionDestroyObjects
        setScriptPause True
        loopMissionState = LMS_SETUPMISSION

    Case LMS_NORMAL
        ' default
    Case LMS_SETUPMISSION
        setScriptPause False
        If Not setUpMission(nextMissionType) Then
            renderLoop = GAMECODE_QUITGAME
            Exit Function
        End If
    Case LMS_SAVECONTINUE
        ' just wait for this to be changed when the new mission starts
    Case LMS_NEWLEVEL
        nextMissionType = LEVEL_TYPE_LDS_NONE
        renderLoop = GAMECODE_NEWLEVEL
        Exit Function
    Case LMS_LOADGAME
        renderLoop = GAMECODE_LOADGAME
        Exit Function
    Case Else
        ASSERT_FAILURE "false", "src/loop.cpp:383", "renderLoop", "unknown loopMissionState"
    End Select

    If quitting Then
        renderLoop = GAMECODE_QUITGAME
        Exit Function
    ElseIf loop_GetVideoStatus() Then
        audio_StopAll
        renderLoop = GAMECODE_PLAYVIDEO
        Exit Function
    End If

    renderLoop = GAMECODE_CONTINUE
End Function

Public Sub countUpdate(ByVal synch As Boolean)
    Dim i As Long, it As Long, psCurr As Long, psCBuilding As Long
    For i = 0 To MAX_PLAYERS - 1
        'set the flag for each player
        setSatUplinkExists False, i

        numCommandDroids(i) = 0
        numConstructorDroids(i) = 0
        numDroids(i) = 0
        numMissionDroids(i) = 0
        numTransporterDroids(i) = 0

        it = apsDroidLists(i).head
        Do While it <> 0
            psCurr = lnObj(it)
            numDroids(i) = numDroids(i) + 1
            Select Case aDroid(psCurr).droidType
            Case DROID_COMMAND
                numCommandDroids(i) = numCommandDroids(i) + 1
            Case DROID_CONSTRUCT, DROID_CYBORG_CONSTRUCT
                numConstructorDroids(i) = numConstructorDroids(i) + 1
            Case DROID_TRANSPORTER, DROID_SUPERTRANSPORTER
                droidCountsInTransporter psCurr, i
            End Select
            it = lnNext(it)
        Loop
        it = mission_apsDroidLists(i).head
        Do While it <> 0
            psCurr = lnObj(it)
            numMissionDroids(i) = numMissionDroids(i) + 1
            Select Case aDroid(psCurr).droidType
            Case DROID_COMMAND
                numCommandDroids(i) = numCommandDroids(i) + 1
            Case DROID_CONSTRUCT, DROID_CYBORG_CONSTRUCT
                numConstructorDroids(i) = numConstructorDroids(i) + 1
            Case DROID_TRANSPORTER, DROID_SUPERTRANSPORTER
                droidCountsInTransporter psCurr, i
            End Select
            it = lnNext(it)
        Loop
        it = apsLimboDroids(i).head
        Do While it <> 0
            psCurr = lnObj(it)
            ' count the type of units
            Select Case aDroid(psCurr).droidType
            Case DROID_COMMAND
                numCommandDroids(i) = numCommandDroids(i) + 1
            Case DROID_CONSTRUCT, DROID_CYBORG_CONSTRUCT
                numConstructorDroids(i) = numConstructorDroids(i) + 1
            End Select
            it = lnNext(it)
        Loop
        ' FIXME: These for-loops are code duplicationo
        setLasSatExists False, i
        it = apsStructLists(i).head
        Do While it <> 0
            psCBuilding = lnObj(it)
            If psCBuilding = 0 Then GoTo next_struct
            If isDead(psCBuilding) Then GoTo next_struct
            If aStruct(psCBuilding).pStructureType <> -1 Then
                If asStructureStats(aStruct(psCBuilding).pStructureType).typ = REF_SAT_UPLINK And aStruct(psCBuilding).status = SS_BUILT Then
                    setSatUplinkExists True, i
                End If
            End If
            'don't wait for the Las Sat to be built - can't build another if one is partially built
            If asWeaponStats(BASE_OBJECT_getWeaponStats(psCBuilding, 0)).weaponSubClass = WSC_LAS_SAT Then
                setLasSatExists True, i
            End If
next_struct:
            it = lnNext(it)
        Loop
        it = mission_apsStructLists(i).head
        Do While it <> 0
            psCBuilding = lnObj(it)
            If psCBuilding = 0 Then GoTo next_mission_struct
            If isDead(psCBuilding) Then GoTo next_mission_struct
            If aStruct(psCBuilding).pStructureType <> -1 Then
                If asStructureStats(aStruct(psCBuilding).pStructureType).typ = REF_SAT_UPLINK And aStruct(psCBuilding).status = SS_BUILT Then
                    setSatUplinkExists True, i
                End If
            End If
            'don't wait for the Las Sat to be built - can't build another if one is partially built
            If asWeaponStats(BASE_OBJECT_getWeaponStats(psCBuilding, 0)).weaponSubClass = WSC_LAS_SAT Then
                setLasSatExists True, i
            End If
next_mission_struct:
            it = lnNext(it)
        Loop
        If synch Then
            syncDebug "countUpdate", "counts[%d] = {droid: %d, command: %d, constructor: %d, mission: %d, transporter: %d}", i, numDroids(i), numCommandDroids(i), numConstructorDroids(i), numMissionDroids(i), numTransporterDroids(i)
        End If
    Next
End Sub

' (static in C++)
Public Sub gameStateUpdate()
    Dim overrideHandleClientBlindNames As Boolean, n As Long, crc As Long, i As Long, it As Long, nxt As Long
    Dim alloc(0 To 9) As Long, position(0 To 9) As Long, playerName As String
    For n = 0 To 9
        alloc(n) = -CLng(NetPlay_players_allocated(n))
        position(n) = NetPlay_players_position(n)
    Next
    syncDebug "gameStateUpdate", "map = ""%s"", pseudorandom 32-bit integer = 0x%08X, allocated = %d %d %d %d %d %d %d %d %d %d, position = %d %d %d %d %d %d %d %d %d %d", game.map, gameRandU32(), _
              alloc(0), alloc(1), alloc(2), alloc(3), alloc(4), alloc(5), alloc(6), alloc(7), alloc(8), alloc(9), _
              position(0), position(1), position(2), position(3), position(4), position(5), position(6), position(7), position(8), position(9)
    overrideHandleClientBlindNames = game.blindMode >= BLIND_MODE_BLIND_GAME
    If overrideHandleClientBlindNames Then overrideHandleClientBlindNames = NetPlay_isHost Or NETisReplay()
    If overrideHandleClientBlindNames Then overrideHandleClientBlindNames = Not ingame.has_endTime
    For n = 0 To MAX_PLAYERS - 1
        If Not overrideHandleClientBlindNames Then playerName = NetPlay_players_name(n) Else playerName = getPlayerGenericName(n)
        syncDebug "gameStateUpdate", "Player %d = ""%s""", n, playerName
    Next

    ' Add version string to desynch logs. Different version strings will not trigger a desynch dump per se, due to the syncDebug{Get, Set}Crc guard.
    crc = syncDebugGetCrc()
    syncDebug "gameStateUpdate", "My client version = %s", version_getVersionString()
    syncDebugSetCrc crc

    ' Actually send pending droid orders.
    sendQueuedDroidInfo

    sendPlayerGameTime
    NETflush  ' Make sure the game time tick message is really sent over the network.

    If Not paused And Not scriptPaused() Then
        updateScripts
        wzapi_processScriptQueuedObjectRemovals
    End If

    ' Update abandoned structures
    handleAbandonedStructures

    ' Update the visibility change stuff
    visUpdateLevel

    ' Put all droids/structures/features into the grid.
    gridReset

    ' Check which objects are visible.
    processVisibility

    ' Update the map.
    mapUpdate

    'update the findpath system
    fpathUpdate

    ' update the command droids
    cmdDroidUpdate

    For i = 0 To MAX_PLAYERS - 1
        'update the current power available for a player
        updatePlayerPower i

        ' mutating_list_iterate
        it = apsDroidLists(i).head
        Do While it <> 0
            nxt = lnNext(it)
            droidUpdate lnObj(it)
            it = nxt
        Loop
        wzapi_processScriptQueuedObjectRemovals
        it = mission_apsDroidLists(i).head
        Do While it <> 0
            nxt = lnNext(it)
            missionDroidUpdate lnObj(it)
            it = nxt
        Loop
        wzapi_processScriptQueuedObjectRemovals
        ' FIXME: These for-loops are code duplication
        it = apsStructLists(i).head
        Do While it <> 0
            nxt = lnNext(it)
            structureUpdate lnObj(it), False
            it = nxt
        Loop
        wzapi_processScriptQueuedObjectRemovals
        it = mission_apsStructLists(i).head
        Do While it <> 0
            nxt = lnNext(it)
            structureUpdate lnObj(it), True  ' update for mission
            it = nxt
        Loop
        wzapi_processScriptQueuedObjectRemovals
    Next

    missionTimerUpdate

    proj_UpdateAll
    wzapi_processScriptQueuedObjectRemovals

    it = apsFeatureList(0).head
    Do While it <> 0
        featureUpdate lnObj(it)
        it = lnNext(it)
    Loop

    ' Free dead droid memory.
    objmemUpdate

    ' accumulate occasional stats / snapshots
    If Not paused And Not scriptPaused() Then
        GameStoryLogger_logGameFrame
    End If

    ' Must end update, since we may or may not have ticked, and some message queue processing code may vary depending on whether it's in an update.
    gameTimeUpdateEnd

    ' Must be at the end of gameStateUpdate, since countUpdate is also called randomly (unsynchronised) between gameStateUpdate calls, but should have no effect if we already called it, and recvMessage requires consistent counts on all clients.
    countUpdate True
End Sub

Public Function getMaxFastForwardTicks() As Long
    getMaxFastForwardTicks = maxFastForwardTicks
End Function

' optional<size_t> value: hasValue False is nullopt
Public Sub setMaxFastForwardTicks(ByVal hasValue As Boolean, ByVal value As Long, ByVal fixedToNormalTickRate As Boolean)
    If hasValue Then maxFastForwardTicks = value Else maxFastForwardTicks = WZ_DEFAULT_MAX_FASTFORWARD_TICKS
    fastForwardTicksFixedToNormalTickRate = fixedToNormalTickRate
End Sub

' The main game loop
Public Function gameLoop() As GAMECODE
    Dim numRegularUpdatesTicks As Long, numFastForwardTicks As Long, timeUpdateResult As GameTimeUpdateResult
    Dim selectedPlayerIsSpectator As Boolean, multiplayerHostDisconnected As Boolean, canFastForwardGameTime As Boolean
    Dim forceTryGameTickUpdate As Boolean, before As Long, after As Long, renderReturn As GAMECODE

    numRegularUpdatesTicks = 0
    numFastForwardTicks = 0
    gameTimeUpdateBegin
    Do
        ' Receive NET_BLAH messages.
        ' Receive GAME_BLAH messages, and if it's time, process exactly as many GAME_BLAH messages as required to be able to tick the gameTime.
        recvMessage
        wzapi_processScriptQueuedObjectRemovals

        ' (VB) the && chains short-circuited as in C
        selectedPlayerIsSpectator = False
        If bMultiPlayer Then selectedPlayerIsSpectator = NetPlay_players_isSpectator(selectedPlayer)
        multiplayerHostDisconnected = bMultiPlayer And Not NetPlay_isHostAlive And NetPlay_bComms And Not NetPlay_isHost ' do not fast-forward after the host has disconnected
        canFastForwardGameTime = False
        If selectedPlayerIsSpectator Then           ' current player must be a spectator
            If Not NetPlay_isHost Then              ' AND NOT THE HOST (!)
                If Not multiplayerHostDisconnected Then ' and the multiplayer host must not be disconnected ("host quit")
                    If numFastForwardTicks < maxFastForwardTicks Then ' and the number of forced updates this call of gameLoop must not exceed the max allowed
                        canFastForwardGameTime = checkPlayerGameTime(NET_ALL_PLAYERS) ' and there must be a new game tick available to process from all players
                    End If
                End If
            End If
        End If

        forceTryGameTickUpdate = False
        If canFastForwardGameTime Then
            If (Not fastForwardTicksFixedToNormalTickRate And numForcedUpdatesLastCall > 0) Or numRegularUpdatesTicks > 0 Then
                forceTryGameTickUpdate = NETgameIsBehindPlayersByAtLeast(4)
            End If
        End If

        ' Update gameTime and graphicsTime, and corresponding deltas. Note that gameTime and graphicsTime pause, if we aren't getting our GAME_GAME_TIME messages.
        timeUpdateResult = gameTimeUpdate(renderBudget > 0 Or previousUpdateWasRender, forceTryGameTickUpdate)

        If timeUpdateResult = GameTimeUpdateResult_GAME_TIME_UPDATED_FORCED Then
            numFastForwardTicks = numFastForwardTicks + 1
            ' FUTURE TODO: Could trigger the display of a UI icon here (if (numFastForwardTicks == maxFastForwardTicks)?) to indicate that the game is fast-forwarding substantially
        ElseIf timeUpdateResult = GameTimeUpdateResult_GAME_TIME_UPDATED Then
            numRegularUpdatesTicks = numRegularUpdatesTicks + 1
        End If

        If deltaGameTime = 0 Then
            Exit Do  ' Not doing a game state update.
        End If

        If Not (Not paused And Not gameUpdatePaused()) Then ASSERT_FAILURE "!paused && !gameUpdatePaused()", "src/loop.cpp:698", "gameLoop", "Nonsensical pause values."

        before = wzGetTicks()
        syncDebug "gameLoop", "Begin game state update, gameTime = %d", gameTime
        gameStateUpdate
        syncDebug "gameLoop", "End game state update, gameTime = %d", gameTime
        after = wzGetTicks()

        renderBudget = renderBudget - (after - before) * renderFraction_n
        If renderBudget < updateFraction_500_neg_floor Then renderBudget = updateFraction_500_neg_floor
        previousUpdateWasRender = False

        If Not (deltaGraphicsTime = 0) Then ASSERT_FAILURE "deltaGraphicsTime == 0", "src/loop.cpp:710", "gameLoop", "Shouldn't update graphics and game state at once."
    Loop
    numForcedUpdatesLastCall = numFastForwardTicks

    If Not ULT(FromU32(U32(realTime) - U32(lastFlushTime)), 400) Then
        lastFlushTime = realTime
        NETflush  ' Make sure that we aren't waiting too long to send data.
    End If

    before = wzGetTicks()
    renderReturn = renderLoop()
    watchtest_frame                     ' (VB) --vbwatchshot
    pie_ScreenFrameRenderEnd ' must happen here for proper renderBudget calculation
    after = wzGetTicks()
    wzapi_processScriptQueuedObjectRemovals

    renderBudget = renderBudget + (after - before) * updateFraction_n
    If renderBudget > renderFraction_500_floor Then renderBudget = renderFraction_500_floor
    previousUpdateWasRender = True

    If headlessGameMode() And autogame_enabled() Then
        ' Output occasional stats to stdout
        stdOutGameSummary
    End If

    gameLoop = renderReturn
End Function

Public Function gamePaused() As Boolean
    gamePaused = paused
End Function

Public Sub setGamePauseStatus(ByVal val As Boolean)
    paused = val
End Sub

Public Function gameUpdatePaused() As Boolean
    gameUpdatePaused = pauseState.gameUpdatePause
End Function

Public Function audioPaused() As Boolean
    audioPaused = pauseState.audioPause
End Function

Public Function scriptPaused() As Boolean
    scriptPaused = pauseState.scriptPause
End Function

Public Function scrollPaused() As Boolean
    scrollPaused = pauseState.scrollPause
End Function

Public Function consolePaused() As Boolean
    consolePaused = pauseState.consolePause
End Function

Public Sub setGameUpdatePause(ByVal state As Boolean)
    pauseState.gameUpdatePause = state
End Sub

Public Sub setAudioPause(ByVal state As Boolean)
    pauseState.audioPause = state
End Sub

Public Sub setScriptPause(ByVal state As Boolean)
    pauseState.scriptPause = state
End Sub

Public Sub setScrollPause(ByVal state As Boolean)
    pauseState.scrollPause = state
End Sub

Public Sub setConsolePause(ByVal state As Boolean)
    pauseState.consolePause = state
End Sub

'set all the pause states to the state value
Public Sub setAllPauseStates(ByVal state As Boolean)
    setGameUpdatePause state
    setAudioPause state
    setScriptPause state
    setScrollPause state
    setConsolePause state
End Sub

Public Function getNumDroids(ByVal player As Long) As Long
    getNumDroids = numDroids(player)
End Function

Public Function getNumTransporterDroids(ByVal player As Long) As Long
    getNumTransporterDroids = numTransporterDroids(player)
End Function

Public Function getNumMissionDroids(ByVal player As Long) As Long
    getNumMissionDroids = numMissionDroids(player)
End Function

Public Function getNumCommandDroids(ByVal player As Long) As Long
    getNumCommandDroids = numCommandDroids(player)
End Function

Public Function getNumConstructorDroids(ByVal player As Long) As Long
    getNumConstructorDroids = numConstructorDroids(player)
End Function

' increase the droid counts - used by update factory to keep the counts in sync
Public Sub adjustDroidCount(ByVal droid_ As Long, ByVal delta As Long)
    Dim player As Long
    player = aDroid(droid_).player
    syncDebug "adjustDroidCount", "numDroids[%d]:%d=%d" & ChrW$(&H2192) & "%d", player, aDroid(droid_).droidType, numDroids(player), numDroids(player) + delta
    numDroids(player) = numDroids(player) + delta
    Select Case aDroid(droid_).droidType
    Case DROID_COMMAND
        numCommandDroids(player) = numCommandDroids(player) + delta
    Case DROID_CONSTRUCT, DROID_CYBORG_CONSTRUCT
        numConstructorDroids(player) = numConstructorDroids(player) + delta
    End Select
End Sub

' Increase counts of droids in a transporter
Public Sub droidCountsInTransporter(ByVal droid_ As Long, ByVal player As Long)
    Dim it As Long, psDroid As Long
    If Not DROID_isTransporter(droid_) Or aDroid(droid_).psGroup = 0 Then
        Exit Sub
    End If

    numTransporterDroids(player) = numTransporterDroids(player) + aGroup(aDroid(droid_).psGroup).refCount - 1

    ' and count the units inside it...
    it = aGroup(aDroid(droid_).psGroup).psList.head
    Do While it <> 0
        psDroid = lnObj(it)
        If psDroid = droid_ Then
            Exit Do
        End If
        If aDroid(psDroid).droidType = DROID_CYBORG_CONSTRUCT Or aDroid(psDroid).droidType = DROID_CONSTRUCT Then
            numConstructorDroids(player) = numConstructorDroids(player) + 1
        End If
        If aDroid(psDroid).droidType = DROID_COMMAND Then
            numCommandDroids(player) = numCommandDroids(player) + 1
        End If
        it = lnNext(it)
    Loop
End Sub

Private Sub resetVideoSkipStates()
    videoTimeSkipState.begin = widgSteadyNowMs()    ' std::chrono::steady_clock::now()
    videoTimeSkipState.canSkip = False
End Sub

' The video playback loop
Public Sub videoLoop()
    Dim videoFinished As Boolean

    If Not (videoMode = 1) Then ASSERT_FAILURE "videoMode == 1", "src/loop.cpp:753", "videoLoop", "videoMode out of sync"

    ' display a frame of the FMV
    videoFinished = Not seq_UpdateFullScreenVideo()

    If Not videoTimeSkipState.canSkip Then
        If (widgSteadyNowMs() - videoTimeSkipState.begin) > sequenceMinSkipTime Then
            videoTimeSkipState.canSkip = True
        End If
    End If

    ' should we stop playing?
    If videoFinished Or (videoTimeSkipState.canSkip And (keyPressed(KEY_ESC) Or mouseReleased(MOUSE_LMB))) Then
        seq_StopFullScreenVideo

        'set the next video off - if any
        If seq_AnySeqLeft() Then
            resetVideoSkipStates
            seq_StartNextFullScreenVideo
        Else
            ' remove the intelligence screen if necessary
            If messageIsImmediate() Then
                intResetScreen True
                setMessageImmediate False
            End If
            If Not bMultiPlayer And getScriptWinLoseVideo() <> 0 Then
                displayGameOver getScriptWinLoseVideo() = PLAY_WIN, False
            End If
            ' executeFnAndProcessScriptQueuedRemovals([]() { triggerEvent(TRIGGER_VIDEO_QUIT); });
            triggerEvent TRIGGER_VIDEO_QUIT
            wzapi_processScriptQueuedObjectRemovals
        End If
    End If
End Sub


Public Sub loop_SetVideoPlaybackMode()
    videoMode = videoMode + 1
    paused = True
    video = True
    resetVideoSkipStates
    gameTimeStop
    pie_SetFogStatus False
    audio_StopAll
    wzShowMouse False
    screen_StopBackDrop
End Sub


Public Sub loop_ClearVideoPlaybackMode()
    videoMode = videoMode - 1
    paused = False
    video = False
    resetVideoSkipStates
    gameTimeStart
    pie_SetFogStatus True
    cdAudio_Resume
    wzShowMouse True
    If Not (videoMode = 0) Then ASSERT_FAILURE "videoMode == 0", "src/loop.cpp:817", "loop_ClearVideoPlaybackMode", "loop_ClearVideoPlaybackMode: out of sync."
End Sub


Public Function loop_GetVideoMode() As Long
    loop_GetVideoMode = videoMode
End Function

Public Function loop_GetVideoStatus() As Boolean
    loop_GetVideoStatus = video
End Function
