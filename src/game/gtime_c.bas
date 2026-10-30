Attribute VB_Name = "gtime_c"
' lib/gamelib/gtime.cpp, gtime.h: the game clock, which runs only when the
' game runs, and its synchronisation through the game queues (each human
' player's GAME_GAME_TIME messages: the latency, the CRC checks).
' Rational modifier is its numerator and denominator.
Option Explicit

' See header file for documentation
Public gameTime As Long                 ' UDWORD: the current time in the game world
Public deltaGameTime As Long            ' UDWORD: the time since the last game tick
Public graphicsTime As Long             ' UDWORD: the current time in the graphical display of the game world
Public realTime As Long                 ' UDWORD: the current time in the real world - never stops, and not reset between games
Public deltaGraphicsTime As Long        ' UDWORD
Public deltaRealTime As Long            ' UDWORD
Public graphicsTimeFraction As Single   ' Private performance calculation. Do not use.
Public realTimeFraction As Single       ' Private performance calculation. Do not use.

Public extraGameTicksProcessed As Long  ' UDWORD

' The number of time units per second of the game clock.
Public Const GAME_TICKS_PER_SEC As Long = 1000
' The number of game state updates per second of the game clock.
Public Const GAME_UPDATES_PER_SEC As Long = 10
' The number of time units per game tick.
Public Const GAME_TICKS_PER_UPDATE As Long = GAME_TICKS_PER_SEC \ GAME_UPDATES_PER_SEC
' The maximum time for one frame (stops the clock running away when debugging)
' changed to /6 by ajl. if this needs to go back to ticks/10 then tell me.
Public Const GTIME_MAXFRAME As Long = GAME_TICKS_PER_SEC \ 6

' enum class GameTimeUpdateResult
Public Enum GameTimeUpdateResult
    GameTimeUpdateResult_NO_UPDATE
    GameTimeUpdateResult_GAME_TIME_UPDATED
    GameTimeUpdateResult_GAME_TIME_UPDATED_FORCED
End Enum

Private Const MAX_CONNECTED_PLAYERS As Long = MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS

' The current clock modifier. Set to speed up the game.
Private modifier_n As Long
Private modifier_d As Long

' The real time, the last time graphicsTime updated.
Private prevRealTime As Long            ' uint32_t

' Count how many times gameTimeStop has been called without a game time start.
' We use this to ensure that we can properly nest stop commands.
Private stopCount As Long               ' UDWORD

Private gameQueueTime(0 To MAX_GAMEQUEUE_SLOTS - 1) As Long          ' uint32_t
Private gameQueueCheckTime(0 To MAX_GAMEQUEUE_SLOTS - 1) As Long     ' uint32_t
Private gameQueueCheckCrc(0 To MAX_GAMEQUEUE_SLOTS - 1) As Long      ' uint32_t
Private crcError As Boolean

Private updateReadyTime As Long         ' uint32_t
Private updateWantedTime As Long        ' uint32_t
Private chosenLatency As Long           ' uint16_t
Private discreteChosenLatency As Long   ' uint16_t
Private wantedLatency As Long           ' uint16_t
Private wantedLatencies(0 To MAX_GAMEQUEUE_SLOTS - 1) As Long        ' uint16_t

Private waitingOnPlayersStartTime As Long   ' optional<uint32_t>
Private has_waitingOnPlayersStartTime As Boolean
Private Const MIN_WAITONPLAYERS_DISPLAYTIME_FOR_SPECTATORS As Long = GAME_TICKS_PER_UPDATE

' the C++ initialisers
Public Sub gtime_c_init()
    chosenLatency = GAME_TICKS_PER_UPDATE
    discreteChosenLatency = GAME_TICKS_PER_UPDATE
    wantedLatency = GAME_TICKS_PER_UPDATE
End Sub

' rational.h: Rational(numerator, denominator)'s reduction
Private Sub Rational_construct(n As Long, d As Long, ByVal numerator As Long, ByVal denominator As Long)
    Dim g As Long, a As Long, b As Long, t As Long
    a = numerator
    b = denominator
    Do While b <> 0
        a = a Mod b
        t = a: a = b: b = t
    Loop
    g = Abs(a)
    If denominator > 0 Then g = g * 1 Else g = g * -1
    n = numerator \ g
    d = denominator \ g
End Sub

' (VB) the elements begin..end (end excluded) printed with format, joined by separator
Private Function listToString(ByVal format As String, ByVal separator As String, list() As Long, ByVal begin As Long, ByVal end_ As Long) As String
    Dim ret As String, i As Long
    i = begin
    Do While i <> end_
        ret = ret & va(format, list(i))

        i = i + 1
        If i <> end_ Then
            ret = ret & separator
        End If
    Loop
    listToString = ret
End Function

' Initialise the game clock
Public Sub gameTimeInit()
    Dim player As Long

    ' Start the timer off at 2 so that when the scripts strip the map of objects
    ' for multiPlayer they will be processed as if they died.
    ' Setting game and graphics time.
    setGameTime 2

    ' Setting real time.
    realTime = wzGetTicks()
    deltaRealTime = 0
    prevRealTime = realTime
    realTimeFraction = 0!

    modifier_n = 1
    modifier_d = 1

    stopCount = 0

    chosenLatency = GAME_TICKS_PER_UPDATE * 2
    discreteChosenLatency = GAME_TICKS_PER_UPDATE * 2
    wantedLatency = GAME_TICKS_PER_UPDATE * 2
    For player = 0 To MAX_GAMEQUEUE_SLOTS - 1
        gameQueueCheckTime(player) = 0
        gameQueueCheckCrc(player) = 0
        wantedLatencies(player) = 0
    Next

    ' Don't let syncDebug from previous games cause a desynch dump at gameTime 102.
    crcError = False
    resetSyncDebug
End Sub

Public Sub setGameTime(ByVal newGameTime As Long)
    ' Setting game time.
    gameTime = newGameTime
    setPlayerGameTime NET_ALL_PLAYERS, newGameTime
    deltaGameTime = 0

    ' Setting graphics time to game time.
    graphicsTime = gameTime
    deltaGraphicsTime = 0
    graphicsTimeFraction = 0!

    ' Not setting real time.
End Sub

' (VB) MAX(1, timePeriod) of an int and a UDWORD compares unsigned
Private Function maxOneU(ByVal timePeriod As Long) As Long
    If ULT(1, timePeriod) Then maxOneU = timePeriod Else maxOneU = 1
End Function

' (VB) the UDWORD expression time % timePeriod * requiredRange / divisor
Private Function modularScaled(ByVal time_ As Long, ByVal timePeriod As Long, ByVal requiredRange As Long, ByVal divisor As Long) As Long
    Dim m As Double
    m = U32(time_) - Int(U32(time_) / U32(timePeriod)) * U32(timePeriod)
    modularScaled = FromU32(Int(U32(MUL32(FromU32(m), requiredRange)) / U32(divisor)))
End Function

Public Function getModularScaledGameTime(ByVal timePeriod As Long, ByVal requiredRange As Long) As Long
    getModularScaledGameTime = modularScaled(gameTime, timePeriod, requiredRange, maxOneU(timePeriod))
End Function

Public Function getModularScaledGraphicsTime(ByVal timePeriod As Long, ByVal requiredRange As Long) As Long
    getModularScaledGraphicsTime = modularScaled(graphicsTime, maxOneU(timePeriod), requiredRange, maxOneU(timePeriod))
End Function

Public Function getModularScaledRealTime(ByVal timePeriod As Long, ByVal requiredRange As Long) As Long
    getModularScaledRealTime = modularScaled(realTime, maxOneU(timePeriod), requiredRange, maxOneU(timePeriod))
End Function

Public Sub gameTimeUpdateBegin()
    deltaGameTime = 0
    deltaGraphicsTime = 0
    extraGameTicksProcessed = 0
End Sub

' Call this each loop to update the game timer
Public Function gameTimeUpdate(ByVal mayUpdate As Boolean, Optional ByVal forceTryGameTickUpdate As Boolean = False) As GameTimeUpdateResult
    Dim result As GameTimeUpdateResult, currTime As Long, newDeltaGraphicsTime As Long, newGraphicsTime As Long
    Dim timeForGameTickUpdate As Boolean, shouldDisplayWaitingStatus As Boolean, player As Long, waitStart As Long
    result = GameTimeUpdateResult_NO_UPDATE

    currTime = wzGetTicks()

    If ULT(currTime, prevRealTime) Then
        ' Warzone 2100, the first relativistic computer game!
        ' Exhibit A: Time travel
        ' force a rebase
        debug_ 175, LOG_WARNING, "gameTimeUpdate", "Time travel is occurring! Clock went back in time a bit from %d to %d!" & vbLf, prevRealTime, currTime
        prevRealTime = currTime
    End If

    ' Do not update the game time if gameTimeStop has been called
    If stopCount <> 0 Then
        gameTimeUpdate = GameTimeUpdateResult_NO_UPDATE
        Exit Function
    End If

    ' Calculate the new game time
    newDeltaGraphicsTime = quantiseFraction(modifier_n, modifier_d, currTime, prevRealTime)
    If Not (newDeltaGraphicsTime >= 0) Then ASSERT_FAILURE "newDeltaGraphicsTime >= 0", "lib/gamelib/gtime.cpp:187", "gameTimeUpdate", "Something very wrong."

    newGraphicsTime = FromU32(U32(graphicsTime) + newDeltaGraphicsTime)

    If ULT(gameTime, newGraphicsTime) And Not mayUpdate Then
        newGraphicsTime = gameTime
        newDeltaGraphicsTime = FromU32(U32(newGraphicsTime) - U32(graphicsTime))
    End If

    If updateWantedTime = 0 And Not ULT(newGraphicsTime, gameTime) Then
        updateWantedTime = currTime  ' This is the time that we wanted to tick.
    End If

    timeForGameTickUpdate = ULT(gameTime, newGraphicsTime)

    If (timeForGameTickUpdate Or forceTryGameTickUpdate) And Not checkPlayerGameTime(NET_ALL_PLAYERS) Then
        ' Pause time at current game time, since we are waiting GAME_GAME_TIME from other players.
        newGraphicsTime = gameTime
        newDeltaGraphicsTime = FromU32(U32(newGraphicsTime) - U32(graphicsTime))
        If Not has_waitingOnPlayersStartTime Then
            waitingOnPlayersStartTime = realTime
            has_waitingOnPlayersStartTime = True
        End If

        shouldDisplayWaitingStatus = True
        If NetPlay_players_isSpectator(selectedPlayer) And Not NetPlay_isHost Then
            If has_waitingOnPlayersStartTime Then waitStart = waitingOnPlayersStartTime Else waitStart = 0
            shouldDisplayWaitingStatus = ULT(MIN_WAITONPLAYERS_DISPLAYTIME_FOR_SPECTATORS, FromU32(U32(realTime) - U32(waitStart)))
        End If
        If timeForGameTickUpdate And shouldDisplayWaitingStatus Then
            debug_ 221, LOG_SYNC, "gameTimeUpdate", "Waiting for other players. gameTime = %u, player times are {%s}", gameTime, listToString("%u", ", ", gameQueueTime, 0, MAX_CONNECTED_PLAYERS)

            For player = 0 To MAX_CONNECTED_PLAYERS - 1
                If Not checkPlayerGameTime(player) Then
                    NETsetPlayerConnectionStatus CONNECTIONSTATUS_WAITING_FOR_PLAYER, player
                    Exit For  ' GAME_GAME_TIME is processed serially, so don't know if waiting for more players.
                End If
            Next
        End If

        timeForGameTickUpdate = False
        forceTryGameTickUpdate = False
    End If

    ' Adjust deltas.
    If timeForGameTickUpdate Or forceTryGameTickUpdate Then
        ' Update the game time.
        deltaGameTime = GAME_TICKS_PER_UPDATE
        gameTime = FromU32(U32(gameTime) + deltaGameTime)

        has_waitingOnPlayersStartTime = False

        If timeForGameTickUpdate Then result = GameTimeUpdateResult_GAME_TIME_UPDATED Else result = GameTimeUpdateResult_GAME_TIME_UPDATED_FORCED

        If Not timeForGameTickUpdate And forceTryGameTickUpdate Then
            extraGameTicksProcessed = extraGameTicksProcessed + 1
        End If

        updateLatency
        If crcError Then
            debug_ 256, LOG_ERROR, "gameTimeUpdate", "Synch error, gameTimes were: {%s}", listToString("%7u", ", ", gameQueueCheckTime, 0, MAX_GAMEQUEUE_SLOTS)
            debug_ 257, LOG_ERROR, "gameTimeUpdate", "Synch error, CRCs were:      {%s}", listToString(" 0x%04X", ", ", gameQueueCheckCrc, 0, MAX_GAMEQUEUE_SLOTS)
            crcError = False
        End If
    Else
        If extraGameTicksProcessed > 0 Then
            ' If extra game ticks have been processed, newGraphicsTime/newDeltaGraphicsTime must be "fast-forwarded" to appropriate values based on gameTime
            newGraphicsTime = gameTime
            newDeltaGraphicsTime = FromU32(U32(newGraphicsTime) - U32(graphicsTime))
        End If

        ' Update the graphics time.
        graphicsTime = newGraphicsTime
        deltaGraphicsTime = newDeltaGraphicsTime

        ' Update prevRealTime, since graphicsTime changed.
        prevRealTime = currTime
    End If

    ' Pre-calculate fraction used in timeAdjustedIncrement
    graphicsTimeFraction = CSng(U32(deltaGraphicsTime)) / CSng(GAME_TICKS_PER_SEC)

    If ULT(gameTime, graphicsTime) Then ASSERT_FAILURE "graphicsTime <= gameTime", "lib/gamelib/gtime.cpp:281", "gameTimeUpdate", "Trying to see the future."
    gameTimeUpdate = result
End Function

Public Sub gameTimeUpdateEnd()
    deltaGameTime = 0
End Sub

Public Sub realTimeUpdate()
    Dim currTime As Long
    currTime = wzGetTicks()

    ' now update realTime which does not pause
    ' Store the real time
    deltaRealTime = FromU32(U32(currTime) - U32(realTime))
    realTime = FromU32(U32(realTime) + U32(deltaRealTime))

    If Not ULT(deltaRealTime, GTIME_MAXFRAME) Then deltaRealTime = GTIME_MAXFRAME  ' Don't scroll across the map suddenly, if computer freezes for a moment.

    ' Pre-calculate fraction used in timeAdjustedIncrement
    realTimeFraction = CSng(U32(deltaRealTime)) / CSng(GAME_TICKS_PER_SEC)
End Sub

' reset the game time modifiers
Public Sub gameTimeResetMod()
    prevRealTime = wzGetTicks()

    modifier_n = 1
    modifier_d = 1
End Sub

' set the time modifier: void gameTimeSetMod(Rational mod), the Rational given by its numerator and denominator
Public Sub gameTimeSetMod(ByVal modNumerator As Long, ByVal modDenominator As Long)
    gameTimeResetMod
    Rational_construct modifier_n, modifier_d, modNumerator, modDenominator
End Sub

' get the current time modifier: Rational gameTimeGetMod(), as its numerator and denominator
Public Sub gameTimeGetMod(modNumerator As Long, modDenominator As Long)
    modNumerator = modifier_n
    modDenominator = modifier_d
End Sub

Public Function gameTimeIsStopped() As Boolean
    gameTimeIsStopped = stopCount <> 0
End Function

' Call this to stop the game timer
Public Sub gameTimeStop()
    stopCount = stopCount + 1

    graphicsTimeFraction = 0!
End Sub

' Call this to restart the game timer after a call to gameTimeStop
Public Sub gameTimeStart()
'   ASSERT(stopCount > 0, "Game started too many times.");

    prevRealTime = wzGetTicks()
    stopCount = stopCount - 1
    If stopCount < 0 Then stopCount = 0
End Sub

' Call this to reset the game timer
Public Sub gameTimeReset(ByVal time_ As Long)
    ' reset the game timers
    setGameTime time_
    gameTimeResetMod
    realTime = wzGetTicks()
    deltaRealTime = 0
End Sub

Public Sub getTimeComponents(ByVal time_ As Long, hours As Long, minutes As Long, seconds As Long, milliseconds As Long)
    Dim t As Double
    t = U32(time_)
    milliseconds = t - Int(t / GAME_TICKS_PER_SEC) * GAME_TICKS_PER_SEC
    t = Int(t / GAME_TICKS_PER_SEC)

    seconds = t - Int(t / 60) * 60
    t = Int(t / 60)

    minutes = t - Int(t / 60) * 60
    t = Int(t / 60)

    hours = t
End Sub

Private Sub updateLatency()
    Dim maxWantedLatency As Long, player As Long, prevDiscreteChosenLatency As Long, w As Double
    maxWantedLatency = 0
    prevDiscreteChosenLatency = discreteChosenLatency

    ' Find out what latency has been agreed on, next.
    For player = 0 To MAX_CONNECTED_PLAYERS - 1
        If NetPlay_players_allocated(player) Then  ' Don't wait for dropped/kicked players.
            If NetPlay_players_position(player) >= 0 And NetPlay_players_position(player) < game.maxPlayers Then ' should always be a valid map / player position
                If Not NetPlay_players_isSpectator(player) Or player = NetPlay_hostPlayer Then ' Don't wait for spectators (that are not the host)
                    'minWantedLatency = MIN(minWantedLatency, wantedLatencies[player]);  // Minimum, so the clients don't increase the latency to try to make one slow computer run faster.
                    If wantedLatencies(player) > maxWantedLatency Then maxWantedLatency = wantedLatencies(player)  ' Maximum, since the host experiences lower latency than everyone else.
                End If
            End If
        End If
    Next
    ' Adjust the agreed latency. (Can maximum decrease by 5ms or increase by 30ms per update.
    chosenLatency = (chosenLatency + clip_(maxWantedLatency - chosenLatency, -5, 60)) And &HFFFF&
    ' Round the chosen latency to an integer number of updates, up to 10.
    discreteChosenLatency = clip_(((chosenLatency + GAME_TICKS_PER_UPDATE \ 2) \ GAME_TICKS_PER_UPDATE) * GAME_TICKS_PER_UPDATE, GAME_TICKS_PER_UPDATE, GAME_TICKS_PER_UPDATE * GAME_UPDATES_PER_SEC)
    If prevDiscreteChosenLatency <> discreteChosenLatency Then
        debug_ 395, LOG_SYNC, "updateLatency", "Adjusting latency %d -> %d", prevDiscreteChosenLatency, discreteChosenLatency
    End If

    ' We want the chosen latency to increase by how much our update was delayed waiting for others, or to decrease by how long after we got the messages from others that it was time to tick. Plus a tiny 10ms buffer.
    ' We will send this number to others.
    ' (VB) the unsigned sum, converted to int, then clipped
    w = U32(discreteChosenLatency) + U32(updateReadyTime) - U32(updateWantedTime) + 10
    wantedLatency = clip_(FromU32(w), 0, 65535)

    ' Reset the times, ready to be set again.
    updateReadyTime = 0
    updateWantedTime = 0
End Sub

Public Sub sendPlayerGameTime()
    Dim player As Long, latencyTicks As Long, checkTime As Long, checkCrc As Long, w As MessageWriter
    latencyTicks = discreteChosenLatency \ GAME_TICKS_PER_UPDATE
    checkTime = gameTime
    checkCrc = nextDebugSync(gameTime)

    For player = 0 To MAX_CONNECTED_PLAYERS - 1
        If Not myResponsibility(player) Or Not isHumanPlayer(player) Then
            GoTo continue_
        End If

        Set w = NETbeginEncode(NETgameQueue(player), GAME_GAME_TIME)
        NETuint32_t w, latencyTicks
        NETuint32_t w, checkTime
        NETuint16_t w, checkCrc
        NETuint16_t w, wantedLatency
        NETend w
continue_:
    Next

    debugVerboseLogSyncIfNeeded gameTime
End Sub

Private Function shouldWaitForPlayerSlot(ByVal player As Long) As Boolean
    If NetPlay_players_allocated(player) Then  ' Don't wait for players that have been kicked/dropped.
        shouldWaitForPlayerSlot = Not NetPlay_players_isSpectator(player) Or Not ingame.has_TimeEveryoneIsInGame Or player = NetPlay_hostPlayer ' Don't wait for spectators (that are not the host) once the game has started
    End If
End Function

' publicly-exposed function
Public Function gtimeShouldWaitForPlayer(ByVal player As Long) As Boolean
    gtimeShouldWaitForPlayer = shouldWaitForPlayerSlot(player)
End Function

Private Function shouldCheckDebugSyncForPlayerSlot(ByVal player As Long) As Boolean
    If NetPlay_players_allocated(player) Then  ' human player
        If Not ingame.has_endTime Then         ' and game hasn't ended
            shouldCheckDebugSyncForPlayerSlot = Not NetPlay_players_isSpectator(player) Or player = NetPlay_hostPlayer
        End If
    End If
End Function

Public Sub recvPlayerGameTime(ByVal queue As Long)
    Dim latencyTicks As Long, checkTime As Long, checkCrc As Long, r As MessageReader, index As Long
    index = NETQUEUE_index(queue)
    If Not (index < MAX_GAMEQUEUE_SLOTS) Then ASSERT_FAILURE "queue.index < (((11 + 2) + 10) + 1)", "lib/gamelib/gtime.cpp:453", "recvPlayerGameTime", "Unexpected queue.index: %u", index

    latencyTicks = 0
    checkTime = 0
    checkCrc = 0

    Set r = NETbeginDecode(queue, GAME_GAME_TIME)
    NETuint32_t r, latencyTicks
    NETuint32_t r, checkTime
    NETuint16_t r, checkCrc
    NETuint16_t r, wantedLatencies(index)
    NETend r

    gameQueueTime(index) = FromU32(U32(checkTime) + U32(MUL32(latencyTicks, GAME_TICKS_PER_UPDATE)))  ' gameTime when future messages shall be processed.

    gameQueueCheckTime(index) = checkTime
    gameQueueCheckCrc(index) = checkCrc

    If shouldCheckDebugSyncForPlayerSlot(index) Then
        syncDebug "recvPlayerGameTime", "GAME_GAME_TIME p%d;lat%u,ct%u,crc%04X,wlat%u", index, latencyTicks, checkTime, checkCrc, wantedLatencies(index)
        If Not checkDebugSync(checkTime, checkCrc) Then
            debug_ 476, LOG_ERROR, "recvPlayerGameTime", "Found CRC error when receiving GAME_GAME_TIME for player: %u (checkTime: %u, checkCrc: %u)", index, checkTime, checkCrc
            crcError = True
            If NetPlay_players_allocated(index) Then
                NETsetPlayerConnectionStatus CONNECTIONSTATUS_DESYNC, index
            End If
        End If
    End If

    If updateReadyTime = 0 Then
        If checkPlayerGameTime(NET_ALL_PLAYERS) Then
            updateReadyTime = wzGetTicks()  ' This is the time we were able to tick.
        End If
    End If
End Sub

Public Function checkPlayerGameTime(ByVal player As Long) As Boolean
    Dim begin As Long, end_ As Long
    begin = player
    end_ = player + 1
    If player = NET_ALL_PLAYERS Then
        begin = 0
        end_ = MAX_CONNECTED_PLAYERS
    End If

    For player = begin To end_ - 1
        If ULT(gameQueueTime(player), gameTime) Then
            If shouldWaitForPlayerSlot(player) Then
                checkPlayerGameTime = False
                Exit Function  ' Still waiting for this player.
            End If
        End If
    Next

    checkPlayerGameTime = True  ' Have GAME_GAME_TIME from all players.
End Function

Public Sub setPlayerGameTime(ByVal player As Long, ByVal time_ As Long)
    If player = NET_ALL_PLAYERS Then
        For player = 0 To MAX_GAMEQUEUE_SLOTS - 1
            gameQueueTime(player) = time_
        Next
    Else
        If Not (player < MAX_GAMEQUEUE_SLOTS) Then ASSERT_FAILURE "player < (((11 + 2) + 10) + 1)", "lib/gamelib/gtime.cpp:523", "setPlayerGameTime", "Unexpected player: %u", player
        gameQueueTime(player) = time_
    End If
End Sub

' Returns numerator/denominator * (newTime - oldTime). Rounds up or down such that the average return value is right, if oldTime is always the previous newTime.
' (VB) the int64_t products are Doubles, exact below 2^53
Public Function quantiseFraction(ByVal numerator As Long, ByVal denominator As Long, ByVal newTime As Long, ByVal oldTime As Long) As Long
    Dim newValue As Double, oldValue As Double
    newValue = Fix(CDbl(newTime) * numerator / denominator)
    oldValue = Fix(CDbl(oldTime) * numerator / denominator)
    quantiseFraction = newValue - oldValue
End Function

' Returns numerator/denominator * (newTime - oldTime), for each of a Vector3i's coordinates.
Public Function quantiseFraction_vec(numerator As Vector3i, ByVal denominator As Long, ByVal newTime As Long, ByVal oldTime As Long) As Vector3i
    quantiseFraction_vec.x = quantiseFraction(numerator.x, denominator, newTime, oldTime)
    quantiseFraction_vec.y = quantiseFraction(numerator.y, denominator, newTime, oldTime)
    quantiseFraction_vec.z = quantiseFraction(numerator.z, denominator, newTime, oldTime)
End Function

' Returns the value times deltaGameTime, converted to seconds. The return value is rounded down.
Public Function gameTimeAdjustedIncrement(ByVal value As Long) As Long
    gameTimeAdjustedIncrement = value * deltaGameTime \ GAME_TICKS_PER_SEC
End Function

' Returns the value times deltaGraphicsTime, converted to seconds.
Public Function graphicsTimeAdjustedIncrement(ByVal value As Single) As Single
    graphicsTimeAdjustedIncrement = value * graphicsTimeFraction
End Function

' Returns the value times deltaGraphicsTime, converted to seconds, as a float.
Public Function realTimeAdjustedIncrement(ByVal value As Single) As Single
    realTimeAdjustedIncrement = value * realTimeFraction
End Function

' Returns the value times deltaRealTime, converted to seconds. The return value is rounded up or down to the nearest integer, such that it is exactly right on average.
Public Function realTimeAdjustedAverage(ByVal value As Long) As Long
    realTimeAdjustedAverage = quantiseFraction(value, GAME_TICKS_PER_SEC, FromU32(U32(realTime) + U32(deltaRealTime)), realTime)
End Function

' Returns the value times deltaGameTime, converted to seconds. The return value is rounded up or down, such that it is exactly right on average.
' gameTimeAdjustedAverage(int numerator, int denominator) is this with the denominator given.
' (VB) the int and int, int overloads are one function with an Optional denominator
Public Function gameTimeAdjustedAverage(ByVal numerator As Long, Optional ByVal denominator As Long = 1) As Long
    gameTimeAdjustedAverage = quantiseFraction(numerator, GAME_TICKS_PER_SEC * denominator, gameTime + deltaGameTime, gameTime)
End Function

' Returns the numerator/denominator times deltaGameTime, converted to seconds, for each of a Vector3i's coordinates.
Public Function gameTimeAdjustedAverage_vec(numerator As Vector3i, ByVal denominator As Long) As Vector3i
    gameTimeAdjustedAverage_vec = quantiseFraction_vec(numerator, GAME_TICKS_PER_SEC * denominator, gameTime + deltaGameTime, gameTime)
End Function
