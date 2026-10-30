Attribute VB_Name = "activity_c"
' src/activity.cpp, activity.h: so far the ActivityManager's game mode (getCurrentGameMode, the game mode its
' startingGame / startingSavedGame / quitGame set): the music's mode, the end screen and the challenges read
' it. The activity sinks (Discord, the game history, the lobby events) are not kept (cluster J): those calls
' stay stand-ins in stubs_game, and the loaded level events are not cached for them.
Option Explicit

Private currentMode As ActivitySink_GameMode       ' = ActivitySink::GameMode::MENUS
Private bEndedCurrentMission As Boolean            ' = false

Private Function currentGameTypeToMode() As ActivitySink_GameMode
    Dim mode As ActivitySink_GameMode
    mode = ActivitySink_GameMode_CAMPAIGN
    If challengeActive Then
        mode = ActivitySink_GameMode_CHALLENGE
    ElseIf game.typ = LEVEL_TYPE_SKIRMISH Then
        If NetPlay_bComms Then mode = ActivitySink_GameMode_MULTIPLAYER Else mode = ActivitySink_GameMode_SKIRMISH
    ElseIf game.typ = LEVEL_TYPE_CAMPAIGN Then
        mode = ActivitySink_GameMode_CAMPAIGN
    End If
    currentGameTypeToMode = mode
End Function

Public Function ActivityManager_getCurrentGameMode() As ActivitySink_GameMode
    ActivityManager_getCurrentGameMode = currentMode
End Function

Public Sub ActivityManager_startingGame()
    Dim mode As ActivitySink_GameMode
    mode = currentGameTypeToMode()
    bEndedCurrentMission = False

    currentMode = mode
End Sub

Public Sub ActivityManager_startingSavedGame()
    Dim mode As ActivitySink_GameMode
    mode = currentGameTypeToMode()
    bEndedCurrentMission = False

    ' (mode SKIRMISH, CHALLENGE or a MULTIPLAYER replay: the synthesized updateMultiplayGameData event goes to the
    ' activity sinks, not kept)

    currentMode = mode

    ' (cachedLoadedLevelEvent: the loaded level events go to the sinks: none is cached)
End Sub

' (the game's end statistics, collectEndGameStatsData, go to the activity sinks: not kept)
Public Sub ActivityManager_quitGame(ByVal cheatsUsed As Boolean)
    If currentMode <> ActivitySink_GameMode_MENUS Then
        ' _endedMission(ActivitySink::GameEndReason::QUIT, stats, cheatsUsed): the sinks'
        bEndedCurrentMission = True
    End If

    currentMode = ActivitySink_GameMode_MENUS
End Sub
