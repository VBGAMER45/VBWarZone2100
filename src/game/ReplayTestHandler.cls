VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReplayTestHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB, for the tests) the replay options of vbtest_vb's maprecord and mapplay
' (statsref's StatsrefReplayOptionsHandler does the same): the game's own
' (WZGameReplayOptionsHandler) for the options saved and the player info, a
' built-in map's embedded data (none), and on restoring the game's
' restoreOptions without its level part (the tests load their maps by
' folder, not as levels) and without its version warning.
Option Explicit

Implements ReplayOptionsHandler

Private gameOptions As ReplayOptionsHandler

Private Sub Class_Initialize()
    Set gameOptions = New WZGameReplayOptionsHandler
End Sub

Private Function ReplayOptionsHandler_saveOptions(ByVal object_ As nlohmann_json) As Boolean
    ReplayOptionsHandler_saveOptions = gameOptions.saveOptions(object_)
End Function

' a built-in map: no embedded data
Private Function ReplayOptionsHandler_saveMap(dataVersion As Long, mapBinaryData() As Byte, mapBinaryDataSize As Long) As Boolean
    dataVersion = 1
    mapBinaryDataSize = 0
    ReplayOptionsHandler_saveMap = True
End Function

Private Function ReplayOptionsHandler_optionsUpdatePlayerInfo(ByVal object_ As nlohmann_json) As Boolean
    ReplayOptionsHandler_optionsUpdatePlayerInfo = gameOptions.optionsUpdatePlayerInfo(object_)
End Function

' WZGameReplayOptionsHandler::restoreOptions without the version warning and the level part
Private Function ReplayOptionsHandler_restoreOptions(ByVal object_ As nlohmann_json, ByVal dataVersion As Long, mapBinaryData() As Byte, ByVal mapBinaryDataSize As Long, ByVal replayNetcodeMajor As Long, ByVal replayNetcodeMinor As Long) As Boolean
    Dim netPlayers As nlohmann_json, i As Long, replaySpectatorIndex As Long

    ' random seed
    gameSRand object_.at("randSeed").get_uint()

    ' restore `game`
    MULTIPLAYERGAME_from_json object_.at("game"), game

    ' restore `ingame`
    MULTIPLAYERINGAME_from_json object_.at("ingame"), ingame

    ' restore NetPlay.players
    Set netPlayers = object_.at("netplay.players")
    For i = 0 To netPlayers.size() - 1
        PLAYER_from_json netPlayers.at_idx(i), i
    Next

    ' restore NetPlay.hostPlayer
    NetPlay_hostPlayer = object_.at("netplay.hostPlayer").get_uint()

    ' restore `NetPlay.bComms` (?)
    NetPlay_bComms = object_.at("netplay.bComms").get_bool()

    ' restore multistats
    If Not loadMultiStatsFromJSON(object_.at("multistats")) Then
        ReplayOptionsHandler_restoreOptions = False
        Exit Function
    End If

    ' Set various other initialization things (see NET_FIREUP)
    ingame.has_TimeEveryoneIsInGame = False      ' reset time
    ingame.has_endTime = False

    ' set selectedPlayer to a *new* spectator slot
    NetPlay_players_push_back
    replaySpectatorIndex = NetPlay_players_size() - 1
    NET_InitPlayer replaySpectatorIndex, False  ' re-init everything
    NetPlay_players_allocated(replaySpectatorIndex) = True
    NetPlay_players_name(replaySpectatorIndex) = "Replay Viewer"
    NetPlay_players_isSpectator(replaySpectatorIndex) = True

    selectedPlayer = replaySpectatorIndex
    realSelectedPlayer = selectedPlayer

    ' Need to initialize net/game queues here before we try to add messages to them
    For i = 0 To MAX_GAMEQUEUE_SLOTS - 1
        NETinitQueue NETgameQueue(i)
        NETsetNoSendOverNetwork NETgameQueue(i)
    Next

    ' simulate "NET_PLAYERRESPONDING" sent by all clients
    ' ensure that all humans players are set to "joined"
    For i = 0 To MAX_GAMEQUEUE_SLOTS - 2
        If isHumanPlayer(i) And ingame.JoiningInProgress(i) Then
            ingame.JoiningInProgress(i) = False
        End If
    Next
    NetPlay_isHostAlive = True ' lie and say the host is alive for multiplayer games

    ReplayOptionsHandler_restoreOptions = True
End Function

' skirmish: the default buffer
Private Function ReplayOptionsHandler_desiredBufferSize() As Double
    ReplayOptionsHandler_desiredBufferSize = 0
End Function

Private Function ReplayOptionsHandler_maximumEmbeddedMapBufferSize() As Double
    ReplayOptionsHandler_maximumEmbeddedMapBufferSize = gameOptions.maximumEmbeddedMapBufferSize()
End Function
