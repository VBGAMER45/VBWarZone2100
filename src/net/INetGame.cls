VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "INetGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' What lib/netplay calls in the game (src, lib/gamelib): the exe implements
' this interface and hands it to NetSetGame at startup. Each method keeps the
' C++ name it stands for. The full list (about 60: the game clock, player
' join and leave, lobby options, player names and stats, quick chat, console
' messages) arrives with cluster J; see docs/PLAN.md. VB refuses underscores
' in the method names of an interface used with Implements, so each name
' drops its underscores and capitalises the letter after each one.
Option Explicit

' src/version.cpp version_getVersionString: the version string netplay sends and checks on joining.
Public Function versionGetVersionString() As String
End Function

' lib/gamelib/gtime.cpp checkPlayerGameTime: true if the GAME_GAME_TIME of player (NET_ALL_PLAYERS: of all) lets the game time move on.
Public Function checkPlayerGameTime(ByVal player As Long) As Boolean
End Function

' lib/gamelib/gtime.cpp recvPlayerGameTime: reads a GAME_GAME_TIME message from the game queue (a NETQUEUE Long).
Public Sub recvPlayerGameTime(ByVal queue As Long)
End Sub

' lib/gamelib/gtime.cpp gtimeShouldWaitForPlayer: whether the game waits for player's GAME_GAME_TIME.
Public Function gtimeShouldWaitForPlayer(ByVal player As Long) As Boolean
End Function

' lib/gamelib/gtime.cpp's realTime (uint32_t), which NETsetPlayerConnectionStatus reads.
Public Function realTime() As Long
End Function

' lib/gamelib/gtime.cpp's gameTime (uint32_t), which NETreplaySaveStop writes.
Public Function gameTime() As Long
End Function

' src/multiplay.cpp clearPlayerName (NET_InitPlayer).
Public Sub clearPlayerName(ByVal player As Long)
End Sub

' src/component.cpp setPlayerColour (NET_InitPlayer).
Public Function setPlayerColour(ByVal player As Long, ByVal col As Long) As Boolean
End Function

' src/multiplay.cpp's game.maxPlayers (uint8_t; NET_InitPlayer, playersPerTeam).
Public Function gameMaxPlayers() As Long
End Function

' src/multiplay.cpp's ingame.localJoiningInProgress (NET_InitPlayer).
Public Function ingameLocalJoiningInProgress() As Boolean
End Function

' netplay.cpp initPlayerNetworkProps's part in src/multiplay.cpp's ingame, for playerIndex (< MAX_CONNECTED_PLAYERS):
' JoiningInProgress and PendingDisconnect false, joinTimes, lastReadyTimes, lastNotReadyTimes and playerLeftGameTime
' reset, secondsNotReady 0, hostChatPermissions set.
Public Sub initPlayerNetworkPropsIngame(ByVal playerIndex As Long, ByVal hostChatPermissions As Boolean)
End Sub

' src/multiint.cpp changeColour (NEThostGame without comms: the host's colour).
Public Function changeColour(ByVal player As Long, ByVal col As Long, ByVal responsibleIdx As Long) As Boolean
End Function

' src/warzoneconfig.cpp war_getMPcolour (NEThostGame).
Public Function warGetMPcolour() As Long
End Function

' src/multiplay.cpp isBlindPlayerInfoState (NETSendNPlayerInfoTo).
Public Function isBlindPlayerInfoState() As Boolean
End Function

' src/multiplay.cpp getPlayerGenericName (NETSendNPlayerInfoTo).
Public Function getPlayerGenericName(ByVal player As Long) As String
End Function
