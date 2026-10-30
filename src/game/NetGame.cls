VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NetGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' The game side of WzNet.dll's INetGame: each method calls the ported game
' function of the same name.
Option Explicit

Implements INetGame

Private Function INetGame_versionGetVersionString() As String
    INetGame_versionGetVersionString = version_getVersionString()
End Function

Private Function INetGame_checkPlayerGameTime(ByVal player As Long) As Boolean
    INetGame_checkPlayerGameTime = checkPlayerGameTime(player)
End Function

Private Sub INetGame_recvPlayerGameTime(ByVal queue As Long)
    recvPlayerGameTime queue
End Sub

Private Function INetGame_gtimeShouldWaitForPlayer(ByVal player As Long) As Boolean
    INetGame_gtimeShouldWaitForPlayer = gtimeShouldWaitForPlayer(player)
End Function

Private Function INetGame_realTime() As Long
    INetGame_realTime = realTime
End Function

Private Function INetGame_gameTime() As Long
    INetGame_gameTime = gameTime
End Function

Private Sub INetGame_clearPlayerName(ByVal player As Long)
    clearPlayerName player
End Sub

Private Function INetGame_setPlayerColour(ByVal player As Long, ByVal col As Long) As Boolean
    INetGame_setPlayerColour = setPlayerColour(player, col)
End Function

Private Function INetGame_gameMaxPlayers() As Long
    INetGame_gameMaxPlayers = game.maxPlayers
End Function

Private Function INetGame_ingameLocalJoiningInProgress() As Boolean
    INetGame_ingameLocalJoiningInProgress = ingame.localJoiningInProgress
End Function

' initPlayerNetworkProps' part in ingame
Private Sub INetGame_initPlayerNetworkPropsIngame(ByVal playerIndex As Long, ByVal hostChatPermissions As Boolean)
    ingame.JoiningInProgress(playerIndex) = False
    ingame.PendingDisconnect(playerIndex) = False
    ingame.has_joinTimes(playerIndex) = False
    ingame.has_lastReadyTimes(playerIndex) = False
    ingame.has_lastNotReadyTimes(playerIndex) = False
    ingame.secondsNotReady(playerIndex) = 0
    ingame.has_playerLeftGameTime(playerIndex) = False
    ingame.hostChatPermissions(playerIndex) = hostChatPermissions
End Sub

Private Function INetGame_changeColour(ByVal player As Long, ByVal col As Long, ByVal responsibleIdx As Long) As Boolean
    INetGame_changeColour = changeColour(player, col, responsibleIdx)
End Function

Private Function INetGame_warGetMPcolour() As Long
    INetGame_warGetMPcolour = war_getMPcolour()
End Function

Private Function INetGame_isBlindPlayerInfoState() As Boolean
    INetGame_isBlindPlayerInfoState = isBlindPlayerInfoState()
End Function

Private Function INetGame_getPlayerGenericName(ByVal player As Long) As String
    INetGame_getPlayerGenericName = getPlayerGenericName(player)
End Function
