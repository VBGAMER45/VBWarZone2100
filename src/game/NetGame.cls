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

'--------------------------------------------------------------------------
' cluster J: the rest of netplay.cpp's calls. An EcKey goes to WzNet as its bytes.

Private Sub INetGame_addConsoleMessage(ByVal text As String)
    addConsoleMessage text, DEFAULT_JUSTIFY, NOTIFY_MESSAGE
End Sub

Private Function INetGame_formatGameName(ByVal name_ As String) As String
    INetGame_formatGameName = formatGameName(name_)
End Function

Private Function INetGame_gameHashString() As String
    If Not Sha256_isZero(game.hash) Then INetGame_gameHashString = Sha256_toString(game.hash)
End Function

Private Function INetGame_gameMap() As String
    INetGame_gameMap = game.map
End Function

Private Function INetGame_gameBlindMode() As Long
    INetGame_gameBlindMode = game.blindMode
End Function

Private Function INetGame_gameMapHasScavengers() As Boolean
    INetGame_gameMapHasScavengers = game.mapHasScavengers
End Function

Private Function INetGame_ingameFlags() As Long
    INetGame_ingameFlags = ingame.flags
End Function

Private Sub INetGame_setIngameLocalJoiningInProgress(ByVal value As Boolean)
    ingame.localJoiningInProgress = value
End Sub

Private Sub INetGame_ingameSetMuteChat(ByVal player As Long, ByVal value As Boolean)
    ingame.muteChat(player) = value
End Sub

Private Function INetGame_ingameVerifiedIdentity(ByVal player As Long) As Boolean
    INetGame_ingameVerifiedIdentity = ingame.VerifiedIdentity(player)
End Function

Private Function INetGame_ingameHostChatPermissions(ByVal player As Long) As Boolean
    INetGame_ingameHostChatPermissions = ingame.hostChatPermissions(player)
End Function

' swapPlayerIndexes: "Swap certain ingame player-associated entries"
Private Sub INetGame_ingameSwapPlayerEntries(ByVal a As Long, ByVal b As Long)
    Dim l As Long, f As Boolean, c As Currency, d As Double
    l = ingame.PingTimes(a): ingame.PingTimes(a) = ingame.PingTimes(b): ingame.PingTimes(b) = l
    l = ingame.LagCounter(a): ingame.LagCounter(a) = ingame.LagCounter(b): ingame.LagCounter(b) = l
    l = ingame.DesyncCounter(a): ingame.DesyncCounter(a) = ingame.DesyncCounter(b): ingame.DesyncCounter(b) = l
    f = ingame.VerifiedIdentity(a): ingame.VerifiedIdentity(a) = ingame.VerifiedIdentity(b): ingame.VerifiedIdentity(b) = f
    f = ingame.JoiningInProgress(a): ingame.JoiningInProgress(a) = ingame.JoiningInProgress(b): ingame.JoiningInProgress(b) = f
    f = ingame.PendingDisconnect(a): ingame.PendingDisconnect(a) = ingame.PendingDisconnect(b): ingame.PendingDisconnect(b) = f
    f = ingame.DataIntegrity(a): ingame.DataIntegrity(a) = ingame.DataIntegrity(b): ingame.DataIntegrity(b) = f
    f = ingame.hostChatPermissions(a): ingame.hostChatPermissions(a) = ingame.hostChatPermissions(b): ingame.hostChatPermissions(b) = f

    f = ingame.has_joinTimes(a): ingame.has_joinTimes(a) = ingame.has_joinTimes(b): ingame.has_joinTimes(b) = f
    c = ingame.joinTimes(a): ingame.joinTimes(a) = ingame.joinTimes(b): ingame.joinTimes(b) = c
    f = ingame.has_lastReadyTimes(a): ingame.has_lastReadyTimes(a) = ingame.has_lastReadyTimes(b): ingame.has_lastReadyTimes(b) = f
    c = ingame.lastReadyTimes(a): ingame.lastReadyTimes(a) = ingame.lastReadyTimes(b): ingame.lastReadyTimes(b) = c
    f = ingame.has_lastNotReadyTimes(a): ingame.has_lastNotReadyTimes(a) = ingame.has_lastNotReadyTimes(b): ingame.has_lastNotReadyTimes(b) = f
    c = ingame.lastNotReadyTimes(a): ingame.lastNotReadyTimes(a) = ingame.lastNotReadyTimes(b): ingame.lastNotReadyTimes(b) = c
    d = ingame.secondsNotReady(a): ingame.secondsNotReady(a) = ingame.secondsNotReady(b): ingame.secondsNotReady(b) = d
    f = ingame.has_playerLeftGameTime(a): ingame.has_playerLeftGameTime(a) = ingame.has_playerLeftGameTime(b): ingame.has_playerLeftGameTime(b) = f
    l = ingame.playerLeftGameTime(a): ingame.playerLeftGameTime(a) = ingame.playerLeftGameTime(b): ingame.playerLeftGameTime(b) = l

    f = ingame.has_lastSentPlayerDataCheck2(a): ingame.has_lastSentPlayerDataCheck2(a) = ingame.has_lastSentPlayerDataCheck2(b): ingame.has_lastSentPlayerDataCheck2(b) = f
    c = ingame.lastSentPlayerDataCheck2(a): ingame.lastSentPlayerDataCheck2(a) = ingame.lastSentPlayerDataCheck2(b): ingame.lastSentPlayerDataCheck2(b) = c
    f = ingame.muteChat(a): ingame.muteChat(a) = ingame.muteChat(b): ingame.muteChat(b) = f
End Sub

Private Function INetGame_getHostLaunch() As Long
    INetGame_getHostLaunch = getHostLaunch()
End Function

Private Function INetGame_getLockedOptionsName() As Boolean
    INetGame_getLockedOptionsName = getLockedOptions().name_
End Function

Private Function INetGame_getPlayerName(ByVal player As Long) As String
    INetGame_getPlayerName = getPlayerName(player)
End Function

Private Function INetGame_setPlayerName(ByVal player As Long, ByVal sName As String) As Boolean
    INetGame_setPlayerName = setPlayerName(player, sName)
End Function

Private Function INetGame_isHumanPlayer(ByVal player As Long) As Boolean
    INetGame_isHumanPlayer = isHumanPlayer(player)
End Function

Private Function INetGame_scavengerSlot() As Long
    INetGame_scavengerSlot = scavengerSlot()
End Function

Private Function INetGame_responsibleFor(ByVal player As Long, ByVal playerinquestion As Long) As Boolean
    INetGame_responsibleFor = responsibleFor(player, playerinquestion)
End Function

Private Sub INetGame_resetReadyStatus(ByVal bSendOptions As Boolean, ByVal ignoreReadyReset As Boolean)
    resetReadyStatus bSendOptions, ignoreReadyReset
End Sub

Private Function INetGame_shouldSkipReadyResetOnPlayerJoinLeaveEvent() As Boolean
    INetGame_shouldSkipReadyResetOnPlayerJoinLeaveEvent = shouldSkipReadyResetOnPlayerJoinLeaveEvent()
End Function

Private Sub INetGame_sendInGameSystemMessage(ByVal text As String)
    sendInGameSystemMessage text
End Sub

Private Sub INetGame_printConsoleNameChange(ByVal oldName As String, ByVal newName As String)
    printConsoleNameChange oldName, newName
End Sub

Private Sub INetGame_multiSyncPlayerSwap(ByVal playerIndexA As Long, ByVal playerIndexB As Long)
    multiSyncPlayerSwap playerIndexA, playerIndexB
End Sub

Private Function INetGame_NetworkTextMessageDecode(ByVal message As NetMessage, ByVal senderIdx As Long, sender As Long, text As String) As Boolean
    Dim m As NetworkTextMessage
    If NetworkTextMessage_decode(m, MessageReader_new(message), senderIdx) Then
        sender = m.sender
        text = m.text
        INetGame_NetworkTextMessageDecode = True
    End If
End Function

Private Function INetGame_decodeSpecInGameTextMessage(ByVal message As NetMessage, ByVal senderIdx As Long, sender As Long, text As String) As Boolean
    Dim m As NetworkTextMessage
    If decodeSpecInGameTextMessage(MessageReader_new(message), senderIdx, m) Then
        sender = m.sender
        text = m.text
        INetGame_decodeSpecInGameTextMessage = True
    End If
End Function

Private Function INetGame_multiplayPlayersShouldCheckReady() As Boolean
    INetGame_multiplayPlayersShouldCheckReady = multiplayPlayersShouldCheckReady()
End Function

Private Sub INetGame_handlePossiblePlayersShouldCheckReadyChange(ByVal previousPlayersShouldCheckReadyValue As Boolean)
    handlePossiblePlayersShouldCheckReadyChange previousPlayersShouldCheckReadyValue
End Sub

Private Sub INetGame_handleAutoReadyRequest()
    handleAutoReadyRequest
End Sub

Private Sub INetGame_kickPlayer(ByVal player_id As Long, ByVal reason As String, ByVal typ As Long, ByVal banPlayer As Boolean)
    kickPlayer player_id, reason, typ, banPlayer
End Sub

Private Sub INetGame_sendRoomSystemMessage(ByVal text As String)
    sendRoomSystemMessage text
End Sub

Private Function INetGame_MultiPlayerJoin(ByVal playerIndex As Long, ByVal hasVerifiedJoinIdentity As Boolean, verifiedJoinIdentity() As Byte, ByVal identityLen As Long) As Boolean
    INetGame_MultiPlayerJoin = MultiPlayerJoin(playerIndex, hasVerifiedJoinIdentity, verifiedJoinIdentity, identityLen)
End Function

' (netplay does not use the result: MultiPlayerLeave is called as a statement, which stubs_ui's Sub and multijoin's
' Function both allow)
Private Function INetGame_MultiPlayerLeave(ByVal playerIndex As Long) As Boolean
    MultiPlayerLeave playerIndex
    INetGame_MultiPlayerLeave = True
End Function

Private Sub INetGame_ShowLobbyStatusMessage(msgs() As String, ByVal count As Long)
    ShowLobbyStatusMessage msgs, count
End Sub

Private Function INetGame_getMultiStatsIdentity(ByVal player As Long, key() As Byte) As Long
    Dim stats As PLAYERSTATS_
    stats = getMultiStats(player)
    INetGame_getMultiStatsIdentity = EcKey_toBytes(stats.identity, EcKey_Public, key)
End Function

Private Function INetGame_getLocalSharedIdentity(key() As Byte) As Long
    Dim k As EcKey
    k = getLocalSharedIdentity()
    If EcKey_hasPrivate(k) Then
        INetGame_getLocalSharedIdentity = EcKey_toBytes(k, EcKey_Private, key)
    Else
        INetGame_getLocalSharedIdentity = EcKey_toBytes(k, EcKey_Public, key)
    End If
End Function

Private Function INetGame_getTruePlayerIdentity(ByVal player As Long, key() As Byte) As Long
    Dim t As TrueIdentity
    t = getTruePlayerIdentity(player)
    INetGame_getTruePlayerIdentity = EcKey_toBytes(t.identity, EcKey_Public, key)
End Function

Private Function INetGame_getOutputPlayerIdentity(ByVal player As Long, key() As Byte) As Long
    Dim k As EcKey
    k = getOutputPlayerIdentity(player)
    INetGame_getOutputPlayerIdentity = EcKey_toBytes(k, EcKey_Public, key)
End Function

Private Function INetGame_swapPlayerMultiStatsLocal(ByVal playerIndexA As Long, ByVal playerIndexB As Long) As Boolean
    INetGame_swapPlayerMultiStatsLocal = swapPlayerMultiStatsLocal(playerIndexA, playerIndexB)
End Function

Private Function INetGame_recvMultiStats(ByVal queue As Long) As Boolean
    INetGame_recvMultiStats = recvMultiStats(queue)
End Function

Private Sub INetGame_resetRecentScoreData()
    resetRecentScoreData
End Sub

Private Sub INetGame_playerSpamMuteReset(ByVal playerIndex As Long)
    playerSpamMuteReset playerIndex
End Sub

Private Sub INetGame_playerSpamMuteNotifyIndexSwap(ByVal playerIndexA As Long, ByVal playerIndexB As Long)
    playerSpamMuteNotifyIndexSwap playerIndexA, playerIndexB
End Sub

Private Sub INetGame_sendQuickChat(ByVal targetPlayer As Long)
    Dim t As WzQuickChatTargeting
    WzQuickChatTargeting_insertSpecificPlayer t, targetPlayer       ' WzQuickChatTargeting::targetSpecificPlayer(targetPlayer)
    sendQuickChat WzQuickChatMessage_INTERNAL_MSG_DELIVERY_FAILURE_TRY_AGAIN, realSelectedPlayer, t
End Sub

Private Function INetGame_wzCommandInterfaceEnabled() As Boolean
    INetGame_wzCommandInterfaceEnabled = wz_command_interface_enabled()
End Function

Private Sub INetGame_wzCommandInterfaceOutput(ByVal str_ As String)
    wz_command_interface_output "%s", str_
End Sub

Private Sub INetGame_wzCommandInterfaceOutputRoomStatusJson(ByVal queued As Boolean)
    wz_command_interface_output_room_status_json queued
End Sub

' (cluster J part 2) the lobby hosting handler: netlobby_c's netplay crossing
Private Sub INetGame_lobbyHostCreate(ByVal lobbyServerAddress As String)
    netlobby_netplayHostCreate lobbyServerAddress
End Sub

Private Sub INetGame_lobbyHostReset()
    netlobby_netplayHostReset
End Sub

Private Function INetGame_lobbyHostCreateGameListing(ByVal identityPlayer As Long, ByVal gameDetailsJson As String, ByVal botProt As Long, _
        ByVal proxyIPs As Long, ByVal hostingIPs As Long, ByVal connectionsJson As String) As Boolean
    INetGame_lobbyHostCreateGameListing = netlobby_netplayHostCreateGameListing(identityPlayer, gameDetailsJson, botProt, proxyIPs, hostingIPs, connectionsJson)
End Function

Private Sub INetGame_lobbyHostUpdateGameListing(ByVal identityPlayer As Long, ByVal gameDetailsJson As String, ByVal botProt As Long, _
        ByVal proxyIPs As Long, ByVal hostingIPs As Long)
    netlobby_netplayHostUpdateGameListing identityPlayer, gameDetailsJson, botProt, proxyIPs, hostingIPs
End Sub

Private Function INetGame_lobbyHostRespondToLobbyConnectionCheckRequest(ByVal jsonRequest As String, ByVal playerName As String, _
        ByVal identityPlayer As Long, response As String, errorCode As Long) As Boolean
    INetGame_lobbyHostRespondToLobbyConnectionCheckRequest = netlobby_netplayHostRespondToLobbyConnectionCheckRequest(jsonRequest, playerName, identityPlayer, response, errorCode)
End Function

Private Function INetGame_lobbyHostCheckJoinRequest(ByVal joinToken As String, ByVal ip As String, ByVal playerName As String, _
        ByVal playerPublicIdentity As String, ByVal connectionMethod As Long, ByVal asSpectator As Boolean, _
        ByVal tmpSocketIndex As Long, ByVal uniqueJoinID As String) As Boolean
    INetGame_lobbyHostCheckJoinRequest = netlobby_netplayHostCheckJoinRequest(joinToken, ip, playerName, playerPublicIdentity, connectionMethod, asSpectator, tmpSocketIndex, uniqueJoinID)
End Function

Private Sub INetGame_lobbyHostRemoveGameListing(ByVal matchStarted As Boolean)
    netlobby_netplayHostRemoveGameListing matchStarted
End Sub

Private Sub INetGame_lobbyHostRun()
    netlobby_netplayHostRun
End Sub

Private Function INetGame_lobbyHostIsConnectedToLobby() As Boolean
    INetGame_lobbyHostIsConnectedToLobby = netlobby_netplayHostIsConnectedToLobby()
End Function

Private Function INetGame_lobbyHostGetCurrentGameId(gameId As String) As Boolean
    INetGame_lobbyHostGetCurrentGameId = netlobby_netplayHostGetCurrentGameId(gameId)
End Function

Private Function INetGame_lobbyHostIsHostJoinOptionSupportedByLobby(ByVal optType As Long, result As Boolean) As Boolean
    INetGame_lobbyHostIsHostJoinOptionSupportedByLobby = netlobby_netplayHostIsHostJoinOptionSupportedByLobby(optType, result)
End Function

' getLoadedMods(): each mod's name and LoadedMod::getHash() (computed as getHash does when still zero; "" for zero)
Private Function INetGame_getLoadedMods(names() As String, hashes() As String) As Long
    Dim mods() As WzMods_LoadedMod, n As Long, i As Long
    n = getLoadedMods(mods)
    If n > 0 Then
        ReDim names(0 To n - 1)
        ReDim hashes(0 To n - 1)
    End If
    For i = 0 To n - 1
        names(i) = mods(i).name_
        If Sha256_isZero(mods(i).fileHash) Then findHashOfFile mods(i).filename, mods(i).fileHash
        If Not Sha256_isZero(mods(i).fileHash) Then hashes(i) = Sha256_toString(mods(i).fileHash) Else hashes(i) = ""
    Next
    INetGame_getLoadedMods = n
End Function

Private Function INetGame_warGetLobbyDisableIPv6() As Boolean
    INetGame_warGetLobbyDisableIPv6 = war_getLobbyDisableIPv6()
End Function
