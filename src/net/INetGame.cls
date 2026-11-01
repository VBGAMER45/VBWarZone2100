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
' this interface (NetGame.cls) and hands it to NetSetGame at startup. Each
' method keeps the C++ name it stands for. VB refuses underscores in the
' method names of an interface used with Implements, so each name drops its
' underscores and capitalises the letter after each one; a field of game or
' ingame is game<Field> / ingame<Field>. An EcKey crosses as its bytes (the
' function's result is the size, the bytes in key(); 0: an empty key).
' Not here: the activity sinks (ActivityManager, not ported), whose calls are
' comments in WzNet; the port mapping is WzNet's own (port_mapping_manager_c).
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

' lib/gamelib/gtime.cpp's realTime (uint32_t).
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

'--------------------------------------------------------------------------
' cluster J: the rest of netplay.cpp's calls

' src/console.cpp addConsoleMessage(text, DEFAULT_JUSTIFY, NOTIFY_MESSAGE) (NETinit's backend message).
Public Sub addConsoleMessage(ByVal text As String)
End Sub

' src/multiint.cpp formatGameName (NETsetLobbyOptField's map name).
Public Function formatGameName(ByVal name_ As String) As String
End Function

' game.hash: (!game.hash.isZero()) ? game.hash.toString() : "" (the lobby game details).
Public Function gameHashString() As String
End Function

' game.map (the lobby game details).
Public Function gameMap() As String
End Function

' game.blindMode (a BLIND_MODE; NETallowJoining's NET_ACCEPTED).
Public Function gameBlindMode() As Long
End Function

' game.mapHasScavengers (NET_findSpectatorSlotToOpen).
Public Function gameMapHasScavengers() As Boolean
End Function

' ingame.flags (uint8_t: the structure limits; NETsetLobbyLimitFlags).
Public Function ingameFlags() As Long
End Function

' ingame.localJoiningInProgress = value (NETclientHandleHostDisconnected).
Public Sub setIngameLocalJoiningInProgress(ByVal value As Boolean)
End Sub

' ingame.muteChat[player] = value (NET_InitPlayers).
Public Sub ingameSetMuteChat(ByVal player As Long, ByVal value As Boolean)
End Sub

' ingame.VerifiedIdentity[player].
Public Function ingameVerifiedIdentity(ByVal player As Long) As Boolean
End Function

' ingame.hostChatPermissions[player].
Public Function ingameHostChatPermissions(ByVal player As Long) As Boolean
End Function

' swapPlayerIndexes' swaps of ingame's player-associated entries (PingTimes, LagCounter, DesyncCounter,
' VerifiedIdentity, JoiningInProgress, PendingDisconnect, DataIntegrity, hostChatPermissions, joinTimes,
' lastReadyTimes, lastNotReadyTimes, secondsNotReady, playerLeftGameTime, lastSentPlayerDataCheck2, muteChat).
Public Sub ingameSwapPlayerEntries(ByVal playerIndexA As Long, ByVal playerIndexB As Long)
End Sub

' src/wrappers.cpp getHostLaunch (a HostLaunch).
Public Function getHostLaunch() As Long
End Function

' src/multiint.cpp getLockedOptions().name.
Public Function getLockedOptionsName() As Boolean
End Function

' src/multiplay.cpp getPlayerName(player).
Public Function getPlayerName(ByVal player As Long) As String
End Function

' src/multiplay.cpp setPlayerName.
Public Function setPlayerName(ByVal player As Long, ByVal sName As String) As Boolean
End Function

' src/multiplay.cpp isHumanPlayer.
Public Function isHumanPlayer(ByVal player As Long) As Boolean
End Function

' src/multiplay.cpp scavengerSlot.
Public Function scavengerSlot() As Long
End Function

' src/multiplay.cpp responsibleFor.
Public Function responsibleFor(ByVal player As Long, ByVal playerinquestion As Long) As Boolean
End Function

' src/multiplay.cpp resetReadyStatus.
Public Sub resetReadyStatus(ByVal bSendOptions As Boolean, ByVal ignoreReadyReset As Boolean)
End Sub

' src/multiplay.cpp shouldSkipReadyResetOnPlayerJoinLeaveEvent.
Public Function shouldSkipReadyResetOnPlayerJoinLeaveEvent() As Boolean
End Function

' src/multiplay.cpp sendInGameSystemMessage.
Public Sub sendInGameSystemMessage(ByVal text As String)
End Sub

' src/multiplay.cpp printConsoleNameChange.
Public Sub printConsoleNameChange(ByVal oldName As String, ByVal newName As String)
End Sub

' src/multiplay.cpp multiSyncPlayerSwap.
Public Sub multiSyncPlayerSwap(ByVal playerIndexA As Long, ByVal playerIndexB As Long)
End Sub

' NetworkTextMessage::decode(MessageReader(message), senderIdx) (LogChatMsg): the decoded sender and text.
Public Function NetworkTextMessageDecode(ByVal message As NetMessage, ByVal senderIdx As Long, sender As Long, text As String) As Boolean
End Function

' src/multiplay.cpp decodeSpecInGameTextMessage(MessageReader(message), senderIdx) (LogChatMsg): has_value(), the
' decoded sender and text.
Public Function decodeSpecInGameTextMessage(ByVal message As NetMessage, ByVal senderIdx As Long, sender As Long, text As String) As Boolean
End Function

' src/multiint.cpp multiplayPlayersShouldCheckReady.
Public Function multiplayPlayersShouldCheckReady() As Boolean
End Function

' src/multiint.cpp handlePossiblePlayersShouldCheckReadyChange.
Public Sub handlePossiblePlayersShouldCheckReadyChange(ByVal previousPlayersShouldCheckReadyValue As Boolean)
End Sub

' src/multiint.cpp handleAutoReadyRequest.
Public Sub handleAutoReadyRequest()
End Sub

' src/multiint.cpp kickPlayer (a LOBBY_ERROR_TYPES).
Public Sub kickPlayer(ByVal player_id As Long, ByVal reason As String, ByVal typ As Long, ByVal banPlayer As Boolean)
End Sub

' src/multiint.cpp sendRoomSystemMessage.
Public Sub sendRoomSystemMessage(ByVal text As String)
End Sub

' src/multijoin.cpp MultiPlayerJoin(playerIndex, optional<EcKey::Key> verifiedJoinIdentity): has_value() and the key's
' bytes verifiedJoinIdentity(0 .. identityLen - 1).
Public Function MultiPlayerJoin(ByVal playerIndex As Long, ByVal hasVerifiedJoinIdentity As Boolean, verifiedJoinIdentity() As Byte, ByVal identityLen As Long) As Boolean
End Function

' src/multijoin.cpp MultiPlayerLeave.
Public Function MultiPlayerLeave(ByVal playerIndex As Long) As Boolean
End Function

' src/multijoin.cpp ShowLobbyStatusMessage(const std::vector<std::string>& msgs): msgs(0 .. count - 1).
Public Sub ShowLobbyStatusMessage(msgs() As String, ByVal count As Long)
End Sub

' src/multistat.cpp getMultiStats(player).identity.toBytes(EcKey::Public).
Public Function getMultiStatsIdentity(ByVal player As Long, key() As Byte) As Long
End Function

' src/multistat.cpp getLocalSharedIdentity(): its private key's bytes (toBytes(EcKey::Private)) when it has one, else its
' public key's.
Public Function getLocalSharedIdentity(key() As Byte) As Long
End Function

' src/multistat.cpp getTruePlayerIdentity(player).identity.toBytes(EcKey::Public).
Public Function getTruePlayerIdentity(ByVal player As Long, key() As Byte) As Long
End Function

' src/multistat.cpp getOutputPlayerIdentity(player).toBytes(EcKey::Public).
Public Function getOutputPlayerIdentity(ByVal player As Long, key() As Byte) As Long
End Function

' src/multistat.cpp swapPlayerMultiStatsLocal.
Public Function swapPlayerMultiStatsLocal(ByVal playerIndexA As Long, ByVal playerIndexB As Long) As Boolean
End Function

' src/multistat.cpp recvMultiStats.
Public Function recvMultiStats(ByVal queue As Long) As Boolean
End Function

' src/multistat.cpp resetRecentScoreData.
Public Sub resetRecentScoreData()
End Sub

' src/hci/quickchat.cpp playerSpamMuteReset.
Public Sub playerSpamMuteReset(ByVal playerIndex As Long)
End Sub

' src/hci/quickchat.cpp playerSpamMuteNotifyIndexSwap.
Public Sub playerSpamMuteNotifyIndexSwap(ByVal playerIndexA As Long, ByVal playerIndexB As Long)
End Sub

' src/hci/quickchat.cpp sendQuickChat(WzQuickChatMessage::INTERNAL_MSG_DELIVERY_FAILURE_TRY_AGAIN, realSelectedPlayer,
' WzQuickChatTargeting::targetSpecificPlayer(targetPlayer)) (NETFilterMessageWhileSwappingPlayer).
Public Sub sendQuickChat(ByVal targetPlayer As Long)
End Sub

' src/stdinreader.cpp wz_command_interface_enabled.
Public Function wzCommandInterfaceEnabled() As Boolean
End Function

' src/stdinreader.cpp wz_command_interface_output (the text formatted already).
Public Sub wzCommandInterfaceOutput(ByVal str_ As String)
End Sub

' src/stdinreader.cpp wz_command_interface_output_room_status_json.
Public Sub wzCommandInterfaceOutputRoomStatusJson(ByVal queued As Boolean)
End Sub

' (cluster J part 2) lib/netplay/netlobby.cpp's lobby hosting handler, in the exe (its netlobby_c): netplay.cpp's
' lobbyHostingHandler (one at a time). Its event callbacks come back through NetApi (NETlobbyGameCreatedCallback,
' NETlobbyGamePublishedCallback, NETlobbyGameGoneCallback, NETtempSocketHandleLobbyJoinCheckResult). The host's
' identity is getMultiStats(identityPlayer).identity (the private key stays in the exe); gameDetailsJson is
' netlobby.cpp's to_json(GameDetails) text, connectionsJson a JSON array of to_json(ConnectionInfo); a
' HostJoinOptionValue or HostJoinOptionType is a Long.

' lobbyHostingHandler = netlobby::MakeLobbyHostListingHandler(lobbyServerAddress, eventCallbacks).
Public Sub lobbyHostCreate(ByVal lobbyServerAddress As String)
End Sub

' lobbyHostingHandler.reset() (the handler's destructor).
Public Sub lobbyHostReset()
End Sub

' createGameListing(hostIdentity, gameDetails, joinOptions, connections).
Public Function lobbyHostCreateGameListing(ByVal identityPlayer As Long, ByVal gameDetailsJson As String, ByVal botProt As Long, _
        ByVal proxyIPs As Long, ByVal hostingIPs As Long, ByVal connectionsJson As String) As Boolean
End Function

' updateGameListing(hostIdentity, gameDetails, joinOptions).
Public Sub lobbyHostUpdateGameListing(ByVal identityPlayer As Long, ByVal gameDetailsJson As String, ByVal botProt As Long, _
        ByVal proxyIPs As Long, ByVal hostingIPs As Long)
End Sub

' respondToLobbyConnectionCheckRequest(jsonRequest, playerName, playerIdentity): True with the response, or False
' with errorCode (LobbyConnectionCheckRequestError: 0 GameNotListed, 1 InvalidRequest).
Public Function lobbyHostRespondToLobbyConnectionCheckRequest(ByVal jsonRequest As String, ByVal playerName As String, _
        ByVal identityPlayer As Long, response As String, errorCode As Long) As Boolean
End Function

' checkJoinRequest(joinRequest, [i, uniqueJoinIDCopy](CheckJoinResult result)): the result comes back as
' NetApi's NETtempSocketHandleLobbyJoinCheckResult(tmpSocketIndex, uniqueJoinID, ...).
Public Function lobbyHostCheckJoinRequest(ByVal joinToken As String, ByVal ip As String, ByVal playerName As String, _
        ByVal playerPublicIdentity As String, ByVal connectionMethod As Long, ByVal asSpectator As Boolean, _
        ByVal tmpSocketIndex As Long, ByVal uniqueJoinID As String) As Boolean
End Function

' removeGameListing(bool matchStarted).
Public Sub lobbyHostRemoveGameListing(ByVal matchStarted As Boolean)
End Sub

' run(): queued updates and keep-alives.
Public Sub lobbyHostRun()
End Sub

Public Function lobbyHostIsConnectedToLobby() As Boolean
End Function

' getCurrentGameId(): optional<std::string> (has_value: the result).
Public Function lobbyHostGetCurrentGameId(gameId As String) As Boolean
End Function

' isHostJoinOptionSupportedByLobby(optType): optional<bool> (has_value: the result).
Public Function lobbyHostIsHostJoinOptionSupportedByLobby(ByVal optType As Long, result As Boolean) As Boolean
End Function

' src/modding.cpp getLoadedMods(): the mods' names and their getHash() strings ("" for a zero hash):
' names(0 .. count - 1), hashes(0 .. count - 1).
Public Function getLoadedMods(names() As String, hashes() As String) As Long
End Function

' src/warzoneconfig.cpp war_getLobbyDisableIPv6.
Public Function warGetLobbyDisableIPv6() As Boolean
End Function
