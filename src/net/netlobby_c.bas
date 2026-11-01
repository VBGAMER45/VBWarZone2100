Attribute VB_Name = "netlobby_c"
' lib/netplay/netplay.cpp's lobby hosting handler (std::shared_ptr<netlobby::LobbyServerHostingHandlerProtocol>
' lobbyHostingHandler) and NETcreateLobbyHostingHandler's eventCallbacks, with netlobby.h's GetDefaultLobbyAddress.
' The handler (lib/netplay/netlobby.cpp) lives in the exe (its netlobby_c: it needs urlrequest, version and the
' verify prompts, which only the exe has): these calls reach it through INetGame (netGame.lobbyHost*), and its
' event callbacks come back through NetApi (NETlobbyGameCreatedCallback, NETlobbyGamePublishedCallback,
' NETlobbyGameGoneCallback, NETtempSocketHandleLobbyJoinCheckResult). handlerExists is the shared_ptr being set.
' The host's identity crosses as the player whose getMultiStats(player).identity it is; gameDetails as
' netlobby.cpp's to_json(GameDetails) JSON text (gameDetailsJson), the connections as a JSON array of
' to_json(ConnectionInfo).
Option Explicit

' std::shared_ptr<netlobby::LobbyServerHostingHandlerProtocol> lobbyHostingHandler (it is set)
Private handlerExists As Boolean

' netlobby.h: lobbyAddressEncoded, each byte xor scr_key (0x5A)
Public Function netlobby_GetDefaultLobbyAddress() As String
    Dim lobbyAddressEncoded As Variant, i As Long, s As String
    lobbyAddressEncoded = Array(&H32, &H2E, &H2E, &H2A, &H29, &H60, &H75, &H75, &H2D, &H20, &H36, &H35, &H38, &H38, &H23, &H74, _
        &H2D, &H20, &H68, &H6B, &H6A, &H6A, &H74, &H34, &H3F, &H2E, &H75, &H36, &H35, &H38, &H38, &H23)
    For i = 0 To UBound(lobbyAddressEncoded)
        s = s & ChrW$(lobbyAddressEncoded(i) Xor &H5A)
    Next
    netlobby_GetDefaultLobbyAddress = s
End Function

Public Function lobbyHostingHandler_exists() As Boolean
    lobbyHostingHandler_exists = handlerExists
End Function

' lobbyHostingHandler.reset()
Public Sub lobbyHostingHandler_reset()
    If Not handlerExists Then Exit Sub
    handlerExists = False
    netGame.lobbyHostReset
End Sub

' static std::shared_ptr<netlobby::LobbyServerHostingHandlerProtocol> NETcreateLobbyHostingHandler() (its result is
' assigned to lobbyHostingHandler: the exe's handler is made here)
Public Sub NETcreateLobbyHostingHandler()
    If handlerExists Then ASSERT_FAILURE "lobbyHostingHandler == nullptr", "lib/netplay/netplay.cpp:3301", "NETcreateLobbyHostingHandler", "Already created lobby hosting handler": Exit Sub

    ' eventCallbacks: lobbyGameCreatedCallback, lobbyGamePublishedCallback, lobbyGameGoneCallback (below, through NetApi)
    netGame.lobbyHostCreate NETgetLobbyserverAddress()
    handlerExists = True
End Sub

' eventCallbacks.lobbyGameCreatedCallback = [](std::string gameId)
Public Sub lobbyGameCreatedCallback(ByVal gameId As String)
    debug_ 3305, LOG_NET, "NETcreateLobbyHostingHandler", "Created lobby game ID: %s - awaiting publish", gameId
End Sub

' eventCallbacks.lobbyGamePublishedCallback = [](std::string gameId, std::string motd)
Public Sub lobbyGamePublishedCallback(ByVal gameId As String, ByVal motd As String)
    Dim gameIdMsg As String, msgs(0 To 1) As String
    ' Messaging the UI / sending wzcmd events that the game is actually listed in the lobby
    debug_ 3309, LOG_NET, "NETcreateLobbyHostingHandler", "Using game ID: %s", gameId

    netGame.wzCommandInterfaceOutput va("WZEVENT: lobbyid: %s" & vbLf, gameId)
    ' ActivityManager::instance().updateMultiplayGameData(game, ingame, NETGameIsLocked()) (the activity sinks: not ported)

    If Len(motd) = 0 Then
        motd = "Game is listed in the lobby"
    End If
    gameIdMsg = "Lobby GameId"
    gameIdMsg = gameIdMsg & ": "    ' build this manually to ensure that the format always separates the prefix text from the gameId with ": ", without relying on the translation
    gameIdMsg = gameIdMsg & gameId
    msgs(0) = motd
    msgs(1) = gameIdMsg
    netGame.ShowLobbyStatusMessage msgs, 2

    ' Signal to the ActivityManager that the lobby hosting status is known
    handleLobbyHostingStatusIsKnown True, gameId
End Sub

' eventCallbacks.lobbyGameGoneCallback = [](optional<netlobby::LobbyError> error) (hasError: has_value())
Public Sub lobbyGameGoneCallback(ByVal hasError As Boolean, ByVal errCode As String, ByVal errMessage As String)
    Dim errorMsg As String, b() As Byte, n As Long

    ' Signal to the ActivityManager that the lobby hosting status is known
    handleLobbyHostingStatusIsKnown False, ""

    If hasError Then
        errorMsg = va("%s" & vbLf & "%s: %s", "Game not listed in the lobby.", errCode, errMessage)
        ShowLobbyStatusMessage1 errorMsg

        ' this is horrible but message can have 0x0a and other junk in it
        n = utf8_encode(errMessage, b)
        If n = 0 Then ReDim b(0 To 0)
        netGame.wzCommandInterfaceOutput va("WZEVENT: lobbyerror (%s): %s" & vbLf, errCode, base64Encode(b, n))
    Else
        errorMsg = "Error connecting to the lobby server."
        errorMsg = errorMsg & vbLf
        errorMsg = errorMsg & "Please check your internet connection, or wait a short while and try hosting again."
        errorMsg = errorMsg & vbLf
        errorMsg = errorMsg & "Your game will only be accessible via direct connection (or local LAN)."
        ShowLobbyStatusMessage1 errorMsg

        netGame.wzCommandInterfaceOutput "WZEVENT: lobbyerror (UNKNOWN): No details" & vbLf
    End If
End Sub

' lobbyHostingHandler->createGameListing(getMultiStats(realSelectedPlayer).identity, gameDetails, hostJoinOptions, connections)
Public Function lobbyHostingHandler_createGameListing(ByVal identityPlayer As Long, ByVal connectionsJson As String) As Boolean
    lobbyHostingHandler_createGameListing = netGame.lobbyHostCreateGameListing(identityPlayer, gameDetailsJson(), hostJoinOptions_botProt, _
        hostJoinOptions_proxyIPs, hostJoinOptions_hostingIPs, connectionsJson)
End Function

' lobbyHostingHandler->updateGameListing(getMultiStats(realSelectedPlayer).identity, gameDetails, hostJoinOptions)
Public Sub lobbyHostingHandler_updateGameListing(ByVal identityPlayer As Long)
    netGame.lobbyHostUpdateGameListing identityPlayer, gameDetailsJson(), hostJoinOptions_botProt, hostJoinOptions_proxyIPs, hostJoinOptions_hostingIPs
End Sub

' lobbyHostingHandler->respondToLobbyConnectionCheckRequest(jsonString, gameDetails.host.name, getMultiStats(realSelectedPlayer).identity)
' (True: response has the value; False: errorCode is the LobbyConnectionCheckRequestError, 0 GameNotListed, 1 InvalidRequest)
Public Function lobbyHostingHandler_respondToLobbyConnectionCheckRequest(ByVal jsonString As String, ByVal playerName As String, _
        ByVal identityPlayer As Long, response As String, errorCode As Long) As Boolean
    lobbyHostingHandler_respondToLobbyConnectionCheckRequest = netGame.lobbyHostRespondToLobbyConnectionCheckRequest(jsonString, playerName, identityPlayer, response, errorCode)
End Function

' lobbyHostingHandler->checkJoinRequest(joinRequest, [i, uniqueJoinIDCopy](CheckJoinResult result) {...}) (the result
' comes back through NetApi's NETtempSocketHandleLobbyJoinCheckResult)
Public Function lobbyHostingHandler_checkJoinRequest(ByVal joinToken As String, ByVal ip As String, ByVal playerName As String, _
        ByVal playerPublicIdentity As String, ByVal connectionMethod As Long, ByVal asSpectator As Boolean, ByVal i As Long, ByVal uniqueJoinIDCopy As String) As Boolean
    lobbyHostingHandler_checkJoinRequest = netGame.lobbyHostCheckJoinRequest(joinToken, ip, playerName, playerPublicIdentity, connectionMethod, asSpectator, i, uniqueJoinIDCopy)
End Function

' removeGameListing(bool matchStarted)
Public Sub lobbyHostingHandler_removeGameListing(ByVal matchStarted As Boolean)
    netGame.lobbyHostRemoveGameListing matchStarted
End Sub

' run(): sending queued updates as well as keep-alives (NETrecvNet, every frame)
Public Sub lobbyHostingHandler_run()
    netGame.lobbyHostRun
End Sub

Public Function lobbyHostingHandler_isConnectedToLobby() As Boolean
    lobbyHostingHandler_isConnectedToLobby = netGame.lobbyHostIsConnectedToLobby()
End Function

' getCurrentGameId().value_or(std::string())
Public Function lobbyHostingHandler_getCurrentGameId() As String
    Dim gameId As String
    If netGame.lobbyHostGetCurrentGameId(gameId) Then lobbyHostingHandler_getCurrentGameId = gameId Else lobbyHostingHandler_getCurrentGameId = ""
End Function

' isHostJoinOptionSupportedByLobby(optType): optional<bool> (has_value: the result)
Public Function lobbyHostingHandler_isHostJoinOptionSupportedByLobby(ByVal optType As Long, result As Boolean) As Boolean
    lobbyHostingHandler_isHostJoinOptionSupportedByLobby = netGame.lobbyHostIsHostJoinOptionSupportedByLobby(optType, result)
End Function

' (VB) ShowLobbyStatusMessage({msg}): a vector of one message
Public Sub ShowLobbyStatusMessage1(ByVal msg As String)
    Dim msgs(0 To 0) As String
    msgs(0) = msg
    netGame.ShowLobbyStatusMessage msgs, 1
End Sub

' (VB) the crossing: gameDetails as netlobby.cpp's to_json(GameDetails) writes it (the exe's from_json reads it)
Private Function gameDetailsJson() As String
    Dim j As nlohmann_json, a As nlohmann_json, o As nlohmann_json, i As Long, mods As nlohmann_json
    Set j = json_object(True)
    j.set_item_ref "v", json_value(gameDetails.versionStr, True)
    Set a = json_array(True)
    a.push_back_ref json_value_unsigned(gameDetails.netcodeVer_major, True)
    a.push_back_ref json_value_unsigned(gameDetails.netcodeVer_minor, True)
    j.set_item_ref "netv", a
    j.set_item_ref "name", json_value(gameDetails.name, True)
    Set o = json_object(True)
    o.set_item_ref "name", json_value(gameDetails.host_name, True)
    o.set_item_ref "pk", json_value(gameDetails.host_publicIdentity, True)
    j.set_item_ref "host", o
    Set a = json_array(True)
    a.push_back_ref json_value(gameDetails.map_name, True)
    a.push_back_ref json_value(gameDetails.map_hash, True)
    j.set_item_ref "map", a
    If gameDetails.modsCount <> 0 Then
        Set mods = json_array(True)
        For i = 0 To gameDetails.modsCount - 1
            Set o = json_object(True)
            o.set_item_ref "name", json_value(gameDetails.mods_name(i), True)
            o.set_item_ref "hash", json_value(gameDetails.mods_hash(i), True)
            mods.push_back_ref o
        Next
        j.set_item_ref "mods", mods
    End If
    Set o = json_object(True)
    o.set_item_ref "cur", json_value_unsigned(gameDetails.players_current, True)
    o.set_item_ref "max", json_value_unsigned(gameDetails.players_max, True)
    j.set_item_ref "players", o
    Set o = json_object(True)
    o.set_item_ref "cur", json_value_unsigned(gameDetails.spectators_current, True)
    o.set_item_ref "max", json_value_unsigned(gameDetails.spectators_max, True)
    j.set_item_ref "specs", o
    j.set_item_ref "ally", json_value_unsigned(gameDetails.alliancesMode And &HFF&, True)
    If (gameDetails.blindMode And &HFF&) <> 0 Then
        j.set_item_ref "blind", json_value_unsigned(gameDetails.blindMode And &HFF&, True)    ' (uint8_t blindMode)
    End If
    j.set_item_ref "lim", json_value_unsigned(gameDetails.limits And &HFF&, True)
    If gameDetails.isPrivate Then
        j.set_item_ref "private", json_value(True, True)
    End If
    gameDetailsJson = j.dump(-1, " ", False, error_handler_t_replace)
End Function
