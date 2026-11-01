VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LobbyServerHostingHandlerImpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/netplay/netlobby.cpp: class LobbyServerHostingHandlerImpl : public LobbyServerHostingHandlerProtocol (the
' protocol's one implementation). construct is its constructor (MakeLobbyHostListingHandler), destruct its
' destructor (lobbyHostingHandler.reset()). Its std::weak_ptr captures are serial (netlobby_c's
' LobbyServerHostingHandlerImpl_lock); the lambdas are netlobby_c's LobbyServerHostingHandlerImpl_fn. The
' std::chrono times are steady-clock milliseconds (a default time_point: 0), the intervals seconds. The PATCH
' request's captured ownedPendingUpdate is kept here (inflight*: one update request at a time, as run() only sends
' when no response is awaited), found again by its token.
Option Explicit

Private Enum LobbyListingState
    LobbyListingState_None
    LobbyListingState_PendingCreate
    LobbyListingState_PendingPublish
    LobbyListingState_Listed
End Enum

' struct PendingUpdateGameDetails
Private Type PendingUpdateGameDetails
    hostIdentity As EcKey
    gameDetails As netlobby_GameDetails
    joinOptions As netlobby_HostJoinOptions
End Type

Private serial_ As Long
Private lobbyGameCreatedCallback As WzFunc      ' EventCallbacks callbacks
Private lobbyGamePublishedCallback As WzFunc
Private lobbyGameGoneCallback As WzFunc

Private lobbyServerAddress As String
Private trustedLobbyServerAddress As Boolean
Private state As LobbyListingState

Private currPendingPublishCheckinInterval As Long   ' std::chrono::seconds

Private lastGameDetails As netlobby_GameDetails
Private lastJoinOptions As netlobby_HostJoinOptions
Private publishedConnectionInfo() As netlobby_ConnectionInfo
Private publishedConnectionInfoCount As Long

Private has_pendingUpdate As Boolean            ' optional<PendingUpdateGameDetails> pendingUpdate
Private pendingUpdate As PendingUpdateGameDetails
Private awaitingAsyncUpdateResponse As Boolean
Private lastLobbyUpdateAttempt As Double
Private lastSuccessfulLobbyUpdate As Double
Private has_lobbyUpdateRateLimitWaitTill As Boolean ' optional<std::chrono::steady_clock::time_point> lobbyUpdateRateLimitWaitTill
Private lobbyUpdateRateLimitWaitTill As Double

Private hostGameCtx As netlobby_HostGameContext
Private createMotd As String
Private has_supportedHostJoinOpts As Boolean    ' optional<std::unordered_set<HostJoinOptionType>> supportedHostJoinOpts
Private supportedHostJoinOpts As Long           ' bit (1 << type) each

Private inflightToken As Long                   ' (VB) the PATCH request's ownedPendingUpdate
Private has_inflightUpdate As Boolean
Private inflightUpdate As PendingUpdateGameDetails

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

' (VB) std::chrono::steady_clock::now(), in milliseconds
Private Function steadyNowMs() As Double
    Dim c As Currency, freq As Currency
    QueryPerformanceFrequency freq
    QueryPerformanceCounter c
    steadyNowMs = CDbl(c) * 1000# / CDbl(freq)
End Function

' (VB) std::chrono::duration_cast<std::chrono::seconds>(now - t)
Private Function secondsSince(ByVal now As Double, ByVal t As Double) As Double
    secondsSince = Fix((now - t) / 1000#)
End Function

Public Property Get serial() As Long
    serial = serial_
End Property

' LobbyServerHostingHandlerImpl(const std::string& lobbyServerAddress, EventCallbacks callbacks)
Friend Sub construct(ByVal lobbyServerAddress_ As String, ByVal trusted As Boolean, ByVal createdCallback As WzFunc, _
        ByVal publishedCallback As WzFunc, ByVal goneCallback As WzFunc)
    Set lobbyGameCreatedCallback = createdCallback
    Set lobbyGamePublishedCallback = publishedCallback
    Set lobbyGameGoneCallback = goneCallback
    lobbyServerAddress = lobbyServerAddress_
    trustedLobbyServerAddress = trusted
    state = LobbyListingState_None
    currPendingPublishCheckinInterval = 1       ' SERVER_UPDATE_PENDING_PUBLISH_CHECKIN_MIN_INTERVAL
    netlobby_HostJoinOptions_init lastJoinOptions
    serial_ = LobbyServerHostingHandlerImpl_register(Me)
End Sub

' LobbyServerHostingHandlerImpl::~LobbyServerHostingHandlerImpl()
Friend Sub destruct()
    If netlobby_HostGameContext_isValid(hostGameCtx) Then
        debug_ 1302, LOG_INFO, "~LobbyServerHostingHandlerImpl", "Removing game listing due to LobbyServerHostingHandler destruction (unexpected)"
        transitionToGameListingGone False, False, "", ""
    End If
    LobbyServerHostingHandlerImpl_unregister serial_
End Sub

' UrlRequestResponseHandler buildHostingAPIResponseHandler(std::unordered_set<long> acceptedResponseCodes, handlerFunc)
' (acceptedResponseCodes as ",200,201,": the set's members between commas)
Private Function buildHostingAPIResponseHandler(ByVal acceptedResponseCodes As String, ByVal handlerFunc As WzFunc) As WzFunc
    Set buildHostingAPIResponseHandler = WzFunc_new(FN_netlobby_hostingAPIResponse, serial_, acceptedResponseCodes, handlerFunc)
End Function

' Builds a request failure handler that treats a request failure as fatal for the lobby game listing
Private Function buildHostingAPIStandardFailureHandler() As WzFunc
    Set buildHostingAPIStandardFailureHandler = WzFunc_new(FN_netlobby_hostingAPIStdFailure, serial_)
End Function

' (VB) try { err = jsonData.get<LobbyError>(); } catch (...) { }
Private Function tryLobbyError(jsonData As nlohmann_json, errCode As String, errMessage As String) As Boolean
    Dim code As String, message As String, it As nlohmann_json
    On Error GoTo failed
    If jsonData Is Nothing Then GoTo failed
    code = jsonData.at("code").get_string()
    Set it = jsonData.find("message")
    If Not it Is Nothing Then message = it.get_string()
    errCode = code
    errMessage = message
    tryLobbyError = True
    Exit Function
failed:
    tryLobbyError = False
End Function

Friend Sub logAPIFailure(ByVal url As String, ByVal httpStatusCode As Long, jsonData As nlohmann_json)
    Dim errCode As String, errMessage As String
    ' Log
    If tryLobbyError(jsonData, errCode, errMessage) Then
        debug_ 1398, LOG_LOBBY, "logAPIFailure", "Lobby returned error: (%ld) [%s] %s", httpStatusCode, errCode, errMessage
    Else
        debug_ 1402, LOG_LOBBY, "logAPIFailure", "Lobby error: (%ld)", httpStatusCode
    End If
End Sub

Friend Sub handleAPIFailure(ByVal url As String, ByVal httpStatusCode As Long, jsonData As nlohmann_json)
    Dim errCode As String, errMessage As String, has_err As Boolean
    has_err = tryLobbyError(jsonData, errCode, errMessage)

    ' Log
    If has_err Then
        debug_ 1420, LOG_LOBBY, "handleAPIFailure", "Lobby returned error: (%ld) [%s] %s", httpStatusCode, errCode, errMessage
    Else
        debug_ 1424, LOG_LOBBY, "handleAPIFailure", "Lobby error: (%ld)", httpStatusCode
    End If

    transitionToGameListingGone False, has_err, errCode, errMessage
End Sub

Friend Sub handleAPIResponseHandlingFailure(ByVal url As String, ByVal httpStatusCode As Long, ByVal what As String)
    ' The API response handler function threw an exception
    ' Very likely due to trying to parse the JSON response

    ' Log
    debug_ 1436, LOG_LOBBY, "handleAPIResponseHandlingFailure", "Lobby response handling failure: (%ld) %s", httpStatusCode, what

    ' Treat this as a critical failure, and disconnected lobby listing
    transitionToGameListingGone False, False, "", ""
End Sub

' bool createGameListing(const EcKey& hostIdentity, const GameDetails& gameDetails, const HostJoinOptions& joinOptions,
'   const std::vector<ConnectionInfo>& connections)
Friend Function createGameListing(hostIdentity As EcKey, gameDetails As netlobby_GameDetails, joinOptions As netlobby_HostJoinOptions, _
        connections() As netlobby_ConnectionInfo, ByVal connectionsCount As Long) As Boolean
    Dim request As URLDataRequest, hostCreateGameInfo As nlohmann_json, attestation As nlohmann_json
    If Not (state = LobbyListingState_None) Then ASSERT_FAILURE "state == LobbyListingState::None", "lib/netplay/netlobby.cpp:1444", "createGameListing", "create called when lobby listing state is: %d", state: Exit Function
    If Not (connectionsCount <> 0) Then ASSERT_FAILURE "!connections.empty()", "lib/netplay/netlobby.cpp:1445", "createGameListing", "create called with no connections?": Exit Function

    initializeGameListingCreationData gameDetails, joinOptions, connections, connectionsCount

    netlobby_buildBaseHostingURLDataRequest request
    request.url = netlobby_buildLobbyRequestBaseUrl(lobbyServerAddress) & "/api/v1/game"
    request.protocol = InternetProtocol_IPv4

    Set hostCreateGameInfo = netlobby_GameDetails_to_json(gameDetails)
    If Not (hostCreateGameInfo.is_object()) Then ASSERT_FAILURE "hostCreateGameInfo.is_object()", "lib/netplay/netlobby.cpp:1454", "createGameListing", "Not an object?": Exit Function
    hostCreateGameInfo.set_item_ref "joinopts", netlobby_HostJoinOptions_to_json(joinOptions)

    If trustedLobbyServerAddress Then
        Set attestation = netlobby_makeAttestationJson(gameDetails.host.name, hostIdentity, lobbyServerAddress, False, "")
        hostCreateGameInfo.set_item_ref "attest", attestation
    End If

    URLDataRequest_setPost request, hostCreateGameInfo.dump(-1, " ", False, error_handler_t_replace)

    request.connectTimeoutMs = 10000

    request.autoRetryStrategy = URLRequestAutoRetryStrategy_RetryAfter(5000, 2, Array(HttpRequestConnectionFailure, HttpStatusTooManyRequests))

    Set request.onResponse = buildHostingAPIResponseHandler(",200,201,202,", WzFunc_new(FN_netlobby_createGameHandler, serial_, lobbyServerAddress))

    Set request.onFailure = buildHostingAPIStandardFailureHandler()

    If urlRequestData(request) = 0 Then
        createGameListing = False
        Exit Function
    End If

    state = LobbyListingState_PendingCreate
    createGameListing = True
End Function

' LobbyConnectionCheckResult respondToLobbyConnectionCheckRequest(const std::string& jsonRequest, const WzString& playerName,
'   const EcKey& playerIdentity): True with the response, or False with the error
Friend Function respondToLobbyConnectionCheckRequest(ByVal jsonRequest As String, ByVal playerName As String, playerIdentity As EcKey, _
        responseString As String, error_ As netlobby_LobbyConnectionCheckRequestError) As Boolean
    Dim lobbyConnectToken As String, challenge As String, what As String, challengeLength As Long, simpleResponse As nlohmann_json, gid As String
    If state = LobbyListingState_None Or Not netlobby_HostGameContext_isValid(hostGameCtx) Then
        ' Game not (yet?) listed
        error_ = netlobby_LobbyConnectionCheckRequestError_GameNotListed
        Exit Function
    End If

    ' JSON decode LobbyConnectCheckRequest from jsonString
    If Not netlobby_parseLobbyConnectCheckRequest(jsonRequest, lobbyConnectToken, challenge, what) Then
        debug_ 1513, LOG_LOBBY, "respondToLobbyConnectionCheckRequest", "Failed to parse lobby connection request: %s", what
        error_ = netlobby_LobbyConnectionCheckRequestError_InvalidRequest
        Exit Function
    End If

    ' Sanity-checks
    If lobbyConnectToken <> hostGameCtx.lobbyConnectToken Then
        debug_ 1520, LOG_LOBBY, "respondToLobbyConnectionCheckRequest", "Invalid lobbyConnectToken"
        error_ = netlobby_LobbyConnectionCheckRequestError_InvalidRequest
        Exit Function
    End If
    challengeLength = utf8Length(challenge)
    If challengeLength < 32 Or challengeLength > 88 Then
        ' Insufficient challenge length
        debug_ 1527, LOG_LOBBY, "respondToLobbyConnectionCheckRequest", "Invalid challenge length: %zu", challengeLength
        error_ = netlobby_LobbyConnectionCheckRequestError_InvalidRequest
        Exit Function
    End If

    If trustedLobbyServerAddress Then
        ' Construct LobbyConnectCheckResponse
        responseString = netlobby_makeLobbyConnectCheckResponseJson(hostGameCtx.gameId, playerName, playerIdentity, challenge).dump(-1, " ", False, error_handler_t_replace)
    Else
        ' Send just the gameId
        Set simpleResponse = json_object(True)
        If Not getCurrentGameId(gid) Then gid = ""
        simpleResponse.set_item_ref "gid", json_value(gid, True)

        responseString = simpleResponse.dump(-1, " ", False, error_handler_t_replace)
    End If

    respondToLobbyConnectionCheckRequest = True
End Function

' (VB) std::string::size(): the UTF-8 length
Private Function utf8Length(ByVal s As String) As Long
    Dim b() As Byte
    utf8Length = utf8_encode(s, b)
End Function

' Queues a game listing update
Friend Sub updateGameListing(hostIdentity As EcKey, gameDetails As netlobby_GameDetails, joinOptions As netlobby_HostJoinOptions)
    If state = LobbyListingState_None Then
        Exit Sub
    End If

    has_pendingUpdate = True
    pendingUpdate.hostIdentity = hostIdentity
    pendingUpdate.gameDetails = gameDetails
    pendingUpdate.joinOptions = joinOptions
End Sub

Private Sub sendListingUpdateImpl()
    Dim request As URLDataRequest, has_ownedPendingUpdate As Boolean, ownedPendingUpdate As PendingUpdateGameDetails, fresh As PendingUpdateGameDetails
    Dim hostCreateGameInfo As nlohmann_json, hostIdentityChanged As Boolean, attestation As nlohmann_json

    netlobby_buildBaseHostingURLDataRequest request
    netlobby_addHostAuthorization request, hostGameCtx
    request.url = netlobby_buildLobbyRequestBaseUrl(lobbyServerAddress) & "/api/v1/game/" & hostGameCtx.gameId
    request.protocol = InternetProtocol_IPv4
    request.connectTimeoutMs = 30 * 1000&       ' SERVER_UPDATE_MAX_INTERVAL

    has_ownedPendingUpdate = has_pendingUpdate
    ownedPendingUpdate = pendingUpdate
    has_pendingUpdate = False
    pendingUpdate = fresh

    If has_ownedPendingUpdate Then
        Set hostCreateGameInfo = netlobby_serializeGameDetailsChangedFields(ownedPendingUpdate.gameDetails, lastGameDetails)
        If Not (hostCreateGameInfo.is_object()) Then ASSERT_FAILURE "hostCreateGameInfo.is_object()", "lib/netplay/netlobby.cpp:1577", "sendListingUpdateImpl", "Not an object?": Exit Sub

        If ownedPendingUpdate.joinOptions.botProt <> lastJoinOptions.botProt Or ownedPendingUpdate.joinOptions.proxyIPs <> lastJoinOptions.proxyIPs _
           Or ownedPendingUpdate.joinOptions.hostingIPs <> lastJoinOptions.hostingIPs Then
            hostCreateGameInfo.set_item_ref "joinopts", netlobby_HostJoinOptions_to_json(ownedPendingUpdate.joinOptions)
        End If

        hostIdentityChanged = (lastGameDetails.host.publicIdentity <> ownedPendingUpdate.gameDetails.host.publicIdentity)
        If hostIdentityChanged And trustedLobbyServerAddress Then
            Set attestation = netlobby_makeAttestationJson(ownedPendingUpdate.gameDetails.host.name, ownedPendingUpdate.hostIdentity, lobbyServerAddress, False, "")
            hostCreateGameInfo.set_item_ref "attest", attestation
        End If

        URLDataRequest_setPatch request, hostCreateGameInfo.dump(-1, " ", False, error_handler_t_replace)
    Else
        URLDataRequest_setPatch request, "{}"
    End If

    ' Don't want the default autoRetryStrategy here because of the asynchronous nature of the requests
    ' Specifically: There is no guarantee that a patch request response is returned before another pendingUpdate is queued by the main thread. So we must coordinate retry strategy with the main thread.

    inflightToken = inflightToken + 1
    has_inflightUpdate = has_ownedPendingUpdate
    inflightUpdate = ownedPendingUpdate

    Set request.onResponse = buildHostingAPIResponseHandler(",200,201,202,429,", WzFunc_new(FN_netlobby_updateGameHandler, serial_, lobbyServerAddress, inflightToken))

    Set request.onFailure = buildHostingAPIStandardFailureHandler()

    If urlRequestData(request) = 0 Then
        ' Failed to dispatch request
        debug_ 1648, LOG_INFO, "sendListingUpdateImpl", "Failed to dispatch url request"
    End If

    awaitingAsyncUpdateResponse = True
    lastLobbyUpdateAttempt = steadyNowMs()
End Sub

' sendListingUpdateImpl's handlerFunc, after strongSelf: (handle only accepted response codes (called on main thread))
Friend Sub listingUpdateResponse(ByVal token As Long, ByVal httpStatusCode As Long, jsonData As nlohmann_json)
    Dim now As Double, has_ownedPendingUpdate As Boolean, ownedPendingUpdate As PendingUpdateGameDetails, hostGameStatus As netlobby_HostGameStatus
    If token = inflightToken Then
        has_ownedPendingUpdate = has_inflightUpdate
        ownedPendingUpdate = inflightUpdate
    End If

    awaitingAsyncUpdateResponse = False
    now = steadyNowMs()

    If httpStatusCode = 429 Then        ' 429 Too Many Requests (RFC 6585)
        If Not has_pendingUpdate And has_ownedPendingUpdate Then
            ' No more recent pending update game info
            ' Requeue the *same* pendingUpdate
            has_pendingUpdate = True
            pendingUpdate = ownedPendingUpdate
        End If

        ' Throttle retry
        has_lobbyUpdateRateLimitWaitTill = True
        lobbyUpdateRateLimitWaitTill = now + 5000#
        Exit Sub
    End If

    lastSuccessfulLobbyUpdate = now
    If has_ownedPendingUpdate Then
        lastGameDetails = ownedPendingUpdate.gameDetails
        lastJoinOptions = ownedPendingUpdate.joinOptions
    End If

    ' Get HostGameStatus from jsonData
    netlobby_HostGameStatus_from_json jsonData, hostGameStatus
    handleHostGameStatus hostGameStatus
End Sub

' Checks a lobby-originated join request (resultHandler gets (has_value, pass, errCode, errMessage))
Friend Function checkJoinRequest(joinRequest As netlobby_JoinRequest, ByVal resultHandler As WzFunc) As Boolean
    Dim request As URLDataRequest, joinRequestJson As nlohmann_json
    If Not (Not resultHandler Is Nothing) Then ASSERT_FAILURE "resultHandler != nullptr", "lib/netplay/netlobby.cpp:1658", "checkJoinRequest", "No result handler": Exit Function

    If state = LobbyListingState_None Or Not netlobby_HostGameContext_isValid(hostGameCtx) Then
        ' Game not (yet?) listed, so can't check join
        checkJoinRequest = False
        Exit Function
    End If

    netlobby_buildBaseHostingURLDataRequest request
    netlobby_addHostAuthorization request, hostGameCtx
    request.url = netlobby_buildLobbyRequestBaseUrl(lobbyServerAddress) & "/api/v1/game/" & hostGameCtx.gameId & "/check-join-request"
    request.protocol = InternetProtocol_IPv4
    request.connectTimeoutMs = 10000

    request.autoRetryStrategy = URLRequestAutoRetryStrategy_RetryAfter(3000, 1, Array(HttpStatusTooManyRequests))

    Set joinRequestJson = netlobby_JoinRequest_to_json(joinRequest)
    If Not (joinRequestJson.is_object()) Then ASSERT_FAILURE "joinRequestJson.is_object()", "lib/netplay/netlobby.cpp:1675", "checkJoinRequest", "Not an object?": Exit Function
    URLDataRequest_setPost request, joinRequestJson.dump(-1, " ", False, error_handler_t_replace)

    Set request.onResponse = WzFunc_new(FN_netlobby_checkJoinResponse, resultHandler)

    Set request.onFailure = WzFunc_new(FN_netlobby_checkJoinFailure, resultHandler)

    ' Dispatch query
    If urlRequestData(request) = 0 Then
        checkJoinRequest = False
        Exit Function
    End If

    checkJoinRequest = True
End Function

' Removes the current game listing from the lobby
Friend Sub removeGameListing(ByVal matchStarted As Boolean)
    ' skip event callback, since this was deliberately removed from the lobby
    transitionToGameListingGone matchStarted, False, "", "", True
End Sub

' void transitionToGameListingGone(bool matchStarted, optional<LobbyError> error, bool skipEventCallback = false)
Private Sub transitionToGameListingGone(ByVal matchStarted As Boolean, ByVal has_error As Boolean, ByVal errCode As String, ByVal errMessage As String, _
        Optional ByVal skipEventCallback As Boolean = False)
    If netlobby_HostGameContext_isValid(hostGameCtx) Then
        netlobby_LobbyDeleteGameImpl lobbyServerAddress, hostGameCtx, matchStarted
    Else
        debug_ 1779, LOG_LOBBY, "transitionToGameListingGone", "No known game listing to remove - skipping removal"
    End If

    eventGameListingGone has_error, errCode, errMessage, skipEventCallback
End Sub

' Handles sending queued updates as well as keep-alives - should be called regularly by the main loop (at least, while in lobby)
Friend Sub run()
    Dim now As Double, secondsSinceLastSuccessfulLobbyUpdate As Double, secondsSinceLastLobbyUpdateAttempt As Double
    Dim successfulLobbyUpdateInterval As Long
    If state = LobbyListingState_None Or Not netlobby_HostGameContext_isValid(hostGameCtx) Then
        Exit Sub        ' nothing to do until / unless state is pending or listed
    End If

    If Not awaitingAsyncUpdateResponse Then
        now = steadyNowMs()

        secondsSinceLastSuccessfulLobbyUpdate = secondsSince(now, lastSuccessfulLobbyUpdate)
        secondsSinceLastLobbyUpdateAttempt = secondsSince(now, lastLobbyUpdateAttempt)

        If secondsSinceLastSuccessfulLobbyUpdate >= 120 Then    ' SERVER_UPDATE_NO_SUCCESS_ABANDON_INTERVAL
            ' Despite trying, we haven't managed to successfully update the server in a while
            ' Treat this as a failure and stop trying
            transitionToGameListingGone False, False, "", ""
            Exit Sub
        End If

        If has_lobbyUpdateRateLimitWaitTill Then
            If now < lobbyUpdateRateLimitWaitTill Then
                ' Wait until lobbyUpdateRateLimitWaitTill passes to try again
                Exit Sub
            Else
                ' clear lobbyUpdateRateLimitWaitTill and proceed as normal
                has_lobbyUpdateRateLimitWaitTill = False
            End If
        End If

        If state = LobbyListingState_PendingPublish Then
            successfulLobbyUpdateInterval = currPendingPublishCheckinInterval
        Else
            successfulLobbyUpdateInterval = 30                  ' SERVER_UPDATE_MAX_INTERVAL
        End If

        If (secondsSinceLastSuccessfulLobbyUpdate >= successfulLobbyUpdateInterval) _
           Or (has_pendingUpdate And secondsSinceLastLobbyUpdateAttempt >= 8) Then ' SERVER_UPDATE_MIN_INTERVAL
            If state = LobbyListingState_PendingPublish Then
                ' increasing delay for initial PendingPublish check-ins (i.e. waiting for lobby to return published status)
                currPendingPublishCheckinInterval = currPendingPublishCheckinInterval + 1
                If currPendingPublishCheckinInterval > 5 Then currPendingPublishCheckinInterval = 5   ' SERVER_UPDATE_PENDING_PUBLISH_CHECKIN_MAX_INTERVAL
            End If
            sendListingUpdateImpl
        End If
    End If
End Sub

Friend Function isConnectedToLobby() As Boolean
    isConnectedToLobby = (state <> LobbyListingState_None)
End Function

' optional<std::string> getCurrentGameId() const (has_value: the result)
Friend Function getCurrentGameId(gameId As String) As Boolean
    If Not netlobby_HostGameContext_isValid(hostGameCtx) Then
        getCurrentGameId = False
        Exit Function
    End If
    gameId = hostGameCtx.gameId
    getCurrentGameId = True
End Function

' optional<bool> isHostJoinOptionSupportedByLobby(HostJoinOptionType optType) const (has_value: the result)
Friend Function isHostJoinOptionSupportedByLobby(ByVal optType As netlobby_HostJoinOptionType, result As Boolean) As Boolean
    If Not has_supportedHostJoinOpts Then
        ' The server hasn't told us yet - don't know if supported
        isHostJoinOptionSupportedByLobby = False
        Exit Function
    End If

    result = (supportedHostJoinOpts And CLng(2 ^ optType)) <> 0
    isHostJoinOptionSupportedByLobby = True
End Function

Private Sub initializeGameListingCreationData(gameDetails As netlobby_GameDetails, joinOptions As netlobby_HostJoinOptions, _
        connections() As netlobby_ConnectionInfo, ByVal connectionsCount As Long)
    lastGameDetails = gameDetails
    lastJoinOptions = joinOptions

    publishedConnectionInfo = connections
    publishedConnectionInfoCount = connectionsCount
End Sub

Private Sub addConnectionsToGame()
    Dim k As Long, conn As netlobby_ConnectionInfo, protocol As InternetProtocol, request As URLDataRequest
    Dim publishConnRequestInfoJson As nlohmann_json, ver As netlobby_IPVersion
    For k = 0 To publishedConnectionInfoCount - 1
        conn = publishedConnectionInfo(k)
        If conn.typ.has_ipVersion Then ver = conn.typ.ipVersion Else ver = netlobby_IPVersion_IPv4
        protocol = netlobby_ipVersionToInternetProtocol(ver)

        netlobby_buildBaseHostingURLDataRequest request
        netlobby_addHostAuthorization request, hostGameCtx
        request.url = netlobby_buildLobbyRequestBaseUrl(lobbyServerAddress) & "/api/v1/game/" & hostGameCtx.gameId & "/connection"
        request.protocol = protocol

        Set publishConnRequestInfoJson = netlobby_ConnectionInfo_to_json(conn)
        If Not (publishConnRequestInfoJson.is_object()) Then ASSERT_FAILURE "publishConnRequestInfoJson.is_object()", "lib/netplay/netlobby.cpp:1893", "addConnectionsToGame", "Not an object?": Exit Sub

        URLDataRequest_setPost request, publishConnRequestInfoJson.dump(-1, " ", False, error_handler_t_replace)

        request.connectTimeoutMs = 15000

        request.autoRetryStrategy = URLRequestAutoRetryStrategy_RetryAfter(3000, 2, Array(HttpRequestConnectionFailure, HttpStatusTooManyRequests))

        ' Response codes:
        ' - 409: May be returned on conflict
        ' - 429: May be returned if autoRetryStrategy doesn't succeed
        Set request.onResponse = buildHostingAPIResponseHandler(",200,201,202,406,409,429,", WzFunc_new(FN_netlobby_addConnectionHandler, serial_))

        Set request.onFailure = WzFunc_new(FN_netlobby_addConnectionFailure, protocol)

        urlRequestData request
    Next
End Sub

Friend Sub handleHostGameStatus(statusInfo As netlobby_HostGameStatus)
    Dim allLobbyReceivedConnectionsAreInErrorState As Boolean, i As Long, translatedHostNotAccessibleString As String
    If statusInfo.gameId <> hostGameCtx.gameId Then
        debug_ 1946, LOG_LOBBY, "handleHostGameStatus", "Old game status reply for old game - ignoring"
        Exit Sub
    End If

    If statusInfo.published Then
        If state = LobbyListingState_PendingPublish Then
            ' transition to published state
            eventGamePublished
        End If
    Else
        ' Check for a heads-up that this game listing will never be published
        ' i.e. if no connections were able to be validated by the lobby server

        allLobbyReceivedConnectionsAreInErrorState = True
        For i = 0 To statusInfo.connectionStatusInfoCount - 1
            If statusInfo.connectionStatusInfo(i).status <> netlobby_ConnectionStatus_Error Then allLobbyReceivedConnectionsAreInErrorState = False
        Next

        ' 1. Check if all connections are Error state
        ' 2. Check that lobby server has received all connections
        If allLobbyReceivedConnectionsAreInErrorState And statusInfo.connectionStatusInfoCount = publishedConnectionInfoCount Then
            ' Every connection we tried to add to the game couldn't be verified by the lobby :(
            ' Host appears to be inaccessible from the outside world
            ' Immediately remove game and transition to state gone (with synthesized errCodeHostNotAccessible)

            translatedHostNotAccessibleString = "Error hosting game in lobby server"
            translatedHostNotAccessibleString = translatedHostNotAccessibleString & vbLf
            translatedHostNotAccessibleString = translatedHostNotAccessibleString & va("Make sure port %d can receive incoming connections.", NETgetGameserverPort())
            translatedHostNotAccessibleString = translatedHostNotAccessibleString & vbLf
            translatedHostNotAccessibleString = translatedHostNotAccessibleString & "If you're using a router, configure it to enable UPnP/NAT-PMP/PCP or to forward the port to your system."

            transitionToGameListingGone False, True, "HOST_NOT_ACCESSIBLE", translatedHostNotAccessibleString
        End If
    End If
End Sub

Friend Sub eventGameCreated(resp As netlobby_WrappedCreateGameResponse)
    If state <> LobbyListingState_PendingCreate Then
        ASSERT_FAILURE "state == LobbyListingState::PendingCreate", "lib/netplay/netlobby.cpp:1990", "eventGameCreated", "Unexpected state (%d)", state
        netlobby_LobbyDeleteGameImpl lobbyServerAddress, resp.ctx, False
        Exit Sub
    End If

    hostGameCtx = resp.ctx
    state = LobbyListingState_PendingPublish

    createMotd = resp.motd      ' store, but don't provide to callbacks until game is published
    has_supportedHostJoinOpts = True
    supportedHostJoinOpts = resp.supportedHostJoinOpts

    lastSuccessfulLobbyUpdate = steadyNowMs()
    lastLobbyUpdateAttempt = steadyNowMs()

    ' Attempt to add connections to game
    addConnectionsToGame

    ' Trigger event callback
    If Not lobbyGameCreatedCallback Is Nothing Then
        WzFunc_call lobbyGameCreatedCallback, hostGameCtx.gameId
    End If
End Sub

Private Sub eventGamePublished()
    If state = LobbyListingState_Listed Then
        Exit Sub
    End If

    state = LobbyListingState_Listed

    ' Trigger event callback
    If Not lobbyGamePublishedCallback Is Nothing Then
        WzFunc_call lobbyGamePublishedCallback, hostGameCtx.gameId, createMotd
    End If
End Sub

Private Sub eventGameListingGone(ByVal has_error As Boolean, ByVal errCode As String, ByVal errMessage As String, ByVal skipEventCallback As Boolean)
    Dim freshDetails As netlobby_GameDetails, freshCtx As netlobby_HostGameContext, freshUpdate As PendingUpdateGameDetails
    If state = LobbyListingState_None Then
        Exit Sub
    End If

    state = LobbyListingState_None

    lastGameDetails = freshDetails
    netlobby_HostJoinOptions_init lastJoinOptions

    has_pendingUpdate = False
    pendingUpdate = freshUpdate
    awaitingAsyncUpdateResponse = False
    lastLobbyUpdateAttempt = 0
    lastSuccessfulLobbyUpdate = 0

    hostGameCtx = freshCtx

    If Not skipEventCallback Then
        ' Trigger lobbyGameGoneCallback event callback
        If Not lobbyGameGoneCallback Is Nothing Then
            WzFunc_call lobbyGameGoneCallback, has_error, errCode, errMessage
        End If
    End If
End Sub
