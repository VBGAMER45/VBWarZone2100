Attribute VB_Name = "netlobby_c"
' lib/netplay/netlobby.cpp, netlobby.h, whole (cluster J part 2): the lobby server's client. It lives in the exe
' (it needs urlrequest, version and lib/widget's verify prompts, which only the exe has); WzNet's netlobby_c
' forwards netplay's calls of the hosting handler to it through INetGame (NetGame.cls: lobbyHost*), and the
' handler's event callbacks (netplay.cpp's lambdas) go back through NetApi (NETlobbyGame*Callback,
' NETtempSocketHandleLobbyJoinCheckResult). GetDefaultLobbyAddress is NetApi's (netlobby_GetDefaultLobbyAddress).
' * The requests are urlrequest_c's (WinHttpRequest, polled each frame); their handlers run on the main thread
'   and still queue their work with wzAsyncExecOnMainThread, as 4.7.0's do from the transfer thread; the
'   mutex of PendingJoinRequestData is not needed (one thread).
' * LobbyServerHostingHandlerImpl is the class of that name; its std::weak_ptr captures are its serial
'   (LobbyServerHostingHandlerImpl_lock: Nothing once the handler is destroyed, its destruct). The pending
'   update a PATCH request captures (ownedPendingUpdate) is kept by the handler (one request at a time:
'   awaitingAsyncUpdateResponse).
' * Results handed to a WzFunc (ListResult, JoinResult) are handles into this module's stores, taken once
'   (UDTs do not fit a Variant). A std::shared_ptr<LobbyErrorResolutionData> is the response's nlohmann_json
'   (LobbyErrorResolutionJsonData::jsonData: its only kind). LobbyError's resolvableLobbyErrorJsonData is the
'   error's "data" nlohmann_json (Nothing = nullptr).
' * parse_url's RE2 pattern ^(?:([^:/?#]+)://)?(?:[^/?#@]+@)?([^/?#:]+) is matched by hand (parse_url).
' * The JSON conversions (to_json / from_json) are <Type>_to_json / <Type>_from_json; nlohmann_json raises
'   (Err) where json.hpp throws, and the try / catch blocks are procedures with an error handler.
Option Explicit

Private Const SERVER_UPDATE_PENDING_PUBLISH_CHECKIN_MIN_INTERVAL As Long = 1    ' seconds
Private Const SERVER_UPDATE_PENDING_PUBLISH_CHECKIN_MAX_INTERVAL As Long = 5
Private Const SERVER_UPDATE_MIN_INTERVAL As Long = 8
Private Const SERVER_UPDATE_MAX_INTERVAL As Long = 30
Private Const SERVER_UPDATE_NO_SUCCESS_ABANDON_INTERVAL As Long = 120

Private Const errCodeHostNotAccessible As String = "HOST_NOT_ACCESSIBLE"
Private Const errCodeHostDisallowedJoinSource As String = "HOST_DISALLOWED_JOIN_SOURCE"
Private Const errCodeGameNotFound As String = "GAME_NOT_FOUND"

'==========================================================================================================
' netlobby.h

' enum class IPVersion : uint8_t
Public Enum netlobby_IPVersion
    netlobby_IPVersion_IPv4 = 1
    netlobby_IPVersion_IPv6
End Enum

' enum class ConnectionMethod
Public Enum netlobby_ConnectionMethod
    netlobby_ConnectionMethod_TCP_DIRECT
    netlobby_ConnectionMethod_GNS_DIRECT
End Enum

' enum class StructureLimits : uint8_t
Public Const netlobby_StructureLimits_NO_TANKS As Long = 1
Public Const netlobby_StructureLimits_NO_BORGS As Long = 2
Public Const netlobby_StructureLimits_NO_VTOLS As Long = 4

' enum class HostJoinOptionValue
Public Enum netlobby_HostJoinOptionValue
    netlobby_HostJoinOptionValue_BlockAll
    netlobby_HostJoinOptionValue_BlockSuspicious
    netlobby_HostJoinOptionValue_Allow
End Enum

' enum class HostJoinOptionType
Public Enum netlobby_HostJoinOptionType
    netlobby_HostJoinOptionType_BotProt
    netlobby_HostJoinOptionType_ProxyIPs
    netlobby_HostJoinOptionType_HostingIPs
End Enum

' enum class ConnectionStatus : uint8_t
Public Enum netlobby_ConnectionStatus
    netlobby_ConnectionStatus_Submitted = 0
    netlobby_ConnectionStatus_Validated
    netlobby_ConnectionStatus_Error
End Enum

' enum class RequestJoinDetailsResults::ConnectivityFailureType
Public Enum netlobby_ConnectivityFailureType
    netlobby_ConnectivityFailureType_ERR_COULD_NOT_CONNECT
    netlobby_ConnectivityFailureType_ERR_TIMEOUT
    netlobby_ConnectivityFailureType_ERR_OTHER
End Enum

' enum class LobbyServerHostingHandlerProtocol::LobbyConnectionCheckRequestError
Public Enum netlobby_LobbyConnectionCheckRequestError
    netlobby_LobbyConnectionCheckRequestError_GameNotListed
    netlobby_LobbyConnectionCheckRequestError_InvalidRequest
End Enum

' (netlobby.cpp) enum class ImageType
Public Enum ImageType
    ImageType_PNG
    ImageType_JPEG
End Enum

Public Type netlobby_NetcodeVer
    major As Long                               ' uint32_t
    minor As Long
End Type

Public Type netlobby_HostDetails
    name As String                              ' WzString
    publicIdentity As String
End Type

Public Type netlobby_MapDetails
    name As String                              ' WzString
    hash As String
End Type

Public Type netlobby_ModDetails
    name As String                              ' WzString
    hash As String                              ' WzString
End Type

Public Type netlobby_PlayerCounts
    current As Long                             ' uint32_t
    max As Long                                 ' uint32_t
End Type

' struct GameDetails
Public Type netlobby_GameDetails
    versionStr As String
    netcodeVer As netlobby_NetcodeVer
    name As String
    host As netlobby_HostDetails
    map As netlobby_MapDetails
    mods() As netlobby_ModDetails               ' std::vector<ModDetails>: mods(0 .. modsCount - 1)
    modsCount As Long
    players As netlobby_PlayerCounts
    spectators As netlobby_PlayerCounts
    alliancesMode As Long                       ' uint8_t
    blindMode As Long                           ' uint8_t
    limits As Long                              ' uint8_t
    isPrivate As Boolean
End Type

' struct HostJoinOptions
Public Type netlobby_HostJoinOptions
    botProt As netlobby_HostJoinOptionValue     ' = BlockAll
    proxyIPs As netlobby_HostJoinOptionValue    ' = BlockSuspicious
    hostingIPs As netlobby_HostJoinOptionValue  ' = BlockAll
End Type

' struct ConnectionType
Public Type netlobby_ConnectionType
    method As netlobby_ConnectionMethod
    has_ipVersion As Boolean                    ' optional<IPVersion> ipVersion
    ipVersion As netlobby_IPVersion
End Type

' struct ConnectionInfo
Public Type netlobby_ConnectionInfo
    typ As netlobby_ConnectionType              ' type
    host As String
    port As Long                                ' uint16_t
End Type

' struct HostConnectionStatusInfo
Public Type netlobby_HostConnectionStatusInfo
    typ As netlobby_ConnectionType
    status As netlobby_ConnectionStatus
End Type

' struct HostGameStatus
Public Type netlobby_HostGameStatus
    gameId As String
    published As Boolean
    connectionStatusInfo() As netlobby_HostConnectionStatusInfo   ' (0 .. connectionStatusInfoCount - 1)
    connectionStatusInfoCount As Long
End Type

' struct JoinRequest
Public Type netlobby_JoinRequest
    joinToken As String
    ip As String
    playerName As String
    playerPublicIdentity As String
    connectionMethod As netlobby_ConnectionMethod
    asSpectator As Boolean
End Type

' struct CheckJoinRequestResult
Public Type netlobby_CheckJoinRequestResult
    pass As Boolean
    error_ As String
End Type

' struct LobbyError
Public Type netlobby_LobbyError
    errCode As String
    errMessage As String
    resolvableLobbyErrorJsonData As nlohmann_json   ' std::shared_ptr<nlohmann::json> (Nothing = nullptr)
End Type

' struct GameListing
Public Type netlobby_GameListing
    gameId As String
    details As netlobby_GameDetails
    availableConnectionTypes() As netlobby_ConnectionType  ' (0 .. availableConnectionTypesCount - 1)
    availableConnectionTypesCount As Long
End Type

' struct ListResponse
Public Type netlobby_ListResponse
    games() As netlobby_GameListing             ' games(0 .. gamesCount - 1)
    gamesCount As Long
    lobbyMOTD As String
End Type

' typedef tl::expected<ListResponse, LobbyError> ListResult
Public Type netlobby_ListResult
    has_value As Boolean
    value As netlobby_ListResponse
    error_ As netlobby_LobbyError
End Type

' struct GameJoinDetails
Public Type netlobby_GameJoinDetails
    connections() As netlobby_ConnectionInfo    ' connections(0 .. connectionsCount - 1)
    connectionsCount As Long
    joinTokens() As String                      ' joinTokens(0 .. joinTokensCount - 1)
    joinTokensCount As Long
End Type

' RequestJoinDetailsResults::errors' entry: std::unordered_map<IPVersion, std::variant<ConnectivityFailure, LobbyError>>
Public Type netlobby_RequestJoinDetailsError
    first As netlobby_IPVersion                 ' the key
    isLobbyError As Boolean                     ' the variant's alternative
    connectivityType As netlobby_ConnectivityFailureType  ' ConnectivityFailure::type
    connectivityDetails As String               ' ConnectivityFailure::details
    lobbyError As netlobby_LobbyError
End Type

' struct RequestJoinDetailsResults
Public Type netlobby_RequestJoinDetailsResults
    joinDetails As netlobby_GameJoinDetails
    errors() As netlobby_RequestJoinDetailsError ' errors(0 .. errorsCount - 1)
    errorsCount As Long
End Type

' typedef tl::expected<RequestJoinDetailsResults, LobbyError> JoinResult
Public Type netlobby_JoinResult
    has_value As Boolean
    value As netlobby_RequestJoinDetailsResults
    error_ As netlobby_LobbyError
End Type

'==========================================================================================================
' netlobby.cpp's own types

' struct HostGameContext
Public Type netlobby_HostGameContext
    gameId As String                            ' The lobby game ID
    hostToken As String                         ' A token the host provides to authenticate host actions with the lobby
    lobbyConnectToken As String                 ' A token the *lobby* provides to the host when reaching out to attempt & verify connection(s)
End Type

' struct WrappedCreateGameResponse
Public Type netlobby_WrappedCreateGameResponse
    ctx As netlobby_HostGameContext
    motd As String
    supportedHostJoinOpts As Long               ' std::unordered_set<HostJoinOptionType>: bit (1 << type) each
End Type

' struct RequestAttestation
Private Type RequestAttestation
    playerNameB64 As String
    playerPublicIdentity As String
    nonceB64 As String
    secondsSinceEpoch As String
    lobbyAddress As String
    has_gameId As Boolean                       ' optional<std::string> gameId
    gameId As String
    signature As String
End Type

' struct LobbyConnectCheckRequest
Private Type LobbyConnectCheckRequest
    lobbyConnectToken As String
    challenge As String
End Type

' struct LobbyConnectCheckResponse
Private Type LobbyConnectCheckResponse
    gameId As String
    playerNameB64 As String
    playerPublicIdentity As String
    secondsSinceEpoch As String
    signature As String
End Type

' struct RequestJoinResponse
Private Type RequestJoinResponse
    connections() As netlobby_ConnectionInfo
    connectionsCount As Long
    joinToken As String
End Type

' struct PendingJoinRequestData (pendingIPVersionRequests: requested / received per IPVersion 1, 2)
Private Type PendingJoinRequestData
    used As Boolean
    requested(1 To 2) As Boolean
    received(1 To 2) As Boolean                 ' PendingIPVersionJoinRequestStatus::received
    pendingResults As netlobby_RequestJoinDetailsResults
    resultHandler As WzFunc
End Type

' (VB) the results handed to a WzFunc: a handle into these stores (taken once)
Private joinResults() As netlobby_JoinResult
Private joinResultsCount As Long
Private listResults() As netlobby_ListResult
Private listResultsCount As Long
Private pendingJoins() As PendingJoinRequestData
Private pendingJoinsCount As Long

' (VB) the live hosting handlers by serial (their std::weak_ptr)
Private hostingHandlers As Collection
Private hostingHandlerSerial As Long
Private netplayLobbyHostingHandler As LobbyServerHostingHandlerImpl  ' (VB) netplay.cpp's lobbyHostingHandler (below)

Public Const FN_netlobby_enumGamesResponse As Long = 17210          ' EnumerateGames' request.onResponse
Public Const FN_netlobby_enumGamesFailure As Long = 17211           ' EnumerateGames' request.onFailure
Public Const FN_netlobby_deliverListResult As Long = 17212          ' [completionHandlerFunc, result]() { completionHandlerFunc(result); }
Public Const FN_netlobby_enumGamesError As Long = 17213             ' [completionHandlerFunc, httpStatusCode, jsonData]()
Public Const FN_netlobby_enumGamesConnError As Long = 17214         ' [completionHandlerFunc, type, httpStatusCode]()
Public Const FN_netlobby_joinResponse As Long = 17215               ' buildResponseHandler(ipVersion)
Public Const FN_netlobby_joinFailure As Long = 17216                 ' buildFailureHandler(ipVersion)
Public Const FN_netlobby_deliverJoinResult As Long = 17217          ' [readyResults, resultHandlerCopy]()
Public Const FN_netlobby_logOnMainThread As Long = 17218            ' logErrorOnMainThread's [errorString]()
Public Const FN_netlobby_captchaResult As Long = 17219              ' makeCaptchaWidget's [resultHandler, token](WzString result)
Public Const FN_netlobby_slideResult As Long = 17220                ' makeSlideWidget's [resultHandler, token](uint32_t pointX, uint32_t pointY)
Public Const FN_netlobby_hostingAPIResponse As Long = 17221         ' buildHostingAPIResponseHandler's onResponse
Public Const FN_netlobby_hostingAPIHandle As Long = 17222           ' [weakSelf, handlerFunc, urlCopy, httpStatusCode, jsonData]() mutable
Public Const FN_netlobby_hostingAPIFailureMain As Long = 17223      ' [weakSelf, urlCopy, httpStatusCode, jsonData]() (handleAPIFailure)
Public Const FN_netlobby_hostingAPIStdFailure As Long = 17224       ' buildHostingAPIStandardFailureHandler's onFailure
Public Const FN_netlobby_createGameHandler As Long = 17225           ' createGameListing's handlerFunc
Public Const FN_netlobby_updateGameHandler As Long = 17226           ' sendListingUpdateImpl's handlerFunc
Public Const FN_netlobby_addConnectionHandler As Long = 17227        ' addConnectionsToGame's handlerFunc
Public Const FN_netlobby_addConnectionFailure As Long = 17228        ' addConnectionsToGame's onFailure ([protocol])
Public Const FN_netlobby_checkJoinResponse As Long = 17229           ' checkJoinRequest's onResponse ([resultHandler])
Public Const FN_netlobby_checkJoinResult As Long = 17230             ' [resultHandler, jsonData]()
Public Const FN_netlobby_checkJoinError As Long = 17231              ' [resultHandler, httpStatusCode, jsonData]()
Public Const FN_netlobby_checkJoinFailure As Long = 17232            ' checkJoinRequest's onFailure ([resultHandler])
Public Const FN_netlobby_checkJoinConnError As Long = 17233          ' [resultHandler, httpStatusCode]()
Public Const FN_netlobby_netplayGameCreated As Long = 17234         ' eventCallbacks.lobbyGameCreatedCallback (WzNet's, through NetApi)
Public Const FN_netlobby_netplayGamePublished As Long = 17235       ' eventCallbacks.lobbyGamePublishedCallback
Public Const FN_netlobby_netplayGameGone As Long = 17236            ' eventCallbacks.lobbyGameGoneCallback
Public Const FN_netlobby_netplayCheckJoinResult As Long = 17237     ' NETallowJoining's [i, uniqueJoinIDCopy](CheckJoinResult result)

'----------------------------------------------------------------------------------------------------------
' (VB) the stores

Public Function netlobby_JoinResult_store(r As netlobby_JoinResult) As Long
    ReDim Preserve joinResults(0 To joinResultsCount)
    joinResults(joinResultsCount) = r
    joinResultsCount = joinResultsCount + 1
    netlobby_JoinResult_store = joinResultsCount
End Function

Public Sub netlobby_JoinResult_take(ByVal handle As Long, r As netlobby_JoinResult)
    Dim empty_ As netlobby_JoinResult
    r = joinResults(handle - 1)
    joinResults(handle - 1) = empty_
End Sub

Public Function netlobby_ListResult_store(r As netlobby_ListResult) As Long
    ReDim Preserve listResults(0 To listResultsCount)
    listResults(listResultsCount) = r
    listResultsCount = listResultsCount + 1
    netlobby_ListResult_store = listResultsCount
End Function

Public Sub netlobby_ListResult_take(ByVal handle As Long, r As netlobby_ListResult)
    Dim empty_ As netlobby_ListResult
    r = listResults(handle - 1)
    listResults(handle - 1) = empty_
End Sub

' LobbyError(errCode, errMessage)
Public Sub netlobby_LobbyError_make(e As netlobby_LobbyError, ByVal errCode As String, ByVal errMessage As String)
    e.errCode = errCode
    e.errMessage = errMessage
    Set e.resolvableLobbyErrorJsonData = Nothing
End Sub

' (VB) json.hpp's type_error::create(302, what, &context), thrown
Private Sub throw_type_error_302(ByVal what As String)
    Err.Raise vbObjectError + 302, "nlohmann::json", "[json.exception.type_error.302] " & what
End Sub

' (VB) std::to_string of the seconds since the epoch (std::chrono::system_clock)
Private Function secondsSinceEpochString() As String
    secondsSinceEpochString = Format$(time_(), "0")
End Function

' (VB) the UTF-8 bytes of a string, as a std::string holds them: b(0 .. n - 1), n returned
Private Function stringBytes(ByVal s As String, b() As Byte) As Long
    stringBytes = utf8_encode(s, b)
    If stringBytes = 0 Then ReDim b(0 To 0)
End Function

'----------------------------------------------------------------------------------------------------------
' MARK: -

' uint32_t PlayerCounts::availableSlots() const
Public Function netlobby_PlayerCounts_availableSlots(pc As netlobby_PlayerCounts) As Long
    If ULT(pc.max, pc.current) Then
        netlobby_PlayerCounts_availableSlots = 0
        Exit Function
    End If
    netlobby_PlayerCounts_availableSlots = pc.max - pc.current
End Function

' bool HostGameContext::isValid() const
Public Function netlobby_HostGameContext_isValid(c As netlobby_HostGameContext) As Boolean
    netlobby_HostGameContext_isValid = Len(c.gameId) <> 0
End Function

' MARK: - Url parsing helpers

' UrlParts parse_url(const std::string& url): has_scheme / scheme and hostname (False: no match, {})
Private Function parse_url(ByVal url As String, has_scheme As Boolean, scheme As String, hostname As String) As Boolean
    Dim p As Long, sepPos As Long, i As Long, c As String, hostStart As Long, atPos As Long
    has_scheme = False
    scheme = ""
    hostname = ""
    p = 1
    ' ^(?:([^:/?#]+)://)?  -> Group 1: Optional scheme (captured)
    sepPos = InStr(url, "://")
    If sepPos > 1 Then
        For i = 1 To sepPos - 1
            c = Mid$(url, i, 1)
            If c = ":" Or c = "/" Or c = "?" Or c = "#" Then Exit For
        Next
        If i = sepPos Then
            scheme = Left$(url, sepPos - 1)
            has_scheme = True
            p = sepPos + 3
        End If
    End If
    ' (?:[^/?#@]+@)?  -> Optional userinfo (username:password@) - NOT captured
    hostStart = p
    atPos = 0
    For i = p To Len(url)
        c = Mid$(url, i, 1)
        If c = "/" Or c = "?" Or c = "#" Then Exit For
        If c = "@" Then
            atPos = i
            Exit For
        End If
    Next
    If atPos > hostStart Then p = atPos + 1
    ' ([^/?#:]+)  -> Group 2: Hostname (stops at /, ?, #, or : for port)
    For i = p To Len(url)
        c = Mid$(url, i, 1)
        If c = "/" Or c = "?" Or c = "#" Or c = ":" Then Exit For
    Next
    If i = p Then
        ' no match: return {}
        has_scheme = False
        scheme = ""
        Exit Function
    End If
    hostname = Mid$(url, p, i - p)
    parse_url = True
End Function

' optional<std::string> url_string_get_protocol(const std::string& url) (has_value: the result)
Private Function url_string_get_protocol(ByVal url As String, protocol As String) As Boolean
    Dim pos As Long
    pos = InStr(url, "://")         ' Find the position of "://"
    If pos <> 0 Then
        ' Extract the part before the delimiter (e.g., "https")
        protocol = Left$(url, pos - 1)
        url_string_get_protocol = True
    Else
        ' delimiter is not found
        url_string_get_protocol = False
    End If
End Function

Private Function buildLobbyRequestBaseUrl(ByVal lobbyAddress As String) As String
    Dim protocol As String
    If Not url_string_get_protocol(lobbyAddress, protocol) Then
        ' Default to https if no protocol is specified
        buildLobbyRequestBaseUrl = "https://" & lobbyAddress
        Exit Function
    End If
    buildLobbyRequestBaseUrl = lobbyAddress
End Function

' Check if the lobbyAddress is trusted / known, with a trusted / known privacy policy (and secure transport)
' Used in various places to simplify matters for alternative lobbies, and to minimize data sent to unknown or insecure servers (with unknown privacy policies)
Public Function netlobby_isTrustedLobbyAddress(ByVal lobbyAddress As String) As Boolean
    Dim has_scheme As Boolean, scheme As String, hostname As String
    parse_url lobbyAddress, has_scheme, scheme, hostname
    If Not has_scheme Then scheme = "https"
    If hostname = "wzlobby.wz2100.net" And scheme = "https" Then
        netlobby_isTrustedLobbyAddress = True
        Exit Function
    End If
    netlobby_isTrustedLobbyAddress = False
End Function

' MARK: - LobbyError

Private Sub LobbyError_from_json(j As nlohmann_json, v As netlobby_LobbyError)
    Dim it As nlohmann_json
    v.errCode = j.at("code").get_string()

    v.errMessage = ""
    Set it = j.find("message")
    If Not it Is Nothing Then
        v.errMessage = it.get_string()
    End If

    Set it = j.find("data")
    If Not it Is Nothing Then
        Set v.resolvableLobbyErrorJsonData = it.clone()
    End If
End Sub

Private Function LobbyError_to_json(v As netlobby_LobbyError) As nlohmann_json
    Dim j As nlohmann_json
    Set j = json_object()
    j.set_item_ref "code", json_value(v.errCode)
    j.set_item_ref "message", json_value(v.errMessage)
    Set LobbyError_to_json = j
End Function

' (VB) try { err = jsonData.get<LobbyError>(); } catch (const std::exception&) { }: has_value returned
Private Function tryGetLobbyError(jsonData As nlohmann_json, err_ As netlobby_LobbyError) As Boolean
    Dim e As netlobby_LobbyError
    On Error GoTo failed
    If jsonData Is Nothing Then GoTo failed
    LobbyError_from_json jsonData, e
    err_ = e
    tryGetLobbyError = True
    Exit Function
failed:
    tryGetLobbyError = False
End Function

' optional<WzString> GetLocalizedLobbyErrorMessage(const LobbyError& lobbyError) (has_value: the result)
Public Function netlobby_GetLocalizedLobbyErrorMessage(lobbyError As netlobby_LobbyError, out As String) As Boolean
    Dim retVal As String
    If lobbyError.errCode = errCodeHostDisallowedJoinSource Then
        retVal = "This host has disallowed joins from certain hosting providers, VPNs, and/or proxies."
        retVal = retVal & vbLf
        retVal = retVal & "Please try another room or disable any proxies / VPNs and try again."
        out = retVal
        netlobby_GetLocalizedLobbyErrorMessage = True
        Exit Function
    End If
    If lobbyError.errCode = errCodeGameNotFound Then
        retVal = "Game not found in lobby."
        retVal = retVal & vbLf
        retVal = retVal & "The game may have already started, or the host may have disbanded the game lobby."
        out = retVal
        netlobby_GetLocalizedLobbyErrorMessage = True
        Exit Function
    End If

    out = ""
    netlobby_GetLocalizedLobbyErrorMessage = False
End Function

' bool LobbyError::isResolvableLobbyError() const
Public Function netlobby_LobbyError_isResolvableLobbyError(e As netlobby_LobbyError) As Boolean
    If e.errCode = "CAPTCHA_REQUIRED" Then
        netlobby_LobbyError_isResolvableLobbyError = True
        Exit Function
    End If
    netlobby_LobbyError_isResolvableLobbyError = False
End Function

' MARK: ImageType

Private Function ImageType_from_json(j As nlohmann_json) As ImageType
    Dim strValue As String
    strValue = j.get_string()
    If strValue = "png" Then
        ImageType_from_json = ImageType_PNG
        Exit Function
    ElseIf strValue = "jpg" Or strValue = "jpeg" Then
        ImageType_from_json = ImageType_JPEG
        Exit Function
    Else
        throw_type_error_302 "unsupported value for ImageType field: " & strValue
    End If
End Function

' MARK: B64ImageData (the class B64ImageData)

Private Function B64ImageData_from_json(j As nlohmann_json) As B64ImageData
    Dim v As B64ImageData
    Set v = New B64ImageData
    v.b64Data = j.at("data").get_string()
    v.type_ = ImageType_from_json(j.at("type"))
    Set B64ImageData_from_json = v
End Function

' MARK: CaptchaVerifyPrompt (token; data), from_json

Private Sub CaptchaVerifyPrompt_from_json(j As nlohmann_json, token As String, data As CaptchaVerifyPromptData)
    Dim d As nlohmann_json
    token = j.at("token").get_string()

    Set d = j.at("data")
    Set data.image = B64ImageData_from_json(d.at("image"))
End Sub

' struct CaptchaVerifyResponse { token; result }: to_json
Private Function CaptchaVerifyResponse_to_json(ByVal token As String, ByVal result As String) As nlohmann_json
    Dim j As nlohmann_json, resp As nlohmann_json
    Set j = json_object()
    j.set_item_ref "token", json_value(token)

    Set resp = json_object()
    resp.set_item_ref "result", json_value(result)

    j.set_item_ref "response", resp
    Set CaptchaVerifyResponse_to_json = j
End Function

' MARK: SlideVerifyPrompt (token; data), from_json

Private Sub SlideVerifyPrompt_from_json(j As nlohmann_json, token As String, data As SlideVerifyPromptData)
    Dim d As nlohmann_json
    token = j.at("token").get_string()

    Set d = j.at("data")
    Set data.image = B64ImageData_from_json(d.at("image"))
    Set data.tile = B64ImageData_from_json(d.at("tile"))
    data.tile_width = d.at("tile_width").get_uint()
    data.tile_height = d.at("tile_height").get_uint()
    data.tile_x = d.at("tile_x").get_uint()
    data.tile_y = d.at("tile_y").get_uint()
End Sub

' struct SlideVerifyResponse { token; x; y }: to_json
Private Function SlideVerifyResponse_to_json(ByVal token As String, ByVal x As Long, ByVal y As Long) As nlohmann_json
    Dim j As nlohmann_json, resp As nlohmann_json
    Set j = json_object()
    j.set_item_ref "token", json_value(token)

    Set resp = json_object()
    resp.set_item_ref "x", json_value_unsigned(x)
    resp.set_item_ref "y", json_value_unsigned(y)

    j.set_item_ref "response", resp
    Set SlideVerifyResponse_to_json = j
End Function

' MARK: GetLobbyErrorResolveWidget

Private Function constructChallengeResolveWidgetFromErrorJson(j As nlohmann_json, ByVal resultHandler As WzFunc) As WIDGET
    Dim it As nlohmann_json, typeStr As String, token As String, captcha As CaptchaVerifyPromptData, slide As SlideVerifyPromptData
    If Not (Not resultHandler Is Nothing) Then ASSERT_FAILURE "resultHandler != nullptr", "lib/netplay/netlobby.cpp:387", "constructChallengeResolveWidgetFromErrorJson", "Null resultHandler": Exit Function

    On Error GoTo failed
    Set it = j.find("type")
    If it Is Nothing Then
        debug_ 393, LOG_LOBBY, "constructChallengeResolveWidgetFromErrorJson", "Missing type field"
        Exit Function
    End If
    If Not it.is_string() Then
        debug_ 393, LOG_LOBBY, "constructChallengeResolveWidgetFromErrorJson", "Missing type field"
        Exit Function
    End If

    typeStr = it.get_string()
    If typeStr = "captcha" Then
        CaptchaVerifyPrompt_from_json j, token, captcha
        Set constructChallengeResolveWidgetFromErrorJson = WzVerifyPrompt_makeCaptchaWidget(captcha, WzFunc_new(FN_netlobby_captchaResult, resultHandler, token))
        Exit Function
    ElseIf typeStr = "slide" Or typeStr = "dragdrop" Then
        SlideVerifyPrompt_from_json j, token, slide
        Set constructChallengeResolveWidgetFromErrorJson = WzVerifyPrompt_makeSlideWidget(slide, typeStr = "dragdrop", WzFunc_new(FN_netlobby_slideResult, resultHandler, token))
        Exit Function
    Else
        ' unknown type
        debug_ 422, LOG_LOBBY, "constructChallengeResolveWidgetFromErrorJson", "Unknown type field value: %s", typeStr
        Exit Function
    End If
    Exit Function
failed:
    debug_ 427, LOG_LOBBY, "constructChallengeResolveWidgetFromErrorJson", "Failed parsing json: %s", Err.Description
    Set constructChallengeResolveWidgetFromErrorJson = Nothing
End Function

' std::shared_ptr<WIDGET> GetLobbyErrorResolveWidget(const LobbyError& lobbyError, const LobbyErrorResolveResultHandler& resultHandler)
' (resultHandler gets the resolution data: an nlohmann_json)
Public Function netlobby_GetLobbyErrorResolveWidget(lobbyError As netlobby_LobbyError, ByVal resultHandler As WzFunc) As WIDGET
    ' Certain special LobbyErrors might be resolvable by the user

    If lobbyError.errCode = "CAPTCHA_REQUIRED" Then
        If lobbyError.resolvableLobbyErrorJsonData Is Nothing Then
            Set netlobby_GetLobbyErrorResolveWidget = Nothing
            Exit Function
        End If

        Set netlobby_GetLobbyErrorResolveWidget = constructChallengeResolveWidgetFromErrorJson(lobbyError.resolvableLobbyErrorJsonData, resultHandler)
        Exit Function
    End If

    Set netlobby_GetLobbyErrorResolveWidget = Nothing
End Function

' MARK: - JSON conversion

' MARK: NetcodeVer
' NetcodeVer has special compressed JSON format of:
' [major,minor]

Private Function NetcodeVer_to_json(v As netlobby_NetcodeVer) As nlohmann_json
    Dim j As nlohmann_json
    Set j = json_array(True)
    j.push_back_ref json_value_unsigned(v.major, True)
    j.push_back_ref json_value_unsigned(v.minor, True)
    Set NetcodeVer_to_json = j
End Function

Private Sub NetcodeVer_from_json(j As nlohmann_json, v As netlobby_NetcodeVer)
    v.major = j.at_idx(0).get_uint()
    v.minor = j.at_idx(1).get_uint()
End Sub

' MARK: HostDetails

Private Function HostDetails_to_json(v As netlobby_HostDetails) As nlohmann_json
    Dim j As nlohmann_json
    Set j = json_object(True)
    j.set_item_ref "name", json_value(v.name, True)
    j.set_item_ref "pk", json_value(v.publicIdentity, True)
    Set HostDetails_to_json = j
End Function

Private Sub HostDetails_from_json(j As nlohmann_json, v As netlobby_HostDetails)
    v.name = j.at("name").get_string()
    v.publicIdentity = j.at("pk").get_string()
End Sub

' MARK: MapDetails
' MapDetails has special compressed JSON format of:
' ["name","hash"]

Private Function MapDetails_to_json(v As netlobby_MapDetails) As nlohmann_json
    Dim j As nlohmann_json
    Set j = json_array(True)
    j.push_back_ref json_value(v.name, True)
    j.push_back_ref json_value(v.hash, True)
    Set MapDetails_to_json = j
End Function

Private Sub MapDetails_from_json(j As nlohmann_json, v As netlobby_MapDetails)
    v.name = j.at_idx(0).get_string()
    v.hash = j.at_idx(1).get_string()
End Sub

' MARK: ModDetails

Private Function ModDetails_to_json(v As netlobby_ModDetails) As nlohmann_json
    Dim j As nlohmann_json
    Set j = json_object(True)
    j.set_item_ref "name", json_value(v.name, True)
    j.set_item_ref "hash", json_value(v.hash, True)
    Set ModDetails_to_json = j
End Function

Private Sub ModDetails_from_json(j As nlohmann_json, v As netlobby_ModDetails)
    v.name = j.at("name").get_string()
    v.hash = j.at("hash").get_string()
End Sub

' MARK: PlayerCounts

Private Function PlayerCounts_to_json(v As netlobby_PlayerCounts) As nlohmann_json
    Dim j As nlohmann_json
    Set j = json_object(True)
    j.set_item_ref "cur", json_value_unsigned(v.current, True)
    j.set_item_ref "max", json_value_unsigned(v.max, True)
    Set PlayerCounts_to_json = j
End Function

Private Sub PlayerCounts_from_json(j As nlohmann_json, v As netlobby_PlayerCounts)
    Dim it As nlohmann_json
    v.max = j.at("max").get_uint()

    v.current = 0
    Set it = j.find("cur")
    If Not it Is Nothing Then
        v.current = it.get_uint()
    End If
End Sub

' MARK: GameDetails

Public Function netlobby_GameDetails_to_json(v As netlobby_GameDetails) As nlohmann_json
    Dim j As nlohmann_json, mods As nlohmann_json, i As Long
    Set j = json_object(True)

    j.set_item_ref "v", json_value(v.versionStr, True)
    j.set_item_ref "netv", NetcodeVer_to_json(v.netcodeVer)

    j.set_item_ref "name", json_value(v.name, True)
    j.set_item_ref "host", HostDetails_to_json(v.host)
    j.set_item_ref "map", MapDetails_to_json(v.map)
    If v.modsCount <> 0 Then
        Set mods = json_array(True)
        For i = 0 To v.modsCount - 1
            mods.push_back_ref ModDetails_to_json(v.mods(i))
        Next
        j.set_item_ref "mods", mods
    End If
    j.set_item_ref "players", PlayerCounts_to_json(v.players)
    j.set_item_ref "specs", PlayerCounts_to_json(v.spectators)
    j.set_item_ref "ally", json_value_unsigned(v.alliancesMode And &HFF&, True)
    If v.blindMode <> 0 Then
        j.set_item_ref "blind", json_value_unsigned(v.blindMode And &HFF&, True)
    End If
    j.set_item_ref "lim", json_value_unsigned(v.limits And &HFF&, True)
    If v.isPrivate Then
        j.set_item_ref "private", json_value(True, True)
    End If
    Set netlobby_GameDetails_to_json = j
End Function

' (the operator== of the netlobby.h structs)
Private Function ModsEqual(a As netlobby_GameDetails, b As netlobby_GameDetails) As Boolean
    Dim i As Long
    If a.modsCount <> b.modsCount Then Exit Function
    For i = 0 To a.modsCount - 1
        If a.mods(i).name <> b.mods(i).name Or a.mods(i).hash <> b.mods(i).hash Then Exit Function
    Next
    ModsEqual = True
End Function

Public Function netlobby_serializeGameDetailsChangedFields(newGameDetails As netlobby_GameDetails, lastPersistedGameDetails As netlobby_GameDetails) As nlohmann_json
    Dim gameDetailsJson As nlohmann_json
    ' Start by converting to json
    Set gameDetailsJson = netlobby_GameDetails_to_json(newGameDetails)

    ' Then check individual fields and erase them if they are the same as lastPersistedGameDetails
    If newGameDetails.versionStr = lastPersistedGameDetails.versionStr Then
        gameDetailsJson.erase_ "v"
    End If
    If newGameDetails.netcodeVer.major = lastPersistedGameDetails.netcodeVer.major And newGameDetails.netcodeVer.minor = lastPersistedGameDetails.netcodeVer.minor Then
        gameDetailsJson.erase_ "netv"
    End If
    If newGameDetails.name = lastPersistedGameDetails.name Then
        gameDetailsJson.erase_ "name"
    End If
    If newGameDetails.host.name = lastPersistedGameDetails.host.name And newGameDetails.host.publicIdentity = lastPersistedGameDetails.host.publicIdentity Then
        gameDetailsJson.erase_ "host"
    End If
    If newGameDetails.map.name = lastPersistedGameDetails.map.name And newGameDetails.map.hash = lastPersistedGameDetails.map.hash Then
        gameDetailsJson.erase_ "map"
    End If
    If ModsEqual(newGameDetails, lastPersistedGameDetails) Then
        gameDetailsJson.erase_ "mods"
    End If
    If newGameDetails.players.current = lastPersistedGameDetails.players.current And newGameDetails.players.max = lastPersistedGameDetails.players.max Then
        gameDetailsJson.erase_ "players"
    End If
    If newGameDetails.spectators.current = lastPersistedGameDetails.spectators.current And newGameDetails.spectators.max = lastPersistedGameDetails.spectators.max Then
        gameDetailsJson.erase_ "specs"
    End If
    If newGameDetails.alliancesMode = lastPersistedGameDetails.alliancesMode Then
        gameDetailsJson.erase_ "ally"
    End If
    If newGameDetails.blindMode = lastPersistedGameDetails.blindMode Then
        gameDetailsJson.erase_ "blind"
    End If
    If newGameDetails.limits = lastPersistedGameDetails.limits Then
        gameDetailsJson.erase_ "lim"
    End If
    If newGameDetails.isPrivate = lastPersistedGameDetails.isPrivate Then
        gameDetailsJson.erase_ "private"
    End If

    Set netlobby_serializeGameDetailsChangedFields = gameDetailsJson
End Function

Public Sub netlobby_GameDetails_from_json(j As nlohmann_json, v As netlobby_GameDetails)
    Dim it As nlohmann_json, i As Long, fresh As netlobby_PlayerCounts
    v.versionStr = j.at("v").get_string()
    NetcodeVer_from_json j.at("netv"), v.netcodeVer

    v.name = j.at("name").get_string()
    HostDetails_from_json j.at("host"), v.host
    MapDetails_from_json j.at("map"), v.map

    Erase v.mods
    v.modsCount = 0
    Set it = j.find("mods")
    If Not it Is Nothing Then
        ' it.value().get<std::vector<ModDetails>>()
        If Not it.is_array() Then throw_type_error_302 "type must be array, but is " & it.type_name()
        If it.size() > 0 Then ReDim v.mods(0 To it.size() - 1)
        For i = 0 To it.size() - 1
            ModDetails_from_json it.at_idx(i), v.mods(i)
        Next
        v.modsCount = it.size()
    End If

    PlayerCounts_from_json j.at("players"), v.players

    v.spectators = fresh
    Set it = j.find("specs")
    If Not it Is Nothing Then
        PlayerCounts_from_json it, v.spectators
    End If

    v.alliancesMode = 0
    Set it = j.find("ally")
    If Not it Is Nothing Then
        v.alliancesMode = it.get_uint() And &HFF&
    End If

    v.blindMode = 0
    Set it = j.find("blind")
    If Not it Is Nothing Then
        v.blindMode = it.get_uint() And &HFF&
    End If

    v.limits = 0
    Set it = j.find("lim")
    If Not it Is Nothing Then
        v.limits = it.get_uint() And &HFF&
    End If

    v.isPrivate = False
    Set it = j.find("private")
    If Not it Is Nothing Then
        v.isPrivate = it.get_bool()
    End If
End Sub

' MARK: HostJoinOptionValue

Private Function HostJoinOptionValue_to_string(ByVal m As netlobby_HostJoinOptionValue) As String
    Select Case m
    Case netlobby_HostJoinOptionValue_BlockAll
        HostJoinOptionValue_to_string = "block"
    Case netlobby_HostJoinOptionValue_BlockSuspicious
        HostJoinOptionValue_to_string = "blockSus"
    Case netlobby_HostJoinOptionValue_Allow
        HostJoinOptionValue_to_string = "allow"
    Case Else
        HostJoinOptionValue_to_string = ""      ' silence compiler warning
    End Select
End Function

' optional<HostJoinOptionValue> host_join_option_value_from_string(std::string_view s) (has_value: the result)
Private Function host_join_option_value_from_string(ByVal s As String, v As netlobby_HostJoinOptionValue) As Boolean
    host_join_option_value_from_string = True
    If s = "block" Then
        v = netlobby_HostJoinOptionValue_BlockAll
    ElseIf s = "blockSus" Then
        v = netlobby_HostJoinOptionValue_BlockSuspicious
    ElseIf s = "allow" Then
        v = netlobby_HostJoinOptionValue_Allow
    Else
        host_join_option_value_from_string = False
    End If
End Function

Private Function HostJoinOptionValue_from_json(j As nlohmann_json) As netlobby_HostJoinOptionValue
    Dim v As netlobby_HostJoinOptionValue
    If Not host_join_option_value_from_string(j.get_string(), v) Then throw_type_error_302 "invalid value for HostJoinOptionValue: " & j.type_name()
    HostJoinOptionValue_from_json = v
End Function

' MARK: HostJoinOptions

Public Function netlobby_HostJoinOptions_to_json(v As netlobby_HostJoinOptions) As nlohmann_json
    Dim j As nlohmann_json
    Set j = json_object(True)
    j.set_item_ref "botProt", json_value(HostJoinOptionValue_to_string(v.botProt), True)
    j.set_item_ref "proxyIps", json_value(HostJoinOptionValue_to_string(v.proxyIPs), True)
    j.set_item_ref "hostingIps", json_value(HostJoinOptionValue_to_string(v.hostingIPs), True)
    Set netlobby_HostJoinOptions_to_json = j
End Function

Private Sub HostJoinOptions_from_json(j As nlohmann_json, v As netlobby_HostJoinOptions)
    v.botProt = HostJoinOptionValue_from_json(j.at("botProt"))
    v.proxyIPs = HostJoinOptionValue_from_json(j.at("proxyIps"))
    v.hostingIPs = HostJoinOptionValue_from_json(j.at("hostingIps"))
End Sub

' HostJoinOptions::HostJoinOptions() (the defaults)
Public Sub netlobby_HostJoinOptions_init(v As netlobby_HostJoinOptions)
    v.botProt = netlobby_HostJoinOptionValue_BlockAll
    v.proxyIPs = netlobby_HostJoinOptionValue_BlockSuspicious
    v.hostingIPs = netlobby_HostJoinOptionValue_BlockAll
End Sub

' MARK: HostGameContext
' (to_json deliberately unimplemented - should never be sending the HostGameContext anywhere directly)

Private Sub HostGameContext_from_json(j As nlohmann_json, v As netlobby_HostGameContext)
    v.gameId = j.at("gameId").get_string()
    v.hostToken = j.at("hostToken").get_string()
    v.lobbyConnectToken = j.at("lobbyToken").get_string()
End Sub

' MARK: HostJoinOptionType

' optional<HostJoinOptionType> host_join_option_type_from_string(std::string_view s) (has_value: the result)
Private Function host_join_option_type_from_string(ByVal s As String, v As netlobby_HostJoinOptionType) As Boolean
    host_join_option_type_from_string = True
    If s = "botProt" Then
        v = netlobby_HostJoinOptionType_BotProt
    ElseIf s = "proxyIps" Then
        v = netlobby_HostJoinOptionType_ProxyIPs
    ElseIf s = "hostingIps" Then
        v = netlobby_HostJoinOptionType_HostingIPs
    Else
        host_join_option_type_from_string = False
    End If
End Function

' MARK: WrappedCreateGameResponse

Public Sub netlobby_WrappedCreateGameResponse_from_json(j As nlohmann_json, v As netlobby_WrappedCreateGameResponse)
    Dim it As nlohmann_json, i As Long
    HostGameContext_from_json j, v.ctx

    Set it = j.find("motd")
    If Not it Is Nothing Then
        v.motd = it.get_string()
    End If

    v.supportedHostJoinOpts = 0
    Set it = j.find("hostJoinOpts")
    If Not it Is Nothing Then
        If it.is_array() Then
            ' Grab any values we know, ignore any we don't
            For i = 0 To it.size() - 1
                addSupportedHostJoinOpt it.at_idx(i), v.supportedHostJoinOpts
            Next
        End If
    End If
End Sub

' (the loop's try block: silently ignore issues converting)
Private Sub addSupportedHostJoinOpt(opt As nlohmann_json, mask As Long)
    Dim strVal As String, optType As netlobby_HostJoinOptionType
    On Error GoTo ignored
    strVal = opt.get_string()
    If host_join_option_type_from_string(strVal, optType) Then
        mask = mask Or (2 ^ optType)
    End If
ignored:
End Sub

' MARK: ConnectionMethod

Private Function ConnectionMethod_to_string(ByVal m As netlobby_ConnectionMethod) As String
    Select Case m
    Case netlobby_ConnectionMethod_TCP_DIRECT
        ConnectionMethod_to_string = "tcp"
    Case netlobby_ConnectionMethod_GNS_DIRECT
        ConnectionMethod_to_string = "gns"
    Case Else
        ConnectionMethod_to_string = ""
    End Select
End Function

Private Function connection_method_from_string(ByVal s As String, v As netlobby_ConnectionMethod) As Boolean
    connection_method_from_string = True
    If s = "tcp" Then
        v = netlobby_ConnectionMethod_TCP_DIRECT
    ElseIf s = "gns" Then
        v = netlobby_ConnectionMethod_GNS_DIRECT
    Else
        connection_method_from_string = False
    End If
End Function

' MARK: IPVersion

' const char* to_string(IPVersion m)
Public Function netlobby_to_string(ByVal m As netlobby_IPVersion) As String
    Select Case m
    Case netlobby_IPVersion_IPv4
        netlobby_to_string = "ip4"
    Case netlobby_IPVersion_IPv6
        netlobby_to_string = "ip6"
    Case Else
        netlobby_to_string = ""
    End Select
End Function

Private Function ip_version_from_string(ByVal s As String, v As netlobby_IPVersion) As Boolean
    ip_version_from_string = True
    If s = "ip4" Then
        v = netlobby_IPVersion_IPv4
    ElseIf s = "ip6" Then
        v = netlobby_IPVersion_IPv6
    Else
        ip_version_from_string = False
    End If
End Function

' MARK: ConnectionType

Private Sub ConnectionType_from_json(j As nlohmann_json, v As netlobby_ConnectionType)
    Dim strVal As String, delimPos As Long, leftPart As String, rightPart As String
    strVal = j.get_string()

    delimPos = InStr(strVal, ":")
    If delimPos = 0 Then
        ' assume it's just a ConnectionMethod string, with IPVersion unspecified
        v.has_ipVersion = False
        If Not connection_method_from_string(strVal, v.method) Then throw_type_error_302 "invalid value for ConnectionMethod: " & j.type_name()
        Exit Sub
    End If

    ' appears to be a <ConnectionMethod>:<IPVersion> string
    ' Extract the left part (before the colon)
    leftPart = Left$(strVal, delimPos - 1)
    ' Extract the right part (after the colon)
    rightPart = Mid$(strVal, delimPos + 1)

    If Not connection_method_from_string(leftPart, v.method) Then throw_type_error_302 "invalid value for ConnectionMethod: " & j.type_name()
    v.has_ipVersion = ip_version_from_string(rightPart, v.ipVersion)
End Sub

Private Function ConnectionType_to_json(v As netlobby_ConnectionType) As nlohmann_json
    Dim str_ As String
    str_ = ConnectionMethod_to_string(v.method)

    If v.has_ipVersion Then
        str_ = str_ & ":"
        str_ = str_ & netlobby_to_string(v.ipVersion)
    End If

    Set ConnectionType_to_json = json_value(str_, True)
End Function

' MARK: ConnectionInfo

Public Function netlobby_ConnectionInfo_to_json(v As netlobby_ConnectionInfo) As nlohmann_json
    Dim j As nlohmann_json
    Set j = json_object(True)
    j.set_item_ref "type", ConnectionType_to_json(v.typ)
    j.set_item_ref "host", json_value(v.host, True)
    j.set_item_ref "port", json_value_unsigned(v.port And &HFFFF&, True)
    Set netlobby_ConnectionInfo_to_json = j
End Function

Public Sub netlobby_ConnectionInfo_from_json(j As nlohmann_json, v As netlobby_ConnectionInfo)
    ConnectionType_from_json j.at("type"), v.typ
    v.host = j.at("host").get_string()
    v.port = j.at("port").get_uint() And &HFFFF&
End Sub

' MARK: ConnectionStatus

Private Function ConnectionStatus_from_json(j As nlohmann_json) As netlobby_ConnectionStatus
    Dim s As String
    s = j.get_string()
    If s = "submitted" Then
        ConnectionStatus_from_json = netlobby_ConnectionStatus_Submitted
    ElseIf s = "validated" Then
        ConnectionStatus_from_json = netlobby_ConnectionStatus_Validated
    ElseIf s = "error" Then
        ConnectionStatus_from_json = netlobby_ConnectionStatus_Error
    Else
        throw_type_error_302 "invalid value for ConnectionStatus: " & j.type_name()
    End If
End Function

' MARK: HostGameStatus

Public Sub netlobby_HostGameStatus_from_json(j As nlohmann_json, v As netlobby_HostGameStatus)
    Dim conns As nlohmann_json, i As Long
    v.gameId = j.at("gameId").get_string()
    v.published = j.at("published").get_bool()
    Set conns = j.at("connections")
    If Not conns.is_array() Then throw_type_error_302 "type must be array, but is " & conns.type_name()
    Erase v.connectionStatusInfo
    v.connectionStatusInfoCount = 0
    If conns.size() > 0 Then ReDim v.connectionStatusInfo(0 To conns.size() - 1)
    For i = 0 To conns.size() - 1
        ' HostConnectionStatusInfo from_json
        ConnectionType_from_json conns.at_idx(i).at("type"), v.connectionStatusInfo(i).typ
        v.connectionStatusInfo(i).status = ConnectionStatus_from_json(conns.at_idx(i).at("status"))
    Next
    v.connectionStatusInfoCount = conns.size()
End Sub

' MARK: JoinRequest

Public Function netlobby_JoinRequest_to_json(v As netlobby_JoinRequest) As nlohmann_json
    Dim j As nlohmann_json
    Set j = json_object(True)
    j.set_item_ref "joinToken", json_value(v.joinToken, True)
    j.set_item_ref "ip", json_value(v.ip, True)
    j.set_item_ref "name", json_value(v.playerName, True)
    j.set_item_ref "pk", json_value(v.playerPublicIdentity, True)
    j.set_item_ref "connMethod", json_value(ConnectionMethod_to_string(v.connectionMethod), True)
    j.set_item_ref "asSpec", json_value(v.asSpectator, True)
    Set netlobby_JoinRequest_to_json = j
End Function

' MARK: CheckJoinRequestResult

Private Sub CheckJoinRequestResult_from_json(j As nlohmann_json, v As netlobby_CheckJoinRequestResult)
    Dim it As nlohmann_json
    v.pass = j.at("pass").get_bool()

    v.error_ = ""
    Set it = j.find("error")
    If Not it Is Nothing Then
        v.error_ = it.get_string()
    End If
End Sub

' MARK: RequestAttestation

' RequestAttestation::RequestAttestation(playerName, playerIdentity, lobbyAddress, gameId = nullopt)
Private Sub RequestAttestation_init(v As RequestAttestation, ByVal playerName As String, playerIdentity As EcKey, ByVal lobbyAddress As String, _
        ByVal has_gameId As Boolean, ByVal gameId As String)
    Dim b() As Byte, n As Long, nonce(0 To 31) As Byte, strToSign As String, sigBytes() As Byte
    v.lobbyAddress = lobbyAddress
    v.has_gameId = has_gameId
    v.gameId = gameId

    n = stringBytes(playerName, b)                      ' vecFromWzString(playerName)
    v.playerNameB64 = base64Encode(b, n)
    n = EcKey_toBytes(playerIdentity, EcKey_Public, b)
    v.playerPublicIdentity = base64Encode(b, n)

    genSecRandomBytes nonce, 0, 32
    v.nonceB64 = base64Encode(nonce, 32)

    ' Get current time in UTC (system_clock)
    v.secondsSinceEpoch = secondsSinceEpochString()

    ' Sign a string containing:
    ' <random nonce b64><seconds since epoch><playerNameB64><lobbyAddress><gameId, if present>
    strToSign = v.nonceB64 & v.secondsSinceEpoch & v.playerNameB64 & v.lobbyAddress
    If has_gameId Then strToSign = strToSign & gameId

    n = stringBytes(strToSign, b)
    If Len(strToSign) = 0 Then n = 0
    n = EcKey_sign(playerIdentity, b, 0, n, sigBytes)
    v.signature = base64Encode(sigBytes, n)
End Sub

Private Function RequestAttestation_to_json(v As RequestAttestation) As nlohmann_json
    Dim j As nlohmann_json
    Set j = json_object(True)
    j.set_item_ref "pk", json_value(v.playerPublicIdentity, True)
    j.set_item_ref "name", json_value(v.playerNameB64, True)
    j.set_item_ref "n", json_value(v.nonceB64, True)
    j.set_item_ref "l", json_value(v.lobbyAddress, True)
    If v.has_gameId Then
        j.set_item_ref "gid", json_value(v.gameId, True)
    End If
    j.set_item_ref "at", json_value(v.secondsSinceEpoch, True)
    j.set_item_ref "sig", json_value(v.signature, True)
    Set RequestAttestation_to_json = j
End Function

' nlohmann::ordered_json(RequestAttestation(playerName, playerIdentity, lobbyAddress, gameId))
Public Function netlobby_makeAttestationJson(ByVal playerName As String, playerIdentity As EcKey, ByVal lobbyAddress As String, _
        ByVal has_gameId As Boolean, ByVal gameId As String) As nlohmann_json
    Dim a As RequestAttestation
    RequestAttestation_init a, playerName, playerIdentity, lobbyAddress, has_gameId, gameId
    Set netlobby_makeAttestationJson = RequestAttestation_to_json(a)
End Function

' MARK: LobbyConnectCheckRequest

Private Sub LobbyConnectCheckRequest_from_json(j As nlohmann_json, v As LobbyConnectCheckRequest)
    v.lobbyConnectToken = j.at("lobbyToken").get_string()
    v.challenge = j.at("challenge").get_string()
End Sub

' (the try block of respondToLobbyConnectionCheckRequest: parse + get; False with the exception's what())
Public Function netlobby_parseLobbyConnectCheckRequest(ByVal jsonRequest As String, lobbyConnectToken As String, challenge As String, what As String) As Boolean
    Dim jsonData As nlohmann_json, request As LobbyConnectCheckRequest
    On Error GoTo failed
    Set jsonData = json_parse(jsonRequest)
    LobbyConnectCheckRequest_from_json jsonData, request
    lobbyConnectToken = request.lobbyConnectToken
    challenge = request.challenge
    netlobby_parseLobbyConnectCheckRequest = True
    Exit Function
failed:
    what = Err.Description
    netlobby_parseLobbyConnectCheckRequest = False
End Function

' MARK: LobbyConnectCheckResponse

' LobbyConnectCheckResponse(gameId, playerName, playerIdentity, request), to_json: the response's JSON
Public Function netlobby_makeLobbyConnectCheckResponseJson(ByVal gameId As String, ByVal playerName As String, playerIdentity As EcKey, ByVal challenge As String) As nlohmann_json
    Dim v As LobbyConnectCheckResponse, b() As Byte, n As Long, strToSign As String, sigBytes() As Byte, j As nlohmann_json
    v.gameId = gameId
    n = stringBytes(playerName, b)
    v.playerNameB64 = base64Encode(b, n)
    n = EcKey_toBytes(playerIdentity, EcKey_Public, b)
    v.playerPublicIdentity = base64Encode(b, n)

    ' Get current time in UTC (system_clock)
    v.secondsSinceEpoch = secondsSinceEpochString()

    ' Sign a string containing:
    ' WZLR:<request.Challenge><seconds since epoch><player name b64><gameId>
    strToSign = "WZLR:" & challenge & v.secondsSinceEpoch & v.playerNameB64 & gameId

    n = stringBytes(strToSign, b)
    n = EcKey_sign(playerIdentity, b, 0, n, sigBytes)
    v.signature = base64Encode(sigBytes, n)

    ' to_json
    Set j = json_object(True)
    j.set_item_ref "gid", json_value(v.gameId, True)
    j.set_item_ref "pk", json_value(v.playerPublicIdentity, True)
    j.set_item_ref "name", json_value(v.playerNameB64, True)
    j.set_item_ref "at", json_value(v.secondsSinceEpoch, True)
    j.set_item_ref "sig", json_value(v.signature, True)
    Set netlobby_makeLobbyConnectCheckResponseJson = j
End Function

'----------------------------------------------------------------------------------------------------------
' MARK: - The hosting handler's helpers

Public Function netlobby_getWZUserAgent() As String
    Static currentUserAgentString As String, initialised As Boolean
    If Not initialised Then
        currentUserAgentString = version_getHTTPUserAgentString()
        initialised = True
    End If
    netlobby_getWZUserAgent = currentUserAgentString
End Function

Public Sub netlobby_buildBaseHostingURLDataRequest(urlRequest As URLDataRequest)
    Dim fresh As URLDataRequest
    urlRequest = fresh

    ' We always want direct connections to the lobby server (skipping any locally-configured proxies or redirects)
    urlRequest.noProxy = True
    urlRequest.noFollowRedirects = True

    urlRequest.userAgent = netlobby_getWZUserAgent()

    URLRequestBase_setRequestHeader urlRequest, "Origin", "wz2100://wz2100.net"
    URLRequestBase_setRequestHeader urlRequest, "Referer", "https://github.com/Warzone2100/warzone2100?wz_internal=mp.host"
    URLRequestBase_setRequestHeader urlRequest, "Accept", "application/json"

    urlRequest.maxDownloadSizeLimit = 20# * 1024 * 1024    ' response should never be > 20 MB
End Sub

Public Function netlobby_addHostAuthorization(request As URLDataRequest, hostGameCtx As netlobby_HostGameContext) As Boolean
    If Not (Len(hostGameCtx.hostToken) <> 0) Then ASSERT_FAILURE "!hostGameCtx.hostToken.empty()", "lib/netplay/netlobby.cpp:1255", "addHostAuthorization", "Missing HostToken": Exit Function
    URLRequestBase_setRequestHeader request, "Authorization", "HostToken " & hostGameCtx.hostToken
    netlobby_addHostAuthorization = True
End Function

Public Function netlobby_LobbyDeleteGameImpl(ByVal lobbyServerAddress As String, hostGameCtx As netlobby_HostGameContext, ByVal matchStarted As Boolean) As Boolean
    Dim request As URLDataRequest
    If Not (netlobby_HostGameContext_isValid(hostGameCtx)) Then ASSERT_FAILURE "hostGameCtx.isValid()", "lib/netplay/netlobby.cpp:1262", "LobbyDeleteGameImpl", "Invalid HostGameContext": Exit Function

    netlobby_buildBaseHostingURLDataRequest request
    netlobby_addHostAuthorization request, hostGameCtx
    request.url = buildLobbyRequestBaseUrl(lobbyServerAddress) & "/api/v1/game/" & hostGameCtx.gameId
    If matchStarted Then
        request.url = request.url & "?started=1"
    End If
    request.protocol = InternetProtocol_IPv4
    URLDataRequest_setDelete request

    request.autoRetryStrategy = URLRequestAutoRetryStrategy_RetryAfter(3000, 3, Array(HttpRequestConnectionFailure, HttpStatusTooManyRequests))

    urlRequestData request

    netlobby_LobbyDeleteGameImpl = True
End Function

Public Sub netlobby_logErrorOnMainThread(ByVal prefix As String, ByVal error_ As String)
    Dim errorString As String
    errorString = prefix
    errorString = errorString & error_
    wzAsyncExecOnMainThread WzFunc_new(FN_netlobby_logOnMainThread, errorString)
End Sub

' (VB) the base URL of the lobby's API (buildLobbyRequestBaseUrl, for the hosting handler)
Public Function netlobby_buildLobbyRequestBaseUrl(ByVal lobbyAddress As String) As String
    netlobby_buildLobbyRequestBaseUrl = buildLobbyRequestBaseUrl(lobbyAddress)
End Function

' static InternetProtocol ipVersionToInternetProtocol(IPVersion ver)
Public Function netlobby_ipVersionToInternetProtocol(ByVal ver As netlobby_IPVersion) As InternetProtocol
    Select Case ver
    Case netlobby_IPVersion_IPv4: netlobby_ipVersionToInternetProtocol = InternetProtocol_IPv4
    Case netlobby_IPVersion_IPv6: netlobby_ipVersionToInternetProtocol = InternetProtocol_IPv6
    Case Else: netlobby_ipVersionToInternetProtocol = InternetProtocol_IP_ANY   ' silence compiler warning
    End Select
End Function

' (VB) parse the response body as JSON: the try block every handler starts with (Nothing on failure, logged with
' logExceptionOnMainThread("Failed to parse JSON response: ", e))
Private Function parseResponseJson(data As MemoryStruct) As nlohmann_json
    Dim b() As Byte
    On Error GoTo failed
    b = data.memory
    Set parseResponseJson = json_parse_bytes(b, 0, data.size)
    Exit Function
failed:
    netlobby_logErrorOnMainThread "Failed to parse JSON response: ", Err.Description
    Set parseResponseJson = json_value(Null)      ' jsonData = nlohmann::json()
End Function

' (VB) nlohmann::json() for a parse that failed: null
Private Function jsonNull() As nlohmann_json
    Set jsonNull = json_value(Null)
End Function

'----------------------------------------------------------------------------------------------------------
' (VB) the hosting handlers' std::weak_ptr

Public Function LobbyServerHostingHandlerImpl_register(ByVal h As LobbyServerHostingHandlerImpl) As Long
    If hostingHandlers Is Nothing Then Set hostingHandlers = New Collection
    hostingHandlerSerial = hostingHandlerSerial + 1
    hostingHandlers.Add h, "h" & hostingHandlerSerial
    LobbyServerHostingHandlerImpl_register = hostingHandlerSerial
End Function

Public Sub LobbyServerHostingHandlerImpl_unregister(ByVal serial As Long)
    On Error Resume Next
    hostingHandlers.Remove "h" & serial
    On Error GoTo 0
End Sub

' weakSelf.lock(): the handler, or Nothing once it is gone
Public Function LobbyServerHostingHandlerImpl_lock(ByVal serial As Long) As LobbyServerHostingHandlerImpl
    If hostingHandlers Is Nothing Then Exit Function
    On Error Resume Next
    Set LobbyServerHostingHandlerImpl_lock = hostingHandlers.Item("h" & serial)
    On Error GoTo 0
End Function

' std::shared_ptr<LobbyServerHostingHandlerProtocol> MakeLobbyHostListingHandler(lobbyAddress, callbacks)
' (callbacks: lobbyGameCreatedCallback(gameId), lobbyGamePublishedCallback(gameId, motd),
' lobbyGameGoneCallback(has_error, errCode, errMessage))
Public Function netlobby_MakeLobbyHostListingHandler(ByVal lobbyAddress As String, ByVal lobbyGameCreatedCallback As WzFunc, _
        ByVal lobbyGamePublishedCallback As WzFunc, ByVal lobbyGameGoneCallback As WzFunc) As LobbyServerHostingHandlerImpl
    Dim h As LobbyServerHostingHandlerImpl
    Set h = New LobbyServerHostingHandlerImpl
    h.construct lobbyAddress, netlobby_isTrustedLobbyAddress(lobbyAddress), lobbyGameCreatedCallback, lobbyGamePublishedCallback, lobbyGameGoneCallback
    Set netlobby_MakeLobbyHostListingHandler = h
End Function

'----------------------------------------------------------------------------------------------------------
' MARK: - Enumerating game listings

' bool EnumerateGames(const std::string& lobbyServerAddress, CompletionHandlerFunc completionHandlerFunc,
'   optional<size_t> resultsLimit, optional<NetcodeVer> matchingVersion)
' (completionHandlerFunc takes a netlobby_ListResult_store handle)
Public Function netlobby_EnumerateGames(ByVal lobbyServerAddress As String, ByVal completionHandlerFunc As WzFunc, _
        Optional ByVal has_resultsLimit As Boolean = False, Optional ByVal resultsLimit As Long = 0, _
        Optional ByVal has_matchingVersion As Boolean = False, Optional ByVal matchingVersion_major As Long = 0, _
        Optional ByVal matchingVersion_minor As Long = 0) As Boolean
    Dim request As URLDataRequest

    ' We always want direct connections to the lobby server (skipping any locally-configured proxies or redirects)
    request.noProxy = True
    request.noFollowRedirects = True

    request.userAgent = netlobby_getWZUserAgent()

    URLRequestBase_setRequestHeader request, "Origin", "wz2100://wz2100.net"
    URLRequestBase_setRequestHeader request, "Referer", "https://github.com/Warzone2100/warzone2100?wz_internal=list"
    URLRequestBase_setRequestHeader request, "Accept", "application/jsonl"

    request.maxDownloadSizeLimit = 20# * 1024 * 1024 ' response should never be > 20 MB

    request.url = buildLobbyRequestBaseUrl(lobbyServerAddress) & "/api/v1/games"

    request.connectTimeoutMs = 10000

    Set request.onResponse = WzFunc_new(FN_netlobby_enumGamesResponse, completionHandlerFunc, has_resultsLimit, resultsLimit, _
        has_matchingVersion, matchingVersion_major, matchingVersion_minor)

    Set request.onFailure = WzFunc_new(FN_netlobby_enumGamesFailure, completionHandlerFunc)

    netlobby_EnumerateGames = urlRequestData(request) <> 0
End Function

' (EnumerateGames' onResponse) [completionHandlerFunc, resultsLimit, matchingVersion](url, responseDetails, data)
Private Function enumGamesResponse(ByVal f As WzFunc, ByVal url As String, ByVal responseDetails As HTTPResponseDetails, ByVal data As MemoryStruct) As Double
    Dim httpStatusCode As Long, result As netlobby_ListResult, jsonData As nlohmann_json, b() As Byte
    Dim start As Long, end_ As Long, has_header As Boolean
    httpStatusCode = responseDetails.httpStatusCode()
    If httpStatusCode = 200 Then
        result.has_value = True
        ' Expected response is JSON-Lines of either header object or GameListing objects
        ' (processMemoryStructLines)
        b = data.memory
        start = 0
        Do
            end_ = -1
            If start < data.size Then end_ = findByte(b, start, data.size, 10)
            If end_ < 0 Then Exit Do
            If Not enumGamesLine(f, b, start, end_ - start, result.value, has_header) Then GoTo dispatch
            start = end_ + 1
        Loop
        ' Handle the last line if no trailing newline
        If start < data.size Then
            enumGamesLine f, b, start, data.size - start, result.value, has_header
        End If
dispatch:
        ' call the completionHandlerFunc on the main thread
        wzAsyncExecOnMainThread WzFunc_new(FN_netlobby_deliverListResult, f.cap(0), netlobby_ListResult_store(result))
    Else
        ' Parse the (expected) json LobbyError response
        Set jsonData = parseJsonOrNull(data)
        If jsonData.is_null() Then
            If httpStatusCode = 429 Then
                ' synthesize expected data for this specific code
                Dim rateLimited As netlobby_LobbyError
                netlobby_LobbyError_make rateLimited, "RATE_LIMITED", "Rate limit exceeded. Please try again later."
                Set jsonData = LobbyError_to_json(rateLimited)
            End If
        End If

        ' Unexpected / failure response code
        wzAsyncExecOnMainThread WzFunc_new(FN_netlobby_enumGamesError, f.cap(0), httpStatusCode, jsonData)
    End If
    enumGamesResponse = URLRequestHandlingBehavior_Done
End Function

' (VB) the index of the first byte value v in b(start .. end_ - 1), or -1
Private Function findByte(b() As Byte, ByVal start As Long, ByVal end_ As Long, ByVal v As Byte) As Long
    Dim i As Long
    For i = start To end_ - 1
        If b(i) = v Then
            findByte = i
            Exit Function
        End If
    Next
    findByte = -1
End Function

' (VB) the try block that parses a failure response: the JSON, or null (logged as the C++ logs it)
Private Function parseJsonOrNull(data As MemoryStruct) As nlohmann_json
    Set parseJsonOrNull = parseResponseJson(data)
End Function

' processMemoryStructLines' function for EnumerateGames: one line, b(start .. start + n - 1); False stops
Private Function enumGamesLine(ByVal f As WzFunc, b() As Byte, ByVal start As Long, ByVal n As Long, result As netlobby_ListResponse, has_header As Boolean) As Boolean
    Dim jsonData As nlohmann_json, it As nlohmann_json, listing As netlobby_GameListing, what As String
    If Not tryParseLine(b, start, n, jsonData, what) Then
        netlobby_logErrorOnMainThread "Failed to parse JSON response: ", what
        enumGamesLine = False
        Exit Function
    End If
    If Not (jsonData.is_object()) Then ASSERT_FAILURE "jsonData.is_object()", "lib/netplay/netlobby.cpp:2151", "EnumerateGames", "Received non-object item": Exit Function

    Set it = jsonData.find("type")
    If Not it Is Nothing Then
        If isHeaderType(it) Then
            ' Found header object
            If Not has_header Then
                ' Try to parse as EnumGamesHeader
                If Not tryGetHeaderMotd(jsonData, result.lobbyMOTD, what) Then
                    ' failed to convert json header object
                    debug_ 2165, LOG_NEVER, "EnumerateGames", "%s", jsonData.dump(-1, " ", False, error_handler_t_replace)
                    netlobby_logErrorOnMainThread "Invalid header: ", what
                    enumGamesLine = True    ' continue enumeration
                    Exit Function
                End If
                has_header = True
            End If
            enumGamesLine = True
            Exit Function
        End If
    End If

    ' Try to parse as GameListing
    If Not tryGetGameListing(jsonData, listing, what) Then
        ' failed to convert json response
        debug_ 2182, LOG_NEVER, "EnumerateGames", "%s", jsonData.dump(-1, " ", False, error_handler_t_replace)
        netlobby_logErrorOnMainThread "Invalid listing: ", what
        enumGamesLine = True        ' continue enumeration
        Exit Function
    End If

    If Not f.cap(3) Or (f.cap(4) = listing.details.netcodeVer.major And f.cap(5) = listing.details.netcodeVer.minor) Then
        ReDim Preserve result.games(0 To result.gamesCount)
        result.games(result.gamesCount) = listing
        result.gamesCount = result.gamesCount + 1
    End If

    If f.cap(1) Then
        enumGamesLine = ULT(result.gamesCount, f.cap(2))
        Exit Function
    End If
    enumGamesLine = True            ' continue enumeration
End Function

' (it.value().get<std::string>() == "header": a non-string type throws, which ends the enumeration's callback
' as an uncaught exception would end the transfer thread's handler: treated as not a header)
Private Function isHeaderType(it As nlohmann_json) As Boolean
    If it.is_string() Then isHeaderType = (it.get_string() = "header")
End Function

Private Function tryParseLine(b() As Byte, ByVal start As Long, ByVal n As Long, jsonData As nlohmann_json, what As String) As Boolean
    On Error GoTo failed
    Set jsonData = json_parse_bytes(b, start, n)
    tryParseLine = True
    Exit Function
failed:
    what = Err.Description
    tryParseLine = False
End Function

' EnumGamesHeader from_json (v.motd = j.at("motd"))
Private Function tryGetHeaderMotd(j As nlohmann_json, motd As String, what As String) As Boolean
    Dim s As String
    On Error GoTo failed
    s = j.at("motd").get_string()
    motd = s
    tryGetHeaderMotd = True
    Exit Function
failed:
    what = Err.Description
End Function

' MARK: GameListing
Private Function tryGetGameListing(j As nlohmann_json, v As netlobby_GameListing, what As String) As Boolean
    Dim listing As netlobby_GameListing, conn As nlohmann_json, i As Long
    On Error GoTo failed
    listing.gameId = j.at("id").get_string()
    netlobby_GameDetails_from_json j, listing.details
    Set conn = j.at("conn")
    If Not conn.is_array() Then throw_type_error_302 "type must be array, but is " & conn.type_name()
    If conn.size() > 0 Then ReDim listing.availableConnectionTypes(0 To conn.size() - 1)
    For i = 0 To conn.size() - 1
        ConnectionType_from_json conn.at_idx(i), listing.availableConnectionTypes(i)
    Next
    listing.availableConnectionTypesCount = conn.size()
    v = listing
    tryGetGameListing = True
    Exit Function
failed:
    what = Err.Description
End Function

'----------------------------------------------------------------------------------------------------------
' MARK: - Requesting join details

' void PendingJoinRequestData::HandleResponseSuccess(IPVersion ver, nlohmann::json &&jsonData)
Private Sub PendingJoinRequestData_HandleResponseSuccess(ByVal id As Long, ByVal ver As netlobby_IPVersion, jsonData As nlohmann_json)
    Dim response As RequestJoinResponse, i As Long, isNowReady As Boolean, e As netlobby_LobbyError
    If Not pendingJoins(id).received(ver) Then
        pendingJoins(id).received(ver) = True

        If tryGetRequestJoinResponse(jsonData, response) Then
            ' Add to pendingResults
            For i = 0 To response.connectionsCount - 1
                appendJoinConnection pendingJoins(id).pendingResults.joinDetails, response.connections(i), response.joinToken
            Next
        Else
            ' Failed to parse response as RequestJoinResponse - Record as error
            netlobby_LobbyError_make e, "RESPONSE_PARSE_ERROR", "Invalid or unrecognized response"
            setJoinError id, ver, True, 0, "", e
        End If
    Else
        ' ERROR - received twice for same ipversion?!?
    End If
    isNowReady = PendingJoinRequestData_isReady(id)
    If isNowReady Then PendingJoinRequestData_DispatchResultsIfReady id
End Sub

' joinDetails.connections.push_back(conn); joinDetails.joinTokens.push_back(joinToken)
Private Sub appendJoinConnection(d As netlobby_GameJoinDetails, conn As netlobby_ConnectionInfo, ByVal joinToken As String)
    ReDim Preserve d.connections(0 To d.connectionsCount)
    d.connections(d.connectionsCount) = conn
    d.connectionsCount = d.connectionsCount + 1
    ReDim Preserve d.joinTokens(0 To d.joinTokensCount)
    d.joinTokens(d.joinTokensCount) = joinToken
    d.joinTokensCount = d.joinTokensCount + 1
End Sub

Private Function tryGetRequestJoinResponse(j As nlohmann_json, v As RequestJoinResponse) As Boolean
    Dim conns As nlohmann_json, i As Long, r As RequestJoinResponse
    On Error GoTo failed
    If j Is Nothing Then GoTo failed
    ' RequestJoinResponse from_json: connections (ConnectionInfo each), joinToken
    Set conns = j.at("connections")
    If Not conns.is_array() Then throw_type_error_302 "type must be array, but is " & conns.type_name()
    If conns.size() > 0 Then ReDim r.connections(0 To conns.size() - 1)
    For i = 0 To conns.size() - 1
        netlobby_ConnectionInfo_from_json conns.at_idx(i), r.connections(i)
    Next
    r.connectionsCount = conns.size()
    r.joinToken = j.at("joinToken").get_string()
    v = r
    tryGetRequestJoinResponse = True
    Exit Function
failed:
    tryGetRequestJoinResponse = False
End Function

' pendingResults.errors[ver] = ... (a ConnectivityFailure or a LobbyError)
Private Sub setJoinError(ByVal id As Long, ByVal ver As netlobby_IPVersion, ByVal isLobbyError As Boolean, ByVal connectivityType As netlobby_ConnectivityFailureType, _
        ByVal connectivityDetails As String, lobbyError As netlobby_LobbyError)
    setResultsError pendingJoins(id).pendingResults, ver, isLobbyError, connectivityType, connectivityDetails, lobbyError
End Sub

Private Sub setResultsError(r As netlobby_RequestJoinDetailsResults, ByVal ver As netlobby_IPVersion, ByVal isLobbyError As Boolean, _
        ByVal connectivityType As netlobby_ConnectivityFailureType, ByVal connectivityDetails As String, lobbyError As netlobby_LobbyError)
    Dim i As Long, empty_ As netlobby_LobbyError
    For i = 0 To r.errorsCount - 1
        If r.errors(i).first = ver Then Exit For
    Next
    If i = r.errorsCount Then
        ReDim Preserve r.errors(0 To r.errorsCount)
        r.errorsCount = r.errorsCount + 1
    End If
    r.errors(i).first = ver
    r.errors(i).isLobbyError = isLobbyError
    r.errors(i).connectivityType = connectivityType
    r.errors(i).connectivityDetails = connectivityDetails
    If isLobbyError Then r.errors(i).lobbyError = lobbyError Else r.errors(i).lobbyError = empty_
End Sub

' void PendingJoinRequestData::HandleResponseError(IPVersion ver, long httpStatusCode, nlohmann::json &&jsonData)
Private Sub PendingJoinRequestData_HandleResponseError(ByVal id As Long, ByVal ver As netlobby_IPVersion, ByVal httpStatusCode As Long, jsonData As nlohmann_json)
    Dim err_ As netlobby_LobbyError, isNowReady As Boolean
    If Not pendingJoins(id).received(ver) Then
        pendingJoins(id).received(ver) = True

        ' Try to decode LobbyError json
        If tryGetLobbyError(jsonData, err_) Then
            setJoinError id, ver, True, 0, "", err_
        Else
            netlobby_LobbyError_make err_, "GENERIC_ERROR", va("Status Code: %ld", httpStatusCode)
            setJoinError id, ver, True, 0, "", err_
        End If
    End If
    isNowReady = PendingJoinRequestData_isReady(id)
    If isNowReady Then PendingJoinRequestData_DispatchResultsIfReady id
End Sub

Private Function JoinConnectivityFailureFromURLRequestFailure(ByVal failureType As URLRequestFailureType) As netlobby_ConnectivityFailureType
    Select Case failureType
    Case URLRequestFailureType_OPERATION_TIMEOUT
        JoinConnectivityFailureFromURLRequestFailure = netlobby_ConnectivityFailureType_ERR_TIMEOUT
        Exit Function
    Case URLRequestFailureType_COULDNT_CONNECT
        JoinConnectivityFailureFromURLRequestFailure = netlobby_ConnectivityFailureType_ERR_COULD_NOT_CONNECT
        Exit Function
    End Select
    JoinConnectivityFailureFromURLRequestFailure = netlobby_ConnectivityFailureType_ERR_OTHER
End Function

' void PendingJoinRequestData::HandleFailure(IPVersion ver, URLRequestFailureType failureType, const std::string& internalErrorReason)
Private Sub PendingJoinRequestData_HandleFailure(ByVal id As Long, ByVal ver As netlobby_IPVersion, ByVal failureType As URLRequestFailureType, ByVal internalErrorReason As String)
    Dim isNowReady As Boolean, none As netlobby_LobbyError
    If Not pendingJoins(id).received(ver) Then
        pendingJoins(id).received(ver) = True

        setJoinError id, ver, False, JoinConnectivityFailureFromURLRequestFailure(failureType), internalErrorReason, none
    End If
    isNowReady = PendingJoinRequestData_isReady(id)
    If isNowReady Then PendingJoinRequestData_DispatchResultsIfReady id
End Sub

Private Function PendingJoinRequestData_isReady(ByVal id As Long) As Boolean
    Dim v As Long
    For v = 1 To 2
        If pendingJoins(id).requested(v) And Not pendingJoins(id).received(v) Then Exit Function
    Next
    PendingJoinRequestData_isReady = True
End Function

Private Sub PendingJoinRequestData_DispatchResultsIfReady(ByVal id As Long)
    Dim readyResults As netlobby_JoinResult, fresh As netlobby_RequestJoinDetailsResults, resultHandler As WzFunc
    If Not pendingJoins(id).used Then Exit Sub
    If PendingJoinRequestData_isReady(id) Then
        readyResults.has_value = True
        readyResults.value = pendingJoins(id).pendingResults
        pendingJoins(id).pendingResults = fresh
        Set resultHandler = pendingJoins(id).resultHandler

        ' Dispatch results to resultHandler on main thread
        wzAsyncExecOnMainThread WzFunc_new(FN_netlobby_deliverJoinResult, resultHandler, netlobby_JoinResult_store(readyResults))

        ' (the last shared_ptr reference goes with the requests' handlers: the slot is freed)
        freePendingJoin id
    End If
End Sub

Private Function allocPendingJoin() As Long
    Dim i As Long
    For i = 0 To pendingJoinsCount - 1
        If Not pendingJoins(i).used Then
            allocPendingJoin = i
            Exit Function
        End If
    Next
    ReDim Preserve pendingJoins(0 To pendingJoinsCount)
    allocPendingJoin = pendingJoinsCount
    pendingJoinsCount = pendingJoinsCount + 1
End Function

Private Sub freePendingJoin(ByVal id As Long)
    Dim empty_ As PendingJoinRequestData
    pendingJoins(id) = empty_
End Sub

' bool RequestJoinDetails(const std::string& lobbyServerAddress, const std::string& lobbyGameId, const WzString& playerName,
'   const EcKey& playerIdentity, bool asSpectator, RequestJoinResultHandlerFunc resultHandler, optional<IPVersion>
'   optIpVersion, std::shared_ptr<LobbyErrorResolutionData> errorResolution)
' (resultHandler takes a netlobby_JoinResult_store handle; errorResolution: the resolution's nlohmann_json or Nothing)
Public Function netlobby_RequestJoinDetails(ByVal lobbyServerAddress As String, ByVal lobbyGameId As String, ByVal playerName As String, _
        playerIdentity As EcKey, ByVal asSpectator As Boolean, ByVal resultHandler As WzFunc, _
        Optional ByVal has_optIpVersion As Boolean = False, Optional ByVal optIpVersion As netlobby_IPVersion = netlobby_IPVersion_IPv4, _
        Optional ByVal errorResolution As Object = Nothing) As Boolean
    Dim baseRequest As URLDataRequest, request As URLDataRequest, requestBody As nlohmann_json, attestation As nlohmann_json
    Dim requestIPVersions(0 To 1) As netlobby_IPVersion, requestIPVersionsCount As Long, sharedPendingData As Long
    Dim requestInitFailures As Long, k As Long, ver As netlobby_IPVersion, errorResolutionJsonData As nlohmann_json

    ' We always want direct connections to the lobby server (skipping any locally-configured proxies or redirects)
    baseRequest.noProxy = True
    baseRequest.noFollowRedirects = True

    baseRequest.userAgent = netlobby_getWZUserAgent()

    URLRequestBase_setRequestHeader baseRequest, "Origin", "wz2100://wz2100.net"
    URLRequestBase_setRequestHeader baseRequest, "Referer", "https://github.com/Warzone2100/warzone2100?wz_internal=request.join"
    URLRequestBase_setRequestHeader baseRequest, "Accept", "application/json"

    baseRequest.maxDownloadSizeLimit = 2# * 1024 * 1024 ' response should never be > 2 MB

    baseRequest.url = buildLobbyRequestBaseUrl(lobbyServerAddress) & "/api/v1/game/" & lobbyGameId & "/request-join"

    baseRequest.connectTimeoutMs = 10000

    baseRequest.autoRetryStrategy = URLRequestAutoRetryStrategy_RetryAfter(2000, 1, Array(HttpStatusTooManyRequests))

    ' construct request body, and setPost
    Set requestBody = json_object(True)
    requestBody.set_item_ref "asSpec", json_value(asSpectator, True)
    If netlobby_isTrustedLobbyAddress(lobbyServerAddress) Then
        Set attestation = netlobby_makeAttestationJson(playerName, playerIdentity, lobbyServerAddress, True, lobbyGameId)
        If Not (attestation.is_object()) Then ASSERT_FAILURE "attestation.is_object()", "lib/netplay/netlobby.cpp:2508", "RequestJoinDetails", "Not an object?": Exit Function
        requestBody.set_item_ref "attest", attestation
    End If
    If Not errorResolution Is Nothing Then
        ' Get errorResolution data json and append to sent body
        If TypeOf errorResolution Is nlohmann_json Then
            Set errorResolutionJsonData = errorResolution
            requestBody.set_item "captcha", errorResolutionJsonData
        Else
            debug_ 2521, LOG_LOBBY, "RequestJoinDetails", "errorResolution was unexpected type?"
        End If
    End If
    URLDataRequest_setPost baseRequest, requestBody.dump(-1, " ", False, error_handler_t_replace)

    ' Determine how many requests to dispatch, depending on optIpVersion
    If Not has_optIpVersion Then
        ' Request both IPv4 and IPv6 results
        requestIPVersions(0) = netlobby_IPVersion_IPv4
        requestIPVersions(1) = netlobby_IPVersion_IPv6
        requestIPVersionsCount = 2
    Else
        requestIPVersions(0) = optIpVersion
        requestIPVersionsCount = 1
    End If

    ' auto sharedPendingData = std::make_shared<PendingJoinRequestData>(requestIPVersions, resultHandler)
    sharedPendingData = allocPendingJoin()
    pendingJoins(sharedPendingData).used = True
    Set pendingJoins(sharedPendingData).resultHandler = resultHandler
    For k = 0 To requestIPVersionsCount - 1
        pendingJoins(sharedPendingData).requested(requestIPVersions(k)) = True
    Next

    requestInitFailures = 0
    For k = 0 To requestIPVersionsCount - 1
        ver = requestIPVersions(k)
        request = baseRequest
        request.protocol = netlobby_ipVersionToInternetProtocol(ver)

        Set request.onResponse = WzFunc_new(FN_netlobby_joinResponse, sharedPendingData, ver)
        Set request.onFailure = WzFunc_new(FN_netlobby_joinFailure, sharedPendingData, ver)

        If urlRequestData(request) = 0 Then
            ' Failed to dispatch request!
            requestInitFailures = requestInitFailures + 1
            PendingJoinRequestData_HandleFailure sharedPendingData, ver, URLRequestFailureType_TRANSFER_FAILED, "Unable to dispatch request"
        End If
    Next

    netlobby_RequestJoinDetails = requestInitFailures < requestIPVersionsCount
End Function

'----------------------------------------------------------------------------------------------------------
' The lambdas

Public Function netlobby_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim r As netlobby_ListResult, e As netlobby_LobbyError, jsonData As nlohmann_json, httpStatusCode As Long
    Dim strongSelf As LobbyServerHostingHandlerImpl, what As String, data As MemoryStruct, details As HTTPResponseDetails
    Dim checkResult As netlobby_CheckJoinRequestResult
    Select Case f.id
    Case FN_netlobby_enumGamesResponse              ' EnumerateGames' onResponse(url, responseDetails, data)
        netlobby_c_fn = enumGamesResponse(f, args(0), args(1), args(2))

    Case FN_netlobby_deliverListResult              ' [completionHandlerFunc, result = std::move(result)]()
        WzFunc_call f.cap(0), f.cap(1)

    Case FN_netlobby_enumGamesError                 ' [completionHandlerFunc, httpStatusCode, jsonData = std::move(jsonData)]()
        Set jsonData = f.cap(2)
        httpStatusCode = f.cap(1)
        If Not tryGetLobbyError(jsonData, e) Then
            ' Failed to parse response as LobbyError - non-fatal
            ' (might not have sent a response besides the httpStatusCode)
            If httpStatusCode = 429 Then
                ' synthesize LobbyError for this specific code
                netlobby_LobbyError_make e, "RATE_LIMITED", "Rate limit exceeded. Please try again later."
                r.error_ = e
            Else
                ' Log + call resultHandler
                debug_ 2246, LOG_LOBBY, "EnumerateGames", "Lobby error: (%ld)", httpStatusCode
                netlobby_LobbyError_make r.error_, "GENERIC_ERROR", va("Status Code: %ld", httpStatusCode)
                r.has_value = False
                WzFunc_call f.cap(0), netlobby_ListResult_store(r)
                Exit Function
            End If
        Else
            r.error_ = e
        End If
        ' Log + call resultHandler
        debug_ 2241, LOG_LOBBY, "EnumerateGames", "Lobby returned error: (%ld) [%s] %s", httpStatusCode, r.error_.errCode, r.error_.errMessage
        r.has_value = False
        WzFunc_call f.cap(0), netlobby_ListResult_store(r)

    Case FN_netlobby_enumGamesFailure               ' [completionHandlerFunc](url, type, transferDetails)
        ' Failure response - call resultHandler on the main thread
        httpStatusCode = 0
        If Not args(2) Is Nothing Then
            Set details = args(2)
            httpStatusCode = details.httpStatusCode()
        End If
        wzAsyncExecOnMainThread WzFunc_new(FN_netlobby_enumGamesConnError, f.cap(0), args(1), httpStatusCode)

    Case FN_netlobby_enumGamesConnError             ' [completionHandlerFunc, type, httpStatusCode]()
        debug_ 2264, LOG_LOBBY, "EnumerateGames", "Lobby connection error: %s (%ld)", to_string_URLRequestFailureType(f.cap(1)), f.cap(2)
        r.has_value = False
        netlobby_LobbyError_make r.error_, "SERVER_CONNECTION_FAILURE", va("Status Code: %ld", f.cap(2))
        WzFunc_call f.cap(0), netlobby_ListResult_store(r)

    Case FN_netlobby_joinResponse                   ' [sharedPendingData, ipVersion](url, responseDetails, data)
        Set details = args(1)
        Set data = args(2)
        httpStatusCode = details.httpStatusCode()

        ' Parse the (expected) json response
        Set jsonData = parseResponseJson(data)

        If httpStatusCode = 200 Then
            PendingJoinRequestData_HandleResponseSuccess f.cap(0), f.cap(1), jsonData
        Else
            PendingJoinRequestData_HandleResponseError f.cap(0), f.cap(1), httpStatusCode, jsonData
        End If
        netlobby_c_fn = URLRequestHandlingBehavior_Done

    Case FN_netlobby_joinFailure                    ' [sharedPendingData, ipVersion](url, type, transferDetails)
        what = ""
        If Not args(2) Is Nothing Then
            Set details = args(2)
            what = details.getInternalResultDescription()
        End If
        PendingJoinRequestData_HandleFailure f.cap(0), f.cap(1), args(1), what

    Case FN_netlobby_deliverJoinResult              ' [readyResults, resultHandlerCopy = resultHandler]()
        WzFunc_call f.cap(0), f.cap(1)

    Case FN_netlobby_logOnMainThread                ' [errorString = std::move(errorString)]()
        debug_ 1289, LOG_LOBBY, "logErrorOnMainThread", "%s", f.cap(0)

    Case FN_netlobby_captchaResult                  ' [resultHandler, token = verifyPrompt.token](WzString result)
        WzFunc_call f.cap(0), CaptchaVerifyResponse_to_json(f.cap(1), args(0))

    Case FN_netlobby_slideResult                    ' [resultHandler, token = verifyPrompt.token](uint32_t pointX, uint32_t pointY)
        WzFunc_call f.cap(0), SlideVerifyResponse_to_json(f.cap(1), args(0), args(1))

    Case FN_netlobby_hostingAPIResponse, FN_netlobby_hostingAPIHandle, FN_netlobby_hostingAPIFailureMain, FN_netlobby_hostingAPIStdFailure, _
         FN_netlobby_createGameHandler, FN_netlobby_updateGameHandler, FN_netlobby_addConnectionHandler, FN_netlobby_addConnectionFailure
        WzFunc_assign netlobby_c_fn, hostingHandlerLambda(f, args)

    Case FN_netlobby_netplayGameCreated, FN_netlobby_netplayGamePublished, FN_netlobby_netplayGameGone, FN_netlobby_netplayCheckJoinResult
        netplayLambda f, args

    Case FN_netlobby_checkJoinResponse              ' [resultHandler](url, responseDetails, data)
        Set details = args(1)
        Set data = args(2)
        ' Parse the (expected) json response
        Set jsonData = parseResponseJson(data)

        httpStatusCode = details.httpStatusCode()
        If httpStatusCode = 200 Then
            ' On accepted response code(s) (which may vary by api)
            ' call the resultHandler on the main thread
            wzAsyncExecOnMainThread WzFunc_new(FN_netlobby_checkJoinResult, f.cap(0), jsonData)
        Else
            ' Unexpected / failure response code
            wzAsyncExecOnMainThread WzFunc_new(FN_netlobby_checkJoinError, f.cap(0), httpStatusCode, jsonData)
        End If
        netlobby_c_fn = URLRequestHandlingBehavior_Done

    Case FN_netlobby_checkJoinResult                ' [resultHandler, jsonData = std::move(jsonData)]()
        Set jsonData = f.cap(1)
        If Not tryGetCheckJoinRequestResult(jsonData, checkResult, what) Then
            ' failed to convert json response
            WzFunc_call f.cap(0), False, False, "WZ_INTERNAL_INVALID_SERVER_RESPONSE", what
            Exit Function
        End If
        WzFunc_call f.cap(0), True, checkResult.pass, "", checkResult.error_

    Case FN_netlobby_checkJoinError                 ' [resultHandler, httpStatusCode, jsonData = std::move(jsonData)]()
        Set jsonData = f.cap(2)
        httpStatusCode = f.cap(1)
        ' Log + call resultHandler
        If tryGetLobbyError(jsonData, e) Then
            debug_ 1727, LOG_LOBBY, "checkJoinRequest", "Lobby returned error: (%ld) [%s] %s", httpStatusCode, e.errCode, e.errMessage
            WzFunc_call f.cap(0), False, False, e.errCode, e.errMessage
        Else
            debug_ 1732, LOG_LOBBY, "checkJoinRequest", "Lobby error: (%ld)", httpStatusCode
            WzFunc_call f.cap(0), False, False, "GENERIC_ERROR", va("Status Code: %ld", httpStatusCode)
        End If

    Case FN_netlobby_checkJoinFailure               ' [resultHandler](url, type, transferDetails)
        ' Failure response - call resultHandler on the main thread
        httpStatusCode = 0
        If Not args(2) Is Nothing Then
            Set details = args(2)
            httpStatusCode = details.httpStatusCode()
        End If
        wzAsyncExecOnMainThread WzFunc_new(FN_netlobby_checkJoinConnError, f.cap(0), httpStatusCode)

    Case FN_netlobby_checkJoinConnError             ' [resultHandler, httpStatusCode]()
        WzFunc_call f.cap(0), False, False, "SERVER_CONNECTION_FAILURE", va("Status Code: %ld", f.cap(1))
    End Select
End Function

Private Function tryGetCheckJoinRequestResult(j As nlohmann_json, v As netlobby_CheckJoinRequestResult, what As String) As Boolean
    Dim r As netlobby_CheckJoinRequestResult
    On Error GoTo failed
    CheckJoinRequestResult_from_json j, r
    v = r
    tryGetCheckJoinRequestResult = True
    Exit Function
failed:
    what = Err.Description
End Function

'----------------------------------------------------------------------------------------------------------
' (VB) the lambdas of LobbyServerHostingHandlerImpl (none has an error handler: a handlerFunc's exception
' reaches callHostingHandlerFunc's, the try block of buildHostingAPIResponseHandler's main-thread lambda)

Private Function hostingHandlerLambda(ByVal f As WzFunc, args() As Variant) As Variant
    Dim strongSelf As LobbyServerHostingHandlerImpl, details As HTTPResponseDetails, data As MemoryStruct
    Dim jsonData As nlohmann_json, httpStatusCode As Long, what As String, resp As netlobby_WrappedCreateGameResponse
    Dim hostGameStatus As netlobby_HostGameStatus
    Select Case f.id
    Case FN_netlobby_hostingAPIResponse             ' [weakSelf, acceptedResponseCodes, handlerFunc](url, responseDetails, data)
        Set details = args(1)
        Set data = args(2)
        ' Parse the (expected) json response
        Set jsonData = parseResponseJson(data)

        httpStatusCode = details.httpStatusCode()
        If InStr(f.cap(1), "," & CStr(httpStatusCode) & ",") <> 0 Then
            ' On accepted response code(s) (which may vary by api)
            ' call the handlerFunc on the main thread
            wzAsyncExecOnMainThread WzFunc_new(FN_netlobby_hostingAPIHandle, f.cap(0), f.cap(2), args(0), httpStatusCode, jsonData)
        Else
            ' Failure response code - call handleAPIFailure on the main thread
            wzAsyncExecOnMainThread WzFunc_new(FN_netlobby_hostingAPIFailureMain, f.cap(0), args(0), httpStatusCode, jsonData, False)
        End If
        hostingHandlerLambda = URLRequestHandlingBehavior_Done

    Case FN_netlobby_hostingAPIHandle               ' [weakSelf, handlerFunc, urlCopy, httpStatusCode, jsonData = std::move(jsonData)]() mutable
        If Not callHostingHandlerFunc(f.cap(1), f.cap(2), f.cap(3), f.cap(4), what) Then
            ' handlerFunc caused an exception - possibly trying to convert json
            ' treat this as a failure
            Set strongSelf = LobbyServerHostingHandlerImpl_lock(f.cap(0))
            If strongSelf Is Nothing Then
                debug_ 1338, LOG_LOBBY, "buildHostingAPIResponseHandler", "LobbyServerHostingHandler instance is gone"
                Exit Function
            End If
            strongSelf.handleAPIResponseHandlingFailure f.cap(2), f.cap(3), what
        End If

    Case FN_netlobby_hostingAPIFailureMain          ' [weakSelf, urlCopy, httpStatusCode, jsonData]() (cap(4): the standard failure handler's)
        Set strongSelf = LobbyServerHostingHandlerImpl_lock(f.cap(0))
        If strongSelf Is Nothing Then
            If f.cap(4) Then
                debug_ 1375, LOG_LOBBY, "buildHostingAPIStandardFailureHandler", "LobbyServerHostingHandler instance is gone"
            Else
                ASSERT_FAILURE "strongSelf != nullptr", "lib/netplay/netlobby.cpp:1350", "buildHostingAPIResponseHandler", "LobbyServerHostingHandler instance is gone"
            End If
            Exit Function
        End If
        Set jsonData = f.cap(3)
        strongSelf.handleAPIFailure f.cap(1), f.cap(2), jsonData

    Case FN_netlobby_hostingAPIStdFailure           ' [weakSelf](url, type, transferDetails)
        ' Failure response - call handleAPIFailure on the main thread
        httpStatusCode = 0
        If Not args(2) Is Nothing Then
            Set details = args(2)
            httpStatusCode = details.httpStatusCode()
        End If
        wzAsyncExecOnMainThread WzFunc_new(FN_netlobby_hostingAPIFailureMain, f.cap(0), args(0), httpStatusCode, jsonNull(), True)

    Case FN_netlobby_createGameHandler              ' [weakSelf, connections, lobbyServerAddressCopy](url, httpStatusCode, jsonData)
        ' handle only accepted response codes (called on main thread)
        Set jsonData = args(2)
        netlobby_WrappedCreateGameResponse_from_json jsonData, resp

        Set strongSelf = LobbyServerHostingHandlerImpl_lock(f.cap(0))
        If strongSelf Is Nothing Then
            debug_ 1478, LOG_LOBBY, "createGameListing", "LobbyServerHostingHandler instance is gone"
            ' Immediately call DELETE API on gameId
            netlobby_LobbyDeleteGameImpl f.cap(1), resp.ctx, False
            Exit Function
        End If

        strongSelf.eventGameCreated resp

    Case FN_netlobby_updateGameHandler              ' [weakSelf, lobbyServerAddressCopy, ownedPendingUpdate](url, httpStatusCode, jsonData)
        ' handle only accepted response codes (called on main thread)
        Set strongSelf = LobbyServerHostingHandlerImpl_lock(f.cap(0))
        If strongSelf Is Nothing Then
            debug_ 1610, LOG_LOBBY, "sendListingUpdateImpl", "LobbyServerHostingHandler instance is gone"
            Exit Function
        End If
        Set jsonData = args(2)
        strongSelf.listingUpdateResponse f.cap(2), args(1), jsonData

    Case FN_netlobby_addConnectionHandler           ' [weakSelf, connections](url, httpStatusCode, jsonData)
        ' handle only accepted response codes (called on main thread)
        httpStatusCode = args(1)
        If httpStatusCode = 406 Or httpStatusCode = 409 Or httpStatusCode = 429 Then  ' 409 Conflict or 429 Too Many Requests
            ' 409: Tried to add a duplicate of the same connection type?
            ' 429: Lobby is rate-limiting adding this connection (and autoRetryStrategy didn't succeed)
            ' (Failing to add a single connection doesn't necessarily mean the whole listing has failed)
            ' Just log it
            debug_ 1915, LOG_LOBBY, "addConnectionsToGame", "Received status code: %ld", httpStatusCode
            Exit Function
        End If

        ' Waiting on the server to test the connection here - status is likely pending

        ' Get HostGameStatus from jsonData
        Set jsonData = args(2)
        netlobby_HostGameStatus_from_json jsonData, hostGameStatus
        Set strongSelf = LobbyServerHostingHandlerImpl_lock(f.cap(0))
        If strongSelf Is Nothing Then
            debug_ 1926, LOG_LOBBY, "addConnectionsToGame", "LobbyServerHostingHandler instance is gone"
            Exit Function
        End If
        strongSelf.handleHostGameStatus hostGameStatus

    Case FN_netlobby_addConnectionFailure           ' [protocol](url, type, transferDetails)
        ' Failing to add a single connection doesn't necessarily mean the whole listing has failed
        ' Just log (on main thread)
        netlobby_logErrorOnMainThread va("AddConnection request failure (protocol: %s, failureType: %s)", to_string_InternetProtocol(f.cap(0)), to_string_URLRequestFailureType(args(1))), ""
    End Select
End Function

' try { handlerFunc(urlCopy, httpStatusCode, std::move(jsonData)); } catch (const std::exception &e): False with e.what()
Private Function callHostingHandlerFunc(ByVal handlerFunc As WzFunc, ByVal url As String, ByVal httpStatusCode As Long, ByVal jsonData As nlohmann_json, what As String) As Boolean
    On Error GoTo failed
    WzFunc_call handlerFunc, url, httpStatusCode, jsonData
    callHostingHandlerFunc = True
    Exit Function
failed:
    what = Err.Description
    callHostingHandlerFunc = False
End Function

'----------------------------------------------------------------------------------------------------------
' (VB) netplay.cpp's static lobbyHostingHandler, which WzNet's netlobby_c reaches through INetGame (NetGame.cls's
' lobbyHost* methods), and the eventCallbacks NETcreateLobbyHostingHandler gives it (WzNet's lambdas, reached
' through NetApi: NETlobbyGameCreatedCallback, NETlobbyGamePublishedCallback, NETlobbyGameGoneCallback). The
' host's identity is getMultiStats(player).identity, taken here (the private key stays in the exe); the game
' details cross as JSON text in netlobby.cpp's to_json(GameDetails) format, the connections as a JSON array of
' to_json(ConnectionInfo).


' lobbyHostingHandler = netlobby::MakeLobbyHostListingHandler(NETgetLobbyserverAddress(), eventCallbacks)
Public Sub netlobby_netplayHostCreate(ByVal lobbyServerAddress As String)
    netlobby_netplayHostReset
    Set netplayLobbyHostingHandler = netlobby_MakeLobbyHostListingHandler(lobbyServerAddress, WzFunc_new(FN_netlobby_netplayGameCreated), _
        WzFunc_new(FN_netlobby_netplayGamePublished), WzFunc_new(FN_netlobby_netplayGameGone))
End Sub

' lobbyHostingHandler.reset()
Public Sub netlobby_netplayHostReset()
    Dim h As LobbyServerHostingHandlerImpl
    If netplayLobbyHostingHandler Is Nothing Then Exit Sub
    Set h = netplayLobbyHostingHandler
    Set netplayLobbyHostingHandler = Nothing
    h.destruct
End Sub

' (VB) the crossing's game details and join options
Private Function netplayGameDetails(ByVal gameDetailsJson As String, gd As netlobby_GameDetails) As Boolean
    On Error GoTo failed
    netlobby_GameDetails_from_json json_parse(gameDetailsJson), gd
    netplayGameDetails = True
    Exit Function
failed:
    ASSERT_FAILURE "false", "lib/netplay/netplay.cpp", "netplayGameDetails", "(VB) invalid game details from WzNet: %s", Err.Description
End Function

Private Sub netplayJoinOptions(ByVal botProt As Long, ByVal proxyIPs As Long, ByVal hostingIPs As Long, opts As netlobby_HostJoinOptions)
    opts.botProt = botProt
    opts.proxyIPs = proxyIPs
    opts.hostingIPs = hostingIPs
End Sub

Private Function netplayConnections(ByVal connectionsJson As String, conns() As netlobby_ConnectionInfo) As Long
    Dim j As nlohmann_json, i As Long
    On Error GoTo failed
    Set j = json_parse(connectionsJson)
    If j.size() > 0 Then ReDim conns(0 To j.size() - 1)
    For i = 0 To j.size() - 1
        netlobby_ConnectionInfo_from_json j.at_idx(i), conns(i)
    Next
    netplayConnections = j.size()
    Exit Function
failed:
    ASSERT_FAILURE "false", "lib/netplay/netplay.cpp", "netplayConnections", "(VB) invalid connections from WzNet: %s", Err.Description
    netplayConnections = 0
End Function

Private Function netplayIdentity(ByVal player As Long) As EcKey
    Dim stats As PLAYERSTATS_
    stats = getMultiStats(player)
    netplayIdentity = stats.identity
End Function

' lobbyHostingHandler->createGameListing(getMultiStats(realSelectedPlayer).identity, gameDetails, hostJoinOptions, buildHostingConnectionDescriptions())
Public Function netlobby_netplayHostCreateGameListing(ByVal identityPlayer As Long, ByVal gameDetailsJson As String, ByVal botProt As Long, _
        ByVal proxyIPs As Long, ByVal hostingIPs As Long, ByVal connectionsJson As String) As Boolean
    Dim identity As EcKey, gd As netlobby_GameDetails, opts As netlobby_HostJoinOptions, conns() As netlobby_ConnectionInfo, n As Long
    If netplayLobbyHostingHandler Is Nothing Then Exit Function
    identity = netplayIdentity(identityPlayer)
    If Not netplayGameDetails(gameDetailsJson, gd) Then Exit Function
    netplayJoinOptions botProt, proxyIPs, hostingIPs, opts
    n = netplayConnections(connectionsJson, conns)
    netlobby_netplayHostCreateGameListing = netplayLobbyHostingHandler.createGameListing(identity, gd, opts, conns, n)
End Function

' lobbyHostingHandler->updateGameListing(getMultiStats(realSelectedPlayer).identity, gameDetails, hostJoinOptions)
Public Sub netlobby_netplayHostUpdateGameListing(ByVal identityPlayer As Long, ByVal gameDetailsJson As String, ByVal botProt As Long, _
        ByVal proxyIPs As Long, ByVal hostingIPs As Long)
    Dim identity As EcKey, gd As netlobby_GameDetails, opts As netlobby_HostJoinOptions
    If netplayLobbyHostingHandler Is Nothing Then Exit Sub
    identity = netplayIdentity(identityPlayer)
    If Not netplayGameDetails(gameDetailsJson, gd) Then Exit Sub
    netplayJoinOptions botProt, proxyIPs, hostingIPs, opts
    netplayLobbyHostingHandler.updateGameListing identity, gd, opts
End Sub

' lobbyHostingHandler->respondToLobbyConnectionCheckRequest(jsonString, gameDetails.host.name, getMultiStats(realSelectedPlayer).identity)
' (True: the response; False: errorCode, a LobbyConnectionCheckRequestError)
Public Function netlobby_netplayHostRespondToLobbyConnectionCheckRequest(ByVal jsonRequest As String, ByVal playerName As String, _
        ByVal identityPlayer As Long, response As String, errorCode As Long) As Boolean
    Dim identity As EcKey, e As netlobby_LobbyConnectionCheckRequestError
    errorCode = netlobby_LobbyConnectionCheckRequestError_GameNotListed
    If netplayLobbyHostingHandler Is Nothing Then Exit Function
    identity = netplayIdentity(identityPlayer)
    netlobby_netplayHostRespondToLobbyConnectionCheckRequest = netplayLobbyHostingHandler.respondToLobbyConnectionCheckRequest(jsonRequest, playerName, identity, response, e)
    errorCode = e
End Function

' lobbyHostingHandler->checkJoinRequest(joinRequest, [i, uniqueJoinIDCopy](CheckJoinResult result) { NETtempSocketHandleLobbyJoinCheckResult(i, uniqueJoinIDCopy, result); })
Public Function netlobby_netplayHostCheckJoinRequest(ByVal joinToken As String, ByVal ip As String, ByVal playerName As String, _
        ByVal playerPublicIdentity As String, ByVal connectionMethod As Long, ByVal asSpectator As Boolean, ByVal i As Long, ByVal uniqueJoinIDCopy As String) As Boolean
    Dim joinRequest As netlobby_JoinRequest
    If netplayLobbyHostingHandler Is Nothing Then Exit Function
    joinRequest.joinToken = joinToken
    joinRequest.ip = ip
    joinRequest.playerName = playerName
    joinRequest.playerPublicIdentity = playerPublicIdentity
    joinRequest.connectionMethod = connectionMethod
    joinRequest.asSpectator = asSpectator
    netlobby_netplayHostCheckJoinRequest = netplayLobbyHostingHandler.checkJoinRequest(joinRequest, WzFunc_new(FN_netlobby_netplayCheckJoinResult, i, uniqueJoinIDCopy))
End Function

Public Sub netlobby_netplayHostRemoveGameListing(ByVal matchStarted As Boolean)
    If netplayLobbyHostingHandler Is Nothing Then Exit Sub
    netplayLobbyHostingHandler.removeGameListing matchStarted
End Sub

Public Sub netlobby_netplayHostRun()
    If netplayLobbyHostingHandler Is Nothing Then Exit Sub
    netplayLobbyHostingHandler.run
End Sub

Public Function netlobby_netplayHostIsConnectedToLobby() As Boolean
    If netplayLobbyHostingHandler Is Nothing Then Exit Function
    netlobby_netplayHostIsConnectedToLobby = netplayLobbyHostingHandler.isConnectedToLobby()
End Function

Public Function netlobby_netplayHostGetCurrentGameId(gameId As String) As Boolean
    If netplayLobbyHostingHandler Is Nothing Then Exit Function
    netlobby_netplayHostGetCurrentGameId = netplayLobbyHostingHandler.getCurrentGameId(gameId)
End Function

Public Function netlobby_netplayHostIsHostJoinOptionSupportedByLobby(ByVal optType As Long, result As Boolean) As Boolean
    If netplayLobbyHostingHandler Is Nothing Then Exit Function
    netlobby_netplayHostIsHostJoinOptionSupportedByLobby = netplayLobbyHostingHandler.isHostJoinOptionSupportedByLobby(optType, result)
End Function

' (VB) the event callbacks' and the join check's crossing into WzNet
Private Function netplayLambda(ByVal f As WzFunc, args() As Variant) As Variant
    Select Case f.id
    Case FN_netlobby_netplayGameCreated             ' (std::string gameId)
        NETlobbyGameCreatedCallback args(0)
    Case FN_netlobby_netplayGamePublished           ' (std::string gameId, std::string motd)
        NETlobbyGamePublishedCallback args(0), args(1)
    Case FN_netlobby_netplayGameGone                ' (optional<netlobby::LobbyError> error)
        NETlobbyGameGoneCallback args(0), args(1), args(2)
    Case FN_netlobby_netplayCheckJoinResult         ' [i, uniqueJoinIDCopy](CheckJoinResult result)
        NETtempSocketHandleLobbyJoinCheckResult f.cap(0), f.cap(1), args(0), args(1), args(2), args(3)
    End Select
End Function
