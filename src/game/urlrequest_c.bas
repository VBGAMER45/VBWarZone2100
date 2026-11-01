Attribute VB_Name = "urlrequest_c"
' src/urlrequest.cpp, urlrequest.h, urlrequest_private.h, whole, and src/urlrequest_curl.cpp for data requests:
' its requests, its transfer loop and its handlers, with curl's part (the easy handle and its transfer: libcurl
' with Schannel on 4.7.0's Windows build) in WzNet's urltransfer_c (NetApi's NETurlTransfer*: HTTP/1.1 over Winsock
' and Schannel; it takes createCURLHandle's options and gives CURLcodes, CURLINFO_RESPONSE_CODE, CURLINFO_PRIMARY_IP,
' CURLINFO_RETRY_AFTER, the header callback's lines and the body; its header comment says how it does curl's part).
' No thread (the project's rule): urlRequestThreadFunc's loop is urlRequestPoll, which main_c's mainLoop runs every
' frame (the delayed transfers, the new requests, the cancelled ones, performTransfers' finished transfers with
' their auto-retry strategy and the handlers' RetryAfter), and its shutdown part (the waitOnShutdown transfers
' waited for up to MAX_WAIT_ON_SHUTDOWN_SECONDS) urlRequestShutdown. The handlers run on the main thread from the
' poll (curl's run on its thread; the ported handlers still queue their main-thread work with
' wzAsyncExecOnMainThread, as 4.7.0's do). The progress callback is not called (no ported caller sets one).
' urlDownloadFile (to a file) is not ported (no caller in the port). An AsyncURLRequestHandle is a Long serial
' (0 = nullptr). The callbacks are WzFuncs: onResponse(url, HTTPResponseDetails, MemoryStruct) returns a
' URLRequestHandlingBehavior (a Double: URLRequestHandlingBehavior_Done, or the RetryAfter delay in ms);
' onFailure(url, URLRequestFailureType, HTTPResponseDetails or Nothing); progressCallback(url, dltotal, dlnow).
Option Explicit

Private Const MAX_WAIT_ON_SHUTDOWN_SECONDS As Long = 15
Private Const MAXIMUM_IN_MEMORY_DOWNLOAD_SIZE As Double = 536870912#   ' 1 << 29: 512 MB, default max download limit

Public Enum URLRequestFailureType
    URLRequestFailureType_INITIALIZE_REQUEST_ERROR
    URLRequestFailureType_OPERATION_TIMEOUT
    URLRequestFailureType_COULDNT_CONNECT
    URLRequestFailureType_TRANSFER_FAILED
    URLRequestFailureType_CANCELLED
    URLRequestFailureType_CANCELLED_BY_SHUTDOWN
End Enum

Public Enum InternetProtocol
    InternetProtocol_IP_ANY
    InternetProtocol_IPv4
    InternetProtocol_IPv6
End Enum

Public Enum URLRequestMethod
    URLRequestMethod_Get
    URLRequestMethod_Post
    URLRequestMethod_Patch
    URLRequestMethod_Delete
End Enum

Public Const HttpRequestConnectionFailure As Long = 0     ' 0 is returned as the http response code if there is an internal curl error or error connecting altogether
Public Const HttpStatusTooManyRequests As Long = 429      ' RFC 6585, 4

' struct URLRequestAutoRetryStrategy (the delays in milliseconds)
Public Type URLRequestAutoRetryStrategy
    m_minDelay As Double
    m_maxDelay As Double
    m_maxRetries As Long
    m_retryResponseCodes() As Long              ' std::unordered_set<long>: m_retryResponseCodes(0 .. m_retryResponseCodesCount - 1)
    m_retryResponseCodesCount As Long
End Type

' struct URLRequestHandlingBehavior: Done, or RetryAfter(delay) (the delay in milliseconds)
Public Const URLRequestHandlingBehavior_Done As Double = -1#

' struct URLDataRequest : public URLRequestBase
Public Type URLDataRequest
    ' URLRequestBase
    url As String
    protocol As InternetProtocol
    noProxy As Boolean                          ' Whether to disable the use of proxy.
    noFollowRedirects As Boolean                ' Whether to disable following redirects.
    connectTimeoutMs As Long                    ' uint32_t: the connection phase's maximum time (0 for default)
    userAgent As String                         ' Explicitly set the User-Agent header value
    progressCallback As WzFunc                  ' UrlProgressCallback
    onFailure As WzFunc                         ' UrlRequestFailure
    requestHeaderNames() As String              ' std::unordered_map<std::string, std::string> requestHeaders
    requestHeaderValues() As String
    requestHeadersCount As Long
    ' URLDataRequest
    maxDownloadSizeLimit As Double              ' curl_off_t
    autoRetryStrategy As URLRequestAutoRetryStrategy
    onResponse As WzFunc                        ' UrlRequestResponseHandler
    m_method As URLRequestMethod
    m_requestBody As String                     ' Data to be provided in a POST/PATCH body
End Type

' (VB) a transfer: urlrequest_curl.cpp's RunningURLDataRequest with its AsyncRequestImpl and its place in
' newUrlRequests, runningTransfers or delayedTransfers
Private Const XFER_FREE As Long = 0
Private Const XFER_NEW As Long = 1
Private Const XFER_RUNNING As Long = 2
Private Const XFER_DELAYED As Long = 3

Private Type RunningTransfer
    state As Long                               ' XFER_*
    handle As Long                              ' its AsyncURLRequestHandle
    cancelFlag As Boolean                       ' AsyncRequestImpl::cancelFlag
    request As URLDataRequest
    xfer As Long                                ' WzNet's transfer (curl's easy handle; 0: none)
    numRetries As Long
    retryAfterTime As Double                    ' (delayed) steady-clock milliseconds
    queueOrder As Double                        ' (VB) the order transfers entered their list
End Type

Private transfers() As RunningTransfer
Private transfersCount As Long
Private nextHandle As Long
Private nextQueueOrder As Double
Private urlRequestQuit As Boolean
Private polling As Boolean                      ' (VB) urlRequestPoll is running (a handler that polls is not re-entered)

' curl.h's CURLcodes the handlers look at (WzNet's urltransfer_c gives CURLcodes)
Private Const CURLE_COULDNT_RESOLVE_HOST As Long = 6
Private Const CURLE_COULDNT_CONNECT As Long = 7
Private Const CURLE_WRITE_ERROR As Long = 23
Private Const CURLE_OPERATION_TIMEDOUT As Long = 28
Private Const CURLE_SSL_CONNECT_ERROR As Long = 35

Public Const FN_urlrequest_logFailure As Long = 17200       ' [url]{ debug(...); } (and the other values the message needs)

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private qpcFreq As Currency

' (VB) std::chrono::steady_clock::now(), in milliseconds
Private Function steadyNowMs() As Double
    Dim c As Currency
    If qpcFreq = 0 Then QueryPerformanceFrequency qpcFreq
    QueryPerformanceCounter c
    steadyNowMs = CDbl(c) * 1000# / CDbl(qpcFreq)
End Function

'----------------------------------------------------------------------------------------------------------
' urlrequest.cpp

Public Function to_string_URLRequestFailureType(ByVal val As URLRequestFailureType) As String
    Select Case val
    Case URLRequestFailureType_INITIALIZE_REQUEST_ERROR: to_string_URLRequestFailureType = "INITIALIZE_REQUEST_ERROR"
    Case URLRequestFailureType_OPERATION_TIMEOUT: to_string_URLRequestFailureType = "OPERATION_TIMEOUT"
    Case URLRequestFailureType_COULDNT_CONNECT: to_string_URLRequestFailureType = "COULDNT_CONNECT"
    Case URLRequestFailureType_TRANSFER_FAILED: to_string_URLRequestFailureType = "TRANSFER_FAILED"
    Case URLRequestFailureType_CANCELLED: to_string_URLRequestFailureType = "CANCELLED"
    Case URLRequestFailureType_CANCELLED_BY_SHUTDOWN: to_string_URLRequestFailureType = "CANCELLED_BY_SHUTDOWN"
    End Select
End Function

Public Function to_string_InternetProtocol(ByVal protocol As InternetProtocol) As String
    Select Case protocol
    Case InternetProtocol_IP_ANY: to_string_InternetProtocol = "IP_ANY"
    Case InternetProtocol_IPv4: to_string_InternetProtocol = "IPv4"
    Case InternetProtocol_IPv6: to_string_InternetProtocol = "IPv6"
    Case Else: to_string_InternetProtocol = ""
    End Select
End Function

Public Function to_string_URLRequestMethod(ByVal method As URLRequestMethod) As String
    Select Case method
    Case URLRequestMethod_Get: to_string_URLRequestMethod = "GET"
    Case URLRequestMethod_Post: to_string_URLRequestMethod = "POST"
    Case URLRequestMethod_Patch: to_string_URLRequestMethod = "PATCH"
    Case URLRequestMethod_Delete: to_string_URLRequestMethod = "DELETE"
    End Select
End Function

' bool URLRequestBase::setRequestHeader(const std::string& name, const std::string& value)
Public Function URLRequestBase_setRequestHeader(request As URLDataRequest, ByVal name_ As String, ByVal value As String) As Boolean
    Dim i As Long
    If Not (Len(name_) <> 0) Then ASSERT_FAILURE "!name.empty()", "src/urlrequest.cpp:83", "setRequestHeader", "Header name must not be empty": Exit Function
    If Not (InStr(name_, ":") = 0) Then ASSERT_FAILURE "name.find_first_of("":"") == std::string::npos", "src/urlrequest.cpp:84", "setRequestHeader", "Header name must not contain ':'": Exit Function
    ' requestHeaders[name] = value;
    For i = 0 To request.requestHeadersCount - 1
        If request.requestHeaderNames(i) = name_ Then
            request.requestHeaderValues(i) = value
            URLRequestBase_setRequestHeader = True
            Exit Function
        End If
    Next
    ReDim Preserve request.requestHeaderNames(0 To request.requestHeadersCount)
    ReDim Preserve request.requestHeaderValues(0 To request.requestHeadersCount)
    request.requestHeaderNames(request.requestHeadersCount) = name_
    request.requestHeaderValues(request.requestHeadersCount) = value
    request.requestHeadersCount = request.requestHeadersCount + 1
    URLRequestBase_setRequestHeader = True
End Function

Public Sub URLDataRequest_setPost(request As URLDataRequest, ByVal jsonData As String)
    request.m_method = URLRequestMethod_Post
    request.m_requestBody = jsonData
    URLRequestBase_setRequestHeader request, "Content-Type", "application/json"
End Sub

Public Sub URLDataRequest_setPatch(request As URLDataRequest, ByVal jsonData As String)
    request.m_method = URLRequestMethod_Patch
    request.m_requestBody = jsonData
    URLRequestBase_setRequestHeader request, "Content-Type", "application/json"
End Sub

Public Sub URLDataRequest_setDelete(request As URLDataRequest)
    request.m_method = URLRequestMethod_Delete
End Sub

Public Function URLDataRequest_method(request As URLDataRequest) As URLRequestMethod
    URLDataRequest_method = request.m_method
End Function

Public Function URLDataRequest_requestBody(request As URLDataRequest) As String
    URLDataRequest_requestBody = request.m_requestBody
End Function

'----------------------------------------------------------------------------------------------------------
' urlrequest.h: URLRequestAutoRetryStrategy

Public Function URLRequestAutoRetryStrategy_None() As URLRequestAutoRetryStrategy
End Function

' RetryAfter(minDelay, maxRetries = 1, retryResponseCodes = {HttpRequestConnectionFailure, HttpStatusTooManyRequests},
' maxDelay = 10000 ms): the codes as a Variant array (Empty: the default set)
Public Function URLRequestAutoRetryStrategy_RetryAfter(ByVal minDelay As Double, Optional ByVal maxRetries As Long = 1, _
        Optional ByVal retryResponseCodes As Variant, Optional ByVal maxDelay As Double = 10000#) As URLRequestAutoRetryStrategy
    Dim result As URLRequestAutoRetryStrategy, i As Long
    result.m_minDelay = minDelay
    result.m_maxRetries = maxRetries
    result.m_maxDelay = maxDelay
    If IsMissing(retryResponseCodes) Or IsEmpty(retryResponseCodes) Then retryResponseCodes = Array(HttpRequestConnectionFailure, HttpStatusTooManyRequests)
    For i = LBound(retryResponseCodes) To UBound(retryResponseCodes)
        ReDim Preserve result.m_retryResponseCodes(0 To result.m_retryResponseCodesCount)
        result.m_retryResponseCodes(result.m_retryResponseCodesCount) = retryResponseCodes(i)
        result.m_retryResponseCodesCount = result.m_retryResponseCodesCount + 1
    Next
    URLRequestAutoRetryStrategy_RetryAfter = result
End Function

Private Function shouldRetryOnHttpResponseCode(s As URLRequestAutoRetryStrategy, ByVal httpResponseCode As Long) As Boolean
    Dim i As Long
    If s.m_maxRetries <= 0 Then Exit Function
    For i = 0 To s.m_retryResponseCodesCount - 1
        If s.m_retryResponseCodes(i) = httpResponseCode Then
            shouldRetryOnHttpResponseCode = True
            Exit Function
        End If
    Next
End Function

' URLRequestHandlingBehavior::RetryAfter(std::chrono::milliseconds delay)
Public Function URLRequestHandlingBehavior_RetryAfter(ByVal delayMs As Double) As Double
    If delayMs < 0 Then delayMs = 0
    URLRequestHandlingBehavior_RetryAfter = delayMs
End Function

'----------------------------------------------------------------------------------------------------------
' urlrequest_curl.cpp: the requests and the transfer loop

' AsyncURLRequestHandle urlRequestData(const URLDataRequest& request)
Public Function urlRequestData(request As URLDataRequest) As Long
    Dim i As Long
    If Not (Len(request.url) <> 0) Then ASSERT_FAILURE "!request.url.empty()", "src/urlrequest_curl.cpp:1419", "urlRequestData", "A valid request must specify a URL": Exit Function
    If Not (request.maxDownloadSizeLimit <= MAXIMUM_IN_MEMORY_DOWNLOAD_SIZE) Then ASSERT_FAILURE "request.maxDownloadSizeLimit <= (curl_off_t)MAXIMUM_IN_MEMORY_DOWNLOAD_SIZE", _
        "src/urlrequest_curl.cpp:1420", "urlRequestData", "Requested maxDownloadSizeLimit exceeds maximum in-memory download size limit %zu", MAXIMUM_IN_MEMORY_DOWNLOAD_SIZE: Exit Function

    ' newUrlRequests.push_back(new RunningURLDataRequest(request, requestHandle)) (started by the next poll, as
    ' the transfer thread starts it when woken)
    i = allocTransfer()
    nextHandle = nextHandle + 1
    transfers(i).handle = nextHandle
    transfers(i).cancelFlag = False
    transfers(i).request = request
    transfers(i).numRetries = 0
    setTransferState i, XFER_NEW
    urlRequestData = nextHandle
End Function

' void urlRequestSetCancelFlag(AsyncURLRequestHandle requestHandle)
Public Sub urlRequestSetCancelFlag(ByVal requestHandle As Long)
    Dim i As Long
    If requestHandle = 0 Then Exit Sub
    For i = 0 To transfersCount - 1
        If transfers(i).state <> XFER_FREE And transfers(i).handle = requestHandle Then
            transfers(i).cancelFlag = True
            Exit Sub
        End If
    Next
End Sub

Public Sub urlRequestInit()
    ' (curl_global_init, the SSL backend's selection and its locks: WzNet's transfers need none)
    ' start background thread for processing HTTP requests (the poll: urlRequestPoll)
    urlRequestQuit = False
End Sub

Public Sub urlRequestOutputDebugInfo()
    debug_ 1502, LOG_WZ, "urlRequestOutputDebugInfo", "cURL: %s", NETurlTransferVersion()
    debug_ 1503, LOG_WZ, "urlRequestOutputDebugInfo", "- SSL: %s", "Schannel"
    debug_ 1522, LOG_WZ, "urlRequestOutputDebugInfo", "- AsynchDNS: %d", 0
End Sub

Public Sub urlRequestShutdown()
    Dim i As Long, waitForTransfersStart As Double, anyRunning As Boolean
    If urlRequestQuit Then Exit Sub
    ' Signal the thread to quit
    urlRequestQuit = True

    ' while signaled to quit, there might be one or more new requests that have yet to be started,
    ' which need to be processed before fully shutting down
    processDelayedTransfers
    processNewUrlRequests True

    ' cancel any running transfers which are *not* flagged to waitOnShutdown()
    For i = 0 To transfersCount - 1
        If transfers(i).state = XFER_RUNNING Then
            If Not waitOnShutdown(i) Then
                abortTransfer i
                requestFailedToFinish i, URLRequestFailureType_CANCELLED_BY_SHUTDOWN
                freeTransfer i
            End If
        End If
    Next

    ' Wait for all remaining ("waitOnShutdown") transfers to complete
    waitForTransfersStart = steadyNowMs()
    Do
        anyRunning = False
        For i = 0 To transfersCount - 1
            If transfers(i).state = XFER_RUNNING Then anyRunning = True
        Next
        If Not anyRunning Then Exit Do
        performTransfers True, 1
        If (steadyNowMs() - waitForTransfersStart) / 1000# > MAX_WAIT_ON_SHUTDOWN_SECONDS Then
            debug_ 0, LOG_WZ, "urlRequestThreadFunc", "Failed to finish all ""waitOnShutdown"" transfers before timeout: %d", MAX_WAIT_ON_SHUTDOWN_SECONDS
            Exit Do
        End If
    Loop
    For i = 0 To transfersCount - 1
        If transfers(i).state = XFER_RUNNING Then
            abortTransfer i
            requestFailedToFinish i, URLRequestFailureType_CANCELLED_BY_SHUTDOWN
            freeTransfer i
        End If
    Next
    For i = 0 To transfersCount - 1
        If transfers(i).state = XFER_DELAYED Then
            requestFailedToFinish i, URLRequestFailureType_CANCELLED_BY_SHUTDOWN
            freeTransfer i
        End If
    Next
End Sub

' (VB) urlRequestThreadFunc's loop, one pass (main_c's mainLoop, every frame)
Public Sub urlRequestPoll()
    Dim i As Long
    If urlRequestQuit Or polling Then Exit Sub
    If transfersCount = 0 Then Exit Sub
    polling = True
    processDelayedTransfers
    ' Start handling new requests
    processNewUrlRequests False
    ' cancel any requests that have been signaled to cancel
    For i = 0 To transfersCount - 1
        If transfers(i).state = XFER_RUNNING And transfers(i).cancelFlag Then
            abortTransfer i
            requestFailedToFinish i, URLRequestFailureType_CANCELLED
            freeTransfer i
        End If
    Next
    performTransfers False, 0
    polling = False
End Sub

' processDelayedTransfers_WhileLockIsHeld: the delayed transfers whose time came go back to the new requests
Private Sub processDelayedTransfers()
    Dim i As Long, now As Double
    now = steadyNowMs()
    For i = 0 To transfersCount - 1
        If transfers(i).state = XFER_DELAYED Then
            If transfers(i).cancelFlag Then
                freeTransfer i
            ElseIf transfers(i).retryAfterTime < now Then
                setTransferState i, XFER_NEW
            End If
        End If
    Next
End Sub

' processNewUrlRequests_WhileLockIsHeld(shuttingDown), in the order the requests came
Private Sub processNewUrlRequests(ByVal shuttingDown As Boolean)
    Dim i As Long
    i = nextInState(XFER_NEW)
    Do While i >= 0
        If shuttingDown And Not waitOnShutdown(i) Then
            ' if shuttingDown, only process waitOnShutdown requests
            requestFailedToFinish i, URLRequestFailureType_CANCELLED_BY_SHUTDOWN
            freeTransfer i
        ElseIf Not createTransferHandle(i) Then
            ' something went wrong creating the handle
            requestFailedToFinish i, URLRequestFailureType_INITIALIZE_REQUEST_ERROR
            freeTransfer i
        Else
            setTransferState i, XFER_RUNNING
        End If
        i = nextInState(XFER_NEW)
    Loop
End Sub

' performTransfers: the finished transfers (WaitForResponse(waitSeconds)), with the auto-retry strategy
Private Sub performTransfers(ByVal shuttingDown As Boolean, ByVal waitSeconds As Long)
    Dim i As Long, done As Boolean, result As Long, resultText As String, httpResponseCode As Long
    Dim retryAfterSeconds As Double, minRetryDelay As Double, delayMS As Double, responseBehavior As Double
    For i = 0 To transfersCount - 1
        If transfers(i).state = XFER_RUNNING Then
            done = pollTransfer(i, waitSeconds, result, resultText)
            If done Then
                If Not shuttingDown And transfers(i).request.autoRetryStrategy.m_maxRetries > 0 Then
                    If result = 0 Then
                        httpResponseCode = transferStatus(i)
                        If shouldRetryOnHttpResponseCode(transfers(i).request.autoRetryStrategy, httpResponseCode) And _
                           transfers(i).numRetries < transfers(i).request.autoRetryStrategy.m_maxRetries Then
                            transfers(i).numRetries = transfers(i).numRetries + 1

                            retryAfterSeconds = transferRetryAfter(i)
                            If retryAfterSeconds < 0 Then retryAfterSeconds = 0
                            If retryAfterSeconds > 3600 Then retryAfterSeconds = 3600

                            minRetryDelay = transfers(i).request.autoRetryStrategy.m_minDelay * transfers(i).numRetries
                            If minRetryDelay > transfers(i).request.autoRetryStrategy.m_maxDelay Then minRetryDelay = transfers(i).request.autoRetryStrategy.m_maxDelay

                            ' std::clamp(retryAfterSeconds * 1000, minRetryDelay, maxDelay)
                            delayMS = retryAfterSeconds * 1000#
                            If delayMS < minRetryDelay Then delayMS = minRetryDelay
                            If delayMS > transfers(i).request.autoRetryStrategy.m_maxDelay Then delayMS = transfers(i).request.autoRetryStrategy.m_maxDelay
                            abortTransfer i
                            transfers(i).retryAfterTime = steadyNowMs() + delayMS
                            setTransferState i, XFER_DELAYED
                            GoTo continue_
                        End If
                    End If
                End If

                responseBehavior = handleRequestDone(i, result, resultText)
                If responseBehavior <> URLRequestHandlingBehavior_Done And Not shuttingDown And transfers(i).state = XFER_RUNNING Then
                    transfers(i).numRetries = transfers(i).numRetries + 1
                    abortTransfer i
                    transfers(i).retryAfterTime = steadyNowMs() + responseBehavior
                    setTransferState i, XFER_DELAYED
                ElseIf transfers(i).state = XFER_RUNNING Then
                    ' If shutting down, don't retry!
                    freeTransfer i
                End If
            End If
        End If
continue_:
    Next
End Sub

' (VB) curl_multi_perform and curl_multi_info_read for the transfer: True when it finished, with result its
' CURLcode and curl_easy_strerror's text (waitSeconds: curl_multi_poll's wait)
Private Function pollTransfer(ByVal i As Long, ByVal waitSeconds As Long, result As Long, resultText As String) As Boolean
    Dim code As Long, ip As String, retryAfter As Double
    If Not NETurlTransferPoll(transfers(i).xfer, waitSeconds * 1000) Then Exit Function
    result = NETurlTransferResult(transfers(i).xfer, resultText, code, ip, retryAfter)
    pollTransfer = True
End Function

' (VB) CURLINFO_RESPONSE_CODE (0 when there is none)
Private Function transferStatus(ByVal i As Long) As Long
    Dim t As String, code As Long, ip As String, retryAfter As Double
    NETurlTransferResult transfers(i).xfer, t, code, ip, retryAfter
    transferStatus = code
End Function

' (VB) CURLINFO_RETRY_AFTER: the Retry-After header's seconds (0 when absent)
Private Function transferRetryAfter(ByVal i As Long) As Double
    Dim t As String, code As Long, ip As String, retryAfter As Double
    NETurlTransferResult transfers(i).xfer, t, code, ip, retryAfter
    transferRetryAfter = retryAfter
End Function

' (VB) CURLINFO_PRIMARY_IP ("" before a connection)
Private Function transferPrimaryIp(ByVal i As Long) As String
    Dim t As String, code As Long, ip As String, retryAfter As Double
    NETurlTransferResult transfers(i).xfer, t, code, ip, retryAfter
    transferPrimaryIp = ip
End Function

' RunningURLTransferRequestBase::handleRequestDone(CURLcode result): the URLRequestHandlingBehavior
Private Function handleRequestDone(ByVal i As Long, ByVal result As Long, ByVal resultText As String) As Double
    Dim code As Long, details As HTTPResponseDetails, chunk As MemoryStruct, failureType As URLRequestFailureType
    Dim body() As Byte, n As Long, maxSize As Double

    code = transferStatus(i)

    Set details = New HTTPResponseDetails
    Set chunk = New MemoryStruct
    If result = 0 Then
        ' the body (WriteMemoryCallback_URLTransferRequest), up to the transfer's maxDownloadSize
        maxSize = MAXIMUM_IN_MEMORY_DOWNLOAD_SIZE
        If transfers(i).request.maxDownloadSizeLimit > 0 And transfers(i).request.maxDownloadSizeLimit < maxSize Then maxSize = transfers(i).request.maxDownloadSizeLimit
        n = transferBody(i, body)
        If n + 1 > maxSize Then
            ' (writeMemoryCallback refuses what would reach maxDownloadSize: CURLE_WRITE_ERROR)
            ASSERT_FAILURE "bytesToWrite == realsize", "src/urlrequest_curl.cpp:991", "writeMemoryCallback", "Writing partial data - reached max size"
            result = CURLE_WRITE_ERROR
            resultText = "Failed writing received data to disk/application"
            n = 0
        Else
            chunk.setMemory body, n
        End If
    End If
    details.init result, resultText, code, True, transferPrimaryIp(i)
    readResponseHeaders i, details

    If result = 0 Then
        handleRequestDone = onResponse(i, details, chunk)
    Else
        failureType = URLRequestFailureType_TRANSFER_FAILED
        Select Case result
        Case CURLE_OPERATION_TIMEDOUT
            failureType = URLRequestFailureType_OPERATION_TIMEOUT
        Case CURLE_COULDNT_RESOLVE_HOST, CURLE_COULDNT_CONNECT, CURLE_SSL_CONNECT_ERROR
            failureType = URLRequestFailureType_COULDNT_CONNECT
        End Select
        onFailure i, failureType, details
        handleRequestDone = URLRequestHandlingBehavior_Done
    End If
End Function

' (VB) the response body's bytes (writeMemoryCallback's chunk): the count
Private Function transferBody(ByVal i As Long, body() As Byte) As Long
    transferBody = NETurlTransferBody(transfers(i).xfer, body)
End Function

' header_callback, for each header line of the responses
Private Sub readResponseHeaders(ByVal i As Long, details As HTTPResponseDetails)
    Dim lines() As String, k As Long, header_line As String, header_separator As Long
    lines = Split(NETurlTransferHeaders(transfers(i).xfer), vbLf)
    For k = 0 To UBound(lines)
        header_line = lines(k)
        header_separator = InStr(header_line, ":")
        If header_separator <> 0 Then
            details.setHeader trim_str(Left$(header_line, header_separator - 1)), trim_str(Mid$(header_line, header_separator + 1))
        End If
    Next
End Sub

' urlrequest_private.h: trim_str (the leading and trailing white space)
Private Function trim_str(ByVal s As String) As String
    Dim a As Long, b As Long
    a = 1
    b = Len(s)
    Do While a <= b
        If InStr(" " & vbTab & vbCr & vbLf & vbVerticalTab & vbFormFeed, Mid$(s, a, 1)) = 0 Then Exit Do
        a = a + 1
    Loop
    Do While b >= a
        If InStr(" " & vbTab & vbCr & vbLf & vbVerticalTab & vbFormFeed, Mid$(s, b, 1)) = 0 Then Exit Do
        b = b - 1
    Loop
    trim_str = Mid$(s, a, b - a + 1)
End Function

' RunningURLDataRequest::onResponse
Private Function onResponse(ByVal i As Long, details As HTTPResponseDetails, chunk As MemoryStruct) As Double
    Dim r As Variant, f As WzFunc, url As String
    ' (copies: a ParamArray argument is passed by reference, which would lock transfers() against a handler's
    ' new request)
    Set f = transfers(i).request.onResponse
    url = transfers(i).request.url
    If Not f Is Nothing Then
        r = WzFunc_call(f, url, details, chunk)
        If IsEmpty(r) Then onResponse = URLRequestHandlingBehavior_Done Else onResponse = CDbl(r)
        Exit Function
    End If
    onResponse = URLRequestHandlingBehavior_Done
End Function

' RunningURLTransferRequestBase::requestFailedToFinish(URLRequestFailureType type)
Private Sub requestFailedToFinish(ByVal i As Long, ByVal typ As URLRequestFailureType)
    If Not (typ <> URLRequestFailureType_TRANSFER_FAILED) Then ASSERT_FAILURE "type != URLRequestFailureType::TRANSFER_FAILED", "src/urlrequest_curl.cpp:851", "requestFailedToFinish", "TRANSFER_FAILED should be handled by handleRequestDone": Exit Sub
    onFailure i, typ, Nothing
End Sub

' RunningURLTransferRequestBase::onFailure(type, transferDetails)
Private Sub onFailure(ByVal i As Long, ByVal typ As URLRequestFailureType, transferDetails As HTTPResponseDetails)
    Dim url As String, f As WzFunc
    url = transfers(i).request.url
    Select Case typ
    Case URLRequestFailureType_INITIALIZE_REQUEST_ERROR
        wzAsyncExecOnMainThread WzFunc_new(FN_urlrequest_logFailure, 866, LOG_NET, "cURL: Failed to initialize request for (%s)", url)
    Case URLRequestFailureType_OPERATION_TIMEOUT
        wzAsyncExecOnMainThread WzFunc_new(FN_urlrequest_logFailure, 871, LOG_NET, "cURL: Request timeout for (%s)", url)
    Case URLRequestFailureType_COULDNT_CONNECT
        wzAsyncExecOnMainThread WzFunc_new(FN_urlrequest_logFailure, 876, LOG_NET, "cURL: Couldn't connect for (%s)", url)
    Case URLRequestFailureType_TRANSFER_FAILED
        If transferDetails Is Nothing Then
            wzAsyncExecOnMainThread WzFunc_new(FN_urlrequest_logFailure, 884, LOG_ERROR, "cURL: Request for (%s) failed - but no transfer failure details provided!", url)
        Else
            wzAsyncExecOnMainThread WzFunc_new(FN_urlrequest_logFailure, 891, LOG_NET, "cURL: Request for (%s) failed with error %d, and HTTP response code: %ld", url, _
                transferDetails.curlResult(), transferDetails.httpStatusCode())
        End If
    Case URLRequestFailureType_CANCELLED
        wzAsyncExecOnMainThread WzFunc_new(FN_urlrequest_logFailure, 898, LOG_NET, "cURL: Request for (%s) was cancelled", url)
    Case URLRequestFailureType_CANCELLED_BY_SHUTDOWN
        wzAsyncExecOnMainThread WzFunc_new(FN_urlrequest_logFailure, 903, LOG_NET, "cURL: Request for (%s) was cancelled by application shutdown", url)
    End Select

    Set f = transfers(i).request.onFailure
    If Not f Is Nothing Then
        WzFunc_call f, url, typ, transferDetails
    End If
End Sub

Public Function urlrequest_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Select Case f.id
    Case FN_urlrequest_logFailure                           ' [line, part, format, url, a, b]
        If IsEmpty(f.cap(4)) Then
            debug_ f.cap(0), f.cap(1), "onFailure", f.cap(2), f.cap(3)
        Else
            debug_ f.cap(0), f.cap(1), "onFailure", f.cap(2), f.cap(3), f.cap(4), f.cap(5)
        End If
    End Select
End Function

' virtual bool waitOnShutdown() const { return method() != URLRequestMethod::Get; }
Private Function waitOnShutdown(ByVal i As Long) As Boolean
    waitOnShutdown = transfers(i).request.m_method <> URLRequestMethod_Get
End Function

' URLTransferRequest::createCURLHandle: the transfer with its options (CURLOPT_CUSTOMREQUEST, CURLOPT_URL,
' CURLOPT_IPRESOLVE, CURLOPT_CONNECTTIMEOUT_MS, CURLOPT_NOPROXY (WzNet's transfers are always direct),
' CURLOPT_HTTPHEADER, CURLOPT_USERAGENT, CURLOPT_POSTFIELDS, CURLOPT_FOLLOWLOCATION, CURLOPT_MAXFILESIZE_LARGE; the
' rest is fixed in urltransfer_c), added to the transfers (curl_multi_add_handle)
Private Function createTransferHandle(ByVal i As Long) As Boolean
    Dim k As Long, headers As String, body() As Byte, n As Long, maxSize As Double
    If transfers(i).xfer <> 0 Then
        ' reuse already-created handle
        createTransferHandle = True
        Exit Function
    End If
    For k = 0 To transfers(i).request.requestHeadersCount - 1
        headers = headers & transfers(i).request.requestHeaderNames(k) & ": " & transfers(i).request.requestHeaderValues(k) & vbCrLf
    Next
    ' CURLOPT_POSTFIELDS for a POST or PATCH whose requestBody() is not empty
    n = -1
    If transfers(i).request.m_method = URLRequestMethod_Post Or transfers(i).request.m_method = URLRequestMethod_Patch Then
        If Len(transfers(i).request.m_requestBody) <> 0 Then n = utf8_encode(transfers(i).request.m_requestBody, body)
    End If
    ' maxDownloadSize(): MAXIMUM_IN_MEMORY_DOWNLOAD_SIZE, or the request's lower limit
    maxSize = MAXIMUM_IN_MEMORY_DOWNLOAD_SIZE
    If transfers(i).request.maxDownloadSizeLimit > 0 And transfers(i).request.maxDownloadSizeLimit < maxSize Then maxSize = transfers(i).request.maxDownloadSizeLimit
    transfers(i).xfer = NETurlTransferStart(to_string_URLRequestMethod(transfers(i).request.m_method), transfers(i).request.url, _
        headers, body, n, transfers(i).request.protocol, transfers(i).request.connectTimeoutMs, Not transfers(i).request.noFollowRedirects, _
        transfers(i).request.userAgent, maxSize)
    createTransferHandle = (transfers(i).xfer <> 0)
End Function

' curl_multi_remove_handle + curl_easy_cleanup (a running transfer is aborted)
Private Sub abortTransfer(ByVal i As Long)
    If transfers(i).xfer <> 0 Then NETurlTransferClose transfers(i).xfer
    transfers(i).xfer = 0
End Sub

'----------------------------------------------------------------------------------------------------------
' (VB) the transfer slots

Private Function allocTransfer() As Long
    Dim i As Long
    For i = 0 To transfersCount - 1
        If transfers(i).state = XFER_FREE Then
            allocTransfer = i
            Exit Function
        End If
    Next
    ReDim Preserve transfers(0 To transfersCount)
    allocTransfer = transfersCount
    transfersCount = transfersCount + 1
End Function

Private Sub setTransferState(ByVal i As Long, ByVal state As Long)
    transfers(i).state = state
    nextQueueOrder = nextQueueOrder + 1
    transfers(i).queueOrder = nextQueueOrder
End Sub

' the first transfer (in queue order) in a state, or -1
Private Function nextInState(ByVal state As Long) As Long
    Dim i As Long, best As Long
    best = -1
    For i = 0 To transfersCount - 1
        If transfers(i).state = state Then
            If best < 0 Then
                best = i
            ElseIf transfers(i).queueOrder < transfers(best).queueOrder Then
                best = i
            End If
        End If
    Next
    nextInState = best
End Function

' delete urlTransfer
Private Sub freeTransfer(ByVal i As Long)
    Dim empty_ As RunningTransfer
    abortTransfer i
    transfers(i) = empty_
End Sub
