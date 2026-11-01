Attribute VB_Name = "urltransfer_c"
' (cluster J part 2) curl's part of src/urlrequest_curl.cpp on Windows, for the exe's urlrequest_c: 4.7.0's
' Windows build links libcurl with its Schannel backend and zlib (vcpkg curl, "ssl"); here an HTTP/1.1 client over
' Winsock 2 and SSPI's Schannel, without threads: urlrequest_c's poll (every frame) runs urlTransferPoll, which goes
' on until a socket would block. A transfer is a slot of this module's table and its handle is index + 1 (0: none).
' What it does with the options urlrequest_curl.cpp's createCURLHandle sets, as libcurl does:
' * CURLOPT_IPRESOLVE: the resolve's family (IPv4: AF_INET only; IPv6: AF_INET6 only; IP_ANY: both, in the
'   system's order) and, with both, curl's happy eyeballs: the first address's family first, the other family's
'   first attempt 200 ms later (CURLOPT_HAPPY_EYEBALLS_TIMEOUT_MS's default), the first to connect wins; each
'   family's addresses are tried in turn, an address with more after it getting half the time left;
' * CURLOPT_CONNECTTIMEOUT_MS (0: curl's default, 300 s) covers the resolve, the connect and the TLS handshake;
'   CURLOPT_LOW_SPEED_TIME 60 / CURLOPT_LOW_SPEED_LIMIT 30: once the request is being sent, under 30 bytes/s (over
'   the last 5 seconds) for 60 seconds; both end in CURLE_OPERATION_TIMEDOUT;
' * CURLOPT_FOLLOWLOCATION with CURLOPT_MAXREDIRS 10 (CURLE_TOO_MANY_REDIRECTS past it), to http and https only
'   (CURLOPT_REDIR_PROTOCOLS_STR); as curl 8 does with a CURLOPT_CUSTOMREQUEST, the method stays, a 301 or 302
'   after a request with a body and a 303 drop the body, 307 and 308 keep it; the header lines of every response
'   (1xx, redirects, chunked trailers) go to the header list, as curl calls the header callback for each;
' * CURLOPT_ACCEPT_ENCODING "" (curl with zlib: "deflate, gzip") and CURLOPT_HTTP_CONTENT_DECODING: the body is
'   decoded with miniz (src\shared\miniz) when the transfer ends (curl decodes as it goes);
' * CURLOPT_MAXFILESIZE_LARGE: a Content-Length past it, or a body that grows past it: CURLE_FILESIZE_EXCEEDED;
' * CURLOPT_NOPROXY "*" or not: the transfer is always direct (curl without noProxy would honour the http_proxy,
'   https_proxy, all_proxy and no_proxy environment variables; this transport has no proxy support);
' * TLS as curl's schannel.c sets it up by default: the certificate checked by Schannel against the system's roots
'   with the host name (SNI and the name check) and revocation checks over the chain (curl's default without
'   CURLSSLOPT_NO_REVOKE), TLS 1.2 at least (curl 8's default), TLS 1.3 offered where Schannel has it (Windows
'   build 20348 and later), SCH_CREDENTIALS from Windows 10 1809 (build 17763) and SCHANNEL_CRED before it; no
'   client certificate is sent; close_notify when the connection closes after a complete response.
' The request as curl writes it: the request line, Host, User-Agent, Accept: */*, Accept-Encoding, the request's
' headers (a header of the same name replaces curl's), Content-Length and curl's default Content-Type for a body;
' no Connection header: the connection is closed after the response (no reuse).
' The resolve (getaddrinfo) blocks, as WzNet's other resolves do (curl's threaded resolver does not).
' The results are CURLcodes, with curl_easy_strerror's texts; each failure is logged (LOG_NET) with its detail,
' and each connection with its address and family.
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)
Private Declare Function QueryPerformanceCounter Lib "kernel32" (c As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (f As Currency) As Long
Private Declare Function SetHandleInformation Lib "kernel32" (ByVal hObject As Long, ByVal dwMask As Long, ByVal dwFlags As Long) As Long
Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Private Declare Function RtlGetVersion Lib "ntdll" (lpVersionInformation As Any) As Long
Private Declare Function ws_WSAStartup Lib "ws2_32" Alias "WSAStartup" (ByVal wVersionRequested As Integer, lpWSAData As Any) As Long
Private Declare Function ws_socket Lib "ws2_32" Alias "socket" (ByVal af As Long, ByVal type_ As Long, ByVal protocol As Long) As Long
Private Declare Function ws_closesocket Lib "ws2_32" Alias "closesocket" (ByVal s As Long) As Long
Private Declare Function ws_connect Lib "ws2_32" Alias "connect" (ByVal s As Long, name_ As Any, ByVal namelen As Long) As Long
Private Declare Function ws_ioctlsocket Lib "ws2_32" Alias "ioctlsocket" (ByVal s As Long, ByVal cmd As Long, argp As Long) As Long
Private Declare Function ws_setsockopt Lib "ws2_32" Alias "setsockopt" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, optval As Long, ByVal optlen As Long) As Long
Private Declare Function ws_getsockopt Lib "ws2_32" Alias "getsockopt" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, optval As Long, optlen As Long) As Long
Private Declare Function ws_select Lib "ws2_32" Alias "select" (ByVal nfds As Long, readfds As Any, writefds As Any, exceptfds As Any, timeout As timeval) As Long
Private Declare Function ws_send Lib "ws2_32" Alias "send" (ByVal s As Long, buf As Any, ByVal len_ As Long, ByVal flags As Long) As Long
Private Declare Function ws_recv Lib "ws2_32" Alias "recv" (ByVal s As Long, buf As Any, ByVal len_ As Long, ByVal flags As Long) As Long
Private Declare Function ws_getaddrinfo Lib "ws2_32" Alias "getaddrinfo" (ByVal pNodeName As String, ByVal pServiceName As String, pHints As ADDRINFOA, ppResult As Long) As Long
Private Declare Sub ws_freeaddrinfo Lib "ws2_32" Alias "freeaddrinfo" (ByVal pAddrInfo As Long)
Private Declare Function ws_inet_ntop Lib "ws2_32" Alias "inet_ntop" (ByVal Family As Long, pAddr As Any, ByVal pStringBuf As String, ByVal StringBufSize As Long) As Long
Private Declare Function AcquireCredentialsHandleW Lib "secur32" (ByVal pszPrincipal As Long, ByVal pszPackage As Long, ByVal fCredentialUse As Long, _
    ByVal pvLogonId As Long, pAuthData As Any, ByVal pGetKeyFn As Long, ByVal pvGetKeyArgument As Long, phCredential As SecHandle, ptsExpiry As Currency) As Long
Private Declare Function InitializeSecurityContextW Lib "secur32" (phCredential As SecHandle, ByVal phContext As Long, ByVal pszTargetName As Long, _
    ByVal fContextReq As Long, ByVal Reserved1 As Long, ByVal TargetDataRep As Long, ByVal pInput As Long, ByVal Reserved2 As Long, _
    phNewContext As SecHandle, pOutput As SecBufferDesc, pfContextAttr As Long, ptsExpiry As Currency) As Long
Private Declare Function DeleteSecurityContext Lib "secur32" (phContext As SecHandle) As Long
Private Declare Function FreeContextBuffer Lib "secur32" (ByVal pvContextBuffer As Long) As Long
Private Declare Function QueryContextAttributesW Lib "secur32" (phContext As SecHandle, ByVal ulAttribute As Long, pBuffer As Any) As Long
Private Declare Function EncryptMessage Lib "secur32" (phContext As SecHandle, ByVal fQOP As Long, pMessage As SecBufferDesc, ByVal MessageSeqNo As Long) As Long
Private Declare Function DecryptMessage Lib "secur32" (phContext As SecHandle, pMessage As SecBufferDesc, ByVal MessageSeqNo As Long, ByVal pfQOP As Long) As Long
Private Declare Function ApplyControlToken Lib "secur32" (phContext As SecHandle, pInput As SecBufferDesc) As Long

' CURLcode (curl.h) and curl's defaults
Private Const CURLE_OK As Long = 0
Private Const CURLE_UNSUPPORTED_PROTOCOL As Long = 1
Private Const CURLE_URL_MALFORMAT As Long = 3
Private Const CURLE_COULDNT_RESOLVE_HOST As Long = 6
Private Const CURLE_COULDNT_CONNECT As Long = 7
Private Const CURLE_WEIRD_SERVER_REPLY As Long = 8
Private Const CURLE_PARTIAL_FILE As Long = 18
Private Const CURLE_WRITE_ERROR As Long = 23
Private Const CURLE_OPERATION_TIMEDOUT As Long = 28
Private Const CURLE_SSL_CONNECT_ERROR As Long = 35
Private Const CURLE_TOO_MANY_REDIRECTS As Long = 47
Private Const CURLE_GOT_NOTHING As Long = 52
Private Const CURLE_SEND_ERROR As Long = 55
Private Const CURLE_RECV_ERROR As Long = 56
Private Const CURLE_PEER_FAILED_VERIFICATION As Long = 60
Private Const CURLE_BAD_CONTENT_ENCODING As Long = 61
Private Const CURLE_FILESIZE_EXCEEDED As Long = 63
Private Const CURLE_TOO_LARGE As Long = 100
Private Const DEFAULT_CONNECT_TIMEOUT As Long = 300000          ' curl's connect.h, ms
Private Const HAPPY_EYEBALLS_TIMEOUT As Long = 200              ' CURLOPT_HAPPY_EYEBALLS_TIMEOUT_MS's default
Private Const MAXREDIRS As Long = 10                            ' urlrequest_curl.cpp's CURLOPT_MAXREDIRS
Private Const LOW_SPEED_LIMIT As Long = 30                      ' urlrequest_curl.cpp's, bytes/s
Private Const LOW_SPEED_TIME As Long = 60000                    ' urlrequest_curl.cpp's (60 s)
Private Const MAX_HTTP_RESP_HEADER_SIZE As Long = 307200        ' curl's http.h (300 KB)
Private Const ACCEPT_ENCODINGS As String = "deflate, gzip"      ' Curl_get_content_encodings() with zlib only
Private Const MAX_STEPS_PER_POLL As Long = 64                   ' (VB) a poll's work, so a large body does not hold a frame

' Winsock
Private Const INVALID_SOCKET As Long = -1
Private Const SOCKET_ERROR As Long = -1
Private Const AF_UNSPEC As Long = 0
Private Const AF_INET As Long = 2
Private Const AF_INET6 As Long = 23
Private Const SOCK_STREAM As Long = 1
Private Const IPPROTO_TCP As Long = 6
Private Const TCP_NODELAY As Long = 1
Private Const SOL_SOCKET As Long = &HFFFF&
Private Const SO_ERROR As Long = &H1007&
Private Const FIONBIO As Long = &H8004667E
Private Const HANDLE_FLAG_INHERIT As Long = 1
Private Const WSAEWOULDBLOCK As Long = 10035
Private Const SOCKADDR_SIZE As Long = 28                        ' sizeof(SOCKADDR_IN6)
Private Const RECV_SIZE As Long = 16384

' SSPI and Schannel (sspi.h, schannel.h)
Private Const SECBUFFER_VERSION As Long = 0
Private Const SECBUFFER_EMPTY As Long = 0
Private Const SECBUFFER_DATA As Long = 1
Private Const SECBUFFER_TOKEN As Long = 2
Private Const SECBUFFER_EXTRA As Long = 5
Private Const SECBUFFER_STREAM_TRAILER As Long = 6
Private Const SECBUFFER_STREAM_HEADER As Long = 7
Private Const SECBUFFER_ALERT As Long = 17
Private Const SECBUFFER_APPLICATION_PROTOCOLS As Long = 18
Private Const SecApplicationProtocolNegotiationExt_ALPN As Long = 2
Private Const SECPKG_CRED_OUTBOUND As Long = 2
Private Const SECPKG_ATTR_STREAM_SIZES As Long = 4
Private Const SECPKG_ATTR_CONNECTION_INFO As Long = &H5A&
Private Const UNISP_NAME As String = "Microsoft Unified Security Protocol Provider"
Private Const SCHANNEL_CRED_VERSION As Long = 4
Private Const SCH_CREDENTIALS_VERSION As Long = 5
Private Const SCHANNEL_SHUTDOWN As Long = 1
Private Const SP_PROT_TLS1_2_CLIENT As Long = &H800&
Private Const SP_PROT_TLS1_3_CLIENT As Long = &H2000&
Private Const SCH_CRED_NO_DEFAULT_CREDS As Long = &H10&
Private Const SCH_CRED_AUTO_CRED_VALIDATION As Long = &H20&
Private Const SCH_CRED_REVOCATION_CHECK_CHAIN As Long = &H200&
Private Const SCH_USE_STRONG_CRYPTO As Long = &H400000
Private Const ISC_REQ_REPLAY_DETECT As Long = &H4&
Private Const ISC_REQ_SEQUENCE_DETECT As Long = &H8&
Private Const ISC_REQ_CONFIDENTIALITY As Long = &H10&
Private Const ISC_REQ_USE_SUPPLIED_CREDS As Long = &H80&
Private Const ISC_REQ_ALLOCATE_MEMORY As Long = &H100&
Private Const ISC_REQ_STREAM As Long = &H8000&
Private Const SEC_E_OK As Long = 0
Private Const SEC_I_CONTINUE_NEEDED As Long = &H90312
Private Const SEC_I_CONTEXT_EXPIRED As Long = &H90317
Private Const SEC_I_INCOMPLETE_CREDENTIALS As Long = &H90320
Private Const SEC_I_RENEGOTIATE As Long = &H90321
Private Const SEC_E_INCOMPLETE_MESSAGE As Long = &H80090318
Private Const SEC_E_WRONG_PRINCIPAL As Long = &H80090322

Private Type SecHandle
    dwLower As Long
    dwUpper As Long
End Type

Private Type SecBuffer
    cbBuffer As Long
    BufferType As Long
    pvBuffer As Long
End Type

Private Type SecBufferDesc
    ulVersion As Long
    cBuffers As Long
    pBuffers As Long
End Type

Private Type SCHANNEL_CRED
    dwVersion As Long
    cCreds As Long
    paCred As Long
    hRootStore As Long
    cMappers As Long
    aphMappers As Long
    cSupportedAlgs As Long
    palgSupportedAlgs As Long
    grbitEnabledProtocols As Long
    dwMinimumCipherStrength As Long
    dwMaximumCipherStrength As Long
    dwSessionLifespan As Long
    dwFlags As Long
    dwCredFormat As Long
End Type

Private Type SCH_CREDENTIALS
    dwVersion As Long
    dwCredFormat As Long
    cCreds As Long
    paCred As Long
    hRootStore As Long
    cMappers As Long
    aphMappers As Long
    dwSessionLifespan As Long
    dwFlags As Long
    cTlsParameters As Long
    pTlsParameters As Long
End Type

Private Type TLS_PARAMETERS
    cAlpnIds As Long
    rgstrAlpnIds As Long
    grbitDisabledProtocols As Long
    cDisabledCrypto As Long
    pDisabledCrypto As Long
    dwFlags As Long
End Type

Private Type SecPkgContext_StreamSizes
    cbHeader As Long
    cbTrailer As Long
    cbMaximumMessage As Long
    cBuffers As Long
    cbBlockSize As Long
End Type

Private Type SecPkgContext_ConnectionInfo
    dwProtocol As Long
    aiCipher As Long
    dwCipherStrength As Long
    aiHash As Long
    dwHashStrength As Long
    aiExch As Long
    dwExchStrength As Long
End Type

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

' the transfer's phases
Private Const XP_FREE As Long = 0
Private Const XP_RESOLVE As Long = 1
Private Const XP_CONNECT As Long = 2
Private Const XP_TLS As Long = 3
Private Const XP_SEND As Long = 4
Private Const XP_HEADERS As Long = 5
Private Const XP_BODY As Long = 6
Private Const XP_DONE As Long = 7

' the body's framing
Private Const BM_NONE As Long = 0
Private Const BM_LENGTH As Long = 1
Private Const BM_CHUNKED As Long = 2
Private Const BM_CLOSE As Long = 3
' chunked's states
Private Const CH_SIZE As Long = 0
Private Const CH_DATA As Long = 1
Private Const CH_DATA_END As Long = 2
Private Const CH_TRAILER As Long = 3

' a transfer: curl's easy handle with its connection
Private Type UrlTransfer
    phase As Long                       ' XP_*
    result As Long                      ' (XP_DONE) the CURLcode
    ' the request (createCURLHandle's options)
    method As String                    ' CURLOPT_CUSTOMREQUEST
    url As String                       ' CURLOPT_URL, then each redirect's
    headers As String                   ' CURLOPT_HTTPHEADER: "Name: value" lines, each ending in vbCrLf
    body() As Byte                      ' CURLOPT_POSTFIELDS
    bodyLen As Long                     ' -1: none (a 301, 302 or 303 may drop it)
    family As Long                      ' CURLOPT_IPRESOLVE: AF_INET, AF_INET6 or AF_UNSPEC
    connectTimeoutMs As Long            ' CURLOPT_CONNECTTIMEOUT_MS
    followRedirects As Boolean          ' CURLOPT_FOLLOWLOCATION
    redirects As Long
    userAgent As String                 ' CURLOPT_USERAGENT
    maxFileSize As Double               ' CURLOPT_MAXFILESIZE_LARGE (0: none)
    ' the URL's parts
    https As Boolean
    host As String                      ' (an IPv6 address without its brackets)
    hostHeader As String                ' the Host header's value
    port As Long
    target As String                    ' the path and the query
    ' the connect (resolve, connect and TLS handshake)
    connectStart As Double
    connectDeadline As Double
    addrs() As Byte                     ' the resolved addresses, SOCKADDR_SIZE bytes each
    addrFam() As Long
    addrLen() As Long
    addrCount As Long
    balFam(0 To 1) As Long              ' happy eyeballs: each family's attempts (0: no address of a second family)
    balNext(0 To 1) As Long             ' the next address to look at
    balSock(0 To 1) As Long             ' the attempt's socket
    balAddr(0 To 1) As Long             ' the attempt's address
    balDeadline(0 To 1) As Double
    balState(0 To 1) As Long            ' 0 not started, 1 trying, 2 no address left
    sock As Long                        ' the connection
    primaryIp As String                 ' CURLINFO_PRIMARY_IP
    ' Schannel
    hasCtx As Boolean
    ctx As SecHandle
    handshakeDone As Boolean
    renegotiating As Boolean
    savedPhase As Long                  ' (renegotiating) the phase to go back to
    reqFlags As Long
    sizes As SecPkgContext_StreamSizes
    tlsIn() As Byte                     ' received, not yet decrypted
    tlsInLen As Long
    tlsNeedMore As Boolean              ' the last decryption wanted more bytes
    ' sending
    sendBuf() As Byte
    sendLen As Long
    sendPos As Long
    ' receiving
    plain() As Byte                     ' received (and decrypted), not yet parsed
    plainLen As Long
    eof As Boolean
    received As Double                  ' bytes received on this connection
    ' the response
    status As Long                      ' CURLINFO_RESPONSE_CODE
    headerText As String                ' every line the header callback would get, each ending in vbLf
    contentLength As Double             ' -1: none
    chunked As Boolean
    contentEncoding As String
    location As String
    retryAfter As Double                ' CURLINFO_RETRY_AFTER
    bodyMode As Long                    ' BM_*
    left_ As Double                     ' BM_LENGTH: the body's bytes left; chunked: the chunk's
    chunkState As Long                  ' CH_*
    resp() As Byte                      ' the body
    respLen As Long
    ' the low-speed check
    total As Double                     ' bytes sent and received since the request started
    sampleTime(0 To 5) As Double        ' one sample a second, the newest last
    sampleBytes(0 To 5) As Double
    sampleCount As Long
    lowSpeedSince As Double             ' 0: not slow
End Type

Private xfers() As UrlTransfer
Private xferCount As Long
Private wsaStarted As Boolean
Private credAcquired As Boolean
Private cred As SecHandle               ' the Schannel credentials (curl's are per connection; one serves all here)
Private credTls13 As Boolean
Private qpcFreq As Currency
Private rbuf(0 To RECV_SIZE - 1) As Byte

'----------------------------------------------------------------------------------------------------------
' the exe's calls (NetApi)

' a transfer started (curl_multi_add_handle of createCURLHandle's handle): method (CURLOPT_CUSTOMREQUEST), url,
' headers ("Name: value" lines ending in vbCrLf), body(0 .. bodyLen - 1) (bodyLen -1: no CURLOPT_POSTFIELDS),
' protocol (urlrequest.h's InternetProtocol: 0 IP_ANY, 1 IPv4, 2 IPv6), connectTimeoutMs (0: curl's default),
' followRedirects, userAgent ("": none), maxFileSize (0: none). The handle (index + 1).
Public Function urlTransferStart(ByVal method As String, ByVal url As String, ByVal headers As String, body() As Byte, _
        ByVal bodyLen As Long, ByVal protocol As Long, ByVal connectTimeoutMs As Long, ByVal followRedirects As Boolean, _
        ByVal userAgent As String, ByVal maxFileSize As Double) As Long
    Dim i As Long, wsa(0 To 511) As Byte, rc As Long
    If Not wsaStarted Then
        If ws_WSAStartup(&H202, wsa(0)) = 0 Then wsaStarted = True
    End If
    i = newTransfer()
    xfers(i).method = method
    xfers(i).url = url
    xfers(i).headers = headers
    xfers(i).bodyLen = -1
    If bodyLen >= 0 Then
        xfers(i).bodyLen = bodyLen
        If bodyLen > 0 Then
            ReDim xfers(i).body(0 To bodyLen - 1)
            CopyMemory xfers(i).body(0), body(0), bodyLen
        End If
    End If
    Select Case protocol
    Case 1: xfers(i).family = AF_INET
    Case 2: xfers(i).family = AF_INET6
    Case Else: xfers(i).family = AF_UNSPEC
    End Select
    xfers(i).connectTimeoutMs = connectTimeoutMs
    xfers(i).followRedirects = followRedirects
    xfers(i).userAgent = userAgent
    xfers(i).maxFileSize = maxFileSize
    rc = beginRequest(i)
    If rc <> CURLE_OK Then finishTransfer i, rc, "URL: " & url
    urlTransferStart = i + 1
End Function

' curl_multi_perform for one transfer: True when it is done. waitMs > 0: when nothing can be done, waits up to that
' long for its socket (curl_multi_poll), once.
Public Function urlTransferPoll(ByVal handle As Long, ByVal waitMs As Long) As Boolean
    Dim i As Long, steps As Long, progressed As Boolean
    i = handle - 1
    If i < 0 Or i >= xferCount Then
        urlTransferPoll = True
        Exit Function
    End If
    If xfers(i).phase = XP_FREE Then
        urlTransferPoll = True
        Exit Function
    End If
    Do While xfers(i).phase <> XP_DONE
        progressed = stepTransfer(i)
        steps = steps + 1
        If Not progressed Then
            If waitMs <= 0 Then Exit Do
            waitForSocket i, waitMs
            waitMs = 0
        ElseIf steps >= MAX_STEPS_PER_POLL Then
            Exit Do
        End If
    Loop
    urlTransferPoll = (xfers(i).phase = XP_DONE)
End Function

' the CURLcode, with curl_easy_strerror's text, CURLINFO_RESPONSE_CODE, CURLINFO_PRIMARY_IP and CURLINFO_RETRY_AFTER
Public Function urlTransferResult(ByVal handle As Long, resultText As String, httpStatus As Long, primaryIp As String, _
        retryAfter As Double) As Long
    Dim i As Long
    i = handle - 1
    If i < 0 Or i >= xferCount Then Exit Function
    resultText = curlStrError(xfers(i).result)
    httpStatus = xfers(i).status
    primaryIp = xfers(i).primaryIp
    retryAfter = xfers(i).retryAfter
    urlTransferResult = xfers(i).result
End Function

' the header lines the header callback got (each ending in vbLf)
Public Function urlTransferHeaders(ByVal handle As Long) As String
    Dim i As Long
    i = handle - 1
    If i < 0 Or i >= xferCount Then Exit Function
    urlTransferHeaders = xfers(i).headerText
End Function

' the body (decoded): body(0 .. n - 1), n returned
Public Function urlTransferBody(ByVal handle As Long, body() As Byte) As Long
    Dim i As Long, n As Long
    i = handle - 1
    If i < 0 Or i >= xferCount Then Exit Function
    n = xfers(i).respLen
    If n > 0 Then
        ReDim body(0 To n - 1)
        CopyMemory body(0), xfers(i).resp(0), n
    End If
    urlTransferBody = n
End Function

' curl_multi_remove_handle + curl_easy_cleanup (a running transfer is aborted)
Public Sub urlTransferClose(ByVal handle As Long)
    Dim i As Long, empty_ As UrlTransfer
    i = handle - 1
    If i < 0 Or i >= xferCount Then Exit Sub
    If xfers(i).phase = XP_FREE Then Exit Sub
    If xfers(i).phase <> XP_DONE Then closeConnection i, False
    xfers(i) = empty_
End Sub

' curl_version()'s place in urlRequestOutputDebugInfo
Public Function urlTransferVersion() As String
    Dim tls As String
    tls = "TLS 1.2"
    If windowsBuild() >= 20348 Then tls = "TLS 1.2, 1.3"
    urlTransferVersion = "(VB) WzNet urltransfer_c: HTTP/1.1 over Winsock, Schannel (" & tls & "), miniz " & mz_version()
End Function

'----------------------------------------------------------------------------------------------------------
' the transfer's steps

Private Function stepTransfer(ByVal i As Long) As Boolean
    Select Case xfers(i).phase
    Case XP_RESOLVE
        stepResolve i
        stepTransfer = True
    Case XP_CONNECT
        stepTransfer = stepConnect(i)
    Case XP_TLS
        stepTransfer = stepTls(i)
    Case XP_SEND
        stepTransfer = stepSend(i)
    Case XP_HEADERS, XP_BODY
        stepTransfer = stepReceive(i)
    End Select
End Function

' the URL's parts; the connect's clock starts
Private Function beginRequest(ByVal i As Long) As Long
    Dim rc As Long, timeout As Long
    rc = parseUrl(i, xfers(i).url)
    If rc <> CURLE_OK Then
        beginRequest = rc
        Exit Function
    End If
    timeout = xfers(i).connectTimeoutMs
    If timeout <= 0 Then timeout = DEFAULT_CONNECT_TIMEOUT
    xfers(i).connectStart = nowMs()
    xfers(i).connectDeadline = xfers(i).connectStart + timeout
    xfers(i).phase = XP_RESOLVE
End Function

' scheme://[user@]host[:port][/path][?query][#fragment] (http and https only, as CURLOPT_PROTOCOLS_STR)
Private Function parseUrl(ByVal i As Long, ByVal url As String) As Long
    Dim p As Long, scheme As String, rest As String, authority As String, k As Long, c As String
    Dim host As String, portStr As String, defaultPort As Long, target As String
    p = InStr(url, "://")
    If p = 0 Then
        scheme = "http"                         ' (curl guesses http for a URL without a scheme)
        rest = url
    Else
        scheme = LCase$(Left$(url, p - 1))
        rest = Mid$(url, p + 3)
    End If
    Select Case scheme
    Case "https": defaultPort = 443
    Case "http": defaultPort = 80
    Case Else
        parseUrl = CURLE_UNSUPPORTED_PROTOCOL
        Exit Function
    End Select
    ' the authority ends at the first "/", "?" or "#"
    For k = 1 To Len(rest)
        c = Mid$(rest, k, 1)
        If c = "/" Or c = "?" Or c = "#" Then Exit For
    Next
    authority = Left$(rest, k - 1)
    target = Mid$(rest, k)
    p = InStrRev(authority, "@")
    If p > 0 Then authority = Mid$(authority, p + 1)
    If Left$(authority, 1) = "[" Then
        p = InStr(authority, "]")
        If p = 0 Then
            parseUrl = CURLE_URL_MALFORMAT
            Exit Function
        End If
        host = Mid$(authority, 2, p - 2)
        If Mid$(authority, p + 1, 1) = ":" Then
            portStr = Mid$(authority, p + 2)
        ElseIf p < Len(authority) Then
            parseUrl = CURLE_URL_MALFORMAT
            Exit Function
        End If
    Else
        p = InStr(authority, ":")
        If p > 0 Then
            host = Left$(authority, p - 1)
            portStr = Mid$(authority, p + 1)
        Else
            host = authority
        End If
    End If
    If Len(host) = 0 Then
        parseUrl = CURLE_URL_MALFORMAT
        Exit Function
    End If
    xfers(i).port = defaultPort
    If Len(portStr) > 0 Then
        If Len(portStr) > 5 Or Not (portStr Like String$(Len(portStr), "#")) Then
            parseUrl = CURLE_URL_MALFORMAT
            Exit Function
        End If
        xfers(i).port = CLng(portStr)
        If xfers(i).port < 1 Or xfers(i).port > 65535 Then
            parseUrl = CURLE_URL_MALFORMAT
            Exit Function
        End If
    End If
    p = InStr(target, "#")
    If p > 0 Then target = Left$(target, p - 1)
    If Left$(target, 1) = "?" Then target = "/" & target
    If Len(target) = 0 Then target = "/"
    xfers(i).https = (scheme = "https")
    xfers(i).host = host
    xfers(i).target = target
    If InStr(host, ":") > 0 Then
        xfers(i).hostHeader = "[" & host & "]"
    Else
        xfers(i).hostHeader = host
    End If
    If xfers(i).port <> defaultPort Then xfers(i).hostHeader = xfers(i).hostHeader & ":" & xfers(i).port
End Function

' getaddrinfo with CURLOPT_IPRESOLVE's family; the happy eyeballs' two families
Private Sub stepResolve(ByVal i As Long)
    Dim hints As ADDRINFOA, res As Long, p As Long, ai As ADDRINFOA, rc As Long, n As Long, k As Long
    hints.ai_family = xfers(i).family
    hints.ai_socktype = SOCK_STREAM
    hints.ai_protocol = IPPROTO_TCP
    rc = ws_getaddrinfo(xfers(i).host, CStr(xfers(i).port), hints, res)
    If rc <> 0 Or res = 0 Then
        finishTransfer i, CURLE_COULDNT_RESOLVE_HOST, "getaddrinfo " & familyName(xfers(i).family) & ": error " & rc & " for " & xfers(i).host
        Exit Sub
    End If
    p = res
    Do While p <> 0
        CopyMemory ai, ByVal p, LenB(ai)
        If (ai.ai_family = AF_INET Or ai.ai_family = AF_INET6) And ai.ai_addrlen <= SOCKADDR_SIZE And ai.ai_addr <> 0 Then
            ReDim Preserve xfers(i).addrs(0 To (n + 1) * SOCKADDR_SIZE - 1)
            ReDim Preserve xfers(i).addrFam(0 To n)
            ReDim Preserve xfers(i).addrLen(0 To n)
            CopyMemory xfers(i).addrs(n * SOCKADDR_SIZE), ByVal ai.ai_addr, ai.ai_addrlen
            xfers(i).addrFam(n) = ai.ai_family
            xfers(i).addrLen(n) = ai.ai_addrlen
            n = n + 1
        End If
        p = ai.ai_next
    Loop
    ws_freeaddrinfo res
    xfers(i).addrCount = n
    If n = 0 Then
        finishTransfer i, CURLE_COULDNT_RESOLVE_HOST, "no address of the family " & familyName(xfers(i).family) & " for " & xfers(i).host
        Exit Sub
    End If
    ' the first address's family first, the other one HAPPY_EYEBALLS_TIMEOUT later
    xfers(i).balFam(0) = xfers(i).addrFam(0)
    xfers(i).balFam(1) = 0
    For k = 1 To n - 1
        If xfers(i).addrFam(k) <> xfers(i).balFam(0) Then
            xfers(i).balFam(1) = xfers(i).addrFam(k)
            Exit For
        End If
    Next
    For k = 0 To 1
        xfers(i).balNext(k) = 0
        xfers(i).balSock(k) = INVALID_SOCKET
        xfers(i).balState(k) = 0
    Next
    xfers(i).phase = XP_CONNECT
End Sub

Private Function stepConnect(ByVal i As Long) As Boolean
    Dim b As Long, now As Double, won As Long, st As Long
    now = nowMs()
    won = -1
    For b = 0 To 1
        If xfers(i).balFam(b) <> 0 Then
            If xfers(i).balState(b) = 0 Then
                If b = 0 Or now - xfers(i).connectStart >= HAPPY_EYEBALLS_TIMEOUT Or xfers(i).balState(0) = 2 Then
                    startAttempt i, b, now
                    stepConnect = True
                End If
            End If
            If xfers(i).balState(b) = 1 Then
                st = attemptStatus(xfers(i).balSock(b))
                If st = 1 Then
                    won = b
                    Exit For
                ElseIf st = -1 Or now >= xfers(i).balDeadline(b) Then
                    closeSocket xfers(i).balSock(b)
                    startAttempt i, b, now
                    stepConnect = True
                End If
            End If
        End If
    Next
    If won >= 0 Then
        connected i, won
        stepConnect = True
        Exit Function
    End If
    If (xfers(i).balFam(0) = 0 Or xfers(i).balState(0) = 2) And (xfers(i).balFam(1) = 0 Or xfers(i).balState(1) = 2) Then
        finishTransfer i, CURLE_COULDNT_CONNECT, "no address of " & xfers(i).host & " port " & xfers(i).port & " accepted the connection"
        stepConnect = True
    ElseIf now >= xfers(i).connectDeadline Then
        finishTransfer i, CURLE_OPERATION_TIMEDOUT, "Connection timed out after " & CLng(now - xfers(i).connectStart) & " milliseconds"
        stepConnect = True
    End If
End Function

' the family's next address: a non-blocking socket connecting (curl's sockopt_callback: not inherited; TCP_NODELAY,
' CURLOPT_TCP_NODELAY's default)
Private Sub startAttempt(ByVal i As Long, ByVal b As Long, ByVal now As Double)
    Dim k As Long, s As Long, one As Long, rc As Long, e As Long, k2 As Long, more As Boolean, remaining As Double
    Do
        k = -1
        For k2 = xfers(i).balNext(b) To xfers(i).addrCount - 1
            If xfers(i).addrFam(k2) = xfers(i).balFam(b) Then
                k = k2
                Exit For
            End If
        Next
        If k < 0 Then
            xfers(i).balState(b) = 2
            Exit Sub
        End If
        xfers(i).balNext(b) = k + 1
        s = ws_socket(xfers(i).addrFam(k), SOCK_STREAM, IPPROTO_TCP)
        If s <> INVALID_SOCKET Then
            SetHandleInformation s, HANDLE_FLAG_INHERIT, 0
            one = 1
            ws_setsockopt s, IPPROTO_TCP, TCP_NODELAY, one, 4
            one = 1
            ws_ioctlsocket s, FIONBIO, one
            rc = ws_connect(s, xfers(i).addrs(k * SOCKADDR_SIZE), xfers(i).addrLen(k))
            e = Err.LastDllError
            If rc = 0 Or e = WSAEWOULDBLOCK Then
                more = False
                For k2 = k + 1 To xfers(i).addrCount - 1
                    If xfers(i).addrFam(k2) = xfers(i).balFam(b) Then more = True
                Next
                remaining = xfers(i).connectDeadline - now
                If more Then
                    xfers(i).balDeadline(b) = now + remaining / 2
                Else
                    xfers(i).balDeadline(b) = xfers(i).connectDeadline
                End If
                xfers(i).balSock(b) = s
                xfers(i).balAddr(b) = k
                xfers(i).balState(b) = 1
                Exit Sub
            End If
            debug_ 0, LOG_NET, "urlTransfer", "(VB) connect to %s failed: %d", addrText(i, k), e
            ws_closesocket s
        End If
    Loop
End Sub

' 1: connected, -1: failed, 0: still connecting
Private Function attemptStatus(ByVal s As Long) As Long
    Dim wr As fd_set, ex As fd_set, tv As timeval, n As Long, soErr As Long, optlen As Long
    wr.fd_count = 1
    wr.fd_array(0) = s
    ex.fd_count = 1
    ex.fd_array(0) = s
    n = ws_select(0, ByVal 0&, wr, ex, tv)
    If n = SOCKET_ERROR Then
        attemptStatus = -1
    ElseIf ex.fd_count > 0 Then
        attemptStatus = -1
    ElseIf wr.fd_count > 0 Then
        optlen = 4
        If ws_getsockopt(s, SOL_SOCKET, SO_ERROR, soErr, optlen) = 0 And soErr = 0 Then
            attemptStatus = 1
        Else
            attemptStatus = -1
        End If
    End If
End Function

Private Sub connected(ByVal i As Long, ByVal won As Long)
    Dim other As Long
    other = 1 - won
    xfers(i).sock = xfers(i).balSock(won)
    xfers(i).balSock(won) = INVALID_SOCKET
    If xfers(i).balSock(other) <> INVALID_SOCKET Then closeSocket xfers(i).balSock(other)
    xfers(i).primaryIp = addrText(i, xfers(i).balAddr(won))
    debug_ 0, LOG_NET, "urlTransfer", "(VB) Connected to %s (%s) port %d for %s", xfers(i).primaryIp, _
        familyName(xfers(i).addrFam(xfers(i).balAddr(won))), xfers(i).port, xfers(i).url
    xfers(i).received = 0
    xfers(i).eof = False
    If xfers(i).https Then
        tlsStart i
    Else
        buildRequest i
    End If
End Sub

'----------------------------------------------------------------------------------------------------------
' Schannel (curl's schannel.c: schannel_acquire_credential_handle, schannel_connect_step1/2/3, schannel_send,
' schannel_recv, schannel_shutdown)

Private Function acquireCredentials() As Boolean
    Dim pkg As String, st As Long, expiry As Currency, enabled As Long, flags As Long, build_ As Long
    Dim sc As SCH_CREDENTIALS, tp As TLS_PARAMETERS, oc As SCHANNEL_CRED
    If credAcquired Then
        acquireCredentials = True
        Exit Function
    End If
    build_ = windowsBuild()
    ' curl 8's default: TLS 1.2 at least; TLS 1.3 where Schannel has it built in
    enabled = SP_PROT_TLS1_2_CLIENT
    If build_ >= 20348 Then enabled = enabled Or SP_PROT_TLS1_3_CLIENT
    ' CURLOPT_SSL_VERIFYPEER and VERIFYHOST (on): Schannel checks the chain, the name and the revocation
    flags = SCH_CRED_AUTO_CRED_VALIDATION Or SCH_CRED_REVOCATION_CHECK_CHAIN Or SCH_CRED_NO_DEFAULT_CREDS Or SCH_USE_STRONG_CRYPTO
    pkg = UNISP_NAME
    If build_ >= 17763 Then
        tp.grbitDisabledProtocols = Not enabled
        sc.dwVersion = SCH_CREDENTIALS_VERSION
        sc.dwFlags = flags
        sc.cTlsParameters = 1
        sc.pTlsParameters = VarPtr(tp)
        st = AcquireCredentialsHandleW(0, StrPtr(pkg), SECPKG_CRED_OUTBOUND, 0, sc, 0, 0, cred, expiry)
    Else
        oc.dwVersion = SCHANNEL_CRED_VERSION
        oc.grbitEnabledProtocols = enabled
        oc.dwFlags = flags
        st = AcquireCredentialsHandleW(0, StrPtr(pkg), SECPKG_CRED_OUTBOUND, 0, oc, 0, 0, cred, expiry)
    End If
    If st <> SEC_E_OK Then
        debug_ 0, LOG_NET, "urlTransfer", "(VB) schannel: AcquireCredentialsHandle failed: 0x%08x", st
        Exit Function
    End If
    credTls13 = (enabled And SP_PROT_TLS1_3_CLIENT) <> 0
    credAcquired = True
    acquireCredentials = True
End Function

' schannel_connect_step1: the first InitializeSecurityContext (the ClientHello), with ALPN offering http/1.1 (curl's
' CURLOPT_SSL_ENABLE_ALPN default; 4.7.0's libcurl has no HTTP/2)
Private Sub tlsStart(ByVal i As Long)
    Dim outBufs(0 To 1) As SecBuffer, outDesc As SecBufferDesc, ctx As SecHandle, attr As Long, expiry As Currency
    Dim target As String, st As Long, alpn(0 To 18) As Byte, inBuf As SecBuffer, inDesc As SecBufferDesc, k As Long
    If Not acquireCredentials() Then
        finishTransfer i, CURLE_SSL_CONNECT_ERROR, "schannel: no credentials"
        Exit Sub
    End If
    xfers(i).reqFlags = ISC_REQ_SEQUENCE_DETECT Or ISC_REQ_REPLAY_DETECT Or ISC_REQ_CONFIDENTIALITY Or _
        ISC_REQ_ALLOCATE_MEMORY Or ISC_REQ_STREAM
    ' SEC_APPLICATION_PROTOCOLS: the extension's length (15), ALPN, the list's length (9), the list: "\x08http/1.1"
    alpn(0) = 15
    alpn(4) = SecApplicationProtocolNegotiationExt_ALPN
    alpn(8) = 9
    alpn(10) = 8
    For k = 1 To 8
        alpn(10 + k) = Asc(Mid$("http/1.1", k, 1))
    Next
    inBuf.BufferType = SECBUFFER_APPLICATION_PROTOCOLS
    inBuf.cbBuffer = 19
    inBuf.pvBuffer = VarPtr(alpn(0))
    inDesc.ulVersion = SECBUFFER_VERSION
    inDesc.cBuffers = 1
    inDesc.pBuffers = VarPtr(inBuf)
    outBufs(0).BufferType = SECBUFFER_TOKEN
    outBufs(1).BufferType = SECBUFFER_ALERT
    outDesc.ulVersion = SECBUFFER_VERSION
    outDesc.cBuffers = 2
    outDesc.pBuffers = VarPtr(outBufs(0))
    target = xfers(i).host
    st = InitializeSecurityContextW(cred, 0, StrPtr(target), xfers(i).reqFlags, 0, 0, VarPtr(inDesc), 0, ctx, outDesc, attr, expiry)
    If st <> SEC_I_CONTINUE_NEEDED Then
        freeOutBuffers outBufs
        finishTransfer i, CURLE_SSL_CONNECT_ERROR, "schannel: initial InitializeSecurityContext failed: 0x" & Hex$(st)
        Exit Sub
    End If
    xfers(i).ctx = ctx
    xfers(i).hasCtx = True
    xfers(i).handshakeDone = False
    xfers(i).tlsInLen = 0
    xfers(i).tlsNeedMore = False
    queueOutToken i, outBufs
    xfers(i).phase = XP_TLS
End Sub

' schannel_connect_step2: the handshake's tokens until SEC_E_OK (also a renegotiation's)
Private Function stepTls(ByVal i As Long) As Boolean
    Dim inBufs(0 To 1) As SecBuffer, inDesc As SecBufferDesc, outBufs(0 To 1) As SecBuffer, outDesc As SecBufferDesc
    Dim ctx As SecHandle, attr As Long, expiry As Currency, target As String, st As Long, r As Long
    If Not xfers(i).renegotiating And nowMs() >= xfers(i).connectDeadline Then
        finishTransfer i, CURLE_OPERATION_TIMEDOUT, "SSL connection timeout"
        stepTls = True
        Exit Function
    End If
    ' the pending tokens first
    If xfers(i).sendPos < xfers(i).sendLen Then
        r = flushSend(i, CURLE_SSL_CONNECT_ERROR)
        If r < 0 Then
            stepTls = True
            Exit Function
        End If
        If r > 0 Then stepTls = True
        If xfers(i).sendPos < xfers(i).sendLen Then Exit Function
    End If
    ' the server's answer
    r = recvSome(i, CURLE_SSL_CONNECT_ERROR)
    If r < 0 Then
        stepTls = True
        Exit Function
    End If
    If r > 0 Then
        stepTls = True
        xfers(i).tlsNeedMore = False
    End If
    If xfers(i).tlsInLen = 0 Or xfers(i).tlsNeedMore Then
        If xfers(i).eof Then
            finishTransfer i, CURLE_SSL_CONNECT_ERROR, "schannel: server closed the connection during the handshake"
            stepTls = True
        End If
        Exit Function
    End If
    target = xfers(i).host
    Do
        inBufs(0).BufferType = SECBUFFER_TOKEN
        inBufs(0).cbBuffer = xfers(i).tlsInLen
        inBufs(0).pvBuffer = VarPtr(xfers(i).tlsIn(0))
        inBufs(1).BufferType = SECBUFFER_EMPTY
        inBufs(1).cbBuffer = 0
        inBufs(1).pvBuffer = 0
        inDesc.ulVersion = SECBUFFER_VERSION
        inDesc.cBuffers = 2
        inDesc.pBuffers = VarPtr(inBufs(0))
        outBufs(0).BufferType = SECBUFFER_TOKEN
        outBufs(0).cbBuffer = 0
        outBufs(0).pvBuffer = 0
        outBufs(1).BufferType = SECBUFFER_ALERT
        outBufs(1).cbBuffer = 0
        outBufs(1).pvBuffer = 0
        outDesc.ulVersion = SECBUFFER_VERSION
        outDesc.cBuffers = 2
        outDesc.pBuffers = VarPtr(outBufs(0))
        ctx = xfers(i).ctx
        st = InitializeSecurityContextW(cred, VarPtr(ctx), StrPtr(target), xfers(i).reqFlags, 0, 0, VarPtr(inDesc), 0, ctx, outDesc, attr, expiry)
        xfers(i).ctx = ctx
        stepTls = True
        Select Case st
        Case SEC_E_INCOMPLETE_MESSAGE
            freeOutBuffers outBufs
            xfers(i).tlsNeedMore = True
            Exit Function
        Case SEC_I_INCOMPLETE_CREDENTIALS
            ' a client certificate has been requested: go on without one
            freeOutBuffers outBufs
            If (xfers(i).reqFlags And ISC_REQ_USE_SUPPLIED_CREDS) <> 0 Then
                finishTransfer i, CURLE_SSL_CONNECT_ERROR, "schannel: the server wants a client certificate"
                Exit Function
            End If
            xfers(i).reqFlags = xfers(i).reqFlags Or ISC_REQ_USE_SUPPLIED_CREDS
        Case SEC_I_CONTINUE_NEEDED, SEC_E_OK
            queueOutToken i, outBufs
            If inBufs(1).BufferType = SECBUFFER_EXTRA And inBufs(1).cbBuffer > 0 Then
                bufConsume xfers(i).tlsIn, xfers(i).tlsInLen, xfers(i).tlsInLen - inBufs(1).cbBuffer
            Else
                xfers(i).tlsInLen = 0
            End If
            If st = SEC_E_OK Then
                handshakeComplete i
                Exit Function
            End If
            If xfers(i).tlsInLen = 0 Then Exit Function
        Case SEC_E_WRONG_PRINCIPAL
            freeOutBuffers outBufs
            finishTransfer i, CURLE_PEER_FAILED_VERIFICATION, "schannel: SNI or certificate check failed: 0x" & Hex$(st)
            Exit Function
        Case Else
            freeOutBuffers outBufs
            finishTransfer i, CURLE_SSL_CONNECT_ERROR, "schannel: next InitializeSecurityContext failed: 0x" & Hex$(st)
            Exit Function
        End Select
    Loop
End Function

' schannel_connect_step3: the stream sizes; the request (or, after a renegotiation, the phase it interrupted)
Private Sub handshakeComplete(ByVal i As Long)
    Dim ctx As SecHandle, sizes As SecPkgContext_StreamSizes, info As SecPkgContext_ConnectionInfo
    ctx = xfers(i).ctx
    If QueryContextAttributesW(ctx, SECPKG_ATTR_STREAM_SIZES, sizes) <> SEC_E_OK Then
        finishTransfer i, CURLE_SSL_CONNECT_ERROR, "schannel: QueryContextAttributes (stream sizes) failed"
        Exit Sub
    End If
    xfers(i).sizes = sizes
    xfers(i).handshakeDone = True
    If xfers(i).renegotiating Then
        xfers(i).renegotiating = False
        xfers(i).phase = xfers(i).savedPhase
        Exit Sub
    End If
    If QueryContextAttributesW(ctx, SECPKG_ATTR_CONNECTION_INFO, info) = SEC_E_OK Then
        debug_ 0, LOG_NET, "urlTransfer", "(VB) schannel: %s with %s", protocolName(info.dwProtocol), xfers(i).host
    End If
    buildRequest i
End Sub

' the output token (sent) and the alert buffers (freed)
Private Sub queueOutToken(ByVal i As Long, outBufs() As SecBuffer)
    If outBufs(0).cbBuffer > 0 And outBufs(0).pvBuffer <> 0 Then
        bufAppendPtr xfers(i).sendBuf, xfers(i).sendLen, outBufs(0).pvBuffer, outBufs(0).cbBuffer
    End If
    freeOutBuffers outBufs
End Sub

Private Sub freeOutBuffers(outBufs() As SecBuffer)
    Dim k As Long
    For k = LBound(outBufs) To UBound(outBufs)
        If outBufs(k).pvBuffer <> 0 Then FreeContextBuffer outBufs(k).pvBuffer
        outBufs(k).pvBuffer = 0
        outBufs(k).cbBuffer = 0
    Next
End Sub

' schannel_send: data(ofs .. ofs + n - 1) as TLS records, queued
Private Function encryptQueue(ByVal i As Long, data() As Byte, ByVal ofs As Long, ByVal n As Long) As Boolean
    Dim m() As Byte, bufs(0 To 3) As SecBuffer, desc As SecBufferDesc, ctx As SecHandle, chunk As Long
    Dim st As Long, hdr As Long, trl As Long, total As Long
    hdr = xfers(i).sizes.cbHeader
    trl = xfers(i).sizes.cbTrailer
    Do While n > 0
        chunk = n
        If chunk > xfers(i).sizes.cbMaximumMessage Then chunk = xfers(i).sizes.cbMaximumMessage
        ReDim m(0 To hdr + chunk + trl - 1)
        CopyMemory m(hdr), data(ofs), chunk
        bufs(0).BufferType = SECBUFFER_STREAM_HEADER
        bufs(0).cbBuffer = hdr
        bufs(0).pvBuffer = VarPtr(m(0))
        bufs(1).BufferType = SECBUFFER_DATA
        bufs(1).cbBuffer = chunk
        bufs(1).pvBuffer = VarPtr(m(hdr))
        bufs(2).BufferType = SECBUFFER_STREAM_TRAILER
        bufs(2).cbBuffer = trl
        bufs(2).pvBuffer = VarPtr(m(hdr + chunk))
        bufs(3).BufferType = SECBUFFER_EMPTY
        bufs(3).cbBuffer = 0
        bufs(3).pvBuffer = 0
        desc.ulVersion = SECBUFFER_VERSION
        desc.cBuffers = 4
        desc.pBuffers = VarPtr(bufs(0))
        ctx = xfers(i).ctx
        st = EncryptMessage(ctx, 0, desc, 0)
        If st <> SEC_E_OK Then
            finishTransfer i, CURLE_SEND_ERROR, "schannel: EncryptMessage failed: 0x" & Hex$(st)
            Exit Function
        End If
        total = bufs(0).cbBuffer + bufs(1).cbBuffer + bufs(2).cbBuffer
        bufAppendPtr xfers(i).sendBuf, xfers(i).sendLen, VarPtr(m(0)), total
        ofs = ofs + chunk
        n = n - chunk
    Loop
    encryptQueue = True
End Function

' schannel_recv: the records received so far, decrypted into plain (False: the transfer failed)
Private Function decryptAll(ByVal i As Long) As Boolean
    Dim bufs(0 To 3) As SecBuffer, desc As SecBufferDesc, ctx As SecHandle, st As Long, k As Long, extra As Long
    Do While xfers(i).tlsInLen > 0
        bufs(0).BufferType = SECBUFFER_DATA
        bufs(0).cbBuffer = xfers(i).tlsInLen
        bufs(0).pvBuffer = VarPtr(xfers(i).tlsIn(0))
        For k = 1 To 3
            bufs(k).BufferType = SECBUFFER_EMPTY
            bufs(k).cbBuffer = 0
            bufs(k).pvBuffer = 0
        Next
        desc.ulVersion = SECBUFFER_VERSION
        desc.cBuffers = 4
        desc.pBuffers = VarPtr(bufs(0))
        ctx = xfers(i).ctx
        st = DecryptMessage(ctx, desc, 0, 0)
        xfers(i).ctx = ctx
        Select Case st
        Case SEC_E_OK, SEC_I_RENEGOTIATE
            extra = 0
            For k = 0 To 3
                If bufs(k).BufferType = SECBUFFER_DATA And bufs(k).cbBuffer > 0 Then
                    bufAppendPtr xfers(i).plain, xfers(i).plainLen, bufs(k).pvBuffer, bufs(k).cbBuffer
                ElseIf bufs(k).BufferType = SECBUFFER_EXTRA Then
                    extra = bufs(k).cbBuffer
                End If
            Next
            If extra > 0 Then
                bufConsume xfers(i).tlsIn, xfers(i).tlsInLen, xfers(i).tlsInLen - extra
            Else
                xfers(i).tlsInLen = 0
            End If
            If st = SEC_I_RENEGOTIATE Then
                ' the extra bytes are a handshake message for InitializeSecurityContext (TLS 1.3's tickets, a renegotiation)
                xfers(i).renegotiating = True
                xfers(i).savedPhase = xfers(i).phase
                xfers(i).phase = XP_TLS
                xfers(i).tlsNeedMore = False
                decryptAll = True
                Exit Function
            End If
        Case SEC_E_INCOMPLETE_MESSAGE
            xfers(i).tlsNeedMore = True
            Exit Do
        Case SEC_I_CONTEXT_EXPIRED
            ' the server's close_notify
            xfers(i).tlsInLen = 0
            xfers(i).eof = True
            Exit Do
        Case Else
            finishTransfer i, CURLE_RECV_ERROR, "schannel: failed to decrypt data: 0x" & Hex$(st)
            Exit Function
        End Select
    Loop
    decryptAll = True
End Function

' schannel_shutdown: the close_notify alert (best effort: one non-blocking send)
Private Sub sendCloseNotify(ByVal i As Long)
    Dim tok As Long, inBufs(0 To 0) As SecBuffer, inDesc As SecBufferDesc, outBufs(0 To 0) As SecBuffer, outDesc As SecBufferDesc
    Dim ctx As SecHandle, attr As Long, expiry As Currency, target As String, st As Long
    tok = SCHANNEL_SHUTDOWN
    inBufs(0).BufferType = SECBUFFER_TOKEN
    inBufs(0).cbBuffer = 4
    inBufs(0).pvBuffer = VarPtr(tok)
    inDesc.ulVersion = SECBUFFER_VERSION
    inDesc.cBuffers = 1
    inDesc.pBuffers = VarPtr(inBufs(0))
    ctx = xfers(i).ctx
    If ApplyControlToken(ctx, inDesc) <> SEC_E_OK Then Exit Sub
    outBufs(0).BufferType = SECBUFFER_TOKEN
    outDesc.ulVersion = SECBUFFER_VERSION
    outDesc.cBuffers = 1
    outDesc.pBuffers = VarPtr(outBufs(0))
    target = xfers(i).host
    st = InitializeSecurityContextW(cred, VarPtr(ctx), StrPtr(target), xfers(i).reqFlags, 0, 0, 0, 0, ctx, outDesc, attr, expiry)
    If (st = SEC_E_OK Or st = SEC_I_CONTEXT_EXPIRED) And outBufs(0).cbBuffer > 0 And outBufs(0).pvBuffer <> 0 Then
        ws_send xfers(i).sock, ByVal outBufs(0).pvBuffer, outBufs(0).cbBuffer, 0
    End If
    freeOutBuffers outBufs
    xfers(i).ctx = ctx
End Sub

'----------------------------------------------------------------------------------------------------------
' HTTP/1.1

' the request as curl's Curl_http writes it, queued (encrypted for https)
Private Sub buildRequest(ByVal i As Long)
    Dim req As String, b() As Byte, n As Long, hdrs As String
    hdrs = xfers(i).headers
    req = xfers(i).method & " " & xfers(i).target & " HTTP/1.1" & vbCrLf
    If Not hasHeader(hdrs, "Host") Then req = req & "Host: " & xfers(i).hostHeader & vbCrLf
    If Len(xfers(i).userAgent) <> 0 And Not hasHeader(hdrs, "User-Agent") Then req = req & "User-Agent: " & xfers(i).userAgent & vbCrLf
    If Not hasHeader(hdrs, "Accept") Then req = req & "Accept: */*" & vbCrLf
    If Not hasHeader(hdrs, "Accept-Encoding") Then req = req & "Accept-Encoding: " & ACCEPT_ENCODINGS & vbCrLf
    req = req & hdrs
    If xfers(i).bodyLen >= 0 Then
        If Not hasHeader(hdrs, "Content-Length") Then req = req & "Content-Length: " & xfers(i).bodyLen & vbCrLf
        If Not hasHeader(hdrs, "Content-Type") Then req = req & "Content-Type: application/x-www-form-urlencoded" & vbCrLf
    End If
    req = req & vbCrLf
    n = utf8_encode(req, b)
    If xfers(i).bodyLen > 0 Then
        ReDim Preserve b(0 To n + xfers(i).bodyLen - 1)
        CopyMemory b(n), xfers(i).body(0), xfers(i).bodyLen
        n = n + xfers(i).bodyLen
    End If
    If xfers(i).https Then
        If Not encryptQueue(i, b, 0, n) Then Exit Sub
    Else
        bufAppendPtr xfers(i).sendBuf, xfers(i).sendLen, VarPtr(b(0)), n
    End If
    ' the response's state
    xfers(i).plainLen = 0
    xfers(i).contentLength = -1
    xfers(i).chunked = False
    xfers(i).contentEncoding = ""
    xfers(i).location = ""
    xfers(i).respLen = 0
    xfers(i).bodyMode = BM_NONE
    ' the low-speed check starts (curl's PERFORMING state)
    xfers(i).total = 0
    xfers(i).sampleCount = 0
    xfers(i).lowSpeedSince = 0
    xfers(i).phase = XP_SEND
End Sub

' Curl_checkheaders: a header of that name in the request's own ("Name:" or "Name;", any case)
Private Function hasHeader(ByVal hdrs As String, ByVal name_ As String) As Boolean
    Dim lines() As String, k As Long, l As String
    If Len(hdrs) = 0 Then Exit Function
    lines = Split(hdrs, vbCrLf)
    For k = 0 To UBound(lines)
        l = lines(k)
        If Len(l) > Len(name_) Then
            If LCase$(Left$(l, Len(name_))) = LCase$(name_) Then
                If Mid$(l, Len(name_) + 1, 1) = ":" Or Mid$(l, Len(name_) + 1, 1) = ";" Then
                    hasHeader = True
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Private Function stepSend(ByVal i As Long) As Boolean
    Dim r As Long
    If lowSpeedCheck(i) Then
        stepSend = True
        Exit Function
    End If
    r = flushSend(i, CURLE_SEND_ERROR)
    If r <> 0 Then stepSend = True
    If r < 0 Then Exit Function
    If xfers(i).sendLen = 0 Then
        xfers(i).phase = XP_HEADERS
        stepSend = True
    End If
End Function

Private Function stepReceive(ByVal i As Long) As Boolean
    Dim r As Long, parsed As Boolean
    If lowSpeedCheck(i) Then
        stepReceive = True
        Exit Function
    End If
    r = recvSome(i, CURLE_RECV_ERROR)
    If r < 0 Then
        stepReceive = True
        Exit Function
    End If
    If r > 0 Then
        stepReceive = True
        xfers(i).tlsNeedMore = False
    End If
    If xfers(i).https And xfers(i).tlsInLen > 0 And Not xfers(i).tlsNeedMore Then
        If Not decryptAll(i) Then
            stepReceive = True
            Exit Function
        End If
        If xfers(i).phase = XP_TLS Then
            stepReceive = True
            Exit Function
        End If
    End If
    Do
        parsed = False
        If xfers(i).phase = XP_HEADERS Then
            parsed = parseHeaders(i)
        ElseIf xfers(i).phase = XP_BODY Then
            parsed = parseBody(i)
        End If
        If parsed Then stepReceive = True
    Loop While parsed And (xfers(i).phase = XP_HEADERS Or xfers(i).phase = XP_BODY)
    If xfers(i).eof And (xfers(i).phase = XP_HEADERS Or xfers(i).phase = XP_BODY) Then
        stepReceive = True
        If xfers(i).phase = XP_HEADERS Then
            If xfers(i).received = 0 Then
                finishTransfer i, CURLE_GOT_NOTHING, "Empty reply from server"
            Else
                finishTransfer i, CURLE_RECV_ERROR, "the connection closed in the response's headers"
            End If
        ElseIf xfers(i).bodyMode = BM_CLOSE Then
            completeTransfer i
        ElseIf xfers(i).bodyMode = BM_LENGTH Then
            finishTransfer i, CURLE_PARTIAL_FILE, "transfer closed with " & xfers(i).left_ & " bytes remaining to read"
        Else
            finishTransfer i, CURLE_PARTIAL_FILE, "transfer closed with outstanding read data remaining"
        End If
    End If
End Function

' a response's headers, when they are all in (False: not yet)
Private Function parseHeaders(ByVal i As Long) As Boolean
    Dim endPos As Long, text As String, lines() As String, k As Long, l As String, c As Long
    Dim name_ As String, value As String, statusLine As String, codeStr As String
    endPos = findHeaderEnd(xfers(i).plain, xfers(i).plainLen)
    If endPos = 0 Then
        If xfers(i).plainLen > MAX_HTTP_RESP_HEADER_SIZE Then
            finishTransfer i, CURLE_TOO_LARGE, "Too large response headers"
            parseHeaders = True
        End If
        Exit Function
    End If
    parseHeaders = True
    text = latin1(xfers(i).plain, 0, endPos)
    bufConsume xfers(i).plain, xfers(i).plainLen, endPos
    lines = Split(Replace(text, vbCrLf, vbLf), vbLf)
    statusLine = lines(0)
    If Left$(statusLine, 5) <> "HTTP/" Then
        finishTransfer i, CURLE_UNSUPPORTED_PROTOCOL, "Received HTTP/0.9 when not allowed"
        Exit Function
    End If
    codeStr = Mid$(statusLine, InStr(statusLine, " ") + 1, 3)
    If InStr(statusLine, " ") = 0 Or Not (codeStr Like "###") Then
        finishTransfer i, CURLE_WEIRD_SERVER_REPLY, "Unsupported response code in HTTP response"
        Exit Function
    End If
    xfers(i).status = CLng(codeStr)
    xfers(i).contentLength = -1
    xfers(i).chunked = False
    xfers(i).contentEncoding = ""
    xfers(i).location = ""
    xfers(i).retryAfter = 0
    xfers(i).headerText = xfers(i).headerText & statusLine & vbLf
    For k = 1 To UBound(lines)
        l = lines(k)
        If Len(l) > 0 Then
            xfers(i).headerText = xfers(i).headerText & l & vbLf
            c = InStr(l, ":")
            If c > 0 Then
                name_ = LCase$(Trim$(Left$(l, c - 1)))
                value = Trim$(Replace(Mid$(l, c + 1), vbTab, " "))
                Select Case name_
                Case "content-length"
                    If Len(value) = 0 Or Len(value) > 15 Or Not (value Like String$(Len(value), "#")) Then
                        finishTransfer i, CURLE_WEIRD_SERVER_REPLY, "Invalid Content-Length: " & value
                        Exit Function
                    End If
                    xfers(i).contentLength = CDbl(value)
                Case "transfer-encoding"
                    If InStr(LCase$(value), "chunked") > 0 Then xfers(i).chunked = True
                Case "content-encoding"
                    xfers(i).contentEncoding = value
                Case "location"
                    xfers(i).location = value
                Case "retry-after"
                    xfers(i).retryAfter = parseRetryAfter(value)
                End Select
            End If
        End If
    Next
    ' 1xx: the final response follows
    If xfers(i).status >= 100 And xfers(i).status < 200 And xfers(i).status <> 101 Then Exit Function
    ' CURLOPT_FOLLOWLOCATION
    If xfers(i).followRedirects And Len(xfers(i).location) > 0 Then
        Select Case xfers(i).status
        Case 301, 302, 303, 307, 308
            redirect i
            Exit Function
        End Select
    End If
    ' the body's framing
    If xfers(i).status = 204 Or xfers(i).status = 304 Or xfers(i).status = 101 Then
        xfers(i).bodyMode = BM_NONE
        completeTransfer i
    ElseIf xfers(i).chunked Then
        xfers(i).bodyMode = BM_CHUNKED
        xfers(i).chunkState = CH_SIZE
        xfers(i).phase = XP_BODY
    ElseIf xfers(i).contentLength >= 0 Then
        If xfers(i).maxFileSize > 0 And xfers(i).contentLength > xfers(i).maxFileSize Then
            finishTransfer i, CURLE_FILESIZE_EXCEEDED, "Maximum file size exceeded (Content-Length " & xfers(i).contentLength & ")"
            Exit Function
        End If
        xfers(i).bodyMode = BM_LENGTH
        xfers(i).left_ = xfers(i).contentLength
        xfers(i).phase = XP_BODY
        If xfers(i).left_ = 0 Then completeTransfer i
    Else
        xfers(i).bodyMode = BM_CLOSE
        xfers(i).phase = XP_BODY
    End If
End Function

' the offset after the empty line ending the headers (0: not in yet)
Private Function findHeaderEnd(b() As Byte, ByVal n As Long) As Long
    Dim k As Long
    For k = 0 To n - 2
        If b(k) = 10 Then
            If b(k + 1) = 10 Then
                findHeaderEnd = k + 2
                Exit Function
            End If
            If k + 2 < n Then
                If b(k + 1) = 13 And b(k + 2) = 10 Then
                    findHeaderEnd = k + 3
                    Exit Function
                End If
            End If
        End If
    Next
End Function

' the body as it arrives (False: nothing to do yet)
Private Function parseBody(ByVal i As Long) As Boolean
    Dim n As Long, l As String, sizeStr As String, p As Long, size As Double, k As Long, c As String
    Select Case xfers(i).bodyMode
    Case BM_LENGTH
        If xfers(i).plainLen = 0 Then Exit Function
        n = xfers(i).plainLen
        If n > xfers(i).left_ Then n = xfers(i).left_
        If Not appendBody(i, n) Then
            parseBody = True
            Exit Function
        End If
        xfers(i).left_ = xfers(i).left_ - n
        If xfers(i).left_ = 0 Then completeTransfer i
        parseBody = True
    Case BM_CLOSE
        If xfers(i).plainLen = 0 Then Exit Function
        appendBody i, xfers(i).plainLen
        parseBody = True
    Case BM_CHUNKED
        Select Case xfers(i).chunkState
        Case CH_SIZE
            If Not takeLine(i, l) Then Exit Function
            parseBody = True
            p = InStr(l, ";")
            If p > 0 Then l = Left$(l, p - 1)
            sizeStr = Trim$(l)
            If Len(sizeStr) = 0 Or Len(sizeStr) > 15 Then
                finishTransfer i, CURLE_RECV_ERROR, "Malformed encoding found in chunked-encoding"
                Exit Function
            End If
            size = 0
            For k = 1 To Len(sizeStr)
                c = UCase$(Mid$(sizeStr, k, 1))
                If c Like "[0-9A-F]" Then
                    size = size * 16 + CLng("&H" & c)
                Else
                    finishTransfer i, CURLE_RECV_ERROR, "Malformed encoding found in chunked-encoding"
                    Exit Function
                End If
            Next
            If size = 0 Then
                xfers(i).chunkState = CH_TRAILER
            Else
                xfers(i).left_ = size
                xfers(i).chunkState = CH_DATA
            End If
        Case CH_DATA
            If xfers(i).plainLen = 0 Then Exit Function
            n = xfers(i).plainLen
            If n > xfers(i).left_ Then n = xfers(i).left_
            parseBody = True
            If Not appendBody(i, n) Then Exit Function
            xfers(i).left_ = xfers(i).left_ - n
            If xfers(i).left_ = 0 Then xfers(i).chunkState = CH_DATA_END
        Case CH_DATA_END
            If Not takeLine(i, l) Then Exit Function
            parseBody = True
            If Len(l) <> 0 Then
                finishTransfer i, CURLE_RECV_ERROR, "Malformed encoding found in chunked-encoding"
                Exit Function
            End If
            xfers(i).chunkState = CH_SIZE
        Case CH_TRAILER
            If Not takeLine(i, l) Then Exit Function
            parseBody = True
            If Len(l) = 0 Then
                completeTransfer i
            Else
                xfers(i).headerText = xfers(i).headerText & l & vbLf
            End If
        End Select
    End Select
End Function

' the first n bytes of plain to the body (CURLOPT_MAXFILESIZE_LARGE: False when the body grew past it)
Private Function appendBody(ByVal i As Long, ByVal n As Long) As Boolean
    bufAppendPtr xfers(i).resp, xfers(i).respLen, VarPtr(xfers(i).plain(0)), n
    bufConsume xfers(i).plain, xfers(i).plainLen, n
    If xfers(i).maxFileSize > 0 And xfers(i).respLen > xfers(i).maxFileSize Then
        finishTransfer i, CURLE_FILESIZE_EXCEEDED, "Exceeded the maximum allowed file size (" & xfers(i).maxFileSize & ")"
        Exit Function
    End If
    appendBody = True
End Function

' a line of plain (without its CR LF), taken (False: not in yet)
Private Function takeLine(ByVal i As Long, l As String) As Boolean
    Dim k As Long, n As Long
    For k = 0 To xfers(i).plainLen - 1
        If xfers(i).plain(k) = 10 Then
            n = k
            If n > 0 Then
                If xfers(i).plain(n - 1) = 13 Then n = n - 1
            End If
            l = latin1(xfers(i).plain, 0, n)
            bufConsume xfers(i).plain, xfers(i).plainLen, k + 1
            takeLine = True
            Exit Function
        End If
    Next
End Function

' Curl_follow: the Location relative to the URL; the method as curl 8 keeps it with a CURLOPT_CUSTOMREQUEST
' (http_switch_to_get: "Stick to %s instead of GET" drops the body only)
Private Sub redirect(ByVal i As Long)
    Dim newUrl As String, rc As Long
    xfers(i).redirects = xfers(i).redirects + 1
    If xfers(i).redirects > MAXREDIRS Then
        finishTransfer i, CURLE_TOO_MANY_REDIRECTS, "Maximum (" & MAXREDIRS & ") redirects followed"
        Exit Sub
    End If
    newUrl = resolveLocation(i, xfers(i).location)
    Select Case xfers(i).status
    Case 301, 302, 303
        xfers(i).bodyLen = -1
    End Select
    debug_ 0, LOG_NET, "urlTransfer", "(VB) %d: following the redirect to %s", xfers(i).status, newUrl
    closeConnection i, True
    xfers(i).url = newUrl
    rc = beginRequest(i)
    If rc <> CURLE_OK Then finishTransfer i, rc, "redirect to " & newUrl
End Sub

Private Function resolveLocation(ByVal i As Long, ByVal loc As String) As String
    Dim base As String, path As String, p As Long
    If xfers(i).https Then base = "https://" Else base = "http://"
    p = InStr(loc, "://")
    If p > 0 And p < InStr(loc & "/", "/") Then
        resolveLocation = loc
    ElseIf Left$(loc, 2) = "//" Then
        resolveLocation = Left$(base, Len(base) - 2) & loc
    ElseIf Left$(loc, 1) = "/" Then
        resolveLocation = base & xfers(i).hostHeader & loc
    Else
        path = xfers(i).target
        p = InStr(path, "?")
        If p > 0 Then path = Left$(path, p - 1)
        If Left$(loc, 1) = "?" Then
            resolveLocation = base & xfers(i).hostHeader & path & loc
        Else
            resolveLocation = base & xfers(i).hostHeader & Left$(path, InStrRev(path, "/")) & loc
        End If
    End If
End Function

' the response is complete: CURLOPT_HTTP_CONTENT_DECODING (the codings undone, the last first)
Private Sub completeTransfer(ByVal i As Long)
    Dim codings() As String, k As Long, c As String, hs As Long
    If Len(xfers(i).contentEncoding) > 0 And xfers(i).respLen > 0 Then
        codings = Split(xfers(i).contentEncoding, ",")
        For k = UBound(codings) To 0 Step -1
            c = LCase$(Trim$(codings(k)))
            Select Case c
            Case "identity", "none", ""
            Case "deflate"
                ' zlib's format; curl takes raw deflate too (servers that send it so)
                If Not inflateBody(i, 0, MZ_DEFAULT_WINDOW_BITS) Then
                    If Not inflateBody(i, 0, -MZ_DEFAULT_WINDOW_BITS) Then
                        finishTransfer i, CURLE_BAD_CONTENT_ENCODING, "Error while processing content unencoding: deflate"
                        Exit Sub
                    End If
                End If
            Case "gzip", "x-gzip"
                hs = gzipHeaderSize(xfers(i).resp, xfers(i).respLen)
                If hs = 0 Then
                    finishTransfer i, CURLE_BAD_CONTENT_ENCODING, "Error while processing content unencoding: not a gzip header"
                    Exit Sub
                End If
                If Not inflateBody(i, hs, -MZ_DEFAULT_WINDOW_BITS) Then
                    finishTransfer i, CURLE_BAD_CONTENT_ENCODING, "Error while processing content unencoding: gzip"
                    Exit Sub
                End If
            Case Else
                finishTransfer i, CURLE_BAD_CONTENT_ENCODING, "Unrecognized content encoding type. libcurl understands " & ACCEPT_ENCODINGS & " content encodings."
                Exit Sub
            End Select
        Next
    End If
    finishTransfer i, CURLE_OK, ""
End Sub

' RFC 1952's member header: its size (0: not gzip)
Private Function gzipHeaderSize(b() As Byte, ByVal n As Long) As Long
    Dim flg As Long, p As Long
    If n < 10 Then Exit Function
    If b(0) <> &H1F Or b(1) <> &H8B Or b(2) <> 8 Then Exit Function
    flg = b(3)
    p = 10
    If (flg And 4) <> 0 Then                    ' FEXTRA
        If p + 2 > n Then Exit Function
        p = p + 2 + b(p) + 256& * b(p + 1)
    End If
    If (flg And 8) <> 0 Then                    ' FNAME
        Do While p < n
            p = p + 1
            If b(p - 1) = 0 Then Exit Do
        Loop
    End If
    If (flg And 16) <> 0 Then                   ' FCOMMENT
        Do While p < n
            p = p + 1
            If b(p - 1) = 0 Then Exit Do
        Loop
    End If
    If (flg And 2) <> 0 Then p = p + 2          ' FHCRC
    If p >= n Then Exit Function
    gzipHeaderSize = p
End Function

' the body from ofs, inflated (windowBits: miniz's, negative for raw deflate) in its place (False: not valid)
Private Function inflateBody(ByVal i As Long, ByVal ofs As Long, ByVal windowBits As Long) As Boolean
    Dim strm As z_stream, outB() As Byte, res() As Byte, resLen As Long, rc As Long, produced As Long
    If ofs >= xfers(i).respLen Then Exit Function
    If mz_inflateInit2(strm, windowBits) <> MZ_OK Then Exit Function
    strm.next_in = ofs
    strm.avail_in = xfers(i).respLen - ofs
    ReDim outB(0 To 65535)
    Do
        strm.next_out = 0
        strm.avail_out = 65536
        rc = inflate(strm, xfers(i).resp, outB, Z_NO_FLUSH)
        produced = 65536 - strm.avail_out
        If produced > 0 Then bufAppendPtr res, resLen, VarPtr(outB(0)), produced
        If rc = Z_STREAM_END Then
            inflateBody = True
            Exit Do
        End If
        If rc <> Z_OK And rc <> Z_BUF_ERROR Then Exit Do
        If produced = 0 Then Exit Do
    Loop
    inflateEnd strm
    If Not inflateBody Then Exit Function
    xfers(i).resp = res
    xfers(i).respLen = resLen
End Function

'----------------------------------------------------------------------------------------------------------
' sockets and buffers

' the queued bytes, sent as far as the socket takes them: the count (-1: the transfer failed with errCode)
Private Function flushSend(ByVal i As Long, ByVal errCode As Long) As Long
    Dim r As Long, e As Long, sent As Long
    Do While xfers(i).sendPos < xfers(i).sendLen
        r = ws_send(xfers(i).sock, xfers(i).sendBuf(xfers(i).sendPos), xfers(i).sendLen - xfers(i).sendPos, 0)
        e = Err.LastDllError
        If r = SOCKET_ERROR Then
            If e = WSAEWOULDBLOCK Then Exit Do
            finishTransfer i, errCode, "send failure: " & e
            flushSend = -1
            Exit Function
        End If
        xfers(i).sendPos = xfers(i).sendPos + r
        xfers(i).total = xfers(i).total + r
        sent = sent + r
    Loop
    If xfers(i).sendPos >= xfers(i).sendLen Then
        xfers(i).sendLen = 0
        xfers(i).sendPos = 0
    End If
    flushSend = sent
End Function

' what the socket has (one recv): the count, 0 (nothing, or the end: eof), -1 (the transfer failed with errCode)
Private Function recvSome(ByVal i As Long, ByVal errCode As Long) As Long
    Dim r As Long, e As Long
    If xfers(i).eof Then Exit Function
    r = ws_recv(xfers(i).sock, rbuf(0), RECV_SIZE, 0)
    e = Err.LastDllError
    If r = SOCKET_ERROR Then
        If e = WSAEWOULDBLOCK Then Exit Function
        finishTransfer i, errCode, "Recv failure: " & e
        recvSome = -1
        Exit Function
    End If
    If r = 0 Then
        xfers(i).eof = True
        Exit Function
    End If
    xfers(i).total = xfers(i).total + r
    xfers(i).received = xfers(i).received + r
    If xfers(i).https Then
        bufAppendPtr xfers(i).tlsIn, xfers(i).tlsInLen, VarPtr(rbuf(0)), r
    Else
        bufAppendPtr xfers(i).plain, xfers(i).plainLen, VarPtr(rbuf(0)), r
    End If
    recvSome = r
End Function

' Curl_speedcheck: True when the transfer failed (too slow)
Private Function lowSpeedCheck(ByVal i As Long) As Boolean
    Dim now As Double, k As Long, speed As Double, dt As Double
    now = nowMs()
    If xfers(i).sampleCount > 0 Then
        If now - xfers(i).sampleTime(xfers(i).sampleCount - 1) < 1000 Then Exit Function
    End If
    If xfers(i).sampleCount = 6 Then
        For k = 0 To 4
            xfers(i).sampleTime(k) = xfers(i).sampleTime(k + 1)
            xfers(i).sampleBytes(k) = xfers(i).sampleBytes(k + 1)
        Next
        xfers(i).sampleCount = 5
    End If
    xfers(i).sampleTime(xfers(i).sampleCount) = now
    xfers(i).sampleBytes(xfers(i).sampleCount) = xfers(i).total
    xfers(i).sampleCount = xfers(i).sampleCount + 1
    If xfers(i).sampleCount < 2 Then Exit Function
    dt = (now - xfers(i).sampleTime(0)) / 1000#
    If dt <= 0 Then Exit Function
    speed = (xfers(i).total - xfers(i).sampleBytes(0)) / dt
    If speed < LOW_SPEED_LIMIT Then
        If xfers(i).lowSpeedSince = 0 Then
            xfers(i).lowSpeedSince = now
        ElseIf now - xfers(i).lowSpeedSince >= LOW_SPEED_TIME Then
            finishTransfer i, CURLE_OPERATION_TIMEDOUT, "Operation too slow. Less than " & LOW_SPEED_LIMIT & " bytes/sec transferred the last " & (LOW_SPEED_TIME \ 1000) & " seconds"
            lowSpeedCheck = True
        End If
    Else
        xfers(i).lowSpeedSince = 0
    End If
End Function

' curl_multi_poll for one transfer: up to ms for its socket(s)
Private Sub waitForSocket(ByVal i As Long, ByVal ms As Long)
    Dim rd As fd_set, wr As fd_set, ex As fd_set, tv As timeval, b As Long
    Select Case xfers(i).phase
    Case XP_CONNECT
        For b = 0 To 1
            If xfers(i).balState(b) = 1 Then
                wr.fd_array(wr.fd_count) = xfers(i).balSock(b)
                wr.fd_count = wr.fd_count + 1
                ex.fd_array(ex.fd_count) = xfers(i).balSock(b)
                ex.fd_count = ex.fd_count + 1
            End If
        Next
        If ms > HAPPY_EYEBALLS_TIMEOUT Then ms = HAPPY_EYEBALLS_TIMEOUT
    Case XP_SEND
        wr.fd_array(0) = xfers(i).sock
        wr.fd_count = 1
    Case XP_TLS, XP_HEADERS, XP_BODY
        If xfers(i).sendPos < xfers(i).sendLen Then
            wr.fd_array(0) = xfers(i).sock
            wr.fd_count = 1
        Else
            rd.fd_array(0) = xfers(i).sock
            rd.fd_count = 1
        End If
    Case Else
        Exit Sub
    End Select
    If rd.fd_count + wr.fd_count + ex.fd_count = 0 Then Exit Sub
    tv.tv_sec = ms \ 1000
    tv.tv_usec = (ms Mod 1000) * 1000
    ws_select 0, rd, wr, ex, tv
End Sub

' the transfer ends with a CURLcode (failures logged with their detail)
Private Sub finishTransfer(ByVal i As Long, ByVal code As Long, ByVal detail As String)
    xfers(i).result = code
    If code <> CURLE_OK Then
        debug_ 0, LOG_NET, "urlTransfer", "(VB) %s (%d): %s (%s)", curlStrError(code), code, detail, xfers(i).url
    End If
    closeConnection i, code = CURLE_OK
    Erase xfers(i).sendBuf
    Erase xfers(i).tlsIn
    Erase xfers(i).plain
    Erase xfers(i).body
    xfers(i).phase = XP_DONE
End Sub

' the connection closed (graceful: close_notify first); its attempts too
Private Sub closeConnection(ByVal i As Long, ByVal graceful As Boolean)
    Dim ctx As SecHandle, b As Long
    If xfers(i).hasCtx Then
        If graceful And xfers(i).handshakeDone And xfers(i).sock <> INVALID_SOCKET Then sendCloseNotify i
        ctx = xfers(i).ctx
        DeleteSecurityContext ctx
        xfers(i).hasCtx = False
        xfers(i).handshakeDone = False
        xfers(i).renegotiating = False
    End If
    If xfers(i).sock <> INVALID_SOCKET Then closeSocket xfers(i).sock
    For b = 0 To 1
        If xfers(i).balSock(b) <> INVALID_SOCKET Then closeSocket xfers(i).balSock(b)
        xfers(i).balState(b) = 0
    Next
    xfers(i).tlsInLen = 0
    xfers(i).tlsNeedMore = False
    xfers(i).plainLen = 0
    xfers(i).sendLen = 0
    xfers(i).sendPos = 0
    xfers(i).eof = False
    xfers(i).addrCount = 0
End Sub

Private Sub closeSocket(s As Long)
    ws_closesocket s
    s = INVALID_SOCKET
End Sub

Private Function newTransfer() As Long
    Dim i As Long, b As Long
    For i = 0 To xferCount - 1
        If xfers(i).phase = XP_FREE Then Exit For
    Next
    If i = xferCount Then
        xferCount = xferCount + 4
        ReDim Preserve xfers(0 To xferCount - 1)
    End If
    xfers(i).sock = INVALID_SOCKET
    For b = 0 To 1
        xfers(i).balSock(b) = INVALID_SOCKET
    Next
    xfers(i).contentLength = -1
    newTransfer = i
End Function

' (VB) buf grows to hold n more bytes from the address p
Private Sub bufAppendPtr(buf() As Byte, bufLen As Long, ByVal p As Long, ByVal n As Long)
    Dim cap As Long
    If n <= 0 Then Exit Sub
    cap = -1
    On Error Resume Next
    cap = UBound(buf)
    On Error GoTo 0
    If cap + 1 < bufLen + n Then
        cap = (cap + 1) * 2
        If cap < bufLen + n Then cap = bufLen + n
        If cap < 4096 Then cap = 4096
        ReDim Preserve buf(0 To cap - 1)
    End If
    CopyMemory buf(bufLen), ByVal p, n
    bufLen = bufLen + n
End Sub

' (VB) the first n bytes of buf dropped
Private Sub bufConsume(buf() As Byte, bufLen As Long, ByVal n As Long)
    If n <= 0 Then Exit Sub
    If n >= bufLen Then
        bufLen = 0
        Exit Sub
    End If
    CopyMemory buf(0), buf(n), bufLen - n
    bufLen = bufLen - n
End Sub

Private Function latin1(b() As Byte, ByVal ofs As Long, ByVal n As Long) As String
    Dim s As String, k As Long
    s = String$(n, 0)
    For k = 0 To n - 1
        Mid$(s, k + 1, 1) = ChrW$(b(ofs + k))
    Next
    latin1 = s
End Function

' an address of the transfer's list as text (inet_ntop)
Private Function addrText(ByVal i As Long, ByVal k As Long) As String
    Dim buf As String, ofs As Long
    buf = String$(64, 0)
    ofs = k * SOCKADDR_SIZE
    If xfers(i).addrFam(k) = AF_INET Then
        ws_inet_ntop AF_INET, xfers(i).addrs(ofs + 4), buf, 64
    Else
        ws_inet_ntop AF_INET6, xfers(i).addrs(ofs + 8), buf, 64
    End If
    addrText = Left$(buf, InStr(buf, vbNullChar) - 1)
End Function

Private Function familyName(ByVal fam As Long) As String
    Select Case fam
    Case AF_INET: familyName = "IPv4"
    Case AF_INET6: familyName = "IPv6"
    Case Else: familyName = "any"
    End Select
End Function

Private Function protocolName(ByVal proto As Long) As String
    Select Case proto
    Case &H800&, &H400&: protocolName = "TLS 1.2"
    Case &H2000&, &H1000&: protocolName = "TLS 1.3"
    Case &H200&, &H100&: protocolName = "TLS 1.1"
    Case &H80&, &H40&: protocolName = "TLS 1.0"
    Case Else: protocolName = "protocol 0x" & Hex$(proto)
    End Select
End Function

' CURLINFO_RETRY_AFTER (http.c): delay-seconds, or an HTTP-date (RFC 7231's IMF-fixdate) as the seconds from now
Private Function parseRetryAfter(ByVal v As String) As Double
    Dim parts() As String, months As String, m As Long, d As Date, st As SYSTEMTIME, now As Date, secs As Double
    v = Trim$(v)
    If Len(v) = 0 Then Exit Function
    If Len(v) < 16 And v Like String$(Len(v), "#") Then
        parseRetryAfter = CDbl(v)
        Exit Function
    End If
    ' "Sun, 06 Nov 1994 08:49:37 GMT"
    parts = Split(v, " ")
    If UBound(parts) <> 5 Then Exit Function
    months = "JanFebMarAprMayJunJulAugSepOctNovDec"
    m = InStr(months, parts(2))
    If m = 0 Or (m - 1) Mod 3 <> 0 Or Len(parts(2)) <> 3 Then Exit Function
    If Not (parts(1) Like "##") Or Not (parts(3) Like "####") Or Not (parts(4) Like "##:##:##") Then Exit Function
    On Error Resume Next
    d = DateSerial(CInt(parts(3)), (m - 1) \ 3 + 1, CInt(parts(1))) + TimeSerial(CInt(Left$(parts(4), 2)), CInt(Mid$(parts(4), 4, 2)), CInt(Right$(parts(4), 2)))
    If Err.Number <> 0 Then
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    GetSystemTime st
    now = DateSerial(st.wYear, st.wMonth, st.wDay) + TimeSerial(st.wHour, st.wMinute, st.wSecond)
    secs = DateDiff("s", now, d)
    If secs > 0 Then parseRetryAfter = secs
End Function

' curl_easy_strerror
Private Function curlStrError(ByVal code As Long) As String
    Select Case code
    Case CURLE_OK: curlStrError = "No error"
    Case CURLE_UNSUPPORTED_PROTOCOL: curlStrError = "Unsupported protocol"
    Case CURLE_URL_MALFORMAT: curlStrError = "URL using bad/illegal format or missing URL"
    Case CURLE_COULDNT_RESOLVE_HOST: curlStrError = "Could not resolve hostname"
    Case CURLE_COULDNT_CONNECT: curlStrError = "Could not connect to server"
    Case CURLE_WEIRD_SERVER_REPLY: curlStrError = "Weird server reply"
    Case CURLE_PARTIAL_FILE: curlStrError = "Transferred a partial file"
    Case CURLE_WRITE_ERROR: curlStrError = "Failed writing received data to disk/application"
    Case CURLE_OPERATION_TIMEDOUT: curlStrError = "Timeout was reached"
    Case CURLE_SSL_CONNECT_ERROR: curlStrError = "SSL connect error"
    Case CURLE_TOO_MANY_REDIRECTS: curlStrError = "Number of redirects hit maximum amount"
    Case CURLE_GOT_NOTHING: curlStrError = "Server returned nothing (no headers, no data)"
    Case CURLE_SEND_ERROR: curlStrError = "Failed sending data to the peer"
    Case CURLE_RECV_ERROR: curlStrError = "Failure when receiving data from the peer"
    Case CURLE_PEER_FAILED_VERIFICATION: curlStrError = "SSL peer certificate or SSH remote key was not OK"
    Case CURLE_BAD_CONTENT_ENCODING: curlStrError = "Unrecognized or bad HTTP Content or Transfer-Encoding"
    Case CURLE_FILESIZE_EXCEEDED: curlStrError = "Maximum file size exceeded"
    Case CURLE_TOO_LARGE: curlStrError = "A value or data field grew larger than allowed"
    Case Else: curlStrError = "Unknown error"
    End Select
End Function

' RtlGetVersion's build number (GetVersionEx's is capped by the manifest)
Private Function windowsBuild() As Long
    Dim v(0 To 69) As Long
    v(0) = 276                                  ' sizeof(RTL_OSVERSIONINFOW): 5 DWORDs + 128 WCHARs
    If RtlGetVersion(v(0)) = 0 Then windowsBuild = v(3)
End Function

' std::chrono::steady_clock::now(), in milliseconds
Private Function nowMs() As Double
    Dim c As Currency
    If qpcFreq = 0 Then QueryPerformanceFrequency qpcFreq
    QueryPerformanceCounter c
    nowMs = CDbl(c) * 1000# / CDbl(qpcFreq)
End Function
