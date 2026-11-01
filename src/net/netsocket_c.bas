Attribute VB_Name = "netsocket_c"
' lib/netplay/tcp/netsocket.cpp, netsocket.h and tcp/sock_error.cpp, sock_error.h (the WZ_OS_WIN branches), over
' Winsock 2 (ws2_32 declares).
' struct tcp::Socket is a slot of this module's table and a Socket* is its handle (index + 1; 0 is nullptr);
' a SocketAddress* (struct addrinfo, getaddrinfo's list) is the Long pointer getaddrinfo gives. A "%p" in a log
' line prints the handle. sock_error.h's Windows defines hold here: EAGAIN and EWOULDBLOCK are WSAEWOULDBLOCK,
' EBADF WSAEBADF, ECONNRESET WSAECONNRESET, EINPROGRESS WSAEINPROGRESS, EINTR WSAEINTR, EISCONN WSAEISCONN,
' ETIMEDOUT WSAETIMEDOUT. SOCK_CLOEXEC is not defined on Windows (setSocketInheritable is used), nor SO_NOSIGPIPE
' (socketBlockSIGPIPE does nothing); IPV6_V6ONLY is (the listen socket is one dual-stack IPv6 socket when it can).
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)
Private Declare Function SetHandleInformation Lib "kernel32" (ByVal hObject As Long, ByVal dwMask As Long, ByVal dwFlags As Long) As Long
Private Declare Function ws_WSAStartup Lib "ws2_32" Alias "WSAStartup" (ByVal wVersionRequested As Integer, lpWSAData As Any) As Long
Private Declare Function WSACleanup Lib "ws2_32" () As Long
Private Declare Function ws_socket Lib "ws2_32" Alias "socket" (ByVal af As Long, ByVal type_ As Long, ByVal protocol As Long) As Long
Private Declare Function ws_closesocket Lib "ws2_32" Alias "closesocket" (ByVal s As Long) As Long
Private Declare Function ws_connect Lib "ws2_32" Alias "connect" (ByVal s As Long, ByVal name_ As Long, ByVal namelen As Long) As Long
Private Declare Function ws_bind Lib "ws2_32" Alias "bind" (ByVal s As Long, name_ As Any, ByVal namelen As Long) As Long
Private Declare Function ws_listen Lib "ws2_32" Alias "listen" (ByVal s As Long, ByVal backlog As Long) As Long
Private Declare Function ws_accept Lib "ws2_32" Alias "accept" (ByVal s As Long, addr As Any, addrlen As Long) As Long
Private Declare Function ws_ioctlsocket Lib "ws2_32" Alias "ioctlsocket" (ByVal s As Long, ByVal cmd As Long, argp As Long) As Long
Private Declare Function ws_setsockopt Lib "ws2_32" Alias "setsockopt" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
Private Declare Function ws_getsockopt Lib "ws2_32" Alias "getsockopt" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, optval As Any, optlen As Long) As Long
Private Declare Function ws_select Lib "ws2_32" Alias "select" (ByVal nfds As Long, ByVal readfds As Long, ByVal writefds As Long, ByVal exceptfds As Long, timeout As timeval) As Long
Private Declare Function getaddrinfo Lib "ws2_32" (ByVal pNodeName As String, ByVal pServiceName As String, pHints As ADDRINFOA, ppResult As Long) As Long
Private Declare Function inet_pton Lib "ws2_32" (ByVal Family As Long, ByVal pszAddrString As String, pAddrBuf As Any) As Long
Private Declare Function inet_ntop Lib "ws2_32" (ByVal Family As Long, pAddr As Any, ByVal pStringBuf As String, ByVal StringBufSize As Long) As Long

Private Const INVALID_SOCKET As Long = -1
Private Const SOCKET_ERROR As Long = -1
Private Const AF_INET As Long = 2
Private Const AF_INET6 As Long = 23
Private Const AF_UNSPEC As Long = 0
Private Const SOCK_STREAM As Long = 1
Private Const IPPROTO_TCP As Long = 6
Private Const IPPROTO_IPV6 As Long = 41
Private Const IPV6_V6ONLY_ As Long = 27        ' IPV6_V6ONLY (VB would take the locals named ipv6_v6only for it)
Private Const TCP_NODELAY As Long = 1
Private Const SOL_SOCKET As Long = &HFFFF&
Private Const SO_REUSEADDR As Long = 4
Private Const FIONBIO As Long = &H8004667E
Private Const HANDLE_FLAG_INHERIT As Long = 1
Private Const AI_V4MAPPED As Long = &H800      ' ws2def.h's (netsocket.h's fallbacks are for older headers)
Private Const AI_ADDRCONFIG As Long = &H400
' tcp/sock_error.h on Windows
Private Const EAGAIN As Long = 10035            ' WSAEWOULDBLOCK
Private Const EWOULDBLOCK As Long = 10035       ' WSAEWOULDBLOCK
Private Const EINPROGRESS As Long = 10036       ' WSAEINPROGRESS
Private Const EINTR As Long = 10004             ' WSAEINTR
Private Const ETIMEDOUT As Long = 10060         ' WSAETIMEDOUT

' enum { SOCK_CONNECTION, SOCK_IPV4_LISTEN = SOCK_CONNECTION, SOCK_IPV6_LISTEN, SOCK_COUNT }
Private Const SOCK_CONNECTION As Long = 0
Private Const SOCK_IPV4_LISTEN As Long = 0
Private Const SOCK_IPV6_LISTEN As Long = 1
Private Const SOCK_COUNT As Long = 2

Public Type timeval
    tv_sec As Long
    tv_usec As Long
End Type

' winsock2.h's fd_set (FD_SETSIZE 64)
Public Type fd_set
    fd_count As Long
    fd_array(0 To 63) As Long
End Type

' ws2def.h's ADDRINFOA (SocketAddress)
Public Type ADDRINFOA
    ai_flags As Long
    ai_family As Long
    ai_socktype As Long
    ai_protocol As Long
    ai_addrlen As Long
    ai_canonname As Long
    ai_addr As Long
    ai_next As Long
End Type

' struct Socket (a slot: fd[SOCK_COUNT], ready, char textAddress[40])
Private sockFd() As Long                ' sockFd(slot * SOCK_COUNT + i)
Private sockReady() As Boolean
Private sockTextAddress() As String
Private sockUsed() As Boolean
Private sockCount As Long

Private wsaData(0 To 511) As Byte       ' static WSADATA stuff;
Private lastSockErr_ As Long            ' (VB) see sockCalled

' (VB) new Socket: a free slot, fd[] = {INVALID_SOCKET, INVALID_SOCKET}, ready = false, textAddress = ""
Private Function newSocket() As Long
    Dim i As Long
    For i = 0 To sockCount - 1
        If Not sockUsed(i) Then Exit For
    Next
    If i = sockCount Then
        sockCount = sockCount + 16
        ReDim Preserve sockFd(0 To sockCount * SOCK_COUNT - 1)
        ReDim Preserve sockReady(0 To sockCount - 1)
        ReDim Preserve sockTextAddress(0 To sockCount - 1)
        ReDim Preserve sockUsed(0 To sockCount - 1)
    End If
    sockUsed(i) = True
    sockFd(i * SOCK_COUNT + SOCK_CONNECTION) = INVALID_SOCKET
    sockFd(i * SOCK_COUNT + SOCK_IPV6_LISTEN) = INVALID_SOCKET
    sockReady(i) = False
    sockTextAddress(i) = ""
    newSocket = i + 1
End Function

' (VB) delete sock
Private Sub deleteSocket(ByVal sock As Long)
    sockUsed(sock - 1) = False
End Sub

' (VB) sock->fd[i] for a handle
Private Function fdOf(ByVal sock As Long, ByVal i As Long) As Long
    fdOf = sockFd((sock - 1) * SOCK_COUNT + i)
End Function

Private Sub setFd(ByVal sock As Long, ByVal i As Long, ByVal fd As Long)
    sockFd((sock - 1) * SOCK_COUNT + i) = fd
End Sub

' (VB) the thread's last error as Winsock leaves it. VB sets it to 0 before each Declare'd call (Err.LastDllError),
' so WSAGetLastError through a Declare always gives 0: each Winsock call goes through a wrapper of its name below,
' which keeps a code the call set (sockCalled); a call that succeeds leaves the older one, as WSAGetLastError does
' in C (4.7.0 logs a connect's WSAEWOULDBLOCK after the select() that found the connection failed).
Public Sub sockCalled()
    Dim e As Long
    e = Err.LastDllError
    If e <> 0 Then lastSockErr_ = e
End Sub

' tcp/sock_error.cpp: int getSockErr() (WSAGetLastError())
Public Function getSockErr() As Long
    getSockErr = lastSockErr_
End Function

Private Function WSAStartup(ByVal wVersionRequested As Integer, lpWSAData As Byte) As Long
    WSAStartup = ws_WSAStartup(wVersionRequested, lpWSAData)
    sockCalled
End Function

Private Function socket_(ByVal af As Long, ByVal type_ As Long, ByVal protocol As Long) As Long
    socket_ = ws_socket(af, type_, protocol)
    sockCalled
End Function

Private Function closesocket(ByVal s As Long) As Long
    closesocket = ws_closesocket(s)
    sockCalled
End Function

Private Function connect_(ByVal s As Long, ByVal name_ As Long, ByVal namelen As Long) As Long
    connect_ = ws_connect(s, name_, namelen)
    sockCalled
End Function

Private Function bind_(ByVal s As Long, name_ As Byte, ByVal namelen As Long) As Long
    bind_ = ws_bind(s, name_, namelen)
    sockCalled
End Function

Private Function listen_(ByVal s As Long, ByVal backlog As Long) As Long
    listen_ = ws_listen(s, backlog)
    sockCalled
End Function

Private Function accept_(ByVal s As Long, addr As Byte, addrlen As Long) As Long
    accept_ = ws_accept(s, addr, addrlen)
    sockCalled
End Function

Private Function ioctlsocket(ByVal s As Long, ByVal cmd As Long, argp As Long) As Long
    ioctlsocket = ws_ioctlsocket(s, cmd, argp)
    sockCalled
End Function

Private Function setsockopt(ByVal s As Long, ByVal level As Long, ByVal optname As Long, optval As Long, ByVal optlen As Long) As Long
    setsockopt = ws_setsockopt(s, level, optname, optval, optlen)
    sockCalled
End Function

Private Function getsockopt(ByVal s As Long, ByVal level As Long, ByVal optname As Long, optval As Long, optlen As Long) As Long
    getsockopt = ws_getsockopt(s, level, optname, optval, optlen)
    sockCalled
End Function

Private Function select_(ByVal nfds As Long, ByVal readfds As Long, ByVal writefds As Long, ByVal exceptfds As Long, timeout As timeval) As Long
    select_ = ws_select(nfds, readfds, writefds, exceptfds, timeout)
    sockCalled
End Function

' SOCKET getRawSocketFd(const Socket& sock)
Public Function getRawSocketFd(ByVal sock As Long) As Long
    getRawSocketFd = fdOf(sock, SOCK_CONNECTION)
End Function

' bool isValidSocket(const Socket& sock)
Public Function isValidSocket(ByVal sock As Long) As Boolean
    isValidSocket = fdOf(sock, SOCK_CONNECTION) <> INVALID_SOCKET
End Function

' bool socketReadReady(const Socket& sock)
Public Function socketReadReady(ByVal sock As Long) As Boolean
    socketReadReady = sockReady(sock - 1)
End Function

' void socketSetReadReady(Socket& sock, bool ready)
Public Sub socketSetReadReady(ByVal sock As Long, ByVal ready As Boolean)
    sockReady(sock - 1) = ready
End Sub

' (VB) ntohs of the two bytes at p
Private Function ntohsAt(ByVal p As Long) As Long
    Dim b(0 To 1) As Byte
    CopyMemory b(0), ByVal p, 2
    ntohsAt = CLng(b(0)) * 256 + b(1)
End Function

' static int addressToText(const struct sockaddr *addr, char *buf, size_t size): the text into buf (char[size]:
' snprintf keeps size - 1 characters)
Private Function addressToText(ByVal addr As Long, buf As String, ByVal size As Long) As Long
    Dim family As Integer, b(0 To 15) As Byte, i As Long, s As String
    CopyMemory family, ByVal addr, 2
    Select Case family
    Case AF_INET
        ' handleIpv4(sin_addr.s_addr): ntohl, then "%u.%u.%u.%u" from the high byte
        CopyMemory b(0), ByVal addr + 4, 4
        s = b(0) & "." & b(1) & "." & b(2) & "." & b(3)
    Case AF_INET6
        CopyMemory b(0), ByVal addr + 8, 16          ' sin6_addr
        ' IN6_IS_ADDR_V4MAPPED: ::ffff:a.b.c.d
        If b(0) = 0 And b(1) = 0 And b(2) = 0 And b(3) = 0 And b(4) = 0 And b(5) = 0 And b(6) = 0 And b(7) = 0 _
                And b(8) = 0 And b(9) = 0 And b(10) = &HFF And b(11) = &HFF Then
            s = b(12) & "." & b(13) & "." & b(14) & "." & b(15)
        Else
            ' "%hx:%hx:%hx:%hx:%hx:%hx:%hx:%hx" of ntohs(address[i])
            For i = 0 To 7
                If i > 0 Then s = s & ":"
                s = s & LCase$(Hex$(CLng(b(i * 2)) * 256 + b(i * 2 + 1)))
            Next
        End If
    Case Else
        ASSERT_FAILURE "!""Unknown address family""", "lib/netplay/tcp/netsocket.cpp:162", "addressToText", "Got non IPv4 or IPv6 address!"
        addressToText = -1
        Exit Function
    End Select
    addressToText = Len(s)
    If Len(s) > size - 1 Then s = Left$(s, size - 1)
    buf = s
End Function

' static const char *strSockError(int error)
Private Function strSockError(ByVal error_ As Long) As String
    strSockError = net_errorMessage(make_network_error_code(error_))   ' the same table as GenericSystemErrorCategory's
End Function

' bool socketSetTCPNoDelay(Socket& sock, bool nodelay)
Public Function socketSetTCPNoDelay(ByVal sock As Long, ByVal nodelay As Boolean) As Boolean
    Dim value As Long, result As Long
    If nodelay Then value = 1 Else value = 0
    result = setsockopt(fdOf(sock, SOCK_CONNECTION), IPPROTO_TCP, TCP_NODELAY, value, 4)
    debug_ 234, LOG_NET, "socketSetTCPNoDelay", "Setting TCP_NODELAY on socket: %s", IIf(result <> SOCKET_ERROR, "success", "failure")
    socketSetTCPNoDelay = result <> SOCKET_ERROR
End Function

' static bool setSocketInheritable(SOCKET fd, bool inheritable) (no SOCK_CLOEXEC on Windows)
Private Function setSocketInheritable(ByVal fd As Long, ByVal inheritable As Boolean) As Boolean
    Dim dwFlags As Long, dwErr As Long
    If inheritable Then dwFlags = HANDLE_FLAG_INHERIT Else dwFlags = 0
    If SetHandleInformation(fd, HANDLE_FLAG_INHERIT, dwFlags) = 0 Then
        dwErr = Err.LastDllError
        debug_ 273, LOG_NET, "setSocketInheritable", "Failed to set socket %sinheritable: %s", IIf(inheritable, "", "non-"), CStr(dwErr)
        setSocketInheritable = False
        Exit Function
    End If
    debug_ 278, LOG_NET, "setSocketInheritable", "Socket is set to %sinheritable.", IIf(inheritable, "", "non-")
    setSocketInheritable = True
End Function

' static bool setSocketBlocking(const SOCKET fd, bool blocking)
Private Function setSocketBlocking(ByVal fd As Long, ByVal blocking As Boolean) As Boolean
    Dim nonblocking As Long
    If blocking Then nonblocking = 0 Else nonblocking = 1
    If ioctlsocket(fd, FIONBIO, nonblocking) = SOCKET_ERROR Then
        debug_ 309, LOG_NET, "setSocketBlocking", "Failed to set socket %sblocking: %s", IIf(blocking, "", "non-"), strSockError(getSockErr())
        setSocketBlocking = False
        Exit Function
    End If
    debug_ 313, LOG_NET, "setSocketBlocking", "Socket is set to %sblocking.", IIf(blocking, "", "non-")
    setSocketBlocking = True
End Function

' static void socketBlockSIGPIPE(const SOCKET fd, bool block_sigpipe): no SO_NOSIGPIPE on Windows
Private Sub socketBlockSIGPIPE(ByVal fd As Long, ByVal block_sigpipe As Boolean)
End Sub

' void socketCloseNow(Socket *sock)
Public Sub socketCloseNow(ByVal sock As Long)
    Dim i As Long, err_ As Long
    For i = 0 To SOCK_COUNT - 1
        If fdOf(sock, i) <> INVALID_SOCKET Then
            err_ = closesocket(fdOf(sock, i))
            If err_ <> 0 Then
                debug_ 348, LOG_ERROR, "socketCloseNow", "Failed to close socket %p: %s", sock, strSockError(getSockErr())
            End If
            ' Make sure that dangling pointers to this structure don't think they've got their hands on a valid socket.
            setFd sock, i, INVALID_SOCKET
        End If
    Next
    deleteSocket sock
End Sub

' Socket *socketAccept(Socket *sock)
Public Function socketAccept(ByVal sock As Long) As Long
    Dim i As Long, j As Long, addr(0 To 127) As Byte, addr_len As Long, conn As Long, newConn As Long
    ' Search for a socket that has a pending connection on it and accept the first one.
    For i = 0 To SOCK_COUNT - 1
        If fdOf(sock, i) <> INVALID_SOCKET Then
            addr_len = 128                       ' sizeof(struct sockaddr_storage)
            newConn = accept_(fdOf(sock, i), addr(0), addr_len)
            If newConn = INVALID_SOCKET Then
                ' Ignore the case where no connection is pending
                If getSockErr() <> EAGAIN And getSockErr() <> EWOULDBLOCK Then
                    debug_ 388, LOG_ERROR, "socketAccept", "accept failed for socket %p: %s", sock, strSockError(getSockErr())
                End If
                GoTo continue_
            End If

            conn = newSocket()

            If Not setSocketInheritable(newConn, False) Then
                debug_ 405, LOG_NET, "socketAccept", "Couldn't set socket (%p) inheritable status (false). Ignoring...", conn
                ' ignore and continue
            End If

            debug_ 410, LOG_NET, "socketAccept", "setting socket (%p) blocking status (false).", conn
            If Not setSocketBlocking(newConn, False) Then
                debug_ 413, LOG_NET, "socketAccept", "Couldn't set socket (%p) blocking status (false).  Closing.", conn
                socketCloseNow conn
                socketAccept = 0
                Exit Function
            End If

            socketBlockSIGPIPE newConn, True

            ' Mark all unused socket handles as invalid
            For j = 0 To SOCK_COUNT - 1
                setFd conn, j, INVALID_SOCKET
            Next

            setFd conn, SOCK_CONNECTION, newConn

            sockReady(sock - 1) = False

            addressToText VarPtr(addr(0)), sockTextAddress(conn - 1), 40
            debug_ 431, LOG_NET, "socketAccept", "Incoming connection from [%s]:/*%%d*/ (FIXME: gives strict-aliasing error)", sockTextAddress(conn - 1)
            debug_ 432, LOG_NET, "socketAccept", "Using socket %p", conn
            socketAccept = conn
            Exit Function
        End If
continue_:
    Next
    socketAccept = 0
End Function

' (VB) a SocketAddress's addrinfo
Public Sub readAddrinfo(ByVal addr As Long, ai As ADDRINFOA)
    CopyMemory ai, ByVal addr, LenB(ai)
End Sub

' (VB) FD_ISSET(fd, set): __WSAFDIsSet
Private Function fdIsSet(ByVal fd As Long, set_ As fd_set) As Boolean
    Dim i As Long
    For i = 0 To set_.fd_count - 1
        If set_.fd_array(i) = fd Then
            fdIsSet = True
            Exit Function
        End If
    Next
End Function

' net::result<Socket*> socketOpen(const SocketAddress *addr, unsigned timeout)
Public Function socketOpen(ByVal addr As Long, ByVal timeout As Long, ec As Long) As Long
    Dim i As Long, ret As Long, ai As ADDRINFOA, conn As Long, sockErr As Long, socket_type As Long
    Dim conReady As fd_set, conFailed As fd_set, tv As timeval, fd As Long
    ec = 0
    conn = newSocket()

    If Not (addr <> 0) Then ASSERT_FAILURE "addr != nullptr", "lib/netplay/tcp/netsocket.cpp:447", "socketOpen", "NULL Socket provided"
    readAddrinfo addr, ai

    addressToText ai.ai_addr, sockTextAddress(conn - 1), 40
    debug_ 454, LOG_NET, "socketOpen", "Connecting to [%s]:%d", sockTextAddress(conn - 1), ntohsAt(ai.ai_addr + 2)

    ' Mark all unused socket handles as invalid
    For i = 0 To SOCK_COUNT - 1
        setFd conn, i, INVALID_SOCKET
    Next

    socket_type = ai.ai_socktype
    setFd conn, SOCK_CONNECTION, socket_(ai.ai_family, socket_type, ai.ai_protocol)
    fd = fdOf(conn, SOCK_CONNECTION)

    If fd = INVALID_SOCKET Then
        sockErr = getSockErr()
        debug_ 474, LOG_ERROR, "socketOpen", "Failed to create a socket (%p): %s", conn, strSockError(sockErr)
        socketCloseNow conn
        ec = make_network_error_code(sockErr)
        Exit Function
    End If

    If Not setSocketInheritable(fd, False) Then
        debug_ 482, LOG_NET, "socketOpen", "Couldn't set socket (%p) inheritable status (false). Ignoring...", conn
        ' ignore and continue
    End If

    debug_ 487, LOG_NET, "socketOpen", "setting socket (%p) blocking status (false).", conn
    If Not setSocketBlocking(fd, False) Then
        sockErr = getSockErr()
        debug_ 491, LOG_NET, "socketOpen", "Couldn't set socket (%p) blocking status (false).  Closing.", conn
        socketCloseNow conn
        ec = make_network_error_code(sockErr)
        Exit Function
    End If

    socketBlockSIGPIPE fd, True

    ret = connect_(fd, ai.ai_addr, ai.ai_addrlen)
    If ret = SOCKET_ERROR Then
        If (getSockErr() <> EINPROGRESS And getSockErr() <> EAGAIN And getSockErr() <> EWOULDBLOCK) Or timeout = 0 Then
            sockErr = getSockErr()
            debug_ 515, LOG_NET, "socketOpen", "Failed to start connecting: %s, using socket %p", strSockError(sockErr), conn
            socketCloseNow conn
            ec = make_network_error_code(sockErr)
            Exit Function
        End If

        Do
            tv.tv_sec = timeout \ 1000                ' unsigned / 1000 (timeouts stay below 2^31)
            tv.tv_usec = (timeout Mod 1000) * 1000

            conReady.fd_count = 0                     ' FD_ZERO, FD_SET
            conReady.fd_array(0) = fd
            conReady.fd_count = 1
            conFailed.fd_count = 0
            conFailed.fd_array(0) = fd
            conFailed.fd_count = 1

            ret = select_(fd + 1, 0, VarPtr(conReady), VarPtr(conFailed), tv)
        Loop While ret = SOCKET_ERROR And getSockErr() = EINTR

        If ret = SOCKET_ERROR Then
            sockErr = getSockErr()
            debug_ 542, LOG_NET, "socketOpen", "Failed to wait for connection: %s, socket %p.  Closing.", strSockError(sockErr), conn
            socketCloseNow conn
            ec = make_network_error_code(sockErr)
            Exit Function
        End If

        If ret = 0 Then
            sockErr = ETIMEDOUT
            debug_ 550, LOG_NET, "socketOpen", "Timed out while waiting for connection to be established: %s, using socket %p.  Closing.", strSockError(sockErr), conn
            socketCloseNow conn
            ec = make_network_error_code(sockErr)
            Exit Function
        End If

        If Not (fdIsSet(fd, conReady) Or fdIsSet(fd, conFailed)) Then
            ASSERT_FAILURE "FD_ISSET(conn->fd[SOCK_CONNECTION], &conReady) || FD_ISSET(conn->fd[SOCK_CONNECTION], &conFailed)", "lib/netplay/tcp/netsocket.cpp:556", "socketOpen", """sock"" is the only file descriptor in set, it should be the one that is set."
        End If

        If fdIsSet(fd, conFailed) Then
            sockErr = getSockErr()
            debug_ 569, LOG_NET, "socketOpen", "Failed to connect: %s, with socket %p.  Closing.", strSockError(sockErr), conn
            socketCloseNow conn
            ec = make_network_error_code(sockErr)
            Exit Function
        End If
    End If

    socketOpen = conn
End Function

' net::result<Socket*> socketListen(unsigned int port)
Public Function socketListen(ByVal port As Long, ec As Long) As Long
    ' Enable the V4 to V6 mapping, but only when available, because it isn't available on all platforms.
    Const ipv6_v6only As Long = 0
    Const so_reuseaddr As Long = 1
    Dim addr4(0 To 15) As Byte, addr6(0 To 27) As Byte, i As Long, conn As Long, errorCode As Long, v As Long

    ec = 0
    conn = newSocket()

    ' Mark all unused socket handles as invalid
    For i = 0 To SOCK_COUNT - 1
        setFd conn, i, INVALID_SOCKET
    Next

    sockTextAddress(conn - 1) = "LISTENING SOCKET"

    ' Listen on all local IPv4 and IPv6 addresses for the given port
    addr4(0) = AF_INET And &HFF                   ' sin_family
    addr4(2) = (port \ 256) And &HFF              ' sin_port = htons(port)
    addr4(3) = port And &HFF                      ' sin_addr.s_addr = INADDR_ANY (zero)

    addr6(0) = AF_INET6 And &HFF                  ' sin6_family
    addr6(2) = (port \ 256) And &HFF              ' sin6_port = htons(port)
    addr6(3) = port And &HFF                      ' sin6_flowinfo, sin6_addr = in6addr_any, sin6_scope_id: zero

    setFd conn, SOCK_IPV4_LISTEN, socket_(AF_INET, SOCK_STREAM, 0)
    setFd conn, SOCK_IPV6_LISTEN, socket_(AF_INET6, SOCK_STREAM, 0)

    If fdOf(conn, SOCK_IPV4_LISTEN) = INVALID_SOCKET And fdOf(conn, SOCK_IPV6_LISTEN) = INVALID_SOCKET Then
        errorCode = getSockErr()
        debug_ 624, LOG_ERROR, "socketListen", "Failed to create an IPv4 and IPv6 (only supported address families) socket (%p): %s.  Closing.", conn, strSockError(errorCode)
        socketCloseNow conn
        ec = make_network_error_code(errorCode)
        Exit Function
    End If

    If fdOf(conn, SOCK_IPV4_LISTEN) <> INVALID_SOCKET Then
        debug_ 631, LOG_NET, "socketListen", "Successfully created an IPv4 socket (%p)", conn
    End If

    If fdOf(conn, SOCK_IPV6_LISTEN) <> INVALID_SOCKET Then
        debug_ 636, LOG_NET, "socketListen", "Successfully created an IPv6 socket (%p)", conn
    End If

    If fdOf(conn, SOCK_IPV6_LISTEN) <> INVALID_SOCKET Then
        v = ipv6_v6only
        If setsockopt(fdOf(conn, SOCK_IPV6_LISTEN), IPPROTO_IPV6, IPV6_V6ONLY_, v, 4) = SOCKET_ERROR Then
            debug_ 644, LOG_INFO, "socketListen", "Failed to set IPv6 socket to perform IPv4 to IPv6 mapping. Falling back to using two sockets. Error: %s", strSockError(getSockErr())
        Else
            debug_ 648, LOG_NET, "socketListen", "Successfully enabled IPv4 to IPv6 mapping. Cleaning up IPv4 socket."
            closesocket fdOf(conn, SOCK_IPV4_LISTEN)
            setFd conn, SOCK_IPV4_LISTEN, INVALID_SOCKET
        End If
    End If

    If fdOf(conn, SOCK_IPV4_LISTEN) <> INVALID_SOCKET Then
        If Not setSocketInheritable(fdOf(conn, SOCK_IPV4_LISTEN), False) Then
            debug_ 664, LOG_NET, "socketListen", "Couldn't set socket (%p) inheritable status (false). Ignoring...", conn
            ' ignore and continue
        End If

        v = so_reuseaddr
        If setsockopt(fdOf(conn, SOCK_IPV4_LISTEN), SOL_SOCKET, SO_REUSEADDR, v, 4) = SOCKET_ERROR Then
            debug_ 671, LOG_WARNING, "socketListen", "Failed to set SO_REUSEADDR on IPv4 socket. Error: %s", strSockError(getSockErr())
        End If

        debug_ 674, LOG_NET, "socketListen", "setting socket (%p) blocking status (false, IPv4).", conn
        If listenFailed(fdOf(conn, SOCK_IPV4_LISTEN), addr4, 16) Then
            debug_ 679, LOG_ERROR, "socketListen", "Failed to set up IPv4 socket for listening on port %u: %s", port, strSockError(getSockErr())
            closesocket fdOf(conn, SOCK_IPV4_LISTEN)
            setFd conn, SOCK_IPV4_LISTEN, INVALID_SOCKET
        End If
    End If

    If fdOf(conn, SOCK_IPV6_LISTEN) <> INVALID_SOCKET Then
        If Not setSocketInheritable(fdOf(conn, SOCK_IPV6_LISTEN), False) Then
            debug_ 694, LOG_NET, "socketListen", "Couldn't set socket (%p) inheritable status (false). Ignoring...", conn
            ' ignore and continue
        End If

        v = so_reuseaddr
        If setsockopt(fdOf(conn, SOCK_IPV6_LISTEN), SOL_SOCKET, SO_REUSEADDR, v, 4) = SOCKET_ERROR Then
            debug_ 701, LOG_INFO, "socketListen", "Failed to set SO_REUSEADDR on IPv6 socket. Error: %s", strSockError(getSockErr())
        End If

        debug_ 704, LOG_NET, "socketListen", "setting socket (%p) blocking status (false, IPv6).", conn
        If listenFailed(fdOf(conn, SOCK_IPV6_LISTEN), addr6, 28) Then
            debug_ 709, LOG_ERROR, "socketListen", "Failed to set up IPv6 socket for listening on port %u: %s", port, strSockError(getSockErr())
            closesocket fdOf(conn, SOCK_IPV6_LISTEN)
            setFd conn, SOCK_IPV6_LISTEN, INVALID_SOCKET
        End If
    End If

    ' Check whether we still have at least a single (operating) socket.
    If fdOf(conn, SOCK_IPV4_LISTEN) = INVALID_SOCKET And fdOf(conn, SOCK_IPV6_LISTEN) = INVALID_SOCKET Then
        errorCode = getSockErr()
        debug_ 724, LOG_NET, "socketListen", "No IPv4 or IPv6 sockets created."
        socketCloseNow conn
        ec = make_network_error_code(errorCode)
        Exit Function
    End If

    socketListen = conn
End Function

' (VB) bind(...) == SOCKET_ERROR || listen(..., 5) == SOCKET_ERROR || !setSocketBlocking(..., false), in that order
Private Function listenFailed(ByVal fd As Long, addr() As Byte, ByVal addrLen As Long) As Boolean
    If bind_(fd, addr(0), addrLen) = SOCKET_ERROR Then
        listenFailed = True
    ElseIf listen_(fd, 5) = SOCKET_ERROR Then
        listenFailed = True
    ElseIf Not setSocketBlocking(fd, False) Then
        listenFailed = True
    End If
End Function

' net::result<Socket*> socketOpenAny(const SocketAddress *addr, unsigned timeout)
Public Function socketOpenAny(ByVal addr As Long, ByVal timeout As Long, ec As Long) As Long
    Dim res As Long, ai As ADDRINFOA
    ec = 0
    res = 0
    Do While addr <> 0
        res = socketOpen(addr, timeout, ec)
        If ec = 0 Then
            socketOpenAny = res
            Exit Function
        End If
        readAddrinfo addr, ai
        addr = ai.ai_next
    Loop
    socketOpenAny = res
End Function

' bool socketHasIPv4(const Socket& sock)
Public Function socketHasIPv4(ByVal sock As Long) As Boolean
    Dim ipv6_v6only As Long, len_ As Long
    If fdOf(sock, SOCK_IPV4_LISTEN) <> INVALID_SOCKET Then
        socketHasIPv4 = True
    Else
        If fdOf(sock, SOCK_IPV6_LISTEN) <> INVALID_SOCKET Then
            ipv6_v6only = 1
            len_ = 4
            If getsockopt(fdOf(sock, SOCK_IPV6_LISTEN), IPPROTO_IPV6, IPV6_V6ONLY_, ipv6_v6only, len_) = 0 Then
                socketHasIPv4 = (ipv6_v6only = 0)
                Exit Function
            End If
        End If
        socketHasIPv4 = False
    End If
End Function

' bool socketHasIPv6(const Socket& sock)
Public Function socketHasIPv6(ByVal sock As Long) As Boolean
    socketHasIPv6 = fdOf(sock, SOCK_IPV6_LISTEN) <> INVALID_SOCKET
End Function

' char const *getSocketTextAddress(const Socket& sock)
Public Function getSocketTextAddress(ByVal sock As Long) As String
    getSocketTextAddress = sockTextAddress(sock - 1)
End Function

' std::vector<unsigned char> ipv4_AddressString_To_NetBinary(const std::string& ipv4Address): the bytes into
' binaryForm, their count the result (0: inet_pton failed, the vector cleared)
Public Function ipv4_AddressString_To_NetBinary(ByVal ipv4Address As String, binaryForm() As Byte) As Long
    ReDim binaryForm(0 To 3)
    If inet_pton(AF_INET, ipv4Address, binaryForm(0)) <= 0 Then
        Erase binaryForm                         ' inet_pton failed
        ipv4_AddressString_To_NetBinary = 0
        Exit Function
    End If
    ipv4_AddressString_To_NetBinary = 4
End Function

' std::string ipv4_NetBinary_To_AddressString(const std::vector<unsigned char>& ip4NetBinaryForm)
Public Function ipv4_NetBinary_To_AddressString(ip4NetBinaryForm() As Byte, ByVal size As Long) As String
    Dim buf As String
    If size <> 4 Then Exit Function
    buf = String$(16, vbNullChar)                ' INET_ADDRSTRLEN
    If inet_ntop(AF_INET, ip4NetBinaryForm(0), buf, 16) = 0 Then Exit Function
    ipv4_NetBinary_To_AddressString = Left$(buf, InStr(buf, vbNullChar) - 1)
End Function

' std::vector<unsigned char> ipv6_AddressString_To_NetBinary(const std::string& ipv6Address)
Public Function ipv6_AddressString_To_NetBinary(ByVal ipv6Address As String, binaryForm() As Byte) As Long
    ReDim binaryForm(0 To 15)
    If inet_pton(AF_INET6, ipv6Address, binaryForm(0)) <= 0 Then
        Erase binaryForm                         ' inet_pton failed
        ipv6_AddressString_To_NetBinary = 0
        Exit Function
    End If
    ipv6_AddressString_To_NetBinary = 16
End Function

' std::string ipv6_NetBinary_To_AddressString(const std::vector<unsigned char>& ip6NetBinaryForm)
Public Function ipv6_NetBinary_To_AddressString(ip6NetBinaryForm() As Byte, ByVal size As Long) As String
    Dim buf As String
    If size <> 16 Then Exit Function
    buf = String$(46, vbNullChar)                ' INET6_ADDRSTRLEN
    If inet_ntop(AF_INET6, ip6NetBinaryForm(0), buf, 46) = 0 Then Exit Function
    ipv6_NetBinary_To_AddressString = Left$(buf, InStr(buf, vbNullChar) - 1)
End Function

' net::result<SocketAddress*> resolveHost(const char *host, unsigned int port)
Public Function resolveHost(ByVal host As String, ByVal port As Long, ec As Long) As Long
    Dim results As Long, service As String, hint As ADDRINFOA, flags As Long, error_ As Long, errMsg As String

    ec = 0
    hint.ai_family = AF_UNSPEC
    hint.ai_socktype = SOCK_STREAM
    hint.ai_protocol = 0
    flags = flags Or AI_V4MAPPED
    flags = flags Or AI_ADDRCONFIG
    hint.ai_flags = flags
    hint.ai_addrlen = 0
    hint.ai_addr = 0
    hint.ai_canonname = 0
    hint.ai_next = 0

    service = va("%u", port)

    error_ = getaddrinfo(host, service, hint, results)
    If error_ <> 0 Then
        ec = make_getaddrinfo_error_code(error_)
        errMsg = net_errorMessage(ec)
        debug_ 872, LOG_NET, "resolveHost", "getaddrinfo failed for %s:%s: %s", host, service, errMsg
        Exit Function
    End If

    resolveHost = results
End Function

' void SOCKETinit()
Public Sub SOCKETinit()
    ' `WSAStartup`/ `WSACleanup` can be called many times, just need to make sure that for each `WSAStartup`
    ' there's a matching `WSACleanup` call.
    If WSAStartup(&H202, wsaData(0)) <> 0 Then     ' (2 << 8) + 2: Winsock 2.2
        debug_ 893, LOG_ERROR, "SOCKETinit", "Failed to initialize Winsock: %s", strSockError(getSockErr())
        Exit Sub
    End If
End Sub

' void SOCKETshutdown()
Public Sub SOCKETshutdown()
    WSACleanup
End Sub
