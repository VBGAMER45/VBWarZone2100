VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TCPClientConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/netplay/tcp/tcp_client_connection.cpp, tcp_client_connection.h: class tcp::TCPClientConnection, the TCP part
' of an IClientConnection (which holds this object: the base). It keeps no reference to its base (the base
' holds it): the methods that need the base (connectionStatus's selfConnList_) take it.
' This file includes tcp/sock_error.h: EBADF, ECONNRESET and EINTR are the Winsock codes.
Option Explicit

Private Declare Function send_ Lib "ws2_32" Alias "send" (ByVal s As Long, buf As Any, ByVal len_ As Long, ByVal flags As Long) As Long
Private Declare Function recv_ Lib "ws2_32" Alias "recv" (ByVal s As Long, buf As Any, ByVal len_ As Long, ByVal flags As Long) As Long
Private Declare Function ioctlsocket Lib "ws2_32" (ByVal s As Long, ByVal cmd As Long, argp As Long) As Long

Private Const SOCKET_ERROR As Long = -1
Private Const MSG_NOSIGNAL As Long = 0
Private Const FIONREAD As Long = &H4004667F
Private Const EBADF As Long = 10009             ' WSAEBADF
Private Const ECONNRESET As Long = 10054        ' WSAECONNRESET
Private Const EINTR As Long = 10004             ' WSAEINTR

Private socket_ As Long                         ' Socket* (netsocket_c's handle; 0 = nullptr)
Private connStatusDescriptorSet_ As SelectDescriptorSet

' explicit TCPClientConnection(WzConnectionProvider& connProvider, WzCompressionProvider& compressionProvider,
'     PendingWritesManager& pwm, Socket* rawSocket): makes its base, the IClientConnection it returns
Friend Function construct(ByVal connProvider As WzConnectionProvider, ByVal pwm As PendingWritesManager, ByVal rawSocket As Long) As IClientConnection
    Dim base As IClientConnection
    Set base = New IClientConnection
    base.construct connProvider, pwm, Me
    socket_ = rawSocket
    Set connStatusDescriptorSet_ = connProvider.newDescriptorSet(0)  ' PollEventType::READABLE
    If Not (socket_ <> 0) Then ASSERT_FAILURE "socket_ != nullptr", "lib/netplay/tcp/tcp_client_connection.cpp:38", "TCPClientConnection", "Null socket passed to TCPClientConnection ctor"
    Set construct = base
End Function

' ~TCPClientConnection()
Friend Sub destroy_()
    If socket_ = 0 Then
        Exit Sub
    End If
    socketCloseNow socket_
    socket_ = 0
End Sub

' virtual net::result<ssize_t> sendImpl(const std::vector<uint8_t>& data) override
Friend Function sendImpl(data() As Byte, ByVal size As Long, ec As Long) As Long
    Dim retSent As Long
    ec = 0
    If Not isValid() Then
        debug_ 55, LOG_ERROR, "sendImpl", "Invalid socket (EBADF)"
        ec = make_network_error_code(EBADF)
        Exit Function
    End If
    If size = 0 Then
        retSent = send_(getRawSocketFd(), ByVal 0&, 0, MSG_NOSIGNAL)
    Else
        retSent = send_(getRawSocketFd(), data(0), size, MSG_NOSIGNAL)
    End If
    sockCalled
    If retSent <> SOCKET_ERROR Then
        sendImpl = retSent
        Exit Function
    End If
    ec = make_network_error_code(getSockErr())
End Function

' virtual net::result<ssize_t> recvImpl(char* dst, size_t maxSize) override
Friend Function recvImpl(dst() As Byte, ByVal ofs As Long, ByVal maxSize As Long, ec As Long) As Long
    Dim fd As Long, lastSockErr As Long, received As Long
    ec = 0
    fd = getRawSocketFd()
    lastSockErr = 0
    Do
        If maxSize = 0 Then
            received = recv_(fd, ByVal 0&, 0, 0)      ' (VB) dst may be one past its array's end then
        Else
            received = recv_(fd, dst(ofs), maxSize, 0)
        End If
        sockCalled
        If received = SOCKET_ERROR Then
            lastSockErr = getSockErr()
        ElseIf received = 0 Then
            ' Socket got disconnected. No reason to do anything else here.
            ec = make_network_error_code(ECONNRESET)
            Exit Function
        End If
    Loop While received = SOCKET_ERROR And lastSockErr = EINTR

    If received = SOCKET_ERROR Then
        ec = make_network_error_code(lastSockErr)
        Exit Function
    End If
    recvImpl = received
End Function

' virtual bool readReady() const override
Friend Function readReady() As Boolean
    readReady = socketReadReady(socket_)
End Function

' virtual void setReadReady(bool ready) override
Friend Sub setReadReady(ByVal ready As Boolean)
    socketSetReadReady socket_, ready
End Sub

' virtual void useNagleAlgorithm(bool enable) override
Friend Sub useNagleAlgorithm(ByVal enable As Boolean)
    socketSetTCPNoDelay socket_, Not enable
End Sub

' virtual std::string textAddress() const override
Friend Function textAddress() As String
    textAddress = getSocketTextAddress(socket_)
End Function

' virtual bool isValid() const override
Friend Function isValid() As Boolean
    If socket_ = 0 Then Exit Function
    isValid = isValidSocket(socket_)
End Function

' virtual net::result<void> connectionStatus() const override: the result is ec; base is this connection's
' IClientConnection (selfConnList_, readReady)
Friend Function connectionStatus(ByVal base As IClientConnection) As Long
    Dim checkConnRes As Long, ec As Long, readQueue As Long, ioctlRet As Long
    Dim conns As Collection
    If Not (isValid()) Then ASSERT_FAILURE "isValid()", "lib/netplay/tcp/tcp_client_connection.cpp:119", "connectionStatus", "Invalid connection object": connectionStatus = make_network_error_code(EBADF): Exit Function

    ' Check whether the socket is still connected
    Set conns = base.selfConnList()
    checkConnRes = checkConnectionsReadable(conns, connStatusDescriptorSet_, 0, ec)
    If ec <> 0 Then
        connectionStatus = ec
        Exit Function
    ElseIf checkConnRes = conns.Count And readReady() Then
        ' The next recv(2) call won't block, but we're writing. So check the read queue to see if the connection
        ' is closed. If there's no data in the queue that means the connection is closed.
        ioctlRet = ioctlsocket(getRawSocketFd(), FIONREAD, readQueue)
        sockCalled
        If ioctlRet = SOCKET_ERROR Then
            debug_ 144, LOG_NET, "connectionStatus", "socket error"
            connectionStatus = make_network_error_code(getSockErr())
            Exit Function
        ElseIf readQueue = 0 Then
            ' Disconnected
            debug_ 150, LOG_NET, "connectionStatus", "Read queue empty - failing (ECONNRESET)"
            connectionStatus = make_network_error_code(ECONNRESET)
            Exit Function
        End If
    End If
    connectionStatus = 0
End Function

' virtual void setConnectedTimeout(std::chrono::milliseconds timeout) override
Friend Sub setConnectedTimeout(ByVal timeoutMs As Long)
    ' NO-OP
End Sub

' SOCKET getRawSocketFd() const
Friend Function getRawSocketFd() As Long
    getRawSocketFd = netsocket_c.getRawSocketFd(socket_)
End Function
