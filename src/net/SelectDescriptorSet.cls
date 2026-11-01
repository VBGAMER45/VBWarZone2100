VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SelectDescriptorSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/netplay/descriptor_set.h (class IDescriptorSet) and tcp/select_descriptor_set.h (template <PollEventType>
' class tcp::SelectDescriptorSet: IDescriptorSet's one implementation on Windows, as TCPConnectionProvider::
' newDescriptorSet chooses). The template argument is construct's eventType (PollEventType: READABLE 0,
' WRITABLE 1). isSet's ::tl::expected<bool, ErroredState> is the Boolean result with errState set to the
' ErroredState + 1 (InvalidConn 1, Error 2, HangUp 3) when it holds an error, 0 otherwise.
Option Explicit

Private Declare Function select_ Lib "ws2_32" Alias "select" (ByVal nfds As Long, ByVal readfds As Long, ByVal writefds As Long, ByVal exceptfds As Long, timeout As timeval) As Long

Private Const SOCKET_ERROR As Long = -1
Private Const EINTR As Long = 10004             ' tcp/sock_error.h: WSAEINTR
Private Const EAGAIN As Long = 10035            ' tcp/sock_error.h: WSAEWOULDBLOCK
Private Const FD_SETSIZE As Long = 64

Private eventType_ As Long
Private fds_ As fd_set
Private maxfd_ As Long                          ' SOCKET (unsigned)

' explicit SelectDescriptorSet()
Friend Sub construct(ByVal eventType As Long)
    eventType_ = eventType
    fds_.fd_count = 0                           ' FD_ZERO(&fds_)
End Sub

' virtual bool add(IClientConnection* conn) override
Friend Function add(ByVal conn As IClientConnection) As Boolean
    Dim tcpConn As TCPClientConnection, fd As Long, i As Long
    Set tcpConn = conn.tcpImpl                  ' dynamic_cast<TCPClientConnection*>(conn)
    If Not (Not tcpConn Is Nothing) Then ASSERT_FAILURE "tcpConn", "lib/netplay/tcp/select_descriptor_set.h:60", "add", "Invalid connection type: expected TCPClientConnection": add = False: Exit Function
    If Not (conn.isValid()) Then ASSERT_FAILURE "tcpConn->isValid()", "lib/netplay/tcp/select_descriptor_set.h:61", "add", "Connection object is not valid: socket is not opened": add = False: Exit Function
    fd = tcpConn.getRawSocketFd()
    ' FD_SET(fd, &fds_)
    For i = 0 To fds_.fd_count - 1
        If fds_.fd_array(i) = fd Then Exit For
    Next
    If i = fds_.fd_count Then
        If fds_.fd_count < FD_SETSIZE Then
            fds_.fd_array(i) = fd
            fds_.fd_count = fds_.fd_count + 1
        End If
    End If
    If ULT(maxfd_, fd) Then maxfd_ = fd         ' std::max(maxfd_, fd)
    add = True
End Function

' virtual bool remove(IClientConnection* conn) override
Friend Function remove(ByVal conn As IClientConnection) As Boolean
    Dim tcpConn As TCPClientConnection, fd As Long, i As Long
    Set tcpConn = conn.tcpImpl
    If Not (Not tcpConn Is Nothing) Then ASSERT_FAILURE "tcpConn", "lib/netplay/tcp/select_descriptor_set.h:72", "remove", "Invalid connection type: expected TCPClientConnection": remove = False: Exit Function
    fd = tcpConn.getRawSocketFd()
    ' FD_CLR(fd, &fds_)
    For i = 0 To fds_.fd_count - 1
        If fds_.fd_array(i) = fd Then
            Do While i < fds_.fd_count - 1
                fds_.fd_array(i) = fds_.fd_array(i + 1)
                i = i + 1
            Loop
            fds_.fd_count = fds_.fd_count - 1
            Exit For
        End If
    Next
    remove = True
End Function

' virtual void clear() override
Friend Sub clear()
    fds_.fd_count = 0                           ' FD_ZERO(&fds_)
    maxfd_ = 0
End Sub

' virtual net::result<int> poll(std::chrono::milliseconds timeout) override
Friend Function poll(ByVal timeoutMs As Long, ec As Long) As Long
    Dim ret As Long
    ec = 0
    Do
        ret = pollImpl(timeoutMs)
    Loop While ret = SOCKET_ERROR And (getSockErr() = EINTR Or getSockErr() = EAGAIN)
    If ret = SOCKET_ERROR Then
        ec = make_network_error_code(getSockErr())
        Exit Function
    End If
    poll = ret
End Function

' virtual ::tl::expected<bool, ErroredState> isSet(const IClientConnection* conn) const override
Friend Function isSet(ByVal conn As IClientConnection, errState As Long) As Boolean
    Dim tcpConn As TCPClientConnection, fd As Long, i As Long
    errState = 0
    Set tcpConn = conn.tcpImpl
    If Not (Not tcpConn Is Nothing) Then ASSERT_FAILURE "tcpConn", "lib/netplay/tcp/select_descriptor_set.h:104", "isSet", "Invalid connection type: expected TCPClientConnection": errState = 1: Exit Function
    fd = tcpConn.getRawSocketFd()
    ' FD_ISSET(tcpConn->getRawSocketFd(), &fds_): __WSAFDIsSet
    For i = 0 To fds_.fd_count - 1
        If fds_.fd_array(i) = fd Then
            isSet = True
            Exit Function
        End If
    Next
End Function

' virtual bool empty() const override
Friend Function empty_() As Boolean
    empty_ = (maxfd_ = 0)
End Function

' int pollImpl(std::chrono::milliseconds timeout)
Private Function pollImpl(ByVal timeoutMs As Long) As Long
    Dim tv As timeval
    tv.tv_sec = timeoutMs \ 1000
    tv.tv_usec = (timeoutMs Mod 1000) * 1000
    Select Case eventType_
    Case 0                                      ' PollEventType::READABLE
        pollImpl = select_(maxfd_ + 1, VarPtr(fds_), 0, 0, tv)
        sockCalled
    Case 1                                      ' PollEventType::WRITABLE
        pollImpl = select_(maxfd_ + 1, 0, VarPtr(fds_), 0, tv)
        sockCalled
    Case Else
        pollImpl = -1
    End Select
End Function
