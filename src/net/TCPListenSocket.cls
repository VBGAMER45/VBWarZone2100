VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TCPListenSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/netplay/tcp/tcp_listen_socket.cpp, tcp_listen_socket.h: class tcp::TCPListenSocket, the TCP part of an
' IListenSocket (which holds this object and passes its base members to accept).
Option Explicit

Private listenSocket_ As Long                   ' tcp::Socket* (netsocket_c's handle)

' explicit TCPListenSocket(WzConnectionProvider& connProvider, WzCompressionProvider& compressionProvider,
'     PendingWritesManager& pwm, tcp::Socket* rawSocket): makes its base, the IListenSocket it returns
Friend Function construct(ByVal connProvider As WzConnectionProvider, ByVal pwm As PendingWritesManager, ByVal rawSocket As Long) As IListenSocket
    Dim base As IListenSocket
    Set base = New IListenSocket
    base.construct connProvider, pwm, Me
    listenSocket_ = rawSocket
    Set construct = base
End Function

' ~TCPListenSocket()
Friend Sub destroy_()
    If listenSocket_ <> 0 Then
        socketCloseNow listenSocket_
        listenSocket_ = 0                        ' (VB) a second close does nothing
    End If
End Sub

' virtual IClientConnection* accept() override (connProvider_, pwm_: the base's members)
Friend Function accept(ByVal connProvider As WzConnectionProvider, ByVal pwm As PendingWritesManager) As IClientConnection
    Dim s As Long, conn As TCPClientConnection
    If Not (listenSocket_ <> 0) Then ASSERT_FAILURE "listenSocket_ != nullptr", "lib/netplay/tcp/tcp_listen_socket.cpp:44", "accept", "Internal socket handle shouldn't be null!"
    If listenSocket_ = 0 Then
        Exit Function
    End If
    s = socketAccept(listenSocket_)
    If s = 0 Then
        Exit Function
    End If
    Set conn = New TCPClientConnection
    Set accept = conn.construct(connProvider, pwm, s)
End Function

' virtual IPVersionsMask supportedIpVersions() const override
Friend Function supportedIpVersions() As Long
    Dim resMask As Long
    resMask = 0
    If socketHasIPv4(listenSocket_) Then
        resMask = resMask Or 1                   ' IPVersions::IPV4
    End If
    If socketHasIPv6(listenSocket_) Then
        resMask = resMask Or 2                   ' IPVersions::IPV6
    End If
    supportedIpVersions = resMask
End Function

' virtual ConnectionProviderType type() const noexcept override
Friend Function type_() As Long
    type_ = 0                                    ' ConnectionProviderType::TCP_DIRECT
End Function
