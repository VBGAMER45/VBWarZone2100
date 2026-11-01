VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TCPConnectionProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/netplay/tcp/tcp_connection_provider.cpp, tcp_connection_provider.h: class tcp::TCPConnectionProvider, the TCP
' part of a WzConnectionProvider (which holds this object; the methods that hand `*this` on take the base, self).
' Windows' newDescriptorSet: select() (SelectDescriptorSet), as 4.7.0's WZ_OS_WIN branch.
' This file includes netsocket.h and so tcp/sock_error.h; EINVAL is not among its defines (MSVC's 22).
Option Explicit

Private Const EINVAL As Long = 22
Private Const TCP_DIRECT As Long = 0

Private initialized_ As Boolean
Private addressResolver_ As TCPAddressResolver

' virtual void initialize() override
Friend Sub initialize()
    If initialized_ Then Exit Sub
    initialized_ = True
    SOCKETinit
    Set addressResolver_ = New TCPAddressResolver
End Sub

' virtual void shutdown() override
Friend Sub shutdown()
    If Not initialized_ Then Exit Sub
    initialized_ = False
    Set addressResolver_ = Nothing
    SOCKETshutdown
End Sub

' virtual ConnectionProviderType type() const noexcept override
Friend Function type_() As Long
    type_ = TCP_DIRECT
End Function

' virtual net::result<std::unique_ptr<IConnectionAddress>> resolveHost(const char* host, uint16_t port) const override
Friend Function resolveHost(ByVal host As String, ByVal port As Long, ec As Long) As IConnectionAddress
    ec = 0
    If Not (Not addressResolver_ Is Nothing) Then ASSERT_FAILURE "addressResolver_ != nullptr", "lib/netplay/tcp/tcp_connection_provider.cpp:69", "resolveHost", "Address resolver not initialized": ec = make_network_error_code(EINVAL): Exit Function
    Set resolveHost = addressResolver_.resolveHost(host, port, ec)
End Function

' virtual net::result<IListenSocket*> openListenSocket(uint16_t port) override
Friend Function openListenSocket(ByVal self As WzConnectionProvider, ByVal port As Long, ec As Long) As IListenSocket
    Dim res As Long, ls As TCPListenSocket
    res = netsocket_c.socketListen(port, ec)
    If ec <> 0 Then
        Exit Function
    End If
    Set ls = New TCPListenSocket
    Set openListenSocket = ls.construct(self, PendingWritesManagerMap_get_type(type_()), res)
End Function

' virtual net::result<IClientConnection*> openClientConnectionAny(const IConnectionAddress& addr, unsigned timeout) override
Friend Function openClientConnectionAny(ByVal self As WzConnectionProvider, ByVal addr As IConnectionAddress, ByVal timeout As Long, ec As Long) As IClientConnection
    Dim tcpAddr As TCPConnectionAddress, rawAddr As Long, res As Long, conn As TCPClientConnection
    ec = 0
    If Not addr Is Nothing Then Set tcpAddr = addr.tcpImpl
    If Not (Not tcpAddr Is Nothing) Then ASSERT_FAILURE "tcpAddr != nullptr", "lib/netplay/tcp/tcp_connection_provider.cpp:86", "openClientConnectionAny", "Expected TCPConnectionAddress instance": ec = make_network_error_code(EINVAL): Exit Function
    rawAddr = tcpAddr.asRawSocketAddress()
    res = socketOpenAny(rawAddr, timeout, ec)
    If ec <> 0 Then
        Exit Function
    End If
    Set conn = New TCPClientConnection
    Set openClientConnectionAny = conn.construct(self, PendingWritesManagerMap_get_type(type_()), res)
End Function

' virtual IConnectionPollGroup* newConnectionPollGroup() override
Friend Function newConnectionPollGroup(ByVal self As WzConnectionProvider) As IConnectionPollGroup
    Dim g As TCPConnectionPollGroup
    Set g = New TCPConnectionPollGroup
    Set newConnectionPollGroup = g.construct(self)
End Function

' virtual std::unique_ptr<IDescriptorSet> newDescriptorSet(PollEventType eventType) override
Friend Function newDescriptorSet(ByVal eventType As Long) As SelectDescriptorSet
    ' For now, use `select()` on Windows instead of `poll()` because of a bug in Windows versions prior to
    ' "Windows 10 2004", which can lead to `poll()` function timing out on socket connection errors instead of
    ' returning an error early.
    Dim s As SelectDescriptorSet
    Select Case eventType
    Case 0, 1                                   ' PollEventType::READABLE, WRITABLE
        Set s = New SelectDescriptorSet
        s.construct eventType
        Set newDescriptorSet = s
    Case Else
        ASSERT_FAILURE "false", "lib/netplay/tcp/tcp_connection_provider.cpp:132", "newDescriptorSet", "Unexpected PollEventType value: %d", eventType
    End Select
End Function

' virtual void processConnectionStateChanges() override {}
Friend Sub processConnectionStateChanges()
End Sub

' virtual PortMappingInternetProtocolMask portMappingProtocolTypes() const override
Friend Function portMappingProtocolTypes() As Long
    portMappingProtocolTypes = 1 Or 2           ' PortMappingInternetProtocol::TCP_IPV4 | TCP_IPV6
End Function

' virtual void disposeConnection(IClientConnection* conn) override
Friend Sub disposeConnection(ByVal conn As IClientConnection)
    ' TODO: Remove the connection from the owning poll group (if there's any).
    ' For this, we'll need to have a reference to the poll group in the `IClientConnection` interface.
End Sub
