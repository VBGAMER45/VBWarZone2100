VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzConnectionProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' lib/netplay/wz_connection_provider.cpp, wz_connection_provider.h: class WzConnectionProvider (cluster J's
' Contract C), PublicNotCreatable: ConnectionProviderRegistry_Get gives it (a std::shared_ptr<WzConnectionProvider>
' is this object). The derived class is WzNet's private TCPConnectionProvider, which this object holds; the
' non-virtual openClientConnectionAsync is here.
' No threads: openClientConnectionAsync opens the connection inside the call (openDirectConnectionAsyncImpl's work,
' 4.7.0's background thread) and calls callback.onResult before it returns; the joining screen's handler defers
' its work with wzAsyncExecOnMainThread, as 4.7.0's does.
Option Explicit

Private tcp_ As TCPConnectionProvider           ' (VB) the derived object
Private shutdownByRegistry_ As Boolean

Friend Sub construct(ByVal derived As TCPConnectionProvider)
    Set tcp_ = derived
End Sub

' (VB) ConnectionProviderRegistry::Shutdown let go of this provider: the weak_ptrs to it expire
Friend Sub markShutdownByRegistry()
    shutdownByRegistry_ = True
End Sub

Friend Function isShutdownByRegistry() As Boolean
    isShutdownByRegistry = shutdownByRegistry_
End Function

' virtual void initialize() = 0
Public Sub initialize()
    tcp_.initialize
End Sub

' virtual void shutdown() = 0
Public Sub shutdown()
    tcp_.shutdown
End Sub

' virtual ConnectionProviderType type() const noexcept = 0
Public Function type_() As Long
    type_ = tcp_.type_()
End Function

' virtual net::result<std::unique_ptr<IConnectionAddress>> resolveHost(const char* host, uint16_t port) const = 0
Public Function resolveHost(ByVal host As String, ByVal port As Long, ec As Long) As IConnectionAddress
    Set resolveHost = tcp_.resolveHost(host, port And &HFFFF&, ec)
End Function

' virtual net::result<IListenSocket*> openListenSocket(uint16_t port) = 0
Public Function openListenSocket(ByVal port As Long, ec As Long) As IListenSocket
    Set openListenSocket = tcp_.openListenSocket(Me, port And &HFFFF&, ec)
End Function

' virtual net::result<IClientConnection*> openClientConnectionAny(const IConnectionAddress& addr, unsigned timeout) = 0
Public Function openClientConnectionAny(ByVal addr As IConnectionAddress, ByVal timeout As Long, ec As Long) As IClientConnection
    Set openClientConnectionAny = tcp_.openClientConnectionAny(Me, addr, timeout, ec)
End Function

' OpenConnectionResult openClientConnectionSyncImpl(const char* host, uint32_t port, std::chrono::milliseconds timeout,
'     std::shared_ptr<WzConnectionProvider> connProvider) (wz_connection_provider.cpp's anonymous namespace)
Private Function openClientConnectionSyncImpl(ByVal host As String, ByVal port As Long, ByVal timeoutMs As Long, ByVal connProvider As WzConnectionProvider) As OpenConnectionResult
    Dim r As OpenConnectionResult, addrResult As IConnectionAddress, hostsErr As Long, hostsErrMsg As String
    Dim connRes As IClientConnection, connErr As Long, connErrMsg As String
    Set r = New OpenConnectionResult
    Set openClientConnectionSyncImpl = r
    If connProvider Is Nothing Then
        r.construct_error make_generic_error_code(errc_connection_aborted), "Connection provider no longer valid"
        Exit Function
    End If
    Set addrResult = connProvider.resolveHost(host, port, hostsErr)
    If hostsErr <> 0 Then
        hostsErrMsg = net_errorMessage(hostsErr)
        r.construct_error hostsErr, va("Cannot resolve host ""%s"": [%d]: %s", host, net_errorValue(hostsErr), hostsErrMsg)
        Exit Function
    End If
    Set connRes = connProvider.openClientConnectionAny(addrResult, timeoutMs, connErr)
    If connErr <> 0 Then
        connErrMsg = net_errorMessage(connErr)
        r.construct_error connErr, va("Cannot resolve host ""%s"": [%d]: %s", host, net_errorValue(connErr), connErrMsg)
        Exit Function
    End If
    r.construct_socket connRes
End Function

' virtual bool openClientConnectionAsync(const std::string& host, uint32_t port, std::chrono::milliseconds timeout,
'     OpenConnectionToHostResultCallback callback)
Public Function openClientConnectionAsync(ByVal host As String, ByVal port As Long, ByVal timeoutMs As Long, ByVal callback As IOpenConnectionCallback) As Boolean
    Dim strongConnProvider As WzConnectionProvider
    ' (VB) the background thread's openDirectConnectionAsyncImpl, here: pRequest->connProvider.lock(), then
    ' pRequestInfo->callback(openClientConnectionSyncImpl(host, port, timeout, strongConnProvider))
    If Not shutdownByRegistry_ Then Set strongConnProvider = Me
    callback.onResult openClientConnectionSyncImpl(host, port, timeoutMs, strongConnProvider)
    openClientConnectionAsync = True
End Function

' virtual IConnectionPollGroup* newConnectionPollGroup() = 0
Public Function newConnectionPollGroup() As IConnectionPollGroup
    Set newConnectionPollGroup = tcp_.newConnectionPollGroup(Me)
End Function

' virtual std::unique_ptr<IDescriptorSet> newDescriptorSet(PollEventType eventType) = 0 (PollEventType: READABLE 0,
' WRITABLE 1; WzNet-private, as the descriptor sets are)
Friend Function newDescriptorSet(ByVal eventType As Long) As SelectDescriptorSet
    Set newDescriptorSet = tcp_.newDescriptorSet(eventType)
End Function

' virtual void processConnectionStateChanges() = 0
Public Sub processConnectionStateChanges()
    tcp_.processConnectionStateChanges
End Sub

' virtual PortMappingInternetProtocolMask portMappingProtocolTypes() const = 0
Public Function portMappingProtocolTypes() As Long
    portMappingProtocolTypes = tcp_.portMappingProtocolTypes()
End Function

' virtual void disposeConnection(IClientConnection* conn) = 0
Public Sub disposeConnection(ByVal conn As IClientConnection)
    tcp_.disposeConnection conn
End Sub
