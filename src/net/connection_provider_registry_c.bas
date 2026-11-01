Attribute VB_Name = "connection_provider_registry_c"
' lib/netplay/connection_provider_registry.cpp, connection_provider_registry.h: class ConnectionProviderRegistry
' (Instance() is this module). A std::shared_ptr<WzConnectionProvider> is the object. enum class
' ConnectionProviderType: TCP_DIRECT 0, GNS_DIRECT 1 (no GameNetworkingSockets backend: GNS is never registered,
' WZ_GNS_NETWORK_BACKEND_ENABLED being undefined). The C++ throws std::runtime_error: Err.Raise here.
Option Explicit

Private Const TCP_DIRECT As Long = 0

Private registeredProviders_ As Collection      ' std::unordered_map<ConnectionProviderType, std::shared_ptr<WzConnectionProvider>>

Private Function providers() As Collection
    If registeredProviders_ Is Nothing Then Set registeredProviders_ = New Collection
    Set providers = registeredProviders_
End Function

' std::shared_ptr<WzConnectionProvider> Get(ConnectionProviderType pt)
Public Function ConnectionProviderRegistry_Get(ByVal pt As Long) As WzConnectionProvider
    Dim p As WzConnectionProvider
    On Error Resume Next
    Set p = providers().Item(CStr(pt))
    On Error GoTo 0
    If p Is Nothing Then
        Err.Raise vbObjectError + 513, "ConnectionProviderRegistry::Get", "Attempt to get nonexistent connection provider"
    End If
    Set ConnectionProviderRegistry_Get = p
End Function

' bool IsRegistered(ConnectionProviderType) const
Public Function ConnectionProviderRegistry_IsRegistered(ByVal pt As Long) As Boolean
    Dim p As WzConnectionProvider
    On Error Resume Next
    Set p = providers().Item(CStr(pt))
    ConnectionProviderRegistry_IsRegistered = Not p Is Nothing
End Function

' void Register(ConnectionProviderType pt)
Public Sub ConnectionProviderRegistry_Register(ByVal pt As Long)
    Dim p As WzConnectionProvider, tcp As TCPConnectionProvider
    If ConnectionProviderRegistry_IsRegistered(pt) Then
        Exit Sub
    End If
    Select Case pt
    Case TCP_DIRECT
        Set tcp = New TCPConnectionProvider
        Set p = New WzConnectionProvider
        p.construct tcp
        providers().Add p, CStr(pt)
    Case Else
        Err.Raise vbObjectError + 514, "ConnectionProviderRegistry::Register", "Unknown connection provider type"
    End Select
End Sub

' void Shutdown()
Public Sub ConnectionProviderRegistry_Shutdown()
    Dim p As WzConnectionProvider
    For Each p In providers()
        p.shutdown
        p.markShutdownByRegistry                 ' (VB) the weak_ptrs to it expire (IClientConnection::connProviderLock)
    Next
    Set registeredProviders_ = New Collection
End Sub
