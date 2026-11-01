Attribute VB_Name = "pending_writes_manager_map_c"
' lib/netplay/pending_writes_manager_map.cpp, pending_writes_manager_map.h: the global singleton mapping connection
' provider types to PendingWritesManager instances (PendingWritesManagerMap::instance() is this module).
' get(ConnectionProviderType) is PendingWritesManagerMap_get_type, get(const WzConnectionProvider&)
' PendingWritesManagerMap_get. PendingWritesManager_pollAll is the port's: every manager's thread pass
' (cluster J's Contract C; netplay calls it from NETrecvNet, NETrecvGame and NETflush).
Option Explicit

Private pendingWritesManagers_ As Collection    ' std::unordered_map<ConnectionProviderType, std::unique_ptr<PendingWritesManager>>

Private Function managers() As Collection
    If pendingWritesManagers_ Is Nothing Then Set pendingWritesManagers_ = New Collection
    Set managers = pendingWritesManagers_
End Function

' PendingWritesManager& get(ConnectionProviderType pt)
Public Function PendingWritesManagerMap_get_type(ByVal pt As Long) As PendingWritesManager
    Dim connProvider As WzConnectionProvider, pwm As PendingWritesManager
    If Not (ConnectionProviderRegistry_IsRegistered(pt)) Then ASSERT_FAILURE "cpr.IsRegistered(pt)", "lib/netplay/pending_writes_manager_map.cpp:36", "get", "Connection provider %d is not registered", pt
    On Error Resume Next
    Set pwm = managers().Item(CStr(pt))
    On Error GoTo 0
    If Not pwm Is Nothing Then
        Set PendingWritesManagerMap_get_type = pwm
        Exit Function
    End If
    Set connProvider = ConnectionProviderRegistry_Get(pt)
    If Not (Not connProvider Is Nothing) Then ASSERT_FAILURE "connProvider != nullptr", "lib/netplay/pending_writes_manager_map.cpp:43", "get", "Null connection provider"
    Set pwm = New PendingWritesManager
    pwm.initialize connProvider
    managers().Add pwm, CStr(pt)
    Set PendingWritesManagerMap_get_type = pwm
End Function

' PendingWritesManager& get(const WzConnectionProvider& connProvider)
Public Function PendingWritesManagerMap_get(ByVal connProvider As WzConnectionProvider) As PendingWritesManager
    Set PendingWritesManagerMap_get = PendingWritesManagerMap_get_type(connProvider.type_())
End Function

' void Shutdown()
Public Sub PendingWritesManagerMap_Shutdown()
    Dim pwm As PendingWritesManager
    For Each pwm In managers()
        pwm.deinitialize
    Next
    Set pendingWritesManagers_ = New Collection
End Sub

' (VB) the pending-writes threads' work: one pass of each manager (see PendingWritesManager.cls)
Public Sub PendingWritesManager_pollAll()
    Dim pwm As PendingWritesManager
    If pendingWritesManagers_ Is Nothing Then Exit Sub
    For Each pwm In pendingWritesManagers_
        pwm.threadPass
    Next
End Sub
