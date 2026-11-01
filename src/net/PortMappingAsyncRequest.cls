VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PortMappingAsyncRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/netplay/port_mapping_manager.cpp, port_mapping_manager.h: class PortMappingManager::PortMappingAsyncRequest
' (a PortMappingAsyncRequestHandle / PortMappingAsyncRequestPtr is this object; Nothing: nullptr). A TimePoint is
' steady_clock's milliseconds (steadyNowMs); TimePoint::max() is TIMEPOINT_MAX. callbacks (a CallbackVector of
' SuccessCallback / FailureCallback pairs) is a Collection of callback numbers (port_mapping_manager_c's PMCB_*).
Option Explicit

Public deadline As Double                       ' TimePoint deadline = TimePoint::max();
Public mappingId As Long                        ' MappingId mappingId;
Public callbacks As Collection                  ' CallbackVector callbacks;
Public s As PortMappingDiscoveryStatus          ' PortMappingDiscoveryStatus s = PortMappingDiscoveryStatus::NOT_STARTED;
Public discoveredIPaddress As String
Public discoveredPort As Long                   ' uint16_t discoveredPort = 0;
Public sourcePort As Long                       ' uint16_t sourcePort = 0;
Public protocol As Long                         ' PortMappingInternetProtocol protocol = PortMappingInternetProtocol::TCP_IPV4;
Public impl As PortMappingImpl                  ' std::shared_ptr<PortMappingImpl> impl = nullptr;

' PortMappingManager::PortMappingAsyncRequest::PortMappingAsyncRequest()
Private Sub Class_Initialize()
    deadline = TIMEPOINT_MAX
    mappingId = WZ_PORT_MAPPING_ID_INVALID
    Set callbacks = New Collection
    s = PortMappingDiscoveryStatus_NOT_STARTED
    protocol = PortMappingInternetProtocol_TCP_IPV4
End Sub

' PortMappingDiscoveryStatus status() const
Public Function status() As PortMappingDiscoveryStatus
    status = s
End Function

' bool in_progress() const
Public Function in_progress() As Boolean
    in_progress = s = PortMappingDiscoveryStatus_IN_PROGRESS
End Function

' bool failed() const
Public Function failed() As Boolean
    failed = s = PortMappingDiscoveryStatus_FAILURE Or s = PortMappingDiscoveryStatus_TIMEOUT
End Function

' bool succeeded() const
Public Function succeeded() As Boolean
    succeeded = s = PortMappingDiscoveryStatus_SUCCESS
End Function

' bool timed_out(TimePoint t) const
Public Function timed_out(ByVal t As Double) As Boolean
    timed_out = t >= deadline
End Function

' void set_port_mapping_in_progress(TimePoint deadline, MappingId mappingId, const std::shared_ptr<PortMappingImpl>& impl)
Public Sub set_port_mapping_in_progress(ByVal deadline_ As Double, ByVal mappingId_ As Long, ByVal impl_ As PortMappingImpl)
    If Not (s = PortMappingDiscoveryStatus_NOT_STARTED Or s = PortMappingDiscoveryStatus_IN_PROGRESS) Then ASSERT_FAILURE "s == PortMappingDiscoveryStatus::NOT_STARTED || s == PortMappingDiscoveryStatus::IN_PROGRESS", "lib/netplay/port_mapping_manager.cpp:90", "set_port_mapping_in_progress", "Invalid state"
    If impl_ Is Nothing Then ASSERT_FAILURE "impl_ != nullptr", "lib/netplay/port_mapping_manager.cpp:91", "set_port_mapping_in_progress", "Missing impl?"
    deadline = deadline_
    s = PortMappingDiscoveryStatus_IN_PROGRESS
    mappingId = mappingId_
    Set impl = impl_
End Sub

' bool destroy_mapping()
Public Function destroy_mapping() As Boolean
    Dim result As Boolean
    ' Destroy the internal mapping.
    If mappingId <> WZ_PORT_MAPPING_ID_INVALID Then
        If impl Is Nothing Then ASSERT_FAILURE "impl != nullptr", "lib/netplay/port_mapping_manager.cpp:103", "destroy_mapping", "Missing impl?": destroy_mapping = False: Exit Function
        result = impl.destroy_port_mapping(mappingId)
        mappingId = WZ_PORT_MAPPING_ID_INVALID
        s = PortMappingDiscoveryStatus_NOT_STARTED
        destroy_mapping = result
    Else
        destroy_mapping = False
    End If
End Function
