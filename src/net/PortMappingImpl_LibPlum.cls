VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PortMappingImpl_LibPlum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/netplay/port_mapping_manager_impl_libplum.cpp, port_mapping_manager_impl_libplum.h: class
' PortMappingImpl_LibPlum (its anonymous namespace's callbacks are port_mapping_manager_impl_libplum_c's).
Option Explicit

Private m_isInit As Boolean                     ' bool m_isInit = false;

' PortMappingImpl_LibPlum::~PortMappingImpl_LibPlum()
Private Sub Class_Terminate()
    shutdown
End Sub

' PortMappingImpl::Type PortMappingImpl_LibPlum::get_impl_type() const
Public Function get_impl_type() As Long
    get_impl_type = PortMappingImpl_Type_Libplum
End Function

' int PortMappingImpl_LibPlum::get_discovery_timeout_seconds() const
Public Function get_discovery_timeout_seconds() As Long
    get_discovery_timeout_seconds = 10  ' 10 seconds for Libplum, to ensure we don't wait too long before falling-back to Miniupnpc
End Function

' bool PortMappingImpl_LibPlum::init()
Public Function init() As Boolean
    Dim config As plum_config_t, res As Long
    config.log_level = PLUM_LOG_LEVEL_INFO
    config.log_callback = FN_PlumLogCallback
    res = plum_init(config)
    m_isInit = (res = PLUM_ERR_SUCCESS)
    init = res = PLUM_ERR_SUCCESS
End Function

' bool PortMappingImpl_LibPlum::shutdown()
Public Function shutdown() As Boolean
    If Not m_isInit Then
        shutdown = False
        Exit Function
    End If
    plum_cleanup
    m_isInit = False
    shutdown = True
End Function

' int PortMappingImpl_LibPlum::create_port_mapping(uint16_t port, PortMappingInternetProtocol protocol)
Public Function create_port_mapping(ByVal port As Long, ByVal protocol As Long) As Long
    Dim m As plum_mapping_t
    If protocol = PortMappingInternetProtocol_TCP_IPV4 Or protocol = PortMappingInternetProtocol_TCP_IPV6 Then
        m.protocol = PLUM_IP_PROTOCOL_TCP
    Else ' UDP_IPV4 / UDP_IPV6
        m.protocol = PLUM_IP_PROTOCOL_UDP
    End If
    m.internal_port = port
    m.external_port = port  ' suggest an external port the same as the internal port (the router may decide otherwise)

    create_port_mapping = plum_create_mapping(m, FN_PlumMappingCallback)
End Function

' bool PortMappingImpl_LibPlum::destroy_port_mapping(int mappingId)
Public Function destroy_port_mapping(ByVal mappingId As Long) As Boolean
    Dim result As Long
    result = plum_destroy_mapping(mappingId)
    destroy_port_mapping = result = PLUM_ERR_SUCCESS
End Function
