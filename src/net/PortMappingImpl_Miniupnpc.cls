VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PortMappingImpl_Miniupnpc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' STAND-IN for lib/netplay/port_mapping_manager_impl_miniupnpc.cpp, class PortMappingImpl_Miniupnpc: miniupnpc (4.7.0's
' fallback after libplum) is not ported (cluster J part 2, the user's choice: libplum only). Its slot is unavailable:
' init fails, as when miniupnpc cannot initialise, so PortMappingManager's switchToNextWorkingImpl passes over it.
Option Explicit

' bool init() override
Public Function init() As Boolean
    init = False
End Function

' bool shutdown() override
Public Function shutdown() As Boolean
    shutdown = False
End Function

' int create_port_mapping(uint16_t port, PortMappingInternetProtocol protocol) override
Public Function create_port_mapping(ByVal port As Long, ByVal protocol As Long) As Long
    create_port_mapping = WZ_PORT_MAPPING_ID_INVALID
End Function

' bool destroy_port_mapping(int mappingId) override
Public Function destroy_port_mapping(ByVal mappingId As Long) As Boolean
    destroy_port_mapping = False
End Function

' PortMappingImpl::Type get_impl_type() const override
Public Function get_impl_type() As Long
    get_impl_type = PortMappingImpl_Type_Miniupnpc
End Function

' int get_discovery_timeout_seconds() const override
Public Function get_discovery_timeout_seconds() As Long
    get_discovery_timeout_seconds = 0
End Function
