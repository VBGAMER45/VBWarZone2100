VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PortMappingImpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/netplay/port_mapping_manager.cpp, port_mapping_manager.h: class PortMappingImpl, the port mapping
' implementations' base class (a std::shared_ptr<PortMappingImpl> is this object). The derived object it holds is a
' PortMappingImpl_LibPlum or a PortMappingImpl_Miniupnpc (miniupnpc is not ported: that class is a stand-in whose
' init fails); each virtual function is the derived object's.
Option Explicit

Private libplum_ As PortMappingImpl_LibPlum     ' (VB) the derived object, one of the two
Private miniupnpc_ As PortMappingImpl_Miniupnpc

Friend Sub constructLibPlum(ByVal derived As PortMappingImpl_LibPlum)
    Set libplum_ = derived
End Sub

Friend Sub constructMiniupnpc(ByVal derived As PortMappingImpl_Miniupnpc)
    Set miniupnpc_ = derived
End Sub

' virtual bool init() = 0
Public Function init() As Boolean
    If Not libplum_ Is Nothing Then init = libplum_.init() Else init = miniupnpc_.init()
End Function

' virtual bool shutdown() = 0
Public Function shutdown() As Boolean
    If Not libplum_ Is Nothing Then shutdown = libplum_.shutdown() Else shutdown = miniupnpc_.shutdown()
End Function

' virtual int create_port_mapping(uint16_t port, PortMappingInternetProtocol protocol) = 0
Public Function create_port_mapping(ByVal port As Long, ByVal protocol As Long) As Long
    If Not libplum_ Is Nothing Then
        create_port_mapping = libplum_.create_port_mapping(port, protocol)
    Else
        create_port_mapping = miniupnpc_.create_port_mapping(port, protocol)
    End If
End Function

' virtual bool destroy_port_mapping(int mappingId) = 0
Public Function destroy_port_mapping(ByVal mappingId As Long) As Boolean
    If Not libplum_ Is Nothing Then
        destroy_port_mapping = libplum_.destroy_port_mapping(mappingId)
    Else
        destroy_port_mapping = miniupnpc_.destroy_port_mapping(mappingId)
    End If
End Function

' virtual Type get_impl_type() const = 0
Public Function get_impl_type() As Long
    If Not libplum_ Is Nothing Then get_impl_type = libplum_.get_impl_type() Else get_impl_type = miniupnpc_.get_impl_type()
End Function

' virtual int get_discovery_timeout_seconds() const = 0
Public Function get_discovery_timeout_seconds() As Long
    If Not libplum_ Is Nothing Then
        get_discovery_timeout_seconds = libplum_.get_discovery_timeout_seconds()
    Else
        get_discovery_timeout_seconds = miniupnpc_.get_discovery_timeout_seconds()
    End If
End Function
