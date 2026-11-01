Attribute VB_Name = "netjoin_c"
' lib/netplay/netjoin.cpp, netjoin.h, whole: struct JoinConnectionDescription (the exe's: the joining screen, the
' game browser, multijoin_helpers and multiint's kick redirect use it; a UDT does not cross into WzNet, so the
' exe has its own copy of these few lines). JoinConnectionType's GNS_DIRECT exists only with
' WZ_GNS_NETWORK_BACKEND_ENABLED, which this port does not define (GameNetworkingSockets is not ported).
Option Explicit

' enum class JoinConnectionDescription::JoinConnectionType
Public Enum JoinConnectionType
    JoinConnectionType_TCP_DIRECT
End Enum

' struct JoinConnectionDescription
Public Type JoinConnectionDescription
    host As String
    port As Long                        ' uint32_t = 0
    type_ As JoinConnectionType         ' type = JoinConnectionType::TCP_DIRECT
End Type

' JoinConnectionDescription(const std::string& host, uint32_t port) / (JoinConnectionType t, const std::string& host, uint32_t port)
Public Function JoinConnectionDescription_make(ByVal t As JoinConnectionType, ByVal host As String, ByVal port As Long) As JoinConnectionDescription
    Dim d As JoinConnectionDescription
    d.host = host
    d.port = port
    d.type_ = t
    JoinConnectionDescription_make = d
End Function
