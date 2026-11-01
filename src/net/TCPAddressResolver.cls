VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TCPAddressResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/netplay/tcp/tcp_address_resolver.cpp, tcp_address_resolver.h and address_resolver.h (class IAddressResolver:
' this is its one implementation): class tcp::TCPAddressResolver.
Option Explicit

' explicit TCPAddressResolver()
Private Sub Class_Initialize()
    SOCKETinit
End Sub

' virtual ~TCPAddressResolver() override
Private Sub Class_Terminate()
    SOCKETshutdown
End Sub

' virtual net::result<std::unique_ptr<IConnectionAddress>> resolveHost(const char* host, uint16_t port) const override
Friend Function resolveHost(ByVal host As String, ByVal port As Long, ec As Long) As IConnectionAddress
    Dim resolved As Long, addr As TCPConnectionAddress
    resolved = netsocket_c.resolveHost(host, port, ec)
    If ec <> 0 Then
        Exit Function
    End If
    Set addr = New TCPConnectionAddress
    Set resolveHost = addr.construct(resolved)
End Function
