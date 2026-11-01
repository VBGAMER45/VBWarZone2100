VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IConnectionAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' lib/netplay/connection_address.h: struct IConnectionAddress (cluster J's Contract C), PublicNotCreatable:
' WzConnectionProvider.resolveHost makes it (a std::unique_ptr<IConnectionAddress>: the last reference's release
' frees the address list). The derived class is WzNet's private TCPConnectionAddress, which this object holds.
Option Explicit

Private tcp_ As TCPConnectionAddress            ' (VB) the derived object

Friend Sub construct(ByVal derived As TCPConnectionAddress)
    Set tcp_ = derived
End Sub

' (VB) dynamic_cast<const TCPConnectionAddress*>(&addr)
Friend Function tcpImpl() As TCPConnectionAddress
    Set tcpImpl = tcp_
End Function

' virtual net::result<std::string> toString() const = 0
Public Function toString(ec As Long) As String
    toString = tcp_.toString(ec)
End Function
