VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IConnectionPollGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' lib/netplay/connection_poll_group.h: class IConnectionPollGroup (cluster J's Contract C), PublicNotCreatable:
' WzConnectionProvider.newConnectionPollGroup makes it. The derived class is WzNet's private
' TCPConnectionPollGroup, which this object holds. C++'s `delete` of a poll group is dropping the object.
Option Explicit

Private tcp_ As TCPConnectionPollGroup          ' (VB) the derived object

Friend Sub construct(ByVal derived As TCPConnectionPollGroup)
    Set tcp_ = derived
End Sub

' virtual net::result<int> checkConnectionsReadable(std::chrono::milliseconds timeout) = 0
Public Function checkConnectionsReadable(ByVal timeoutMs As Long, ec As Long) As Long
    checkConnectionsReadable = tcp_.checkConnectionsReadable(timeoutMs, ec)
End Function

' virtual void add(IClientConnection* conn) = 0
Public Sub add(ByVal conn As IClientConnection)
    tcp_.add conn
End Sub

' virtual void remove(IClientConnection* conn) = 0
Public Sub remove(ByVal conn As IClientConnection)
    tcp_.remove conn
End Sub
