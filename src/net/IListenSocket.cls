VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IListenSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' lib/netplay/listen_socket.cpp, listen_socket.h: class IListenSocket (cluster J's Contract C), PublicNotCreatable:
' WzConnectionProvider.openListenSocket makes it. The derived class is WzNet's private TCPListenSocket, which this
' object holds. close_() (Close is a VB keyword) is C++'s `delete` of the listen socket (~TCPListenSocket
' closes it).
' enum class IPVersions: IPV4 1, IPV6 2 (supportedIpVersions' mask).
Option Explicit

Private connProvider_ As WzConnectionProvider
Private compressionProviderSet_ As Boolean      ' WzCompressionProvider* compressionProvider_ (the singleton)
Private pwm_ As PendingWritesManager
Private tcp_ As TCPListenSocket                 ' (VB) the derived object

' IListenSocket(WzConnectionProvider& connProvider, WzCompressionProvider& compressionProvider, PendingWritesManager& pwm)
Friend Sub construct(ByVal connProvider As WzConnectionProvider, ByVal pwm As PendingWritesManager, ByVal derived As TCPListenSocket)
    Set connProvider_ = connProvider
    compressionProviderSet_ = True
    Set pwm_ = pwm
    Set tcp_ = derived
End Sub

' virtual IClientConnection* accept() = 0 (Nothing: no connection waiting)
Public Function accept() As IClientConnection
    Set accept = tcp_.accept(connProvider_, pwm_)
End Function

' virtual IPVersionsMask supportedIpVersions() const = 0
Public Function supportedIpVersions() As Long
    supportedIpVersions = tcp_.supportedIpVersions()
End Function

' virtual ConnectionProviderType type() const noexcept = 0
Public Function type_() As Long
    type_ = tcp_.type_()
End Function

' (VB) delete listenSocket
Public Sub close_()
    If Not tcp_ Is Nothing Then tcp_.destroy_
End Sub
