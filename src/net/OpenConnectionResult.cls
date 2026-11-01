VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OpenConnectionResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' lib/netplay/open_connection_result.h: struct OpenConnectionResult (cluster J's Contract C), PublicNotCreatable:
' WzConnectionProvider.openClientConnectionAsync hands it to the callback. Its std::unique_ptr<IClientConnection,
' ConnectionCloser> open_socket is open_socket() (get()) and open_socket_release() (release(): the caller owns the
' connection then); the last reference to this object closes a connection not released (ConnectionCloser).
' optional<std::error_code> errorCode is hasError() and errorCode() (the ec, error_categories_c).
Option Explicit

Private open_socket_ As IClientConnection
Private hasErrorCode_ As Boolean
Private errorCode_ As Long
Private errorString_ As String

' OpenConnectionResult(std::error_code ec, std::string errorString)
Friend Sub construct_error(ByVal ec As Long, ByVal errorString As String)
    hasErrorCode_ = True
    errorCode_ = ec
    errorString_ = errorString
End Sub

' OpenConnectionResult(IClientConnection* open_socket)
Friend Sub construct_socket(ByVal open_socket As IClientConnection)
    Set open_socket_ = open_socket
End Sub

' bool hasError() const
Public Function hasError() As Boolean
    hasError = hasErrorCode_
End Function

' open_socket.get()
Public Function open_socket() As IClientConnection
    Set open_socket = open_socket_
End Function

' open_socket.release()
Public Function open_socket_release() As IClientConnection
    Set open_socket_release = open_socket_
    Set open_socket_ = Nothing
End Function

' errorCode.value() (0 when there is none)
Public Function errorCode() As Long
    If hasErrorCode_ Then errorCode = errorCode_
End Function

' std::string errorString
Public Function errorString() As String
    errorString = errorString_
End Function

' struct ConnectionCloser: the unique_ptr's deleter
Private Sub Class_Terminate()
    If Not open_socket_ Is Nothing Then
        open_socket_.close_
        Set open_socket_ = Nothing
    End If
End Sub
