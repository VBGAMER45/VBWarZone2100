VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IOpenConnectionCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' lib/netplay/open_connection_result.h: typedef std::function<void(OpenConnectionResult&& result)>
' OpenConnectionToHostResultCallback, as a callback interface (cluster J's Contract C): the caller of
' WzConnectionProvider.openClientConnectionAsync implements it (Implements IOpenConnectionCallback) and passes the
' object, which is called once, before openClientConnectionAsync returns.
Option Explicit

' void operator()(OpenConnectionResult&& result)
Public Sub onResult(ByVal result As OpenConnectionResult)
End Sub
