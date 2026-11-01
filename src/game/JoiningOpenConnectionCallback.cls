VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JoiningOpenConnectionCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/screens/joiningscreen.cpp: attemptToOpenConnection's callback lambda ([weakSelf, connectionIdx](OpenConnectionResult&&
' result)) as J_DESIGN's Contract C's IOpenConnectionCallback, which the exe implements (joiningscreen_c).
Option Explicit

Implements IOpenConnectionCallback

Private weakSelf_ As Long               ' std::weak_ptr<WzJoiningGameScreen_HandlerRoot>: the widgweak_vb serial
Private connectionIdx_ As Long          ' size_t connectionIdx

Public Sub init(ByVal weakSelf As Long, ByVal connectionIdx As Long)
    weakSelf_ = weakSelf
    connectionIdx_ = connectionIdx
End Sub

Private Sub IOpenConnectionCallback_onResult(ByVal result As OpenConnectionResult)
    Dim strongSelf As WIDGET
    Set strongSelf = WeakWidget_lock(weakSelf_)
    If strongSelf Is Nothing Then
        ' background thread ultimately returned after the requester has gone away (join was cancelled?) - just return
        Exit Sub
    End If
    WzJoiningGameScreen_HandlerRoot_processOpenConnectionResultOnMainThread strongSelf, connectionIdx_, result
End Sub
