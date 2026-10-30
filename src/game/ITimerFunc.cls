VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ITimerFunc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' wzapi.h: typedef std::function<void (uniqueTimerID, BASE_OBJECT*,
' timerAdditionalData *)> TimerFunc, as an interface: the script backend's
' timer functions implement it (the lambdas SetQuickJSTimer and
' restoreTimerFunction make). uniqueTimerID (uint64_t) is a Double; a
' timerAdditionalData * is an Object (Nothing = nullptr); a BASE_OBJECT * a
' handle.
Option Explicit

Public Sub invoke(ByVal timerID As Double, ByVal baseObject As Long, ByVal additionalData As Object)
End Sub
