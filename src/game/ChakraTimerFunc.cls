VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChakraTimerFunc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/quickjs_backend.cpp: the TimerFunc lambda of SetQuickJSTimer and of
' quickjs_scripting_instance::restoreTimerFunction ([ctx, funcName]
' (uniqueTimerID, BASE_OBJECT*, timerAdditionalData*)): calls funcName in the
' script with the timer's object or its string argument
' (chakra_backend_c's runQuickJSTimerFunction).
Option Explicit

Implements ITimerFunc

Private ctx As Long                     ' JSContext *: the captured context
Private funcName As String

Friend Sub init(ByVal ctx_ As Long, ByVal funcName_ As String)
    ctx = ctx_
    funcName = funcName_
End Sub

Private Sub ITimerFunc_invoke(ByVal timerID As Double, ByVal baseObject As Long, ByVal additionalData As Object)
    runQuickJSTimerFunction ctx, funcName, baseObject, additionalData
End Sub
