VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChakraPlayerCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/quickjs_backend.cpp: the playerCallbackFunc lambda of
' quickjs_execution_context::getNamedScriptCallback ([pCtx, func](const int
' player)): calls the script's function func with the player.
Option Explicit

Private pCtx As Long                    ' JSContext *
Private func As String

Friend Sub init(ByVal ctx_ As Long, ByVal func_ As String)
    pCtx = ctx_
    func = func_
End Sub

Public Sub invoke(ByVal player As Long)
    Dim args() As Long, n As Long, prev As Long
    prev = jsEnterContext(pCtx)
    jsArgsPush args, n, jsInt(player)
    callFunction pCtx, func, args, n
    jsArgsRelease args, n
    jsLeaveContext prev
End Sub
