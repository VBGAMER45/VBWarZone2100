VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "wzapi_execution_context"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/wzapi.h, wzapi.cpp: class wzapi::execution_context, and
' src/quickjs_backend.cpp's quickjs_execution_context (the one
' implementation): what an API function knows of its caller. Made by
' wzapi_execution_context_new(ctx) (chakra_backend_c) for each call of a
' native function.
' throwError: QuickJS's JS_ThrowReferenceError("%s failed in %s at line %d")
' followed by the API function's normal return, which QuickJS never raises
' (see chakra_backend_c's header): the message is kept (lastError), nothing
' is thrown.
Option Explicit

Private ctx As Long                     ' JSContext *
Private m_lastError As String

Friend Sub construct(ByVal ctx_ As Long)
    ctx = ctx_
End Sub

Public Function currentInstance() As ChakraScriptingInstance
    Set currentInstance = engineToInstanceMap_at(ctx)
End Function

Public Function player() As Long
    player = currentInstance().player()
End Function

Public Sub set_isReceivingAllEvents(ByVal value As Boolean)
    currentInstance().setReceiveAllEvents value
End Sub

Public Function get_isReceivingAllEvents() As Boolean
    get_isReceivingAllEvents = currentInstance().isReceivingAllEvents()
End Function

' playerCallbackFunc: a ChakraPlayerCallback (Sub invoke(ByVal player As Long))
Public Function getNamedScriptCallback(ByVal func As String) As Object
    Dim cb As ChakraPlayerCallback
    Set cb = New ChakraPlayerCallback
    cb.init ctx, func
    Set getNamedScriptCallback = cb
End Function

Public Sub doNotSaveGlobal(ByVal global_ As String)
    engineToInstanceMap_at(ctx).doNotSaveGlobal global_
End Sub

Public Sub throwError(ByVal expr As String, ByVal line_ As Long, ByVal function_ As String)
    m_lastError = expr & " failed in " & function_ & " at line " & CStr(line_)
End Sub

' (VB) the last throwError's message ("" if none)
Public Function lastError() As String
    lastError = m_lastError
End Function
