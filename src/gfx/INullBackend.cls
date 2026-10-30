VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "INullBackend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' lib/ivis_opengl/gfx_api_null.h: gfx_api::backend_Null_Impl, the window
' system's side of the null (headless) backend. WzPlatform implements it
' (win_Null_Impl, the port of lib/sdl/gfx_api_null_sdl.cpp) and hands it to
' gfx_api_context_initialize with backend_type_null_backend.
Option Explicit

Public Sub swapWindow()
End Sub

Public Function setSwapInterval(ByVal mode As swap_interval_mode) As Boolean
End Function

Public Function getSwapInterval() As swap_interval_mode
End Function
