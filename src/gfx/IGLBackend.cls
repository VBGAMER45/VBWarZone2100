VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IGLBackend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' lib/ivis_opengl/gfx_api_gl.h: gfx_api::backend_OpenGL_Impl, the window
' system's side of the OpenGL backend. WzPlatform implements it
' (win_OpenGL_Impl, the port of lib/sdl/gfx_api_gl_sdl.cpp) and hands it to
' gfx_api_context_initialize.
Option Explicit

' C: GLADloadproc getGLGetProcAddress(): returns the loader; here the loader itself:
' the address of the GL function name (0 if there is none).
Public Function getGLGetProcAddress(ByVal procName As String) As Long
End Function

' Creates an OpenGL context (double-buffered)
Public Function createGLContext() As Boolean
End Function

Public Function destroyGLContext() As Boolean
End Function

Public Sub swapWindow()
End Sub

' The size of the window's underlying drawable in pixels (for use with glViewport).
Public Sub getDrawableSize(w As Long, h As Long)
End Sub

Public Function isOpenGLES() As Boolean
End Function

Public Function setSwapInterval(ByVal mode As swap_interval_mode) As Boolean
End Function

Public Function getSwapInterval() As swap_interval_mode
End Function
