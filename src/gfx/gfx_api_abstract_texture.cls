VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "gfx_api_abstract_texture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' gfx_api.h: gfx_api::abstract_texture. The textures of every backend implement
' it besides their own interface (gfx_api_texture, gfx_api_texture_array; VB
' interfaces do not inherit): a texture bound through bind_textures, or one the
' depth and scene passes render to (gl_gpurendered_texture), is one of these.
' Its methods are also in gfx_api_texture and gfx_api_texture_array, so that a
' texture is used as C++ uses it without a conversion. Methods have no
' underscores (Implements refuses them): backend_internal_value is
' backendInternalValue.
Option Explicit

Public Sub bind()
End Sub

Public Function isArray() As Boolean
End Function

Public Sub flush()
End Sub

' used for backend-specific internal purposes
Public Function backendInternalValue() As Long
End Function
