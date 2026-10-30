VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "gl_gpurendered_texture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/ivis_opengl/gfx_api_gl.cpp: gl_gpurendered_texture (gfx_api_gl.h), a texture the GPU renders
' to (the depth map, the scene texture, the default depth texture). gl_context (gfx_api_gl_c)
' sets its fields, as C++'s friend struct does; the destructor is Class_Terminate.
Option Explicit

Implements gfx_api_abstract_texture

Private id_ As Long                     ' GLuint _id
Public gles As Boolean                  ' = false
Public isArray_ As Boolean              ' _isArray = false

' MARK: gl_gpurendered_texture

Private Sub Class_Initialize()
    glGenTextures 1, VarPtr(id_)
End Sub

Private Sub Class_Terminate()
    glDeleteTextures 1, VarPtr(id_)
End Sub

Public Sub bind()
    If isArray_ Then
        glBindTexture GL_TEXTURE_2D_ARRAY, id_
    Else
        glBindTexture GL_TEXTURE_2D, id_
    End If
End Sub

Public Function isArray() As Boolean
    isArray = isArray_
End Function

Public Function target() As Long
    If isArray_ Then target = GL_TEXTURE_2D_ARRAY Else target = GL_TEXTURE_2D
End Function

Public Function id() As Long
    id = id_
End Function

Public Function backend_internal_value() As Long
    ' not currently used in GL backend
    backend_internal_value = 0
End Function

Public Sub unbind()
    If isArray_ Then
        glBindTexture GL_TEXTURE_2D_ARRAY, 0
    Else
        glBindTexture GL_TEXTURE_2D, 0
    End If
End Sub

'--------------------------------------------------------------------------
' gfx_api::abstract_texture

Private Sub gfx_api_abstract_texture_bind()
    bind
End Sub

Private Function gfx_api_abstract_texture_isArray() As Boolean
    gfx_api_abstract_texture_isArray = isArray_
End Function

Private Sub gfx_api_abstract_texture_flush()
End Sub

Private Function gfx_api_abstract_texture_backendInternalValue() As Long
    gfx_api_abstract_texture_backendInternalValue = backend_internal_value()
End Function
