VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "gl_buffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/ivis_opengl/gfx_api_gl.cpp: gl_buffer (gfx_api_gl.h), a vertex or index buffer. The
' constructor gl_buffer(usage, hint) is gfx_api_gl_c's gl_buffer_new (Friend construct); the
' destructor is Class_Terminate. The data is an address (0 = NULL).
Option Explicit

Implements gfx_api_buffer

Public usage As buffer_usage
Public hint As buffer_storage_hint
Public buffer As Long                   ' GLuint buffer = 0
Public buffer_size As Long              ' = 0
Public lastUploaded_FrameNum As Long    ' = 0

' MARK: gl_buffer

Friend Sub construct(ByVal usage_ As buffer_usage, ByVal hint_ As buffer_storage_hint)
    usage = usage_
    hint = hint_
    glGenBuffers 1, VarPtr(buffer)
End Sub

Private Sub Class_Terminate()
    glDeleteBuffers 1, VarPtr(buffer)
End Sub

Public Sub bind()
    glBindBuffer to_gl_usage(usage), buffer
End Sub

Public Sub unbind()
    glBindBuffer to_gl_usage(usage), 0
End Sub

Public Sub upload(ByVal size As Long, ByVal data As Long)
    Dim current_FrameNum As Long
    current_FrameNum = gfx_api_context_current_FrameNum()
    lastUploaded_FrameNum = current_FrameNum

    If Not (size > 0) Then ASSERT_FAILURE "size > 0", "lib/ivis_opengl/gfx_api_gl.cpp:781", "upload", "Attempt to upload buffer of size 0"
    glBindBuffer to_gl_usage(usage), buffer
    glBufferData to_gl_usage(usage), size, data, to_gl_hint(hint)
    buffer_size = size
    glBindBuffer to_gl_usage(usage), 0
End Sub

Public Sub update(ByVal start As Long, ByVal size As Long, ByVal data As Long, Optional ByVal flag As update_flag = update_flag_none)
    Dim current_FrameNum As Long
    current_FrameNum = gfx_api_context_current_FrameNum()
    lastUploaded_FrameNum = current_FrameNum

    If Not (start < buffer_size) Then ASSERT_FAILURE "start < buffer_size", "lib/ivis_opengl/gfx_api_gl.cpp:796", "update", "Starting offset (%zu) is past end of buffer (length: %zu)", start, buffer_size
    If Not (start + size <= buffer_size) Then ASSERT_FAILURE "start + size <= buffer_size", "lib/ivis_opengl/gfx_api_gl.cpp:797", "update", "Attempt to write past end of buffer"
    If size = 0 Then
        debug_ 800, LOG_WARNING, "update", "Attempt to update buffer with 0 bytes of new data"
        Exit Sub
    End If
    glBindBuffer to_gl_usage(usage), buffer
    glBufferSubData to_gl_usage(usage), start, size, data
    glBindBuffer to_gl_usage(usage), 0
End Sub

Public Function current_buffer_size() As Long
    current_buffer_size = buffer_size
End Function

'--------------------------------------------------------------------------
' gfx_api::buffer

Private Sub gfx_api_buffer_upload(ByVal size As Long, ByVal data As Long)
    upload size, data
End Sub

Private Sub gfx_api_buffer_update(ByVal start As Long, ByVal size As Long, ByVal data As Long, Optional ByVal flag As update_flag = update_flag_none)
    update start, size, data, flag
End Sub

Private Function gfx_api_buffer_currentBufferSize() As Long
    gfx_api_buffer_currentBufferSize = buffer_size
End Function

Private Sub gfx_api_buffer_bind()
    bind
End Sub
