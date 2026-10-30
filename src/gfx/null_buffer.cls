VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "null_buffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/ivis_opengl/gfx_api_null.cpp, gfx_api_null.h: null_buffer, the null
' (headless) backend's buffer: it keeps its size and stores nothing. The
' constructor null_buffer(usage, hint) is gfx_api_null_c's null_buffer_new.
' The once-per-frame asserts are `#if defined(DEBUG)` (not compiled).
Option Explicit

Implements gfx_api_buffer

Public usage As buffer_usage
Public hint As buffer_storage_hint
Public buffer_size As Long              ' = 0
Public lastUploaded_FrameNum As Long    ' = 0

' null_buffer::~null_buffer(): no-op

Private Sub gfx_api_buffer_bind()
    ' no-op
End Sub

Private Sub gfx_api_buffer_upload(ByVal size As Long, ByVal data As Long)
    Dim current_FrameNum As Long
    current_FrameNum = gfx_api_context_current_FrameNum()
    lastUploaded_FrameNum = current_FrameNum

    If Not (size > 0) Then ASSERT_FAILURE "size > 0", "lib/ivis_opengl/gfx_api_null.cpp:143", "upload", "Attempt to upload buffer of size 0"
    buffer_size = size
    ' no-op
End Sub

Private Sub gfx_api_buffer_update(ByVal start As Long, ByVal size As Long, ByVal data As Long, Optional ByVal flag As update_flag = update_flag_none)
    Dim current_FrameNum As Long
    current_FrameNum = gfx_api_context_current_FrameNum()
    lastUploaded_FrameNum = current_FrameNum

    If Not (start < buffer_size) Then ASSERT_FAILURE "start < buffer_size", "lib/ivis_opengl/gfx_api_null.cpp:156", "update", "Starting offset (%zu) is past end of buffer (length: %zu)", start, buffer_size
    If Not (start + size <= buffer_size) Then ASSERT_FAILURE "start + size <= buffer_size", "lib/ivis_opengl/gfx_api_null.cpp:157", "update", "Attempt to write past end of buffer"
    If size = 0 Then
        debug_ 160, LOG_WARNING, "update", "Attempt to update buffer with 0 bytes of new data"
        Exit Sub
    End If
    ' no-op
End Sub

Private Function gfx_api_buffer_currentBufferSize() As Long
    gfx_api_buffer_currentBufferSize = buffer_size
End Function
