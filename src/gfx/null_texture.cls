VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "null_texture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/ivis_opengl/gfx_api_null.cpp, gfx_api_null.h: null_texture, the null
' (headless) backend's 2D texture: it keeps its format and size and uploads
' nothing. gfx_api::texture's isArray() is false; abstract_texture's flush()
' does nothing.
Option Explicit

Implements gfx_api_texture
Implements gfx_api_abstract_texture

' protected (null_context sets them: friend struct null_context)
Public internal_format As pixel_format  ' = gfx_api::pixel_format::invalid
Public tex_width As Long                ' = 0
Public tex_height As Long               ' = 0

' null_texture::null_texture(), ~null_texture(): no-op

Private Sub gfx_api_texture_bind()
    ' no-op
End Sub

Private Function gfx_api_texture_isArray() As Boolean
    gfx_api_texture_isArray = False
End Function

Private Sub gfx_api_texture_flush()
End Sub

Private Function gfx_api_texture_backendInternalValue() As Long
    gfx_api_texture_backendInternalValue = 0
End Function

Private Function gfx_api_texture_upload(ByVal mip_level As Long, ByVal image As iV_Image) As Boolean
    Dim width As Long, height As Long
    If Not (image.data() <> 0) Then ASSERT_FAILURE "image.data() != nullptr", "lib/ivis_opengl/gfx_api_null.cpp:43", "upload", "Attempt to upload image without data": gfx_api_texture_upload = False: Exit Function
    If Not (image.pixel_format_() = internal_format) Then ASSERT_FAILURE "image.pixel_format() == internal_format", "lib/ivis_opengl/gfx_api_null.cpp:44", "upload", "Uploading image to texture with different format": gfx_api_texture_upload = False: Exit Function
    width = image.width()
    height = image.height()
    If Not (width > 0 And height > 0) Then ASSERT_FAILURE "width > 0 && height > 0", "lib/ivis_opengl/gfx_api_null.cpp:47", "upload", "Attempt to upload texture with width or height of 0 (width: %zu, height: %zu)", width, height
    ' no-op
    gfx_api_texture_upload = True
End Function

Private Function gfx_api_texture_uploadSub(ByVal mip_level As Long, ByVal offset_x As Long, ByVal offset_y As Long, ByVal image As iV_Image) As Boolean
    Dim width As Long, height As Long
    If Not (image.data() <> 0) Then ASSERT_FAILURE "image.data() != nullptr", "lib/ivis_opengl/gfx_api_null.cpp:54", "upload_sub", "Attempt to upload image without data": gfx_api_texture_uploadSub = False: Exit Function
    If Not (image.pixel_format_() = internal_format) Then ASSERT_FAILURE "image.pixel_format() == internal_format", "lib/ivis_opengl/gfx_api_null.cpp:55", "upload_sub", "Uploading image to texture with different format": gfx_api_texture_uploadSub = False: Exit Function
    width = image.width()
    height = image.height()
    If Not (width > 0 And height > 0) Then ASSERT_FAILURE "width > 0 && height > 0", "lib/ivis_opengl/gfx_api_null.cpp:58", "upload_sub", "Attempt to upload texture with width or height of 0 (width: %zu, height: %zu)", width, height
    ' no-op
    gfx_api_texture_uploadSub = True
End Function

Private Function gfx_api_texture_id() As Long
    gfx_api_texture_id = 0
End Function

Private Sub gfx_api_texture_getDimensions(width As Long, height As Long)
    width = tex_width
    height = tex_height
End Sub

' (gfx_api::abstract_texture: the same methods)
Private Sub gfx_api_abstract_texture_bind()
    gfx_api_texture_bind
End Sub

Private Function gfx_api_abstract_texture_isArray() As Boolean
    gfx_api_abstract_texture_isArray = gfx_api_texture_isArray()
End Function

Private Sub gfx_api_abstract_texture_flush()
    gfx_api_texture_flush
End Sub

Private Function gfx_api_abstract_texture_backendInternalValue() As Long
    gfx_api_abstract_texture_backendInternalValue = gfx_api_texture_backendInternalValue()
End Function
