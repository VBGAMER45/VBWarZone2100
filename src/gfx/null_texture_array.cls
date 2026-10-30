VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "null_texture_array"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/ivis_opengl/gfx_api_null.cpp, gfx_api_null.h: null_texture_array, the
' null (headless) backend's texture array. gfx_api::texture_array's isArray()
' is true; abstract_texture's flush() does nothing.
Option Explicit

Implements gfx_api_texture_array
Implements gfx_api_abstract_texture

' protected (null_context sets it: friend struct null_context)
Public internal_format As pixel_format  ' = gfx_api::pixel_format::invalid

' null_texture_array::null_texture_array(), ~null_texture_array(): no-op

Private Sub gfx_api_texture_array_bind()
    ' no-op
End Sub

Private Function gfx_api_texture_array_isArray() As Boolean
    gfx_api_texture_array_isArray = True
End Function

Private Sub gfx_api_texture_array_flush()
End Sub

Private Function gfx_api_texture_array_backendInternalValue() As Long
    gfx_api_texture_array_backendInternalValue = 0
End Function

Private Function gfx_api_texture_array_uploadLayer(ByVal layer As Long, ByVal mip_level As Long, ByVal image As iV_Image) As Boolean
    Dim width As Long, height As Long
    If Not (image.data() <> 0) Then ASSERT_FAILURE "image.data() != nullptr", "lib/ivis_opengl/gfx_api_null.cpp:97", "upload_layer", "Attempt to upload image without data": gfx_api_texture_array_uploadLayer = False: Exit Function
    If Not (image.pixel_format_() = internal_format) Then ASSERT_FAILURE "image.pixel_format() == internal_format", "lib/ivis_opengl/gfx_api_null.cpp:98", "upload_layer", "Uploading image to texture with different format": gfx_api_texture_array_uploadLayer = False: Exit Function
    width = image.width()
    height = image.height()
    If Not (width > 0 And height > 0) Then ASSERT_FAILURE "width > 0 && height > 0", "lib/ivis_opengl/gfx_api_null.cpp:101", "upload_layer", "Attempt to upload texture with width or height of 0 (width: %zu, height: %zu)", width, height
    ' no-op
    gfx_api_texture_array_uploadLayer = True
End Function

Private Function gfx_api_texture_array_id() As Long
    gfx_api_texture_array_id = 0
End Function

' (gfx_api::abstract_texture: the same methods)
Private Sub gfx_api_abstract_texture_bind()
    gfx_api_texture_array_bind
End Sub

Private Function gfx_api_abstract_texture_isArray() As Boolean
    gfx_api_abstract_texture_isArray = gfx_api_texture_array_isArray()
End Function

Private Sub gfx_api_abstract_texture_flush()
    gfx_api_texture_array_flush
End Sub

Private Function gfx_api_abstract_texture_backendInternalValue() As Long
    gfx_api_abstract_texture_backendInternalValue = gfx_api_texture_array_backendInternalValue()
End Function
