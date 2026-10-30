VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "gl_texture_array"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/ivis_opengl/gfx_api_gl.cpp: gl_texture_array (gfx_api_gl.h), a 2D texture array: the layers
' are copied into a client-side buffer (texture_array_mip_level_buffer) and uploaded on flush.
' gl_context (gfx_api_gl_c) sets its fields, as C++'s friend struct does; the destructor is
' Class_Terminate. The asserts that a size_t fits a GLint / GLsizei are left out (Longs).
Option Explicit

Implements gfx_api_texture_array
Implements gfx_api_abstract_texture

Private id_ As Long                     ' GLuint _id
Public mip_count As Long                ' = 0
Public layer_count As Long              ' = 0
Public internal_format As pixel_format  ' = gfx_api::pixel_format::invalid
Public gles As Boolean                  ' = false
Public pInternalBuffer As texture_array_mip_level_buffer ' = nullptr

' MARK: gl_texture_array

Private Sub Class_Initialize()
    glGenTextures 1, VarPtr(id_)
End Sub

Private Sub Class_Terminate()
    glDeleteTextures 1, VarPtr(id_)
    Set pInternalBuffer = Nothing
End Sub

Public Sub bind()
    glBindTexture GL_TEXTURE_2D_ARRAY, id_
End Sub

Public Sub unbind()
    glBindTexture GL_TEXTURE_2D_ARRAY, 0
End Sub

Public Function backend_internal_value() As Long
    ' not currently used in GL backend
    backend_internal_value = 0
End Function

Private Function upload_internal(ByVal layer As Long, ByVal mip_level As Long, ByVal offset_x As Long, ByVal offset_y As Long, ByVal image As iV_Image) As Boolean
    Dim width As Long, height As Long
    If Not (image.data() <> 0) Then ASSERT_FAILURE "image.data() != nullptr", "lib/ivis_opengl/gfx_api_gl.cpp:686", "upload_internal", "Attempt to upload image without data": upload_internal = False: Exit Function
    If Not (image.pixel_format_() = internal_format) Then ASSERT_FAILURE "image.pixel_format() == internal_format", "lib/ivis_opengl/gfx_api_gl.cpp:687", "upload_internal", "Uploading image to texture with different format": upload_internal = False: Exit Function
    If Not (mip_level < mip_count) Then ASSERT_FAILURE "mip_level < mip_count", "lib/ivis_opengl/gfx_api_gl.cpp:688", "upload_internal", "mip_level (%zu) >= mip_count (%zu)", mip_level, mip_count: upload_internal = False: Exit Function
    width = image.width()
    height = image.height()
    If Not (width > 0 And height > 0) Then ASSERT_FAILURE "width > 0 && height > 0", "lib/ivis_opengl/gfx_api_gl.cpp:691", "upload_internal", "Attempt to upload texture with width or height of 0 (width: %zu, height: %zu)", width, height
    bind
    If Not (gfx_api_format_memory_size(image.pixel_format_(), width, height) = image.data_size()) Then ASSERT_FAILURE "gfx_api::format_memory_size(image.pixel_format(), width, height) == image.data_size()", "lib/ivis_opengl/gfx_api_gl.cpp:701", "upload_internal", "data_size (%zu) does not match expected format_memory_size(%s, %zu, %zu)=%zu", image.data_size(), gfx_api_format_to_str(image.pixel_format_()), width, height, gfx_api_format_memory_size(image.pixel_format_(), width, height)

    ' Copy to an internal buffer for upload on flush
    If Not (offset_x = 0 And offset_y = 0) Then ASSERT_FAILURE "offset_x == 0 && offset_y == 0", "lib/ivis_opengl/gfx_api_gl.cpp:704", "upload_internal", "Trying to upload compressed sub texture": upload_internal = False: Exit Function
    pInternalBuffer.copy_data_to_buffer mip_level, layer, image.data(), image.data_size()

    unbind
    upload_internal = True
End Function

Public Function upload_layer(ByVal layer As Long, ByVal mip_level As Long, ByVal image As iV_Image) As Boolean
    upload_layer = upload_internal(layer, mip_level, 0, 0, image)
End Function

Public Function id() As Long
    id = id_
End Function

Public Sub flush()
    Dim i As Long, level_width As Long, level_height As Long, level_buffer_size As Long
    If Not pInternalBuffer Is Nothing Then
        ' If compressed texture, upload each mip level with glCompressedTexImage3D from the client-side buffer
        bind
        For i = 0 To mip_count - 1
            If Not pInternalBuffer.get_mip_level(i, level_width, level_height, level_buffer_size) Then
                GoTo continue_
            End If

            If gfx_api_is_uncompressed_format(internal_format) Then
                glTexImage3D GL_TEXTURE_2D_ARRAY, i, to_gl_internalformat(internal_format, gles), level_width, level_height, layer_count, 0, to_gl_format(internal_format, gles), GL_UNSIGNED_BYTE, pInternalBuffer.get_read_pointer(i)
            Else
                glCompressedTexImage3D GL_TEXTURE_2D_ARRAY, i, to_gl_internalformat(internal_format, gles), level_width, level_height, layer_count, 0, level_buffer_size, pInternalBuffer.get_read_pointer(i)
            End If
continue_:
        Next
        unbind
        pInternalBuffer.clear
    End If
End Sub

'--------------------------------------------------------------------------
' gfx_api::texture_array

Private Sub gfx_api_texture_array_bind()
    bind
End Sub

Private Function gfx_api_texture_array_isArray() As Boolean
    gfx_api_texture_array_isArray = True
End Function

Private Sub gfx_api_texture_array_flush()
    flush
End Sub

Private Function gfx_api_texture_array_backendInternalValue() As Long
    gfx_api_texture_array_backendInternalValue = backend_internal_value()
End Function

Private Function gfx_api_texture_array_uploadLayer(ByVal layer As Long, ByVal mip_level As Long, ByVal image As iV_Image) As Boolean
    gfx_api_texture_array_uploadLayer = upload_layer(layer, mip_level, image)
End Function

Private Function gfx_api_texture_array_id() As Long
    gfx_api_texture_array_id = id_
End Function

'--------------------------------------------------------------------------
' gfx_api::abstract_texture

Private Sub gfx_api_abstract_texture_bind()
    bind
End Sub

Private Function gfx_api_abstract_texture_isArray() As Boolean
    gfx_api_abstract_texture_isArray = True
End Function

Private Sub gfx_api_abstract_texture_flush()
    flush
End Sub

Private Function gfx_api_abstract_texture_backendInternalValue() As Long
    gfx_api_abstract_texture_backendInternalValue = backend_internal_value()
End Function
