VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "gl_texture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/ivis_opengl/gfx_api_gl.cpp: gl_texture (gfx_api_gl.h), a 2D texture. gl_context
' (gfx_api_gl_c) sets its fields, as C++'s friend struct does; the destructor is Class_Terminate.
' upload takes an iV_BaseImage (an iV_Image or an iV_CompressedImage), upload_sub an iV_Image. The asserts that a
' size_t fits a GLint / GLsizei are left out: every size is a Long already.
Option Explicit

Implements gfx_api_texture
Implements gfx_api_abstract_texture

Private id_ As Long                     ' GLuint _id
Public mip_count As Long                ' = 0
Public internal_format As pixel_format  ' = gfx_api::pixel_format::invalid
Public gles As Boolean                  ' = false
Public tex_width As Long                ' = 0
Public tex_height As Long               ' = 0

' MARK: gl_texture

Private Sub Class_Initialize()
    glGenTextures 1, VarPtr(id_)
End Sub

Private Sub Class_Terminate()
    glDeleteTextures 1, VarPtr(id_)
End Sub

Public Sub bind()
    glBindTexture GL_TEXTURE_2D, id_
End Sub

Public Function backend_internal_value() As Long
    ' not currently used in GL backend
    backend_internal_value = 0
End Function

Public Sub unbind()
    glBindTexture GL_TEXTURE_2D, 0
End Sub

Private Function upload_internal(ByVal mip_level As Long, ByVal offset_x As Long, ByVal offset_y As Long, ByVal image As iV_BaseImage) As Boolean
    Dim width As Long, height As Long, glFormat As Long
    If Not (image.data() <> 0) Then ASSERT_FAILURE "image.data() != nullptr", "lib/ivis_opengl/gfx_api_gl.cpp:513", "upload_internal", "Attempt to upload image without data": upload_internal = False: Exit Function
    If Not (image.pixelFormat() = internal_format) Then ASSERT_FAILURE "image.pixel_format() == internal_format", "lib/ivis_opengl/gfx_api_gl.cpp:514", "upload_internal", "Uploading image to texture with different format": upload_internal = False: Exit Function
    width = image.width()
    height = image.height()
    If Not (width > 0 And height > 0) Then ASSERT_FAILURE "width > 0 && height > 0", "lib/ivis_opengl/gfx_api_gl.cpp:517", "upload_internal", "Attempt to upload texture with width or height of 0 (width: %zu, height: %zu)", width, height
    bind
    If Not (gfx_api_format_memory_size(image.pixelFormat(), width, height) = image.dataSize()) Then ASSERT_FAILURE "gfx_api::format_memory_size(image.pixel_format(), width, height) == image.data_size()", "lib/ivis_opengl/gfx_api_gl.cpp:526", "upload_internal", "data_size (%zu) does not match expected format_memory_size(%s, %zu, %zu)=%zu", image.dataSize(), gfx_api_format_to_str(image.pixelFormat()), width, height, gfx_api_format_memory_size(image.pixelFormat(), width, height)
    If gfx_api_is_uncompressed_format(image.pixelFormat()) Then
        glTexSubImage2D GL_TEXTURE_2D, mip_level, offset_x, offset_y, width, height, to_gl_format(image.pixelFormat(), gles), GL_UNSIGNED_BYTE, image.data()
    Else
        If Not (offset_x = 0 And offset_y = 0) Then ASSERT_FAILURE "offset_x == 0 && offset_y == 0", "lib/ivis_opengl/gfx_api_gl.cpp:533", "upload_internal", "Trying to upload compressed sub texture": upload_internal = False: Exit Function
        glFormat = to_gl_internalformat(image.pixelFormat(), gles)
        glCompressedTexImage2D GL_TEXTURE_2D, mip_level, glFormat, width, height, 0, image.dataSize(), image.data()
    End If
    unbind
    upload_internal = True
End Function

Public Function upload(ByVal mip_level As Long, ByVal image As iV_BaseImage) As Boolean
    upload = upload_internal(mip_level, 0, 0, image)
End Function

Public Function upload_sub(ByVal mip_level As Long, ByVal offset_x As Long, ByVal offset_y As Long, ByVal image As iV_Image) As Boolean
    upload_sub = upload_internal(mip_level, offset_x, offset_y, image)
End Function

Public Function id() As Long
    id = id_
End Function

' texture2dDimensions get_dimensions() const
Public Sub get_dimensions(width As Long, height As Long)
    width = tex_width
    height = tex_height
End Sub

'--------------------------------------------------------------------------
' gfx_api::texture

Private Sub gfx_api_texture_bind()
    bind
End Sub

Private Function gfx_api_texture_isArray() As Boolean
    gfx_api_texture_isArray = False
End Function

Private Sub gfx_api_texture_flush()
End Sub

Private Function gfx_api_texture_backendInternalValue() As Long
    gfx_api_texture_backendInternalValue = backend_internal_value()
End Function

Private Function gfx_api_texture_upload(ByVal mip_level As Long, ByVal image As iV_BaseImage) As Boolean
    gfx_api_texture_upload = upload(mip_level, image)
End Function

Private Function gfx_api_texture_uploadSub(ByVal mip_level As Long, ByVal offset_x As Long, ByVal offset_y As Long, ByVal image As iV_Image) As Boolean
    gfx_api_texture_uploadSub = upload_sub(mip_level, offset_x, offset_y, image)
End Function

Private Function gfx_api_texture_id() As Long
    gfx_api_texture_id = id_
End Function

Private Sub gfx_api_texture_getDimensions(width As Long, height As Long)
    get_dimensions width, height
End Sub

'--------------------------------------------------------------------------
' gfx_api::abstract_texture

Private Sub gfx_api_abstract_texture_bind()
    bind
End Sub

Private Function gfx_api_abstract_texture_isArray() As Boolean
    gfx_api_abstract_texture_isArray = False
End Function

Private Sub gfx_api_abstract_texture_flush()
End Sub

Private Function gfx_api_abstract_texture_backendInternalValue() As Long
    gfx_api_abstract_texture_backendInternalValue = backend_internal_value()
End Function
