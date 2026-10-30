VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "texture_array_mip_level_buffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/ivis_opengl/gfx_api_gl.cpp: texture_array_mip_level_buffer, the client-side copy of a
' texture array's layers until gl_texture_array::flush uploads them. The
' std::unordered_map<size_t, MipLevel> is an array over the mip levels with a "present" flag; a
' MipLevel's fields are read with get_mip_level (a Boolean, the fields ByRef) and
' get_read_pointer.
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

Private Type MipLevel
    present As Boolean                  ' (VB) the map holds this level
    buffer() As Byte                    ' std::vector<uint8_t>
    buffer_size As Long
    width As Long                       ' = 0
    height As Long                      ' = 0
    memorySizePerLayer As Long          ' = 0
End Type

Private mip_levels As Long
Private layer_count As Long
Private width As Long
Private height As Long
Private internal_format As pixel_format
Private mipLevelDataBuffer() As MipLevel

Private Function getMipLevel(ByVal mip_level As Long) As Long
    Dim mipLevelBufferSize As Long
    If mipLevelDataBuffer(mip_level).present Then
        getMipLevel = mip_level
        Exit Function
    End If
    mipLevelDataBuffer(mip_level).present = True
    ' std::max<size_t>(1, width >> mip_level)
    mipLevelDataBuffer(mip_level).width = SHRU(width, mip_level)
    If mipLevelDataBuffer(mip_level).width < 1 Then mipLevelDataBuffer(mip_level).width = 1
    mipLevelDataBuffer(mip_level).height = SHRU(height, mip_level)
    If mipLevelDataBuffer(mip_level).height < 1 Then mipLevelDataBuffer(mip_level).height = 1
    mipLevelDataBuffer(mip_level).memorySizePerLayer = gfx_api_format_memory_size(internal_format, mipLevelDataBuffer(mip_level).width, mipLevelDataBuffer(mip_level).height)
    mipLevelBufferSize = mipLevelDataBuffer(mip_level).memorySizePerLayer * layer_count
    If mipLevelBufferSize > 0 Then ReDim mipLevelDataBuffer(mip_level).buffer(0 To mipLevelBufferSize - 1)
    mipLevelDataBuffer(mip_level).buffer_size = mipLevelBufferSize
    getMipLevel = mip_level
End Function

' texture_array_mip_level_buffer(const size_t& _mipmap_count, const size_t& _layer_count, const size_t& _width, const size_t& _height, const gfx_api::pixel_format& _internal_format)
Friend Sub construct(ByVal mipmap_count_ As Long, ByVal layer_count_ As Long, ByVal width_ As Long, ByVal height_ As Long, ByVal internal_format_ As pixel_format)
    mip_levels = mipmap_count_
    layer_count = layer_count_
    width = width_
    height = height_
    internal_format = internal_format_
    If mip_levels > 0 Then ReDim mipLevelDataBuffer(0 To mip_levels - 1)
End Sub

Public Function copy_data_to_buffer(ByVal mip_level As Long, ByVal layer As Long, ByVal data As Long, ByVal data_size As Long) As Boolean
    Dim level As Long, startingBufferOffset As Long
    If Not (mip_level < mip_levels) Then ASSERT_FAILURE "mip_level < mip_levels", "lib/ivis_opengl/gfx_api_gl.cpp:610", "copy_data_to_buffer", "Invalid mip_level (%zu)", mip_level: copy_data_to_buffer = False: Exit Function
    If Not (layer < layer_count) Then ASSERT_FAILURE "layer < layer_count", "lib/ivis_opengl/gfx_api_gl.cpp:611", "copy_data_to_buffer", "Invalid layer (%zu)", layer: copy_data_to_buffer = False: Exit Function
    level = getMipLevel(mip_level)
    If Not (data_size = mipLevelDataBuffer(level).memorySizePerLayer) Then ASSERT_FAILURE "data_size == level.memorySizePerLayer", "lib/ivis_opengl/gfx_api_gl.cpp:613", "copy_data_to_buffer", "Invalid data_size (%zu, expecting: %zu)", data_size, mipLevelDataBuffer(level).memorySizePerLayer: copy_data_to_buffer = False: Exit Function
    startingBufferOffset = mipLevelDataBuffer(level).memorySizePerLayer * layer
    If data_size > 0 Then CopyMemory mipLevelDataBuffer(level).buffer(startingBufferOffset), ByVal data, data_size
    copy_data_to_buffer = True
End Function

' const MipLevel* get_mip_level(size_t mip_level) const: False = nullptr
Public Function get_mip_level(ByVal mip_level As Long, level_width As Long, level_height As Long, level_buffer_size As Long) As Boolean
    If mip_level < 0 Or mip_level >= mip_levels Then Exit Function
    If Not mipLevelDataBuffer(mip_level).present Then Exit Function
    level_width = mipLevelDataBuffer(mip_level).width
    level_height = mipLevelDataBuffer(mip_level).height
    level_buffer_size = mipLevelDataBuffer(mip_level).buffer_size
    get_mip_level = True
End Function

' 0 = nullptr
Public Function get_read_pointer(ByVal mip_level As Long) As Long
    If mip_level < 0 Or mip_level >= mip_levels Then Exit Function
    If Not mipLevelDataBuffer(mip_level).present Then Exit Function
    If mipLevelDataBuffer(mip_level).buffer_size = 0 Then Exit Function
    get_read_pointer = VarPtr(mipLevelDataBuffer(mip_level).buffer(0))
End Function

Public Function get_buffer_size(ByVal mip_level As Long, ByVal layer As Long) As Long
    If mip_level < 0 Or mip_level >= mip_levels Then Exit Function
    If Not mipLevelDataBuffer(mip_level).present Then Exit Function
    get_buffer_size = mipLevelDataBuffer(mip_level).memorySizePerLayer
End Function

Public Sub clear()
    If mip_levels > 0 Then ReDim mipLevelDataBuffer(0 To mip_levels - 1)
End Sub
