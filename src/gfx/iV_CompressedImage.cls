VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iV_CompressedImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/ivis_opengl/gfx_api_image_compress_priv.cpp, gfx_api_image_compress_priv.h: iV_CompressedImage, an image
' in a compressed format (implements iV_BaseImage). The data are a Byte array (C++: a uint64_t buffer, its size
' rounded up to whole uint64_ts); data() / uint64_w() give its address (0 if none). The transcoder writes the
' blocks into the array itself: swapData lends it (gfx_api_image_basis_priv_c). C++'s move constructor and
' assignment are moveFrom; the destructor is Class_Terminate.
Option Explicit

Implements iV_BaseImage

Private m_width As Long                 ' unsigned int = 0
Private m_height As Long                ' unsigned int = 0
Private m_bufferRowLength As Long       ' unsigned int = 0
Private m_bufferImageHeight As Long     ' unsigned int = 0
Private m_format As pixel_format        ' = gfx_api::pixel_format::invalid
Private m_data() As Byte                ' uint64_t *m_data = nullptr
Private m_hasData As Boolean            ' (VB) m_data != nullptr
Private m_data_size As Long             ' size_t = 0

Private Sub Class_Initialize()
    m_format = pixel_format_invalid
End Sub

Private Sub Class_Terminate()
    clear
End Sub

' (VB) m_data = (uint64_t*)malloc(n): buf is a new buffer of n bytes
Private Function NewBuffer(buf() As Byte, ByVal n As Long) As Boolean
    On Error GoTo fail
    ReDim buf(0 To n - 1)
    NewBuffer = True
    Exit Function
fail:
End Function

' Allocate a new iV_CompressedImage buffer
Public Function allocate(ByVal format As pixel_format, ByVal data_size As Long, Optional ByVal bufferRowLength As Long = 0, Optional ByVal bufferImageHeight As Long = 0, Optional ByVal newWidth As Long = 0, Optional ByVal newHeight As Long = 0, Optional ByVal zeroMemory As Boolean = False) As Boolean
    Dim malloc_size As Long
    If m_hasData Then
        Erase m_data
        m_hasData = False
        m_data_size = 0
    End If
    If format <> pixel_format_invalid And data_size > 0 And newWidth > 0 And newHeight > 0 Then
        malloc_size = ((data_size + (8 - 1)) \ 8) * 8
        If Not NewBuffer(m_data, malloc_size) Then
            ASSERT_FAILURE "false", "lib/ivis_opengl/gfx_api_image_compress_priv.cpp:153", "allocate", "Failed to allocate memory buffer of size: %zu", data_size
            allocate = False
            Exit Function
        End If
        m_hasData = True
        ' (VB) a new VB array is zeroed, so zeroMemory needs no memset
    End If
    m_width = newWidth
    m_height = newHeight
    m_bufferRowLength = bufferRowLength
    m_bufferImageHeight = bufferImageHeight
    m_format = format
    m_data_size = data_size
    allocate = True
End Function

' Clear (free) the image contents / data
Public Sub clear()
    allocate pixel_format_invalid, 0, 0, 0, 0, 0
End Sub

Public Function width() As Long
    width = m_width
End Function

Public Function height() As Long
    height = m_height
End Function

' Get the current image pixel format
Public Function pixel_format_() As pixel_format
    pixel_format_ = m_format
End Function

' Get a pointer to the bitmap data that can be read
Public Function data() As Long
    If m_hasData Then data = VarPtr(m_data(0))
End Function

Public Function data_size() As Long
    data_size = m_data_size
End Function

Public Function bufferRowLength() As Long
    bufferRowLength = 0
End Function

Public Function bufferImageHeight() As Long
    bufferImageHeight = 0
End Function

' Get a pointer to the bitmap data that can be written to
Public Function uint64_w() As Long
    If m_hasData Then uint64_w = VarPtr(m_data(0))
End Function

' (VB) iV_CompressedImage& operator=(iV_CompressedImage&& other): takes other's data, leaves it empty
Public Sub moveFrom(ByVal other As iV_CompressedImage)
    If Not other Is Me Then
        clear
        other.swapMembers m_data, m_hasData, m_width, m_height, m_format, m_data_size
    End If
End Sub

' (VB) std::swap of the members the move swaps (m_width, m_height, m_format, m_data, m_data_size)
Friend Sub swapMembers(buf() As Byte, hasBuf As Boolean, w As Long, h As Long, format As pixel_format, size As Long)
    Dim t As Long, tb As Boolean, tf As pixel_format
    ByteArraySwap m_data, buf
    tb = m_hasData: m_hasData = hasBuf: hasBuf = tb
    t = m_width: m_width = w: w = t
    t = m_height: m_height = h: h = t
    tf = m_format: m_format = format: format = tf
    t = m_data_size: m_data_size = size: size = t
End Sub

' (VB) the buffer uint64_w() points at, lent as an array (swapped with buf; swap again to give it back)
Friend Sub swapData(buf() As Byte)
    ByteArraySwap m_data, buf
End Sub

'--------------------------------------------------------------------------
' iV_BaseImage

Private Function iV_BaseImage_width() As Long
    iV_BaseImage_width = m_width
End Function

Private Function iV_BaseImage_height() As Long
    iV_BaseImage_height = m_height
End Function

Private Function iV_BaseImage_pixelFormat() As pixel_format
    iV_BaseImage_pixelFormat = m_format
End Function

Private Function iV_BaseImage_data() As Long
    iV_BaseImage_data = data()
End Function

Private Function iV_BaseImage_dataSize() As Long
    iV_BaseImage_dataSize = m_data_size
End Function

Private Function iV_BaseImage_bufferRowLength() As Long
    iV_BaseImage_bufferRowLength = bufferRowLength()
End Function

Private Function iV_BaseImage_bufferImageHeight() As Long
    iV_BaseImage_bufferImageHeight = bufferImageHeight()
End Function
