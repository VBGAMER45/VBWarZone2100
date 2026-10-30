Attribute VB_Name = "pieimage_c"
' lib/ivis_opengl/pieimage.cpp: iV_Image's static member (the class is
' iV_Image.cls) and (VB) ByteArraySwap, which exchanges two dynamic Byte arrays
' in O(1), as C++ swaps the images' unsigned char * pointers.
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Var() As Any) As Long

' static gfx_api::pixel_format iV_Image::pixel_format_for_channels(unsigned int channels)
Public Function iV_Image_pixel_format_for_channels(ByVal channels As Long) As pixel_format
    Select Case channels
    Case 1
        iV_Image_pixel_format_for_channels = pixel_format_FORMAT_R8_UNORM
    Case 2
        iV_Image_pixel_format_for_channels = pixel_format_FORMAT_RG8_UNORM
    Case 3
        iV_Image_pixel_format_for_channels = pixel_format_FORMAT_RGB8_UNORM_PACK8
    Case 4
        iV_Image_pixel_format_for_channels = pixel_format_FORMAT_RGBA8_UNORM_PACK8
    Case Else
        debug_ 155, LOG_FATAL, "pixel_format_for_channels", "iV_getPixelFormat: Unsupported image channels: %u", channels
        iV_Image_pixel_format_for_channels = pixel_format_invalid
    End Select
End Function

' (VB) std::swap of two buffers: exchanges the arrays' descriptors (both dynamic, neither locked)
Public Sub ByteArraySwap(a() As Byte, b() As Byte)
    Dim pa As Long, pb As Long, t As Long
    pa = VarPtrArray(a)
    pb = VarPtrArray(b)
    CopyMemory t, ByVal pa, 4
    CopyMemory ByVal pa, ByVal pb, 4
    CopyMemory ByVal pb, t, 4
End Sub
