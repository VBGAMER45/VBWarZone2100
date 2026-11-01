Attribute VB_Name = "jpeg_util_gdip"
' lib/ivis_opengl/jpeg_util.cpp, jpeg_util.h: iV_loadImage_JPEG, with GDI+ (gdiplus.dll's flat API, its own
' GdiplusStartup) in place of libjpeg-turbo: tjDecompressHeader2's size is the GDI+ bitmap's, tjDecompress2 to
' 8-bit RGB is GDI+'s decode read as 32bpp ARGB and kept as R, G, B. GDI+'s IDCT is not TJFLAG_FASTDCT's: a
' pixel may differ from 4.7.0's by a few levels (the image a lobby prompt shows: nothing compares its bytes). A
' buffer GDI+ cannot read is "Invalid width and/or height", as tjDecompressHeader2 leaves the size at 0.
' JPEGResult is its error text ("" = JPEGResult::NoError()).
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As IUnknown) As Long

Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Type GpRect
    x As Long
    y As Long
    Width As Long
    Height As Long
End Type

Private Type GpBitmapData
    Width As Long
    Height As Long
    stride As Long
    PixelFormat As Long
    Scan0 As Long
    Reserved As Long
End Type

Private Declare Function GdiplusStartup Lib "gdiplus" (token As Long, inputbuf As GdiplusStartupInput, ByVal outputbuf As Long) As Long
Private Declare Function GdipCreateBitmapFromStream Lib "gdiplus" (ByVal stream As IUnknown, bitmap As Long) As Long
Private Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal image As Long, Width As Long) As Long
Private Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal image As Long, Height As Long) As Long
Private Declare Function GdipBitmapLockBits Lib "gdiplus" (ByVal bitmap As Long, rect As GpRect, ByVal flags As Long, ByVal PixelFormat As Long, lockedBitmapData As GpBitmapData) As Long
Private Declare Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal bitmap As Long, lockedBitmapData As GpBitmapData) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal image As Long) As Long

Private Const GMEM_MOVEABLE As Long = &H2
Private Const ImageLockModeRead As Long = 1
Private Const PixelFormat32bppARGB As Long = &H26200A

' jpeg_util.h
Public Const WZ_JPEG_MAX_IMAGE_DIMENSIONS As Long = 4096

Private gdipToken As Long
Private gdipStarted As Boolean

Private Function GdipEnsureJpeg() As Boolean
    Dim si As GdiplusStartupInput
    If Not gdipStarted Then
        si.GdiplusVersion = 1
        If GdiplusStartup(gdipToken, si, 0) = 0 Then gdipStarted = True
    End If
    GdipEnsureJpeg = gdipStarted
End Function

' (VB) a GDI+ bitmap of the bytes data(0 .. n - 1) (0 when GDI+ cannot read them)
Private Function bitmapFromMemory(data() As Byte, ByVal n As Long) As Long
    Dim hMem As Long, p As Long, stm As IUnknown, bitmap As Long
    If n <= 0 Then Exit Function
    If Not GdipEnsureJpeg() Then Exit Function
    hMem = GlobalAlloc(GMEM_MOVEABLE, n)
    If hMem = 0 Then Exit Function
    p = GlobalLock(hMem)
    CopyMemory ByVal p, data(LBound(data)), n
    GlobalUnlock hMem
    If CreateStreamOnHGlobal(hMem, 1, stm) <> 0 Then
        GlobalFree hMem
        Exit Function
    End If
    If GdipCreateBitmapFromStream(stm, bitmap) <> 0 Then bitmap = 0
    Set stm = Nothing
    bitmapFromMemory = bitmap
End Function

' JPEGResult iV_loadImage_JPEG(std::vector<unsigned char>&& memoryBuffer, iV_Image *image)
Public Function iV_loadImage_JPEG(memoryBuffer() As Byte, ByVal image As iV_Image) As String
    Dim n As Long, bitmap As Long, width_ As Long, height_ As Long, r As GpRect, bd As GpBitmapData
    Dim row() As Byte, outRow() As Byte, x As Long, y As Long, dst As Long

    On Error Resume Next
    n = UBound(memoryBuffer) - LBound(memoryBuffer) + 1
    If Err.Number <> 0 Then n = 0
    On Error GoTo 0

    ' Initialize; Get image info
    bitmap = bitmapFromMemory(memoryBuffer, n)
    If bitmap <> 0 Then
        GdipGetImageWidth bitmap, width_
        GdipGetImageHeight bitmap, height_
    End If

    If width_ <= 0 Or height_ <= 0 Then
        If bitmap <> 0 Then GdipDisposeImage bitmap
        iV_loadImage_JPEG = "Invalid width and/or height"
        Exit Function
    End If
    If width_ > WZ_JPEG_MAX_IMAGE_DIMENSIONS Or height_ > WZ_JPEG_MAX_IMAGE_DIMENSIONS Then
        GdipDisposeImage bitmap
        iV_loadImage_JPEG = "Too big width and/or height"
        Exit Function
    End If

    ' Decompress to 8-bit RGB
    If Not image.allocate(width_, height_, 3, True) Then
        GdipDisposeImage bitmap
        iV_loadImage_JPEG = "Failed to allocate memory"
        Exit Function
    End If

    r.Width = width_
    r.Height = height_
    If GdipBitmapLockBits(bitmap, r, ImageLockModeRead, PixelFormat32bppARGB, bd) = 0 Then
        ReDim row(0 To width_ * 4 - 1)
        ReDim outRow(0 To width_ * 3 - 1)
        dst = image.bmp_w()
        For y = 0 To height_ - 1
            CopyMemory row(0), ByVal (bd.Scan0 + y * bd.stride), width_ * 4
            ' GDI+'s 32bpp ARGB is B, G, R, A in memory
            For x = 0 To width_ - 1
                outRow(x * 3) = row(x * 4 + 2)
                outRow(x * 3 + 1) = row(x * 4 + 1)
                outRow(x * 3 + 2) = row(x * 4)
            Next
            CopyMemory ByVal (dst + y * width_ * 3), outRow(0), width_ * 3
        Next
        GdipBitmapUnlockBits bitmap, bd
    End If

    ' Cleanup
    GdipDisposeImage bitmap

    iV_loadImage_JPEG = ""      ' JPEGResult::NoError()
End Function
