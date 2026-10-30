Attribute VB_Name = "png_util_gdip"
' lib/ivis_opengl/png_util_libpng.cpp, lib/ivis_opengl/png_util.h: PNG files,
' with GDI+ (gdiplus.dll's flat API) in place of libpng. This is the png_util
' of the Windows build (png_util_spng.cpp is the Emscripten one): the same
' flow, messages and channel counts as libpng's transforms give them
' (iV_loadImage_PNG: always RGBA; iV_loadImage_PNG2: grey 1, grey + alpha 2,
' RGB and palette 3, with tRNS one more, RGBA 4; forceRGBA8: 4; 16-bit
' samples cut to their high byte). The file's bytes are read through PhysFS,
' the IHDR and tRNS chunks are read here, and GDI+ decodes the rest from a
' memory stream to 32bpp ARGB after the gAMA, cHRM, sRGB and iCCP chunks are
' dropped, so that it applies no colour correction (libpng applies none
' without png_set_gamma): checked byte for byte against an exact decode on
' every kind of PNG in the game data. libpng's error longjmp (png_error,
' setjmp) is a GoTo to the setjmp branch's code.
' Saving: GDI+'s PNG encoder for RGB and RGBA; the other colour types (grey,
' used by iV_saveImage_PNG_Gray) are written here with stored (uncompressed)
' deflate blocks, which every PNG reader reads (larger files than libpng's).
' IMGSaveError is its text ("" = IMGSaveError::None). iV_saveImage_JPEG is not
' ported (no JPEG encoder: it logs and writes nothing).
' The memory-buffer overload of iV_loadImage_PNG is iV_loadImage_PNG_mem.
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As IUnknown) As Long
Private Declare Function GetHGlobalFromStream Lib "ole32" (ByVal pstm As IUnknown, phglobal As Long) As Long

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

Private Type GUID_
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Long, pclsid As GUID_) As Long
Private Declare Function GdiplusStartup Lib "gdiplus" (token As Long, inputbuf As GdiplusStartupInput, ByVal outputbuf As Long) As Long
Private Declare Function GdipCreateBitmapFromStream Lib "gdiplus" (ByVal stream As IUnknown, bitmap As Long) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal Width As Long, ByVal Height As Long, ByVal stride As Long, ByVal PixelFormat As Long, ByVal Scan0 As Long, bitmap As Long) As Long
Private Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal image As Long, Width As Long) As Long
Private Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal image As Long, Height As Long) As Long
Private Declare Function GdipBitmapLockBits Lib "gdiplus" (ByVal bitmap As Long, rect As GpRect, ByVal flags As Long, ByVal PixelFormat As Long, lockedBitmapData As GpBitmapData) As Long
Private Declare Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal bitmap As Long, lockedBitmapData As GpBitmapData) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal image As Long) As Long
Private Declare Function GdipSaveImageToStream Lib "gdiplus" (ByVal image As Long, ByVal stream As IUnknown, clsidEncoder As GUID_, ByVal encoderParams As Long) As Long

Private Const GMEM_MOVEABLE As Long = &H2
Private Const ImageLockModeRead As Long = 1
Private Const PixelFormat24bppRGB As Long = &H21808
Private Const PixelFormat32bppARGB As Long = &H26200A
Private Const PNG_ENCODER_CLSID As String = "{557CF406-1A04-11D3-9A73-0000F81EF32E}"

' png_util.h
Public Const WZ_PNG_MAX_IMAGE_DIMENSIONS As Long = 4096
Public Const WZ_PNG_MAX_DECODED_MEMSIZE As Long = 67108864     ' (WZ_PNG_MAX_IMAGE_DIMENSIONS * WZ_PNG_MAX_IMAGE_DIMENSIONS * 4)
Public Const WZ_PNG_MAX_CHUNKSIZE_LIMIT As Long = 67108864     ' (1024 * 1024 * 64)

Private Const PNG_BYTES_TO_CHECK As Long = 8

' png.h
Private Const PNG_COLOR_TYPE_GRAY As Long = 0
Private Const PNG_COLOR_TYPE_RGB As Long = 2
Private Const PNG_COLOR_TYPE_PALETTE As Long = 3
Private Const PNG_COLOR_TYPE_GRAY_ALPHA As Long = 4
Private Const PNG_COLOR_TYPE_RGB_ALPHA As Long = 6
Private Const PNG_INTERLACE_NONE As Long = 0

' (VB) what libpng's png_read_info gives: the IHDR chunk and whether a valid tRNS chunk came
Private Type PngInfo
    Width As Long
    Height As Long
    bit_depth As Long
    color_type As Long
    interlace_type As Long
    has_tRNS As Boolean                 ' png_get_valid(png_ptr, info_ptr, PNG_INFO_tRNS)
End Type

Private gdipToken As Long
Private gdipStarted As Boolean
Private crcTable(0 To 255) As Long
Private crcTableReady As Boolean

'--------------------------------------------------------------------------
' (VB) the libpng stand-in

Private Function GdipEnsure() As Boolean
    Dim si As GdiplusStartupInput
    If Not gdipStarted Then
        si.GdiplusVersion = 1
        If GdiplusStartup(gdipToken, si, 0) = 0 Then gdipStarted = True
    End If
    GdipEnsure = gdipStarted
End Function

' png_sig_cmp(PNGheader, 0, PNG_BYTES_TO_CHECK) == 0
Private Function PngSigOk(header() As Byte, ByVal first As Long) As Boolean
    PngSigOk = header(first) = 137 And header(first + 1) = 80 And header(first + 2) = 78 And header(first + 3) = 71 _
        And header(first + 4) = 13 And header(first + 5) = 10 And header(first + 6) = 26 And header(first + 7) = 10
End Function

Private Function BE32(b() As Byte, ByVal pos As Long) As Long
    If b(pos) >= &H80 Then
        BE32 = ((b(pos) - &H100&) * &H1000000) Or (b(pos + 1) * &H10000) Or (b(pos + 2) * &H100&) Or b(pos + 3)
    Else
        BE32 = (b(pos) * &H1000000) Or (b(pos + 1) * &H10000) Or (b(pos + 2) * &H100&) Or b(pos + 3)
    End If
End Function

' The file's chunks (src(0 .. n - 1), the signature first): the IHDR and tRNS facts, and the file again
' without its colour-correction chunks (dst, dstN). False where libpng would png_error: no IHDR first, a
' chunk past the end of the data or over the chunk limit, or a size over the user limits.
Private Function PngReadInfo(src() As Byte, ByVal n As Long, info As PngInfo, dst() As Byte, dstN As Long) As Boolean
    Dim pos As Long, ln As Long, typ As String, sawIHDR As Boolean, i As Long
    ReDim dst(0 To n - 1)
    CopyMemory dst(0), src(0), PNG_BYTES_TO_CHECK
    dstN = PNG_BYTES_TO_CHECK
    pos = PNG_BYTES_TO_CHECK
    Do While pos + 8 <= n
        ln = BE32(src, pos)
        If ln < 0 Or ln > WZ_PNG_MAX_CHUNKSIZE_LIMIT Then Exit Function
        If pos + 12 + ln > n Then Exit Function
        typ = Chr$(src(pos + 4)) & Chr$(src(pos + 5)) & Chr$(src(pos + 6)) & Chr$(src(pos + 7))
        If Not sawIHDR And typ <> "IHDR" Then Exit Function
        Select Case typ
        Case "IHDR"
            If ln < 13 Then Exit Function
            info.Width = BE32(src, pos + 8)
            info.Height = BE32(src, pos + 12)
            info.bit_depth = src(pos + 16)
            info.color_type = src(pos + 17)
            info.interlace_type = src(pos + 20)
            sawIHDR = True
        Case "tRNS"
            ' libpng keeps a tRNS only for the types without an alpha channel
            If info.color_type = PNG_COLOR_TYPE_GRAY Or info.color_type = PNG_COLOR_TYPE_RGB Or info.color_type = PNG_COLOR_TYPE_PALETTE Then info.has_tRNS = True
        End Select
        If typ <> "gAMA" And typ <> "cHRM" And typ <> "sRGB" And typ <> "iCCP" Then
            CopyMemory dst(dstN), src(pos), 12 + ln
            dstN = dstN + 12 + ln
        End If
        pos = pos + 12 + ln
        If typ = "IEND" Then Exit Do
    Loop
    If Not sawIHDR Then Exit Function
    ' png_set_user_limits(png_ptr, WZ_PNG_MAX_IMAGE_DIMENSIONS, WZ_PNG_MAX_IMAGE_DIMENSIONS)
    If info.Width <= 0 Or info.Height <= 0 Or info.Width > WZ_PNG_MAX_IMAGE_DIMENSIONS Or info.Height > WZ_PNG_MAX_IMAGE_DIMENSIONS Then Exit Function
    PngReadInfo = True
End Function

' GDI+ decodes png(0 .. n - 1) (w x h) into image, as channels channels per pixel: 1 grey (R), 2 grey and
' alpha, 3 RGB, 4 RGBA (the image is allocated first by the caller). False if GDI+ fails.
Private Function GdipDecodeInto(png() As Byte, ByVal n As Long, ByVal w As Long, ByVal h As Long, ByVal channels As Long, ByVal image As iV_Image) As Boolean
    Dim hMem As Long, p As Long, stm As IUnknown, bitmap As Long, gw As Long, gh As Long
    Dim r As GpRect, bd As GpBitmapData, row() As Byte, outRow() As Byte, x As Long, y As Long, dst As Long
    If Not GdipEnsure() Then Exit Function
    hMem = GlobalAlloc(GMEM_MOVEABLE, n)
    If hMem = 0 Then Exit Function
    p = GlobalLock(hMem)
    CopyMemory ByVal p, png(0), n
    GlobalUnlock hMem
    If CreateStreamOnHGlobal(hMem, 1, stm) <> 0 Then
        GlobalFree hMem
        Exit Function
    End If
    If GdipCreateBitmapFromStream(stm, bitmap) <> 0 Then Exit Function
    GdipGetImageWidth bitmap, gw
    GdipGetImageHeight bitmap, gh
    If gw <> w Or gh <> h Then GoTo done
    r.Width = w
    r.Height = h
    If GdipBitmapLockBits(bitmap, r, ImageLockModeRead, PixelFormat32bppARGB, bd) <> 0 Then GoTo done
    ReDim row(0 To w * 4 - 1)
    ReDim outRow(0 To w * channels - 1)
    dst = image.bmp_w()
    For y = 0 To h - 1
        CopyMemory row(0), ByVal (bd.Scan0 + y * bd.stride), w * 4
        ' GDI+'s 32bpp ARGB is B, G, R, A in memory
        Select Case channels
        Case 4
            For x = 0 To w - 1
                outRow(x * 4) = row(x * 4 + 2)
                outRow(x * 4 + 1) = row(x * 4 + 1)
                outRow(x * 4 + 2) = row(x * 4)
                outRow(x * 4 + 3) = row(x * 4 + 3)
            Next
        Case 3
            For x = 0 To w - 1
                outRow(x * 3) = row(x * 4 + 2)
                outRow(x * 3 + 1) = row(x * 4 + 1)
                outRow(x * 3 + 2) = row(x * 4)
            Next
        Case 2
            For x = 0 To w - 1
                outRow(x * 2) = row(x * 4 + 2)
                outRow(x * 2 + 1) = row(x * 4 + 3)
            Next
        Case Else
            For x = 0 To w - 1
                outRow(x) = row(x * 4 + 2)
            Next
        End Select
        CopyMemory ByVal (dst + y * w * channels), outRow(0), w * channels
    Next
    GdipBitmapUnlockBits bitmap, bd
    GdipDecodeInto = True
done:
    GdipDisposeImage bitmap
    Set stm = Nothing
End Function

' (VB) the rest of an open file after its first PNG_BYTES_TO_CHECK bytes (wzpng_read_data's reads): the whole
' file in file(0 .. n - 1). False on a read failure (libpng: png_error).
Private Function ReadRestOfFile(ByVal fileHandle As Long, PNGheader() As Byte, file() As Byte, n As Long) As Boolean
    Dim total As Double, chunk() As Byte, got As Double
    total = PHYSFS_fileLength(fileHandle)
    If total >= PNG_BYTES_TO_CHECK Then
        n = CLng(total)
        ReDim file(0 To n - 1)
        CopyMemory file(0), PNGheader(0), PNG_BYTES_TO_CHECK
        If n > PNG_BYTES_TO_CHECK Then
            If WZ_PHYSFS_readBytes(fileHandle, file, PNG_BYTES_TO_CHECK, n - PNG_BYTES_TO_CHECK) <> n - PNG_BYTES_TO_CHECK Then Exit Function
        End If
        ReadRestOfFile = True
        Exit Function
    End If
    ' the length is unknown: read to the end
    n = PNG_BYTES_TO_CHECK
    ReDim file(0 To 65535)
    CopyMemory file(0), PNGheader(0), PNG_BYTES_TO_CHECK
    Do
        If n + 65536 > UBound(file) + 1 Then ReDim Preserve file(0 To 2 * (UBound(file) + 1) - 1)
        got = WZ_PHYSFS_readBytes(fileHandle, file, n, 65536)
        If got < 0 Then Exit Function
        n = n + CLng(got)
    Loop While got > 0
    ReadRestOfFile = True
End Function

' (VB) the channels libpng's transforms give iV_loadImage_PNG2 (png_get_channels after png_read_update_info)
Private Function PngOutputChannels(info As PngInfo, ByVal forceRGBA8 As Boolean) As Long
    Dim channels As Long
    Select Case info.color_type
    Case PNG_COLOR_TYPE_GRAY: channels = 1
    Case PNG_COLOR_TYPE_RGB: channels = 3
    Case PNG_COLOR_TYPE_PALETTE: channels = 3      ' png_set_palette_to_rgb
    Case PNG_COLOR_TYPE_GRAY_ALPHA: channels = 2
    Case PNG_COLOR_TYPE_RGB_ALPHA: channels = 4
    Case Else: channels = 0
    End Select
    If info.has_tRNS Then channels = channels + 1  ' png_set_tRNS_to_alpha
    If forceRGBA8 Then
        ' png_set_gray_to_rgb, then png_set_filler(png_ptr, 0xff, PNG_FILLER_AFTER)
        If channels <= 2 Then channels = channels + 2
        If channels = 3 Then channels = 4
    End If
    PngOutputChannels = channels
End Function

'--------------------------------------------------------------------------

Public Function iV_loadImage_PNG(ByVal fileName As String, ByVal image As iV_Image) As Boolean
    Dim PNGheader() As Byte, readSize As Double, fileHandle As Long
    Dim file() As Byte, n As Long, stripped() As Byte, strippedN As Long, info As PngInfo
    Dim width_ As Long, height_ As Long, depth As Long, decoded_image_size As Double

    ReDim PNGheader(0 To PNG_BYTES_TO_CHECK - 1)
    ' Open file
    fileHandle = PHYSFS_openRead(fileName)
    If Not (fileHandle <> 0) Then ASSERT_FAILURE "fileHandle != nullptr", "lib/ivis_opengl/png_util_libpng.cpp:119", "iV_loadImage_PNG", "Could not open %s: %s", fileName, WZ_PHYSFS_getLastError(): iV_loadImage_PNG = False: Exit Function
    ' WZ_PHYSFS_SETBUFFER(fileHandle, 4096)

    ' Read PNG header from file
    readSize = WZ_PHYSFS_readBytes(fileHandle, PNGheader, 0, PNG_BYTES_TO_CHECK)
    If readSize < PNG_BYTES_TO_CHECK Then
        debug_ 126, LOG_FATAL, "iV_loadImage_PNG", "pie_PNGLoadFile: WZ_WZ_PHYSFS_readBytes(%s) failed with error: %s" & vbLf, fileName, WZ_PHYSFS_getLastError()
        PHYSFS_close fileHandle
        iV_loadImage_PNG = False
        Exit Function
    End If

    ' Verify the PNG header to be correct
    If Not PngSigOk(PNGheader, 0) Then
        debug_ 134, LOG_FATAL, "iV_loadImage_PNG", "pie_PNGLoadMem: Did not recognize PNG header in %s", fileName
        PHYSFS_close fileHandle
        iV_loadImage_PNG = False
        Exit Function
    End If

    ' (png_create_read_struct, png_set_user_limits, png_create_info_struct: PngReadInfo and GDI+)

    ' Tell libpng how many byte we already read; set up the input control
    If Not ReadRestOfFile(fileHandle, PNGheader, file, n) Then GoTo pngError

    ' Most of the following transformations are seemingly not needed
    ' Filler is, however, for an unknown reason required for tertilesc[23]
    ' png_set_strip_16, png_set_gray_to_rgb, png_set_palette_to_rgb, png_set_tRNS_to_alpha,
    ' png_set_filler(png_ptr, 0xff, PNG_FILLER_AFTER): (VB) 4 channels, from GDI+'s 32bpp ARGB
    If Not PngReadInfo(file, n, info, stripped, strippedN) Then GoTo pngError

    width_ = info.Width
    height_ = info.Height
    depth = 4
    decoded_image_size = CDbl(width_) * CDbl(height_) * CDbl(depth)
    If decoded_image_size > WZ_PNG_MAX_DECODED_MEMSIZE Then
        GoTo pngError                   ' png_error(png_ptr, "unexpectedly large decoded image size");
    End If

    image.allocate width_, height_, depth
    If Not GdipDecodeInto(stripped, strippedN, width_, height_, depth, image) Then GoTo pngError

    PHYSFS_close fileHandle

    If Not (image.channels() > 3) Then ASSERT_FAILURE "image->channels() > 3", "lib/ivis_opengl/png_util_libpng.cpp:226", "iV_loadImage_PNG", "Unsupported image channels (%d) found.  We only support 3 (RGB) or 4 (ARGB)", image.channels(): iV_loadImage_PNG = False: Exit Function

    iV_loadImage_PNG = True
    Exit Function

pngError:
    ' libpng's failure jump position (setjmp)
    debug_ 167, LOG_FATAL, "iV_loadImage_PNG", "pie_PNGLoadMem: Error decoding PNG data in %s", fileName
    PHYSFS_close fileHandle
    iV_loadImage_PNG = False
End Function

Public Function iV_loadImage_PNG2(ByVal fileName As String, ByVal image As iV_Image, Optional ByVal forceRGBA8 As Boolean = False, Optional ByVal quietOnOpenFail As Boolean = False) As Boolean
    Dim PNGheader() As Byte, readSize As Double, fileHandle As Long
    Dim file() As Byte, n As Long, stripped() As Byte, strippedN As Long, info As PngInfo
    Dim width_ As Long, height_ As Long, channels As Long, row_bytes As Long, destFormat As pixel_format
    Dim decoded_image_size As Double

    ReDim PNGheader(0 To PNG_BYTES_TO_CHECK - 1)
    ' Open file
    fileHandle = PHYSFS_openRead(fileName)
    If fileHandle = 0 Then
        If Not quietOnOpenFail Then
            If Not (fileHandle <> 0) Then ASSERT_FAILURE "fileHandle != nullptr", "lib/ivis_opengl/png_util_libpng.cpp:255", "iV_loadImage_PNG2", "Could not open %s: %s", fileName, WZ_PHYSFS_getLastError(): iV_loadImage_PNG2 = False: Exit Function
        Else
            iV_loadImage_PNG2 = False
            Exit Function
        End If
    End If
    ' WZ_PHYSFS_SETBUFFER(fileHandle, 4096)

    ' Read PNG header from file
    readSize = WZ_PHYSFS_readBytes(fileHandle, PNGheader, 0, PNG_BYTES_TO_CHECK)
    If readSize < PNG_BYTES_TO_CHECK Then
        debug_ 268, LOG_FATAL, "iV_loadImage_PNG2", "pie_PNGLoadFile: WZ_WZ_PHYSFS_readBytes(%s) failed with error: %s" & vbLf, fileName, WZ_PHYSFS_getLastError()
        PHYSFS_close fileHandle
        iV_loadImage_PNG2 = False
        Exit Function
    End If

    ' Verify the PNG header to be correct
    If Not PngSigOk(PNGheader, 0) Then
        debug_ 276, LOG_FATAL, "iV_loadImage_PNG2", "pie_PNGLoadMem: Did not recognize PNG header in %s", fileName
        PHYSFS_close fileHandle
        iV_loadImage_PNG2 = False
        Exit Function
    End If

    ' Tell libpng how many byte we already read; set up the input control
    If Not ReadRestOfFile(fileHandle, PNGheader, file, n) Then GoTo pngError

    ' Read the PNG info / header
    If Not PngReadInfo(file, n, info, stripped, strippedN) Then GoTo pngError
    width_ = info.Width
    height_ = info.Height

    If info.interlace_type <> PNG_INTERLACE_NONE Then
        GoTo pngError                   ' png_error(png_ptr, "interlaced PNGs are not currently supported");
    End If

    ' Most of the following transformations are seemingly not needed
    ' Filler is, however, for an unknown reason required for tertilesc[23]
    ' (VB) png_set_strip_16, png_set_palette_to_rgb, png_set_tRNS_to_alpha, png_set_expand_gray_1_2_4_to_8,
    ' png_set_packing, and with forceRGBA8 png_set_gray_to_rgb and png_set_filler: their channels, at 8 bits
    channels = PngOutputChannels(info, forceRGBA8)
    row_bytes = width_ * channels

    ' Figure out output image type
    destFormat = iV_Image_pixel_format_for_channels(channels)
    If destFormat = pixel_format_invalid Then
        PHYSFS_close fileHandle
        iV_loadImage_PNG2 = False
        Exit Function
    End If

    decoded_image_size = CDbl(width_) * CDbl(height_) * CDbl(channels)
    If decoded_image_size > WZ_PNG_MAX_DECODED_MEMSIZE Then
        debug_ 420, LOG_FATAL, "iV_loadImage_PNG2", "pie_PNGLoadMem: Unexpectedly large decoded image size (%zu): %s", decoded_image_size, fileName
        PHYSFS_close fileHandle
        iV_loadImage_PNG2 = False
        Exit Function
    End If

    ' Allocate buffer
    If Not image.allocate(width_, height_, channels) Then
        debug_ 428, LOG_FATAL, "iV_loadImage_PNG2", "pie_PNGLoadMem: Unable to allocate memory to load: %s", fileName
        PHYSFS_close fileHandle
        iV_loadImage_PNG2 = False
        Exit Function
    End If

    ' read image rows
    If Not GdipDecodeInto(stripped, strippedN, width_, height_, channels, image) Then GoTo pngError

    PHYSFS_close fileHandle

    iV_loadImage_PNG2 = True
    Exit Function

pngError:
    ' libpng's failure jump position (setjmp)
    debug_ 309, LOG_FATAL, "iV_loadImage_PNG2", "pie_PNGLoadMem: Error decoding PNG data in %s", fileName
    PHYSFS_close fileHandle
    iV_loadImage_PNG2 = False
End Function

' Note: This function must be thread-safe.
'       It does not call the debug() macro directly, but instead returns an IMGSaveError structure with the text of any error.
' IMGSaveError iV_loadImage_PNG(const std::vector<unsigned char>& memoryBuffer, iV_Image *image): memoryBuffer(0 .. UBound)
Public Function iV_loadImage_PNG_mem(memoryBuffer() As Byte, ByVal image As iV_Image) As String
    Dim n As Long, stripped() As Byte, strippedN As Long, info As PngInfo
    Dim width_ As Long, height_ As Long, depth As Long, decoded_image_size As Double, file() As Byte

    n = ByteArraySize(memoryBuffer)
    If n < PNG_BYTES_TO_CHECK Then
        iV_loadImage_PNG_mem = "iV_loadImage_PNG: Insufficient length for PNG header"
        Exit Function
    End If
    ReDim file(0 To n - 1)
    CopyMemory file(0), memoryBuffer(LBound(memoryBuffer)), n

    ' Verify the PNG header to be correct
    If Not PngSigOk(file, 0) Then
        iV_loadImage_PNG_mem = "iV_loadImage_PNG: Did not recognize PNG header"
        Exit Function
    End If

    ' png_set_strip_16, png_set_gray_to_rgb, png_set_palette_to_rgb, png_set_tRNS_to_alpha,
    ' png_set_filler(png_ptr, 0xff, PNG_FILLER_AFTER): (VB) 4 channels, from GDI+'s 32bpp ARGB
    If Not PngReadInfo(file, n, info, stripped, strippedN) Then GoTo pngError

    width_ = info.Width
    height_ = info.Height
    depth = 4
    decoded_image_size = CDbl(width_) * CDbl(height_) * CDbl(depth)
    If decoded_image_size > WZ_PNG_MAX_DECODED_MEMSIZE Then
        GoTo pngError                   ' png_error(png_ptr, "unexpectedly large decoded image size");
    End If

    image.allocate width_, height_, depth
    If Not GdipDecodeInto(stripped, strippedN, width_, height_, depth, image) Then GoTo pngError

    If image.channels() < 3 Or image.channels() > 4 Then
        iV_loadImage_PNG_mem = "Unsupported image channels (" & CStr(image.channels()) & ") found.  We only support 3 (RGB) or 4 (ARGB)"
        Exit Function
    End If

    iV_loadImage_PNG_mem = ""
    Exit Function

pngError:
    iV_loadImage_PNG_mem = "iV_loadImage_PNG: Error decoding PNG data"
End Function

'--------------------------------------------------------------------------
' (VB) PNG writing without libpng

Private Sub crcTableInit()
    Dim n As Long, c As Long, k As Long
    For n = 0 To 255
        c = n
        For k = 0 To 7
            If (c And 1) <> 0 Then
                c = &HEDB88320 Xor SHRU(c, 1)
            Else
                c = SHRU(c, 1)
            End If
        Next
        crcTable(n) = c
    Next
    crcTableReady = True
End Sub

' the PNG chunk CRC (CRC-32) of b(first .. first + n - 1)
Private Function pngCrc(b() As Byte, ByVal first As Long, ByVal n As Long) As Long
    Dim c As Long, i As Long
    If Not crcTableReady Then crcTableInit
    c = &HFFFFFFFF
    For i = first To first + n - 1
        c = crcTable((c Xor b(i)) And &HFF&) Xor SHRU(c, 8)
    Next
    pngCrc = c Xor &HFFFFFFFF
End Function

Private Sub PutBE32(b() As Byte, ByVal pos As Long, ByVal v As Long)
    b(pos) = SHRU(v, 24) And &HFF&
    b(pos + 1) = SHRU(v, 16) And &HFF&
    b(pos + 2) = SHRU(v, 8) And &HFF&
    b(pos + 3) = v And &HFF&
End Sub

' writes a chunk of type typ with data(0 .. n - 1) at out(pos); returns the position after it
Private Function PutChunk(out() As Byte, ByVal pos As Long, ByVal typ As String, data() As Byte, ByVal n As Long) As Long
    Dim i As Long
    PutBE32 out, pos, n
    For i = 0 To 3
        out(pos + 4 + i) = Asc(Mid$(typ, i + 1, 1))
    Next
    If n > 0 Then CopyMemory out(pos + 8), data(0), n
    PutBE32 out, pos + 8 + n, pngCrc(out, pos + 4, 4 + n)
    PutChunk = pos + 12 + n
End Function

' A PNG of w x h 8-bit pixels of color_type, the rows taken from rows(0 .. rowsSize - 1) at rowStride (the
' first w * bytesPerPixel bytes of each, as libpng reads them; bytes past the end are 0), in rowOrder (the
' source row of each output row), into out; returns its size. The zlib stream has stored blocks.
Private Function PngEncodeStored(rows() As Byte, ByVal rowsSize As Long, ByVal rowStride As Long, rowOrder() As Long, ByVal w As Long, ByVal h As Long, ByVal color_type As Long, out() As Byte) As Long
    Dim bpp As Long, rawLen As Long, raw() As Byte, y As Long, src As Long, n As Long, p As Long
    Dim zlen As Long, z() As Byte, blocks As Long, i As Long, blk As Long, s1 As Long, s2 As Long
    Dim ihdr(0 To 12) As Byte, empty_() As Byte, pos As Long
    Select Case color_type
    Case PNG_COLOR_TYPE_GRAY: bpp = 1
    Case PNG_COLOR_TYPE_GRAY_ALPHA: bpp = 2
    Case PNG_COLOR_TYPE_RGB: bpp = 3
    Case Else: bpp = 4
    End Select
    rawLen = h * (1 + w * bpp)
    ReDim raw(0 To rawLen - 1)
    p = 0
    For y = 0 To h - 1
        raw(p) = 0                      ' filter type None
        p = p + 1
        src = rowOrder(y) * rowStride
        n = w * bpp
        If src + n > rowsSize Then n = rowsSize - src
        If n > 0 Then CopyMemory raw(p), rows(src), n
        p = p + w * bpp
    Next
    ' zlib: CMF, FLG, stored deflate blocks of at most 65535 bytes, Adler-32
    blocks = (rawLen + 65534) \ 65535
    If blocks = 0 Then blocks = 1
    zlen = 2 + blocks * 5 + rawLen + 4
    ReDim z(0 To zlen - 1)
    z(0) = &H78
    z(1) = &H1
    pos = 2
    For i = 0 To blocks - 1
        blk = rawLen - i * 65535
        If blk > 65535 Then blk = 65535
        If i = blocks - 1 Then z(pos) = 1 Else z(pos) = 0
        z(pos + 1) = blk And &HFF&
        z(pos + 2) = (blk \ &H100&) And &HFF&
        z(pos + 3) = (65535 - blk) And &HFF&                ' NLEN, the one's complement of LEN
        z(pos + 4) = (65535 - blk) \ &H100&
        If blk > 0 Then CopyMemory z(pos + 5), raw(i * 65535), blk
        pos = pos + 5 + blk
    Next
    s1 = 1
    s2 = 0
    For i = 0 To rawLen - 1
        s1 = s1 + raw(i)
        If s1 >= 65521 Then s1 = s1 - 65521
        s2 = s2 + s1
        If s2 >= 65521 Then s2 = s2 - 65521
    Next
    z(pos) = (s2 \ &H100&) And &HFF&
    z(pos + 1) = s2 And &HFF&
    z(pos + 2) = (s1 \ &H100&) And &HFF&
    z(pos + 3) = s1 And &HFF&

    ReDim out(0 To 8 + 25 + 12 + zlen + 12 - 1)
    out(0) = 137: out(1) = 80: out(2) = 78: out(3) = 71: out(4) = 13: out(5) = 10: out(6) = 26: out(7) = 10
    PutBE32 ihdr, 0, w
    PutBE32 ihdr, 4, h
    ihdr(8) = 8                         ' bit depth
    ihdr(9) = color_type
    ihdr(10) = 0                        ' PNG_COMPRESSION_TYPE_DEFAULT
    ihdr(11) = 0                        ' PNG_FILTER_TYPE_DEFAULT
    ihdr(12) = PNG_INTERLACE_NONE
    pos = PutChunk(out, 8, "IHDR", ihdr, 13)
    pos = PutChunk(out, pos, "IDAT", z, zlen)
    pos = PutChunk(out, pos, "IEND", empty_, 0)
    PngEncodeStored = pos
End Function

' GDI+'s PNG encoder: RGB (3 bytes a pixel) or RGBA (4) rows as above; False if GDI+ fails
Private Function GdipEncodePng(rows() As Byte, ByVal rowStride As Long, rowOrder() As Long, ByVal w As Long, ByVal h As Long, ByVal bpp As Long, out() As Byte, outN As Long) As Boolean
    Dim stride As Long, buf() As Byte, x As Long, y As Long, s As Long, d As Long
    Dim bitmap As Long, stm As IUnknown, clsid As GUID_, hMem As Long, p As Long, size As Long
    Dim all() As Byte, pos As Long, ln As Long, isEnd As Boolean
    If Not GdipEnsure() Then Exit Function
    stride = (w * bpp + 3) And Not 3
    ReDim buf(0 To stride * h - 1)
    For y = 0 To h - 1
        s = rowOrder(y) * rowStride
        d = y * stride
        For x = 0 To w - 1
            ' GDI+'s 24bpp RGB is B, G, R and its 32bpp ARGB B, G, R, A in memory
            buf(d + x * bpp) = rows(s + x * bpp + 2)
            buf(d + x * bpp + 1) = rows(s + x * bpp + 1)
            buf(d + x * bpp + 2) = rows(s + x * bpp)
            If bpp = 4 Then buf(d + x * 4 + 3) = rows(s + x * 4 + 3)
        Next
    Next
    If bpp = 4 Then
        If GdipCreateBitmapFromScan0(w, h, stride, PixelFormat32bppARGB, VarPtr(buf(0)), bitmap) <> 0 Then Exit Function
    Else
        If GdipCreateBitmapFromScan0(w, h, stride, PixelFormat24bppRGB, VarPtr(buf(0)), bitmap) <> 0 Then Exit Function
    End If
    If CreateStreamOnHGlobal(0, 1, stm) <> 0 Then GoTo done
    CLSIDFromString StrPtr(PNG_ENCODER_CLSID), clsid
    If GdipSaveImageToStream(bitmap, stm, clsid, 0) <> 0 Then GoTo done
    If GetHGlobalFromStream(stm, hMem) <> 0 Then GoTo done
    size = GlobalSize(hMem)
    If size < PNG_BYTES_TO_CHECK Then GoTo done
    ReDim all(0 To size - 1)
    p = GlobalLock(hMem)
    CopyMemory all(0), ByVal p, size
    GlobalUnlock hMem
    ' the stream's memory may be larger than the PNG: up to the end of its IEND chunk
    pos = PNG_BYTES_TO_CHECK
    Do While pos + 12 <= size
        ln = BE32(all, pos)
        If ln < 0 Or pos + 12 + ln > size Then Exit Do
        isEnd = (all(pos + 4) = 73 And all(pos + 5) = 69 And all(pos + 6) = 78 And all(pos + 7) = 68)  ' "IEND"
        pos = pos + 12 + ln
        If isEnd Then Exit Do
    Loop
    outN = pos
    ReDim out(0 To outN - 1)
    CopyMemory out(0), all(0), outN
    GdipEncodePng = True
done:
    GdipDisposeImage bitmap
    Set stm = Nothing
End Function

'--------------------------------------------------------------------------

' Note: This function must be thread-safe.
'       It does not call the debug() macro directly, but instead returns an IMGSaveError structure with the text of any error.
Private Function internal_saveImage_PNG(ByVal fileName As String, ByVal image As iV_Image, ByVal color_type As Long, Optional ByVal bottom_up_write As Boolean = True) As String
    Dim fileHandle As Long, channelsPerPixel As Long, currentRow As Long, row_stride As Long, err_ As String
    Dim scanlines() As Long, bitmap_bytes() As Byte, bytesSize As Long, png() As Byte, pngN As Long, ok As Boolean
    If Not (image.channels() <> 0) Then ASSERT_FAILURE "image->channels() != 0", "lib/ivis_opengl/png_util_libpng.cpp:629", "internal_saveImage_PNG", "Bad channels"

    fileHandle = PHYSFS_openWrite(fileName)
    If fileHandle = 0 Then
        err_ = "pie_PNGSaveFile: PHYSFS_openWrite failed (while opening file "
        err_ = err_ & fileName
        err_ = err_ & ") with error: "
        err_ = err_ & WZ_PHYSFS_getLastError()
        internal_saveImage_PNG = err_
        Exit Function
    End If

    channelsPerPixel = image.channels()

    row_stride = image.width() * channelsPerPixel

    If image.height() > 0 Then ReDim scanlines(0 To image.height() - 1)

    ' Set the compression level of ZLIB
    ' Right now we stick with the default, since that one is the
    ' fastest which still produces acceptable filesizes.
    ' The highest compression level hardly produces smaller files than default.
    ' (VB) GDI+'s encoder for RGB and RGBA, stored blocks for the rest (PngEncodeStored)

    ' Create an array of scanlines
    bytesSize = image.size_in_bytes()
    If bytesSize > 0 Then
        ReDim bitmap_bytes(0 To bytesSize - 1)
        CopyMemory bitmap_bytes(0), ByVal image.bmp(), bytesSize
    End If
    For currentRow = 0 To image.height() - 1
        If bottom_up_write Then
            ' We're filling the scanline from the bottom up here,
            ' otherwise we'd have a vertically mirrored image.
            scanlines(currentRow) = (image.height() - currentRow - 1)
        Else
            scanlines(currentRow) = currentRow
        End If
    Next

    If image.width() <= 0 Or image.height() <= 0 Or bytesSize = 0 Then GoTo pngError
    If (color_type = PNG_COLOR_TYPE_RGB And channelsPerPixel = 3) Or (color_type = PNG_COLOR_TYPE_RGB_ALPHA And channelsPerPixel = 4) Then
        ok = GdipEncodePng(bitmap_bytes, row_stride, scanlines, image.width(), image.height(), channelsPerPixel, png, pngN)
    Else
        pngN = PngEncodeStored(bitmap_bytes, bytesSize, row_stride, scanlines, image.width(), image.height(), color_type, png)
        ok = True
    End If
    If Not ok Then GoTo pngError
    WZ_PHYSFS_writeBytes fileHandle, png, 0, pngN

    PHYSFS_close fileHandle
    internal_saveImage_PNG = ""
    Exit Function

pngError:
    ' If libpng encounters an error, it will jump into this if-branch
    PHYSFS_close fileHandle
    internal_saveImage_PNG = "pie_PNGSaveFile: Error encoding PNG data"
End Function

' Note: This function must be thread-safe.
Public Function iV_saveImage_PNG(ByVal fileName As String, ByVal image As iV_Image) As String
    Dim color_type As Long
    color_type = PNG_COLOR_TYPE_RGB
    Select Case image.channels()
    Case 4
        color_type = PNG_COLOR_TYPE_RGB_ALPHA
    Case 3
        color_type = PNG_COLOR_TYPE_RGB
    End Select
    iV_saveImage_PNG = internal_saveImage_PNG(fileName, image, color_type)
End Function

' Note: This function must be thread-safe.
Public Function iV_saveImage_PNG_Gray(ByVal fileName As String, ByVal image As iV_Image) As String
    iV_saveImage_PNG_Gray = internal_saveImage_PNG(fileName, image, PNG_COLOR_TYPE_GRAY)
End Function

' Note: This function is *NOT* thread-safe (jpeg_encode_image is not thread-safe).
' (VB) not ported: no JPEG encoder
Public Sub iV_saveImage_JPEG(ByVal fileName As String, ByVal image As iV_Image)
    debug_ 0, LOG_ERROR, "iV_saveImage_JPEG", "(VB) JPEG screenshots are not ported; not saved: %s", fileName
End Sub

' Note: This function is *NOT* thread-safe.
Public Function iV_LoadAndSavePNG_AsLumaSingleChannel(ByVal inputFilename As String, ByVal outputFilename As String, Optional ByVal check As Boolean = False) As Boolean
    Dim image As New iV_Image, result As Boolean, saveResult As String, outputImage As iV_Image

    ' 1.) Load the PNG into an iV_Image
    If Not iV_loadImage_PNG2(inputFilename, image, False, False) Then
        ' Failed to load the input image
        iV_LoadAndSavePNG_AsLumaSingleChannel = False
        Exit Function
    End If

    result = image.convert_to_luma()
    If Not (result) Then ASSERT_FAILURE "result", "lib/ivis_opengl/png_util_libpng.cpp:823", "iV_LoadAndSavePNG_AsLumaSingleChannel", "(%s): Failed to convert specular map", inputFilename: iV_LoadAndSavePNG_AsLumaSingleChannel = False: Exit Function

    saveResult = internal_saveImage_PNG(outputFilename, image, PNG_COLOR_TYPE_GRAY, False)
    If Not (Len(saveResult) = 0) Then ASSERT_FAILURE "saveResult.noError()", "lib/ivis_opengl/png_util_libpng.cpp:826", "iV_LoadAndSavePNG_AsLumaSingleChannel", "(%s): Failed to save specular map output, with error: %s", inputFilename, saveResult: iV_LoadAndSavePNG_AsLumaSingleChannel = False: Exit Function

    If check Then
        ' Test round-trip - does loading the written image yield the same iV_Image bitmap data?
        Set outputImage = New iV_Image
        If Not iV_loadImage_PNG2(outputFilename, outputImage, False, False) Then
            ' Failed to load the output image
            iV_LoadAndSavePNG_AsLumaSingleChannel = False
            Exit Function
        End If

        If Not (outputImage.channels() = 1) Then ASSERT_FAILURE "outputImage.channels() == 1", "lib/ivis_opengl/png_util_libpng.cpp:838", "iV_LoadAndSavePNG_AsLumaSingleChannel", "Output image unexpectedly has %u channels", outputImage.channels(): iV_LoadAndSavePNG_AsLumaSingleChannel = False: Exit Function
        If Not (image.compare_equal(outputImage)) Then ASSERT_FAILURE "image.compare_equal(outputImage)", "lib/ivis_opengl/png_util_libpng.cpp:839", "iV_LoadAndSavePNG_AsLumaSingleChannel", "Output image doesn't seem to match expected": iV_LoadAndSavePNG_AsLumaSingleChannel = False: Exit Function
        debug_ 840, LOG_INFO, "iV_LoadAndSavePNG_AsLumaSingleChannel", "Validated that output image loads the same bitmap data that was saved."
    End If

    iV_LoadAndSavePNG_AsLumaSingleChannel = True
End Function

' (VB) the number of elements of a 1-D Byte array (0 if it has none)
Private Function ByteArraySize(a() As Byte) As Long
    On Error GoTo none
    ByteArraySize = UBound(a) - LBound(a) + 1
    Exit Function
none:
    ByteArraySize = 0
End Function
