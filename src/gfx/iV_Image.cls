VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iV_Image"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' lib/ivis_opengl/pieimage.cpp, lib/ivis_opengl/pietypes.h: iV_Image, an
' uncompressed image (it implements iV_BaseImage, as iV_CompressedImage does: the
' interface's methods forward to this class's own). The pixels are a Byte
' array; bmp() / bmp_w() / data() give its address (0 if none) for uploads and
' CopyMemory. C++'s move (iV_Image&&) is moveFrom; a function taking iV_Image&&
' takes the object, and the caller drops it. The destructor is Class_Terminate.
' stb_image_resize (resizeInternal) is replaced (VB): a box filter for a size
' that halves or keeps each dimension (the mip levels), a bilinear filter
' otherwise, both as stb with STBIR_FLAG_ALPHA_PREMULTIPLIED (every channel
' filtered alike) and STBIR_EDGE_CLAMP.
' The static pixel_format_for_channels is pieimage_c's
' iV_Image_pixel_format_for_channels; pixel_format() is pixel_format_ (the
' enum's name).
Option Explicit

Implements iV_BaseImage

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)
Private Declare Function RtlCompareMemory Lib "ntdll" (ByVal Source1 As Long, ByVal Source2 As Long, ByVal Length As Long) As Long

Private m_width As Long                 ' unsigned int = 0
Private m_height As Long                ' unsigned int = 0
Private m_channels As Long              ' unsigned int = 0
Private m_bmp() As Byte                 ' unsigned char *m_bmp = nullptr
Private m_hasBmp As Boolean             ' (VB) m_bmp != nullptr
Private m_colorOrder As ColorOrder      ' = ColorOrder::RGB

Private Sub Class_Terminate()
    clear
End Sub

Public Function width() As Long
    width = m_width
End Function

Public Function height() As Long
    height = m_height
End Function

Public Function channels() As Long
    channels = m_channels
End Function

Public Function depth() As Long
    depth = m_channels * 8
End Function

' const unsigned char* data() const: the address of the pixels
Public Function data() As Long
    data = bmp()
End Function

Public Function data_size() As Long
    data_size = size_in_bytes()
End Function

Public Function bufferRowLength() As Long
    bufferRowLength = m_width
End Function

Public Function bufferImageHeight() As Long
    bufferImageHeight = m_height
End Function

Public Function size_in_bytes() As Long
    If bmp() = 0 Then
        size_in_bytes = 0
        Exit Function
    End If
    size_in_bytes = m_width * m_height * m_channels
End Function

' (VB) m_bmp = (unsigned char*)malloc(n): buf is a new buffer of n bytes (none for 0)
Private Function NewBuffer(buf() As Byte, ByVal n As Long) As Boolean
    On Error GoTo fail
    If n > 0 Then
        ReDim buf(0 To n - 1)
    Else
        Erase buf
    End If
    NewBuffer = True
    Exit Function
fail:
End Function

' Allocate a new iV_Image buffer
Public Function allocate(Optional ByVal newWidth As Long = 0, Optional ByVal newHeight As Long = 0, Optional ByVal newChannels As Long = 0, Optional ByVal zeroMemory As Boolean = False) As Boolean
    Dim sizeOfBuffer As Long
    If m_hasBmp Then
        Erase m_bmp
        m_hasBmp = False
    End If
    If newWidth > 0 And newHeight > 0 And newChannels > 0 Then
        sizeOfBuffer = newWidth * newHeight * newChannels
        If Not NewBuffer(m_bmp, sizeOfBuffer) Then
            ASSERT_FAILURE "false", "lib/ivis_opengl/pieimage.cpp:101", "allocate", "Failed to allocate memory buffer of size: %zu", sizeOfBuffer
            allocate = False
            Exit Function
        End If
        m_hasBmp = True
        ' (VB) a new VB array is zeroed, so zeroMemory needs no memset
    End If
    m_width = newWidth
    m_height = newHeight
    m_channels = newChannels
    allocate = True
End Function

Public Function duplicate(ByVal other As iV_Image) As Boolean
    If Not allocate(other.width(), other.height(), other.channels(), False) Then
        duplicate = False
        Exit Function
    End If
    If other.size_in_bytes() > 0 Then CopyMemory ByVal bmp_w(), ByVal other.bmp(), other.size_in_bytes()
    duplicate = True
End Function

Public Sub clear()
    allocate 0, 0, 0
End Sub

' Get a pointer to the bitmap data that can be read
Public Function bmp() As Long
    If m_hasBmp Then bmp = VarPtr(m_bmp(0))
End Function

' Get a pointer to the bitmap data that can be written to
Public Function bmp_w() As Long
    If m_hasBmp Then bmp_w = VarPtr(m_bmp(0))
End Function

Public Function pixel_format_() As pixel_format
    If m_colorOrder = ColorOrder_BGR Then
        Select Case m_channels
        Case 4
            pixel_format_ = pixel_format_FORMAT_BGRA8_UNORM_PACK8
            Exit Function
        Case Else
            ASSERT_FAILURE "false", "lib/ivis_opengl/pieimage.cpp:169", "pixel_format", "BGR color order, but unexpected number of channels: %u", m_channels
            ' fall-through
        End Select
    End If
    pixel_format_ = iV_Image_pixel_format_for_channels(m_channels)
End Function

' template <size_t readChannels, size_t writeChannels, int writeAlphaChannel = -1, unsigned char defaultAlphaValue = 255>
' static void internal_convert_image_contents(unsigned int width, unsigned int height, const unsigned char *pSrc, unsigned char *pDst)
Private Sub internal_convert_image_contents(ByVal readChannels As Long, ByVal writeChannels As Long, ByVal writeAlphaChannel As Long, ByVal defaultAlphaValue As Long, ByVal width_ As Long, ByVal height_ As Long, pSrc() As Byte, pDst() As Byte)
    Dim x As Long, y As Long, i As Long, rgba(0 To 3) As Byte, pSrcPixel As Long, pDstPixel As Long
    For y = 0 To height_ - 1
        For x = 0 To width_ - 1
            rgba(0) = 0: rgba(1) = 0: rgba(2) = 0: rgba(3) = defaultAlphaValue
            pSrcPixel = ((y * width_) + x) * readChannels
            For i = 0 To readChannels - 1
                rgba(i) = pSrc(pSrcPixel + i)
            Next

            pDstPixel = ((y * width_) + x) * writeChannels
            For i = 0 To writeChannels - 1
                pDst(pDstPixel + i) = rgba(i)
            Next
            If writeAlphaChannel >= 0 Then
                pDst(pDstPixel + writeAlphaChannel) = rgba(3)
            End If
        Next
    Next
End Sub

Public Function expand_channels_towards_rgba() As Boolean
    Dim originalBmpData() As Byte, originalChannels As Long, numPixels As Long
    If Not (bmp() <> 0) Then ASSERT_FAILURE "bmp() != nullptr", "lib/ivis_opengl/pieimage.cpp:208", "expand_channels_towards_rgba", "Not initialized": expand_channels_towards_rgba = False: Exit Function
    If Not (m_width > 0 And m_height > 0) Then ASSERT_FAILURE "m_width > 0 && m_height > 0", "lib/ivis_opengl/pieimage.cpp:209", "expand_channels_towards_rgba", "No size": expand_channels_towards_rgba = False: Exit Function
    If m_channels >= 4 Then
        expand_channels_towards_rgba = True
        Exit Function
    End If
    ByteArraySwap originalBmpData, m_bmp ' TODO: Make this bmp()?
    originalChannels = m_channels
    numPixels = m_height * m_width
    m_channels = originalChannels + 1
    If Not NewBuffer(m_bmp, numPixels * m_channels) Then
        ByteArraySwap m_bmp, originalBmpData
        m_channels = originalChannels
        expand_channels_towards_rgba = False
        Exit Function
    End If
    Select Case originalChannels
    Case 1
        internal_convert_image_contents 1, 2, -1, 255, m_width, m_height, originalBmpData, m_bmp
    Case 2
        internal_convert_image_contents 2, 3, -1, 255, m_width, m_height, originalBmpData, m_bmp
    Case 3
        internal_convert_image_contents 3, 4, -1, 255, m_width, m_height, originalBmpData, m_bmp
    Case Else
        ASSERT_FAILURE "false", "lib/ivis_opengl/pieimage.cpp:237", "expand_channels_towards_rgba", "Invalid depth: %u", m_channels
        Erase m_bmp
        ByteArraySwap m_bmp, originalBmpData
        m_channels = originalChannels
        expand_channels_towards_rgba = False
        Exit Function
    End Select
    ' free the original bitmap data
    Erase originalBmpData
    expand_channels_towards_rgba = True
End Function

Public Function convert_to_rgba() As Boolean
    Dim originalBmpData() As Byte, originalChannels As Long, numPixels As Long
    If Not (m_hasBmp) Then ASSERT_FAILURE "m_bmp != nullptr", "lib/ivis_opengl/pieimage.cpp:250", "convert_to_rgba", "Not initialized": convert_to_rgba = False: Exit Function
    If Not (m_width > 0 And m_height > 0) Then ASSERT_FAILURE "m_width > 0 && m_height > 0", "lib/ivis_opengl/pieimage.cpp:251", "convert_to_rgba", "No size": convert_to_rgba = False: Exit Function
    If m_channels = 4 Then
        convert_to_rgba = True
        Exit Function
    End If
    ByteArraySwap originalBmpData, m_bmp
    originalChannels = m_channels
    numPixels = m_height * m_width
    m_channels = 4
    NewBuffer m_bmp, numPixels * m_channels
    Select Case originalChannels
    Case 1
        internal_convert_image_contents 1, 4, -1, 255, m_width, m_height, originalBmpData, m_bmp
    Case 2
        internal_convert_image_contents 2, 4, -1, 255, m_width, m_height, originalBmpData, m_bmp
    Case 3
        internal_convert_image_contents 3, 4, -1, 255, m_width, m_height, originalBmpData, m_bmp
    Case Else
        ASSERT_FAILURE "false", "lib/ivis_opengl/pieimage.cpp:273", "convert_to_rgba", "Invalid depth: %u", m_channels
        Erase m_bmp
        ByteArraySwap m_bmp, originalBmpData
        m_channels = originalChannels
        convert_to_rgba = False
        Exit Function
    End Select
    ' free the original bitmap data
    Erase originalBmpData
    convert_to_rgba = True
End Function

Public Function convert_rg_to_ra_rgba() As Boolean
    Dim numPixels As Long, pixelIdx As Long, pixelStartPos As Long
    If Not (m_hasBmp) Then ASSERT_FAILURE "m_bmp != nullptr", "lib/ivis_opengl/pieimage.cpp:286", "convert_rg_to_ra_rgba", "Not initialized": convert_rg_to_ra_rgba = False: Exit Function
    If Not (m_width > 0 And m_height > 0) Then ASSERT_FAILURE "m_width > 0 && m_height > 0", "lib/ivis_opengl/pieimage.cpp:287", "convert_rg_to_ra_rgba", "No size": convert_rg_to_ra_rgba = False: Exit Function
    If Not (m_channels >= 2) Then ASSERT_FAILURE "m_channels >= 2", "lib/ivis_opengl/pieimage.cpp:288", "convert_rg_to_ra_rgba", "Must have at least 2 channels": convert_rg_to_ra_rgba = False: Exit Function
    If Not (convert_to_rgba()) Then ASSERT_FAILURE "convert_to_rgba()", "lib/ivis_opengl/pieimage.cpp:289", "convert_rg_to_ra_rgba", "Format must be RGBA": convert_rg_to_ra_rgba = False: Exit Function

    numPixels = m_height * m_width
    For pixelIdx = 0 To numPixels - 1
        pixelStartPos = (pixelIdx * m_channels)
        m_bmp(pixelStartPos + 3) = m_bmp(pixelStartPos + 1)
        m_bmp(pixelStartPos + 2) = 0
        m_bmp(pixelStartPos + 1) = 0
    Next
    convert_rg_to_ra_rgba = True
End Function

Public Function convert_to_luma() As Boolean
    Dim originalBmpData() As Byte, numPixels As Long, pixelIdx As Long
    Dim red As Long, green As Long, blue As Long, lum As Long
    If m_channels = 1 Then convert_to_luma = True: Exit Function
    ' Otherwise, expecting 3 or 4-channel (RGB/RGBA) (but handle 2 channels by taking only the first channel)
    If Not (m_channels <= 4) Then ASSERT_FAILURE "m_channels <= 4", "lib/ivis_opengl/pieimage.cpp:306", "convert_to_luma", "Does not have <= 4 channels": convert_to_luma = False: Exit Function
    ByteArraySwap originalBmpData, m_bmp
    numPixels = m_height * m_width
    NewBuffer m_bmp, numPixels
    m_hasBmp = (numPixels > 0)
    If m_channels >= 3 Then
        For pixelIdx = 0 To numPixels - 1
            red = originalBmpData((pixelIdx * m_channels))
            green = originalBmpData((pixelIdx * m_channels) + 1)
            blue = originalBmpData((pixelIdx * m_channels) + 2)
            If red = green And red = blue Then
                ' all channels are the same - just use the first channel
                m_bmp(pixelIdx) = red
            Else
                ' quick approximation of a weighted RGB -> Luma method
                ' (R+R+B+G+G+G)/6
                lum = (red + red + blue + green + green + green) \ 6
                If lum > 255 Then lum = 255
                m_bmp(pixelIdx) = lum
            End If
        Next
    Else
        If Not (m_channels = 2) Then ASSERT_FAILURE "m_channels == 2", "lib/ivis_opengl/pieimage.cpp:333", "convert_to_luma", "Expecting 2 channels??"
        ' For two channels, just take the first, ignore the second
        internal_convert_image_contents 2, 1, -1, 255, m_width, m_height, originalBmpData, m_bmp
    End If
    m_channels = 1
    ' free the original bitmap data
    Erase originalBmpData
    convert_to_luma = True
End Function

' (VB) stb_image_resize's stbir_resize_uint8_generic, replaced: src (sw x sh, channels) into dst (dw x dh).
' A dimension that halves or stays is box filtered (the mip levels' 2x2, 2x1 or 1x2 average, rounded);
' any other size is sampled bilinearly at the pixel centres, the edges clamped.
Private Sub ResizePixels(src() As Byte, ByVal sw As Long, ByVal sh As Long, ByVal ch As Long, dst() As Byte, ByVal dw As Long, ByVal dh As Long)
    Dim x As Long, y As Long, c As Long, fx As Long, fy As Long, sum As Long, n As Long, i As Long, j As Long
    Dim o As Long, sx As Double, sy As Double, x0 As Long, y0 As Long, x1 As Long, y1 As Long, ax As Double, ay As Double
    Dim v As Double, r00 As Long, r01 As Long, r10 As Long, r11 As Long
    If (sw = dw * 2 Or sw = dw) And (sh = dh * 2 Or sh = dh) Then
        fx = sw \ dw
        fy = sh \ dh
        n = fx * fy
        For y = 0 To dh - 1
            For x = 0 To dw - 1
                o = (y * dw + x) * ch
                For c = 0 To ch - 1
                    sum = 0
                    For j = 0 To fy - 1
                        For i = 0 To fx - 1
                            sum = sum + src(((y * fy + j) * sw + (x * fx + i)) * ch + c)
                        Next
                    Next
                    dst(o + c) = (sum + n \ 2) \ n
                Next
            Next
        Next
        Exit Sub
    End If
    For y = 0 To dh - 1
        sy = (y + 0.5) * sh / dh - 0.5
        If sy < 0 Then sy = 0
        y0 = Int(sy)
        If y0 > sh - 1 Then y0 = sh - 1
        y1 = y0 + 1
        If y1 > sh - 1 Then y1 = sh - 1
        ay = sy - y0
        If ay < 0 Then ay = 0
        For x = 0 To dw - 1
            sx = (x + 0.5) * sw / dw - 0.5
            If sx < 0 Then sx = 0
            x0 = Int(sx)
            If x0 > sw - 1 Then x0 = sw - 1
            x1 = x0 + 1
            If x1 > sw - 1 Then x1 = sw - 1
            ax = sx - x0
            If ax < 0 Then ax = 0
            o = (y * dw + x) * ch
            For c = 0 To ch - 1
                r00 = src((y0 * sw + x0) * ch + c)
                r01 = src((y0 * sw + x1) * ch + c)
                r10 = src((y1 * sw + x0) * ch + c)
                r11 = src((y1 * sw + x1) * ch + c)
                v = (r00 * (1 - ax) + r01 * ax) * (1 - ay) + (r10 * (1 - ax) + r11 * ax) * ay
                v = Int(v + 0.5)
                If v > 255 Then v = 255
                If v < 0 Then v = 0
                dst(o + c) = CByte(v)
            Next
        Next
    Next
End Sub

' source: Nothing for *this (resize)
Private Function resizeInternal(ByVal source As iV_Image, ByVal output_w As Long, ByVal output_h As Long, ByVal has_alphaChannelOverride As Boolean, ByVal alphaChannelOverride As Long) As Boolean
    Dim srcBytes() As Byte, srcW As Long, srcH As Long, srcChannels As Long, srcOrder As ColorOrder
    Dim output_pixels() As Byte
    ' (VB) stb's filter (STBIR_FILTER_MITCHELL when both dimensions shrink, else STBIR_FILTER_DEFAULT), its alpha
    ' channel (3 for 4 channels, or alphaChannelOverride) and STBIR_FLAG_ALPHA_PREMULTIPLIED: the replacement
    ' filters every channel alike, as stb does with that flag
    If source Is Nothing Then
        srcW = m_width
        srcH = m_height
        srcChannels = m_channels
        srcOrder = m_colorOrder
        ByteArraySwap srcBytes, m_bmp
        m_hasBmp = False
    Else
        srcW = source.width()
        srcH = source.height()
        srcChannels = source.channels()
        srcOrder = source.colorOrder()
        If source.size_in_bytes() > 0 Then
            ReDim srcBytes(0 To source.size_in_bytes() - 1)
            CopyMemory srcBytes(0), ByVal source.bmp(), source.size_in_bytes()
        End If
    End If

    If Not NewBuffer(output_pixels, output_w * output_h * srcChannels) Then
        If source Is Nothing Then
            ByteArraySwap m_bmp, srcBytes
            m_hasBmp = True
        End If
        resizeInternal = False
        Exit Function
    End If
    If output_w > 0 And output_h > 0 And srcW > 0 And srcH > 0 And srcChannels > 0 Then
        ResizePixels srcBytes, srcW, srcH, srcChannels, output_pixels, output_w, output_h
    End If

    If m_hasBmp Then
        Erase m_bmp
    End If
    m_width = output_w
    m_height = output_h
    ByteArraySwap m_bmp, output_pixels
    m_hasBmp = (output_w * output_h * srcChannels > 0)
    m_channels = srcChannels
    m_colorOrder = srcOrder

    resizeInternal = True
End Function

Public Function resize(ByVal output_w As Long, ByVal output_h As Long, Optional ByVal has_alphaChannelOverride As Boolean = False, Optional ByVal alphaChannelOverride As Long = 0) As Boolean
    If output_w = m_width And output_h = m_height Then
        resize = True
        Exit Function
    End If

    resize = resizeInternal(Nothing, output_w, output_h, has_alphaChannelOverride, alphaChannelOverride)
End Function

Public Function resizedFromOther(ByVal other As iV_Image, ByVal output_w As Long, ByVal output_h As Long, Optional ByVal has_alphaChannelOverride As Boolean = False, Optional ByVal alphaChannelOverride As Long = 0) As Boolean
    If other Is Me Then ASSERT_FAILURE "&other != this", "lib/ivis_opengl/pieimage.cpp:396", "resizedFromOther", "Other is this": resizedFromOther = False: Exit Function
    resizedFromOther = resizeInternal(other, output_w, output_h, has_alphaChannelOverride, alphaChannelOverride)
End Function

' If resizing succeeds (or is not required), returns true. Returns false if resizing failed (memory allocation failure).
Public Function scale_image_max_size(ByVal maxWidth As Long, ByVal maxHeight As Long) As Boolean
    Dim scalingRatio As Double, widthRatio As Double, heightRatio As Double, output_w As Long, output_h As Long
    If (maxWidth <= 0 Or m_width <= maxWidth) And (maxHeight <= 0 Or m_height <= maxHeight) Then
        scale_image_max_size = True
        Exit Function
    End If

    widthRatio = CDbl(maxWidth) / CDbl(m_width)
    heightRatio = CDbl(maxHeight) / CDbl(m_height)
    If maxWidth > 0 And maxHeight > 0 Then
        If widthRatio < heightRatio Then scalingRatio = widthRatio Else scalingRatio = heightRatio
    Else
        If maxWidth > 0 Then scalingRatio = widthRatio Else scalingRatio = heightRatio
    End If

    output_w = Fix(m_width * scalingRatio)
    output_h = Fix(m_height * scalingRatio)

    scale_image_max_size = resize(output_w, output_h)
End Function

' const std::vector<unsigned int>& channelMap: channelMap(0 .. UBound)
Public Function convert_channels(channelMap() As Long) As Boolean
    Dim originalChannels As Long, newChannels As Long, originalBmpData() As Byte, numPixels As Long
    Dim pixelIdx As Long, i As Long, mapSize As Long, allBelow As Boolean
    mapSize = LongArraySize(channelMap)
    If Not (mapSize > 0) Then ASSERT_FAILURE "!channelMap.empty()", "lib/ivis_opengl/pieimage.cpp:428", "convert_channels", "No channels supplied to extract?": convert_channels = False: Exit Function
    If Not (mapSize <= 4) Then ASSERT_FAILURE "channelMap.size() <= 4", "lib/ivis_opengl/pieimage.cpp:429", "convert_channels", "iV_Image does not support > 4 channel textures (channelMap has %zu entries)", mapSize: convert_channels = False: Exit Function
    originalChannels = m_channels
    allBelow = True
    For i = 0 To mapSize - 1
        If Not (channelMap(LBound(channelMap) + i) >= 0 And channelMap(LBound(channelMap) + i) < originalChannels) Then allBelow = False
    Next
    If Not (allBelow) Then ASSERT_FAILURE "std::all_of(channelMap.begin(), channelMap.end(), [originalChannels](unsigned int srcChannel){ return srcChannel < originalChannels; })", "lib/ivis_opengl/pieimage.cpp:431", "convert_channels", "Channel swizzle contains channel > originalChannels (%u)", m_channels: convert_channels = False: Exit Function
    If Not (m_width > 0 And m_height > 0) Then ASSERT_FAILURE "m_width > 0 && m_height > 0", "lib/ivis_opengl/pieimage.cpp:432", "convert_channels", "No size": convert_channels = False: Exit Function

    newChannels = mapSize

    ByteArraySwap originalBmpData, m_bmp
    numPixels = m_height * m_width
    m_channels = newChannels
    NewBuffer m_bmp, m_height * m_width * newChannels
    For pixelIdx = 0 To numPixels - 1
        For i = 0 To newChannels - 1
            m_bmp((pixelIdx * m_channels) + i) = originalBmpData((pixelIdx * originalChannels) + channelMap(LBound(channelMap) + i))
        Next
    Next
    ' free the original bitmap data
    Erase originalBmpData
    convert_channels = True
End Function

Public Function convert_to_single_channel(Optional ByVal channel As Long = 0) As Boolean
    Dim originalBmpData() As Byte, originalChannels As Long, numPixels As Long, pixelIdx As Long
    If Not (channel < m_channels) Then ASSERT_FAILURE "channel < m_channels", "lib/ivis_opengl/pieimage.cpp:454", "convert_to_single_channel", "Cannot extract channel %u from image with %u channels", channel, m_channels: convert_to_single_channel = False: Exit Function
    If Not (m_width > 0 And m_height > 0) Then ASSERT_FAILURE "m_width > 0 && m_height > 0", "lib/ivis_opengl/pieimage.cpp:455", "convert_to_single_channel", "No size": convert_to_single_channel = False: Exit Function
    If channel = 0 And m_channels = 1 Then
        ' nothing to do
        convert_to_single_channel = True
        Exit Function
    End If

    ByteArraySwap originalBmpData, m_bmp
    originalChannels = m_channels
    numPixels = m_height * m_width
    ' copy just the desired channel over
    m_channels = 1
    If Not NewBuffer(m_bmp, numPixels) Then
        ByteArraySwap m_bmp, originalBmpData
        m_channels = originalChannels
        convert_to_single_channel = False
        Exit Function
    End If
    For pixelIdx = 0 To numPixels - 1
        m_bmp(pixelIdx) = originalBmpData((pixelIdx * originalChannels) + channel)
    Next
    ' free the original bitmap data
    Erase originalBmpData

    convert_to_single_channel = True
End Function

Public Function pad_image(ByVal newWidth As Long, ByVal newHeight As Long, ByVal useSmearing As Boolean) As Boolean
    Dim originalBmpData() As Byte, originalWidth As Long, originalHeight As Long, sizeOfBuffer As Long
    Dim x As Long, y As Long
    If newWidth = m_width And newHeight = m_height Then
        ' no padding needed
        pad_image = True
        Exit Function
    End If
    If newWidth < m_width Or newHeight < m_height Then
        pad_image = False
        Exit Function
    End If

    ByteArraySwap originalBmpData, m_bmp
    originalWidth = m_width
    originalHeight = m_height
    sizeOfBuffer = newWidth * newHeight * m_channels
    m_height = newHeight
    m_width = newWidth
    NewBuffer m_bmp, sizeOfBuffer
    m_hasBmp = (sizeOfBuffer > 0)
    If originalWidth > 0 And originalHeight > 0 Then
        For y = 0 To originalHeight - 1
            ' Copy the original image row
            CopyMemory m_bmp(newWidth * y * m_channels), originalBmpData(originalWidth * y * m_channels), (m_channels * originalWidth)
            If useSmearing Then
                ' Smear extra in x direction
                For x = originalWidth To newWidth - 1
                    CopyMemory m_bmp((newWidth * y + x) * m_channels), m_bmp((newWidth * y + x - 1) * m_channels), m_channels
                Next
            End If
        Next
        If useSmearing Then
            ' Smear extra in y direction
            For y = originalHeight To newHeight - 1
                ' Copy the entire prior row (above) in the y direction
                CopyMemory m_bmp(newWidth * y * m_channels), m_bmp((newWidth * (y - 1)) * m_channels), (m_channels * newWidth)
            Next
        End If
    End If
    ' free the original bitmap data
    Erase originalBmpData

    pad_image = True
End Function

Public Function blit_image(ByVal src As iV_Image, ByVal xOffset As Long, ByVal yOffset As Long) As Boolean
    Dim sourceImageHeight As Long, sourceImageWidth As Long, rowWidthToCopy As Long, rowsToCopy As Long
    Dim y As Long, copyOffset As Long, srcAddr As Long
    If Not (xOffset < m_width) Then ASSERT_FAILURE "xOffset < m_width", "lib/ivis_opengl/pieimage.cpp:536", "blit_image", "xOffset (%u) >= destination width (%u)", xOffset, m_width: blit_image = False: Exit Function
    If Not (yOffset < m_height) Then ASSERT_FAILURE "yOffset < m_height", "lib/ivis_opengl/pieimage.cpp:537", "blit_image", "yOffset (%u) >= destination height (%u)", yOffset, m_height: blit_image = False: Exit Function

    ' For now, the source and destination images must share the same format / # of channels
    If Not (src.channels() = channels() And src.pixel_format_() = pixel_format_()) Then ASSERT_FAILURE "src.channels() == channels() && src.pixel_format() == pixel_format()", "lib/ivis_opengl/pieimage.cpp:540", "blit_image", "Source and destination must share the same number of channels / format": blit_image = False: Exit Function

    sourceImageHeight = src.height()
    sourceImageWidth = src.width()

    rowWidthToCopy = sourceImageWidth
    If xOffset + sourceImageWidth > m_width Then
        debug_ 548, LOG_ERROR, "blit_image", "Cropping source image, as it would extend beyond the width of the destination image."
        rowWidthToCopy = m_width - xOffset
    End If
    rowsToCopy = sourceImageHeight
    If yOffset + sourceImageHeight > m_height Then
        debug_ 554, LOG_ERROR, "blit_image", "Cropping source image, as it would extend beyond the height of the destination image."
        rowsToCopy = m_height - yOffset
    End If

    srcAddr = src.bmp()
    For y = 0 To rowsToCopy - 1
        ' Copy the source image row
        copyOffset = (((m_width * (yOffset + y)) + xOffset) * 4)
        CopyMemory m_bmp(copyOffset), ByVal (srcAddr + sourceImageWidth * y * m_channels), (m_channels * rowWidthToCopy)
    Next

    blit_image = True
End Function

Public Function convert_color_order(ByVal newOrder As ColorOrder) As Boolean
    Dim numPixels As Long, pixelIdx As Long, t As Byte
    If m_colorOrder = newOrder Then
        convert_color_order = True
        Exit Function
    End If
    If Not (m_channels = 4) Then ASSERT_FAILURE "m_channels == 4", "lib/ivis_opengl/pieimage.cpp:574", "convert_color_order", "Does not have 4 channels - currently only supported for 4 channels": convert_color_order = False: Exit Function
    numPixels = m_height * m_width
    For pixelIdx = 0 To numPixels - 1
        t = m_bmp((pixelIdx * m_channels))
        m_bmp((pixelIdx * m_channels)) = m_bmp((pixelIdx * m_channels) + 2)
        m_bmp((pixelIdx * m_channels) + 2) = t
    Next
    m_colorOrder = newOrder
    convert_color_order = True
End Function

Public Function compare_equal(ByVal other As iV_Image) As Boolean
    Dim sizeOfBuffers As Long
    If m_width <> other.width() Or m_height <> other.height() Or m_channels <> other.channels() Then
        compare_equal = False
        Exit Function
    End If
    If m_colorOrder <> other.colorOrder() Then
        compare_equal = False
        Exit Function
    End If
    If (bmp() = 0 Or other.bmp() = 0) And (bmp() <> other.bmp()) Then
        compare_equal = False
        Exit Function
    End If

    sizeOfBuffers = m_width * m_height * m_channels
    If sizeOfBuffers = 0 Then
        compare_equal = True
        Exit Function
    End If
    compare_equal = (RtlCompareMemory(bmp(), other.bmp(), sizeOfBuffers) = sizeOfBuffers)
End Function

' (VB) iV_Image& operator=(iV_Image&& other): takes other's pixels, leaves it empty
Public Sub moveFrom(ByVal other As iV_Image)
    If Not other Is Me Then
        clear
        other.swapMembers m_bmp, m_hasBmp, m_width, m_height, m_channels, m_colorOrder
    End If
End Sub

' (VB) m_colorOrder, for another iV_Image (compare_equal, resizedFromOther)
Public Function colorOrder() As ColorOrder
    colorOrder = m_colorOrder
End Function

' (VB) std::swap of every member with another image's (the move constructor and assignment)
Friend Sub swapMembers(buf() As Byte, hasBuf As Boolean, w As Long, h As Long, c As Long, order As ColorOrder)
    Dim t As Long, tb As Boolean, tOrder As ColorOrder
    ByteArraySwap m_bmp, buf
    tb = m_hasBmp: m_hasBmp = hasBuf: hasBuf = tb
    t = m_width: m_width = w: w = t
    t = m_height: m_height = h: h = t
    t = m_channels: m_channels = c: c = t
    tOrder = m_colorOrder: m_colorOrder = order: order = tOrder
End Sub

' (VB) the buffer bmp_w() points at, lent as an array (swapped with buf; swap again to give it back): the KTX2
' transcoder writes the pixels into it (gfx_api_image_basis_priv_c)
Friend Sub swapBmpBuffer(buf() As Byte)
    ByteArraySwap m_bmp, buf
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
    iV_BaseImage_pixelFormat = pixel_format_()
End Function

Private Function iV_BaseImage_data() As Long
    iV_BaseImage_data = data()
End Function

Private Function iV_BaseImage_dataSize() As Long
    iV_BaseImage_dataSize = data_size()
End Function

Private Function iV_BaseImage_bufferRowLength() As Long
    iV_BaseImage_bufferRowLength = bufferRowLength()
End Function

Private Function iV_BaseImage_bufferImageHeight() As Long
    iV_BaseImage_bufferImageHeight = bufferImageHeight()
End Function

' (VB) the number of elements of a 1-D Long array (0 if it has none)
Private Function LongArraySize(a() As Long) As Long
    On Error GoTo none
    LongArraySize = UBound(a) - LBound(a) + 1
    Exit Function
none:
    LongArraySize = 0
End Function
