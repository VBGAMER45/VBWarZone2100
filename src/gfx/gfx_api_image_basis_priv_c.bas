Attribute VB_Name = "gfx_api_image_basis_priv_c"
' lib/ivis_opengl/gfx_api_image_basis_priv.cpp, gfx_api_image_basis_priv.h (BASIS_ENABLED, as the official
' builds): the KTX2 files' images and textures, through WzBasis's Basis Universal transcoder (BasisApi's
' functions, a ktx2_transcoder object for C++'s local basist::ktx2_transcoder). WzBasis is Basis Universal at
' the commit 4.7.0 pins (BASISD_LIB_VERSION 150: basis_is_format_supported takes transcoder.get_format()).
' optional<T> parameters and results are a has_ flag and the value; a std::vector of images is an array of
' iV_BaseImage (0-based) and a count; uint32_t maxWidth / maxHeight are Longs (UINT32_MAX is -1, compared
' unsigned). The transcoder writes into the images' own arrays (iV_Image.swapBmpBuffer,
' iV_CompressedImage.swapData lend them).
Option Explicit

' #define WZ_BASIS_UNCOMPRESSED_FORMAT gfx_api::pixel_format::FORMAT_RGBA8_UNORM_PACK8
Public Const WZ_BASIS_UNCOMPRESSED_FORMAT As Long = pixel_format_FORMAT_RGBA8_UNORM_PACK8

Private Const UINT32_MAX As Long = -1

' std::array<optional<gfx_api::pixel_format>, gfx_api::PIXEL_FORMAT_TARGET_COUNT> (PIXEL_FORMAT_TARGET_COUNT is 2;
' (VB) a third element, never set, answers depth_map, as gfx_api_c's real-time compression arrays)
Private has_bestAvailableBasisCompressionFormat_GameTextureRGBA(0 To 2) As Boolean
Private bestAvailableBasisCompressionFormat_GameTextureRGBA(0 To 2) As pixel_format
Private has_bestAvailableBasisCompressionFormat_GameTextureRGB(0 To 2) As Boolean
Private bestAvailableBasisCompressionFormat_GameTextureRGB(0 To 2) As pixel_format
Private has_bestAvailableBasisCompressionFormat_AlphaMask(0 To 2) As Boolean
Private bestAvailableBasisCompressionFormat_AlphaMask(0 To 2) As pixel_format
Private has_bestAvailableBasisCompressionFormat_NormalMap(0 To 2) As Boolean
Private bestAvailableBasisCompressionFormat_NormalMap(0 To 2) As pixel_format
Private has_bestAvailableBasisCompressionFormat_SpecularMap(0 To 2) As Boolean
Private bestAvailableBasisCompressionFormat_SpecularMap(0 To 2) As pixel_format
Private has_bestAvailableBasisCompressionFormat_HeightMap(0 To 2) As Boolean
Private bestAvailableBasisCompressionFormat_HeightMap(0 To 2) As pixel_format

Private basisInitialized As Boolean     ' initBasisTranscoder's static bool basisInitialized = false

' (VB) TEST-ONLY: WZ_TEST_KTX2_DUMP=<folder> (tools\ktx2_game_test.py's dump test) writes every level the transcoder
' outputs, before any channel conversion, to <folder>\<path with / as #>.<transcoder_texture_format>.L<level>, for a
' byte-for-byte comparison with tools\ktx2ref's output, and every image of a .ktx2 file handed to a texture upload
' (after the channel conversion) to <folder>\<path with / as #>.U<pixel_format>.<width>x<height>.L<level> (testDumpUpload).
' Not in 4.7.0; nothing is written without the variable.
Private ktx2DumpChecked As Boolean
Private ktx2DumpDir As String

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

' (VB) C++'s dynamic_cast<iV_Image*>(pBaseImage): Nothing when the image is not an iV_Image
Public Function asIV_Image(ByVal image As iV_BaseImage) As iV_Image
    If image Is Nothing Then Exit Function
    If TypeOf image Is iV_Image Then Set asIV_Image = image
End Function

' (VB) std::max(0, maxWidth) passed on as a uint32_t
Public Function maxZero(ByVal v As Long) As Long
    If v > 0 Then maxZero = v
End Function

' MARK: - Basis Universal transcoder implementation

' optional<basist::transcoder_texture_format>: has_value returned, the value in result
Private Function to_basis_format(ByVal desiredFormat As pixel_format, result As transcoder_texture_format) As Boolean
    Select Case desiredFormat
        ' UNCOMPRESSED FORMATS
    Case pixel_format_FORMAT_RGBA8_UNORM_PACK8
        result = cTFRGBA32
        to_basis_format = True
        Exit Function
    Case pixel_format_FORMAT_BGRA8_UNORM_PACK8, pixel_format_FORMAT_RGB8_UNORM_PACK8, pixel_format_FORMAT_RG8_UNORM, pixel_format_FORMAT_R8_UNORM
        ' not supported by basis transcoder - requires separate transforms
        to_basis_format = False
        Exit Function
        ' COMPRESSED FORMAT
    Case pixel_format_FORMAT_RGB_BC1_UNORM
        result = cTFBC1_RGB
        to_basis_format = True
        Exit Function
    Case pixel_format_FORMAT_RGBA_BC2_UNORM
        to_basis_format = False
        Exit Function
    Case pixel_format_FORMAT_RGBA_BC3_UNORM
        result = cTFBC3_RGBA
        to_basis_format = True
        Exit Function
    Case pixel_format_FORMAT_R_BC4_UNORM
        result = cTFBC4_R
        to_basis_format = True
        Exit Function
    Case pixel_format_FORMAT_RG_BC5_UNORM
        result = cTFBC5_RG
        to_basis_format = True
        Exit Function
    Case pixel_format_FORMAT_RGBA_BPTC_UNORM
        result = cTFBC7_RGBA
        to_basis_format = True
        Exit Function
    Case pixel_format_FORMAT_RGB8_ETC1
        result = cTFETC1_RGB
        to_basis_format = True
        Exit Function
    Case pixel_format_FORMAT_RGB8_ETC2
        to_basis_format = False
        Exit Function
    Case pixel_format_FORMAT_RGBA8_ETC2_EAC
        result = cTFETC2_RGBA
        to_basis_format = True
        Exit Function
    Case pixel_format_FORMAT_R11_EAC
        result = cTFETC2_EAC_R11
        to_basis_format = True
        Exit Function
    Case pixel_format_FORMAT_RG11_EAC
        result = cTFETC2_EAC_RG11
        to_basis_format = True
        Exit Function
    Case pixel_format_FORMAT_ASTC_4x4_UNORM
        result = cTFASTC_4x4_RGBA
        to_basis_format = True
        Exit Function
    Case Else
        debug_ 102, LOG_ERROR, "to_basis_format", "Unrecognised pixel format: %zu", CLng(desiredFormat)
    End Select

    to_basis_format = False
End Function

Private Function basisLibrarySupportsFormat(ByVal format As pixel_format) As Boolean
    Dim bFormat As transcoder_texture_format
    ' Check if it's supported by *at least* one of the basis_tex_formats
    If Not to_basis_format(format, bFormat) Then
        basisLibrarySupportsFormat = False
        Exit Function
    End If
    If basis_is_format_supported(bFormat, cUASTC4x4) Then
        basisLibrarySupportsFormat = True
    Else
        basisLibrarySupportsFormat = basis_is_format_supported(bFormat, cETC1S)
    End If
End Function

' (VB) the format's name for initBasisTranscoder's lines: format_to_str, or "<none>"
Private Function formatOrNone(ByVal has_format As Boolean, ByVal format As pixel_format) As String
    If has_format Then formatOrNone = gfx_api_format_to_str(format) Else formatOrNone = "<none>"
End Function

' void gfx_api::initBasisTranscoder()
Public Sub gfx_api_initBasisTranscoder()
    Dim target As Long, target_idx As Long, k As Long, format As pixel_format
    Dim qualityOrderRGBA(0 To 3) As pixel_format, qualityOrderRGB(0 To 1) As pixel_format, qualityOrderR(0 To 2) As pixel_format

    If Not basisInitialized Then
        basisu_transcoder_init
        basisInitialized = True
    End If

    For target = 0 To 1                         ' PIXEL_FORMAT_TARGET_COUNT
        has_bestAvailableBasisCompressionFormat_GameTextureRGBA(target) = False
        has_bestAvailableBasisCompressionFormat_GameTextureRGB(target) = False
        has_bestAvailableBasisCompressionFormat_AlphaMask(target) = False
        has_bestAvailableBasisCompressionFormat_NormalMap(target) = False
        has_bestAvailableBasisCompressionFormat_SpecularMap(target) = False
    Next target

    If Not get_wz_texture_compression() Then
        ' Texture compression is disabled - leave all bestAvailableBasisCompressionFormat_* variables as nullopt
        debug_ 141, LOG_3D, "initBasisTranscoder", "Basis texture compression formats: disabled"
        Exit Sub
    End If

    ' gfx_api::texture_type::game_texture: a RGB / RGBA texture, possibly stored in a compressed format
    ' Overall quality ranking:
    '   FORMAT_ASTC_4x4_UNORM > FORMAT_RGBA_BPTC_UNORM > FORMAT_RGBA8_ETC2_EAC (/ FORMAT_RGB8_ETC2) > FORMAT_RGBA_BC3_UNORM (DXT5) / FORMAT_RGB_BC1_UNORM (for RGB - 4bpp) > ETC1 (only RGB) > PVRTC (is generally the lowest quality)
    qualityOrderRGBA(0) = pixel_format_FORMAT_ASTC_4x4_UNORM ' NOTE: On OpenGL, only power-of-two ASTC textures are supported.
    qualityOrderRGBA(1) = pixel_format_FORMAT_RGBA_BPTC_UNORM
    qualityOrderRGBA(2) = pixel_format_FORMAT_RGBA8_ETC2_EAC
    qualityOrderRGBA(3) = pixel_format_FORMAT_RGBA_BC3_UNORM
    qualityOrderRGB(0) = pixel_format_FORMAT_RGB_BC1_UNORM
    qualityOrderRGB(1) = pixel_format_FORMAT_RGB8_ETC1

    For target_idx = 0 To 1                     ' PIXEL_FORMAT_TARGET_COUNT
        target = target_idx
        debug_ 158, LOG_3D, "initBasisTranscoder", "Basis target compression formats [%zu]", target_idx
        For k = 0 To 3
            format = qualityOrderRGBA(k)
            If Not basisLibrarySupportsFormat(format) Then
                GoTo continue_rgba
            End If
            If gfx_api_context_textureFormatIsSupported(target, format, pixel_format_usage_sampled_image) Then
                has_bestAvailableBasisCompressionFormat_GameTextureRGBA(target_idx) = True
                bestAvailableBasisCompressionFormat_GameTextureRGBA(target_idx) = format
                Exit For
            End If
continue_rgba:
        Next k
        debug_ 171, LOG_3D, "initBasisTranscoder", "  * GameTextureRGBA: %s", formatOrNone(has_bestAvailableBasisCompressionFormat_GameTextureRGBA(target_idx), bestAvailableBasisCompressionFormat_GameTextureRGBA(target_idx))
        For k = 0 To 1
            format = qualityOrderRGB(k)
            If Not basisLibrarySupportsFormat(format) Then
                GoTo continue_rgb
            End If
            If gfx_api_context_textureFormatIsSupported(target, format, pixel_format_usage_sampled_image) Then
                has_bestAvailableBasisCompressionFormat_GameTextureRGB(target_idx) = True
                bestAvailableBasisCompressionFormat_GameTextureRGB(target_idx) = format
                Exit For
            End If
continue_rgb:
        Next k
        debug_ 184, LOG_3D, "initBasisTranscoder", "  * GameTextureRGB: %s", formatOrNone(has_bestAvailableBasisCompressionFormat_GameTextureRGB(target_idx), bestAvailableBasisCompressionFormat_GameTextureRGB(target_idx))

        ' gfx_api::texture_type::alpha_mask:    // a single-channel texture, containing the alpha values
        ' gfx_api::texture_type::specular_map: // a single-channel texture, containing the specular / luma value
        ' gfx_api::texture_type::height_map:    // a single-channel texture, containing the height value
            ' Is is expected that for either of the above, the single-channel value is stored in the R (and possibly GB) channels
        ' Overall quality rankings: FORMAT_R11_EAC (4bpp) > FORMAT_R_BC4_UNORM (4bpp)
        qualityOrderR(0) = pixel_format_FORMAT_ASTC_4x4_UNORM
        qualityOrderR(1) = pixel_format_FORMAT_R11_EAC
        qualityOrderR(2) = pixel_format_FORMAT_R_BC4_UNORM
        For k = 0 To 2
            format = qualityOrderR(k)
            If Not basisLibrarySupportsFormat(format) Then
                GoTo continue_r
            End If
            If gfx_api_context_textureFormatIsSupported(target, format, pixel_format_usage_sampled_image) Then
                has_bestAvailableBasisCompressionFormat_AlphaMask(target_idx) = True
                bestAvailableBasisCompressionFormat_AlphaMask(target_idx) = format
                has_bestAvailableBasisCompressionFormat_SpecularMap(target_idx) = True
                bestAvailableBasisCompressionFormat_SpecularMap(target_idx) = format
                has_bestAvailableBasisCompressionFormat_HeightMap(target_idx) = True
                bestAvailableBasisCompressionFormat_HeightMap(target_idx) = format
                Exit For
            End If
continue_r:
        Next k
        debug_ 206, LOG_3D, "initBasisTranscoder", "  * AlphaMask: %s", formatOrNone(has_bestAvailableBasisCompressionFormat_AlphaMask(target_idx), bestAvailableBasisCompressionFormat_AlphaMask(target_idx))
        debug_ 207, LOG_3D, "initBasisTranscoder", "  * SpecularMap: %s", formatOrNone(has_bestAvailableBasisCompressionFormat_SpecularMap(target_idx), bestAvailableBasisCompressionFormat_SpecularMap(target_idx))
        debug_ 208, LOG_3D, "initBasisTranscoder", "  * HeightMap: %s", formatOrNone(has_bestAvailableBasisCompressionFormat_HeightMap(target_idx), bestAvailableBasisCompressionFormat_HeightMap(target_idx))
    Next target_idx
End Sub

' Determine the best available live compressed image format for the current system (+ textureType)
' (optional<gfx_api::pixel_format>: has_value returned, the value in result)
Private Function getBestAvailableTranscodeFormatForBasisFile(ByVal target As pixel_format_target, ByVal channel_id As ktx2_df_channel_id, ByVal dfd_transfer_func As Long, ByVal textureType As texture_type, result As pixel_format) As Boolean
    Dim target_idx As Long
    target_idx = target
    Select Case textureType
    Case texture_type_user_interface
        ' FUTURE TODO:
    Case texture_type_game_texture  ' a RGB / RGBA texture, generally
        If has_bestAvailableBasisCompressionFormat_GameTextureRGBA(target_idx) Then
            ' just use the best available RGBA format (if one is available)
            result = bestAvailableBasisCompressionFormat_GameTextureRGBA(target_idx)
            getBestAvailableTranscodeFormatForBasisFile = True
            Exit Function
        ElseIf has_bestAvailableBasisCompressionFormat_GameTextureRGB(target_idx) Then
            ' verify the input file is RGB
            Select Case channel_id
            Case KTX2_DF_CHANNEL_UASTC_RGB
                result = bestAvailableBasisCompressionFormat_GameTextureRGB(target_idx)
                getBestAvailableTranscodeFormatForBasisFile = True
                Exit Function
            Case Else
            End Select
        End If
    Case texture_type_alpha_mask    ' a single-channel texture, containing the alpha values
        result = bestAvailableBasisCompressionFormat_AlphaMask(target_idx)
        getBestAvailableTranscodeFormatForBasisFile = has_bestAvailableBasisCompressionFormat_AlphaMask(target_idx)
        Exit Function
    Case texture_type_normal_map
        result = bestAvailableBasisCompressionFormat_NormalMap(target_idx)
        getBestAvailableTranscodeFormatForBasisFile = has_bestAvailableBasisCompressionFormat_NormalMap(target_idx)
        Exit Function
    Case texture_type_specular_map  ' a single-channel texture, containing the specular / luma value
        result = bestAvailableBasisCompressionFormat_SpecularMap(target_idx)
        getBestAvailableTranscodeFormatForBasisFile = has_bestAvailableBasisCompressionFormat_SpecularMap(target_idx)
        Exit Function
    Case texture_type_height_map    ' a single-channel texture, containing the height values
        result = bestAvailableBasisCompressionFormat_HeightMap(target_idx)
        getBestAvailableTranscodeFormatForBasisFile = has_bestAvailableBasisCompressionFormat_HeightMap(target_idx)
        Exit Function
    Case Else
        ' unsupported
    End Select

    getBestAvailableTranscodeFormatForBasisFile = False
End Function

' optional<gfx_api::pixel_format> gfx_api::getBestAvailableTranscodeFormatForBasis(target, textureType): has_value
' returned, the value in result
Public Function gfx_api_getBestAvailableTranscodeFormatForBasis(ByVal target As pixel_format_target, ByVal textureType As texture_type, result As pixel_format) As Boolean
    Select Case textureType
    Case texture_type_user_interface
        ' FUTURE TODO:
    Case texture_type_game_texture  ' a RGB / RGBA texture, generally
        ' just use the best available RGBA format (if one is available)
        result = bestAvailableBasisCompressionFormat_GameTextureRGBA(target)
        gfx_api_getBestAvailableTranscodeFormatForBasis = has_bestAvailableBasisCompressionFormat_GameTextureRGBA(target)
        Exit Function
    Case texture_type_alpha_mask    ' a single-channel texture, containing the alpha values
        result = bestAvailableBasisCompressionFormat_AlphaMask(target)
        gfx_api_getBestAvailableTranscodeFormatForBasis = has_bestAvailableBasisCompressionFormat_AlphaMask(target)
        Exit Function
    Case texture_type_normal_map
        result = bestAvailableBasisCompressionFormat_NormalMap(target)
        gfx_api_getBestAvailableTranscodeFormatForBasis = has_bestAvailableBasisCompressionFormat_NormalMap(target)
        Exit Function
    Case texture_type_specular_map  ' a single-channel texture, containing the specular / luma value
        result = bestAvailableBasisCompressionFormat_SpecularMap(target)
        gfx_api_getBestAvailableTranscodeFormatForBasis = has_bestAvailableBasisCompressionFormat_SpecularMap(target)
        Exit Function
    Case texture_type_height_map    ' a single-channel texture, containing the height values
        result = bestAvailableBasisCompressionFormat_HeightMap(target)
        gfx_api_getBestAvailableTranscodeFormatForBasis = has_bestAvailableBasisCompressionFormat_HeightMap(target)
        Exit Function
    Case Else
        ' unsupported
    End Select

    gfx_api_getBestAvailableTranscodeFormatForBasis = False
End Function

Private Function iVImage_Basis_Convert_Channels(ByVal target As pixel_format_target, ByVal textureType As texture_type, ByVal uncompressedImage As iV_Image) As Boolean
    Dim channelMap(0 To 2) As Long
    ' Convert to expected # (and arrangement) of channels based on textureType
    Select Case textureType
    Case texture_type_specular_map, texture_type_alpha_mask, texture_type_height_map
        ' extract single channel (should always be in R)
        iVImage_Basis_Convert_Channels = uncompressedImage.convert_to_single_channel(0)
        Exit Function
    Case texture_type_normal_map
        ' TODO: the following must match how the build process encodes normal maps - currently, this assumes they are just encoded as RGBA (with no swizzling)
        If uncompressedImage.channels() > 3 Then
            If gfx_api_context_textureFormatIsSupported(target, pixel_format_FORMAT_RGB8_UNORM_PACK8, pixel_format_usage_sampled_image) Then
                channelMap(0) = 0
                channelMap(1) = 1
                channelMap(2) = 2
                iVImage_Basis_Convert_Channels = uncompressedImage.convert_channels(channelMap)
                Exit Function
            End If
        End If
        ' (falls through to default)
    Case Else
    End Select
    iVImage_Basis_Convert_Channels = False
End Function

Private Function checkFormatVersusMaxCompressionLevel_Basis(ByVal has_maxLevel As Boolean, ByVal maxLevel As max_texture_compression_level, ByVal desiredFormat As pixel_format) As Boolean
    If has_maxLevel Then
        Select Case maxLevel
        Case max_texture_compression_level_same_as_source
            checkFormatVersusMaxCompressionLevel_Basis = gfx_api_is_uncompressed_format(desiredFormat) Or desiredFormat = pixel_format_FORMAT_ASTC_4x4_UNORM
            Exit Function
        Case max_texture_compression_level_highest_quality
            checkFormatVersusMaxCompressionLevel_Basis = (desiredFormat = pixel_format_FORMAT_ASTC_4x4_UNORM Or desiredFormat = pixel_format_FORMAT_RGBA_BPTC_UNORM)
            Exit Function
        End Select
    End If

    checkFormatVersusMaxCompressionLevel_Basis = True
End Function

' (VB) TEST-ONLY: the WZ_TEST_KTX2_DUMP folder ("" without the variable; see ktx2DumpDir)
Private Function testDumpDir() As String
    If Not ktx2DumpChecked Then
        ktx2DumpChecked = True
        ktx2DumpDir = Environ$("WZ_TEST_KTX2_DUMP")
        If Len(ktx2DumpDir) > 0 Then debug_ 0, LOG_INFO, "testDumpDir", "(VB) TEST-ONLY: WZ_TEST_KTX2_DUMP is set: the transcoded KTX2 levels and the uploaded images are written to %s", ktx2DumpDir
    End If
    testDumpDir = ktx2DumpDir
End Function

' (VB) TEST-ONLY: the WZ_TEST_KTX2_DUMP file of a transcoded level (see ktx2DumpDir)
Private Sub testDumpLevel(ByVal filename As String, ByVal levelIndex As Long, ByVal format As transcoder_texture_format, buf() As Byte, ByVal size As Long)
    Dim path As String, f As Integer, out() As Byte
    If Len(testDumpDir()) = 0 Or size <= 0 Then Exit Sub
    path = ktx2DumpDir & "\" & Replace(filename, "/", "#") & "." & CLng(format) & ".L" & levelIndex
    On Error Resume Next
    Kill path
    On Error GoTo 0
    ReDim out(0 To size - 1)
    CopyMemory out(0), buf(0), size
    f = FreeFile
    Open path For Binary Access Write As #f
    Put #f, , out
    Close #f
End Sub

' (VB) TEST-ONLY: the WZ_TEST_KTX2_DUMP file of an image of a .ktx2 file handed to a texture upload, after the
' channel conversion (see ktx2DumpDir); level is the upload's mip level (a skipped level is not counted)
Public Sub testDumpUpload(ByVal filename As String, ByVal level As Long, ByVal image As iV_BaseImage)
    Dim path As String, f As Integer, out() As Byte, size As Long
    If Len(testDumpDir()) = 0 Or Not strEndsWith(filename, ".ktx2") Then Exit Sub
    If image Is Nothing Then Exit Sub
    size = image.dataSize()
    If size <= 0 Or image.data() = 0 Then Exit Sub
    path = ktx2DumpDir & "\" & Replace(filename, "/", "#") & ".U" & CLng(image.pixelFormat()) & "." & image.width() & "x" & image.height() & ".L" & level
    On Error Resume Next
    Kill path
    On Error GoTo 0
    ReDim out(0 To size - 1)
    CopyMemory out(0), ByVal image.data(), size
    f = FreeFile
    Open path For Binary Access Write As #f
    Put #f, , out
    Close #f
End Sub

' (VB) outputImages.push_back
Private Sub pushImage(images() As iV_BaseImage, n As Long, ByVal image As iV_BaseImage)
    If n = 0 Then ReDim images(0 To 0) Else ReDim Preserve images(0 To n)
    Set images(n) = image
    n = n + 1
End Sub

' Returns an iV_BaseImage for each mipLevel in a basis file, in the best possible format for the given texture_type (and input file) for the current system
' - game_texture: A compressed RGBA (or RGB) format, or an uncompressed format
' - alpha_mask: A compressed single-channel format, or R8 uncompressed
' - normal_map: TODO
' - specular_map: A compressed single-channel format, or R8 uncompressed
' Pass `nullopt` to desiredFormat to use the "best available" format for the current system
' (pData(0 .. dataSize - 1); the images in outputImages, their number returned: 0 = {})
Private Function loadiVImagesFromFile_Basis_Data(pData() As Byte, ByVal dataSize As Long, ByVal filename As String, ByVal textureType As texture_type, ByVal target As pixel_format_target, ByVal has_desiredFormat As Boolean, ByVal desiredFormat As pixel_format, ByVal maxWidth As Long, ByVal maxHeight As Long, ByVal has_maxMips As Boolean, ByVal maxMips As Long, outputImages() As iV_BaseImage) As Long
    Dim transcoder As ktx2_transcoder, format As transcoder_texture_format, internalFormat As pixel_format
    Dim has_maxLevel As Boolean, maxLevel As max_texture_compression_level, basisFormat As transcoder_texture_format
    Dim formatName As String, mipmapLevels As Long, width As Long, height As Long, layerIndex As Long, faceIndex As Long
    Dim levelIndex As Long, level_info As ktx2_image_level_info, n As Long
    Dim uncompressedOutput As iV_Image, compressedOutput As iV_CompressedImage
    Dim bytes_per_block_or_pixel As Long, numBlocksOrPixels As Long, outputSize As Long
    Dim decode_flags As Long, pOutput_blocks() As Byte, transcoded As Boolean

    Erase outputImages
    Set transcoder = New ktx2_transcoder
    If Not transcoder.init(pData, dataSize) Then
        debug_ 333, LOG_ERROR, "loadiVImagesFromFile_Basis_Data", "Failed to initialize ktx2_transcoder for file: %s", filename
        Exit Function
    End If

    If Not transcoder.is_uastc() Then
        debug_ 339, LOG_ERROR, "loadiVImagesFromFile_Basis_Data", "Unsupported: ktx2 file does not use UASTC supercompression format: %s", filename
        Exit Function
    End If

    format = cTFRGBA32                  ' standard full uncompressed format
    internalFormat = pixel_format_FORMAT_RGBA8_UNORM_PACK8
    If Not has_desiredFormat Then
        has_desiredFormat = getBestAvailableTranscodeFormatForBasisFile(target, transcoder.get_dfd_channel_id0(), transcoder.get_dfd_transfer_func(), textureType, desiredFormat)
    End If
    If has_desiredFormat Then
        If Not gfx_api_is_uncompressed_format(desiredFormat) Then
            has_maxLevel = gfx_api_getMaxTextureCompressionLevelOverride(filename, maxLevel)
            If checkFormatVersusMaxCompressionLevel_Basis(has_maxLevel, maxLevel, desiredFormat) Then
                If to_basis_format(desiredFormat, basisFormat) Then
                    format = basisFormat
                    internalFormat = desiredFormat
                Else
                    ASSERT_FAILURE "basisFormat.has_value()", "lib/ivis_opengl/gfx_api_image_basis_priv.cpp:361", "loadiVImagesFromFile_Basis_Data", "Failed to convert target internal format (%u) to basis format", CLng(desiredFormat)
                End If
            Else
                debug_ 366, LOG_WZ, "loadiVImagesFromFile_Basis_Data", "Texture compression override prevented transcoding to %s for file: %s", gfx_api_format_to_str(desiredFormat), filename
            End If
        End If
    End If

    ' (BASISD_LIB_VERSION 150: < 160)
    If Not basis_is_format_supported(format, transcoder.get_format()) Then
        formatName = basis_get_format_name(format)
        debug_ 377, LOG_ERROR, "loadiVImagesFromFile_Basis_Data", "Basis transcoder was not compiled with support for format: %s; failed to transcode file: %s", formatName, filename
        Exit Function
    End If

    mipmapLevels = transcoder.get_levels()
    width = transcoder.get_width()
    height = transcoder.get_height()
    debug_ 386, LOG_WZ, "loadiVImagesFromFile_Basis_Data", "%s: Width x Height: (%u x %u); Mipmap levels: %u", filename, width, height, mipmapLevels

    ' for now, these are both zero - if we need support for layers or faces in the future, we'll need some nested loops
    layerIndex = 0
    faceIndex = 0

    For levelIndex = 0 To transcoder.get_levels() - 1
        Set level_info = New ktx2_image_level_info
        If Not transcoder.get_image_level_info(level_info, levelIndex, layerIndex, faceIndex) Then
            debug_ 397, LOG_ERROR, "loadiVImagesFromFile_Basis_Data", "Failed retrieving image level info (%u %u %u) for: %s", levelIndex, layerIndex, faceIndex, filename
            Erase outputImages
            Exit Function
        End If

        If ULT(maxWidth, level_info.m_orig_width) Or ULT(maxHeight, level_info.m_orig_height) Then
            debug_ 403, LOG_WZ, "loadiVImagesFromFile_Basis_Data", "Skipping level [%u] (%u x %u) - exceeds maxWidth or maxHeight: %d x %d", levelIndex, level_info.m_orig_width, level_info.m_orig_height, maxWidth, maxHeight
            GoTo continue_
        End If

        Set uncompressedOutput = Nothing
        Set compressedOutput = Nothing
        bytes_per_block_or_pixel = basis_get_bytes_per_block_or_pixel(format)
        numBlocksOrPixels = 0
        If basis_transcoder_format_is_uncompressed(format) Then
            If Not (format = cTFRGBA32) Then ASSERT_FAILURE "format == basist::transcoder_texture_format::cTFRGBA32", "lib/ivis_opengl/gfx_api_image_basis_priv.cpp:413", "loadiVImagesFromFile_Basis_Data", "Unsupported uncompressed format: %u", CLng(format): Erase outputImages: Exit Function
            numBlocksOrPixels = level_info.m_orig_width * level_info.m_orig_height

            Set uncompressedOutput = New iV_Image
            If Not uncompressedOutput.allocate(level_info.m_orig_width, level_info.m_orig_height, 4, False) Then ' hard-coded for RGBA32 for now
                debug_ 419, LOG_ERROR, "loadiVImagesFromFile_Basis_Data", "Failed to allocate memory for uncompressed image buffer"
                Erase outputImages
                Exit Function
            End If
        Else
            numBlocksOrPixels = level_info.m_total_blocks
            outputSize = numBlocksOrPixels * bytes_per_block_or_pixel

            Set compressedOutput = New iV_CompressedImage
            If Not compressedOutput.allocate(internalFormat, outputSize, level_info.m_width, level_info.m_height, level_info.m_orig_width, level_info.m_orig_height, False) Then
                debug_ 431, LOG_ERROR, "loadiVImagesFromFile_Basis_Data", "Failed to allocate memory for buffer"
                Erase outputImages
                Exit Function
            End If
        End If

        decode_flags = 0
        ' pOutput_blocks: uncompressedOutput->bmp_w() or compressedOutput->uint64_w() ((VB) the image's array, lent)
        If Not uncompressedOutput Is Nothing Then
            uncompressedOutput.swapBmpBuffer pOutput_blocks
        Else
            compressedOutput.swapData pOutput_blocks
        End If
        transcoded = transcoder.transcode_image_level(levelIndex, layerIndex, faceIndex, pOutput_blocks, 0, numBlocksOrPixels, format, decode_flags)
        If transcoded Then testDumpLevel filename, levelIndex, format, pOutput_blocks, numBlocksOrPixels * bytes_per_block_or_pixel ' (VB) TEST-ONLY
        If Not uncompressedOutput Is Nothing Then
            uncompressedOutput.swapBmpBuffer pOutput_blocks
        Else
            compressedOutput.swapData pOutput_blocks
        End If
        If Not transcoded Then
            debug_ 448, LOG_ERROR, "loadiVImagesFromFile_Basis_Data", "Failed transcoding image level (%u %u %u) for: %s", levelIndex, layerIndex, faceIndex, filename
            Erase outputImages
            Exit Function
        End If

        If basis_transcoder_format_is_uncompressed(format) Then
            ' Convert to expected # (and arrangement) of channels based on textureType
            iVImage_Basis_Convert_Channels target, textureType, uncompressedOutput
            If has_desiredFormat Then
                If Not (uncompressedOutput.pixel_format_() = desiredFormat) Then ASSERT_FAILURE "uncompressedOutput->pixel_format() == desiredFormat.value()", "lib/ivis_opengl/gfx_api_image_basis_priv.cpp:458", "loadiVImagesFromFile_Basis_Data", "Input desiredFormat (%d) does not match converted output format (%d)", CLng(desiredFormat), CLng(uncompressedOutput.pixel_format_())
            End If

            pushImage outputImages, n, uncompressedOutput
        Else
            pushImage outputImages, n, compressedOutput
        End If

        If has_maxMips Then
            If n >= maxMips Then
                Exit For
            End If
        End If
continue_:
    Next levelIndex

    transcoder.clear

    loadiVImagesFromFile_Basis_Data = n
End Function

' Returns an iV_BaseImage for each mipLevel in a basis file, in the desiredFormat (if supported by the basis transcoder)
Private Function loadiVImagesFromFile_Basis_internal(ByVal filename As String, ByVal textureType As texture_type, ByVal target As pixel_format_target, ByVal has_desiredFormat As Boolean, ByVal desiredFormat As pixel_format, ByVal maxWidth As Long, ByVal maxHeight As Long, ByVal has_maxMips As Boolean, ByVal maxMips As Long, results() As iV_BaseImage) As Long
    Dim fp As Long, filesize As Double, buffer() As Byte, readSize As Double
    Erase results
    fp = PHYSFS_openRead(filename)
    If debugPartEnabled(LOG_3D) Then debug_ 483, LOG_3D, "loadiVImagesFromFile_Basis_internal", "Reading...[directory: %s] %s", PHYSFS_getRealDir(filename), filename
    If Not (fp <> 0) Then ASSERT_FAILURE "fp != nullptr", "lib/ivis_opengl/gfx_api_image_basis_priv.cpp:484", "loadiVImagesFromFile_Basis_internal", "Could not open %s", filename: Exit Function
    filesize = PHYSFS_fileLength(fp)
    If Not (filesize < 4294967295#) Then ASSERT_FAILURE "filesize < static_cast<PHYSFS_sint64>(std::numeric_limits<uint32_t>::max())", "lib/ivis_opengl/gfx_api_image_basis_priv.cpp:486", "loadiVImagesFromFile_Basis_internal", """%s"" filesize >= std::numeric_limits<uint32_t>::max()", filename: Exit Function
    If filesize > 0 Then ReDim buffer(0 To CLng(filesize) - 1)
    readSize = WZ_PHYSFS_readBytes(fp, buffer, 0, CLng(filesize))
    If readSize < filesize Then
        debug_ 492, LOG_FATAL, "loadiVImagesFromFile_Basis_internal", "WZ_PHYSFS_readBytes(%s) failed with error: %s" & vbLf, filename, WZ_PHYSFS_getLastError()
        PHYSFS_close fp
        Exit Function
    End If
    PHYSFS_close fp
    loadiVImagesFromFile_Basis_internal = loadiVImagesFromFile_Basis_Data(buffer, CLng(filesize), filename, textureType, target, has_desiredFormat, desiredFormat, maxWidth, maxHeight, has_maxMips, maxMips, results)
End Function

' Returns an iV_BaseImage for each mipLevel in a basis file, in the desiredFormat (if supported by the basis transcoder)
' (desiredFormat: has_desiredFormat False = nullopt; maxWidth / maxHeight: UINT32_MAX (-1) by default)
Public Function gfx_api_loadiVImagesFromFile_Basis(ByVal filename As String, ByVal textureType As texture_type, ByVal target As pixel_format_target, ByVal has_desiredFormat As Boolean, ByVal desiredFormat As pixel_format, ByVal maxWidth As Long, ByVal maxHeight As Long, results() As iV_BaseImage) As Long
    gfx_api_loadiVImagesFromFile_Basis = loadiVImagesFromFile_Basis_internal(filename, textureType, target, has_desiredFormat, desiredFormat, maxWidth, maxHeight, False, 0, results)
End Function

' gfx_api::texture* gfx_api::loadImageTextureFromFile_KTX2 (Nothing = nullptr)
Public Function gfx_api_loadImageTextureFromFile_KTX2(ByVal filename As String, ByVal textureType As texture_type, Optional ByVal maxWidth As Long = -1, Optional ByVal maxHeight As Long = -1) As gfx_api_texture
    Dim maxWidth_u32 As Long, maxHeight_u32 As Long, images() As iV_BaseImage, images_size As Long
    Dim mipmap_levels As Long, toplevel_width As Long, toplevel_height As Long, uploadFormat As pixel_format
    Dim pTexture As gfx_api_texture, i As Long, expected_width As Long, expected_height As Long, uploadResult As Boolean

    If maxWidth > 0 Then maxWidth_u32 = maxWidth Else maxWidth_u32 = UINT32_MAX
    If maxHeight > 0 Then maxHeight_u32 = maxHeight Else maxHeight_u32 = UINT32_MAX

    images_size = gfx_api_loadiVImagesFromFile_Basis(filename, textureType, pixel_format_target_texture_2d, False, pixel_format_invalid, maxWidth_u32, maxHeight_u32, images) ' nullopt: auto-detect best possible format
    If images_size = 0 Then
        ' Failed to load
        Exit Function
    End If

    ' Create a new compatible gpu texture object
    mipmap_levels = images_size
    toplevel_width = images(0).width()
    toplevel_height = images(0).height()
    uploadFormat = images(0).pixelFormat()
    Set pTexture = gfx_api_context_create_texture(mipmap_levels, toplevel_width, toplevel_width, uploadFormat, filename) ' (sic: the width twice, as 4.7.0)

    ' Upload image levels to texture
    For i = 0 To mipmap_levels - 1
        expected_width = SHRU(toplevel_width, i)
        If expected_width < 1 Then expected_width = 1
        expected_height = SHRU(toplevel_height, i)
        If expected_height < 1 Then expected_height = 1

        If Not (images(i).width() = expected_width And images(i).height() = expected_height) Then ASSERT_FAILURE "images[i]->width() == expected_width && images[i]->height() == expected_height", "lib/ivis_opengl/gfx_api_image_basis_priv.cpp:531", "loadImageTextureFromFile_KTX2", "Actual mip-level dimensions (%u x %u) do not match expected dimensions (%u x %u) for level[%zu] in image: %s", images(i).width(), images(i).height(), expected_width, expected_height, i, filename

        testDumpUpload filename, i, images(i)   ' (VB) TEST-ONLY
        uploadResult = pTexture.upload(i, images(i))
        If Not (uploadResult) Then ASSERT_FAILURE "uploadResult", "lib/ivis_opengl/gfx_api_image_basis_priv.cpp:534", "loadImageTextureFromFile_KTX2", "Failed to upload buffer to image": Exit Function
    Next i

    Set gfx_api_loadImageTextureFromFile_KTX2 = pTexture     ' pTexture.release()
End Function

' std::unique_ptr<iV_Image> gfx_api::loadUncompressedImageFromFile_KTX2 (Nothing = nullptr)
Public Function gfx_api_loadUncompressedImageFromFile_KTX2(ByVal filename As String, ByVal textureType As texture_type, ByVal target As pixel_format_target, Optional ByVal maxWidth As Long = -1, Optional ByVal maxHeight As Long = -1) As iV_Image
    Dim maxWidth_u32 As Long, maxHeight_u32 As Long, images() As iV_BaseImage, images_size As Long
    Dim pUncompressedImage As iV_Image

    If maxWidth > 0 Then maxWidth_u32 = maxWidth Else maxWidth_u32 = UINT32_MAX
    If maxHeight > 0 Then maxHeight_u32 = maxHeight Else maxHeight_u32 = UINT32_MAX

    images_size = loadiVImagesFromFile_Basis_internal(filename, textureType, target, True, WZ_BASIS_UNCOMPRESSED_FORMAT, maxWidth_u32, maxHeight_u32, True, 1, images)
    If images_size = 0 Then
        ' Failed to load
        Exit Function
    End If
    If Not (images_size = 1) Then ASSERT_FAILURE "images.size() == 1", "lib/ivis_opengl/gfx_api_image_basis_priv.cpp:551", "loadUncompressedImageFromFile_KTX2", "Should only be a single - top-level - mip level"

    ' Verify it's an iV_Image (uncompressed) - which it should be!
    Set pUncompressedImage = asIV_Image(images(0))
    If Not (Not pUncompressedImage Is Nothing) Then ASSERT_FAILURE "pUncompressedImage != nullptr", "lib/ivis_opengl/gfx_api_image_basis_priv.cpp:555", "loadUncompressedImageFromFile_KTX2", "Returned image is not uncompressed??"
    Set gfx_api_loadUncompressedImageFromFile_KTX2 = pUncompressedImage
End Function
