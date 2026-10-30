Attribute VB_Name = "gfx_api_c"
' lib/ivis_opengl/gfx_api.cpp: the context's creation (context::initialize,
' get, isInitialized), the per-texture compression overrides and the image
' and texture loaders; gfx_api_formats_def.h's is_uncompressed_format; and
' lib/ivis_opengl/gfx_api_image_compress_priv.cpp's real-time compression
' choices, as built without ETCPAK_ENABLED (no built-in compressor: every
' texture stays uncompressed; compressImage fails). gfx_api.cpp is ported as
' built without BASIS_ENABLED (there are no .ktx2 files in the data), except
' imageLoadFilenameFromInputFilename, ported as built with it (as for the
' official builds; the .png files take precedence, so the result is the same).
' context::get() is the backend chosen in current_backend_context
' (gfx_api_context_vb); the context's non-virtual members are
' gfx_api_context_<name>, the namespace's functions gfx_api_<name>.
' std::unique_ptr<iV_Image> / gfx_api::texture * results are objects (Nothing
' = nullptr); a std::vector of images is an array of iV_Image and a count.
' The GenerateDefaultTextureFunc and progress callbacks the game passes to
' loadTextureArrayFromFiles are numbers (GfxApi's GenerateDefaultTextureFunc,
' ProgressCallbackFunc): the lambdas of terrain.cpp and texture.cpp, run here.
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

Private Const wz_png_extension As String = ".png"

' gfx_api::max_texture_compression_level
Public Enum max_texture_compression_level
    max_texture_compression_level_same_as_source = 0
    max_texture_compression_level_highest_quality = 1
End Enum

Private backend As backend_type         ' = gfx_api::backend_type::opengl_backend
Public uses_gfx_debug As Boolean        ' = false
Private gfxApiReady As Boolean          ' (VB) the initialisers have run

' MARK: - Per-texture compression overrides
Private textureCompressionOverrides As Object ' std::unordered_map<std::string, max_texture_compression_level> (a Scripting.Dictionary)
Private has_textureCompressionOverrideFileRealDir As Boolean
Private textureCompressionOverrideFileRealDir As String

' gfx_api_image_compress_priv.cpp: std::array<optional<gfx_api::pixel_format>, PIXEL_FORMAT_TARGET_COUNT>
' (PIXEL_FORMAT_TARGET_COUNT is 2; (VB) a third element, never set, answers depth_map, which C++ would
' read past the end). builtInRealTimeFormatCompressors is empty (built without ETCPAK_ENABLED).
Private has_bestAvailableCompressionFormat_GameTextureRGBA(0 To 2) As Boolean
Private bestAvailableCompressionFormat_GameTextureRGBA(0 To 2) As pixel_format
Private has_bestAvailableCompressionFormat_GameTextureRGB(0 To 2) As Boolean
Private bestAvailableCompressionFormat_GameTextureRGB(0 To 2) As pixel_format

' the C++ initialisers
Private Sub InitGfxApi()
    If gfxApiReady Then Exit Sub
    gfxApiReady = True
    backend = backend_type_opengl_backend
    Set textureCompressionOverrides = CreateObject("Scripting.Dictionary")
    textureCompressionOverrides.CompareMode = vbBinaryCompare
End Sub

Private Function to_string(ByVal backendType As backend_type) As String
    Select Case backendType
    Case backend_type_null_backend: to_string = "Null backend": Exit Function
    Case backend_type_opengl_backend: to_string = "GL backend": Exit Function
    Case backend_type_vulkan_backend: to_string = "Vulkan backend": Exit Function
    End Select
    to_string = ""
End Function

' gfx_api::context::initialize. impl: the backend_Impl_Factory's product for backendType
' (an IGLBackend for opengl_backend, an INullBackend for null_backend)
Public Function gfx_api_context_initialize(ByVal impl As Object, ByVal antialiasing As Long, ByVal swapMode As swap_interval_mode, ByVal has_mipLodBias As Boolean, ByVal mipLodBias As Single, ByVal depthMapResolution As Long, ByVal backendType As backend_type) As Boolean
    Dim result As Boolean
    InitGfxApi
    If current_backend_context <> CONTEXT_NONE And backend = backendType Then
        ' ignore re-init for same backendType (for now)
        debug_ 49, LOG_ERROR, "initialize", "Attempt to re-initialize gfx_api::context for the same backend type - ignoring (for now)"
        gfx_api_context_initialize = True
        Exit Function
    End If
    backend = backendType
    If current_backend_context <> CONTEXT_NONE Then
        debug_ 55, LOG_FATAL, "initialize", "Attempt to reinitialize gfx_api::context for a new backend type once initialized - currently unsupported"
        gfx_api_context_initialize = False
        Exit Function
    End If
    Select Case backend
    Case backend_type_null_backend
        current_backend_context = CONTEXT_NULL    ' new null_context(uses_gfx_debug)
    Case backend_type_opengl_backend
        current_backend_context = CONTEXT_GL      ' new gl_context(uses_gfx_debug)
    Case backend_type_vulkan_backend
        debug_ 70, LOG_FATAL, "initialize", "Warzone 2100 was not compiled with support for the Vulkan backend."
        gfx_api_context_initialize = False
        Exit Function
    End Select
    If Not (current_backend_context <> CONTEXT_NONE) Then ASSERT_FAILURE "current_backend_context != nullptr", "lib/ivis_opengl/gfx_api.cpp:75", "initialize", "Failed to initialize gfx backend context": gfx_api_context_initialize = False: Exit Function
    result = gfx_api_context_initialize_(current_backend_context, impl, antialiasing, swapMode, has_mipLodBias, mipLodBias, depthMapResolution)
    If Not result Then
        debug_ 79, LOG_INFO, "initialize", "Failed to initialize gfx_api::context for: %s", to_string(backend)
        gfx_api_context_shutdown
        current_backend_context = CONTEXT_NONE    ' delete current_backend_context; = nullptr
        gfx_api_context_initialize = False
        Exit Function
    End If

    ' Calculate the best available run-time compression formats
    gfx_api_initBestRealTimeCompressionFormats

    gfx_api_context_initialize = result
End Function

' gfx_api::context::isInitialized
Public Function gfx_api_context_isInitialized() As Boolean
    gfx_api_context_isInitialized = (current_backend_context <> CONTEXT_NONE)
End Function

' MARK: - Per-texture compression overrides

Private Sub processTexCompressOverrideLine(ByVal line_ As String)
    Dim delimeterPos As Long, maxLevelStr As String, maxLevel As max_texture_compression_level
    If Len(line_) = 0 Then
        Exit Sub
    End If
    If Left$(line_, 1) = "#" Then
        Exit Sub                                ' ignore comment lines
    End If
    delimeterPos = InStrRev(line_, "=") - 1     ' line.rfind('=') (-1 = npos)
    If delimeterPos = -1 Then
        ' line is missing delimeter
        debug_ 127, LOG_WZ, "processTexCompressOverrideLine", "TexCompressOverride line is incorrectly formatted: %s", line_
        Exit Sub
    End If
    If delimeterPos = 0 Then
        ' line is missing filename
        debug_ 133, LOG_WZ, "processTexCompressOverrideLine", "TexCompressOverride line is incorrectly formatted: %s", line_
        Exit Sub
    End If
    If delimeterPos = Len(line_) - 1 Then
        ' line is missing parameter after delimeter
        debug_ 139, LOG_WZ, "processTexCompressOverrideLine", "TexCompressOverride line is incorrectly formatted: %s", line_
        Exit Sub
    End If

    maxLevelStr = Mid$(line_, delimeterPos + 2)
    If maxLevelStr = "0" Then
        maxLevel = max_texture_compression_level_same_as_source
    ElseIf maxLevelStr = "1" Then
        maxLevel = max_texture_compression_level_highest_quality
    Else
        debug_ 155, LOG_WZ, "processTexCompressOverrideLine", "TexCompressOverride line is incorrectly formatted (unknown value): %s", line_
        Exit Sub
    End If

    textureCompressionOverrides(Left$(line_, delimeterPos)) = maxLevel
End Sub

Public Function gfx_api_loadTextureCompressionOverrides() As Boolean
    Const fileName As String = "texpages/compression_overrides.txt"
    Dim has_newRealDirPath As Boolean, newRealDirPath As String
    Dim loadedTexConfigData() As Byte, loadedTexConfigData_size As Long, text As String, lines_() As String, i As Long
    InitGfxApi

    If PHYSFS_exists(fileName) <> 0 Then
        has_newRealDirPath = True
        newRealDirPath = PHYSFS_getRealDir(fileName)    ' WZ_PHYSFS_getRealDir_String
        If has_textureCompressionOverrideFileRealDir Then
            If newRealDirPath = textureCompressionOverrideFileRealDir Then
                ' no need to reload
                gfx_api_loadTextureCompressionOverrides = True
                Exit Function
            End If
        End If
    End If

    ' load the latest file
    has_textureCompressionOverrideFileRealDir = has_newRealDirPath
    textureCompressionOverrideFileRealDir = newRealDirPath
    textureCompressionOverrides.RemoveAll

    ' [FILE FORMAT]
    ' Comments: Line begins with #
    ' Each line is:
    ' <filename>=<max texture compression quality level>
    ' Where <filename> is the path inside the data directory (example: texpages/page-6.png), *case-sensitive*
    ' Where <max texture compression quality level> corresponds to the following table:
    '   0: same as source (for loading a PNG: no gpu texture compression, for loading a KTX2: ASTC4x4 or uncompressed)
    '   1: highest quality compressed texture formats only (ASTC4x4, BPTC - for PNGs, this will still currently lead to uncompressed textures in memory)
    ' To get the default behavior, which is to use the best texture compression format available on a particular system, omit listing the texture filename at all.
    ' (This file is intended only for overriding the default behavior.)
    ' NOTE: This is only intended for compatibility with some old manually-created pixel-art texture atlases that are tightly packed, low resolution, and may have "bleed" issues as a result of texture compression.
    ' In general, modern texture files with appropriate format, resolution, and padding (if needed between elements) should not need to be listed here, or should be reformatted to avoid it.

    debug_ 195, LOG_WZ, "loadTextureCompressionOverrides", "Loading texture compression config overrides: %s", fileName
    If Not loadFileToBufferVector(fileName, loadedTexConfigData, loadedTexConfigData_size, False, False) Then
        debug_ 198, LOG_WZ, "loadTextureCompressionOverrides", "No texture compression config override file: %s", fileName
        gfx_api_loadTextureCompressionOverrides = False
        Exit Function
    End If

    If loadedTexConfigData_size > 0 Then
        text = Left$(StrConv(loadedTexConfigData, vbUnicode), loadedTexConfigData_size)
    End If
    ' Remove all \r from the string
    text = Replace(text, vbCr, "")

    ' each line up to a \n, and the rest after the last one (an empty last line is ignored)
    If Len(text) > 0 Then
        lines_ = Split(text, vbLf)
        For i = 0 To UBound(lines_)
            processTexCompressOverrideLine lines_(i)
        Next i
    End If

    debug_ 221, LOG_WZ, "loadTextureCompressionOverrides", "Loaded %zu texture compression config overrides", textureCompressionOverrides.Count
    gfx_api_loadTextureCompressionOverrides = True
End Function

' optional<max_texture_compression_level>: has_value returned, the value in maxLevel
Public Function gfx_api_getMaxTextureCompressionLevelOverride(ByVal filename As String, maxLevel As max_texture_compression_level) As Boolean
    InitGfxApi
    If Not textureCompressionOverrides.Exists(filename) Then
        gfx_api_getMaxTextureCompressionLevelOverride = False
        Exit Function
    End If
    maxLevel = textureCompressionOverrides(filename)
    gfx_api_getMaxTextureCompressionLevelOverride = True
End Function

' MARK: - Image load helpers

Private Function loadImageTextureFromFile_PNG(ByVal filename As String, ByVal textureType As texture_type, ByVal maxWidth As Long, ByVal maxHeight As Long, ByVal quiet As Boolean) As gfx_api_texture
    Dim loadedUncompressedImage As iV_Image, forceRGB As Boolean
    Set loadedUncompressedImage = New iV_Image

    ' 1.) Load the PNG into an iV_Image
    forceRGB = (textureType = texture_type_game_texture) Or (textureType = texture_type_user_interface)
    If Not iV_loadImage_PNG2(filename, loadedUncompressedImage, forceRGB, quiet) Then
        ' Failed to load the image
        Exit Function
    End If

    Set loadImageTextureFromFile_PNG = gfx_api_context_loadTextureFromUncompressedImage(loadedUncompressedImage, textureType, filename, maxWidth, maxHeight)
End Function

Public Function gfx_api_imageLoadFilenameFromInputFilename(ByVal filename As String) As String
    Dim ktx2Filename As String
    ' For backwards-compatibility support, filenames that end in ".png" are used as "keys" and we first check for a .ktx2 file with the same filename
    If strEndsWith(filename, wz_png_extension) Then
        ' First, check if the .png file exists
        ' (.png files take precedence, so that existing mods work - and only have to include png files)
        If PHYSFS_exists(filename) <> 0 Then
            gfx_api_imageLoadFilenameFromInputFilename = filename
            Exit Function
        End If

        ' Check for presence of .ktx2 file
        ktx2Filename = Left$(filename, Len(filename) - Len(wz_png_extension))
        ktx2Filename = ktx2Filename & ".ktx2"
        If PHYSFS_exists(ktx2Filename) <> 0 Then
            gfx_api_imageLoadFilenameFromInputFilename = ktx2Filename
            Exit Function
        End If
        ' Fall-back to the .png file (which presumably doesn't exist because of the first check above, but this will cause it to be named as such in the later error message)
    End If
    gfx_api_imageLoadFilenameFromInputFilename = filename
End Function

Public Function gfx_api_checkImageFilesWouldLoadFromSameParentMountPath(filenames() As String, ByVal filenames_size As Long, ByVal ignoreNotFound As Boolean) As Boolean
    Dim imageLoadFilename1 As String, has_realDir1 As Boolean, realDir1 As String
    Dim i As Long, imageLoadFilename_other As String, realDir_other As String
    If filenames_size = 0 Then gfx_api_checkImageFilesWouldLoadFromSameParentMountPath = False: Exit Function
    If filenames_size = 1 Then gfx_api_checkImageFilesWouldLoadFromSameParentMountPath = True: Exit Function
    imageLoadFilename1 = gfx_api_imageLoadFilenameFromInputFilename(filenames(0))
    If PHYSFS_exists(imageLoadFilename1) <> 0 Then
        has_realDir1 = True
        realDir1 = PHYSFS_getRealDir(imageLoadFilename1)
    End If

    For i = 1 To filenames_size - 1
        imageLoadFilename_other = gfx_api_imageLoadFilenameFromInputFilename(filenames(i))
        If PHYSFS_exists(imageLoadFilename_other) = 0 Then
            If Not has_realDir1 Or ignoreNotFound Then
                GoTo continue_
            End If
            gfx_api_checkImageFilesWouldLoadFromSameParentMountPath = False
            Exit Function
        End If
        realDir_other = PHYSFS_getRealDir(imageLoadFilename_other)
        ' optional<std::string> != std::string: unequal when realDir1 has no value
        If Not has_realDir1 Then
            gfx_api_checkImageFilesWouldLoadFromSameParentMountPath = False
            Exit Function
        End If
        If realDir_other <> realDir1 Then
            gfx_api_checkImageFilesWouldLoadFromSameParentMountPath = False
            Exit Function
        End If
continue_:
    Next i
    gfx_api_checkImageFilesWouldLoadFromSameParentMountPath = True
End Function

' MARK: - High-level texture loading

' Load a texture from a file
' (which loads straight to a texture based on the appropriate texture_type, handling mip_maps, compression, etc)
Public Function gfx_api_context_loadTextureFromFile(ByVal filename As String, ByVal textureType As texture_type, Optional ByVal maxWidth As Long = -1, Optional ByVal maxHeight As Long = -1, Optional ByVal quiet As Boolean = False) As gfx_api_texture
    Dim imageLoadFilename As String
    imageLoadFilename = gfx_api_imageLoadFilenameFromInputFilename(filename)

    If strEndsWith(imageLoadFilename, ".png") Then
        Set gfx_api_context_loadTextureFromFile = loadImageTextureFromFile_PNG(imageLoadFilename, textureType, maxWidth, maxHeight, quiet)
    Else
        debug_ 336, LOG_ERROR, "loadTextureFromFile", "Unable to load image file: %s", filename
    End If
End Function

' 1 + floor(log2(n)) for n > 0, written with integers (log2 of a power of two is exact in C)
Private Function floorLog2(ByVal n As Long) As Long
    Dim k As Long
    Do While n >= 2
        n = n \ 2
        k = k + 1
    Loop
    floorLog2 = k
End Function

Private Function calcMipmapLevelsForUncompressedImage(ByVal image As iV_Image, ByVal textureType As texture_type) As Long
    Dim generateMipMaps As Boolean, mipmap_levels As Long, w As Long, h As Long
    ' 3.) Determine mipmap levels (if needed / desired)
    generateMipMaps = (textureType <> texture_type_user_interface)
    mipmap_levels = 1
    If generateMipMaps Then
        ' Calculate how many mip-map levels (with a target minimum level dimension of 4)
        w = image.width()
        h = image.height()
        If h > w Then w = h
        mipmap_levels = 1 + floorLog2(w)    ' 1 + static_cast<size_t>(floor(log2(std::max(image.width(), image.height()))))
    End If
    calcMipmapLevelsForUncompressedImage = mipmap_levels
End Function

Private Function uncompressedPNGImageConvertChannels(ByVal image As iV_Image, ByVal target As pixel_format_target, ByVal textureType As texture_type, ByVal filename As String) As Boolean
    Dim result As Boolean, channelMap(0 To 2) As Long
    ' 1.) Convert to expected # of channels based on textureType
    Select Case textureType
    Case texture_type_specular_map
        result = image.convert_to_luma()
        If Not (result) Then ASSERT_FAILURE "result", "lib/ivis_opengl/gfx_api.cpp:362", "uncompressedPNGImageConvertChannels", "(%s): Failed to convert specular map", filename: uncompressedPNGImageConvertChannels = False: Exit Function
    Case texture_type_alpha_mask
        If image.channels() > 1 Then
            If Not (image.channels() = 4) Then ASSERT_FAILURE "image.channels() == 4", "lib/ivis_opengl/gfx_api.cpp:369", "uncompressedPNGImageConvertChannels", "(%s): Alpha mask does not have 1 or 4 channels, as expected", filename: uncompressedPNGImageConvertChannels = False: Exit Function
            result = image.convert_to_single_channel(3) ' extract alpha channel
            If Not (result) Then ASSERT_FAILURE "result", "lib/ivis_opengl/gfx_api.cpp:371", "uncompressedPNGImageConvertChannels", "(%s): Failed extract alpha channel", filename: uncompressedPNGImageConvertChannels = False: Exit Function
        End If
    Case texture_type_height_map
        If image.channels() > 1 Then
            result = image.convert_to_single_channel(0) ' extract first channel
            If Not (result) Then ASSERT_FAILURE "result", "lib/ivis_opengl/gfx_api.cpp:380", "uncompressedPNGImageConvertChannels", "(%s): Failed to extract first channel", filename: uncompressedPNGImageConvertChannels = False: Exit Function
        End If
    Case texture_type_normal_map
        If image.channels() > 3 Then
            If gfx_api_context_textureFormatIsSupported(target, pixel_format_FORMAT_RGB8_UNORM_PACK8, pixel_format_usage_sampled_image) Then
                channelMap(0) = 0
                channelMap(1) = 1
                channelMap(2) = 2
                image.convert_channels channelMap ' convert to RGB (which is supported)
            End If
        End If
        ' (falls through to default)
    Case Else
    End Select

    uncompressedPNGImageConvertChannels = True
End Function

Private Function checkFormatVersusMaxCompressionLevel_FromUncompressed(ByVal has_maxLevel As Boolean, ByVal maxLevel As max_texture_compression_level, ByVal desiredFormat As pixel_format) As Boolean
    If has_maxLevel Then
        Select Case maxLevel
        Case max_texture_compression_level_same_as_source
            checkFormatVersusMaxCompressionLevel_FromUncompressed = gfx_api_is_uncompressed_format(desiredFormat)
            Exit Function
        Case max_texture_compression_level_highest_quality
            checkFormatVersusMaxCompressionLevel_FromUncompressed = gfx_api_is_uncompressed_format(desiredFormat) Or desiredFormat = pixel_format_FORMAT_ASTC_4x4_UNORM Or desiredFormat = pixel_format_FORMAT_RGBA_BPTC_UNORM
            Exit Function
        End Select
    End If

    checkFormatVersusMaxCompressionLevel_FromUncompressed = True
End Function

' std::vector<std::unique_ptr<iV_Image>>: the levels in results (0-based), their number returned
Private Function generateMipMapsFromUncompressedImage(ByVal image As iV_Image, ByVal mipmap_levels As Long, ByVal textureType As texture_type, results() As iV_Image) As Long
    Dim pLastImage As iV_Image, i As Long, n As Long
    Dim has_alphaChannelOverride As Boolean, alphaChannelOverride As Long
    Dim output_w As Long, output_h As Long, pNewLevel As iV_Image

    Set pLastImage = image
    For i = 1 To mipmap_levels - 1
        has_alphaChannelOverride = False
        If textureType = texture_type_alpha_mask Then
            has_alphaChannelOverride = True
            alphaChannelOverride = 0
        End If

        output_w = pLastImage.width() \ 2
        If output_w < 1 Then output_w = 1    ' std::max<unsigned int>(1, pLastImage->width() >> 1)
        output_h = pLastImage.height() \ 2
        If output_h < 1 Then output_h = 1

        Set pNewLevel = New iV_Image
        pNewLevel.resizedFromOther pLastImage, output_w, output_h, has_alphaChannelOverride, alphaChannelOverride

        ReDim Preserve results(0 To n)
        Set results(n) = pNewLevel
        n = n + 1

        Set pLastImage = pNewLevel
    Next i

    generateMipMapsFromUncompressedImage = n
End Function

' Takes an iv_Image and texture_type and loads a texture as appropriate / possible
' (image: the iV_Image&&; it is changed as the C++ changes it)
Public Function gfx_api_context_loadTextureFromUncompressedImage(ByVal image As iV_Image, ByVal textureType As texture_type, ByVal filename As String, Optional ByVal maxWidth As Long = -1, Optional ByVal maxHeight As Long = -1) As gfx_api_texture
    Dim imgScaleResult As Boolean, mipmap_levels As Long, channels As Long, closestSupportedChannels As Long, i As Long
    Dim expandResult As Boolean, uploadFormat As pixel_format, bestAvailableCompressedFormat As pixel_format
    Dim has_maxLevel As Boolean, maxLevel As max_texture_compression_level
    Dim pTexture As gfx_api_texture, uploadResult As Boolean, compressedImage As iV_Image
    Dim has_alphaChannelOverride As Boolean, alphaChannelOverride As Long
    Dim output_w As Long, output_h As Long, resizeResult As Boolean

    ' 1.) Convert to expected # of channels based on textureType
    If Not uncompressedPNGImageConvertChannels(image, pixel_format_target_texture_2d, textureType, filename) Then
        Exit Function
    End If

    ' 2.) If maxWidth / maxHeight exceed current image dimensions, resize()
    imgScaleResult = image.scale_image_max_size(maxWidth, maxHeight)
    If Not (imgScaleResult) Then ASSERT_FAILURE "imgScaleResult", "lib/ivis_opengl/gfx_api.cpp:455", "loadTextureFromUncompressedImage", "Failed to scale image to max size (%d x %d): %s", maxWidth, maxHeight, filename: Exit Function

    ' 3.) Determine mipmap levels (if needed / desired, based on textureType)
    mipmap_levels = calcMipmapLevelsForUncompressedImage(image, textureType)

    ' 4.) Extend channels, if needed, to a supported uncompressed format
    channels = image.channels()
    ' Verify that the gfx backend supports this format
    If Not (gfx_api_context_getClosestSupportedUncompressedImageFormatChannels(pixel_format_target_texture_2d, channels, closestSupportedChannels)) Then ASSERT_FAILURE "closestSupportedChannels.has_value()", "lib/ivis_opengl/gfx_api.cpp:464", "loadTextureFromUncompressedImage", "Exhausted all possible uncompressed formats??": Exit Function
    For i = image.channels() To closestSupportedChannels - 1
        expandResult = image.expand_channels_towards_rgba()
        If Not (expandResult) Then ASSERT_FAILURE "expandResult", "lib/ivis_opengl/gfx_api.cpp:468", "loadTextureFromUncompressedImage", "Failed to expand channels: %s", filename: Exit Function
    Next i

    uploadFormat = image.pixel_format_()
    If gfx_api_bestRealTimeCompressionFormatForImage(pixel_format_target_texture_2d, image, textureType, bestAvailableCompressedFormat) Then
        If bestAvailableCompressedFormat <> pixel_format_invalid Then
            has_maxLevel = gfx_api_getMaxTextureCompressionLevelOverride(filename, maxLevel)
            If checkFormatVersusMaxCompressionLevel_FromUncompressed(has_maxLevel, maxLevel, bestAvailableCompressedFormat) Then
                uploadFormat = bestAvailableCompressedFormat
            Else
                debug_ 481, LOG_WZ, "loadTextureFromUncompressedImage", "Texture compression override prevented compressing to %s for file: %s", gfx_api_format_to_str(bestAvailableCompressedFormat), filename
            End If
        End If
    End If

    ' 5.) Create a new compatible gpu texture object
    Set pTexture = gfx_api_context_create_texture(mipmap_levels, image.width(), image.height(), uploadFormat, filename)

    ' 6.) Upload initial (full) level
    If uploadFormat = image.pixel_format_() Then
        uploadResult = pTexture.upload(0, image)
        If Not (uploadResult) Then ASSERT_FAILURE "uploadResult", "lib/ivis_opengl/gfx_api.cpp:492", "loadTextureFromUncompressedImage", "Failed to upload buffer to image": Exit Function
    Else
        ' Run-time compression
        Set compressedImage = gfx_api_compressImage(image, uploadFormat)
        If Not (Not compressedImage Is Nothing) Then ASSERT_FAILURE "compressedImage != nullptr", "lib/ivis_opengl/gfx_api.cpp:498", "loadTextureFromUncompressedImage", "Failed to compress image to format: %zu", CLng(uploadFormat): Exit Function
        uploadResult = pTexture.upload(0, compressedImage)
        If Not (uploadResult) Then ASSERT_FAILURE "uploadResult", "lib/ivis_opengl/gfx_api.cpp:500", "loadTextureFromUncompressedImage", "Failed to upload buffer to image": Exit Function
    End If

    ' 7.) Generate and upload mipmaps (if needed)
    For i = 1 To mipmap_levels - 1
        has_alphaChannelOverride = False
        If textureType = texture_type_alpha_mask Then
            has_alphaChannelOverride = True
            alphaChannelOverride = 0
        End If

        output_w = image.width() \ 2
        If output_w < 1 Then output_w = 1    ' std::max<unsigned int>(1, image.width() >> 1)
        output_h = image.height() \ 2
        If output_h < 1 Then output_h = 1

        resizeResult = image.resize(output_w, output_h, has_alphaChannelOverride, alphaChannelOverride)
        If Not (resizeResult) Then ASSERT_FAILURE "resizeResult", "lib/ivis_opengl/gfx_api.cpp:516", "loadTextureFromUncompressedImage", "Failed to resize image mipmap [%zu] to output size (%u x %u): %s", i, output_w, output_h, filename: Exit Function

        If uploadFormat = image.pixel_format_() Then
            uploadResult = pTexture.upload(i, image)
            If Not (uploadResult) Then ASSERT_FAILURE "uploadResult", "lib/ivis_opengl/gfx_api.cpp:521", "loadTextureFromUncompressedImage", "Failed to upload buffer to image": Exit Function
        Else
            ' Run-time compression
            Set compressedImage = gfx_api_compressImage(image, uploadFormat)
            If Not (Not compressedImage Is Nothing) Then ASSERT_FAILURE "compressedImage != nullptr", "lib/ivis_opengl/gfx_api.cpp:527", "loadTextureFromUncompressedImage", "Failed to compress image to format: %zu", CLng(uploadFormat): Exit Function
            uploadResult = pTexture.upload(i, compressedImage)
            If Not (uploadResult) Then ASSERT_FAILURE "uploadResult", "lib/ivis_opengl/gfx_api.cpp:529", "loadTextureFromUncompressedImage", "Failed to upload buffer to image": Exit Function
        End If
    Next i

    Set gfx_api_context_loadTextureFromUncompressedImage = pTexture     ' pTexture.release()
End Function

' gfx_api::loadUncompressedImageFromFile (Nothing = nullptr)
Public Function gfx_api_loadUncompressedImageFromFile(ByVal filename As String, ByVal target As pixel_format_target, ByVal textureType As texture_type, Optional ByVal maxWidth As Long = -1, Optional ByVal maxHeight As Long = -1, Optional ByVal forceRGBA8 As Boolean = False) As iV_Image
    Dim imageLoadFilename As String, loadedUncompressedImage As iV_Image
    imageLoadFilename = gfx_api_imageLoadFilenameFromInputFilename(filename)

    If strEndsWith(imageLoadFilename, ".png") Then
        ' Load the PNG into an iV_Image
        Set loadedUncompressedImage = New iV_Image
        If Not iV_loadImage_PNG2(imageLoadFilename, loadedUncompressedImage, forceRGBA8) Then
            ' Failed to load the image
            Exit Function
        End If
        ' Convert to expected # of channels based on textureType (if forceRGBA8 is not set)
        If Not forceRGBA8 Then
            If Not uncompressedPNGImageConvertChannels(loadedUncompressedImage, target, textureType, imageLoadFilename) Then
                Exit Function
            End If
        End If
        ' If maxWidth / maxHeight exceed current image dimensions, resize()
        loadedUncompressedImage.scale_image_max_size maxWidth, maxHeight
    Else
        debug_ 576, LOG_ERROR, "loadUncompressedImageFromFile", "Unable to load image file: %s", filename
        Exit Function
    End If

    Set gfx_api_loadUncompressedImageFromFile = loadedUncompressedImage
End Function

' MARK: - texture

' optional<unsigned int>: has_value returned, the value in result
Public Function gfx_api_context_getClosestSupportedUncompressedImageFormatChannels(ByVal target As pixel_format_target, ByVal channels As Long, result As Long) As Boolean
    Dim format As pixel_format
    format = iV_Image_pixel_format_for_channels(channels)

    ' Verify that the gfx backend supports this format
    Do While Not gfx_api_context_textureFormatIsSupported(target, format, pixel_format_usage_sampled_image)
        If Not (channels < 4) Then ASSERT_FAILURE "channels < 4", "lib/ivis_opengl/gfx_api.cpp:592", "getClosestSupportedUncompressedImageFormatChannels", "Exhausted all possible uncompressed formats??": gfx_api_context_getClosestSupportedUncompressedImageFormatChannels = False: Exit Function
        channels = channels + 1
        format = iV_Image_pixel_format_for_channels(channels)
    Loop

    result = channels
    gfx_api_context_getClosestSupportedUncompressedImageFormatChannels = True
End Function

' Determine the best available uncompressed image format for the current system (+ textureType)
Public Function gfx_api_context_bestUncompressedPixelFormat(ByVal target As pixel_format_target, ByVal textureType As texture_type) As pixel_format
    Select Case textureType
    Case texture_type_user_interface, texture_type_game_texture ' a RGB / RGBA texture, possibly stored in a compressed format
        ' Since we don't have an image, to check whether it's actually RGB (no alpha), we have to err on the side of RGBA
        gfx_api_context_bestUncompressedPixelFormat = pixel_format_FORMAT_RGBA8_UNORM_PACK8
        Exit Function
    Case texture_type_specular_map, texture_type_alpha_mask, texture_type_height_map ' a single-channel texture, containing the alpha values
        gfx_api_context_bestUncompressedPixelFormat = pixel_format_FORMAT_R8_UNORM
        Exit Function
    Case texture_type_normal_map
        ' *MUST* check if 3-channel is supported (not guaranteed on all systems)
        If gfx_api_context_textureFormatIsSupported(target, pixel_format_FORMAT_RGB8_UNORM_PACK8, pixel_format_usage_sampled_image) Then
            gfx_api_context_bestUncompressedPixelFormat = pixel_format_FORMAT_RGB8_UNORM_PACK8
        Else
            gfx_api_context_bestUncompressedPixelFormat = pixel_format_FORMAT_RGBA8_UNORM_PACK8
        End If
        Exit Function
    End Select

    gfx_api_context_bestUncompressedPixelFormat = pixel_format_FORMAT_RGBA8_UNORM_PACK8
End Function

Public Function gfx_api_context_createTextureForCompatibleImageUploads(ByVal mipmap_count As Long, ByVal bitmap As iV_Image, ByVal filename As String) As gfx_api_texture
    Dim channels As Long, closestSupportedChannels As Long, target_pixel_format As pixel_format
    ' Get the channels of this iV_Image
    channels = bitmap.channels()
    ' Verify that the gfx backend supports this format
    If Not (gfx_api_context_getClosestSupportedUncompressedImageFormatChannels(pixel_format_target_texture_2d, channels, closestSupportedChannels)) Then ASSERT_FAILURE "closestSupportedChannels.has_value()", "lib/ivis_opengl/gfx_api.cpp:636", "createTextureForCompatibleImageUploads", "Exhausted all possible uncompressed formats??": Exit Function

    target_pixel_format = iV_Image_pixel_format_for_channels(closestSupportedChannels)

    Set gfx_api_context_createTextureForCompatibleImageUploads = gfx_api_context_create_texture(1, bitmap.width(), bitmap.height(), target_pixel_format, filename)
End Function

Public Function gfx_api_format_to_str(ByVal format As pixel_format) As String
    Select Case format
    ' UNCOMPRESSED FORMATS
    Case pixel_format_FORMAT_RGBA8_UNORM_PACK8: gfx_api_format_to_str = "RGBA8_UNORM"
    Case pixel_format_FORMAT_BGRA8_UNORM_PACK8: gfx_api_format_to_str = "BGRA8_UNORM"
    Case pixel_format_FORMAT_RGB8_UNORM_PACK8: gfx_api_format_to_str = "RGB8_UNORM"
    Case pixel_format_FORMAT_RG8_UNORM: gfx_api_format_to_str = "RG8_UNORM"
    Case pixel_format_FORMAT_R8_UNORM: gfx_api_format_to_str = "R8_UNORM"
    ' COMPRESSED FORMAT
    Case pixel_format_FORMAT_RGB_BC1_UNORM: gfx_api_format_to_str = "RGB_BC1_UNORM"
    Case pixel_format_FORMAT_RGBA_BC2_UNORM: gfx_api_format_to_str = "RGBA_BC2_UNORM"
    Case pixel_format_FORMAT_RGBA_BC3_UNORM: gfx_api_format_to_str = "RGBA_BC3_UNORM"
    Case pixel_format_FORMAT_R_BC4_UNORM: gfx_api_format_to_str = "R_BC4_UNORM"
    Case pixel_format_FORMAT_RG_BC5_UNORM: gfx_api_format_to_str = "RG_BC5_UNORM"
    Case pixel_format_FORMAT_RGBA_BPTC_UNORM: gfx_api_format_to_str = "RGBA_BPTC_UNORM"
    Case pixel_format_FORMAT_RGB8_ETC1: gfx_api_format_to_str = "RGB8_ETC1"
    Case pixel_format_FORMAT_RGB8_ETC2: gfx_api_format_to_str = "RGB8_ETC2"
    Case pixel_format_FORMAT_RGBA8_ETC2_EAC: gfx_api_format_to_str = "RGBA8_ETC2_EAC"
    Case pixel_format_FORMAT_R11_EAC: gfx_api_format_to_str = "R11_EAC"
    Case pixel_format_FORMAT_RG11_EAC: gfx_api_format_to_str = "RG11_EAC"
    Case pixel_format_FORMAT_ASTC_4x4_UNORM: gfx_api_format_to_str = "ASTC_4x4_UNORM"
    Case Else
        debug_ 668, LOG_ERROR, "format_to_str", "Unrecognised pixel format: %zu", CLng(format)
        gfx_api_format_to_str = ""
    End Select
End Function

Public Function gfx_api_format_channels(ByVal format As pixel_format) As Long
    Select Case format
    Case pixel_format_invalid
        gfx_api_format_channels = 1     ' just return 1 for now...
    ' UNCOMPRESSED FORMATS
    Case pixel_format_FORMAT_RGBA8_UNORM_PACK8, pixel_format_FORMAT_BGRA8_UNORM_PACK8
        gfx_api_format_channels = 4
    Case pixel_format_FORMAT_RGB8_UNORM_PACK8
        gfx_api_format_channels = 3
    Case pixel_format_FORMAT_RG8_UNORM
        gfx_api_format_channels = 2
    Case pixel_format_FORMAT_R8_UNORM
        gfx_api_format_channels = 1
    ' COMPRESSED FORMAT
    Case pixel_format_FORMAT_RGB_BC1_UNORM
        gfx_api_format_channels = 3
    Case pixel_format_FORMAT_RGBA_BC2_UNORM
        gfx_api_format_channels = 4
    Case pixel_format_FORMAT_RGBA_BC3_UNORM
        gfx_api_format_channels = 4
    Case pixel_format_FORMAT_R_BC4_UNORM
        gfx_api_format_channels = 1
    Case pixel_format_FORMAT_RG_BC5_UNORM
        gfx_api_format_channels = 2
    Case pixel_format_FORMAT_RGBA_BPTC_UNORM
        gfx_api_format_channels = 4
    Case pixel_format_FORMAT_RGB8_ETC1, pixel_format_FORMAT_RGB8_ETC2
        gfx_api_format_channels = 3
    Case pixel_format_FORMAT_RGBA8_ETC2_EAC
        gfx_api_format_channels = 4
    Case pixel_format_FORMAT_R11_EAC
        gfx_api_format_channels = 1
    Case pixel_format_FORMAT_RG11_EAC
        gfx_api_format_channels = 2
    Case pixel_format_FORMAT_ASTC_4x4_UNORM
        gfx_api_format_channels = 4     ' just return 4 for now...
    Case Else
        gfx_api_format_channels = 1     ' silence warning
    End Select
End Function

' template <size_t blockWidth, size_t blockHeight> calculate_astc_size
Private Function calculate_astc_size(ByVal blockWidth As Long, ByVal blockHeight As Long, ByVal width As Long, ByVal height As Long) As Long
    calculate_astc_size = ((width + blockWidth - 1) \ blockWidth) * ((height + blockHeight - 1) \ blockHeight) * 16
End Function

Public Function gfx_api_format_texel_block_width(ByVal format As pixel_format) As Long
    Select Case format
    Case pixel_format_invalid
        gfx_api_format_texel_block_width = 1    ' just return 1 for now...
    ' UNCOMPRESSED FORMATS
    Case pixel_format_FORMAT_RGBA8_UNORM_PACK8, pixel_format_FORMAT_BGRA8_UNORM_PACK8, pixel_format_FORMAT_RGB8_UNORM_PACK8, pixel_format_FORMAT_RG8_UNORM, pixel_format_FORMAT_R8_UNORM
        gfx_api_format_texel_block_width = 1
    ' COMPRESSED FORMAT
    Case pixel_format_FORMAT_RGB_BC1_UNORM, pixel_format_FORMAT_RGBA_BC2_UNORM, pixel_format_FORMAT_RGBA_BC3_UNORM, pixel_format_FORMAT_R_BC4_UNORM, pixel_format_FORMAT_RG_BC5_UNORM
        gfx_api_format_texel_block_width = 4
    Case pixel_format_FORMAT_RGBA_BPTC_UNORM
        gfx_api_format_texel_block_width = 4
    Case pixel_format_FORMAT_RGB8_ETC1, pixel_format_FORMAT_RGB8_ETC2, pixel_format_FORMAT_RGBA8_ETC2_EAC, pixel_format_FORMAT_R11_EAC, pixel_format_FORMAT_RG11_EAC
        gfx_api_format_texel_block_width = 4
    Case pixel_format_FORMAT_ASTC_4x4_UNORM
        gfx_api_format_texel_block_width = 4
    Case Else
        gfx_api_format_texel_block_width = 1    ' silence warning
    End Select
End Function

Public Function gfx_api_format_memory_size(ByVal format As pixel_format, ByVal width As Long, ByVal height As Long) As Long
    Select Case format
    Case pixel_format_invalid
        gfx_api_format_memory_size = 0
    ' [UNCOMPRESSED FORMATS]
    Case pixel_format_FORMAT_RGBA8_UNORM_PACK8, pixel_format_FORMAT_BGRA8_UNORM_PACK8
        gfx_api_format_memory_size = width * height * 4
    Case pixel_format_FORMAT_RGB8_UNORM_PACK8
        gfx_api_format_memory_size = width * height * 3
    Case pixel_format_FORMAT_RG8_UNORM
        gfx_api_format_memory_size = width * height * 2
    Case pixel_format_FORMAT_R8_UNORM
        gfx_api_format_memory_size = width * height
    ' [COMPRESSED FORMATS]
    ' BC / DXT formats
    ' 4x4 blocks, each block having a certain number of bytes
    Case pixel_format_FORMAT_RGB_BC1_UNORM
        gfx_api_format_memory_size = ((width + 3) \ 4) * ((height + 3) \ 4) * 8
    Case pixel_format_FORMAT_RGBA_BC2_UNORM, pixel_format_FORMAT_RGBA_BC3_UNORM
        gfx_api_format_memory_size = ((width + 3) \ 4) * ((height + 3) \ 4) * 16
    Case pixel_format_FORMAT_R_BC4_UNORM
        gfx_api_format_memory_size = ((width + 3) \ 4) * ((height + 3) \ 4) * 8
    Case pixel_format_FORMAT_RG_BC5_UNORM, pixel_format_FORMAT_RGBA_BPTC_UNORM
        gfx_api_format_memory_size = ((width + 3) \ 4) * ((height + 3) \ 4) * 16
    ' ETC
    ' Size calculations from: https://www.khronos.org/registry/OpenGL-Refpages/es3.0/html/glCompressedTexImage2D.xhtml
    Case pixel_format_FORMAT_RGB8_ETC1, pixel_format_FORMAT_RGB8_ETC2
        gfx_api_format_memory_size = ((width + 3) \ 4) * ((height + 3) \ 4) * 8
    Case pixel_format_FORMAT_RGBA8_ETC2_EAC
        gfx_api_format_memory_size = ((width + 3) \ 4) * ((height + 3) \ 4) * 16
    Case pixel_format_FORMAT_R11_EAC
        gfx_api_format_memory_size = ((width + 3) \ 4) * ((height + 3) \ 4) * 8
    Case pixel_format_FORMAT_RG11_EAC
        gfx_api_format_memory_size = ((width + 3) \ 4) * ((height + 3) \ 4) * 16
    ' ASTC
    Case pixel_format_FORMAT_ASTC_4x4_UNORM
        gfx_api_format_memory_size = calculate_astc_size(4, 4, width, height)
    ' no default case to ensure compiler error if more formats are added
    Case Else
        gfx_api_format_memory_size = 0  ' silence warning
    End Select
End Function

' gfx_api_formats_def.h: static inline bool is_uncompressed_format(const pixel_format& format)
Public Function gfx_api_is_uncompressed_format(ByVal format As pixel_format) As Boolean
    Select Case format
    Case pixel_format_FORMAT_RGBA8_UNORM_PACK8, pixel_format_FORMAT_BGRA8_UNORM_PACK8, pixel_format_FORMAT_RGB8_UNORM_PACK8, pixel_format_FORMAT_RG8_UNORM, pixel_format_FORMAT_R8_UNORM
        gfx_api_is_uncompressed_format = True
    Case Else
        gfx_api_is_uncompressed_format = False
    End Select
End Function

' texture arrays

' loads an image into a texture array, generating mip levels as needed (based on textureType)
Public Function gfx_api_context_loadTextureArrayLayerFromUncompressedImage(ByVal array_ As gfx_api_texture_array, ByVal layer As Long, ByVal image As iV_Image, ByVal textureType As texture_type, ByVal uploadFormat As pixel_format, ByVal filename As String, Optional ByVal maxWidth As Long = -1, Optional ByVal maxHeight As Long = -1) As Boolean
    Dim mipmap_levels As Long
    mipmap_levels = calcMipmapLevelsForUncompressedImage(image, textureType)
    gfx_api_context_loadTextureArrayLayerFromUncompressedImage = gfx_api_context_loadTextureArrayLayerFromUncompressedImage_mips(array_, layer, image, textureType, mipmap_levels, uploadFormat, filename, maxWidth, maxHeight)
End Function

' (the overload with mipmap_levels)
Public Function gfx_api_context_loadTextureArrayLayerFromUncompressedImage_mips(ByVal array_ As gfx_api_texture_array, ByVal layer As Long, ByVal rootImage As iV_Image, ByVal textureType As texture_type, ByVal mipmap_levels As Long, ByVal uploadFormat As pixel_format, ByVal filename As String, Optional ByVal maxWidth As Long = -1, Optional ByVal maxHeight As Long = -1) As Boolean
    Dim miplevels() As iV_Image, miplevels_size As Long, level As Long, image As iV_Image
    miplevels_size = generateMipMapsFromUncompressedImage(rootImage, mipmap_levels, textureType, miplevels)

    If Not uploadMipLevel(array_, layer, uploadFormat, rootImage, 0) Then
        Exit Function
    End If

    For level = 1 To miplevels_size
        Set image = miplevels(level - 1)
        If Not (Not image Is Nothing) Then ASSERT_FAILURE "image != nullptr", "lib/ivis_opengl/gfx_api.cpp:846", "loadTextureArrayLayerFromUncompressedImage", "Image wasn't an iV_Image?": Exit Function

        If Not uploadMipLevel(array_, layer, uploadFormat, image, level) Then
            Exit Function
        End If
    Next level

    gfx_api_context_loadTextureArrayLayerFromUncompressedImage_mips = True
End Function

' the lambda uploadMipLevel of loadTextureArrayLayerFromUncompressedImage
Private Function uploadMipLevel(ByVal array_ As gfx_api_texture_array, ByVal layer As Long, ByVal uploadFormat As pixel_format, ByVal image As iV_Image, ByVal level As Long) As Boolean
    Dim uploadResult As Boolean, compressedImage As iV_Image
    If uploadFormat = image.pixel_format_() Then
        uploadResult = array_.uploadLayer(layer, level, image)
        If Not (uploadResult) Then ASSERT_FAILURE "uploadResult", "lib/ivis_opengl/gfx_api.cpp:825", "loadTextureArrayLayerFromUncompressedImage", "Failed to upload buffer to image": Exit Function
    Else
        ' Run-time compression
        Set compressedImage = gfx_api_compressImage(image, uploadFormat)
        If Not (Not compressedImage Is Nothing) Then ASSERT_FAILURE "compressedImage != nullptr", "lib/ivis_opengl/gfx_api.cpp:831", "loadTextureArrayLayerFromUncompressedImage", "Failed to compress image to format: %zu", CLng(uploadFormat): Exit Function
        uploadResult = array_.uploadLayer(layer, level, compressedImage)
        If Not (uploadResult) Then ASSERT_FAILURE "uploadResult", "lib/ivis_opengl/gfx_api.cpp:833", "loadTextureArrayLayerFromUncompressedImage", "Failed to upload buffer to image": Exit Function
    End If
    uploadMipLevel = True
End Function

' used to load basis mip levels (already encoded in the desired target format) into a texture array
' (images: images_size levels, 0-based)
Public Function gfx_api_context_loadTextureArrayLayerFromBaseImages(ByVal array_ As gfx_api_texture_array, ByVal layer As Long, images() As iV_Image, ByVal images_size As Long, ByVal filename As String, Optional ByVal maxWidth As Long = -1, Optional ByVal maxHeight As Long = -1) As Boolean
    Dim level As Long, pMipLevelImage As iV_Image, uploadSuccess As Boolean
    For level = 0 To images_size - 1
        Set pMipLevelImage = images(level)
        If Not (Not pMipLevelImage Is Nothing) Then ASSERT_FAILURE "pMipLevelImage != nullptr", "lib/ivis_opengl/gfx_api.cpp:863", "loadTextureArrayLayerFromBaseImages", "Null image for level: %zu", level: Exit Function
        uploadSuccess = array_.uploadLayer(layer, level, pMipLevelImage)
        If Not (uploadSuccess) Then ASSERT_FAILURE "uploadSuccess", "lib/ivis_opengl/gfx_api.cpp:865", "loadTextureArrayLayerFromBaseImages", "Failed to upload image for level: %zu", level: Exit Function
    Next level

    gfx_api_context_loadTextureArrayLayerFromBaseImages = True
End Function

' std::vector<std::unique_ptr<iV_BaseImage>>: the images in results (0-based), their number returned
Private Function loadUncompressedImageWithMips(ByVal imageLoadFilename As String, ByVal textureType As texture_type, ByVal maxWidth As Long, ByVal maxHeight As Long, ByVal forceRGBA8 As Boolean, results() As iV_Image) As Long
    Dim pCurrentImage As iV_Image, mipmap_levels As Long, miplevels() As iV_Image, miplevels_size As Long, i As Long
    If strEndsWith(imageLoadFilename, ".png") Then
        Set pCurrentImage = gfx_api_loadUncompressedImageFromFile(imageLoadFilename, pixel_format_target_texture_2d_array, textureType, maxWidth, maxHeight, forceRGBA8)
        If pCurrentImage Is Nothing Then
            debug_ 902, LOG_ERROR, "loadUncompressedImageWithMips", "Unable to load image file: %s", imageLoadFilename
            loadUncompressedImageWithMips = 0
            Exit Function
        End If
        mipmap_levels = calcMipmapLevelsForUncompressedImage(pCurrentImage, textureType)
        miplevels_size = generateMipMapsFromUncompressedImage(pCurrentImage, mipmap_levels, textureType, miplevels)
        ReDim results(0 To miplevels_size)
        Set results(0) = pCurrentImage
        For i = 0 To miplevels_size - 1
            Set results(i + 1) = miplevels(i)
        Next i
        loadUncompressedImageWithMips = miplevels_size + 1
    Else
        debug_ 913, LOG_ERROR, "loadUncompressedImageWithMips", "Unable to load image file: %s", imageLoadFilename
        loadUncompressedImageWithMips = 0
    End If
End Function

' (VB) the GenerateDefaultTextureFunc lambdas the game passes (terrain.cpp, texture.cpp)
Private Function generateDefaultTexture(ByVal defaultTextureGenerator As GenerateDefaultTextureFunc, ByVal width As Long, ByVal height As Long, ByVal channels As Long) As iV_Image
    Dim pDefaultTexture As iV_Image, buf() As Byte, b As Long, n As Long
    Set pDefaultTexture = New iV_Image
    Select Case defaultTextureGenerator
    Case GenerateDefaultTexture_zeroed
        ' default specular map / height map: 0; the blank decal texture
        pDefaultTexture.allocate width, height, channels, True
    Case GenerateDefaultTexture_normalMap
        pDefaultTexture.allocate width, height, channels, True
        ' default normal map: (0,0,1)
        n = pDefaultTexture.data_size()
        If n > 0 Then
            ReDim buf(0 To n - 1)
            For b = 0 To n - 1
                buf(b) = &H7F                   ' memset(pBmpWrite, 0x7f, pDefaultNormalMap->data_size());
            Next b
            If channels >= 3 Then
                For b = 0 To n - 1 Step channels
                    buf(b + 2) = &HFF           ' blue=z
                Next b
            End If
            CopyMemory ByVal pDefaultTexture.bmp_w(), buf(0), n
        End If
    Case Else
        Exit Function
    End Select
    Set generateDefaultTexture = pDefaultTexture
End Function

' (VB) the progress callbacks the game passes: []() { resDoResLoadCallback(); }
Private Sub progressCallbackInvoke(ByVal progressCallback As ProgressCallbackFunc)
    Select Case progressCallback
    Case ProgressCallback_resDoResLoadCallback
        resDoResLoadCallback
    End Select
End Sub

' the lambda getDefaultTextureMipsP of loadTextureArrayFromFiles (False = nullptr; the
' images in defaultTextureMips)
Private Function getDefaultTextureMipsP(ByVal layer As Long, ByVal width As Long, ByVal height As Long, ByVal mipmap_levels As Long, ByVal format As pixel_format, defaultTextureMips() As iV_Image, defaultTextureMips_size As Long, ByVal defaultTextureGenerator As GenerateDefaultTextureFunc, ByVal textureType As texture_type, ByVal maxWidth As Long, ByVal maxHeight As Long) As Boolean
    Dim pCurrentImage As iV_Image, expectedMipLevels As Long, miplevels() As iV_Image, miplevels_size As Long, i As Long
    Dim w As Long, h As Long
    If defaultTextureMips_size = 0 Then
        If Not (defaultTextureGenerator <> GenerateDefaultTexture_none) Then ASSERT_FAILURE "defaultTextureGenerator != nullptr", "lib/ivis_opengl/gfx_api.cpp:981", "loadTextureArrayFromFiles", "Failed to generate default texture - no default texture generator provided": Exit Function
        If Not (gfx_api_is_uncompressed_format(format)) Then ASSERT_FAILURE "gfx_api::is_uncompressed_format(format)", "lib/ivis_opengl/gfx_api.cpp:982", "loadTextureArrayFromFiles", "Expected uncompressed extraction format, but received: %s", gfx_api_format_to_str(format): Exit Function
        If defaultTextureGenerator <> GenerateDefaultTexture_none Then
            If layer > 0 Then
                w = width
                h = height
            Else
                w = maxWidth
                h = maxHeight
            End If
            Set pCurrentImage = generateDefaultTexture(defaultTextureGenerator, w, h, gfx_api_format_channels(format))
            If Not (Not pCurrentImage Is Nothing) Then ASSERT_FAILURE "pCurrentImage != nullptr", "lib/ivis_opengl/gfx_api.cpp:986", "loadTextureArrayFromFiles", "Failed to generate default texture": Exit Function
            If layer > 0 Then
                If Not (pCurrentImage.width() = width And pCurrentImage.height() = height) Then ASSERT_FAILURE "pCurrentImage->width() == width && pCurrentImage->height() == height", "lib/ivis_opengl/gfx_api.cpp:989", "loadTextureArrayFromFiles", "Failed to generate matching default texture": Exit Function
            End If
            expectedMipLevels = calcMipmapLevelsForUncompressedImage(pCurrentImage, textureType)
            If layer > 0 Then
                If Not (expectedMipLevels = mipmap_levels) Then ASSERT_FAILURE "expectedMipLevels == mipmap_levels", "lib/ivis_opengl/gfx_api.cpp:994", "loadTextureArrayFromFiles", "Failed to generate matching default texture": Exit Function
            End If
            miplevels_size = generateMipMapsFromUncompressedImage(pCurrentImage, expectedMipLevels, textureType, miplevels)
            ReDim defaultTextureMips(0 To miplevels_size)
            Set defaultTextureMips(0) = pCurrentImage
            For i = 0 To miplevels_size - 1
                Set defaultTextureMips(i + 1) = miplevels(i)
            Next i
            defaultTextureMips_size = miplevels_size + 1
        End If
    End If
    getDefaultTextureMipsP = True
End Function

' filenames: filenames_size names (0-based; "" = a layer of the default texture)
Public Function gfx_api_context_loadTextureArrayFromFiles(filenames() As String, ByVal filenames_size As Long, ByVal textureType As texture_type, Optional ByVal maxWidth As Long = -1, Optional ByVal maxHeight As Long = -1, Optional ByVal defaultTextureGenerator As GenerateDefaultTextureFunc = GenerateDefaultTexture_none, Optional ByVal progressCallback As ProgressCallbackFunc = ProgressCallback_none, Optional ByVal debugName As String = "") As gfx_api_texture_array
    Dim imageLoadFilenames() As String, i As Long
    Dim allFilesAreKTX2 As Boolean, anyFileIsKTX2 As Boolean, anyFileIsPNG As Boolean
    Dim uploadFormat As pixel_format, desiredImageExtractionFormat As pixel_format, bestAvailableCompressedFormat As pixel_format
    Dim uncompressedExtractionFormat As Boolean
    Dim defaultTextureMips() As iV_Image, defaultTextureMips_size As Long
    Dim texture_array As gfx_api_texture_array, width As Long, height As Long, mipmap_levels As Long, layers_count As Long
    Dim layer As Long, imageLoadFilename As String
    Dim loadedImagesForLayer() As iV_Image, loadedImagesForLayer_size As Long
    Dim pImagesForLayer() As iV_Image, pImagesForLayer_size As Long, has_pImagesForLayer As Boolean
    Dim uploadSuccess As Boolean, level As Long, image As iV_Image, compressedImage As iV_Image, uploadResult As Boolean

    If Not (filenames_size <= gfx_api_context_get_context_value(context_value_MAX_ARRAY_TEXTURE_LAYERS)) Then ASSERT_FAILURE "filenames.size() <= gfx_api::context::get().get_context_value(gfx_api::context::context_value::MAX_ARRAY_TEXTURE_LAYERS)", "lib/ivis_opengl/gfx_api.cpp:921", "loadTextureArrayFromFiles", "Too many layers": Exit Function

    If filenames_size > 0 Then ReDim imageLoadFilenames(0 To filenames_size - 1)
    For i = 0 To filenames_size - 1
        imageLoadFilenames(i) = gfx_api_imageLoadFilenameFromInputFilename(filenames(i))
    Next i

    allFilesAreKTX2 = True                      ' std::all_of
    For i = 0 To filenames_size - 1
        If Not strEndsWith(imageLoadFilenames(i), ".ktx2") Then allFilesAreKTX2 = False: Exit For
    Next i
    anyFileIsKTX2 = allFilesAreKTX2
    If Not anyFileIsKTX2 Then
        For i = 0 To filenames_size - 1
            If strEndsWith(imageLoadFilenames(i), ".ktx2") Then anyFileIsKTX2 = True: Exit For
        Next i
    End If
    If Not allFilesAreKTX2 Then
        For i = 0 To filenames_size - 1
            If strEndsWith(imageLoadFilenames(i), ".png") Then anyFileIsPNG = True: Exit For
        Next i
    End If

    ' (#if !defined(BASIS_ENABLED))
    If anyFileIsKTX2 Then
        For i = 0 To filenames_size - 1
            If strEndsWith(imageLoadFilenames(i), ".ktx2") Then
                debug_ 940, LOG_ERROR, "loadTextureArrayFromFiles", "Unable to load image file: %s", imageLoadFilenames(i)
            End If
        Next i
        Exit Function
    End If

    uploadFormat = gfx_api_context_bestUncompressedPixelFormat(pixel_format_target_texture_2d_array, textureType)
    desiredImageExtractionFormat = uploadFormat
    ' just use the best runtime compression format
    If gfx_api_bestRealTimeCompressionFormat(pixel_format_target_texture_2d_array, textureType, bestAvailableCompressedFormat) Then
        If bestAvailableCompressedFormat <> pixel_format_invalid Then
            uploadFormat = bestAvailableCompressedFormat
            desiredImageExtractionFormat = pixel_format_FORMAT_RGBA8_UNORM_PACK8 ' must extract to RGBA for run-time compression
        End If
    End If
    If anyFileIsKTX2 And anyFileIsPNG Then
        debug_ 971, LOG_INFO, "loadTextureArrayFromFiles", "Performance info: Some files are .ktx2, but others are .png"
    End If

    uncompressedExtractionFormat = gfx_api_is_uncompressed_format(desiredImageExtractionFormat)

    width = 0
    height = 0
    mipmap_levels = 0
    layers_count = filenames_size

    For layer = 0 To layers_count - 1
        imageLoadFilename = imageLoadFilenames(layer)

        ' load the file to an array of base images
        loadedImagesForLayer_size = 0
        has_pImagesForLayer = False
        If Len(imageLoadFilename) = 0 Then
            has_pImagesForLayer = getDefaultTextureMipsP(layer, width, height, mipmap_levels, desiredImageExtractionFormat, defaultTextureMips, defaultTextureMips_size, defaultTextureGenerator, textureType, maxWidth, maxHeight)
            If Not (has_pImagesForLayer) Then ASSERT_FAILURE "pImagesForLayer != nullptr", "lib/ivis_opengl/gfx_api.cpp:1021", "loadTextureArrayFromFiles", "Failed to generate matching default texture": Exit Function
            If defaultTextureMips_size > 0 Then pImagesForLayer = defaultTextureMips
            pImagesForLayer_size = defaultTextureMips_size
        ElseIf uncompressedExtractionFormat Or strEndsWith(imageLoadFilename, ".png") Then
            ' load into an uncompressed format
            loadedImagesForLayer_size = loadUncompressedImageWithMips(imageLoadFilename, textureType, maxWidth, maxHeight, desiredImageExtractionFormat = pixel_format_FORMAT_RGBA8_UNORM_PACK8, loadedImagesForLayer)
            If loadedImagesForLayer_size > 0 Then
                pImagesForLayer = loadedImagesForLayer
                pImagesForLayer_size = loadedImagesForLayer_size
                has_pImagesForLayer = True
            Else
                ' failed to load image
                debug_ 1034, LOG_INFO, "loadTextureArrayFromFiles", "Using default texture generator for failed image: %s", imageLoadFilename
                has_pImagesForLayer = getDefaultTextureMipsP(layer, width, height, mipmap_levels, desiredImageExtractionFormat, defaultTextureMips, defaultTextureMips_size, defaultTextureGenerator, textureType, maxWidth, maxHeight)
                If has_pImagesForLayer And defaultTextureMips_size > 0 Then pImagesForLayer = defaultTextureMips
                pImagesForLayer_size = defaultTextureMips_size
            End If
        Else
            ' load directly into a compressed format
            debug_ 1050, LOG_ERROR, "loadTextureArrayFromFiles", "Unable to load image file: %s", imageLoadFilename
            Exit Function
        End If

        progressCallbackInvoke progressCallback

        If Not (has_pImagesForLayer And pImagesForLayer_size > 0) Then ASSERT_FAILURE "pImagesForLayer && !pImagesForLayer->empty()", "lib/ivis_opengl/gfx_api.cpp:1060", "loadTextureArrayFromFiles", "Unable to load images: %s", imageLoadFilename: Exit Function

        If layer = 0 Then
            width = pImagesForLayer(0).width()
            height = pImagesForLayer(0).height()
            mipmap_levels = pImagesForLayer_size

            Set texture_array = gfx_api_context_create_texture_array(mipmap_levels, layers_count, width, height, uploadFormat, debugName)
            If Not (Not texture_array Is Nothing) Then ASSERT_FAILURE "texture_array.get() != nullptr", "lib/ivis_opengl/gfx_api.cpp:1069", "loadTextureArrayFromFiles", "Failed to create texture array (miplevels: %zu, layers: %zu, width: %u, height: %u): %s", mipmap_levels, layers_count, width, height, imageLoadFilename: Exit Function
        Else
            If Not (width = pImagesForLayer(0).width() And height = pImagesForLayer(0).height()) Then ASSERT_FAILURE "width == pImagesForLayer->front()->width() && height == pImagesForLayer->front()->height()", "lib/ivis_opengl/gfx_api.cpp:1073", "loadTextureArrayFromFiles", "Unexpected image dimensions (%u x %u) does not match the first image dimensions (%u x %u): %s", pImagesForLayer(0).width(), pImagesForLayer(0).height(), width, height, imageLoadFilename: Exit Function
            If Not (pImagesForLayer_size = mipmap_levels) Then ASSERT_FAILURE "pImagesForLayer->size() == mipmap_levels", "lib/ivis_opengl/gfx_api.cpp:1074", "loadTextureArrayFromFiles", "Unexpected number of mip levels (%zu; expected: %zu): %s", pImagesForLayer_size, mipmap_levels, imageLoadFilename: Exit Function
        End If

        ' upload the layer

        ' If already in the uploadFormat
        If uploadFormat = desiredImageExtractionFormat Then
            ' just load directly
            uploadSuccess = gfx_api_context_loadTextureArrayLayerFromBaseImages(texture_array, layer, pImagesForLayer, pImagesForLayer_size, imageLoadFilename, width, height)
            If Not (uploadSuccess) Then ASSERT_FAILURE "uploadSuccess", "lib/ivis_opengl/gfx_api.cpp:1084", "loadTextureArrayFromFiles", "Failed to loadTextureArrayLayerFromBaseImages": Exit Function
        Else
            ' convert from (presumably uncompressed) to desired run-time compressed target format (for each mip level)
            If Not (uncompressedExtractionFormat) Then ASSERT_FAILURE "uncompressedExtractionFormat", "lib/ivis_opengl/gfx_api.cpp:1089", "loadTextureArrayFromFiles", "Expected uncompressed extraction format, but received: %s", gfx_api_format_to_str(desiredImageExtractionFormat): Exit Function

            For level = 0 To pImagesForLayer_size - 1
                Set image = pImagesForLayer(level)
                If Not (Not image Is Nothing) Then ASSERT_FAILURE "image != nullptr", "lib/ivis_opengl/gfx_api.cpp:1094", "loadTextureArrayFromFiles", "Image wasn't an iV_Image?": Exit Function
                ' Run-time compression
                Set compressedImage = gfx_api_compressImage(image, uploadFormat)
                If Not (Not compressedImage Is Nothing) Then ASSERT_FAILURE "compressedImage != nullptr", "lib/ivis_opengl/gfx_api.cpp:1097", "loadTextureArrayFromFiles", "Failed to compress image to format: %zu", CLng(uploadFormat): Exit Function
                uploadResult = texture_array.uploadLayer(layer, level, compressedImage)
                If Not (uploadResult) Then ASSERT_FAILURE "uploadResult", "lib/ivis_opengl/gfx_api.cpp:1099", "loadTextureArrayFromFiles", "Failed to upload buffer to image": Exit Function
            Next level
        End If
    Next layer

    If Not texture_array Is Nothing Then
        texture_array.flush
    End If

    Set gfx_api_context_loadTextureArrayFromFiles = texture_array   ' texture_array.release()
End Function

'--------------------------------------------------------------------------
' gfx_api_image_compress_priv.cpp (as built without ETCPAK_ENABLED)

Private Function hasBuiltInRealTimeFormatCompressor(ByVal desiredFormat As pixel_format) As Boolean
    ' std::any_of over builtInRealTimeFormatCompressors (empty)
    hasBuiltInRealTimeFormatCompressor = False
End Function

Public Sub gfx_api_initBestRealTimeCompressionFormats()
    Dim target As Long, target_idx As Long, k As Long, format As pixel_format, s As String
    Dim qualityOrderRGBA(0 To 1) As pixel_format, qualityOrderRGB(0 To 2) As pixel_format
    For target = 0 To 1                         ' PIXEL_FORMAT_TARGET_COUNT
        has_bestAvailableCompressionFormat_GameTextureRGBA(target) = False
        has_bestAvailableCompressionFormat_GameTextureRGB(target) = False
    Next target

    If Not get_wz_texture_compression() Then
        ' Texture compression is disabled - leave all bestAvailableCompressionFormat_* variables as nullopt
        debug_ 60, LOG_3D, "initBestRealTimeCompressionFormats", "Real-time texture compression formats: disabled"
        Exit Sub
    End If

    ' gfx_api::texture_type::game_texture: a RGB / RGBA texture, possibly stored in a compressed format
    ' Overall quality ranking:
    '   FORMAT_ASTC_4x4_UNORM > FORMAT_RGBA_BPTC_UNORM > FORMAT_RGBA8_ETC2_EAC (/ FORMAT_RGB8_ETC2) > FORMAT_RGBA_BC3_UNORM (DXT5) / FORMAT_RGB_BC1_UNORM (for RGB - 4bpp) > ETC1 (only RGB) > PVRTC (is generally the lowest quality)
    qualityOrderRGBA(0) = pixel_format_FORMAT_RGBA8_ETC2_EAC
    qualityOrderRGBA(1) = pixel_format_FORMAT_RGBA_BC3_UNORM
    qualityOrderRGB(0) = pixel_format_FORMAT_RGB8_ETC2
    qualityOrderRGB(1) = pixel_format_FORMAT_RGB_BC1_UNORM
    qualityOrderRGB(2) = pixel_format_FORMAT_RGB8_ETC1

    For target_idx = 0 To 1                     ' PIXEL_FORMAT_TARGET_COUNT
        debug_ 73, LOG_3D, "initBestRealTimeCompressionFormats", "Real-time compression formats [%zu]", target_idx
        For k = 0 To 1
            format = qualityOrderRGBA(k)
            If Not hasBuiltInRealTimeFormatCompressor(format) Then
                GoTo continue_rgba
            End If
            If gfx_api_context_textureFormatIsSupported(target_idx, format, pixel_format_usage_sampled_image) Then
                has_bestAvailableCompressionFormat_GameTextureRGBA(target_idx) = True
                bestAvailableCompressionFormat_GameTextureRGBA(target_idx) = format
                Exit For
            End If
continue_rgba:
        Next k
        If has_bestAvailableCompressionFormat_GameTextureRGBA(target_idx) Then
            s = gfx_api_format_to_str(bestAvailableCompressionFormat_GameTextureRGBA(target_idx))
        Else
            s = "<none>"
        End If
        debug_ 86, LOG_3D, "initBestRealTimeCompressionFormats", "  * GameTextureRGBA: %s", s
        For k = 0 To 2
            format = qualityOrderRGB(k)
            If Not hasBuiltInRealTimeFormatCompressor(format) Then
                GoTo continue_rgb
            End If
            If gfx_api_context_textureFormatIsSupported(target_idx, format, pixel_format_usage_sampled_image) Then
                has_bestAvailableCompressionFormat_GameTextureRGB(target_idx) = True
                bestAvailableCompressionFormat_GameTextureRGB(target_idx) = format
                Exit For
            End If
continue_rgb:
        Next k
        If has_bestAvailableCompressionFormat_GameTextureRGB(target_idx) Then
            s = gfx_api_format_to_str(bestAvailableCompressionFormat_GameTextureRGB(target_idx))
        Else
            s = "<none>"
        End If
        debug_ 99, LOG_3D, "initBestRealTimeCompressionFormats", "  * GameTextureRGB: %s", s
    Next target_idx
End Sub

' MARK: - Live compression functions

' Determine the best available live compressed image format for the current system (+ textureType)
' (optional<pixel_format>: has_value returned, the value in result)
Public Function gfx_api_bestRealTimeCompressionFormatForImage(ByVal target As pixel_format_target, ByVal image As iV_Image, ByVal textureType As texture_type, result As pixel_format) As Boolean
    Dim target_idx As Long
    If image.width() = 0 Or image.height() = 0 Then
        Exit Function
    End If

    target_idx = target

    ' Only certain formats can be computed real-time - much more limited set
    Select Case textureType
    Case texture_type_user_interface
        ' POSSIBLE FUTURE TODO: Real-time: FORMAT_RGBA8_ETC2_EAC
    Case texture_type_game_texture  ' a RGB / RGBA texture, possibly stored in a compressed format
        Select Case image.pixel_format_()
        Case pixel_format_FORMAT_RGBA8_UNORM_PACK8
            result = bestAvailableCompressionFormat_GameTextureRGBA(target_idx)
            gfx_api_bestRealTimeCompressionFormatForImage = has_bestAvailableCompressionFormat_GameTextureRGBA(target_idx)
            Exit Function
        Case pixel_format_FORMAT_RGB8_UNORM_PACK8
            result = bestAvailableCompressionFormat_GameTextureRGB(target_idx)
            gfx_api_bestRealTimeCompressionFormatForImage = has_bestAvailableCompressionFormat_GameTextureRGB(target_idx)
            Exit Function
        End Select
    Case texture_type_alpha_mask    ' a single-channel texture, containing the alpha values
        ' Do not run-time compress this - just store in a single-channel uncompressed texture
    Case Else
        ' unsupported
    End Select

    gfx_api_bestRealTimeCompressionFormatForImage = False
End Function

' Determine the best available live compressed image format for the current system (+ textureType)
Public Function gfx_api_bestRealTimeCompressionFormat(ByVal target As pixel_format_target, ByVal textureType As texture_type, result As pixel_format) As Boolean
    Dim target_idx As Long
    target_idx = target

    ' Only certain formats can be computed real-time - much more limited set
    Select Case textureType
    Case texture_type_user_interface
        ' POSSIBLE FUTURE TODO: Real-time: FORMAT_RGBA8_ETC2_EAC
    Case texture_type_game_texture  ' a RGB / RGBA texture, possibly stored in a compressed format
        ' Since we don't have an image, to check whether it's actually RGB (no alpha), we have to err on the side of RGBA
        result = bestAvailableCompressionFormat_GameTextureRGBA(target_idx)
        gfx_api_bestRealTimeCompressionFormat = has_bestAvailableCompressionFormat_GameTextureRGBA(target_idx)
        Exit Function
    Case texture_type_specular_map, texture_type_alpha_mask, texture_type_height_map
        ' Do not run-time compress this - just store in a single-channel uncompressed texture
    Case Else
        ' unsupported
    End Select

    gfx_api_bestRealTimeCompressionFormat = False
End Function

' Compresses an iV_Image to the desired compressed image format (if possible); Nothing = nullptr
Public Function gfx_api_compressImage(ByVal image As iV_Image, ByVal desiredFormat As pixel_format) As iV_Image
    ' Runtime compress to the desired format: EtcPak for ETC1, ETC2, ETC2_EAC, BC1 and BC3 (not built in: falls through)
    debug_ 411, LOG_ERROR, "compressImage", "Unsupported compressed image format: %u", CLng(desiredFormat)
End Function
