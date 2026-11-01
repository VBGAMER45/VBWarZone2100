Attribute VB_Name = "basisu_transcoder_c"
' Basis Universal's transcoder (ref/basis_universal @ df079ec, transcoder/basisu_transcoder.cpp): basisu_transcoder_init,
' the transcoder_texture_format helpers (basis_get_bytes_per_block_or_pixel, basis_get_format_name,
' basis_transcoder_format_is_uncompressed, basis_is_format_supported, basis_validate_output_buffer_size) and
' basisu_lowlevel_uastc_transcoder (transcode_slice, transcode_image) for UASTC LDR 4x4 data; the blocks are
' basisu_uastc_c's. ktx2_transcoder (the class) calls transcode_image.
'
' The port transcodes UASTC LDR to RGBA32, RGB565, BGR565, RGBA4444, ASTC 4x4, BC7, BC1, BC3, BC4 and BC5. The other
' targets (ETC1, ETC2, EAC R11 / RG11, PVRTC1 / 2, ATC, FXT1, BC6H, ASTC HDR, the half and 9E5 formats) and the
' ETC1S and UASTC HDR inputs are not ported: basis_is_format_supported answers false for them, as the library does
' when it is built with their BASISD_SUPPORT_* switches at 0 (ETC1 and ETC1S have no switch: Windows' desktop
' OpenGL never needs them, see docs/CONVENTIONS.md), and transcoding to them fails.
Option Explicit

' basisu_transcoder_internal.h: enum class block_format (the ones the UASTC LDR transcoder writes)
Private Enum block_format
    cBF_ETC1
    cBF_ETC2_RGBA
    cBF_BC1
    cBF_BC3
    cBF_BC4
    cBF_BC5
    cBF_PVRTC1_4_RGB
    cBF_PVRTC1_4_RGBA
    cBF_BC7
    cBF_BC7_M5_COLOR
    cBF_ASTC_4x4
    cBF_ETC2_EAC_R11
    cBF_ETC2_EAC_RG11
    cBF_RGBA32
    cBF_RGB565
    cBF_BGR565
    cBF_RGBA4444
    cBF_UASTC_4x4
    cBF_FXT1_RGB
End Enum

Private Const cTFBC7_ALT As Long = 7                    ' transcoder_texture_format::cTFBC7_ALT (deprecated: BC7_RGBA)
Private Const UASTC_BLOCK_SIZE As Long = 16             ' sizeof(uastc_block)

Private g_transcoder_initialized As Boolean
Private px(0 To 63) As Long                             ' (VB) a block's color32 pixels (transcode_slice's block_pixels)

' void basisu_transcoder_init()
Public Sub basisu_transcoder_init()
    If g_transcoder_initialized Then
        ' BASISU_DEVEL_ERROR("basisu_transcoder::basisu_transcoder_init: Called more than once\n");
        Exit Sub
    End If

    ' #if BASISD_SUPPORT_UASTC
    uastc_init

    ' (BASISD_SUPPORT_UASTC_HDR, BASISD_SUPPORT_ASTC's ETC1S tables and the ETC1S conversion tables: not ported)

    ' #if BASISD_SUPPORT_DXT1 || BASISD_SUPPORT_UASTC
    basisu_uastc_bc1_init

    g_transcoder_initialized = True
End Sub

' uint32_t basis_get_bytes_per_block_or_pixel(transcoder_texture_format fmt)
Public Function basis_get_bytes_per_block_or_pixel_(ByVal fmt As Long) As Long
    Select Case fmt
    Case cTFETC1_RGB, cTFBC1_RGB, cTFBC4_R, cTFPVRTC1_4_RGB, cTFPVRTC1_4_RGBA, cTFATC_RGB, cTFPVRTC2_4_RGB, cTFPVRTC2_4_RGBA, cTFETC2_EAC_R11
        basis_get_bytes_per_block_or_pixel_ = 8
    Case cTFBC7_RGBA, cTFBC7_ALT, cTFBC6H, cTFETC2_RGBA, cTFBC3_RGBA, cTFBC5_RG, cTFASTC_4x4_RGBA, cTFASTC_HDR_4x4_RGBA, cTFATC_RGBA, cTFFXT1_RGB, cTFETC2_EAC_RG11
        basis_get_bytes_per_block_or_pixel_ = 16
    Case cTFRGBA32, cTFRGB_9E5
        basis_get_bytes_per_block_or_pixel_ = 4
    Case cTFRGB565, cTFBGR565, cTFRGBA4444
        basis_get_bytes_per_block_or_pixel_ = 2
    Case cTFRGB_HALF
        basis_get_bytes_per_block_or_pixel_ = 2 * 3
    Case cTFRGBA_HALF
        basis_get_bytes_per_block_or_pixel_ = 2 * 4
    Case Else
        ' BASISU_DEVEL_ERROR("basis_get_basisu_texture_format: Invalid fmt\n");
        basis_get_bytes_per_block_or_pixel_ = 0
    End Select
End Function

' const char* basis_get_format_name(transcoder_texture_format fmt)
Public Function basis_get_format_name_(ByVal fmt As Long) As String
    Select Case fmt
    Case cTFETC1_RGB: basis_get_format_name_ = "ETC1_RGB"
    Case cTFBC1_RGB: basis_get_format_name_ = "BC1_RGB"
    Case cTFBC4_R: basis_get_format_name_ = "BC4_R"
    Case cTFPVRTC1_4_RGB: basis_get_format_name_ = "PVRTC1_4_RGB"
    Case cTFPVRTC1_4_RGBA: basis_get_format_name_ = "PVRTC1_4_RGBA"
    Case cTFBC7_RGBA: basis_get_format_name_ = "BC7_RGBA"
    Case cTFBC7_ALT: basis_get_format_name_ = "BC7_RGBA"
    Case cTFETC2_RGBA: basis_get_format_name_ = "ETC2_RGBA"
    Case cTFBC3_RGBA: basis_get_format_name_ = "BC3_RGBA"
    Case cTFBC5_RG: basis_get_format_name_ = "BC5_RG"
    Case cTFASTC_4x4_RGBA: basis_get_format_name_ = "ASTC_RGBA"
    Case cTFASTC_HDR_4x4_RGBA: basis_get_format_name_ = "ASTC_HDR_RGBA"
    Case cTFATC_RGB: basis_get_format_name_ = "ATC_RGB"
    Case cTFATC_RGBA: basis_get_format_name_ = "ATC_RGBA"
    Case cTFRGBA32: basis_get_format_name_ = "RGBA32"
    Case cTFRGB565: basis_get_format_name_ = "RGB565"
    Case cTFBGR565: basis_get_format_name_ = "BGR565"
    Case cTFRGBA4444: basis_get_format_name_ = "RGBA4444"
    Case cTFRGBA_HALF: basis_get_format_name_ = "RGBA_HALF"
    Case cTFRGB_9E5: basis_get_format_name_ = "RGB_9E5"
    Case cTFRGB_HALF: basis_get_format_name_ = "RGB_HALF"
    Case cTFFXT1_RGB: basis_get_format_name_ = "FXT1_RGB"
    Case cTFPVRTC2_4_RGB: basis_get_format_name_ = "PVRTC2_4_RGB"
    Case cTFPVRTC2_4_RGBA: basis_get_format_name_ = "PVRTC2_4_RGBA"
    Case cTFETC2_EAC_R11: basis_get_format_name_ = "ETC2_EAC_R11"
    Case cTFETC2_EAC_RG11: basis_get_format_name_ = "ETC2_EAC_RG11"
    Case cTFBC6H: basis_get_format_name_ = "BC6H"
    Case Else
        ' BASISU_DEVEL_ERROR("basis_get_basisu_texture_format: Invalid fmt\n");
        basis_get_format_name_ = ""
    End Select
End Function

' bool basis_transcoder_format_is_uncompressed(transcoder_texture_format tex_type)
Public Function basis_transcoder_format_is_uncompressed_(ByVal tex_type As Long) As Boolean
    Select Case tex_type
    Case cTFRGBA32, cTFRGB565, cTFBGR565, cTFRGBA4444, cTFRGB_HALF, cTFRGBA_HALF, cTFRGB_9E5
        basis_transcoder_format_is_uncompressed_ = True
    End Select
End Function

' bool basis_block_format_is_uncompressed(block_format blk_fmt) (the block formats the UASTC LDR transcoder writes)
Private Function basis_block_format_is_uncompressed(ByVal blk_fmt As block_format) As Boolean
    Select Case blk_fmt
    Case cBF_RGBA32, cBF_RGB565, cBF_BGR565, cBF_RGBA4444
        basis_block_format_is_uncompressed = True
    End Select
End Function

' bool basis_is_format_supported(transcoder_texture_format tex_type, basis_tex_format fmt), as the port is built: UASTC
' LDR to the ported targets only (the header's comment)
Public Function basis_is_format_supported_(ByVal tex_type As Long, ByVal fmt As Long) As Boolean
    If fmt = cUASTC_HDR_4x4 Then
        ' UASTC HDR (#if BASISD_SUPPORT_UASTC_HDR: not ported)
    ElseIf fmt = cUASTC4x4 Then
        ' UASTC LDR
        Select Case tex_type
        Case cTFRGBA32, cTFRGB565, cTFBGR565, cTFRGBA4444, cTFASTC_4x4_RGBA, cTFBC7_RGBA, cTFBC7_ALT, cTFBC1_RGB, cTFBC3_RGBA, cTFBC4_R, cTFBC5_RG
            basis_is_format_supported_ = True
        End Select
    Else
        ' ETC1S (not ported)
    End If
End Function

' bool basis_validate_output_buffer_size(transcoder_texture_format target_format, uint32_t output_blocks_buf_size_in_blocks_or_pixels,
'     uint32_t orig_width, uint32_t orig_height, uint32_t output_row_pitch_in_blocks_or_pixels, uint32_t output_rows_in_pixels,
'     uint32_t total_slice_blocks)
Private Function basis_validate_output_buffer_size(ByVal target_format As Long, ByVal output_blocks_buf_size_in_blocks_or_pixels As Long, ByVal orig_width As Long, ByVal orig_height As Long, ByVal output_row_pitch_in_blocks_or_pixels As Long, ByVal output_rows_in_pixels As Long, ByVal total_slice_blocks As Long) As Boolean
    Dim num_blocks_fxt1_x As Long, num_blocks_fxt1_y As Long, total_blocks_fxt1 As Long
    If basis_transcoder_format_is_uncompressed_(target_format) Then
        ' Assume the output buffer is orig_width by orig_height
        If output_row_pitch_in_blocks_or_pixels = 0 Then output_row_pitch_in_blocks_or_pixels = orig_width

        If output_rows_in_pixels = 0 Then output_rows_in_pixels = orig_height

        ' Now make sure the output buffer is large enough, or we'll overwrite memory.
        If ULT_(output_blocks_buf_size_in_blocks_or_pixels, output_rows_in_pixels * output_row_pitch_in_blocks_or_pixels) Then
            ' BASISU_DEVEL_ERROR("basis_validate_output_buffer_size: output_blocks_buf_size_in_blocks_or_pixels < (output_rows_in_pixels * output_row_pitch_in_blocks_or_pixels)\n");
            Exit Function
        End If
    ElseIf target_format = cTFFXT1_RGB Then
        num_blocks_fxt1_x = (orig_width + 7) \ 8
        num_blocks_fxt1_y = (orig_height + 3) \ 4
        total_blocks_fxt1 = num_blocks_fxt1_x * num_blocks_fxt1_y

        If ULT_(output_blocks_buf_size_in_blocks_or_pixels, total_blocks_fxt1) Then Exit Function
    Else
        If ULT_(output_blocks_buf_size_in_blocks_or_pixels, total_slice_blocks) Then
            ' BASISU_DEVEL_ERROR("basis_validate_output_buffer_size: output_blocks_buf_size_in_blocks_or_pixels < transcode_image\n");
            Exit Function
        End If
    End If
    basis_validate_output_buffer_size = True
End Function

' (VB) C's (a < b) for uint32_t
Private Function ULT_(ByVal a As Long, ByVal b As Long) As Boolean
    ULT_ = (a Xor &H80000000) < (b Xor &H80000000)
End Function

' basisu_lowlevel_uastc_transcoder::transcode_slice(void* pDst_blocks, uint32_t num_blocks_x, uint32_t num_blocks_y,
'     const uint8_t* pImage_data, uint32_t image_data_size, block_format fmt, uint32_t output_block_or_pixel_stride_in_bytes,
'     bool bc1_allow_threecolor_blocks, bool has_alpha, const uint32_t orig_width, const uint32_t orig_height,
'     uint32_t output_row_pitch_in_blocks_or_pixels = 0, basisu_transcoder_state* pState = nullptr, uint32_t output_rows_in_pixels = 0,
'     int channel0 = -1, int channel1 = -1, uint32_t decode_flags = 0)
' pDst_blocks is dst(dst_ofs); pImage_data is src(src_ofs)
Private Function transcode_slice(dst() As Byte, ByVal dst_ofs As Long, ByVal num_blocks_x As Long, ByVal num_blocks_y As Long, src() As Byte, ByVal src_ofs As Long, ByVal image_data_size As Long, ByVal fmt As block_format, ByVal output_block_or_pixel_stride_in_bytes As Long, ByVal bc1_allow_threecolor_blocks As Boolean, ByVal has_alpha As Boolean, ByVal orig_width As Long, ByVal orig_height As Long, Optional ByVal output_row_pitch_in_blocks_or_pixels As Long = 0, Optional ByVal output_rows_in_pixels As Long = 0, Optional ByVal channel0 As Long = -1, Optional ByVal channel1 As Long = -1, Optional ByVal decode_flags As Long = 0) As Boolean
    Dim total_blocks As Long, total_expected_block_bytes As Long, high_quality As Boolean, from_alpha As Boolean
    Dim status As Boolean, block_y As Long, block_x As Long, pDst_block As Long, pSource_block As Long
    Dim pDst_pixels As Long, max_x As Long, max_y As Long, x As Long, y As Long, q As Long, packed As Long, row_bytes As Long

    If Not g_transcoder_initialized Then
        ' BASISU_DEVEL_ERROR("basisu_lowlevel_uastc_transcoder::transcode_slice: Transcoder not globally initialized.\n");
        Exit Function
    End If

    ' #if BASISD_SUPPORT_UASTC
    total_blocks = num_blocks_x * num_blocks_y

    If output_row_pitch_in_blocks_or_pixels = 0 Then
        If basis_block_format_is_uncompressed(fmt) Then
            output_row_pitch_in_blocks_or_pixels = orig_width
        Else
            If fmt = cBF_FXT1_RGB Then
                output_row_pitch_in_blocks_or_pixels = (orig_width + 7) \ 8
            Else
                output_row_pitch_in_blocks_or_pixels = num_blocks_x
            End If
        End If
    End If

    If basis_block_format_is_uncompressed(fmt) Then
        If output_rows_in_pixels = 0 Then output_rows_in_pixels = orig_height
    End If

    total_expected_block_bytes = UASTC_BLOCK_SIZE * total_blocks
    If ULT_(image_data_size, total_expected_block_bytes) Then
        ' BASISU_DEVEL_ERROR("basisu_lowlevel_uastc_transcoder::transcode_slice: image_data_size < total_expected_block_bytes The file is corrupted or this is a bug.\n");
        Exit Function
    End If

    ' (VB) the C writes through its caller's pointer: here the destination array must hold what the loops write
    If basis_block_format_is_uncompressed(fmt) Then
        row_bytes = output_rows_in_pixels * output_row_pitch_in_blocks_or_pixels * output_block_or_pixel_stride_in_bytes
    Else
        row_bytes = ((num_blocks_y - 1) * output_row_pitch_in_blocks_or_pixels + num_blocks_x) * output_block_or_pixel_stride_in_bytes
    End If
    If dst_ofs < LBound(dst) Or dst_ofs + row_bytes - 1 > UBound(dst) Then Exit Function

    pSource_block = src_ofs

    high_quality = (decode_flags And cDecodeFlagsHighQuality) <> 0
    from_alpha = has_alpha And ((decode_flags And cDecodeFlagsTranscodeAlphaDataToOpaqueFormats) <> 0)

    status = False
    If (fmt = cBF_PVRTC1_4_RGB) Or (fmt = cBF_PVRTC1_4_RGBA) Then
        ' transcode_uastc_to_pvrtc1_4_rgba / _rgb: not ported (basis_is_format_supported says so)
        Exit Function
    Else
        For block_y = 0 To num_blocks_y - 1
            pDst_block = dst_ofs + block_y * output_row_pitch_in_blocks_or_pixels * output_block_or_pixel_stride_in_bytes

            For block_x = 0 To num_blocks_x - 1
                Select Case fmt
                Case cBF_UASTC_4x4
                    For q = 0 To UASTC_BLOCK_SIZE - 1
                        dst(pDst_block + q) = src(pSource_block + q)
                    Next
                    status = True
                Case cBF_BC1
                    status = transcode_uastc_to_bc1(src, pSource_block, dst, pDst_block, high_quality)
                Case cBF_BC3
                    status = transcode_uastc_to_bc3(src, pSource_block, dst, pDst_block, high_quality)
                Case cBF_BC4
                    If channel0 < 0 Then channel0 = 0
                    status = transcode_uastc_to_bc4(src, pSource_block, dst, pDst_block, high_quality, channel0)
                Case cBF_BC5
                    If channel0 < 0 Then channel0 = 0
                    If channel1 < 0 Then channel1 = 3
                    status = transcode_uastc_to_bc5(src, pSource_block, dst, pDst_block, high_quality, channel0, channel1)
                Case cBF_BC7, cBF_BC7_M5_COLOR                  ' for consistently with ETC1S
                    status = transcode_uastc_to_bc7(src, pSource_block, dst, pDst_block)
                Case cBF_ASTC_4x4
                    status = transcode_uastc_to_astc(src, pSource_block, dst, pDst_block)
                Case cBF_RGBA32
                    status = unpack_uastc_rgba(src, pSource_block, px, False)

                    pDst_pixels = dst_ofs + (block_x * 4 + block_y * 4 * output_row_pitch_in_blocks_or_pixels) * 4

                    max_x = output_row_pitch_in_blocks_or_pixels - block_x * 4
                    If max_x > 4 Then max_x = 4
                    max_y = output_rows_in_pixels - block_y * 4
                    If max_y > 4 Then max_y = 4

                    For y = 0 To max_y - 1
                        For x = 0 To max_x - 1
                            q = (y * 4 + x) * 4                 ' block_pixels[y][x]

                            dst(pDst_pixels + 0 + 4 * x) = px(q + 0)
                            dst(pDst_pixels + 1 + 4 * x) = px(q + 1)
                            dst(pDst_pixels + 2 + 4 * x) = px(q + 2)
                            dst(pDst_pixels + 3 + 4 * x) = px(q + 3)
                        Next

                        pDst_pixels = pDst_pixels + output_row_pitch_in_blocks_or_pixels * 4
                    Next
                Case cBF_RGB565, cBF_BGR565
                    status = unpack_uastc_rgba(src, pSource_block, px, False)

                    pDst_pixels = dst_ofs + (block_x * 4 + block_y * 4 * output_row_pitch_in_blocks_or_pixels) * 2

                    max_x = output_row_pitch_in_blocks_or_pixels - block_x * 4
                    If max_x > 4 Then max_x = 4
                    max_y = output_rows_in_pixels - block_y * 4
                    If max_y > 4 Then max_y = 4

                    For y = 0 To max_y - 1
                        For x = 0 To max_x - 1
                            q = (y * 4 + x) * 4

                            If fmt = cBF_RGB565 Then
                                packed = (mul_8(px(q + 0), 31) * 2048) Or (mul_8(px(q + 1), 63) * 32) Or mul_8(px(q + 2), 31)
                            Else
                                packed = (mul_8(px(q + 2), 31) * 2048) Or (mul_8(px(q + 1), 63) * 32) Or mul_8(px(q + 0), 31)
                            End If

                            dst(pDst_pixels + x * 2 + 0) = packed And &HFF
                            dst(pDst_pixels + x * 2 + 1) = (packed \ 256) And &HFF
                        Next

                        pDst_pixels = pDst_pixels + output_row_pitch_in_blocks_or_pixels * 2
                    Next
                Case cBF_RGBA4444
                    status = unpack_uastc_rgba(src, pSource_block, px, False)

                    pDst_pixels = dst_ofs + (block_x * 4 + block_y * 4 * output_row_pitch_in_blocks_or_pixels) * 2

                    max_x = output_row_pitch_in_blocks_or_pixels - block_x * 4
                    If max_x > 4 Then max_x = 4
                    max_y = output_rows_in_pixels - block_y * 4
                    If max_y > 4 Then max_y = 4

                    For y = 0 To max_y - 1
                        For x = 0 To max_x - 1
                            q = (y * 4 + x) * 4

                            packed = (mul_8(px(q + 0), 15) * 4096) Or (mul_8(px(q + 1), 15) * 256) Or (mul_8(px(q + 2), 15) * 16) Or mul_8(px(q + 3), 15)

                            dst(pDst_pixels + x * 2 + 0) = packed And &HFF
                            dst(pDst_pixels + x * 2 + 1) = (packed \ 256) And &HFF
                        Next

                        pDst_pixels = pDst_pixels + output_row_pitch_in_blocks_or_pixels * 2
                    Next
                Case Else
                    ' cETC1, cETC2_RGBA, cETC2_EAC_R11, cETC2_EAC_RG11: not ported (basis_is_format_supported says so)
                    status = False
                End Select

                If Not status Then
                    ' BASISU_DEVEL_ERROR("basisu_lowlevel_uastc_transcoder::transcode_slice: Transcoder failed to unpack a UASTC block - this is a bug, or the data was corrupted\n");
                    Exit Function
                End If

                pSource_block = pSource_block + UASTC_BLOCK_SIZE
                pDst_block = pDst_block + output_block_or_pixel_stride_in_bytes
            Next
        Next
    End If

    transcode_slice = True
End Function

' static inline uint8_t mul_8(uint32_t v, uint32_t q)
Private Function mul_8(ByVal v As Long, ByVal q As Long) As Long
    v = v * q + 128
    mul_8 = ((v + (v \ 256)) \ 256) And &HFF
End Function

' bool basisu_lowlevel_uastc_transcoder::transcode_image(transcoder_texture_format target_format, void* pOutput_blocks,
'     uint32_t output_blocks_buf_size_in_blocks_or_pixels, const uint8_t* pCompressed_data, uint32_t compressed_data_length,
'     uint32_t num_blocks_x, uint32_t num_blocks_y, uint32_t orig_width, uint32_t orig_height, uint32_t level_index,
'     uint32_t slice_offset, uint32_t slice_length, uint32_t decode_flags = 0, bool has_alpha = false, bool is_video = false,
'     uint32_t output_row_pitch_in_blocks_or_pixels = 0, basisu_transcoder_state* pState = nullptr,
'     uint32_t output_rows_in_pixels = 0, int channel0 = -1, int channel1 = -1)
' pOutput_blocks is out(out_ofs); pCompressed_data is comp(comp_ofs)
Public Function basisu_uastc_transcode_image(ByVal target_format As Long, out() As Byte, ByVal out_ofs As Long, ByVal output_blocks_buf_size_in_blocks_or_pixels As Long, comp() As Byte, ByVal comp_ofs As Long, ByVal compressed_data_length As Long, ByVal num_blocks_x As Long, ByVal num_blocks_y As Long, ByVal orig_width As Long, ByVal orig_height As Long, ByVal level_index As Long, ByVal slice_offset As Long, ByVal slice_length As Long, ByVal decode_flags As Long, ByVal has_alpha As Boolean, ByVal is_video As Boolean, ByVal output_row_pitch_in_blocks_or_pixels As Long, ByVal output_rows_in_pixels As Long, ByVal channel0 As Long, ByVal channel1 As Long) As Boolean
    Dim transcode_alpha_data_to_opaque_formats As Boolean, bytes_per_block_or_pixel As Long, total_slice_blocks As Long
    Dim status As Boolean, ch0 As Long, so As Long

    If CDbl(U32_(slice_offset)) + U32_(slice_length) > U32_(compressed_data_length) Then
        ' BASISU_DEVEL_ERROR("basisu_lowlevel_uastc_transcoder::transcode_image: source data buffer too small\n");
        Exit Function
    End If

    If (target_format = cTFPVRTC1_4_RGB) Or (target_format = cTFPVRTC1_4_RGBA) Then
        If (Not is_pow2(num_blocks_x * 4)) Or (Not is_pow2(num_blocks_y * 4)) Then
            ' PVRTC1 only supports power of 2 dimensions
            Exit Function
        End If
    End If

    If (target_format = cTFPVRTC1_4_RGBA) And (Not has_alpha) Then
        ' Switch to PVRTC1 RGB if the input doesn't have alpha.
        target_format = cTFPVRTC1_4_RGB
    End If

    transcode_alpha_data_to_opaque_formats = (decode_flags And cDecodeFlagsTranscodeAlphaDataToOpaqueFormats) <> 0
    bytes_per_block_or_pixel = basis_get_bytes_per_block_or_pixel_(target_format)
    total_slice_blocks = num_blocks_x * num_blocks_y

    If Not basis_validate_output_buffer_size(target_format, output_blocks_buf_size_in_blocks_or_pixels, orig_width, orig_height, output_row_pitch_in_blocks_or_pixels, output_rows_in_pixels, total_slice_blocks) Then
        ' BASISU_DEVEL_ERROR("basisu_lowlevel_uastc_transcoder::transcode_image: output buffer size too small\n");
        Exit Function
    End If

    status = False
    so = comp_ofs + slice_offset

    ' UASTC4x4
    ' (transcode_slice's decode_flags is left at its default, 0: the C passes none)
    Select Case target_format
    Case cTFETC1_RGB
        status = transcode_slice(out, out_ofs, num_blocks_x, num_blocks_y, comp, so, slice_length, cBF_ETC1, bytes_per_block_or_pixel, False, has_alpha, orig_width, orig_height, output_row_pitch_in_blocks_or_pixels, output_rows_in_pixels, channel0, channel1)
    Case cTFETC2_RGBA
        status = transcode_slice(out, out_ofs, num_blocks_x, num_blocks_y, comp, so, slice_length, cBF_ETC2_RGBA, bytes_per_block_or_pixel, False, has_alpha, orig_width, orig_height, output_row_pitch_in_blocks_or_pixels, output_rows_in_pixels, channel0, channel1)
    Case cTFBC1_RGB
        ' TODO: ETC1S allows BC1 from alpha channel. That doesn't seem actually useful, though.
        status = transcode_slice(out, out_ofs, num_blocks_x, num_blocks_y, comp, so, slice_length, cBF_BC1, bytes_per_block_or_pixel, True, has_alpha, orig_width, orig_height, output_row_pitch_in_blocks_or_pixels, output_rows_in_pixels, channel0, channel1)
    Case cTFBC3_RGBA
        status = transcode_slice(out, out_ofs, num_blocks_x, num_blocks_y, comp, so, slice_length, cBF_BC3, bytes_per_block_or_pixel, False, has_alpha, orig_width, orig_height, output_row_pitch_in_blocks_or_pixels, output_rows_in_pixels, channel0, channel1)
    Case cTFBC4_R
        If has_alpha And transcode_alpha_data_to_opaque_formats Then ch0 = 3 Else ch0 = 0
        status = transcode_slice(out, out_ofs, num_blocks_x, num_blocks_y, comp, so, slice_length, cBF_BC4, bytes_per_block_or_pixel, False, has_alpha, orig_width, orig_height, output_row_pitch_in_blocks_or_pixels, output_rows_in_pixels, ch0)
    Case cTFBC5_RG
        status = transcode_slice(out, out_ofs, num_blocks_x, num_blocks_y, comp, so, slice_length, cBF_BC5, bytes_per_block_or_pixel, False, has_alpha, orig_width, orig_height, output_row_pitch_in_blocks_or_pixels, output_rows_in_pixels, 0, 3)
    Case cTFBC7_RGBA, cTFBC7_ALT
        status = transcode_slice(out, out_ofs, num_blocks_x, num_blocks_y, comp, so, slice_length, cBF_BC7, bytes_per_block_or_pixel, False, has_alpha, orig_width, orig_height, output_row_pitch_in_blocks_or_pixels, output_rows_in_pixels)
    Case cTFPVRTC1_4_RGB
        status = transcode_slice(out, out_ofs, num_blocks_x, num_blocks_y, comp, so, slice_length, cBF_PVRTC1_4_RGB, bytes_per_block_or_pixel, False, has_alpha, orig_width, orig_height, output_row_pitch_in_blocks_or_pixels, output_rows_in_pixels)
    Case cTFPVRTC1_4_RGBA
        status = transcode_slice(out, out_ofs, num_blocks_x, num_blocks_y, comp, so, slice_length, cBF_PVRTC1_4_RGBA, bytes_per_block_or_pixel, False, has_alpha, orig_width, orig_height, output_row_pitch_in_blocks_or_pixels, output_rows_in_pixels)
    Case cTFASTC_4x4_RGBA
        status = transcode_slice(out, out_ofs, num_blocks_x, num_blocks_y, comp, so, slice_length, cBF_ASTC_4x4, bytes_per_block_or_pixel, False, has_alpha, orig_width, orig_height, output_row_pitch_in_blocks_or_pixels, output_rows_in_pixels)
    Case cTFATC_RGB, cTFATC_RGBA
        ' BASISU_DEVEL_ERROR("basisu_lowlevel_uastc_transcoder::transcode_image: UASTC->ATC currently unsupported\n");
        Exit Function
    Case cTFFXT1_RGB
        ' BASISU_DEVEL_ERROR("basisu_lowlevel_uastc_transcoder::transcode_image: UASTC->FXT1 currently unsupported\n");
        Exit Function
    Case cTFPVRTC2_4_RGB, cTFPVRTC2_4_RGBA
        ' BASISU_DEVEL_ERROR("basisu_lowlevel_uastc_transcoder::transcode_image: UASTC->PVRTC2 currently unsupported\n");
        Exit Function
    Case cTFETC2_EAC_R11
        If has_alpha And transcode_alpha_data_to_opaque_formats Then ch0 = 3 Else ch0 = 0
        status = transcode_slice(out, out_ofs, num_blocks_x, num_blocks_y, comp, so, slice_length, cBF_ETC2_EAC_R11, bytes_per_block_or_pixel, False, has_alpha, orig_width, orig_height, output_row_pitch_in_blocks_or_pixels, output_rows_in_pixels, ch0)
    Case cTFETC2_EAC_RG11
        status = transcode_slice(out, out_ofs, num_blocks_x, num_blocks_y, comp, so, slice_length, cBF_ETC2_EAC_RG11, bytes_per_block_or_pixel, False, has_alpha, orig_width, orig_height, output_row_pitch_in_blocks_or_pixels, output_rows_in_pixels, 0, 3)
    Case cTFRGBA32
        status = transcode_slice(out, out_ofs, num_blocks_x, num_blocks_y, comp, so, slice_length, cBF_RGBA32, bytes_per_block_or_pixel, False, has_alpha, orig_width, orig_height, output_row_pitch_in_blocks_or_pixels, output_rows_in_pixels)
    Case cTFRGB565
        status = transcode_slice(out, out_ofs, num_blocks_x, num_blocks_y, comp, so, slice_length, cBF_RGB565, bytes_per_block_or_pixel, False, has_alpha, orig_width, orig_height, output_row_pitch_in_blocks_or_pixels, output_rows_in_pixels)
    Case cTFBGR565
        status = transcode_slice(out, out_ofs, num_blocks_x, num_blocks_y, comp, so, slice_length, cBF_BGR565, bytes_per_block_or_pixel, False, has_alpha, orig_width, orig_height, output_row_pitch_in_blocks_or_pixels, output_rows_in_pixels)
    Case cTFRGBA4444
        status = transcode_slice(out, out_ofs, num_blocks_x, num_blocks_y, comp, so, slice_length, cBF_RGBA4444, bytes_per_block_or_pixel, False, has_alpha, orig_width, orig_height, output_row_pitch_in_blocks_or_pixels, output_rows_in_pixels)
    Case Else
        ' BASISU_DEVEL_ERROR("basisu_lowlevel_uastc_transcoder::transcode_image: Invalid format\n");
    End Select

    basisu_uastc_transcode_image = status
End Function

' (VB) a uint32_t's value
Private Function U32_(ByVal x As Long) As Double
    If x < 0 Then U32_ = x + 4294967296# Else U32_ = x
End Function

' basisu::is_pow2(uint32_t x)
Private Function is_pow2(ByVal x As Long) As Boolean
    is_pow2 = (x <> 0) And ((x And (x - 1)) = 0)
End Function
