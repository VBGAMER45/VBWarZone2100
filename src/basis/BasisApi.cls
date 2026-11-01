VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BasisApi"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' WzBasis.dll's public interface: Basis Universal's transcoder (ref/basis_universal @ df079ec, the commit 4.7.0
' pins: transcoder/basisu_transcoder.h) as lib/ivis_opengl/gfx_api_image_basis_priv.cpp uses it, over the zstd
' decoder its KTX2 files need (zstd/zstddeclib.c: ZSTD_decompress). GlobalMultiUse: WzGfx calls these procedures
' as if they were its own; a basist::ktx2_transcoder is a ktx2_transcoder object (New). The transcoder is ported
' into this DLL's modules; this class only forwards to them.
Option Explicit

' basisu_transcoder.h: basist::transcoder_texture_format
Public Enum transcoder_texture_format
    cTFETC1_RGB = 0
    cTFETC2_RGBA = 1
    cTFBC1_RGB = 2
    cTFBC3_RGBA = 3
    cTFBC4_R = 4
    cTFBC5_RG = 5
    cTFBC7_RGBA = 6
    cTFPVRTC1_4_RGB = 8
    cTFPVRTC1_4_RGBA = 9
    cTFASTC_4x4_RGBA = 10
    cTFATC_RGB = 11
    cTFATC_RGBA = 12
    cTFFXT1_RGB = 17
    cTFPVRTC2_4_RGB = 18
    cTFPVRTC2_4_RGBA = 19
    cTFETC2_EAC_R11 = 20
    cTFETC2_EAC_RG11 = 21
    cTFBC6H = 22
    cTFASTC_HDR_4x4_RGBA = 23
    cTFRGBA32 = 13
    cTFRGB565 = 14
    cTFBGR565 = 15
    cTFRGBA4444 = 16
    cTFRGB_HALF = 24
    cTFRGBA_HALF = 25
    cTFRGB_9E5 = 26
    cTFTotalTextureFormats = 27
End Enum

' basisu_file_headers.h: basist::basis_tex_format
Public Enum basis_tex_format
    cETC1S = 0
    cUASTC4x4 = 1
    cUASTC_HDR_4x4 = 2
End Enum

' basisu_transcoder.h: basist::ktx2_df_channel_id
Public Enum ktx2_df_channel_id
    KTX2_DF_CHANNEL_ETC1S_RGB = 0
    KTX2_DF_CHANNEL_ETC1S_RRR = 3
    KTX2_DF_CHANNEL_ETC1S_GGG = 4
    KTX2_DF_CHANNEL_ETC1S_AAA = 15
    KTX2_DF_CHANNEL_UASTC_DATA = 0
    KTX2_DF_CHANNEL_UASTC_RGB = 0
    KTX2_DF_CHANNEL_UASTC_RGBA = 3
    KTX2_DF_CHANNEL_UASTC_RRR = 4
    KTX2_DF_CHANNEL_UASTC_RRRG = 5
    KTX2_DF_CHANNEL_UASTC_RG = 6
End Enum

' basisu_transcoder.h: basist::basisu_decode_flags
Public Enum basisu_decode_flags
    cDecodeFlagsPVRTCDecodeToNextPow2 = 2
    cDecodeFlagsTranscodeAlphaDataToOpaqueFormats = 4
    cDecodeFlagsBC1ForbidThreeColorBlocks = 8
    cDecodeFlagsOutputHasAlphaIndices = 16
    cDecodeFlagsHighQuality = 32
End Enum

' Name and version of this DLL, for the log header.
Public Function BasisBuildInfo() As String
    BasisBuildInfo = App.EXEName & " " & App.Major & "." & App.Minor & "." & App.Revision
End Function

' void basisu_transcoder_init()
Public Sub basisu_transcoder_init()
    basisu_transcoder_c.basisu_transcoder_init
End Sub

' bool basis_is_format_supported(transcoder_texture_format tex_type, basis_tex_format fmt = basis_tex_format::cETC1S)
' (the port: UASTC LDR to RGBA32, RGB565, BGR565, RGBA4444, ASTC 4x4, BC7, BC1, BC3, BC4, BC5; see basisu_transcoder_c)
Public Function basis_is_format_supported(ByVal tex_type As transcoder_texture_format, Optional ByVal fmt As basis_tex_format = cETC1S) As Boolean
    basis_is_format_supported = basis_is_format_supported_(tex_type, fmt)
End Function

' const char* basis_get_format_name(transcoder_texture_format fmt)
Public Function basis_get_format_name(ByVal fmt As transcoder_texture_format) As String
    basis_get_format_name = basis_get_format_name_(fmt)
End Function

' uint32_t basis_get_bytes_per_block_or_pixel(transcoder_texture_format fmt)
Public Function basis_get_bytes_per_block_or_pixel(ByVal fmt As transcoder_texture_format) As Long
    basis_get_bytes_per_block_or_pixel = basis_get_bytes_per_block_or_pixel_(fmt)
End Function

' bool basis_transcoder_format_is_uncompressed(transcoder_texture_format tex_type)
Public Function basis_transcoder_format_is_uncompressed(ByVal tex_type As transcoder_texture_format) As Boolean
    basis_transcoder_format_is_uncompressed = basis_transcoder_format_is_uncompressed_(tex_type)
End Function
