VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ktx2_transcoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' basisu_transcoder.h / basisu_transcoder.cpp: class basist::ktx2_transcoder (KTX2 files with UASTC LDR 4x4 data,
' optionally Zstandard supercompressed). MultiUse: the caller makes one with New (C++: a local object); Class_Terminate
' is the destructor. init keeps a copy of the file's bytes (C++ keeps the caller's pointer, which outlives it).
' The transcoder state is the object's own (C++: its m_def_transcoder_state when no state is passed).
'
' The header's packed_uint fields are Longs holding their uint32_t values (packed_uint<8>'s operator uint32_t gives
' the low 32 bits, as here). ETC1S (BasisLZ) and UASTC HDR files are read by init as in C, but their transcoding is
' not ported: start_transcoding fails for BasisLZ (its global data), transcode_image_level fails for both
' (basisu_transcoder_c's header).
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

' basisu_transcoder.h's KTX2 constants
Private Const KTX2_VK_FORMAT_UNDEFINED As Long = 0
Private Const KTX2_FORMAT_UASTC_4x4_SFLOAT_BLOCK As Long = 1000066000
Private Const KTX2_KDF_DF_MODEL_UASTC As Long = 166
Private Const KTX2_KDF_DF_MODEL_UASTC_HDR As Long = 167
Private Const KTX2_KDF_DF_MODEL_ETC1S As Long = 163
Private Const KTX2_UASTC_BLOCK_SIZE As Long = 16
Private Const KTX2_MAX_SUPPORTED_LEVEL_COUNT As Long = 16
Private Const KTX2_KHR_DF_TRANSFER_LINEAR As Long = 1
Private Const KTX2_KHR_DF_TRANSFER_SRGB As Long = 2
Private Const KTX2_SS_NONE As Long = 0
Private Const KTX2_SS_BASISLZ As Long = 1
Private Const KTX2_SS_ZSTANDARD As Long = 2
Private Const KTX2_DF_PRIMARIES_UNSPECIFIED As Long = 0
Private Const SIZEOF_KTX2_HEADER As Long = 80
Private Const SIZEOF_KTX2_LEVEL_INDEX As Long = 24
Private Const SIZEOF_KTX2_ETC1S_GLOBAL_DATA_HEADER As Long = 20

' const uint8_t* m_pData; uint32_t m_data_size;
Private m_pData() As Byte
Private m_has_data As Boolean                   ' (m_pData != nullptr)
Private m_data_size As Long

' ktx2_header m_header;
Private m_header_vk_format As Long
Private m_header_type_size As Long
Private m_header_pixel_width As Long
Private m_header_pixel_height As Long
Private m_header_pixel_depth As Long
Private m_header_layer_count As Long
Private m_header_face_count As Long
Private m_header_level_count As Long
Private m_header_supercompression_scheme As Long
Private m_header_dfd_byte_offset As Long
Private m_header_dfd_byte_length As Long
Private m_header_kvd_byte_offset As Long
Private m_header_kvd_byte_length As Long
Private m_header_sgd_byte_offset As Long
Private m_header_sgd_byte_length As Long

' basisu::vector<ktx2_level_index> m_levels;
Private m_levels_byte_offset() As Long
Private m_levels_byte_length() As Long
Private m_levels_uncompressed_byte_length() As Long
Private m_levels_size As Long

' basisu::uint8_vec m_dfd;
Private m_dfd() As Byte

' key_value_vec m_key_values: the keys as strings of their bytes (with the terminating NUL), the values as Byte arrays
Private m_key_values_key() As String
Private m_key_values_value() As Variant
Private m_key_values_size As Long

Private m_format As basis_tex_format

Private m_dfd_color_model As Long
Private m_dfd_color_prims As Long
Private m_dfd_transfer_func As Long
Private m_dfd_flags As Long
Private m_dfd_samples As Long
Private m_dfd_chan0 As ktx2_df_channel_id
Private m_dfd_chan1 As ktx2_df_channel_id

' ktx2_transcoder_state m_def_transcoder_state (its m_level_uncomp_data and m_uncomp_data_level_index)
Private m_level_uncomp_data() As Byte
Private m_level_uncomp_data_size As Long
Private m_uncomp_data_level_index As Long

Private m_has_alpha As Boolean
Private m_is_video As Boolean

Private Sub Class_Initialize()
    clear
End Sub

' (VB) uint32_t read_le_dword(const uint8_t *pBytes) of m_pData(p)
Private Function read_le_dword(b() As Byte, ByVal p As Long) As Long
    Dim v As Long
    CopyMemory v, b(p), 4
    read_le_dword = v
End Function

' (VB) C's (a < b) and (a > b) for uint32_t
Private Function ULT_(ByVal a As Long, ByVal b As Long) As Boolean
    ULT_ = (a Xor &H80000000) < (b Xor &H80000000)
End Function

' (VB) a uint32_t's value
Private Function U32_(ByVal x As Long) As Double
    If x < 0 Then U32_ = x + 4294967296# Else U32_ = x
End Function

' void clear()
Public Sub clear()
    Erase m_pData
    m_has_data = False
    m_data_size = 0

    m_header_vk_format = 0
    m_header_type_size = 0
    m_header_pixel_width = 0
    m_header_pixel_height = 0
    m_header_pixel_depth = 0
    m_header_layer_count = 0
    m_header_face_count = 0
    m_header_level_count = 0
    m_header_supercompression_scheme = 0
    m_header_dfd_byte_offset = 0
    m_header_dfd_byte_length = 0
    m_header_kvd_byte_offset = 0
    m_header_kvd_byte_length = 0
    m_header_sgd_byte_offset = 0
    m_header_sgd_byte_length = 0
    m_levels_size = 0
    Erase m_levels_byte_offset, m_levels_byte_length, m_levels_uncompressed_byte_length
    Erase m_dfd
    m_key_values_size = 0
    Erase m_key_values_key, m_key_values_value

    m_format = cETC1S

    m_dfd_color_model = 0
    m_dfd_color_prims = KTX2_DF_PRIMARIES_UNSPECIFIED
    m_dfd_transfer_func = 0
    m_dfd_flags = 0
    m_dfd_samples = 0
    m_dfd_chan0 = KTX2_DF_CHANNEL_UASTC_RGB
    m_dfd_chan1 = KTX2_DF_CHANNEL_UASTC_RGB

    ' m_def_transcoder_state.clear()
    Erase m_level_uncomp_data
    m_level_uncomp_data_size = 0
    m_uncomp_data_level_index = -1

    m_has_alpha = False
    m_is_video = False
End Sub

' bool init(const void* pData, uint32_t data_size): data(0 .. data_size - 1)
Public Function init(data() As Byte, ByVal data_size As Long) As Boolean
    Dim i As Long, level_index_size_in_bytes As Long, n As Long, p As Long, dfd_total_size As Long, dfd_bits As Long
    Dim sample_channel0 As Long, sample_channel1 As Long, ident As Variant
    Const DFD_MINIMUM_SIZE As Long = 44, DFD_MAXIMUM_SIZE As Long = 60

    clear

    n = array_size(data)
    If n < 0 Then
        ' BASISU_DEVEL_ERROR("ktx2_transcoder::init: pData is nullptr\n");
        Exit Function
    End If

    If Not ULT_(SIZEOF_KTX2_HEADER, data_size) Then
        ' BASISU_DEVEL_ERROR("ktx2_transcoder::init: File is impossibly too small to be a valid KTX2 file\n");
        Exit Function
    End If

    ' (VB) the caller's buffer must hold data_size bytes (C reads them through the pointer)
    If ULT_(n, data_size) Then Exit Function

    ' memcmp(pData, g_ktx2_file_identifier, sizeof(g_ktx2_file_identifier))
    ident = Array(&HAB, &H4B, &H54, &H58, &H20, &H32, &H30, &HBB, &HD, &HA, &H1A, &HA)
    For i = 0 To 11
        If data(LBound(data) + i) <> ident(i) Then
            ' BASISU_DEVEL_ERROR("ktx2_transcoder::init: KTX2 file identifier is not present\n");
            Exit Function
        End If
    Next

    ReDim m_pData(0 To data_size - 1)
    CopyMemory m_pData(0), data(LBound(data)), data_size
    m_has_data = True
    m_data_size = data_size

    ' memcpy(&m_header, pData, sizeof(m_header));
    m_header_vk_format = read_le_dword(m_pData, 12)
    m_header_type_size = read_le_dword(m_pData, 16)
    m_header_pixel_width = read_le_dword(m_pData, 20)
    m_header_pixel_height = read_le_dword(m_pData, 24)
    m_header_pixel_depth = read_le_dword(m_pData, 28)
    m_header_layer_count = read_le_dword(m_pData, 32)
    m_header_face_count = read_le_dword(m_pData, 36)
    m_header_level_count = read_le_dword(m_pData, 40)
    m_header_supercompression_scheme = read_le_dword(m_pData, 44)
    m_header_dfd_byte_offset = read_le_dword(m_pData, 48)
    m_header_dfd_byte_length = read_le_dword(m_pData, 52)
    m_header_kvd_byte_offset = read_le_dword(m_pData, 56)
    m_header_kvd_byte_length = read_le_dword(m_pData, 60)
    m_header_sgd_byte_offset = read_le_dword(m_pData, 64)     ' packed_uint<8>: its low 32 bits
    m_header_sgd_byte_length = read_le_dword(m_pData, 72)

    ' We only support UASTC LDR, UASTC HDR and ETC1S.
    ' Note the DFD's contents are what we are guided by for decoding the KTX2 file, not this format field (currently).
    If (m_header_vk_format <> KTX2_VK_FORMAT_UNDEFINED) And (m_header_vk_format <> KTX2_FORMAT_UASTC_4x4_SFLOAT_BLOCK) Then
        ' BASISU_DEVEL_ERROR("ktx2_transcoder::init: KTX2 file must be in ETC1S or UASTC LDR/HDR format\n");
        Exit Function
    End If

    ' 3.3: "When format is VK_FORMAT_UNDEFINED, typeSize must equal 1."
    If m_header_type_size <> 1 Then Exit Function

    ' We only currently support 2D textures (plain, cubemapped, or texture array), which is by far the most common use case.
    ' The BasisU library does not support 1D or 3D textures at all.
    If (m_header_pixel_width = 0) Or (m_header_pixel_height = 0) Or (m_header_pixel_depth <> 0) Then
        ' BASISU_DEVEL_ERROR("ktx2_transcoder::init: Only 2D or cubemap textures are supported\n");
        Exit Function
    End If

    ' Face count must be 1 or 6
    If (m_header_face_count <> 1) And (m_header_face_count <> 6) Then Exit Function

    If m_header_face_count > 1 Then
        ' 3.4: Make sure cubemaps are square.
        If m_header_pixel_width <> m_header_pixel_height Then Exit Function
    End If

    ' 3.7 levelCount: "levelCount=0 is allowed, except for block-compressed formats"
    If m_header_level_count = 0 Then Exit Function

    ' Sanity check the level count.
    If ULT_(KTX2_MAX_SUPPORTED_LEVEL_COUNT, m_header_level_count) Then Exit Function

    If ULT_(KTX2_SS_ZSTANDARD, m_header_supercompression_scheme) Then
        ' BASISU_DEVEL_ERROR("ktx2_transcoder::init: Invalid/unsupported supercompression or file is corrupted or invalid\n");
        Exit Function
    End If

    If m_header_supercompression_scheme = KTX2_SS_BASISLZ Then
        If Not ULT_(SIZEOF_KTX2_ETC1S_GLOBAL_DATA_HEADER, m_header_sgd_byte_length) Then Exit Function

        If ULT_(m_header_sgd_byte_offset, SIZEOF_KTX2_HEADER) Then Exit Function

        If ULT_(m_data_size, m_header_sgd_byte_offset + m_header_sgd_byte_length) Then Exit Function
    End If

    ' m_levels.try_resize(m_header.m_level_count)
    m_levels_size = m_header_level_count
    ReDim m_levels_byte_offset(0 To m_levels_size - 1)
    ReDim m_levels_byte_length(0 To m_levels_size - 1)
    ReDim m_levels_uncompressed_byte_length(0 To m_levels_size - 1)

    n = m_header_level_count
    If n < 1 Then n = 1                                 ' basisu::maximum(1U, (uint32_t)m_header.m_level_count)
    level_index_size_in_bytes = n * SIZEOF_KTX2_LEVEL_INDEX

    If ULT_(m_data_size, SIZEOF_KTX2_HEADER + level_index_size_in_bytes) Then
        ' BASISU_DEVEL_ERROR("ktx2_transcoder::init: File is too small (can't read level index array)\n");
        Exit Function
    End If

    ' memcpy(&m_levels[0], m_pData + sizeof(ktx2_header), level_index_size_in_bytes) (packed_uint<8>: the low 32 bits)
    For i = 0 To m_levels_size - 1
        p = SIZEOF_KTX2_HEADER + i * SIZEOF_KTX2_LEVEL_INDEX
        m_levels_byte_offset(i) = read_le_dword(m_pData, p)
        m_levels_byte_length(i) = read_le_dword(m_pData, p + 8)
        m_levels_uncompressed_byte_length(i) = read_le_dword(m_pData, p + 16)
    Next

    ' Sanity check the level offsets and byte sizes
    For i = 0 To m_levels_size - 1
        If ULT_(m_levels_byte_offset(i), SIZEOF_KTX2_HEADER) Then
            ' BASISU_DEVEL_ERROR("ktx2_transcoder::init: Invalid level offset (too low)\n");
            Exit Function
        End If

        If m_levels_byte_length(i) = 0 Then
            ' BASISU_DEVEL_ERROR("ktx2_transcoder::init: Invalid level byte length\n");  (no return in C)
        End If

        If ULT_(m_data_size, m_levels_byte_offset(i) + m_levels_byte_length(i)) Then
            ' BASISU_DEVEL_ERROR("ktx2_transcoder::init: Invalid level offset and/or length\n");
            Exit Function
        End If

        ' const uint64_t MAX_SANE_LEVEL_UNCOMP_SIZE = 2048ULL * 1024ULL * 1024ULL;
        If U32_(m_levels_uncompressed_byte_length(i)) >= 2147483648# Then
            ' BASISU_DEVEL_ERROR("ktx2_transcoder::init: Invalid level offset (too large)\n");
            Exit Function
        End If

        If m_header_supercompression_scheme = KTX2_SS_BASISLZ Then
            If m_levels_uncompressed_byte_length(i) <> 0 Then Exit Function
        ElseIf Not ULT_(m_header_supercompression_scheme, KTX2_SS_ZSTANDARD) Then
            If m_levels_uncompressed_byte_length(i) = 0 Then Exit Function
        End If
    Next

    If (m_header_dfd_byte_length <> DFD_MINIMUM_SIZE) And (m_header_dfd_byte_length <> DFD_MAXIMUM_SIZE) Then
        ' BASISU_DEVEL_ERROR("ktx2_transcoder::init: Unsupported DFD size\n");
        Exit Function
    End If

    If ULT_(m_data_size, m_header_dfd_byte_offset + m_header_dfd_byte_length) Or ULT_(m_header_dfd_byte_offset, SIZEOF_KTX2_HEADER) Then
        ' BASISU_DEVEL_ERROR("ktx2_transcoder::init: Invalid DFD offset and/or length\n");
        Exit Function
    End If

    p = m_header_dfd_byte_offset                        ' pDFD

    ReDim m_dfd(0 To m_header_dfd_byte_length - 1)
    CopyMemory m_dfd(0), m_pData(p), m_header_dfd_byte_length

    ' This is all hard coded for only ETC1S and UASTC.
    dfd_total_size = read_le_dword(m_pData, p)

    ' 3.10.3: Sanity check
    If dfd_total_size <> m_header_dfd_byte_length Then Exit Function

    ' 3.10.3: More sanity checking
    If m_header_kvd_byte_length <> 0 Then
        If dfd_total_size <> m_header_kvd_byte_offset - m_header_dfd_byte_offset Then Exit Function
    End If

    dfd_bits = read_le_dword(m_pData, p + 3 * 4)
    sample_channel0 = read_le_dword(m_pData, p + 7 * 4)

    m_dfd_color_model = dfd_bits And 255
    m_dfd_color_prims = (dfd_bits And &HFF00&) \ &H100&
    m_dfd_transfer_func = (dfd_bits And &HFF0000) \ &H10000
    m_dfd_flags = ((dfd_bits And &H7FFFFFFF) \ &H1000000) Or (((dfd_bits < 0) And 1) * 128)

    ' See 3.10.1.Restrictions
    If (m_dfd_transfer_func <> KTX2_KHR_DF_TRANSFER_LINEAR) And (m_dfd_transfer_func <> KTX2_KHR_DF_TRANSFER_SRGB) Then
        ' BASISU_DEVEL_ERROR("ktx2_transcoder::init: Invalid DFD transfer function\n");
        Exit Function
    End If

    If m_dfd_color_model = KTX2_KDF_DF_MODEL_ETC1S Then
        m_format = cETC1S

        ' 3.10.2: "Whether the image has 1 or 2 slices can be determined from the DFD's sample count."
        ' If m_has_alpha is true it may be 2-channel RRRG or 4-channel RGBA, but we let the caller deal with that.
        m_has_alpha = (m_header_dfd_byte_length = 60)

        If m_has_alpha Then m_dfd_samples = 2 Else m_dfd_samples = 1
        m_dfd_chan0 = (sample_channel0 And &H7FFFFFFF) \ &H1000000 And 15

        If m_has_alpha Then
            sample_channel1 = read_le_dword(m_pData, p + 11 * 4)
            m_dfd_chan1 = (sample_channel1 And &H7FFFFFFF) \ &H1000000 And 15
        End If
    ElseIf m_dfd_color_model = KTX2_KDF_DF_MODEL_UASTC Then
        m_format = cUASTC4x4

        m_dfd_samples = 1
        m_dfd_chan0 = (sample_channel0 And &H7FFFFFFF) \ &H1000000 And 15

        ' We're assuming "DATA" means RGBA so it has alpha.
        m_has_alpha = (m_dfd_chan0 = KTX2_DF_CHANNEL_UASTC_RGBA) Or (m_dfd_chan0 = KTX2_DF_CHANNEL_UASTC_RRRG)
    ElseIf m_dfd_color_model = KTX2_KDF_DF_MODEL_UASTC_HDR Then
        m_format = cUASTC_HDR_4x4

        m_dfd_samples = 1
        m_dfd_chan0 = (sample_channel0 And &H7FFFFFFF) \ &H1000000 And 15

        ' We're assuming "DATA" means RGBA so it has alpha.
        m_has_alpha = (m_dfd_chan0 = KTX2_DF_CHANNEL_UASTC_RGBA) Or (m_dfd_chan0 = KTX2_DF_CHANNEL_UASTC_RRRG)
    Else
        ' Unsupported DFD color model.
        Exit Function
    End If

    If Not read_key_values() Then
        ' BASISU_DEVEL_ERROR("ktx2_transcoder::init: read_key_values() failed\n");
        Exit Function
    End If

    ' Check for a KTXanimData key
    For i = 0 To m_key_values_size - 1
        If m_key_values_key(i) = "KTXanimData" & Chr$(0) Then
            m_is_video = True
            Exit For
        End If
    Next

    init = True
End Function

' bool read_key_values()
Private Function read_key_values() As Boolean
    Dim pSrc As Long, src_left As Long, l As Long, key_data As String, kb As Long, value_data() As Byte
    Dim ofs As Long, alignment_bytes As Long

    If m_header_kvd_byte_length = 0 Then
        If m_header_kvd_byte_offset <> 0 Then
            ' BASISU_DEVEL_ERROR("ktx2_transcoder::read_key_values: Invalid KVD byte offset (it should be zero when the length is zero)\n");
            Exit Function
        End If

        read_key_values = True
        Exit Function
    End If

    If ULT_(m_header_kvd_byte_offset, SIZEOF_KTX2_HEADER) Then Exit Function

    If ULT_(m_data_size, m_header_kvd_byte_offset + m_header_kvd_byte_length) Then Exit Function

    pSrc = m_header_kvd_byte_offset
    src_left = m_header_kvd_byte_length

    Do While ULT_(4, src_left)
        l = read_le_dword(m_pData, pSrc)

        pSrc = pSrc + 4
        src_left = src_left - 4

        If ULT_(l, 2) Then
            ' BASISU_DEVEL_ERROR("ktx2_transcoder::read_key_values: Failed reading key value fields (0)\n");
            Exit Function
        End If

        If ULT_(src_left, l) Then
            ' BASISU_DEVEL_ERROR("ktx2_transcoder::read_key_values: Failed reading key value fields (1)\n");
            Exit Function
        End If

        ' m_key_values.try_resize(m_key_values.size() + 1): key_data and value_data are its new element's
        key_data = ""
        Do
            If l = 0 Then
                ' BASISU_DEVEL_ERROR("ktx2_transcoder::read_key_values: Failed reading key value fields (2)\n");
                Exit Function
            End If

            kb = m_pData(pSrc)
            key_data = key_data & ChrW$(kb)
            pSrc = pSrc + 1

            src_left = src_left - 1
            l = l - 1
        Loop While kb <> 0

        If l <> 0 Then
            ReDim value_data(0 To l - 1)
            CopyMemory value_data(0), m_pData(pSrc), l
            pSrc = pSrc + l
            src_left = src_left - l
        Else
            Erase value_data
        End If

        ReDim Preserve m_key_values_key(0 To m_key_values_size)
        ReDim Preserve m_key_values_value(0 To m_key_values_size)
        m_key_values_key(m_key_values_size) = key_data
        m_key_values_value(m_key_values_size) = value_data
        m_key_values_size = m_key_values_size + 1

        ofs = pSrc And 3
        alignment_bytes = (4 - ofs) And 3

        If ULT_(src_left, alignment_bytes) Then
            ' BASISU_DEVEL_ERROR("ktx2_transcoder::read_key_values: Failed reading key value fields (3)\n");
            Exit Function
        End If

        pSrc = pSrc + alignment_bytes
        src_left = src_left - alignment_bytes
    Loop

    read_key_values = True
End Function

' const basisu::uint8_vec *find_key(const std::string& key_name) const: the value (a Byte array, Empty when the key
' is absent; an uninitialised array for an empty value)
Public Function find_key(ByVal key_name As String) As Variant
    Dim i As Long
    For i = 0 To m_key_values_size - 1
        If m_key_values_key(i) = key_name & Chr$(0) Then
            find_key = m_key_values_value(i)
            Exit Function
        End If
    Next
End Function

' bool start_transcoding()
Public Function start_transcoding() As Boolean
    If Not m_has_data Then
        ' BASISU_DEVEL_ERROR("ktx2_transcoder::start_transcoding: Must call init() first\n");
        Exit Function
    End If

    If m_header_supercompression_scheme = KTX2_SS_BASISLZ Then
        ' decompress_etc1s_global_data(): ETC1S is not ported (fails, as when the global data cannot be read)
        ' BASISU_DEVEL_ERROR("ktx2_transcoder::start_transcoding: decompress_etc1s_global_data() failed\n");
        Exit Function
    ElseIf m_header_supercompression_scheme = KTX2_SS_ZSTANDARD Then
        ' (BASISD_SUPPORT_KTX2_ZSTD is 1)
    End If

    start_transcoding = True
End Function

' uint32_t get_width() const
Public Function get_width() As Long
    get_width = m_header_pixel_width
End Function

' uint32_t get_height() const
Public Function get_height() As Long
    get_height = m_header_pixel_height
End Function

' uint32_t get_levels() const
Public Function get_levels() As Long
    get_levels = m_header_level_count
End Function

' uint32_t get_layers() const
Public Function get_layers() As Long
    get_layers = m_header_layer_count
End Function

' uint32_t get_faces() const
Public Function get_faces() As Long
    get_faces = m_header_face_count
End Function

' basist::basis_tex_format get_format() const (BASISD_LIB_VERSION 150: 4.7.0 calls get_format)
Public Function get_format() As basis_tex_format
    get_format = m_format
End Function

' bool is_etc1s() const
Public Function is_etc1s() As Boolean
    is_etc1s = (get_format() = cETC1S)
End Function

' bool is_uastc() const
Public Function is_uastc() As Boolean
    is_uastc = (get_format() = cUASTC4x4)
End Function

' bool get_has_alpha() const
Public Function get_has_alpha() As Boolean
    get_has_alpha = m_has_alpha
End Function

' ktx2_df_channel_id get_dfd_channel_id0() const
Public Function get_dfd_channel_id0() As ktx2_df_channel_id
    get_dfd_channel_id0 = m_dfd_chan0
End Function

' ktx2_df_channel_id get_dfd_channel_id1() const
Public Function get_dfd_channel_id1() As ktx2_df_channel_id
    get_dfd_channel_id1 = m_dfd_chan1
End Function

' uint32_t get_dfd_transfer_func() const
Public Function get_dfd_transfer_func() As Long
    get_dfd_transfer_func = m_dfd_transfer_func
End Function

' (VB) maximum(m_header.m_layer_count, 1)
Private Function layer_count_1() As Long
    If m_header_layer_count = 0 Then layer_count_1 = 1 Else layer_count_1 = m_header_layer_count
End Function

' (VB) maximum<uint32_t>(m_header.m_pixel_width >> level_index, 1) (level_index < 16)
Private Function level_dim(ByVal pixels As Long, ByVal level_index As Long) As Long
    Dim v As Long, i As Long
    v = pixels
    If level_index > 0 Then
        If v < 0 Then v = ((v And &H7FFFFFFF) \ 2) Or &H40000000 Else v = v \ 2
        For i = 2 To level_index
            v = v \ 2
        Next
    End If
    If v = 0 Then v = 1
    level_dim = v
End Function

' (VB) the number of elements of b, -1 when it has none (C: a null pointer)
Private Function array_size(b() As Byte) As Long
    On Error GoTo none
    array_size = UBound(b) - LBound(b) + 1
    Exit Function
none:
    array_size = -1
End Function

' bool get_image_level_info(ktx2_image_level_info& level_info, uint32_t level_index, uint32_t layer_index, uint32_t face_index) const
Public Function get_image_level_info(ByVal level_info As ktx2_image_level_info, ByVal level_index As Long, ByVal layer_index As Long, ByVal face_index As Long) As Boolean
    Dim level_width As Long, level_height As Long, num_blocks_x As Long, num_blocks_y As Long

    If Not ULT_(level_index, m_levels_size) Then
        ' BASISU_DEVEL_ERROR("ktx2_transcoder::get_image_level_info: level_index >= m_levels.size()\n");
        Exit Function
    End If

    If m_header_face_count > 1 Then
        If Not ULT_(face_index, 6) Then Exit Function
    ElseIf face_index <> 0 Then
        Exit Function
    End If

    If Not ULT_(layer_index, layer_count_1()) Then Exit Function

    level_width = level_dim(m_header_pixel_width, level_index)
    level_height = level_dim(m_header_pixel_height, level_index)
    num_blocks_x = (level_width + 3) \ 4
    num_blocks_y = (level_height + 3) \ 4

    level_info.m_face_index = face_index
    level_info.m_layer_index = layer_index
    level_info.m_level_index = level_index
    level_info.m_orig_width = level_width
    level_info.m_orig_height = level_height
    level_info.m_width = num_blocks_x * 4
    level_info.m_height = num_blocks_y * 4
    level_info.m_num_blocks_x = num_blocks_x
    level_info.m_num_blocks_y = num_blocks_y
    level_info.m_total_blocks = num_blocks_x * num_blocks_y
    level_info.m_alpha_flag = m_has_alpha
    level_info.m_iframe_flag = False
    ' (m_etc1s_image_descs: always empty, ETC1S is not ported)

    get_image_level_info = True
End Function

' bool transcode_image_level(uint32_t level_index, uint32_t layer_index, uint32_t face_index, void* pOutput_blocks,
'     uint32_t output_blocks_buf_size_in_blocks_or_pixels, basist::transcoder_texture_format fmt, uint32_t decode_flags = 0,
'     uint32_t output_row_pitch_in_blocks_or_pixels = 0, uint32_t output_rows_in_pixels = 0, int channel0 = -1, int channel1 = -1,
'     ktx2_transcoder_state* pState = nullptr)
' pOutput_blocks is output(output_offset ...) (the caller's Byte buffer); the state is the object's own.
Public Function transcode_image_level(ByVal level_index As Long, ByVal layer_index As Long, ByVal face_index As Long, output() As Byte, ByVal output_offset As Long, ByVal output_blocks_buf_size_in_blocks_or_pixels As Long, ByVal fmt As transcoder_texture_format, Optional ByVal decode_flags As Long = 0, Optional ByVal output_row_pitch_in_blocks_or_pixels As Long = 0, Optional ByVal output_rows_in_pixels As Long = 0, Optional ByVal channel0 As Long = -1, Optional ByVal channel1 As Long = -1) As Boolean
    Dim comp_level_data_ofs As Long, comp_level_data_size As Double, uncomp_level_data_size As Double
    Dim level_width As Long, level_height As Long, num_blocks_x As Long, num_blocks_y As Long
    Dim total_2D_image_size As Long, uncomp_ofs As Long

    If Not m_has_data Then
        ' BASISU_DEVEL_ERROR("ktx2_transcoder::transcode_image_2D: Must call init() first\n");
        Exit Function
    End If

    ' if (!pState) pState = &m_def_transcoder_state;

    If Not ULT_(level_index, m_levels_size) Then Exit Function

    If m_header_face_count > 1 Then
        If Not ULT_(face_index, 6) Then Exit Function
    ElseIf face_index <> 0 Then
        Exit Function
    End If

    If Not ULT_(layer_index, layer_count_1()) Then Exit Function

    comp_level_data_ofs = m_levels_byte_offset(level_index)          ' pComp_level_data
    comp_level_data_size = U32_(m_levels_byte_length(level_index))

    uncomp_level_data_size = comp_level_data_size                     ' pUncomp_level_data = pComp_level_data

    If m_header_supercompression_scheme = KTX2_SS_ZSTANDARD Then
        ' Check if we've already decompressed this level's supercompressed data.
        If level_index <> m_uncomp_data_level_index Then
            ' Uncompress the entire level's supercompressed data.
            If Not decompress_level_data(level_index) Then
                ' BASISU_DEVEL_ERROR("ktx2_transcoder::transcode_image_2D: decompress_level_data() failed\n");
                Exit Function
            End If
            m_uncomp_data_level_index = level_index
        End If

        uncomp_level_data_size = m_level_uncomp_data_size
    End If

    level_width = level_dim(m_header_pixel_width, level_index)
    level_height = level_dim(m_header_pixel_height, level_index)
    num_blocks_x = (level_width + 3) \ 4
    num_blocks_y = (level_height + 3) \ 4

    If m_format = cETC1S Then
        ' Ensure start_transcoding() was called. (ETC1S is not ported: its endpoints are always empty)
        ' BASISU_DEVEL_ERROR("ktx2_transcoder::transcode_image_2D: must call start_transcoding() first\n");
        Exit Function
    ElseIf (m_format = cUASTC4x4) Or (m_format = cUASTC_HDR_4x4) Then
        ' Compute length and offset to uncompressed 2D UASTC texture data, given the face/layer indices.
        total_2D_image_size = num_blocks_x * num_blocks_y * KTX2_UASTC_BLOCK_SIZE

        uncomp_ofs = (layer_index * m_header_face_count + face_index) * total_2D_image_size

        ' Sanity checks
        If U32_(uncomp_ofs) >= uncomp_level_data_size Then Exit Function

        If (uncomp_level_data_size - U32_(uncomp_ofs)) < U32_(total_2D_image_size) Then Exit Function

        If m_format = cUASTC_HDR_4x4 Then
            ' m_uastc_hdr_transcoder.transcode_image: UASTC HDR is not ported
            ' BASISU_DEVEL_ERROR("ktx2_transcoder::transcode_image_2D: UASTC HDR transcode_image() failed, this is either a bug or the file is corrupted/invalid\n");
            Exit Function
        Else
            If m_header_supercompression_scheme = KTX2_SS_ZSTANDARD Then
                If Not basisu_uastc_transcode_image(fmt, output, output_offset, output_blocks_buf_size_in_blocks_or_pixels, m_level_uncomp_data, uncomp_ofs, total_2D_image_size, num_blocks_x, num_blocks_y, level_width, level_height, level_index, 0, total_2D_image_size, decode_flags, m_has_alpha, m_is_video, output_row_pitch_in_blocks_or_pixels, output_rows_in_pixels, channel0, channel1) Then
                    ' BASISU_DEVEL_ERROR("ktx2_transcoder::transcode_image_2D: UASTC transcode_image() failed, this is either a bug or the file is corrupted/invalid\n");
                    Exit Function
                End If
            Else
                If Not basisu_uastc_transcode_image(fmt, output, output_offset, output_blocks_buf_size_in_blocks_or_pixels, m_pData, comp_level_data_ofs + uncomp_ofs, total_2D_image_size, num_blocks_x, num_blocks_y, level_width, level_height, level_index, 0, total_2D_image_size, decode_flags, m_has_alpha, m_is_video, output_row_pitch_in_blocks_or_pixels, output_rows_in_pixels, channel0, channel1) Then
                    Exit Function
                End If
            End If
        End If
    Else
        ' Shouldn't get here.
        Exit Function
    End If

    transcode_image_level = True
End Function

' bool decompress_level_data(uint32_t level_index, basisu::uint8_vec& uncomp_data) (uncomp_data: the state's
' m_level_uncomp_data)
Private Function decompress_level_data(ByVal level_index As Long) As Boolean
    Dim comp_ofs As Long, comp_size As Long, uncomp_size As Long, actualUncompSize As Long

    comp_ofs = m_levels_byte_offset(level_index)                     ' pComp_data
    comp_size = m_levels_byte_length(level_index)

    uncomp_size = m_levels_uncompressed_byte_length(level_index)

    ' ((size_t)comp_size) != comp_size, ((size_t)uncomp_size) != uncomp_size: never (uint32_t values)

    ' uncomp_data.try_resize((size_t)uncomp_size)
    If uncomp_size > 0 Then
        ReDim m_level_uncomp_data(0 To uncomp_size - 1)
    Else
        Erase m_level_uncomp_data
    End If
    m_level_uncomp_data_size = uncomp_size

    If m_header_supercompression_scheme = KTX2_SS_ZSTANDARD Then
        ' #if BASISD_SUPPORT_KTX2_ZSTD
        actualUncompSize = ZSTD_decompress(m_level_uncomp_data, 0, uncomp_size, m_pData, comp_ofs, comp_size)
        If ZSTD_isError(actualUncompSize) Then
            ' BASISU_DEVEL_ERROR("ktx2_transcoder::decompress_level_data: Zstd decompression failed, file is invalid or corrupted\n");
            Exit Function
        End If
        If actualUncompSize <> uncomp_size Then
            ' BASISU_DEVEL_ERROR("ktx2_transcoder::decompress_level_data: Zstd decompression returned too few bytes, file is invalid or corrupted\n");
            Exit Function
        End If
    End If

    decompress_level_data = True
End Function
