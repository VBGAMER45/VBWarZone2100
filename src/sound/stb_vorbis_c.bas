Attribute VB_Name = "stb_vorbis_c"
' ref/stb_vorbis/stb_vorbis.c: stb_vorbis v1.22 (Sean Barrett and contributors, public domain), the Ogg
' Vorbis decoder WzSound uses in place of 4.7.0's libvorbis (PLAN, "Sound (WzSound.dll)"). Configured as
' STB_VORBIS_NO_STDIO and STB_VORBIS_NO_PUSHDATA_API: the pull API on a file in memory. Not ported:
' stb_vorbis_seek and seek_frame with their helpers (seek_to_sample_coarse, get_seek_page_info,
' go_to_page_before, peek_decode_initial; WzSound only rewinds with stb_vorbis_seek_start), the float
' output API (get_samples_float*), get_comment and get_sample_offset, the pushdata API, the stdio API.
' (VB) A stb_vorbis * is a handle: the decoder pool's index + 1 (0 = NULL); the internal functions take
' the decoder (vorb) ByRef. A uint32 is a Long holding its bits. A pointer into a buffer is the buffer
' and an index; a float * of the decoder (channel_buffers[i], ...) is a FloatArr. The data is copied into
' the decoder (stream()); stream - stream_start is stream_pos. The C's asserts are comments (stb_vorbis
' builds with NDEBUG in a release build). FAST_SCALED_FLOAT_TO_INT rounds x * 32768 to the nearest
' integer, ties to even: CLng does that too. Floats are VB Singles (the x87 rounds the intermediate
' results more precisely than the C's SSE floats: the decoded samples may differ by a unit or so).
' Interface for WzSound: H_DESIGN.md's "Fork H1's API"; tools/run_tests.py --sound checks every game
' sound against ffmpeg's decoder.
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (dest As Any, ByVal numBytes As Long)

'//////   ERROR CODES

Public Enum STBVorbisError
    VORBIS__no_error

    VORBIS_need_more_data = 1               ' not a real error

    VORBIS_invalid_api_mixing               ' can't mix API modes
    VORBIS_outofmem                         ' not enough memory
    VORBIS_feature_not_supported            ' uses floor 0
    VORBIS_too_many_channels                ' STB_VORBIS_MAX_CHANNELS is too small
    VORBIS_file_open_failure                ' fopen() failed
    VORBIS_seek_without_length              ' can't seek in unknown-length file

    VORBIS_unexpected_eof = 10              ' file is truncated?
    VORBIS_seek_invalid                     ' seek past EOF

    ' decoding errors (corrupt/invalid stream) -- you probably
    ' don't care about the exact details of these

    ' vorbis errors:
    VORBIS_invalid_setup = 20
    VORBIS_invalid_stream

    ' ogg errors:
    VORBIS_missing_capture_pattern = 30
    VORBIS_invalid_stream_structure_version
    VORBIS_continued_packet_flag_invalid
    VORBIS_incorrect_stream_serial_number
    VORBIS_invalid_first_page
    VORBIS_bad_packet_type
    VORBIS_cant_find_last_page
    VORBIS_seek_failed
    VORBIS_ogg_skeleton_not_supported
End Enum

Public Type stb_vorbis_info
    sample_rate As Long                     ' unsigned int
    channels As Long

    setup_memory_required As Long           ' unsigned int
    setup_temp_memory_required As Long      ' unsigned int
    temp_memory_required As Long            ' unsigned int

    max_frame_size As Long
End Type

Private Const STB_VORBIS_MAX_CHANNELS As Long = 16      ' enough for anyone?
Private Const STB_VORBIS_FAST_HUFFMAN_LENGTH As Long = 10
Private Const FAST_HUFFMAN_TABLE_SIZE As Long = 1024     ' (1 << STB_VORBIS_FAST_HUFFMAN_LENGTH)
Private Const FAST_HUFFMAN_TABLE_MASK As Long = 1023     ' (FAST_HUFFMAN_TABLE_SIZE - 1)

' code length assigned to a value with no huffman encoding
Private Const NO_CODE As Long = 255

Private Const CRC32_POLY As Long = &H4C11DB7   ' from spec

Private Const PAGEFLAG_continued_packet As Long = 1
Private Const PAGEFLAG_first_page As Long = 2
Private Const PAGEFLAG_last_page As Long = 4

Private Const EOP As Long = -1
Private Const INVALID_BITS As Long = -1

Private Const VORBIS_packet_id As Long = 1
Private Const VORBIS_packet_comment As Long = 3
Private Const VORBIS_packet_setup As Long = 5

Private Const SAMPLE_unknown As Long = &HFFFFFFFF

Private Const PLAYBACK_MONO As Long = 1
Private Const PLAYBACK_LEFT As Long = 2
Private Const PLAYBACK_RIGHT As Long = 4

Private Const M_PI As Double = 3.14159265358979     ' math.h's (double)

' (VB) a float * the decoder allocates (setup_malloc): its own array
Private Type FloatArr
    v() As Single
End Type

' (VB) an int16 * (finalY), a uint16 * (bit_reverse) as Longs
Private Type IntArr
    v() As Integer
End Type

Private Type LongArr
    v() As Long
End Type

Private Type CodebookT
    dimensions As Long
    entries As Long
    codeword_lengths() As Byte              ' uint8 *
    minimum_value As Single
    delta_value As Single
    value_bits As Long                      ' uint8
    lookup_type As Long                     ' uint8
    sequence_p As Long                      ' uint8
    sparse As Long                          ' uint8
    lookup_values As Long                   ' uint32
    multiplicands() As Single               ' codetype *
    codewords() As Long                     ' uint32 *
    has_codewords As Boolean                ' (VB) codewords != NULL
    fast_huffman(0 To FAST_HUFFMAN_TABLE_SIZE - 1) As Integer   ' STB_VORBIS_FAST_HUFFMAN_SHORT: int16
    sorted_codewords() As Long              ' uint32 *
    has_sorted As Boolean                   ' (VB) sorted_codewords != NULL
    sorted_values() As Long                 ' int * (-1 To sorted_entries - 1: sorted_values[-1] is defined)
    sorted_entries As Long
End Type

Private Type Floor0T
    order As Long                           ' uint8
    rate As Long                            ' uint16
    bark_map_size As Long                   ' uint16
    amplitude_bits As Long                  ' uint8
    amplitude_offset As Long                ' uint8
    number_of_books As Long                 ' uint8
    book_list(0 To 15) As Byte              ' varies
End Type

Private Type Floor1T
    partitions As Long                      ' uint8
    partition_class_list(0 To 31) As Byte   ' varies
    class_dimensions(0 To 15) As Byte       ' varies
    class_subclasses(0 To 15) As Byte       ' varies
    class_masterbooks(0 To 15) As Byte      ' varies
    subclass_books(0 To 15, 0 To 7) As Integer  ' int16: varies
    Xlist(0 To 31 * 8 + 1) As Long          ' uint16: varies
    sorted_order(0 To 31 * 8 + 1) As Byte
    neighbors(0 To 31 * 8 + 1, 0 To 1) As Byte
    floor1_multiplier As Long               ' uint8
    rangebits As Long                       ' uint8
    values As Long
End Type

' typedef union Floor (VB: both members)
Private Type FloorT
    floor0 As Floor0T
    floor1 As Floor1T
End Type

Private Type ResidueT
    begin_ As Long                          ' uint32 begin, end
    end_ As Long
    part_size As Long                       ' uint32
    classifications As Long                 ' uint8
    classbook As Long                       ' uint8
    classdata() As Byte                     ' uint8 **: (VB) classdata[j][k] is classdata(j * classwords + k)
    classwords As Long                      ' (VB) the classbook's dimensions: classdata's row length
    residue_books() As Integer              ' int16 (*)[8]: residue_books(classifications - 1, 7)
End Type

Private Type MappingChannelT
    magnitude As Long                       ' uint8
    angle As Long                           ' uint8
    mux As Long                             ' uint8
End Type

Private Type MappingT
    coupling_steps As Long                  ' uint16
    chan() As MappingChannelT
    submaps As Long                         ' uint8
    submap_floor(0 To 14) As Byte           ' varies
    submap_residue(0 To 14) As Byte         ' varies
End Type

Private Type ModeT
    blockflag As Long                       ' uint8
    mapping As Long                         ' uint8
    windowtype As Long                      ' uint16
    transformtype As Long                   ' uint16
End Type

Private Type ProbedPageT
    page_start As Long                      ' uint32
    page_end As Long                        ' uint32
    last_decoded_sample As Long             ' uint32
End Type

' struct stb_vorbis (vorb)
Private Type vorb
    ' user-accessible info
    sample_rate As Long                     ' unsigned int
    channels As Long

    setup_memory_required As Long           ' unsigned int
    temp_memory_required As Long            ' unsigned int
    setup_temp_memory_required As Long      ' unsigned int

    vendor As String
    comment_list_length As Long
    comment_list() As String

    ' input config
    stream() As Byte                        ' (VB) the data: stream_start is stream(0), stream_end stream(stream_len)
    stream_pos As Long                      ' (VB) stream - stream_start
    stream_len As Long                      ' uint32

    ' the page to seek to when seeking to start, may be zero
    first_audio_page_offset As Long         ' uint32

    ' p_first is the page on which the first audio packet ends
    ' (but not necessarily the page on which it starts)
    p_first As ProbedPageT
    p_last As ProbedPageT

    ' run-time results
    eof_ As Long
    error_ As Long                          ' enum STBVorbisError

    ' header info
    blocksize(0 To 1) As Long
    blocksize_0 As Long
    blocksize_1 As Long
    codebook_count As Long
    codebooks() As CodebookT
    floor_count As Long
    floor_types(0 To 63) As Long            ' uint16: varies
    floor_config() As FloorT
    residue_count As Long
    residue_types(0 To 63) As Long          ' uint16: varies
    residue_config() As ResidueT
    mapping_count As Long
    mapping() As MappingT
    mode_count As Long
    mode_config(0 To 63) As ModeT           ' varies

    total_samples As Long                   ' uint32

    ' decode buffer
    channel_buffers(0 To STB_VORBIS_MAX_CHANNELS - 1) As FloatArr
    ' (outputs[i] = channel_buffers[i] + channel_buffer_start: get_frame_float's left)

    previous_window(0 To STB_VORBIS_MAX_CHANNELS - 1) As FloatArr
    previous_length As Long

    finalY(0 To STB_VORBIS_MAX_CHANNELS - 1) As IntArr   ' int16 * (!STB_VORBIS_NO_DEFER_FLOOR)

    current_loc As Long                     ' uint32: sample location of next frame to decode
    current_loc_valid As Long

    ' per-blocksize precomputed data

    ' twiddle factors
    A(0 To 1) As FloatArr
    B(0 To 1) As FloatArr
    C(0 To 1) As FloatArr
    window_(0 To 1) As FloatArr
    bit_reverse(0 To 1) As LongArr          ' uint16 *

    ' current page/packet/segment streaming info
    serial As Long                          ' uint32: stream serial number for verification
    last_page As Long
    segment_count As Long
    segments(0 To 254) As Byte
    page_flag As Long                       ' uint8
    bytes_in_seg As Long                    ' uint8
    first_decode As Long                    ' uint8
    next_seg As Long
    last_seg As Long                        ' flag that we're on the last segment
    last_seg_which As Long                  ' what was the segment number of the last seg?
    acc As Long                             ' uint32
    valid_bits As Long
    packet_bytes As Long
    end_seg_with_known_loc As Long
    known_loc_for_packet As Long            ' uint32
    discard_samples_deferred As Long
    samples_output As Long                  ' uint32

    ' push mode scanning
    page_crc_tests As Long                  ' only in push_mode: number of tests active; -1 if not searching

    ' sample-access
    channel_buffer_start As Long
    channel_buffer_end As Long

    in_use As Boolean                       ' (VB) the pool's slot holds a decoder
End Type

' (VB) the decoders: a stb_vorbis * is an index + 1
Private pool() As vorb
Private poolCount As Long

' (VB) the tables: 1 << n (pow2tab(31) is &H80000000), (1 << n) - 1, the bits of a byte reversed
Private pow2tab(0 To 31) As Long
Private masktab(0 To 32) As Long
Private rev8(0 To 255) As Long
Private tablesReady As Boolean

Private crc_table(0 To 255) As Long         ' uint32
Private log2_4(0 To 15) As Long             ' ilog's static signed char log2_4[16]
Private inverse_db_table(0 To 255) As Single
Private ogg_page_header(0 To 3) As Long
Private channel_position(0 To 6, 0 To 5) As Long    ' int8
Private range_list(0 To 3) As Long          ' vorbis_decode_packet_rest's static int range_list[4]

' (VB) temp_alloc's buffers: inverse_mdct's buf2, decode_residue's part_classdata
Private buf2() As Single
Private buf2Size As Long
Private part_classdata() As Long            ' (VB) part_classdata[j][class_set] is the classdata row: (j * part_read + class_set)
Private part_classdataSize As Long

' (VB) the module's tables (the C's static initialisers)
Private Sub stb_tables()
    Dim i As Long, k As Long, b As Long
    pow2tab(0) = 1
    For i = 1 To 30
        pow2tab(i) = pow2tab(i - 1) * 2
    Next
    pow2tab(31) = &H80000000
    For i = 0 To 30
        masktab(i) = pow2tab(i) - 1
    Next
    masktab(31) = &H7FFFFFFF
    masktab(32) = -1
    For i = 0 To 255
        b = 0
        For k = 0 To 7
            If (i And pow2tab(k)) <> 0 Then b = b Or pow2tab(7 - k)
        Next
        rev8(i) = b
    Next
    log2_4(0) = 0: log2_4(1) = 1: log2_4(2) = 2: log2_4(3) = 2
    For i = 4 To 7: log2_4(i) = 3: Next
    For i = 8 To 15: log2_4(i) = 4: Next
    ogg_page_header(0) = &H4F: ogg_page_header(1) = &H67: ogg_page_header(2) = &H67: ogg_page_header(3) = &H53
    ' channel_position: L = LEFT | MONO, C = LEFT | RIGHT | MONO, R = RIGHT | MONO
    channel_position(1, 0) = PLAYBACK_LEFT Or PLAYBACK_RIGHT Or PLAYBACK_MONO
    channel_position(2, 0) = PLAYBACK_LEFT Or PLAYBACK_MONO: channel_position(2, 1) = PLAYBACK_RIGHT Or PLAYBACK_MONO
    channel_position(3, 0) = PLAYBACK_LEFT Or PLAYBACK_MONO: channel_position(3, 1) = PLAYBACK_LEFT Or PLAYBACK_RIGHT Or PLAYBACK_MONO: channel_position(3, 2) = PLAYBACK_RIGHT Or PLAYBACK_MONO
    channel_position(4, 0) = PLAYBACK_LEFT Or PLAYBACK_MONO: channel_position(4, 1) = PLAYBACK_RIGHT Or PLAYBACK_MONO: channel_position(4, 2) = PLAYBACK_LEFT Or PLAYBACK_MONO: channel_position(4, 3) = PLAYBACK_RIGHT Or PLAYBACK_MONO
    channel_position(5, 0) = PLAYBACK_LEFT Or PLAYBACK_MONO: channel_position(5, 1) = PLAYBACK_LEFT Or PLAYBACK_RIGHT Or PLAYBACK_MONO: channel_position(5, 2) = PLAYBACK_RIGHT Or PLAYBACK_MONO: channel_position(5, 3) = PLAYBACK_LEFT Or PLAYBACK_MONO: channel_position(5, 4) = PLAYBACK_RIGHT Or PLAYBACK_MONO
    channel_position(6, 0) = PLAYBACK_LEFT Or PLAYBACK_MONO: channel_position(6, 1) = PLAYBACK_LEFT Or PLAYBACK_RIGHT Or PLAYBACK_MONO: channel_position(6, 2) = PLAYBACK_RIGHT Or PLAYBACK_MONO: channel_position(6, 3) = PLAYBACK_LEFT Or PLAYBACK_MONO: channel_position(6, 4) = PLAYBACK_RIGHT Or PLAYBACK_MONO: channel_position(6, 5) = PLAYBACK_LEFT Or PLAYBACK_RIGHT Or PLAYBACK_MONO
    range_list(0) = 256: range_list(1) = 128: range_list(2) = 86: range_list(3) = 64
    inverse_db_table_init
    tablesReady = True
End Sub

' the following table is block-copied from the specification
Private Sub inverse_db_table_init()
    inverse_db_table(0) = 1.0649863e-07: inverse_db_table(1) = 1.1341951e-07: inverse_db_table(2) = 1.2079015e-07: inverse_db_table(3) = 1.2863978e-07
    inverse_db_table(4) = 1.3699951e-07: inverse_db_table(5) = 1.4590251e-07: inverse_db_table(6) = 1.5538408e-07: inverse_db_table(7) = 1.6548181e-07
    inverse_db_table(8) = 1.7623575e-07: inverse_db_table(9) = 1.8768855e-07: inverse_db_table(10) = 1.9988561e-07: inverse_db_table(11) = 2.1287530e-07
    inverse_db_table(12) = 2.2670913e-07: inverse_db_table(13) = 2.4144197e-07: inverse_db_table(14) = 2.5713223e-07: inverse_db_table(15) = 2.7384213e-07
    inverse_db_table(16) = 2.9163793e-07: inverse_db_table(17) = 3.1059021e-07: inverse_db_table(18) = 3.3077411e-07: inverse_db_table(19) = 3.5226968e-07
    inverse_db_table(20) = 3.7516214e-07: inverse_db_table(21) = 3.9954229e-07: inverse_db_table(22) = 4.2550680e-07: inverse_db_table(23) = 4.5315863e-07
    inverse_db_table(24) = 4.8260743e-07: inverse_db_table(25) = 5.1396998e-07: inverse_db_table(26) = 5.4737065e-07: inverse_db_table(27) = 5.8294187e-07
    inverse_db_table(28) = 6.2082472e-07: inverse_db_table(29) = 6.6116941e-07: inverse_db_table(30) = 7.0413592e-07: inverse_db_table(31) = 7.4989464e-07
    inverse_db_table(32) = 7.9862701e-07: inverse_db_table(33) = 8.5052630e-07: inverse_db_table(34) = 9.0579828e-07: inverse_db_table(35) = 9.6466216e-07
    inverse_db_table(36) = 1.0273513e-06: inverse_db_table(37) = 1.0941144e-06: inverse_db_table(38) = 1.1652161e-06: inverse_db_table(39) = 1.2409384e-06
    inverse_db_table(40) = 1.3215816e-06: inverse_db_table(41) = 1.4074654e-06: inverse_db_table(42) = 1.4989305e-06: inverse_db_table(43) = 1.5963394e-06
    inverse_db_table(44) = 1.7000785e-06: inverse_db_table(45) = 1.8105592e-06: inverse_db_table(46) = 1.9282195e-06: inverse_db_table(47) = 2.0535261e-06
    inverse_db_table(48) = 2.1869758e-06: inverse_db_table(49) = 2.3290978e-06: inverse_db_table(50) = 2.4804557e-06: inverse_db_table(51) = 2.6416497e-06
    inverse_db_table(52) = 2.8133190e-06: inverse_db_table(53) = 2.9961443e-06: inverse_db_table(54) = 3.1908506e-06: inverse_db_table(55) = 3.3982101e-06
    inverse_db_table(56) = 3.6190449e-06: inverse_db_table(57) = 3.8542308e-06: inverse_db_table(58) = 4.1047004e-06: inverse_db_table(59) = 4.3714470e-06
    inverse_db_table(60) = 4.6555282e-06: inverse_db_table(61) = 4.9580707e-06: inverse_db_table(62) = 5.2802740e-06: inverse_db_table(63) = 5.6234160e-06
    inverse_db_table(64) = 5.9888572e-06: inverse_db_table(65) = 6.3780469e-06: inverse_db_table(66) = 6.7925283e-06: inverse_db_table(67) = 7.2339451e-06
    inverse_db_table(68) = 7.7040476e-06: inverse_db_table(69) = 8.2047000e-06: inverse_db_table(70) = 8.7378876e-06: inverse_db_table(71) = 9.3057248e-06
    inverse_db_table(72) = 9.9104632e-06: inverse_db_table(73) = 1.0554501e-05: inverse_db_table(74) = 1.1240392e-05: inverse_db_table(75) = 1.1970856e-05
    inverse_db_table(76) = 1.2748789e-05: inverse_db_table(77) = 1.3577278e-05: inverse_db_table(78) = 1.4459606e-05: inverse_db_table(79) = 1.5399272e-05
    inverse_db_table(80) = 1.6400004e-05: inverse_db_table(81) = 1.7465768e-05: inverse_db_table(82) = 1.8600792e-05: inverse_db_table(83) = 1.9809576e-05
    inverse_db_table(84) = 2.1096914e-05: inverse_db_table(85) = 2.2467911e-05: inverse_db_table(86) = 2.3928002e-05: inverse_db_table(87) = 2.5482978e-05
    inverse_db_table(88) = 2.7139006e-05: inverse_db_table(89) = 2.8902651e-05: inverse_db_table(90) = 3.0780908e-05: inverse_db_table(91) = 3.2781225e-05
    inverse_db_table(92) = 3.4911534e-05: inverse_db_table(93) = 3.7180282e-05: inverse_db_table(94) = 3.9596466e-05: inverse_db_table(95) = 4.2169667e-05
    inverse_db_table(96) = 4.4910090e-05: inverse_db_table(97) = 4.7828601e-05: inverse_db_table(98) = 5.0936773e-05: inverse_db_table(99) = 5.4246931e-05
    inverse_db_table(100) = 5.7772202e-05: inverse_db_table(101) = 6.1526565e-05: inverse_db_table(102) = 6.5524908e-05: inverse_db_table(103) = 6.9783085e-05
    inverse_db_table(104) = 7.4317983e-05: inverse_db_table(105) = 7.9147585e-05: inverse_db_table(106) = 8.4291040e-05: inverse_db_table(107) = 8.9768747e-05
    inverse_db_table(108) = 9.5602426e-05: inverse_db_table(109) = 0.00010181521: inverse_db_table(110) = 0.00010843174: inverse_db_table(111) = 0.00011547824
    inverse_db_table(112) = 0.00012298267: inverse_db_table(113) = 0.00013097477: inverse_db_table(114) = 0.00013948625: inverse_db_table(115) = 0.00014855085
    inverse_db_table(116) = 0.00015820453: inverse_db_table(117) = 0.00016848555: inverse_db_table(118) = 0.00017943469: inverse_db_table(119) = 0.00019109536
    inverse_db_table(120) = 0.00020351382: inverse_db_table(121) = 0.00021673929: inverse_db_table(122) = 0.00023082423: inverse_db_table(123) = 0.00024582449
    inverse_db_table(124) = 0.00026179955: inverse_db_table(125) = 0.00027881276: inverse_db_table(126) = 0.00029693158: inverse_db_table(127) = 0.00031622787
    inverse_db_table(128) = 0.00033677814: inverse_db_table(129) = 0.00035866388: inverse_db_table(130) = 0.00038197188: inverse_db_table(131) = 0.00040679456
    inverse_db_table(132) = 0.00043323036: inverse_db_table(133) = 0.00046138411: inverse_db_table(134) = 0.00049136745: inverse_db_table(135) = 0.00052329927
    inverse_db_table(136) = 0.00055730621: inverse_db_table(137) = 0.00059352311: inverse_db_table(138) = 0.00063209358: inverse_db_table(139) = 0.00067317058
    inverse_db_table(140) = 0.00071691700: inverse_db_table(141) = 0.00076350630: inverse_db_table(142) = 0.00081312324: inverse_db_table(143) = 0.00086596457
    inverse_db_table(144) = 0.00092223983: inverse_db_table(145) = 0.00098217216: inverse_db_table(146) = 0.0010459992: inverse_db_table(147) = 0.0011139742
    inverse_db_table(148) = 0.0011863665: inverse_db_table(149) = 0.0012634633: inverse_db_table(150) = 0.0013455702: inverse_db_table(151) = 0.0014330129
    inverse_db_table(152) = 0.0015261382: inverse_db_table(153) = 0.0016253153: inverse_db_table(154) = 0.0017309374: inverse_db_table(155) = 0.0018434235
    inverse_db_table(156) = 0.0019632195: inverse_db_table(157) = 0.0020908006: inverse_db_table(158) = 0.0022266726: inverse_db_table(159) = 0.0023713743
    inverse_db_table(160) = 0.0025254795: inverse_db_table(161) = 0.0026895994: inverse_db_table(162) = 0.0028643847: inverse_db_table(163) = 0.0030505286
    inverse_db_table(164) = 0.0032487691: inverse_db_table(165) = 0.0034598925: inverse_db_table(166) = 0.0036847358: inverse_db_table(167) = 0.0039241906
    inverse_db_table(168) = 0.0041792066: inverse_db_table(169) = 0.0044507950: inverse_db_table(170) = 0.0047400328: inverse_db_table(171) = 0.0050480668
    inverse_db_table(172) = 0.0053761186: inverse_db_table(173) = 0.0057254891: inverse_db_table(174) = 0.0060975636: inverse_db_table(175) = 0.0064938176
    inverse_db_table(176) = 0.0069158225: inverse_db_table(177) = 0.0073652516: inverse_db_table(178) = 0.0078438871: inverse_db_table(179) = 0.0083536271
    inverse_db_table(180) = 0.0088964928: inverse_db_table(181) = 0.009474637: inverse_db_table(182) = 0.010090352: inverse_db_table(183) = 0.010746080
    inverse_db_table(184) = 0.011444421: inverse_db_table(185) = 0.012188144: inverse_db_table(186) = 0.012980198: inverse_db_table(187) = 0.013823725
    inverse_db_table(188) = 0.014722068: inverse_db_table(189) = 0.015678791: inverse_db_table(190) = 0.016697687: inverse_db_table(191) = 0.017782797
    inverse_db_table(192) = 0.018938423: inverse_db_table(193) = 0.020169149: inverse_db_table(194) = 0.021479854: inverse_db_table(195) = 0.022875735
    inverse_db_table(196) = 0.024362330: inverse_db_table(197) = 0.025945531: inverse_db_table(198) = 0.027631618: inverse_db_table(199) = 0.029427276
    inverse_db_table(200) = 0.031339626: inverse_db_table(201) = 0.033376252: inverse_db_table(202) = 0.035545228: inverse_db_table(203) = 0.037855157
    inverse_db_table(204) = 0.040315199: inverse_db_table(205) = 0.042935108: inverse_db_table(206) = 0.045725273: inverse_db_table(207) = 0.048696758
    inverse_db_table(208) = 0.051861348: inverse_db_table(209) = 0.055231591: inverse_db_table(210) = 0.058820850: inverse_db_table(211) = 0.062643361
    inverse_db_table(212) = 0.066714279: inverse_db_table(213) = 0.071049749: inverse_db_table(214) = 0.075666962: inverse_db_table(215) = 0.080584227
    inverse_db_table(216) = 0.085821044: inverse_db_table(217) = 0.091398179: inverse_db_table(218) = 0.097337747: inverse_db_table(219) = 0.10366330
    inverse_db_table(220) = 0.11039993: inverse_db_table(221) = 0.11757434: inverse_db_table(222) = 0.12521498: inverse_db_table(223) = 0.13335215
    inverse_db_table(224) = 0.14201813: inverse_db_table(225) = 0.15124727: inverse_db_table(226) = 0.16107617: inverse_db_table(227) = 0.17154380
    inverse_db_table(228) = 0.18269168: inverse_db_table(229) = 0.19456402: inverse_db_table(230) = 0.20720788: inverse_db_table(231) = 0.22067342
    inverse_db_table(232) = 0.23501402: inverse_db_table(233) = 0.25028656: inverse_db_table(234) = 0.26655159: inverse_db_table(235) = 0.28387361
    inverse_db_table(236) = 0.30232132: inverse_db_table(237) = 0.32196786: inverse_db_table(238) = 0.34289114: inverse_db_table(239) = 0.36517414
    inverse_db_table(240) = 0.38890521: inverse_db_table(241) = 0.41417847: inverse_db_table(242) = 0.44109412: inverse_db_table(243) = 0.46975890
    inverse_db_table(244) = 0.50028648: inverse_db_table(245) = 0.53279791: inverse_db_table(246) = 0.56742212: inverse_db_table(247) = 0.60429640
    inverse_db_table(248) = 0.64356699: inverse_db_table(249) = 0.68538959: inverse_db_table(250) = 0.72993007: inverse_db_table(251) = 0.77736504
    inverse_db_table(252) = 0.82788260: inverse_db_table(253) = 0.88168307: inverse_db_table(254) = 0.9389798: inverse_db_table(255) = 1.0
End Sub

' (VB) C's uint32 a + b (b a signed int added as C converts it), modulo 2^32
Private Function add32(ByVal a As Long, ByVal b As Long) As Long
    Dim d As Double
    d = CDbl(a) + CDbl(b)
    If d > 2147483647# Then
        d = d - 4294967296#
    ElseIf d < -2147483648# Then
        d = d + 4294967296#
    End If
    add32 = d
End Function

' (VB) C's uint32 a - b, modulo 2^32
Private Function sub32(ByVal a As Long, ByVal b As Long) As Long
    Dim d As Double
    d = CDbl(a) - CDbl(b)
    If d > 2147483647# Then
        d = d - 4294967296#
    ElseIf d < -2147483648# Then
        d = d + 4294967296#
    End If
    sub32 = d
End Function

' (VB) C's a < b for uint32
Private Function ult32(ByVal a As Long, ByVal b As Long) As Boolean
    ult32 = (a Xor &H80000000) < (b Xor &H80000000)
End Function

' (VB) C's x >> n for a uint32, 0 <= n <= 32
Private Function ushr(ByVal x As Long, ByVal n As Long) As Long
    If n = 0 Then
        ushr = x
    ElseIf n >= 32 Then
        ushr = 0
    ElseIf n = 31 Then
        If x < 0 Then ushr = 1 Else ushr = 0
    ElseIf x >= 0 Then
        ushr = x \ pow2tab(n)
    Else
        ushr = ((x And &H7FFFFFFF) \ pow2tab(n)) Or pow2tab(31 - n)
    End If
End Function

' (VB) C's x << n for a uint32 (the bits shifted out are lost), 0 <= n <= 31
Private Function ushl(ByVal x As Long, ByVal n As Long) As Long
    If n = 0 Then
        ushl = x
    Else
        x = x And masktab(32 - n)               ' the bits that survive
        If (x And pow2tab(31 - n)) <> 0 Then
            ushl = ((x And masktab(31 - n)) * pow2tab(n)) Or &H80000000
        Else
            ushl = x * pow2tab(n)
        End If
    End If
End Function

Private Function error_(f As vorb, ByVal e As Long) As Long
    f.error_ = e
    If f.eof_ = 0 And e <> VORBIS_need_more_data Then
        f.error_ = e ' breakpoint for debugging
    End If
    error_ = 0
End Function

' these functions are used for allocating temporary memory
' while decoding. if you can afford the stack space, use
' alloca(); otherwise, provide a temp buffer and it will
' allocate out of those.
' (VB) the memory is VB's arrays; setup_malloc's count of setup_memory_required is kept (get_info)
Private Sub setup_malloc(f As vorb, ByVal sz As Long)
    sz = (sz + 7) And Not 7 ' round up to nearest 8 for alignment of future allocs.
    f.setup_memory_required = f.setup_memory_required + sz
End Sub

Private Sub crc32_init()
    Dim i As Long, j As Long, s As Long
    For i = 0 To 255
        s = ushl(i, 24)
        For j = 0 To 7
            ' s = (s << 1) ^ (s >= (1U<<31) ? CRC32_POLY : 0);
            If s < 0 Then
                s = ushl(s, 1) Xor CRC32_POLY
            Else
                s = ushl(s, 1)
            End If
        Next
        crc_table(i) = s
    Next
End Sub

Private Function crc32_update(ByVal crc As Long, ByVal byte_ As Long) As Long
    ' return (crc << 8) ^ crc_table[byte ^ (crc >> 24)];
    crc32_update = ushl(crc, 8) Xor crc_table(byte_ Xor ushr(crc, 24))
End Function

' used in setup, and for huffman that doesn't go fast path
' (VB) the bits reversed a byte at a time (the C's masks and shifts give the same 32 bits reversed)
Private Function bit_reverse(ByVal n As Long) As Long
    Dim b0 As Long, b3 As Long, r As Long
    b3 = (n And &H7F000000) \ &H1000000
    If n < 0 Then b3 = b3 Or &H80
    r = rev8(b3) Or rev8((n And &HFF0000) \ &H10000) * &H100& Or rev8((n And &HFF00&) \ &H100&) * &H10000
    b0 = rev8(n And &HFF&)
    If b0 >= &H80 Then
        r = r Or ((b0 And &H7F) * &H1000000) Or &H80000000
    Else
        r = r Or b0 * &H1000000
    End If
    bit_reverse = r
End Function

Private Function square(ByVal x As Single) As Single
    square = x * x
End Function

' this is a weird definition of log2() for which log2(1) = 1, log2(2) = 2, log2(4) = 3
' as required by the specification. fast(?) implementation from stb.h
' @OPTIMIZE: called multiple times per-packet with "constants"; move to setup
Private Function ilog(ByVal n As Long) As Long
    If n < 0 Then ilog = 0: Exit Function ' signed n returns 0

    ' 2 compares if n < 16, 3 compares otherwise (4 if signed or n > 1<<29)
    If n < 16384 Then
        If n < 16 Then
            ilog = 0 + log2_4(n)
        ElseIf n < 512 Then
            ilog = 5 + log2_4(n \ 32)
        Else
            ilog = 10 + log2_4(n \ 1024)
        End If
    ElseIf n < 16777216 Then
        If n < 524288 Then
            ilog = 15 + log2_4(n \ 32768)
        Else
            ilog = 20 + log2_4(n \ 1048576)
        End If
    ElseIf n < 536870912 Then
        ilog = 25 + log2_4(n \ 33554432)
    Else
        ilog = 30 + log2_4(n \ 1073741824)
    End If
End Function

'/////////////////////// LEAF SETUP FUNCTIONS //////////////////////////
'
' these functions are only called at setup, and only a few times
' per file

Private Function float32_unpack(ByVal x As Long) As Single
    ' from the specification
    Dim mantissa As Long, exp_ As Long, res As Double
    mantissa = x And &H1FFFFF
    exp_ = (x And &H7FE00000) \ &H200000
    If x < 0 Then res = -CDbl(mantissa) Else res = mantissa   ' sign = x & 0x80000000
    float32_unpack = CSng(CSng(res) * 2# ^ (exp_ - 788))     ' (float) ldexp((float)res, (int)exp-788)
End Function

' zlib & jpeg huffman tables assume that the output symbols
' can either be arbitrarily arranged, or have monotonically
' increasing frequencies--they rely on the lengths being sorted;
' this makes for a very simple generation algorithm.
' vorbis allows a huffman table with non-sorted lengths. This
' requires a more sophisticated construction, since symbols in
' order do not map to huffman codes "in order".
Private Sub add_entry(c As CodebookT, ByVal huff_code As Long, ByVal symbol As Long, ByVal count As Long, ByVal len_ As Long, values() As Long)
    If c.sparse = 0 Then
        c.codewords(symbol) = huff_code
    Else
        c.codewords(count) = huff_code
        c.codeword_lengths(count) = len_
        values(count) = symbol
    End If
End Sub

Private Function compute_codewords(c As CodebookT, len_() As Byte, ByVal n As Long, values() As Long) As Long
    Dim i As Long, k As Long, m As Long, available(0 To 31) As Long, res As Long, z As Long, y As Long

    m = 0
    ' find the first entry
    For k = 0 To n - 1
        If len_(k) < NO_CODE Then Exit For
    Next
    If k = n Then compute_codewords = 1: Exit Function ' assert(c->sorted_entries == 0)
    ' assert(len[k] < 32); // no error return required, code reading lens checks this
    ' add to the list
    add_entry c, 0, k, m, len_(k), values
    m = m + 1
    ' add all available leaves
    For i = 1 To len_(k)
        available(i) = pow2tab(32 - i)          ' 1U << (32-i)
    Next
    ' note that the above code treats the first case specially,
    ' but it's really the same as the following code, so they
    ' could probably be combined (except the initial code is 0,
    ' and I use 0 in available[] to mean 'empty')
    For i = k + 1 To n - 1
        z = len_(i)
        If z <> NO_CODE Then
            ' assert(z < 32); // no error return required, code reading lens checks this
            ' find lowest available leaf (should always be earliest,
            ' which is what the specification calls for)
            ' note that this property, and the fact we can never have
            ' more than one free leaf at a given level, isn't totally
            ' trivial to prove, but it seems true and the assert never
            ' fires, so!
            Do While z > 0
                If available(z) <> 0 Then Exit Do
                z = z - 1
            Loop
            If z = 0 Then compute_codewords = 0: Exit Function
            res = available(z)
            available(z) = 0
            add_entry c, bit_reverse(res), i, m, len_(i), values
            m = m + 1
            ' propagate availability up the tree
            If z <> len_(i) Then
                For y = len_(i) To z + 1 Step -1
                    ' assert(available[y] == 0);
                    available(y) = add32(res, pow2tab(32 - y))  ' res + (1 << (32-y))
                Next
            End If
        End If
    Next
    compute_codewords = 1
End Function

' accelerated huffman table allows fast O(1) match of all symbols
' of length <= STB_VORBIS_FAST_HUFFMAN_LENGTH
Private Sub compute_accelerated_huffman(c As CodebookT)
    Dim i As Long, len_ As Long, z As Long
    For i = 0 To FAST_HUFFMAN_TABLE_SIZE - 1
        c.fast_huffman(i) = -1
    Next

    If c.sparse <> 0 Then len_ = c.sorted_entries Else len_ = c.entries
    If len_ > 32767 Then len_ = 32767 ' largest possible value we can encode!
    For i = 0 To len_ - 1
        If c.codeword_lengths(i) <= STB_VORBIS_FAST_HUFFMAN_LENGTH Then
            If c.sparse <> 0 Then z = bit_reverse(c.sorted_codewords(i)) Else z = c.codewords(i)
            ' set table entries for all bit combinations in the higher bits
            Do While z >= 0 And z < FAST_HUFFMAN_TABLE_SIZE   ' (uint32 z)
                c.fast_huffman(z) = i
                z = z + pow2tab(c.codeword_lengths(i))
            Loop
        End If
    Next
End Sub

' (VB) qsort(..., uint32_compare): the codewords are distinct, so any sort gives the C's order (a heap sort)
Private Sub sort_uint32(a() As Long, ByVal n As Long)
    Dim i As Long, t As Long
    For i = n \ 2 - 1 To 0 Step -1
        sift_uint32 a, i, n
    Next
    For i = n - 1 To 1 Step -1
        t = a(0): a(0) = a(i): a(i) = t
        sift_uint32 a, 0, i
    Next
End Sub

Private Sub sift_uint32(a() As Long, ByVal root As Long, ByVal n As Long)
    Dim child As Long, t As Long
    Do
        child = root * 2 + 1
        If child >= n Then Exit Do
        If child + 1 < n Then
            If ult32(a(child), a(child + 1)) Then child = child + 1
        End If
        If Not ult32(a(root), a(child)) Then Exit Do
        t = a(root): a(root) = a(child): a(child) = t
        root = child
    Loop
End Sub

Private Function include_in_sort(c As CodebookT, ByVal len_ As Long) As Long
    If c.sparse <> 0 Then include_in_sort = 1: Exit Function ' assert(len != NO_CODE)
    If len_ = NO_CODE Then include_in_sort = 0: Exit Function
    If len_ > STB_VORBIS_FAST_HUFFMAN_LENGTH Then include_in_sort = 1: Exit Function
    include_in_sort = 0
End Function

' if the fast table above doesn't work, we want to binary
' search them... need to reverse the bits
Private Sub compute_sorted_huffman(c As CodebookT, lengths() As Byte, values() As Long)
    Dim i As Long, len_ As Long, k As Long, huff_len As Long, code As Long, x As Long, n As Long, m As Long
    ' build a list of all the entries
    ' OPTIMIZATION: don't include the short ones, since they'll be caught by FAST_HUFFMAN.
    ' this is kind of a frivolous optimization--I don't see any performance improvement,
    ' but it's like 4 extra lines of code, so.
    If c.sparse = 0 Then
        k = 0
        For i = 0 To c.entries - 1
            If include_in_sort(c, lengths(i)) <> 0 Then
                c.sorted_codewords(k) = bit_reverse(c.codewords(i))
                k = k + 1
            End If
        Next
        ' assert(k == c->sorted_entries);
    Else
        For i = 0 To c.sorted_entries - 1
            c.sorted_codewords(i) = bit_reverse(c.codewords(i))
        Next
    End If

    sort_uint32 c.sorted_codewords, c.sorted_entries
    c.sorted_codewords(c.sorted_entries) = &HFFFFFFFF

    If c.sparse <> 0 Then len_ = c.sorted_entries Else len_ = c.entries
    ' now we need to indicate how they correspond; we could either
    '   #1: sort a different data structure that says who they correspond to
    '   #2: for each sorted entry, search the original list to find who corresponds
    '   #3: for each original entry, find the sorted entry
    ' #1 requires extra storage, #2 is slow, #3 can use binary search!
    For i = 0 To len_ - 1
        If c.sparse <> 0 Then huff_len = lengths(values(i)) Else huff_len = lengths(i)
        If include_in_sort(c, huff_len) <> 0 Then
            code = bit_reverse(c.codewords(i))
            x = 0: n = c.sorted_entries
            Do While n > 1
                ' invariant: sc[x] <= code < sc[x+n]
                m = x + n \ 2
                If Not ult32(code, c.sorted_codewords(m)) Then
                    x = m
                    n = n - n \ 2
                Else
                    n = n \ 2
                End If
            Loop
            ' assert(c->sorted_codewords[x] == code);
            If c.sparse <> 0 Then
                c.sorted_values(x) = values(i)
                c.codeword_lengths(x) = huff_len
            Else
                c.sorted_values(x) = i
            End If
        End If
    Next
End Sub

' only run while parsing the header (3 times)
Private Function vorbis_validate(data() As Byte) As Long
    If data(0) = &H76 And data(1) = &H6F And data(2) = &H72 And data(3) = &H62 And data(4) = &H69 And data(5) = &H73 Then vorbis_validate = 1 ' "vorbis"
End Function

' called from setup only, once per code book
' (formula implied by specification)
Private Function lookup1_values(ByVal entries As Long, ByVal dim_ As Long) As Long
    Dim r As Long
    If entries > 0 Then
        r = Int(Exp(CSng(CSng(Log(CSng(entries))) / dim_)))     ' (int) floor(exp((float) log((float) entries) / dim))
    Else
        r = 0   ' (VB) log(0) is -inf: exp(-inf) = 0 (VB's Log raises)
    End If
    If Int(CDbl(CSng(r) + 1!) ^ dim_) <= entries Then       ' (int) cast for MinGW warning;
        r = r + 1                                           ' floor() to avoid _ftol() when non-CRT
    End If
    If CDbl(CSng(r) + 1!) ^ dim_ <= entries Then
        lookup1_values = -1
        Exit Function
    End If
    If Int(CDbl(CSng(r)) ^ dim_) > entries Then
        lookup1_values = -1
        Exit Function
    End If
    lookup1_values = r
End Function

' called twice per file
Private Sub compute_twiddle_factors(ByVal n As Long, A() As Single, B() As Single, C() As Single)
    Dim n4 As Long, n8 As Long, k As Long, k2 As Long
    n4 = n \ 4: n8 = n \ 8

    k2 = 0
    For k = 0 To n4 - 1
        A(k2) = CSng(Cos(4 * k * M_PI / n))
        A(k2 + 1) = CSng(-Sin(4 * k * M_PI / n))
        B(k2) = CSng(CSng(Cos((k2 + 1) * M_PI / n / 2)) * 0.5!)
        B(k2 + 1) = CSng(CSng(Sin((k2 + 1) * M_PI / n / 2)) * 0.5!)
        k2 = k2 + 2
    Next
    k2 = 0
    For k = 0 To n8 - 1
        C(k2) = CSng(Cos(2 * (k2 + 1) * M_PI / n))
        C(k2 + 1) = CSng(-Sin(2 * (k2 + 1) * M_PI / n))
        k2 = k2 + 2
    Next
End Sub

Private Sub compute_window(ByVal n As Long, window_() As Single)
    Dim n2 As Long, i As Long
    n2 = n \ 2
    For i = 0 To n2 - 1
        window_(i) = CSng(Sin(0.5 * M_PI * square(CSng(Sin((i - 0 + 0.5) / n2 * 0.5 * M_PI)))))
    Next
End Sub

Private Sub compute_bitreverse(ByVal n As Long, rev() As Long)
    Dim ld As Long, i As Long, n8 As Long
    ld = ilog(n) - 1 ' ilog is off-by-one from normal definitions
    n8 = n \ 8
    For i = 0 To n8 - 1
        rev(i) = ushr(bit_reverse(i), 32 - ld + 3) * 4     ' (bit_reverse(i) >> (32-ld+3)) << 2
    Next
End Sub

Private Function init_blocksize(f As vorb, ByVal b As Long, ByVal n As Long) As Long
    Dim n2 As Long, n4 As Long, n8 As Long
    n2 = n \ 2: n4 = n \ 4: n8 = n \ 8
    ReDim f.A(b).v(0 To n2 - 1): setup_malloc f, 4 * n2
    ReDim f.B(b).v(0 To n2 - 1): setup_malloc f, 4 * n2
    ReDim f.C(b).v(0 To n4 - 1): setup_malloc f, 4 * n4
    compute_twiddle_factors n, f.A(b).v, f.B(b).v, f.C(b).v
    ReDim f.window_(b).v(0 To n2 - 1): setup_malloc f, 4 * n2
    compute_window n, f.window_(b).v
    ReDim f.bit_reverse(b).v(0 To n8 - 1): setup_malloc f, 2 * n8
    compute_bitreverse n, f.bit_reverse(b).v
    init_blocksize = 1
End Function

Private Sub neighbors(x() As Long, ByVal n As Long, plow As Long, phigh As Long)
    Dim low As Long, high As Long, i As Long
    low = -1
    high = 65536
    For i = 0 To n - 1
        If x(i) > low And x(i) < x(n) Then plow = i: low = x(i)
        If x(i) < high And x(i) > x(n) Then phigh = i: high = x(i)
    Next
End Sub

'
'/////////////////////// END LEAF SETUP FUNCTIONS //////////////////////////

' (STB_VORBIS_NO_STDIO: USE_MEMORY(z) is TRUE)
Private Function get8(z As vorb) As Long
    If z.stream_pos >= z.stream_len Then z.eof_ = 1: get8 = 0: Exit Function
    get8 = z.stream(z.stream_pos)
    z.stream_pos = z.stream_pos + 1
End Function

Private Function get32(f As vorb) As Long
    Dim x As Long
    x = get8(f)
    x = x + get8(f) * &H100&
    x = x + get8(f) * &H10000
    x = x Or ushl(get8(f), 24)              ' x += (uint32) get8(f) << 24 (the bits above 24 are clear)
    get32 = x
End Function

Private Function getn(z As vorb, data() As Byte, ByVal n As Long) As Long
    If z.stream_pos + n > z.stream_len Then z.eof_ = 1: getn = 0: Exit Function
    If n > 0 Then CopyMemory data(0), z.stream(z.stream_pos), n
    z.stream_pos = z.stream_pos + n
    getn = 1
End Function

Private Sub skip(z As vorb, ByVal n As Long)
    z.stream_pos = z.stream_pos + n
    If z.stream_pos >= z.stream_len Then z.eof_ = 1
End Sub

' loc: unsigned int
Private Function set_file_offset(f As vorb, ByVal loc As Long) As Long
    f.eof_ = 0
    ' if (f->stream_start + loc >= f->stream_end || f->stream_start + loc < f->stream_start)
    If loc < 0 Or loc >= f.stream_len Then
        f.stream_pos = f.stream_len
        f.eof_ = 1
        set_file_offset = 0
    Else
        f.stream_pos = loc
        set_file_offset = 1
    End If
End Function

Private Function capture_pattern(f As vorb) As Long
    If &H4F <> get8(f) Then Exit Function
    If &H67 <> get8(f) Then Exit Function
    If &H67 <> get8(f) Then Exit Function
    If &H53 <> get8(f) Then Exit Function
    capture_pattern = 1
End Function

Private Function start_page_no_capturepattern(f As vorb) As Long
    Dim loc0 As Long, loc1 As Long, n As Long, i As Long, len_ As Long
    If f.first_decode <> 0 Then             ' && !IS_PUSH_MODE(f)
        f.p_first.page_start = stb_vorbis_get_file_offset(f) - 4
    End If
    ' stream structure version
    If 0 <> get8(f) Then start_page_no_capturepattern = error_(f, VORBIS_invalid_stream_structure_version): Exit Function
    ' header flag
    f.page_flag = get8(f)
    ' absolute granule position
    loc0 = get32(f)
    loc1 = get32(f)
    ' @TODO: validate loc0,loc1 as valid positions?
    ' stream serial number -- vorbis doesn't interleave, so discard
    get32 f
    'if (f->serial != get32(f)) return error(f, VORBIS_incorrect_stream_serial_number);
    ' page sequence number
    n = get32(f)
    f.last_page = n
    ' CRC32
    get32 f
    ' page_segments
    f.segment_count = get8(f)
    If getn(f, f.segments, f.segment_count) = 0 Then
        start_page_no_capturepattern = error_(f, VORBIS_unexpected_eof)
        Exit Function
    End If
    ' assume we _don't_ know any the sample position of any segments
    f.end_seg_with_known_loc = -2
    If loc0 <> -1 Or loc1 <> -1 Then
        ' determine which packet is the last one that will complete
        For i = f.segment_count - 1 To 0 Step -1
            If f.segments(i) < 255 Then Exit For
        Next
        ' 'i' is now the index of the _last_ segment of a packet that ends
        If i >= 0 Then
            f.end_seg_with_known_loc = i
            f.known_loc_for_packet = loc0
        End If
    End If
    If f.first_decode <> 0 Then
        len_ = 0
        For i = 0 To f.segment_count - 1
            len_ = len_ + f.segments(i)
        Next
        len_ = len_ + 27 + f.segment_count
        f.p_first.page_end = f.p_first.page_start + len_
        f.p_first.last_decoded_sample = loc0
    End If
    f.next_seg = 0
    start_page_no_capturepattern = 1
End Function

Private Function start_page(f As vorb) As Long
    If capture_pattern(f) = 0 Then start_page = error_(f, VORBIS_missing_capture_pattern): Exit Function
    start_page = start_page_no_capturepattern(f)
End Function

Private Function start_packet(f As vorb) As Long
    Do While f.next_seg = -1
        If start_page(f) = 0 Then Exit Function
        If (f.page_flag And PAGEFLAG_continued_packet) <> 0 Then
            start_packet = error_(f, VORBIS_continued_packet_flag_invalid)
            Exit Function
        End If
    Loop
    f.last_seg = 0
    f.valid_bits = 0
    f.packet_bytes = 0
    f.bytes_in_seg = 0
    ' f->next_seg is now valid
    start_packet = 1
End Function

Private Function maybe_start_packet(f As vorb) As Long
    Dim x As Long
    If f.next_seg = -1 Then
        x = get8(f)
        If f.eof_ <> 0 Then Exit Function ' EOF at page boundary is not an error!
        If &H4F <> x Then maybe_start_packet = error_(f, VORBIS_missing_capture_pattern): Exit Function
        If &H67 <> get8(f) Then maybe_start_packet = error_(f, VORBIS_missing_capture_pattern): Exit Function
        If &H67 <> get8(f) Then maybe_start_packet = error_(f, VORBIS_missing_capture_pattern): Exit Function
        If &H53 <> get8(f) Then maybe_start_packet = error_(f, VORBIS_missing_capture_pattern): Exit Function
        If start_page_no_capturepattern(f) = 0 Then Exit Function
        If (f.page_flag And PAGEFLAG_continued_packet) <> 0 Then
            ' set up enough state that we can read this packet if we want,
            ' e.g. during recovery
            f.last_seg = 0
            f.bytes_in_seg = 0
            maybe_start_packet = error_(f, VORBIS_continued_packet_flag_invalid)
            Exit Function
        End If
    End If
    maybe_start_packet = start_packet(f)
End Function

Private Function next_segment(f As vorb) As Long
    Dim len_ As Long
    If f.last_seg <> 0 Then next_segment = 0: Exit Function
    If f.next_seg = -1 Then
        f.last_seg_which = f.segment_count - 1 ' in case start_page fails
        If start_page(f) = 0 Then f.last_seg = 1: next_segment = 0: Exit Function
        If (f.page_flag And PAGEFLAG_continued_packet) = 0 Then
            next_segment = error_(f, VORBIS_continued_packet_flag_invalid)
            Exit Function
        End If
    End If
    len_ = f.segments(f.next_seg)
    f.next_seg = f.next_seg + 1
    If len_ < 255 Then
        f.last_seg = 1
        f.last_seg_which = f.next_seg - 1
    End If
    If f.next_seg >= f.segment_count Then
        f.next_seg = -1
    End If
    ' assert(f->bytes_in_seg == 0);
    f.bytes_in_seg = len_
    next_segment = len_
End Function

Private Function get8_packet_raw(f As vorb) As Long
    If f.bytes_in_seg = 0 Then  ' CLANG!
        If f.last_seg <> 0 Then
            get8_packet_raw = EOP
            Exit Function
        ElseIf next_segment(f) = 0 Then
            get8_packet_raw = EOP
            Exit Function
        End If
    End If
    ' assert(f->bytes_in_seg > 0);
    f.bytes_in_seg = f.bytes_in_seg - 1
    f.packet_bytes = f.packet_bytes + 1
    ' get8(f), written out
    If f.stream_pos >= f.stream_len Then f.eof_ = 1: get8_packet_raw = 0: Exit Function
    get8_packet_raw = f.stream(f.stream_pos)
    f.stream_pos = f.stream_pos + 1
End Function

Private Function get8_packet(f As vorb) As Long
    Dim x As Long
    x = get8_packet_raw(f)
    f.valid_bits = 0
    get8_packet = x
End Function

Private Function get32_packet(f As vorb) As Long
    Dim x As Long
    ' uint32 x: the bytes added as C converts each int (EOP is -1) to uint32
    x = get8_packet(f)
    x = add32(x, get8_packet(f) * &H100&)
    x = add32(x, get8_packet(f) * &H10000)
    x = add32(x, ushl(get8_packet(f) And &HFF&, 24))   ' (uint32) get8_packet(f) << 24
    get32_packet = x
End Function

Private Sub flush_packet(f As vorb)
    Do While get8_packet_raw(f) <> EOP
    Loop
End Sub

' @OPTIMIZE: this is the secondary bit decoder, so it's probably not as important
' as the huffman decoder?
Private Function get_bits(f As vorb, ByVal n As Long) As Long
    Dim z As Long, b As Long

    If f.valid_bits < 0 Then Exit Function  ' 0
    If f.valid_bits < n Then
        If n > 24 Then
            ' the accumulator technique below would not work correctly in this case
            z = get_bits(f, 24)
            z = z Or ushl(get_bits(f, n - 24), 24)  ' z += get_bits(f, n-24) << 24
            get_bits = z
            Exit Function
        End If
        If f.valid_bits = 0 Then f.acc = 0
        Do While f.valid_bits < n
            b = get8_packet_raw(f)
            If b = EOP Then
                f.valid_bits = INVALID_BITS
                Exit Function   ' 0
            End If
            f.acc = f.acc + b * pow2tab(f.valid_bits)   ' valid_bits < n <= 24: below bit 31
            f.valid_bits = f.valid_bits + 8
        Loop
    End If

    ' assert(f->valid_bits >= n);
    z = f.acc And masktab(n)
    f.acc = ushr(f.acc, n)
    f.valid_bits = f.valid_bits - n
    get_bits = z
End Function

' @OPTIMIZE: primary accumulator for huffman
' expand the buffer to as many bits as possible without reading off end of packet
' it might be nice to allow f->valid_bits and f->acc to be stored in registers,
' e.g. cache them locally and decode locally
Private Sub prep_huffman(f As vorb)
    Dim z As Long, s As Long
    If f.valid_bits <= 24 Then
        If f.valid_bits = 0 Then f.acc = 0
        Do
            If f.last_seg <> 0 And f.bytes_in_seg = 0 Then Exit Sub
            z = get8_packet_raw(f)
            If z = EOP Then Exit Sub
            ' f->acc += (unsigned) z << f->valid_bits; (the bits above valid_bits are clear)
            s = f.valid_bits And 31
            If s <= 23 Then
                f.acc = f.acc Or z * pow2tab(s)
            Else
                f.acc = f.acc Or ushl(z, s)
            End If
            f.valid_bits = f.valid_bits + 8
        Loop While f.valid_bits <= 24
    End If
End Sub

Private Function codebook_decode_scalar_raw(f As vorb, c As CodebookT) As Long
    Dim i As Long, code As Long, x As Long, n As Long, len_ As Long, m As Long, useBinary As Boolean
    prep_huffman f

    If Not c.has_codewords And Not c.has_sorted Then codebook_decode_scalar_raw = -1: Exit Function

    ' cases to use binary search: sorted_codewords && !c->codewords
    '                             sorted_codewords && c->entries > 8
    If c.entries > 8 Then useBinary = c.has_sorted Else useBinary = Not c.has_codewords
    If useBinary Then
        ' binary search
        code = bit_reverse(f.acc)
        x = 0: n = c.sorted_entries

        Do While n > 1
            ' invariant: sc[x] <= code < sc[x+n]
            m = x + n \ 2
            If Not ult32(code, c.sorted_codewords(m)) Then
                x = m
                n = n - n \ 2
            Else
                n = n \ 2
            End If
        Loop
        ' x is now the sorted index
        If c.sparse = 0 Then x = c.sorted_values(x)
        ' x is now sorted index if sparse, or symbol otherwise
        len_ = c.codeword_lengths(x)
        If f.valid_bits >= len_ Then
            f.acc = ushr(f.acc, len_)
            f.valid_bits = f.valid_bits - len_
            codebook_decode_scalar_raw = x
            Exit Function
        End If

        f.valid_bits = 0
        codebook_decode_scalar_raw = -1
        Exit Function
    End If

    ' if small, linear search
    ' assert(!c->sparse);
    For i = 0 To c.entries - 1
        If c.codeword_lengths(i) <> NO_CODE Then
            If c.codewords(i) = (f.acc And masktab(c.codeword_lengths(i))) Then
                If f.valid_bits >= c.codeword_lengths(i) Then
                    f.acc = ushr(f.acc, c.codeword_lengths(i))
                    f.valid_bits = f.valid_bits - c.codeword_lengths(i)
                    codebook_decode_scalar_raw = i
                    Exit Function
                End If
                f.valid_bits = 0
                codebook_decode_scalar_raw = -1
                Exit Function
            End If
        End If
    Next

    error_ f, VORBIS_invalid_stream
    f.valid_bits = 0
    codebook_decode_scalar_raw = -1
End Function

' (VB) DECODE_RAW(var, f, c) (!STB_VORBIS_NO_INLINE_DECODE): the fast huffman table, the accumulator's
' shift written out
Private Function DECODE_RAW(f As vorb, c As CodebookT) As Long
    Dim v As Long, n As Long, a As Long
    If f.valid_bits < STB_VORBIS_FAST_HUFFMAN_LENGTH Then prep_huffman f
    v = f.acc And FAST_HUFFMAN_TABLE_MASK
    v = c.fast_huffman(v)
    If v >= 0 Then
        n = c.codeword_lengths(v)
        a = f.acc
        If a >= 0 Then
            f.acc = a \ pow2tab(n)
        Else
            f.acc = ((a And &H7FFFFFFF) \ pow2tab(n)) Or pow2tab(31 - n)
        End If
        f.valid_bits = f.valid_bits - n
        If f.valid_bits < 0 Then f.valid_bits = 0: v = -1
    Else
        v = codebook_decode_scalar_raw(f, c)
    End If
    DECODE_RAW = v
End Function

' DECODE(var,f,c): DECODE_RAW, then if (c->sparse) var = c->sorted_values[var];
Private Function DECODE(f As vorb, c As CodebookT) As Long
    Dim v As Long
    v = DECODE_RAW(f, c)
    If c.sparse <> 0 Then v = c.sorted_values(v)
    DECODE = v
End Function

' (!STB_VORBIS_DIVIDES_IN_CODEBOOK: DECODE_VQ is DECODE_RAW)

' CODEBOOK_ELEMENT_FAST is an optimization for the CODEBOOK_FLOATS case
' where we avoid one addition
' (CODEBOOK_ELEMENT(c,off) and CODEBOOK_ELEMENT_FAST(c,off) are c->multiplicands[off], CODEBOOK_ELEMENT_BASE(c) 0)

Private Function codebook_decode_start(f As vorb, c As CodebookT) As Long
    Dim z As Long
    z = -1

    ' type 0 is only legal in a scalar context
    If c.lookup_type = 0 Then
        error_ f, VORBIS_invalid_stream
    Else
        z = DECODE_RAW(f, c)    ' DECODE_VQ
        ' if (c->sparse) assert(z < c->sorted_entries);
        If z < 0 Then  ' check for EOP
            If f.bytes_in_seg = 0 Then
                If f.last_seg <> 0 Then
                    codebook_decode_start = z
                    Exit Function
                End If
            End If
            error_ f, VORBIS_invalid_stream
        End If
    End If
    codebook_decode_start = z
End Function

' output: a float * (output() from index o)
Private Function codebook_decode(f As vorb, c As CodebookT, output_() As Single, ByVal o As Long, ByVal len_ As Long) As Long
    Dim i As Long, z As Long, last As Single, val_ As Single
    z = codebook_decode_start(f, c)
    If z < 0 Then codebook_decode = 0: Exit Function
    If len_ > c.dimensions Then len_ = c.dimensions

    z = z * c.dimensions
    If c.sequence_p <> 0 Then
        last = 0!   ' CODEBOOK_ELEMENT_BASE(c)
        For i = 0 To len_ - 1
            val_ = c.multiplicands(z + i) + last
            output_(o + i) = output_(o + i) + val_
            last = val_ + c.minimum_value
        Next
    Else
        last = 0!   ' CODEBOOK_ELEMENT_BASE(c)
        For i = 0 To len_ - 1
            output_(o + i) = output_(o + i) + (c.multiplicands(z + i) + last)
        Next
    End If

    codebook_decode = 1
End Function

Private Function codebook_decode_step(f As vorb, c As CodebookT, output_() As Single, ByVal o As Long, ByVal len_ As Long, ByVal step_ As Long) As Long
    Dim i As Long, z As Long, last As Single, val_ As Single
    z = codebook_decode_start(f, c)
    last = 0!   ' CODEBOOK_ELEMENT_BASE(c)
    If z < 0 Then codebook_decode_step = 0: Exit Function
    If len_ > c.dimensions Then len_ = c.dimensions

    z = z * c.dimensions
    For i = 0 To len_ - 1
        val_ = c.multiplicands(z + i) + last
        output_(o + i * step_) = output_(o + i * step_) + val_
        If c.sequence_p <> 0 Then last = val_
    Next

    codebook_decode_step = 1
End Function

' outputs: residue_buffers' channel numbers (-1 = NULL)
Private Function codebook_decode_deinterleave_repeat(f As vorb, c As CodebookT, outputs() As Long, ByVal ch As Long, c_inter_p As Long, p_inter_p As Long, ByVal len_ As Long, ByVal total_decode As Long) As Long
    Dim c_inter As Long, p_inter As Long, i As Long, z As Long, effective As Long, last As Single, val_ As Single, o As Long
    c_inter = c_inter_p
    p_inter = p_inter_p
    effective = c.dimensions

    ' type 0 is only legal in a scalar context
    If c.lookup_type = 0 Then codebook_decode_deinterleave_repeat = error_(f, VORBIS_invalid_stream): Exit Function

    Do While total_decode > 0
        last = 0!   ' CODEBOOK_ELEMENT_BASE(c)
        z = DECODE_RAW(f, c)    ' DECODE_VQ
        ' assert(!c->sparse || z < c->sorted_entries);
        If z < 0 Then
            If f.bytes_in_seg = 0 Then
                If f.last_seg <> 0 Then codebook_decode_deinterleave_repeat = 0: Exit Function
            End If
            codebook_decode_deinterleave_repeat = error_(f, VORBIS_invalid_stream)
            Exit Function
        End If

        ' if this will take us off the end of the buffers, stop short!
        ' we check by computing the length of the virtual interleaved
        ' buffer (len*ch), our current offset within it (p_inter*ch)+(c_inter),
        ' and the length we'll be using (effective)
        If c_inter + p_inter * ch + effective > len_ * ch Then
            effective = len_ * ch - (p_inter * ch - c_inter)
        End If

        z = z * c.dimensions
        If c.sequence_p <> 0 Then
            For i = 0 To effective - 1
                val_ = c.multiplicands(z + i) + last
                o = outputs(c_inter)
                If o >= 0 Then f.channel_buffers(o).v(p_inter) = f.channel_buffers(o).v(p_inter) + val_
                c_inter = c_inter + 1
                If c_inter = ch Then c_inter = 0: p_inter = p_inter + 1
                last = val_
            Next
        Else
            For i = 0 To effective - 1
                val_ = c.multiplicands(z + i) + last
                o = outputs(c_inter)
                If o >= 0 Then f.channel_buffers(o).v(p_inter) = f.channel_buffers(o).v(p_inter) + val_
                c_inter = c_inter + 1
                If c_inter = ch Then c_inter = 0: p_inter = p_inter + 1
            Next
        End If

        total_decode = total_decode - effective
    Loop
    c_inter_p = c_inter
    p_inter_p = p_inter
    codebook_decode_deinterleave_repeat = 1
End Function

Private Function predict_point(ByVal x As Long, ByVal x0 As Long, ByVal x1 As Long, ByVal y0 As Long, ByVal y1 As Long) As Long
    Dim dy As Long, adx As Long, err As Long, off As Long
    dy = y1 - y0
    adx = x1 - x0
    ' @OPTIMIZE: force int division to round in the right direction... is this necessary on x86?
    err = Abs(dy) * (x - x0)
    off = err \ adx
    If dy < 0 Then predict_point = y0 - off Else predict_point = y0 + off
End Function

' @OPTIMIZE: if you want to replace this bresenham line-drawing routine,
' note that you must produce bit-identical output to decode correctly;
' this specific sequence of operations is specified in the spec (it's
' drawing integer-quantized frequency-space lines that the encoder
' expects to be exactly the same)
'     ... also, isn't the whole point of Bresenham's algorithm to NOT
' have to divide in the setup? sigh.
' (!STB_VORBIS_NO_DEFER_FLOOR: LINE_OP(a,b) is a *= b; no STB_VORBIS_DIVIDE_TABLE)
Private Sub draw_line(output_() As Single, ByVal x0 As Long, ByVal y0 As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal n As Long)
    Dim dy As Long, adx As Long, ady As Long, base_ As Long, x As Long, y As Long, err As Long, sy As Long
    dy = y1 - y0
    adx = x1 - x0
    ady = Abs(dy)
    x = x0: y = y0
    err = 0

    base_ = dy \ adx
    If dy < 0 Then
        sy = base_ - 1
    Else
        sy = base_ + 1
    End If
    ady = ady - Abs(base_) * adx
    If x1 > n Then x1 = n
    If x < x1 Then
        output_(x) = output_(x) * inverse_db_table(y And 255)
        x = x + 1
        Do While x < x1
            err = err + ady
            If err >= adx Then
                err = err - adx
                y = y + sy
            Else
                y = y + base_
            End If
            output_(x) = output_(x) * inverse_db_table(y And 255)
            x = x + 1
        Loop
    End If
End Sub

' target: channel_buffers' channel number
Private Function residue_decode(f As vorb, book As CodebookT, ByVal target As Long, ByVal offset As Long, ByVal n As Long, ByVal rtype As Long) As Long
    Dim k As Long, step_ As Long
    If rtype = 0 Then
        step_ = n \ book.dimensions
        For k = 0 To step_ - 1
            If codebook_decode_step(f, book, f.channel_buffers(target).v, offset + k, n - offset - k, step_) = 0 Then
                residue_decode = 0
                Exit Function
            End If
        Next
    Else
        k = 0
        Do While k < n
            If codebook_decode(f, book, f.channel_buffers(target).v, offset, n - k) = 0 Then
                residue_decode = 0
                Exit Function
            End If
            k = k + book.dimensions
            offset = offset + book.dimensions
        Loop
    End If
    residue_decode = 1
End Function

' n is 1/2 of the blocksize --
' specification: "Correct per-vector decode length is [n]/2"
' residue_buffers: the channels' numbers (-1 = NULL); r: f->residue_config + rn (!STB_VORBIS_DIVIDES_IN_RESIDUE)
Private Sub decode_residue(f As vorb, r As ResidueT, residue_buffers() As Long, ByVal ch As Long, ByVal n As Long, ByVal rn As Long, do_not_decode() As Long)
    Dim i As Long, j As Long, pass As Long, rtype As Long, c As Long, classwords As Long
    Dim actual_size As Long, limit_r_begin As Long, limit_r_end As Long, n_read As Long, part_read As Long
    Dim pcount As Long, class_set As Long, z As Long, c_inter As Long, p_inter As Long, q As Long, cc As Long, b As Long
    Dim temp As Long, offset As Long, nn As Long

    rtype = f.residue_types(rn)
    c = r.classbook
    classwords = f.codebooks(c).dimensions
    If rtype = 2 Then actual_size = n * 2 Else actual_size = n
    If r.begin_ < actual_size Then limit_r_begin = r.begin_ Else limit_r_begin = actual_size
    If r.end_ < actual_size Then limit_r_end = r.end_ Else limit_r_end = actual_size
    n_read = limit_r_end - limit_r_begin
    part_read = n_read \ r.part_size
    ' uint8 ***part_classdata = (uint8 ***) temp_block_array(f,f->channels, part_read * sizeof(**part_classdata));
    If part_classdataSize < f.channels * part_read + 1 Then
        part_classdataSize = f.channels * part_read + 1
        ReDim part_classdata(0 To part_classdataSize - 1)
    End If

    For i = 0 To ch - 1
        If do_not_decode(i) = 0 Then
            ZeroMemory f.channel_buffers(residue_buffers(i)).v(0), 4 * n     ' memset(residue_buffers[i], 0, sizeof(float) * n)
        End If
    Next

    If rtype = 2 And ch <> 1 Then
        For j = 0 To ch - 1
            If do_not_decode(j) = 0 Then Exit For
        Next
        If j = ch Then GoTo done

        For pass = 0 To 7
            pcount = 0: class_set = 0
            If ch = 2 Then
                Do While pcount < part_read
                    z = r.begin_ + pcount * r.part_size
                    c_inter = z And 1: p_inter = z \ 2
                    If pass = 0 Then
                        q = DECODE(f, f.codebooks(r.classbook))
                        If q = EOP Then GoTo done
                        part_classdata(class_set) = q   ' part_classdata[0][class_set] = r->classdata[q];
                    End If
                    i = 0
                    Do While i < classwords And pcount < part_read
                        z = r.begin_ + pcount * r.part_size
                        cc = r.classdata(part_classdata(class_set) * classwords + i)
                        b = r.residue_books(cc, pass)
                        If b >= 0 Then
                            ' saves 1%
                            If codebook_decode_deinterleave_repeat(f, f.codebooks(b), residue_buffers, ch, c_inter, p_inter, n, r.part_size) = 0 Then GoTo done
                        Else
                            z = z + r.part_size
                            c_inter = z And 1
                            p_inter = z \ 2
                        End If
                        i = i + 1: pcount = pcount + 1
                    Loop
                    class_set = class_set + 1
                Loop
            ElseIf ch > 2 Then
                Do While pcount < part_read
                    z = r.begin_ + pcount * r.part_size
                    c_inter = z Mod ch: p_inter = z \ ch
                    If pass = 0 Then
                        q = DECODE(f, f.codebooks(r.classbook))
                        If q = EOP Then GoTo done
                        part_classdata(class_set) = q   ' part_classdata[0][class_set] = r->classdata[q];
                    End If
                    i = 0
                    Do While i < classwords And pcount < part_read
                        z = r.begin_ + pcount * r.part_size
                        cc = r.classdata(part_classdata(class_set) * classwords + i)
                        b = r.residue_books(cc, pass)
                        If b >= 0 Then
                            If codebook_decode_deinterleave_repeat(f, f.codebooks(b), residue_buffers, ch, c_inter, p_inter, n, r.part_size) = 0 Then GoTo done
                        Else
                            z = z + r.part_size
                            c_inter = z Mod ch
                            p_inter = z \ ch
                        End If
                        i = i + 1: pcount = pcount + 1
                    Loop
                    class_set = class_set + 1
                Loop
            End If
        Next
        GoTo done
    End If

    For pass = 0 To 7
        pcount = 0: class_set = 0
        Do While pcount < part_read
            If pass = 0 Then
                For j = 0 To ch - 1
                    If do_not_decode(j) = 0 Then
                        temp = DECODE(f, f.codebooks(r.classbook))
                        If temp = EOP Then GoTo done
                        part_classdata(j * part_read + class_set) = temp    ' part_classdata[j][class_set] = r->classdata[temp];
                    End If
                Next
            End If
            i = 0
            Do While i < classwords And pcount < part_read
                For j = 0 To ch - 1
                    If do_not_decode(j) = 0 Then
                        cc = r.classdata(part_classdata(j * part_read + class_set) * classwords + i)
                        b = r.residue_books(cc, pass)
                        If b >= 0 Then
                            offset = r.begin_ + pcount * r.part_size
                            nn = r.part_size
                            If residue_decode(f, f.codebooks(b), residue_buffers(j), offset, nn, rtype) = 0 Then GoTo done
                        End If
                    End If
                Next
                i = i + 1: pcount = pcount + 1
            Loop
            class_set = class_set + 1
        Loop
    Next
done:
End Sub

' the following were split out into separate functions while optimizing;
' they could be pushed back up but eh. __forceinline showed no change;
' they're probably already being inlined.
' (VB) e is a float array and the C's pointers are indexes into it (ee0 = e + i_off ...), A from index 0
Private Sub imdct_step3_iter0_loop(ByVal n As Long, e() As Single, ByVal i_off As Long, ByVal k_off As Long, A() As Single)
    Dim ee0 As Long, ee2 As Long, i As Long, ai As Long, k00_20 As Single, k01_21 As Single
    ee0 = i_off
    ee2 = ee0 + k_off

    ' assert((n & 3) == 0);
    For i = n \ 4 To 1 Step -1
        k00_20 = e(ee0) - e(ee2)
        k01_21 = e(ee0 - 1) - e(ee2 - 1)
        e(ee0) = e(ee0) + e(ee2)
        e(ee0 - 1) = e(ee0 - 1) + e(ee2 - 1)
        e(ee2) = k00_20 * A(ai) - k01_21 * A(ai + 1)
        e(ee2 - 1) = k01_21 * A(ai) + k00_20 * A(ai + 1)
        ai = ai + 8

        k00_20 = e(ee0 - 2) - e(ee2 - 2)
        k01_21 = e(ee0 - 3) - e(ee2 - 3)
        e(ee0 - 2) = e(ee0 - 2) + e(ee2 - 2)
        e(ee0 - 3) = e(ee0 - 3) + e(ee2 - 3)
        e(ee2 - 2) = k00_20 * A(ai) - k01_21 * A(ai + 1)
        e(ee2 - 3) = k01_21 * A(ai) + k00_20 * A(ai + 1)
        ai = ai + 8

        k00_20 = e(ee0 - 4) - e(ee2 - 4)
        k01_21 = e(ee0 - 5) - e(ee2 - 5)
        e(ee0 - 4) = e(ee0 - 4) + e(ee2 - 4)
        e(ee0 - 5) = e(ee0 - 5) + e(ee2 - 5)
        e(ee2 - 4) = k00_20 * A(ai) - k01_21 * A(ai + 1)
        e(ee2 - 5) = k01_21 * A(ai) + k00_20 * A(ai + 1)
        ai = ai + 8

        k00_20 = e(ee0 - 6) - e(ee2 - 6)
        k01_21 = e(ee0 - 7) - e(ee2 - 7)
        e(ee0 - 6) = e(ee0 - 6) + e(ee2 - 6)
        e(ee0 - 7) = e(ee0 - 7) + e(ee2 - 7)
        e(ee2 - 6) = k00_20 * A(ai) - k01_21 * A(ai + 1)
        e(ee2 - 7) = k01_21 * A(ai) + k00_20 * A(ai + 1)
        ai = ai + 8
        ee0 = ee0 - 8
        ee2 = ee2 - 8
    Next
End Sub

Private Sub imdct_step3_inner_r_loop(ByVal lim As Long, e() As Single, ByVal d0 As Long, ByVal k_off As Long, A() As Single, ByVal k1 As Long)
    Dim i As Long, k00_20 As Single, k01_21 As Single, e0 As Long, e2 As Long, ai As Long

    e0 = d0
    e2 = e0 + k_off

    For i = lim \ 4 To 1 Step -1
        k00_20 = e(e0) - e(e2)
        k01_21 = e(e0 - 1) - e(e2 - 1)
        e(e0) = e(e0) + e(e2)
        e(e0 - 1) = e(e0 - 1) + e(e2 - 1)
        e(e2) = (k00_20) * A(ai) - (k01_21) * A(ai + 1)
        e(e2 - 1) = (k01_21) * A(ai) + (k00_20) * A(ai + 1)

        ai = ai + k1

        k00_20 = e(e0 - 2) - e(e2 - 2)
        k01_21 = e(e0 - 3) - e(e2 - 3)
        e(e0 - 2) = e(e0 - 2) + e(e2 - 2)
        e(e0 - 3) = e(e0 - 3) + e(e2 - 3)
        e(e2 - 2) = (k00_20) * A(ai) - (k01_21) * A(ai + 1)
        e(e2 - 3) = (k01_21) * A(ai) + (k00_20) * A(ai + 1)

        ai = ai + k1

        k00_20 = e(e0 - 4) - e(e2 - 4)
        k01_21 = e(e0 - 5) - e(e2 - 5)
        e(e0 - 4) = e(e0 - 4) + e(e2 - 4)
        e(e0 - 5) = e(e0 - 5) + e(e2 - 5)
        e(e2 - 4) = (k00_20) * A(ai) - (k01_21) * A(ai + 1)
        e(e2 - 5) = (k01_21) * A(ai) + (k00_20) * A(ai + 1)

        ai = ai + k1

        k00_20 = e(e0 - 6) - e(e2 - 6)
        k01_21 = e(e0 - 7) - e(e2 - 7)
        e(e0 - 6) = e(e0 - 6) + e(e2 - 6)
        e(e0 - 7) = e(e0 - 7) + e(e2 - 7)
        e(e2 - 6) = (k00_20) * A(ai) - (k01_21) * A(ai + 1)
        e(e2 - 7) = (k01_21) * A(ai) + (k00_20) * A(ai + 1)

        e0 = e0 - 8
        e2 = e2 - 8

        ai = ai + k1
    Next
End Sub

' A: the twiddle factors from index a0 (the C's A pointer)
Private Sub imdct_step3_inner_s_loop(ByVal n As Long, e() As Single, ByVal i_off As Long, ByVal k_off As Long, A() As Single, ByVal a0 As Long, ByVal a_off As Long, ByVal k0 As Long)
    Dim i As Long, A0_ As Single, A1 As Single, A2 As Single, A3 As Single, A4 As Single, A5 As Single, A6 As Single, A7 As Single
    Dim k00 As Single, k11 As Single, ee0 As Long, ee2 As Long
    A0_ = A(a0 + 0)
    A1 = A(a0 + 0 + 1)
    A2 = A(a0 + 0 + a_off)
    A3 = A(a0 + 0 + a_off + 1)
    A4 = A(a0 + 0 + a_off * 2 + 0)
    A5 = A(a0 + 0 + a_off * 2 + 1)
    A6 = A(a0 + 0 + a_off * 3 + 0)
    A7 = A(a0 + 0 + a_off * 3 + 1)

    ee0 = i_off
    ee2 = ee0 + k_off

    For i = n To 1 Step -1
        k00 = e(ee0) - e(ee2)
        k11 = e(ee0 - 1) - e(ee2 - 1)
        e(ee0) = e(ee0) + e(ee2)
        e(ee0 - 1) = e(ee0 - 1) + e(ee2 - 1)
        e(ee2) = (k00) * A0_ - (k11) * A1
        e(ee2 - 1) = (k11) * A0_ + (k00) * A1

        k00 = e(ee0 - 2) - e(ee2 - 2)
        k11 = e(ee0 - 3) - e(ee2 - 3)
        e(ee0 - 2) = e(ee0 - 2) + e(ee2 - 2)
        e(ee0 - 3) = e(ee0 - 3) + e(ee2 - 3)
        e(ee2 - 2) = (k00) * A2 - (k11) * A3
        e(ee2 - 3) = (k11) * A2 + (k00) * A3

        k00 = e(ee0 - 4) - e(ee2 - 4)
        k11 = e(ee0 - 5) - e(ee2 - 5)
        e(ee0 - 4) = e(ee0 - 4) + e(ee2 - 4)
        e(ee0 - 5) = e(ee0 - 5) + e(ee2 - 5)
        e(ee2 - 4) = (k00) * A4 - (k11) * A5
        e(ee2 - 5) = (k11) * A4 + (k00) * A5

        k00 = e(ee0 - 6) - e(ee2 - 6)
        k11 = e(ee0 - 7) - e(ee2 - 7)
        e(ee0 - 6) = e(ee0 - 6) + e(ee2 - 6)
        e(ee0 - 7) = e(ee0 - 7) + e(ee2 - 7)
        e(ee2 - 6) = (k00) * A6 - (k11) * A7
        e(ee2 - 7) = (k11) * A6 + (k00) * A7

        ee0 = ee0 - k0
        ee2 = ee2 - k0
    Next
End Sub

' z: an index into e (the C's float *z)
Private Sub iter_54(e() As Single, ByVal z As Long)
    Dim k00 As Single, k11 As Single, k22 As Single, k33 As Single
    Dim y0 As Single, y1 As Single, y2 As Single, y3 As Single

    k00 = e(z) - e(z - 4)
    y0 = e(z) + e(z - 4)
    y2 = e(z - 2) + e(z - 6)
    k22 = e(z - 2) - e(z - 6)

    e(z) = y0 + y2      ' z0 + z4 + z2 + z6
    e(z - 2) = y0 - y2  ' z0 + z4 - z2 - z6

    ' done with y0,y2

    k33 = e(z - 3) - e(z - 7)

    e(z - 4) = k00 + k33    ' z0 - z4 + z3 - z7
    e(z - 6) = k00 - k33    ' z0 - z4 - z3 + z7

    ' done with k33

    k11 = e(z - 1) - e(z - 5)
    y1 = e(z - 1) + e(z - 5)
    y3 = e(z - 3) + e(z - 7)

    e(z - 1) = y1 + y3      ' z1 + z5 + z3 + z7
    e(z - 3) = y1 - y3      ' z1 + z5 - z3 - z7
    e(z - 5) = k11 - k22    ' z1 - z5 + z2 - z6
    e(z - 7) = k11 + k22    ' z1 - z5 - z2 + z6
End Sub

Private Sub imdct_step3_inner_s_loop_ld654(ByVal n As Long, e() As Single, ByVal i_off As Long, A() As Single, ByVal base_n As Long)
    Dim a_off As Long, A2 As Single, z As Long, base_ As Long
    Dim k00 As Single, k11 As Single, l00 As Single, l11 As Single
    a_off = base_n \ 8
    A2 = A(0 + a_off)
    z = i_off
    base_ = z - 16 * n

    Do While z > base_
        k00 = e(z) - e(z - 8)
        k11 = e(z - 1) - e(z - 9)
        l00 = e(z - 2) - e(z - 10)
        l11 = e(z - 3) - e(z - 11)
        e(z) = e(z) + e(z - 8)
        e(z - 1) = e(z - 1) + e(z - 9)
        e(z - 2) = e(z - 2) + e(z - 10)
        e(z - 3) = e(z - 3) + e(z - 11)
        e(z - 8) = k00
        e(z - 9) = k11
        e(z - 10) = (l00 + l11) * A2
        e(z - 11) = (l11 - l00) * A2

        k00 = e(z - 4) - e(z - 12)
        k11 = e(z - 5) - e(z - 13)
        l00 = e(z - 6) - e(z - 14)
        l11 = e(z - 7) - e(z - 15)
        e(z - 4) = e(z - 4) + e(z - 12)
        e(z - 5) = e(z - 5) + e(z - 13)
        e(z - 6) = e(z - 6) + e(z - 14)
        e(z - 7) = e(z - 7) + e(z - 15)
        e(z - 12) = k11
        e(z - 13) = -k00
        e(z - 14) = (l11 - l00) * A2
        e(z - 15) = (l00 + l11) * -A2

        iter_54 e, z
        iter_54 e, z - 8
        z = z - 16
    Loop
End Sub

' (VB) the blocktype's tables are passed in: A = f->A[blocktype], B, C, bitrev = f->bit_reverse[blocktype]
Private Sub inverse_mdct(buffer() As Single, ByVal n As Long, A() As Single, B() As Single, C() As Single, bitrev() As Long)
    Dim n2 As Long, n4 As Long, n8 As Long, l As Long, ld As Long
    Dim d As Long, e As Long, AA As Long, e_stop As Long
    Dim d0 As Long, d1 As Long, d2 As Long, d3 As Long, e0 As Long, e1 As Long
    Dim v40_20 As Single, v41_21 As Single
    Dim k0 As Long, k0_2 As Long, k1 As Long, lim As Long, i As Long, rlim As Long, r As Long, i_off As Long, A0 As Long
    Dim k4 As Long, br As Long, Ci As Long, Bi As Long
    Dim a02 As Single, a11 As Single, b0 As Single, b1 As Single, b2 As Single, b3 As Single
    Dim p0 As Single, p1 As Single, p2 As Single, p3 As Single

    n2 = n \ 2: n4 = n \ 4: n8 = n \ 8
    ' @OPTIMIZE: reduce register pressure by using fewer variables?
    ' float *buf2 = (float *) temp_alloc(f, n2 * sizeof(*buf2));
    If buf2Size < n2 Then
        buf2Size = n2
        ReDim buf2(0 To n2 - 1)
    End If
    ' twiddle factors: A

    ' IMDCT algorithm from "The use of multirate filter banks for coding of high quality digital audio"
    ' See notes about bugs in that paper in less-optimal implementation 'inverse_mdct_old' after this function.

    ' kernel from paper


    ' merged:
    '   copy and reflect spectral data
    '   step 0

    ' note that it turns out that the items added together during
    ' this step are, in fact, being added to themselves (as reflected
    ' by step 0). inexplicable inefficiency! this became obvious
    ' once I combined the passes.

    ' so there's a missing 'times 2' here (for adding X to itself).
    ' this propagates through linearly to the end, where the numbers
    ' are 1/2 too small, and need to be compensated for.

    d = n2 - 2      ' &buf2[n2-2]
    AA = 0
    e = 0           ' &buffer[0]
    e_stop = n2
    Do While e <> e_stop
        buf2(d + 1) = (buffer(e) * A(AA) - buffer(e + 2) * A(AA + 1))
        buf2(d) = (buffer(e) * A(AA + 1) + buffer(e + 2) * A(AA))
        d = d - 2
        AA = AA + 2
        e = e + 4
    Loop

    e = n2 - 3
    Do While d >= 0
        buf2(d + 1) = (-buffer(e + 2) * A(AA) - -buffer(e) * A(AA + 1))
        buf2(d) = (-buffer(e + 2) * A(AA + 1) + -buffer(e) * A(AA))
        d = d - 2
        AA = AA + 2
        e = e - 4
    Loop

    ' now we use symbolic names for these, so that we can
    ' possibly swap their meaning as we change which operations
    ' are in place

    ' u = buffer; v = buf2;

    ' step 2    (paper output is w, now u)
    ' this could be in place, but the data ends up in the wrong
    ' place... _somebody_'s got to swap it, so this is nominated
    AA = n2 - 8

    e0 = n4     ' &v[n4]
    e1 = 0      ' &v[0]

    d0 = n4     ' &u[n4]
    d1 = 0      ' &u[0]

    Do While AA >= 0
        v41_21 = buf2(e0 + 1) - buf2(e1 + 1)
        v40_20 = buf2(e0) - buf2(e1)
        buffer(d0 + 1) = buf2(e0 + 1) + buf2(e1 + 1)
        buffer(d0) = buf2(e0) + buf2(e1)
        buffer(d1 + 1) = v41_21 * A(AA + 4) - v40_20 * A(AA + 5)
        buffer(d1) = v40_20 * A(AA + 4) + v41_21 * A(AA + 5)

        v41_21 = buf2(e0 + 3) - buf2(e1 + 3)
        v40_20 = buf2(e0 + 2) - buf2(e1 + 2)
        buffer(d0 + 3) = buf2(e0 + 3) + buf2(e1 + 3)
        buffer(d0 + 2) = buf2(e0 + 2) + buf2(e1 + 2)
        buffer(d1 + 3) = v41_21 * A(AA) - v40_20 * A(AA + 1)
        buffer(d1 + 2) = v40_20 * A(AA) + v41_21 * A(AA + 1)

        AA = AA - 8

        d0 = d0 + 4
        d1 = d1 + 4
        e0 = e0 + 4
        e1 = e1 + 4
    Loop

    ' step 3
    ld = ilog(n) - 1 ' ilog is off-by-one from normal definitions

    ' optimized step 3:

    ' the original step3 loop can be nested r inside s or s inside r;
    ' it's written originally as s inside r, but this is dumb when r
    ' iterates many times, and s few. So I have two copies of it and
    ' switch between them halfway.

    ' this is iteration 0 of step 3
    imdct_step3_iter0_loop n \ 16, buffer, n2 - 1 - n4 * 0, -(n \ 8), A
    imdct_step3_iter0_loop n \ 16, buffer, n2 - 1 - n4 * 1, -(n \ 8), A

    ' this is iteration 1 of step 3
    imdct_step3_inner_r_loop n \ 32, buffer, n2 - 1 - n8 * 0, -(n \ 16), A, 16
    imdct_step3_inner_r_loop n \ 32, buffer, n2 - 1 - n8 * 1, -(n \ 16), A, 16
    imdct_step3_inner_r_loop n \ 32, buffer, n2 - 1 - n8 * 2, -(n \ 16), A, 16
    imdct_step3_inner_r_loop n \ 32, buffer, n2 - 1 - n8 * 3, -(n \ 16), A, 16

    l = 2
    Do While l < (ld - 3) \ 2
        k0 = n \ pow2tab(l + 2): k0_2 = k0 \ 2
        lim = pow2tab(l + 1)
        For i = 0 To lim - 1
            imdct_step3_inner_r_loop n \ pow2tab(l + 4), buffer, n2 - 1 - k0 * i, -k0_2, A, pow2tab(l + 3)
        Next
        l = l + 1
    Loop

    Do While l < ld - 6
        k0 = n \ pow2tab(l + 2): k1 = pow2tab(l + 3): k0_2 = k0 \ 2
        rlim = n \ pow2tab(l + 6)
        lim = pow2tab(l + 1)
        A0 = 0
        i_off = n2 - 1
        For r = rlim To 1 Step -1
            imdct_step3_inner_s_loop lim, buffer, i_off, -k0_2, A, A0, k1, k0
            A0 = A0 + k1 * 4
            i_off = i_off - 8
        Next
        l = l + 1
    Loop

    ' iterations with count:
    '   ld-6,-5,-4 all interleaved together
    '       the big win comes from getting rid of needless flops
    '         due to the constants on pass 5 & 4 being all 1 and 0;
    '       combining them to be simultaneous to improve cache made little difference
    imdct_step3_inner_s_loop_ld654 n \ 32, buffer, n2 - 1, A, n

    ' output is u

    ' step 4, 5, and 6
    ' cannot be in-place because of step 5
    ' weirdly, I'd have thought reading sequentially and writing
    ' erratically would have been better than vice-versa, but in
    ' fact that's not what my testing showed. (That is, with
    ' j = bitreverse(i), do you read i and write j, or read j and write i.)
    br = 0
    d0 = n4 - 4     ' &v[n4-4]
    d1 = n2 - 4     ' &v[n2-4]
    Do While d0 >= 0
        k4 = bitrev(br)
        buf2(d1 + 3) = buffer(k4 + 0)
        buf2(d1 + 2) = buffer(k4 + 1)
        buf2(d0 + 3) = buffer(k4 + 2)
        buf2(d0 + 2) = buffer(k4 + 3)

        k4 = bitrev(br + 1)
        buf2(d1 + 1) = buffer(k4 + 0)
        buf2(d1 + 0) = buffer(k4 + 1)
        buf2(d0 + 1) = buffer(k4 + 2)
        buf2(d0 + 0) = buffer(k4 + 3)

        d0 = d0 - 4
        d1 = d1 - 4
        br = br + 2
    Loop
    ' (paper output is u, now v)


    ' data must be in buf2
    ' assert(v == buf2);

    ' step 7   (paper output is v, now v)
    ' this is now in place
    Ci = 0
    d = 0
    e = n2 - 4

    Do While d < e
        a02 = buf2(d) - buf2(e + 2)
        a11 = buf2(d + 1) + buf2(e + 3)

        b0 = C(Ci + 1) * a02 + C(Ci) * a11
        b1 = C(Ci + 1) * a11 - C(Ci) * a02

        b2 = buf2(d) + buf2(e + 2)
        b3 = buf2(d + 1) - buf2(e + 3)

        buf2(d) = b2 + b0
        buf2(d + 1) = b3 + b1
        buf2(e + 2) = b2 - b0
        buf2(e + 3) = b1 - b3

        a02 = buf2(d + 2) - buf2(e)
        a11 = buf2(d + 3) + buf2(e + 1)

        b0 = C(Ci + 3) * a02 + C(Ci + 2) * a11
        b1 = C(Ci + 3) * a11 - C(Ci + 2) * a02

        b2 = buf2(d + 2) + buf2(e)
        b3 = buf2(d + 3) - buf2(e + 1)

        buf2(d + 2) = b2 + b0
        buf2(d + 3) = b3 + b1
        buf2(e) = b2 - b0
        buf2(e + 1) = b1 - b3

        Ci = Ci + 4
        d = d + 4
        e = e - 4
    Loop

    ' data must be in buf2


    ' step 8+decode   (paper output is X, now buffer)
    ' this generates pairs of data a la 8 and pushes them directly through
    ' the decode kernel (pushing rather than pulling) to avoid having
    ' to make another pass later

    ' this cannot POSSIBLY be in place, so we refer to the buffers directly

    Bi = n2 - 8     ' f->B[blocktype] + n2 - 8
    e = n2 - 8      ' buf2 + n2 - 8
    d0 = 0          ' &buffer[0]
    d1 = n2 - 4     ' &buffer[n2-4]
    d2 = n2         ' &buffer[n2]
    d3 = n - 4      ' &buffer[n-4]
    Do While e >= 0
        p3 = buf2(e + 6) * B(Bi + 7) - buf2(e + 7) * B(Bi + 6)
        p2 = -buf2(e + 6) * B(Bi + 6) - buf2(e + 7) * B(Bi + 7)

        buffer(d0) = p3
        buffer(d1 + 3) = -p3
        buffer(d2) = p2
        buffer(d3 + 3) = p2

        p1 = buf2(e + 4) * B(Bi + 5) - buf2(e + 5) * B(Bi + 4)
        p0 = -buf2(e + 4) * B(Bi + 4) - buf2(e + 5) * B(Bi + 5)

        buffer(d0 + 1) = p1
        buffer(d1 + 2) = -p1
        buffer(d2 + 1) = p0
        buffer(d3 + 2) = p0

        p3 = buf2(e + 2) * B(Bi + 3) - buf2(e + 3) * B(Bi + 2)
        p2 = -buf2(e + 2) * B(Bi + 2) - buf2(e + 3) * B(Bi + 3)

        buffer(d0 + 2) = p3
        buffer(d1 + 1) = -p3
        buffer(d2 + 2) = p2
        buffer(d3 + 1) = p2

        p1 = buf2(e + 0) * B(Bi + 1) - buf2(e + 1) * B(Bi + 0)
        p0 = -buf2(e + 0) * B(Bi + 0) - buf2(e + 1) * B(Bi + 1)

        buffer(d0 + 3) = p1
        buffer(d1 + 0) = -p1
        buffer(d2 + 3) = p0
        buffer(d3 + 0) = p0

        Bi = Bi - 8
        e = e - 8
        d0 = d0 + 4
        d2 = d2 + 4
        d1 = d1 - 4
        d3 = d3 - 4
    Loop
End Sub

' the window's index (0 or 1) for len, -1 for NULL
Private Function get_window(f As vorb, ByVal len_ As Long) As Long
    len_ = len_ * 2
    If len_ = f.blocksize_0 Then get_window = 0: Exit Function
    If len_ = f.blocksize_1 Then get_window = 1: Exit Function
    get_window = -1
End Function

' (!STB_VORBIS_NO_DEFER_FLOOR: YTYPE is int16; step2_flag is not used)
Private Function do_floor(f As vorb, map As MappingT, ByVal i As Long, ByVal n As Long, target() As Single, finalY() As Integer) As Long
    Dim n2 As Long, s As Long, floor_ As Long
    n2 = n \ 2
    s = map.chan(i).mux
    floor_ = map.submap_floor(s)
    If f.floor_types(floor_) = 0 Then
        do_floor = error_(f, VORBIS_invalid_stream)
    Else
        do_floor1 f.floor_config(floor_).floor1, n2, target, finalY
        do_floor = 1
    End If
End Function

' (VB) do_floor's floor 1 part: g is &f->floor_config[floor].floor1
Private Sub do_floor1(g As Floor1T, ByVal n2 As Long, target() As Single, finalY() As Integer)
    Dim j As Long, q As Long, lx As Long, ly As Long, hy As Long, hx As Long, w As Single
    lx = 0: ly = finalY(0) * g.floor1_multiplier
    For q = 1 To g.values - 1
        j = g.sorted_order(q)
        If finalY(j) >= 0 Then
            hy = finalY(j) * g.floor1_multiplier
            hx = g.Xlist(j)
            If lx <> hx Then draw_line target, lx, ly, hx, hy, n2
            lx = hx: ly = hy
        End If
    Next
    If lx < n2 Then
        ' optimization of: draw_line(target, lx,ly, n,ly, n2);
        w = inverse_db_table(ly)
        For j = lx To n2 - 1
            target(j) = target(j) * w       ' LINE_OP(target[j], inverse_db_table[ly])
        Next
    End If
End Sub

' The meaning of "left" and "right"
'
' For a given frame:
'     we compute samples from 0..n
'     window_center is n/2
'     we'll window and mix the samples from left_start to left_end with data from the previous frame
'     all of the samples from left_end to right_start can be output without mixing; however,
'        this interval is 0-length except when transitioning between short and long frames
'     all of the samples from right_start to right_end need to be mixed with the next frame,
'        which we don't have, so those get saved in a buffer
'     frame N's right_end-right_start, the number of samples to mix with the next frame,
'        has to be the same as frame N+1's left_end-left_start (which they are by
'        construction)

Private Function vorbis_decode_initial(f As vorb, p_left_start As Long, p_left_end As Long, p_right_start As Long, p_right_end As Long, mode As Long) As Long
    Dim i As Long, n As Long, prev As Long, next_ As Long, window_center As Long, blockflag As Long
    f.channel_buffer_start = 0: f.channel_buffer_end = 0

retry:
    If f.eof_ <> 0 Then Exit Function
    If maybe_start_packet(f) = 0 Then Exit Function
    ' check packet type
    If get_bits(f, 1) <> 0 Then
        ' (IS_PUSH_MODE(f) is FALSE)
        Do While EOP <> get8_packet(f)
        Loop
        GoTo retry
    End If

    i = get_bits(f, ilog(f.mode_count - 1))
    If i = EOP Then Exit Function
    If i >= f.mode_count Then Exit Function
    mode = i
    blockflag = f.mode_config(i).blockflag  ' m = f->mode_config + i
    If blockflag <> 0 Then
        n = f.blocksize_1
        prev = get_bits(f, 1)
        next_ = get_bits(f, 1)
    Else
        prev = 0: next_ = 0
        n = f.blocksize_0
    End If

' WINDOWING

    window_center = n \ 2
    If blockflag <> 0 And prev = 0 Then
        p_left_start = (n - f.blocksize_0) \ 4
        p_left_end = (n + f.blocksize_0) \ 4
    Else
        p_left_start = 0
        p_left_end = window_center
    End If
    If blockflag <> 0 And next_ = 0 Then
        p_right_start = (n * 3 - f.blocksize_0) \ 4
        p_right_end = (n * 3 + f.blocksize_0) \ 4
    Else
        p_right_start = window_center
        p_right_end = n
    End If

    vorbis_decode_initial = 1
End Function

' (VB) the floor 1 decode of a channel in vorbis_decode_packet_rest (the "if (get_bits(f, 1))" block):
' 0 for its "goto error"
Private Function decode_floor1(f As vorb, g As Floor1T, finalY() As Integer, step2_flag() As Byte) As Long
    Dim range_ As Long, offset As Long, j As Long, k As Long, pclass As Long, cdim As Long, cbits As Long, csub As Long
    Dim cval As Long, book As Long, temp As Long, low As Long, high As Long, pred As Long, highroom As Long
    Dim lowroom As Long, room As Long, val_ As Long

    range_ = range_list(g.floor1_multiplier - 1)
    offset = 2
    finalY(0) = get_bits(f, ilog(range_) - 1)
    finalY(1) = get_bits(f, ilog(range_) - 1)
    For j = 0 To g.partitions - 1
        pclass = g.partition_class_list(j)
        cdim = g.class_dimensions(pclass)
        cbits = g.class_subclasses(pclass)
        csub = pow2tab(cbits) - 1
        cval = 0
        If cbits <> 0 Then
            cval = DECODE(f, f.codebooks(g.class_masterbooks(pclass)))
        End If
        For k = 0 To cdim - 1
            book = g.subclass_books(pclass, cval And csub)
            ' cval = cval >> cbits (arithmetic)
            If cval >= 0 Then cval = cval \ pow2tab(cbits) Else cval = -((-cval - 1) \ pow2tab(cbits)) - 1
            If book >= 0 Then
                temp = DECODE(f, f.codebooks(book))
                finalY(offset) = temp
                offset = offset + 1
            Else
                finalY(offset) = 0
                offset = offset + 1
            End If
        Next
    Next
    If f.valid_bits = INVALID_BITS Then decode_floor1 = 0: Exit Function ' behavior according to spec
    step2_flag(0) = 1: step2_flag(1) = 1
    For j = 2 To g.values - 1
        low = g.neighbors(j, 0)
        high = g.neighbors(j, 1)
        'neighbors(g->Xlist, j, &low, &high);
        pred = predict_point(g.Xlist(j), g.Xlist(low), g.Xlist(high), finalY(low), finalY(high))
        val_ = finalY(j)
        highroom = range_ - pred
        lowroom = pred
        If highroom < lowroom Then
            room = highroom * 2
        Else
            room = lowroom * 2
        End If
        If val_ <> 0 Then
            step2_flag(low) = 1: step2_flag(high) = 1
            step2_flag(j) = 1
            If val_ >= room Then
                If highroom > lowroom Then
                    finalY(j) = val_ - lowroom + pred
                Else
                    finalY(j) = pred - val_ + highroom - 1
                End If
            Else
                If (val_ And 1) <> 0 Then
                    finalY(j) = pred - (val_ + 1) \ 2
                Else
                    finalY(j) = pred + val_ \ 2
                End If
            End If
        Else
            step2_flag(j) = 0
            finalY(j) = pred
        End If
    Next

    ' defer final floor computation until _after_ residue
    For j = 0 To g.values - 1
        If step2_flag(j) = 0 Then finalY(j) = -1
    Next
    decode_floor1 = 1
End Function

' mi: the mode's index (m = f->mode_config + mi)
Private Function vorbis_decode_packet_rest(f As vorb, len_ As Long, ByVal mi As Long, ByVal left_start As Long, ByVal left_end As Long, ByVal right_start As Long, ByVal right_end As Long, p_left As Long) As Long
    Static zero_channel(0 To 255) As Long, really_zero_channel(0 To 255) As Long
    Static step2_flag(0 To 255) As Byte
    Static residue_buffers(0 To STB_VORBIS_MAX_CHANNELS - 1) As Long, do_not_decode(0 To 255) As Long
    Dim mapi As Long, i As Long, j As Long, n As Long, n2 As Long, s As Long, floor_ As Long, ch As Long, r As Long
    Dim mg As Long, an As Long, bf As Long, current_end As Long

' WINDOWING

    ' STBV_NOTUSED(left_end);
    bf = f.mode_config(mi).blockflag
    n = f.blocksize(bf)
    mapi = f.mode_config(mi).mapping    ' map = &f->mapping[m->mapping]

' FLOORS
    n2 = n \ 2

    For i = 0 To f.channels - 1
        s = f.mapping(mapi).chan(i).mux
        zero_channel(i) = 0
        floor_ = f.mapping(mapi).submap_floor(s)
        If f.floor_types(floor_) = 0 Then
            vorbis_decode_packet_rest = error_(f, VORBIS_invalid_stream)
            Exit Function
        Else
            If get_bits(f, 1) <> 0 Then
                If decode_floor1(f, f.floor_config(floor_).floor1, f.finalY(i).v, step2_flag) = 0 Then
                    zero_channel(i) = 1     ' error:
                End If
            Else
                zero_channel(i) = 1
            End If
            ' So we just defer everything else to later

            ' at this point we've decoded the floor into buffer
        End If
    Next
    ' at this point we've decoded all floors

    ' re-enable coupled channels if necessary
    For i = 0 To f.channels - 1
        really_zero_channel(i) = zero_channel(i)
    Next
    For i = 0 To f.mapping(mapi).coupling_steps - 1
        mg = f.mapping(mapi).chan(i).magnitude
        an = f.mapping(mapi).chan(i).angle
        If zero_channel(mg) = 0 Or zero_channel(an) = 0 Then
            zero_channel(mg) = 0: zero_channel(an) = 0
        End If
    Next

' RESIDUE DECODE
    For i = 0 To f.mapping(mapi).submaps - 1
        ch = 0
        For j = 0 To f.channels - 1
            If f.mapping(mapi).chan(j).mux = i Then
                If zero_channel(j) <> 0 Then
                    do_not_decode(ch) = 1
                    residue_buffers(ch) = -1
                Else
                    do_not_decode(ch) = 0
                    residue_buffers(ch) = j
                End If
                ch = ch + 1
            End If
        Next
        r = f.mapping(mapi).submap_residue(i)
        decode_residue f, f.residue_config(r), residue_buffers, ch, n2, r, do_not_decode
    Next

' INVERSE COUPLING
    For i = f.mapping(mapi).coupling_steps - 1 To 0 Step -1
        inverse_coupling f.channel_buffers(f.mapping(mapi).chan(i).magnitude).v, f.channel_buffers(f.mapping(mapi).chan(i).angle).v, n2
    Next

    ' finish decoding the floors
    For i = 0 To f.channels - 1
        If really_zero_channel(i) <> 0 Then
            ZeroMemory f.channel_buffers(i).v(0), 4 * n2
        Else
            do_floor f, f.mapping(mapi), i, n, f.channel_buffers(i).v, f.finalY(i).v
        End If
    Next

' INVERSE MDCT
    For i = 0 To f.channels - 1
        inverse_mdct f.channel_buffers(i).v, n, f.A(bf).v, f.B(bf).v, f.C(bf).v, f.bit_reverse(bf).v
    Next

    ' this shouldn't be necessary, unless we exited on an error
    ' and want to flush to get to the next packet
    flush_packet f

    If f.first_decode <> 0 Then
        ' assume we start so first non-discarded sample is sample 0
        ' this isn't to spec, but spec would require us to read ahead
        ' and decode the size of all current frames--could be done,
        ' but presumably it's not a commonly used feature
        f.current_loc = -n2 ' 0u - n2: start of first frame is positioned for discard (NB this is an intentional unsigned overflow/wrap-around)
        ' we might have to discard samples "from" the next frame too,
        ' if we're lapping a large block then a small at the start?
        f.discard_samples_deferred = n - right_end
        f.current_loc_valid = 1
        f.first_decode = 0
    ElseIf f.discard_samples_deferred <> 0 Then
        If f.discard_samples_deferred >= right_start - left_start Then
            f.discard_samples_deferred = f.discard_samples_deferred - (right_start - left_start)
            left_start = right_start
            p_left = left_start
        Else
            left_start = left_start + f.discard_samples_deferred
            p_left = left_start
            f.discard_samples_deferred = 0
        End If
    ElseIf f.previous_length = 0 And f.current_loc_valid <> 0 Then
        ' we're recovering from a seek... that means we're going to discard
        ' the samples from this packet even though we know our position from
        ' the last page header, so we need to update the position based on
        ' the discarded samples here
        ' but wait, the code below is going to add this in itself even
        ' on a discard, so we don't need to do it here...
    End If

    ' check if we have ogg information about the sample # for this packet
    If f.last_seg_which = f.end_seg_with_known_loc Then
        ' if we have a valid current loc, and this is final:
        If f.current_loc_valid <> 0 And (f.page_flag And PAGEFLAG_last_page) <> 0 Then
            current_end = f.known_loc_for_packet
            ' then let's infer the size of the (probably) short final frame
            If ult32(current_end, add32(f.current_loc, right_end - left_start)) Then
                If ult32(current_end, f.current_loc) Then
                    ' negative truncation, that's impossible!
                    len_ = 0
                Else
                    len_ = sub32(current_end, f.current_loc)
                End If
                len_ = len_ + left_start ' this doesn't seem right, but has no ill effect on my test files
                If len_ > right_end Then len_ = right_end ' this should never happen
                f.current_loc = add32(f.current_loc, len_)
                vorbis_decode_packet_rest = 1
                Exit Function
            End If
        End If
        ' otherwise, just set our sample loc
        ' guess that the ogg granule pos refers to the _middle_ of the
        ' last frame?
        ' set f->current_loc to the position of left_start
        f.current_loc = sub32(f.known_loc_for_packet, n2 - left_start)
        f.current_loc_valid = 1
    End If
    If f.current_loc_valid <> 0 Then
        f.current_loc = add32(f.current_loc, right_start - left_start)
    End If

    len_ = right_end  ' ignore samples after the window goes to 0

    vorbis_decode_packet_rest = 1
End Function

' (VB) the inverse coupling of a step: m = f->channel_buffers[map->chan[i].magnitude], a = ...angle
Private Sub inverse_coupling(m() As Single, a() As Single, ByVal n2 As Long)
    Dim j As Long, a2 As Single, m2 As Single, mj As Single, aj As Single
    For j = 0 To n2 - 1
        mj = m(j): aj = a(j)
        If mj > 0 Then
            If aj > 0 Then
                m2 = mj: a2 = mj - aj
            Else
                a2 = mj: m2 = mj + aj
            End If
        Else
            If aj > 0 Then
                m2 = mj: a2 = mj + aj
            Else
                a2 = mj: m2 = mj - aj
            End If
        End If
        m(j) = m2
        a(j) = a2
    Next
End Sub

Private Function vorbis_decode_packet(f As vorb, len_ As Long, p_left As Long, p_right As Long) As Long
    Dim mode As Long, left_end As Long, right_end As Long
    If vorbis_decode_initial(f, p_left, left_end, p_right, right_end, mode) = 0 Then Exit Function
    vorbis_decode_packet = vorbis_decode_packet_rest(f, len_, mode, p_left, left_end, p_right, right_end, p_left)
End Function

Private Function vorbis_finish_frame(f As vorb, ByVal len_ As Long, ByVal left_ As Long, ByVal right_ As Long) As Long
    Dim prev As Long, i As Long, n As Long, w As Long
    ' we use right&left (the start of the right- and left-window sin()-regions)
    ' to determine how much to return, rather than inferring from the rules
    ' (same result, clearer code); 'left' indicates where our sin() window
    ' starts, therefore where the previous window's right edge starts, and
    ' therefore where to start mixing from the previous buffer. 'right'
    ' indicates where our sin() ending-window starts, therefore that's where
    ' we start saving, and where our returned-data ends.

    ' mixin from previous window
    If f.previous_length <> 0 Then
        n = f.previous_length
        w = get_window(f, n)
        If w < 0 Then vorbis_finish_frame = 0: Exit Function
        For i = 0 To f.channels - 1
            mix_previous f.channel_buffers(i).v, f.previous_window(i).v, f.window_(w).v, left_, n
        Next
    End If

    prev = f.previous_length

    ' last half of this data becomes previous window
    f.previous_length = len_ - right_

    ' @OPTIMIZE: could avoid this copy by double-buffering the
    ' output (flipping previous_window with channel_buffers), but
    ' then previous_window would have to be 2x as large, and
    ' channel_buffers couldn't be temp mem (although they're NOT
    ' currently temp mem, they could be (unless we want to level
    ' performance by spreading out the computation))
    If len_ - right_ > 0 Then
        For i = 0 To f.channels - 1
            CopyMemory f.previous_window(i).v(0), f.channel_buffers(i).v(right_), 4 * (len_ - right_)
        Next
    End If

    If prev = 0 Then
        ' there was no previous packet, so this data isn't valid...
        ' this isn't entirely true, only the would-have-overlapped data
        ' isn't valid, but this seems to be what the spec requires
        vorbis_finish_frame = 0
        Exit Function
    End If

    ' truncate a short frame
    If len_ < right_ Then right_ = len_

    f.samples_output = add32(f.samples_output, right_ - left_)

    vorbis_finish_frame = right_ - left_
End Function

' (VB) vorbis_finish_frame's mixing of a channel: cb = f->channel_buffers[i], pw = f->previous_window[i]
Private Sub mix_previous(cb() As Single, pw() As Single, w() As Single, ByVal left_ As Long, ByVal n As Long)
    Dim j As Long
    For j = 0 To n - 1
        cb(left_ + j) = cb(left_ + j) * w(j) + pw(j) * w(n - 1 - j)
    Next
End Sub

Private Function vorbis_pump_first_frame(f As vorb) As Long
    Dim len_ As Long, right_ As Long, left_ As Long, res As Long
    res = vorbis_decode_packet(f, len_, left_, right_)
    If res <> 0 Then
        vorbis_finish_frame f, len_, left_, right_
    End If
    vorbis_pump_first_frame = res
End Function

Private Function start_decoder(f As vorb) As Long
    Dim header(0 To 5) As Byte, x As Long, len_ As Long, i As Long, j As Long, max_submaps As Long
    Dim longest_floorlist As Long, log0 As Long, log1 As Long, isSkeleton As Boolean, b() As Byte, z As Long
    Dim imdct_mem As Long, classify_mem As Long, max_part_read As Long, actual_size As Long
    Dim limit_r_begin As Long, limit_r_end As Long, n_read As Long, part_read As Long

    max_submaps = 0
    longest_floorlist = 0

    ' first page, first packet
    f.first_decode = 1

    If start_page(f) = 0 Then Exit Function
    ' validate page flag
    If (f.page_flag And PAGEFLAG_first_page) = 0 Then start_decoder = error_(f, VORBIS_invalid_first_page): Exit Function
    If (f.page_flag And PAGEFLAG_last_page) <> 0 Then start_decoder = error_(f, VORBIS_invalid_first_page): Exit Function
    If (f.page_flag And PAGEFLAG_continued_packet) <> 0 Then start_decoder = error_(f, VORBIS_invalid_first_page): Exit Function
    ' check for expected packet length
    If f.segment_count <> 1 Then start_decoder = error_(f, VORBIS_invalid_first_page): Exit Function
    If f.segments(0) <> 30 Then
        ' check for the Ogg skeleton fishead identifying header to refine our error
        If f.segments(0) = 64 Then
            If getn(f, header, 6) <> 0 Then
                If header(0) = &H66 And header(1) = &H69 And header(2) = &H73 And header(3) = &H68 And header(4) = &H65 And header(5) = &H61 Then
                    If get8(f) = &H64 Then
                        If get8(f) = 0 Then isSkeleton = True
                    End If
                End If
            End If
        End If
        If isSkeleton Then
            start_decoder = error_(f, VORBIS_ogg_skeleton_not_supported)
        Else
            start_decoder = error_(f, VORBIS_invalid_first_page)
        End If
        Exit Function
    End If

    ' read packet
    ' check packet header
    If get8(f) <> VORBIS_packet_id Then start_decoder = error_(f, VORBIS_invalid_first_page): Exit Function
    If getn(f, header, 6) = 0 Then start_decoder = error_(f, VORBIS_unexpected_eof): Exit Function
    If vorbis_validate(header) = 0 Then start_decoder = error_(f, VORBIS_invalid_first_page): Exit Function
    ' vorbis_version
    If get32(f) <> 0 Then start_decoder = error_(f, VORBIS_invalid_first_page): Exit Function
    f.channels = get8(f): If f.channels = 0 Then start_decoder = error_(f, VORBIS_invalid_first_page): Exit Function
    If f.channels > STB_VORBIS_MAX_CHANNELS Then start_decoder = error_(f, VORBIS_too_many_channels): Exit Function
    f.sample_rate = get32(f): If f.sample_rate = 0 Then start_decoder = error_(f, VORBIS_invalid_first_page): Exit Function
    get32 f ' bitrate_maximum
    get32 f ' bitrate_nominal
    get32 f ' bitrate_minimum
    x = get8(f)
    log0 = x And 15
    log1 = x \ 16
    f.blocksize_0 = pow2tab(log0)
    f.blocksize_1 = pow2tab(log1)
    If log0 < 6 Or log0 > 13 Then start_decoder = error_(f, VORBIS_invalid_setup): Exit Function
    If log1 < 6 Or log1 > 13 Then start_decoder = error_(f, VORBIS_invalid_setup): Exit Function
    If log0 > log1 Then start_decoder = error_(f, VORBIS_invalid_setup): Exit Function

    ' framing_flag
    x = get8(f)
    If (x And 1) = 0 Then start_decoder = error_(f, VORBIS_invalid_first_page): Exit Function

    ' second packet!
    If start_page(f) = 0 Then Exit Function

    If start_packet(f) = 0 Then Exit Function

    If next_segment(f) = 0 Then Exit Function

    If get8_packet(f) <> VORBIS_packet_comment Then start_decoder = error_(f, VORBIS_invalid_setup): Exit Function
    For i = 0 To 5
        header(i) = get8_packet(f) And &HFF&
    Next
    If vorbis_validate(header) = 0 Then start_decoder = error_(f, VORBIS_invalid_setup): Exit Function
    'file vendor
    len_ = get32_packet(f)
    setup_malloc f, len_ + 1
    f.vendor = read_packet_string(f, len_)
    'user comments
    f.comment_list_length = get32_packet(f)
    Erase f.comment_list
    If f.comment_list_length > 0 Then
        ReDim f.comment_list(0 To f.comment_list_length - 1)
        setup_malloc f, 4 * f.comment_list_length
    End If

    For i = 0 To f.comment_list_length - 1
        len_ = get32_packet(f)
        setup_malloc f, len_ + 1
        f.comment_list(i) = read_packet_string(f, len_)
    Next

    ' framing_flag
    x = get8_packet(f) And &HFF&
    If (x And 1) = 0 Then start_decoder = error_(f, VORBIS_invalid_setup): Exit Function


    skip f, f.bytes_in_seg
    f.bytes_in_seg = 0

    Do
        len_ = next_segment(f)
        skip f, len_
        f.bytes_in_seg = 0
    Loop While len_ <> 0

    ' third packet!
    If start_packet(f) = 0 Then Exit Function

    crc32_init ' always init it, to avoid multithread race conditions

    If get8_packet(f) <> VORBIS_packet_setup Then start_decoder = error_(f, VORBIS_invalid_setup): Exit Function
    For i = 0 To 5
        header(i) = get8_packet(f) And &HFF&
    Next
    If vorbis_validate(header) = 0 Then start_decoder = error_(f, VORBIS_invalid_setup): Exit Function

    ' codebooks

    f.codebook_count = get_bits(f, 8) + 1
    ReDim f.codebooks(0 To f.codebook_count - 1)
    setup_malloc f, 2120 * f.codebook_count   ' sizeof(*f->codebooks) (32-bit)
    For i = 0 To f.codebook_count - 1
        If start_decoder_codebook(f, f.codebooks(i)) = 0 Then Exit Function
    Next

    ' time domain transfers (notused)

    x = get_bits(f, 6) + 1
    For i = 0 To x - 1
        z = get_bits(f, 16)
        If z <> 0 Then start_decoder = error_(f, VORBIS_invalid_setup): Exit Function
    Next

    ' Floors
    f.floor_count = get_bits(f, 6) + 1
    ReDim f.floor_config(0 To f.floor_count - 1)
    setup_malloc f, 1356 * f.floor_count      ' f->floor_count * sizeof(*f->floor_config)
    For i = 0 To f.floor_count - 1
        f.floor_types(i) = get_bits(f, 16)
        If f.floor_types(i) > 1 Then start_decoder = error_(f, VORBIS_invalid_setup): Exit Function
        If f.floor_types(i) = 0 Then
            start_decoder_floor0 f, f.floor_config(i).floor0
            start_decoder = error_(f, VORBIS_feature_not_supported)
            Exit Function
        Else
            If start_decoder_floor1(f, f.floor_config(i).floor1, longest_floorlist) = 0 Then Exit Function
        End If
    Next

    ' Residue
    f.residue_count = get_bits(f, 6) + 1
    ReDim f.residue_config(0 To f.residue_count - 1)
    setup_malloc f, 20 * f.residue_count      ' f->residue_count * sizeof(f->residue_config[0])
    For i = 0 To f.residue_count - 1
        If start_decoder_residue(f, f.residue_config(i), i) = 0 Then Exit Function
    Next

    f.mapping_count = get_bits(f, 6) + 1
    ReDim f.mapping(0 To f.mapping_count - 1)
    setup_malloc f, 36 * f.mapping_count      ' f->mapping_count * sizeof(*f->mapping)
    For i = 0 To f.mapping_count - 1
        If start_decoder_mapping(f, f.mapping(i), max_submaps) = 0 Then Exit Function
    Next

    ' Modes
    f.mode_count = get_bits(f, 6) + 1
    For i = 0 To f.mode_count - 1
        f.mode_config(i).blockflag = get_bits(f, 1)
        f.mode_config(i).windowtype = get_bits(f, 16)
        f.mode_config(i).transformtype = get_bits(f, 16)
        f.mode_config(i).mapping = get_bits(f, 8)
        If f.mode_config(i).windowtype <> 0 Then start_decoder = error_(f, VORBIS_invalid_setup): Exit Function
        If f.mode_config(i).transformtype <> 0 Then start_decoder = error_(f, VORBIS_invalid_setup): Exit Function
        If f.mode_config(i).mapping >= f.mapping_count Then start_decoder = error_(f, VORBIS_invalid_setup): Exit Function
    Next

    flush_packet f

    f.previous_length = 0

    For i = 0 To f.channels - 1
        ReDim f.channel_buffers(i).v(0 To f.blocksize_1 - 1)     ' (memset to 0: VB's new array)
        setup_malloc f, 4 * f.blocksize_1
        ReDim f.previous_window(i).v(0 To f.blocksize_1 \ 2 - 1)
        setup_malloc f, 4 * f.blocksize_1 \ 2
        If longest_floorlist > 0 Then ReDim f.finalY(i).v(0 To longest_floorlist - 1)
        setup_malloc f, 2 * longest_floorlist
    Next

    If init_blocksize(f, 0, f.blocksize_0) = 0 Then Exit Function
    If init_blocksize(f, 1, f.blocksize_1) = 0 Then Exit Function
    f.blocksize(0) = f.blocksize_0
    f.blocksize(1) = f.blocksize_1

    ' compute how much temporary memory is needed

    ' 1.
    imdct_mem = (f.blocksize_1 * 4) \ 2     ' (f->blocksize_1 * sizeof(float) >> 1)
    max_part_read = 0
    For i = 0 To f.residue_count - 1
        actual_size = f.blocksize_1 \ 2
        If f.residue_config(i).begin_ < actual_size Then limit_r_begin = f.residue_config(i).begin_ Else limit_r_begin = actual_size
        If f.residue_config(i).end_ < actual_size Then limit_r_end = f.residue_config(i).end_ Else limit_r_end = actual_size
        n_read = limit_r_end - limit_r_begin
        part_read = n_read \ f.residue_config(i).part_size
        If part_read > max_part_read Then max_part_read = part_read
    Next
    classify_mem = f.channels * (4 + max_part_read * 4)   ' (sizeof(void*) + max_part_read * sizeof(uint8 *))

    ' maximum reasonable partition size is f->blocksize_1

    f.temp_memory_required = classify_mem
    If imdct_mem > f.temp_memory_required Then f.temp_memory_required = imdct_mem

    ' @TODO: stb_vorbis_seek_start expects first_audio_page_offset to point to a page
    ' without PAGEFLAG_continued_packet, so this either points to the first page, or
    ' the page after the end of the headers. It might be cleaner to point to a page
    ' in the middle of the headers, when that's the page where the first audio packet
    ' starts, but we'd have to also correctly skip the end of any continued packet in
    ' stb_vorbis_seek_start.
    If f.next_seg = -1 Then
        f.first_audio_page_offset = stb_vorbis_get_file_offset(f)
    Else
        f.first_audio_page_offset = 0
    End If

    start_decoder = 1
End Function

' (VB) the vendor's and the comments' chars: len get8_packet()s
Private Function read_packet_string(f As vorb, ByVal len_ As Long) As String
    Dim b() As Byte, i As Long
    If len_ <= 0 Then Exit Function
    ReDim b(0 To len_ - 1)
    For i = 0 To len_ - 1
        b(i) = get8_packet(f) And &HFF&
    Next
    read_packet_string = StrConv(b, vbUnicode)
End Function

' (VB) start_decoder's codebook loop body: c is f->codebooks + i
Private Function start_decoder_codebook(f As vorb, c As CodebookT) As Long
    Dim values() As Long, ordered As Long, sorted_count As Long, total As Long, lengths() As Byte
    Dim x As Long, y As Long, j As Long, k As Long, current_entry As Long, current_length As Long, limit As Long
    Dim n As Long, present As Long, size As Double, mults() As Long, q As Long, vals As Long
    Dim len_ As Long, sparse As Long, last As Single, z As Double, div As Double, off As Long, val_ As Single

    total = 0
    x = get_bits(f, 8): If x <> &H42 Then start_decoder_codebook = error_(f, VORBIS_invalid_setup): Exit Function
    x = get_bits(f, 8): If x <> &H43 Then start_decoder_codebook = error_(f, VORBIS_invalid_setup): Exit Function
    x = get_bits(f, 8): If x <> &H56 Then start_decoder_codebook = error_(f, VORBIS_invalid_setup): Exit Function
    x = get_bits(f, 8)
    c.dimensions = get_bits(f, 8) * &H100& + x
    x = get_bits(f, 8)
    y = get_bits(f, 8)
    c.entries = get_bits(f, 8) * &H10000 + y * &H100& + x
    ordered = get_bits(f, 1)
    If ordered <> 0 Then c.sparse = 0 Else c.sparse = get_bits(f, 1)

    If c.dimensions = 0 And c.entries <> 0 Then start_decoder_codebook = error_(f, VORBIS_invalid_setup): Exit Function

    ' (VB) lengths is c->codeword_lengths when not sparse (copied in below), a temp array when sparse
    If c.entries > 0 Then ReDim lengths(0 To c.entries - 1) Else ReDim lengths(0 To 0)
    If c.sparse = 0 Then setup_malloc f, c.entries

    If ordered <> 0 Then
        current_entry = 0
        current_length = get_bits(f, 5) + 1
        Do While current_entry < c.entries
            limit = c.entries - current_entry
            n = get_bits(f, ilog(limit))
            If current_length >= 32 Then start_decoder_codebook = error_(f, VORBIS_invalid_setup): Exit Function
            If current_entry + n > c.entries Then start_decoder_codebook = error_(f, VORBIS_invalid_setup): Exit Function
            For j = current_entry To current_entry + n - 1
                lengths(j) = current_length     ' memset(lengths + current_entry, current_length, n)
            Next
            current_entry = current_entry + n
            current_length = current_length + 1
        Loop
    Else
        For j = 0 To c.entries - 1
            If c.sparse <> 0 Then present = get_bits(f, 1) Else present = 1
            If present <> 0 Then
                lengths(j) = get_bits(f, 5) + 1
                total = total + 1
                If lengths(j) = 32 Then start_decoder_codebook = error_(f, VORBIS_invalid_setup): Exit Function
            Else
                lengths(j) = NO_CODE
            End If
        Next
    End If

    If c.sparse <> 0 And total >= c.entries \ 4 Then
        ' convert sparse items to non-sparse!
        If c.entries > f.setup_temp_memory_required Then f.setup_temp_memory_required = c.entries

        setup_malloc f, c.entries   ' c->codeword_lengths = (uint8 *) setup_malloc(f, c->entries); memcpy
        ' setup_temp_free(f, lengths, c->entries); // note this is only safe if there have been no intervening temp mallocs!
        c.sparse = 0
    End If
    If c.sparse = 0 Then c.codeword_lengths = lengths   ' lengths = c->codeword_lengths

    ' compute the size of the sorted tables
    If c.sparse <> 0 Then
        sorted_count = total
    Else
        sorted_count = 0
        For j = 0 To c.entries - 1
            If lengths(j) > STB_VORBIS_FAST_HUFFMAN_LENGTH And lengths(j) <> NO_CODE Then
                sorted_count = sorted_count + 1
            End If
        Next
    End If

    c.sorted_entries = sorted_count
    ' values = NULL;

    If c.sparse = 0 Then
        If c.entries > 0 Then ReDim c.codewords(0 To c.entries - 1) Else ReDim c.codewords(0 To 0)
        c.has_codewords = True
        setup_malloc f, 4 * c.entries
    Else
        If c.sorted_entries <> 0 Then
            ReDim c.codeword_lengths(0 To c.sorted_entries - 1)
            setup_malloc f, c.sorted_entries
            ReDim c.codewords(0 To c.sorted_entries - 1)
            c.has_codewords = True
            ReDim values(0 To c.sorted_entries - 1)
        End If
        size = CDbl(c.entries) + (4 + 4) * CDbl(c.sorted_entries)
        If size > CDbl(f.setup_temp_memory_required) Then f.setup_temp_memory_required = size
    End If

    If compute_codewords(c, lengths, c.entries, values) = 0 Then
        ' if (c->sparse) setup_temp_free(f, values, 0);
        start_decoder_codebook = error_(f, VORBIS_invalid_setup)
        Exit Function
    End If

    If c.sorted_entries <> 0 Then
        ' allocate an extra slot for sentinels
        ReDim c.sorted_codewords(0 To c.sorted_entries)
        setup_malloc f, 4 * (c.sorted_entries + 1)
        ' allocate an extra slot at the front so that c->sorted_values[-1] is defined
        ' so that we can catch that case without an extra if
        ReDim c.sorted_values(-1 To c.sorted_entries - 1)
        setup_malloc f, 4 * (c.sorted_entries + 1)
        c.sorted_values(-1) = -1
        c.has_sorted = True
        compute_sorted_huffman c, lengths, values
    End If

    If c.sparse <> 0 Then
        ' setup_temp_free(f, values, ...); setup_temp_free(f, c->codewords, ...); setup_temp_free(f, lengths, ...);
        Erase c.codewords
        c.has_codewords = False     ' c->codewords = NULL;
    End If

    compute_accelerated_huffman c

    c.lookup_type = get_bits(f, 4)
    If c.lookup_type > 2 Then start_decoder_codebook = error_(f, VORBIS_invalid_setup): Exit Function
    If c.lookup_type > 0 Then
        c.minimum_value = float32_unpack(get_bits(f, 32))
        c.delta_value = float32_unpack(get_bits(f, 32))
        c.value_bits = get_bits(f, 4) + 1
        c.sequence_p = get_bits(f, 1)
        If c.lookup_type = 1 Then
            vals = lookup1_values(c.entries, c.dimensions)
            If vals < 0 Then start_decoder_codebook = error_(f, VORBIS_invalid_setup): Exit Function
            c.lookup_values = vals
        Else
            c.lookup_values = c.entries * c.dimensions     ' (uint32: below 2^31 for the game's files)
        End If
        If c.lookup_values = 0 Then start_decoder_codebook = error_(f, VORBIS_invalid_setup): Exit Function
        ReDim mults(0 To c.lookup_values - 1)   ' uint16 *mults
        For j = 0 To c.lookup_values - 1
            q = get_bits(f, c.value_bits)
            If q = EOP Then start_decoder_codebook = error_(f, VORBIS_invalid_setup): Exit Function
            mults(j) = q And &HFFFF&
        Next

        ' (!STB_VORBIS_DIVIDES_IN_CODEBOOK)
        If c.lookup_type = 1 Then
            sparse = c.sparse
            last = 0
            ' pre-expand the lookup1-style multiplicands, to avoid a divide in the inner loop
            If sparse <> 0 Then
                If c.sorted_entries = 0 Then GoTo skip_
                ReDim c.multiplicands(0 To c.sorted_entries * c.dimensions - 1)
                setup_malloc f, 4 * c.sorted_entries * c.dimensions
            Else
                ReDim c.multiplicands(0 To c.entries * c.dimensions - 1)
                setup_malloc f, 4 * c.entries * c.dimensions
            End If
            If sparse <> 0 Then len_ = c.sorted_entries Else len_ = c.entries
            For j = 0 To len_ - 1
                If sparse <> 0 Then z = c.sorted_values(j) Else z = j    ' unsigned int z
                div = 1     ' unsigned int div
                For k = 0 To c.dimensions - 1
                    off = (Int(z / div)) - Int(Int(z / div) / c.lookup_values) * c.lookup_values  ' (z / div) % c->lookup_values
                    val_ = mults(off) * c.delta_value + c.minimum_value + last
                    c.multiplicands(j * c.dimensions + k) = val_
                    If c.sequence_p <> 0 Then
                        last = val_
                    End If
                    If k + 1 < c.dimensions Then
                        If div > Int(4294967295# / c.lookup_values) Then   ' UINT_MAX / (unsigned int) c->lookup_values
                            ' setup_temp_free(f, mults,sizeof(mults[0])*c->lookup_values);
                            start_decoder_codebook = error_(f, VORBIS_invalid_setup)
                            Exit Function
                        End If
                        div = div * c.lookup_values
                    End If
                Next
            Next
            c.lookup_type = 2
        Else
            last = 0
            ReDim c.multiplicands(0 To c.lookup_values - 1)
            setup_malloc f, 4 * c.lookup_values
            For j = 0 To c.lookup_values - 1
                val_ = mults(j) * c.delta_value + c.minimum_value + last
                c.multiplicands(j) = val_
                If c.sequence_p <> 0 Then
                    last = val_
                End If
            Next
        End If
skip_:
        ' setup_temp_free(f, mults, sizeof(mults[0])*c->lookup_values);
    End If
    start_decoder_codebook = 1
End Function

' (VB) start_decoder's floor 0 header (the floor is not supported: the decoder fails after it)
Private Sub start_decoder_floor0(f As vorb, g As Floor0T)
    Dim j As Long
    g.order = get_bits(f, 8)
    g.rate = get_bits(f, 16)
    g.bark_map_size = get_bits(f, 16)
    g.amplitude_bits = get_bits(f, 6)
    g.amplitude_offset = get_bits(f, 8)
    g.number_of_books = get_bits(f, 4) + 1
    For j = 0 To g.number_of_books - 1
        g.book_list(j) = get_bits(f, 8)
    Next
End Sub

' (VB) start_decoder's floor 1: g is &f->floor_config[i].floor1
Private Function start_decoder_floor1(f As vorb, g As Floor1T, longest_floorlist As Long) As Long
    Dim p_x(0 To 31 * 8 + 1) As Long, p_id(0 To 31 * 8 + 1) As Long     ' stbv__floor_ordering p[31*8+2]
    Dim max_class As Long, j As Long, k As Long, c As Long, low As Long, hi As Long, t As Long, u As Long, m As Long
    max_class = -1
    g.partitions = get_bits(f, 5)
    For j = 0 To g.partitions - 1
        g.partition_class_list(j) = get_bits(f, 4)
        If g.partition_class_list(j) > max_class Then
            max_class = g.partition_class_list(j)
        End If
    Next
    For j = 0 To max_class
        g.class_dimensions(j) = get_bits(f, 3) + 1
        g.class_subclasses(j) = get_bits(f, 2)
        If g.class_subclasses(j) <> 0 Then
            g.class_masterbooks(j) = get_bits(f, 8)
            If g.class_masterbooks(j) >= f.codebook_count Then start_decoder_floor1 = error_(f, VORBIS_invalid_setup): Exit Function
        End If
        For k = 0 To pow2tab(g.class_subclasses(j)) - 1
            g.subclass_books(j, k) = get_bits(f, 8) - 1
            If g.subclass_books(j, k) >= f.codebook_count Then start_decoder_floor1 = error_(f, VORBIS_invalid_setup): Exit Function
        Next
    Next
    g.floor1_multiplier = get_bits(f, 2) + 1
    g.rangebits = get_bits(f, 4)
    g.Xlist(0) = 0
    g.Xlist(1) = pow2tab(g.rangebits)
    g.values = 2
    For j = 0 To g.partitions - 1
        c = g.partition_class_list(j)
        For k = 0 To g.class_dimensions(c) - 1
            g.Xlist(g.values) = get_bits(f, g.rangebits)
            g.values = g.values + 1
        Next
    Next
    ' precompute the sorting
    For j = 0 To g.values - 1
        p_x(j) = g.Xlist(j)
        p_id(j) = j
    Next
    ' qsort(p, g->values, sizeof(p[0]), point_compare): (VB) an insertion sort (equal x are an error below)
    For j = 1 To g.values - 1
        t = p_x(j): u = p_id(j): m = j - 1
        Do While m >= 0
            If p_x(m) <= t Then Exit Do
            p_x(m + 1) = p_x(m): p_id(m + 1) = p_id(m)
            m = m - 1
        Loop
        p_x(m + 1) = t: p_id(m + 1) = u
    Next
    For j = 0 To g.values - 2
        If p_x(j) = p_x(j + 1) Then start_decoder_floor1 = error_(f, VORBIS_invalid_setup): Exit Function
    Next
    For j = 0 To g.values - 1
        g.sorted_order(j) = p_id(j)
    Next
    ' precompute the neighbors
    For j = 2 To g.values - 1
        low = 0: hi = 0
        neighbors g.Xlist, j, low, hi
        g.neighbors(j, 0) = low
        g.neighbors(j, 1) = hi
    Next

    If g.values > longest_floorlist Then longest_floorlist = g.values
    start_decoder_floor1 = 1
End Function

' (VB) start_decoder's residue loop body: r is f->residue_config+i
Private Function start_decoder_residue(f As vorb, r As ResidueT, ByVal i As Long) As Long
    Dim residue_cascade(0 To 63) As Long, j As Long, k As Long, high_bits As Long, low_bits As Long
    Dim classwords As Long, entries As Long, temp As Long
    f.residue_types(i) = get_bits(f, 16)
    If f.residue_types(i) > 2 Then start_decoder_residue = error_(f, VORBIS_invalid_setup): Exit Function
    r.begin_ = get_bits(f, 24)
    r.end_ = get_bits(f, 24)
    If r.end_ < r.begin_ Then start_decoder_residue = error_(f, VORBIS_invalid_setup): Exit Function
    r.part_size = get_bits(f, 24) + 1
    r.classifications = get_bits(f, 6) + 1
    r.classbook = get_bits(f, 8)
    If r.classbook >= f.codebook_count Then start_decoder_residue = error_(f, VORBIS_invalid_setup): Exit Function
    For j = 0 To r.classifications - 1
        high_bits = 0
        low_bits = get_bits(f, 3)
        If get_bits(f, 1) <> 0 Then
            high_bits = get_bits(f, 5)
        End If
        residue_cascade(j) = (high_bits * 8 + low_bits) And &HFF&    ' uint8
    Next
    ReDim r.residue_books(0 To r.classifications - 1, 0 To 7)
    setup_malloc f, 16 * r.classifications
    For j = 0 To r.classifications - 1
        For k = 0 To 7
            If (residue_cascade(j) And pow2tab(k)) <> 0 Then
                r.residue_books(j, k) = get_bits(f, 8)
                If r.residue_books(j, k) >= f.codebook_count Then start_decoder_residue = error_(f, VORBIS_invalid_setup): Exit Function
            Else
                r.residue_books(j, k) = -1
            End If
        Next
    Next
    ' precompute the classifications[] array to avoid inner-loop mod/divide
    ' call it 'classdata' since we already have r->classifications
    entries = f.codebooks(r.classbook).entries
    classwords = f.codebooks(r.classbook).dimensions
    r.classwords = classwords
    If entries * classwords > 0 Then ReDim r.classdata(0 To entries * classwords - 1) Else ReDim r.classdata(0 To 0)
    setup_malloc f, 4 * entries
    For j = 0 To entries - 1
        temp = j
        setup_malloc f, classwords
        For k = classwords - 1 To 0 Step -1
            r.classdata(j * classwords + k) = temp Mod r.classifications
            temp = temp \ r.classifications
        Next
    Next
    start_decoder_residue = 1
End Function

' (VB) start_decoder's mapping loop body: m is f->mapping + i
Private Function start_decoder_mapping(f As vorb, m As MappingT, max_submaps As Long) As Long
    Dim mapping_type As Long, j As Long, k As Long
    mapping_type = get_bits(f, 16)
    If mapping_type <> 0 Then start_decoder_mapping = error_(f, VORBIS_invalid_setup): Exit Function
    ReDim m.chan(0 To f.channels - 1)
    setup_malloc f, 3 * f.channels
    If get_bits(f, 1) <> 0 Then
        m.submaps = get_bits(f, 4) + 1
    Else
        m.submaps = 1
    End If
    If m.submaps > max_submaps Then max_submaps = m.submaps
    If get_bits(f, 1) <> 0 Then
        m.coupling_steps = get_bits(f, 8) + 1
        If m.coupling_steps > f.channels Then start_decoder_mapping = error_(f, VORBIS_invalid_setup): Exit Function
        For k = 0 To m.coupling_steps - 1
            m.chan(k).magnitude = get_bits(f, ilog(f.channels - 1))
            m.chan(k).angle = get_bits(f, ilog(f.channels - 1))
            If m.chan(k).magnitude >= f.channels Then start_decoder_mapping = error_(f, VORBIS_invalid_setup): Exit Function
            If m.chan(k).angle >= f.channels Then start_decoder_mapping = error_(f, VORBIS_invalid_setup): Exit Function
            If m.chan(k).magnitude = m.chan(k).angle Then start_decoder_mapping = error_(f, VORBIS_invalid_setup): Exit Function
        Next
    Else
        m.coupling_steps = 0
    End If

    ' reserved field
    If get_bits(f, 2) <> 0 Then start_decoder_mapping = error_(f, VORBIS_invalid_setup): Exit Function
    If m.submaps > 1 Then
        For j = 0 To f.channels - 1
            m.chan(j).mux = get_bits(f, 4)
            If m.chan(j).mux >= m.submaps Then start_decoder_mapping = error_(f, VORBIS_invalid_setup): Exit Function
        Next
    Else
        ' @SPECIFICATION: this case is missing from the spec
        For j = 0 To f.channels - 1
            m.chan(j).mux = 0
        Next
    End If

    For j = 0 To m.submaps - 1
        get_bits f, 8 ' discard
        m.submap_floor(j) = get_bits(f, 8)
        m.submap_residue(j) = get_bits(f, 8)
        If m.submap_floor(j) >= f.floor_count Then start_decoder_mapping = error_(f, VORBIS_invalid_setup): Exit Function
        If m.submap_residue(j) >= f.residue_count Then start_decoder_mapping = error_(f, VORBIS_invalid_setup): Exit Function
    Next
    start_decoder_mapping = 1
End Function

' (VB) vorbis_deinit and vorbis_init: the slot back to a blank decoder (VB frees the arrays)
Private Sub vorbis_reset(p As vorb)
    Dim blank As vorb
    p = blank
    p.eof_ = 0
    p.error_ = VORBIS__no_error
    p.page_crc_tests = -1
End Sub

Public Sub stb_vorbis_close(ByVal p As Long)
    If p = 0 Then Exit Sub
    vorbis_reset pool(p - 1)    ' vorbis_deinit(p); setup_free(p,p);
    pool(p - 1).in_use = False
End Sub

Public Function stb_vorbis_get_info(ByVal f As Long) As stb_vorbis_info
    Dim d As stb_vorbis_info
    d.channels = pool(f - 1).channels
    d.sample_rate = pool(f - 1).sample_rate
    d.setup_memory_required = pool(f - 1).setup_memory_required
    d.setup_temp_memory_required = pool(f - 1).setup_temp_memory_required
    d.temp_memory_required = pool(f - 1).temp_memory_required
    d.max_frame_size = pool(f - 1).blocksize_1 \ 2
    stb_vorbis_get_info = d
End Function

Public Function stb_vorbis_get_error(ByVal f As Long) As Long
    Dim e As Long
    e = pool(f - 1).error_
    pool(f - 1).error_ = VORBIS__no_error
    stb_vorbis_get_error = e
End Function

Private Function stb_vorbis_get_file_offset(f As vorb) As Long
    ' (!push_mode; USE_MEMORY)
    stb_vorbis_get_file_offset = f.stream_pos      ' (unsigned int) (f->stream - f->stream_start)
End Function

'
' DATA-PULLING API
'

' end_, last: uint32 *end, uint32 *last (the ported callers pass both)
Private Function vorbis_find_page(f As vorb, end_ As Long, last As Long) As Long
    Dim n As Long, retry_loc As Long, i As Long, header(0 To 26) As Long, crc As Long, goal As Long, len_ As Long, s As Long
    Dim valid As Boolean
    Do
        If f.eof_ <> 0 Then vorbis_find_page = 0: Exit Function
        n = get8(f)
        If n = &H4F Then ' page header candidate
            retry_loc = stb_vorbis_get_file_offset(f)
            ' check if we're off the end of a file_section stream
            If ult32(f.stream_len, sub32(retry_loc, 25)) Then vorbis_find_page = 0: Exit Function
            ' check the rest of the header
            For i = 1 To 3
                If get8(f) <> ogg_page_header(i) Then Exit For
            Next
            If f.eof_ <> 0 Then vorbis_find_page = 0: Exit Function
            If i = 4 Then
                For i = 0 To 3
                    header(i) = ogg_page_header(i)
                Next
                For i = 4 To 26
                    header(i) = get8(f)
                Next
                If f.eof_ <> 0 Then vorbis_find_page = 0: Exit Function
                valid = (header(4) = 0)     ' if (header[4] != 0) goto invalid;
                If valid Then
                    goal = header(22) + header(23) * &H100& + header(24) * &H10000 + ushl(header(25), 24)
                    For i = 22 To 25
                        header(i) = 0
                    Next
                    crc = 0
                    For i = 0 To 26
                        crc = crc32_update(crc, header(i))
                    Next
                    len_ = 0
                    For i = 0 To header(26) - 1
                        s = get8(f)
                        crc = crc32_update(crc, s)
                        len_ = len_ + s
                    Next
                    If len_ <> 0 And f.eof_ <> 0 Then vorbis_find_page = 0: Exit Function
                    For i = 0 To len_ - 1
                        crc = crc32_update(crc, get8(f))
                    Next
                    ' finished parsing probable page
                    If crc = goal Then
                        ' we could now check that it's either got the last
                        ' page flag set, OR it's followed by the capture
                        ' pattern, but I guess TECHNICALLY you could have
                        ' a file with garbage between each ogg page and recover
                        ' from it automatically? So even though that paranoia
                        ' might decrease the chance of an invalid decode by
                        ' another 2^32, not worth it since it would hose those
                        ' invalid-but-useful files?
                        end_ = stb_vorbis_get_file_offset(f)
                        If (header(5) And &H4) <> 0 Then
                            last = 1
                        Else
                            last = 0
                        End If
                        set_file_offset f, retry_loc - 1
                        vorbis_find_page = 1
                        Exit Function
                    End If
                End If
            End If
            ' invalid:
            ' not a valid page, so rewind and look for next one
            set_file_offset f, retry_loc
        End If
    Loop
End Function

' seeking is implemented with a binary search, which narrows down the range to
' 64K, before using a linear search (because finding the synchronization
' pattern can be expensive, and the chance we'd find the end page again is
' relatively high for small ranges)
' (VB: stb_vorbis_seek_frame, stb_vorbis_seek and their helpers are not ported)

Public Function stb_vorbis_seek_start(ByVal f As Long) As Long
    stb_vorbis_seek_start = vorbis_seek_start(pool(f - 1))
End Function

Private Function vorbis_seek_start(f As vorb) As Long
    ' (IS_PUSH_MODE(f) is FALSE)
    set_file_offset f, f.first_audio_page_offset
    f.previous_length = 0
    f.first_decode = 1
    f.next_seg = -1
    vorbis_seek_start = vorbis_pump_first_frame(f)
End Function

Public Function stb_vorbis_stream_length_in_samples(ByVal f As Long) As Long
    stb_vorbis_stream_length_in_samples = vorbis_stream_length_in_samples(pool(f - 1))
End Function

Private Function vorbis_stream_length_in_samples(f As vorb) As Long
    Dim restore_offset As Long, previous_safe As Long, end_ As Long, last_page_loc As Long, last As Long
    Dim lo As Long, hi As Long, header(0 To 5) As Byte

    ' (IS_PUSH_MODE(f) is FALSE)
    If f.total_samples = 0 Then
        ' first, store the current decode position so we can restore it
        restore_offset = stb_vorbis_get_file_offset(f)

        ' now we want to seek back 64K from the end (the last page must
        ' be at most a little less than 64K, but let's allow a little slop)
        If f.stream_len >= 65536 And f.stream_len - 65536 >= f.first_audio_page_offset Then
            previous_safe = f.stream_len - 65536
        Else
            previous_safe = f.first_audio_page_offset
        End If

        set_file_offset f, previous_safe
        ' previous_safe is now our candidate 'earliest known place that seeking
        ' to will lead to the final page'

        If vorbis_find_page(f, end_, last) = 0 Then
            ' if we can't find a page, we're hosed!
            f.error_ = VORBIS_cant_find_last_page
            f.total_samples = &HFFFFFFFF
            GoTo done
        End If

        ' check if there are more pages
        last_page_loc = stb_vorbis_get_file_offset(f)

        ' stop when the last_page flag is set, not when we reach eof;
        ' this allows us to stop short of a 'file_section' end without
        ' explicitly checking the length of the section
        Do While last = 0
            set_file_offset f, end_
            If vorbis_find_page(f, end_, last) = 0 Then
                ' the last page we found didn't have the 'last page' flag
                ' set. whoops!
                Exit Do
            End If
            'previous_safe = last_page_loc+1; // NOTE: not used after this point, but note for debugging
            last_page_loc = stb_vorbis_get_file_offset(f)
        Loop

        set_file_offset f, last_page_loc

        ' parse the header
        getn f, header, 6
        ' extract the absolute granule position
        lo = get32(f)
        hi = get32(f)
        If lo = &HFFFFFFFF And hi = &HFFFFFFFF Then
            f.error_ = VORBIS_cant_find_last_page
            f.total_samples = SAMPLE_unknown
            GoTo done
        End If
        If hi <> 0 Then lo = &HFFFFFFFE ' saturate
        f.total_samples = lo

        f.p_last.page_start = last_page_loc
        f.p_last.page_end = end_
        f.p_last.last_decoded_sample = lo

done:
        set_file_offset f, restore_offset
    End If
    If f.total_samples = SAMPLE_unknown Then vorbis_stream_length_in_samples = 0 Else vorbis_stream_length_in_samples = f.total_samples
End Function

Public Function stb_vorbis_stream_length_in_seconds(ByVal f As Long) As Single
    Dim n As Long
    n = stb_vorbis_stream_length_in_samples(f)
    stb_vorbis_stream_length_in_seconds = CSng(CSng(IIf(n < 0, n + 4294967296#, n)) / CSng(IIf(pool(f - 1).sample_rate < 0, pool(f - 1).sample_rate + 4294967296#, pool(f - 1).sample_rate)))
End Function

' (VB) channels, output: the C's int *channels (the decoder's) and float ***output (f->outputs: the
' channel buffers from channel_buffer_start)
Private Function vorbis_get_frame_float(f As vorb) As Long
    Dim len_ As Long, right_ As Long, left_ As Long
    ' (IS_PUSH_MODE(f) is FALSE)

    If vorbis_decode_packet(f, len_, left_, right_) = 0 Then
        f.channel_buffer_start = 0: f.channel_buffer_end = 0
        vorbis_get_frame_float = 0
        Exit Function
    End If

    len_ = vorbis_finish_frame(f, len_, left_, right_)
    ' for (i=0; i < f->channels; ++i) f->outputs[i] = f->channel_buffers[i] + left;

    f.channel_buffer_start = left_
    f.channel_buffer_end = left_ + len_

    vorbis_get_frame_float = len_
End Function

Public Function stb_vorbis_open_memory(data() As Byte, ByVal len_ As Long, error_code As Long) As Long
    Dim h As Long, i As Long
    If Not tablesReady Then stb_tables
    ' if (!data) { if (error) *error = VORBIS_unexpected_eof; return NULL; }
    ' (VB) a free slot of the pool (vorbis_alloc)
    h = -1
    For i = 0 To poolCount - 1
        If Not pool(i).in_use Then h = i: Exit For
    Next
    If h < 0 Then
        h = poolCount
        poolCount = poolCount * 2 + 4
        ReDim Preserve pool(0 To poolCount - 1)
    End If
    vorbis_reset pool(h)                ' vorbis_init(&p, alloc)
    pool(h).in_use = True
    If len_ > 0 Then
        ReDim pool(h).stream(0 To len_ - 1)
        CopyMemory pool(h).stream(0), data(LBound(data)), len_
    Else
        ReDim pool(h).stream(0 To 0)
    End If
    pool(h).stream_pos = 0
    pool(h).stream_len = len_
    If start_decoder(pool(h)) <> 0 Then
        vorbis_pump_first_frame pool(h)
        error_code = VORBIS__no_error
        stb_vorbis_open_memory = h + 1
        Exit Function
    End If
    error_code = pool(h).error_
    vorbis_reset pool(h)                ' vorbis_deinit(&p)
    pool(h).in_use = False
    stb_vorbis_open_memory = 0
End Function

' (STB_VORBIS_NO_INTEGER_CONVERSION is not defined)
' (VB) FAST_SCALED_FLOAT_TO_INT(temp, x, 15) and the clamp to 16 bits: the magic number rounds x * 32768
' to the nearest integer, ties to even, as CLng does
Private Function float_to_short(ByVal x As Single) As Integer
    Dim t As Single
    t = x * 32768!
    If t >= 32767! Then
        float_to_short = 32767
    ElseIf t <= -32768! Then
        float_to_short = -32768
    Else
        float_to_short = CLng(t)
    End If
End Function

' dest[0 .. len-1] at dest + d, src from s
Private Sub copy_samples(dest() As Integer, ByVal d As Long, src() As Single, ByVal s As Long, ByVal len_ As Long)
    Dim i As Long, t As Single
    For i = 0 To len_ - 1
        t = src(s + i) * 32768!
        If t >= 32767! Then
            dest(d + i) = 32767
        ElseIf t <= -32768! Then
            dest(d + i) = -32768
        Else
            dest(d + i) = CLng(t)
        End If
    Next
End Sub

' data: the decoder's channel buffers
Private Sub compute_samples(f As vorb, ByVal mask As Long, output_() As Integer, ByVal o_ As Long, ByVal num_c As Long, ByVal d_offset As Long, ByVal len_ As Long)
    Dim buffer(0 To 31) As Single, i As Long, j As Long, o As Long, n As Long
    n = 32  ' STB_BUFFER_SIZE
    For o = 0 To len_ - 1 Step 32
        For i = 0 To 31
            buffer(i) = 0!
        Next
        If o + n > len_ Then n = len_ - o
        For j = 0 To num_c - 1
            If (channel_position(num_c, j) And mask) <> 0 Then
                For i = 0 To n - 1
                    buffer(i) = buffer(i) + f.channel_buffers(j).v(d_offset + o + i)
                Next
            End If
        Next
        For i = 0 To n - 1
            output_(o_ + o + i) = float_to_short(buffer(i))
        Next
    Next
End Sub

Private Sub compute_stereo_samples(f As vorb, output_() As Integer, ByVal o_ As Long, ByVal num_c As Long, ByVal d_offset As Long, ByVal len_ As Long)
    Dim buffer(0 To 31) As Single, i As Long, j As Long, o As Long, n As Long, o2 As Long, m As Long
    n = 32 \ 2
    ' o is the offset in the source data
    For o = 0 To len_ - 1 Step 16
        ' o2 is the offset in the output data
        o2 = o * 2
        For i = 0 To 31
            buffer(i) = 0!
        Next
        If o + n > len_ Then n = len_ - o
        For j = 0 To num_c - 1
            m = channel_position(num_c, j) And (PLAYBACK_LEFT Or PLAYBACK_RIGHT)
            If m = (PLAYBACK_LEFT Or PLAYBACK_RIGHT) Then
                For i = 0 To n - 1
                    buffer(i * 2 + 0) = buffer(i * 2 + 0) + f.channel_buffers(j).v(d_offset + o + i)
                    buffer(i * 2 + 1) = buffer(i * 2 + 1) + f.channel_buffers(j).v(d_offset + o + i)
                Next
            ElseIf m = PLAYBACK_LEFT Then
                For i = 0 To n - 1
                    buffer(i * 2 + 0) = buffer(i * 2 + 0) + f.channel_buffers(j).v(d_offset + o + i)
                Next
            ElseIf m = PLAYBACK_RIGHT Then
                For i = 0 To n - 1
                    buffer(i * 2 + 1) = buffer(i * 2 + 1) + f.channel_buffers(j).v(d_offset + o + i)
                Next
            End If
        Next
        For i = 0 To n * 2 - 1
            output_(o_ + o2 + i) = float_to_short(buffer(i))
        Next
    Next
End Sub

' (VB) short **buffer with buf_c channels: only buffer[0] (the ported callers ask for one channel); b_offset
' into it
Private Sub convert_samples_short(f As vorb, ByVal buf_c As Long, buffer() As Integer, ByVal b_offset As Long, ByVal data_c As Long, ByVal d_offset As Long, ByVal samples As Long)
    Dim i As Long, limit As Long
    If buf_c <> data_c And buf_c <= 2 And data_c <= 6 Then
        ' static int channel_selector[3][2] = { {0}, {PLAYBACK_MONO}, {PLAYBACK_LEFT, PLAYBACK_RIGHT} };
        For i = 0 To buf_c - 1
            If buf_c = 1 Then compute_samples f, PLAYBACK_MONO, buffer, b_offset, data_c, d_offset, samples
        Next
    Else
        If buf_c < data_c Then limit = buf_c Else limit = data_c
        For i = 0 To limit - 1
            If i = 0 Then copy_samples buffer, b_offset, f.channel_buffers(i).v, d_offset, samples
        Next
        For i = limit To buf_c - 1
            ' memset(buffer[i]+b_offset, 0, sizeof(short) * samples);
        Next
    End If
End Sub

Private Function vorbis_get_frame_short(f As vorb, ByVal num_c As Long, buffer() As Integer, ByVal b As Long, ByVal num_samples As Long) As Long
    Dim len_ As Long
    len_ = vorbis_get_frame_float(f)    ' (output = f->outputs: the buffers from channel_buffer_start)
    If len_ > num_samples Then len_ = num_samples
    If len_ <> 0 Then
        convert_samples_short f, num_c, buffer, b, f.channels, f.channel_buffer_start, len_
    End If
    vorbis_get_frame_short = len_
End Function

' buffer from index b
Private Sub convert_channels_short_interleaved(f As vorb, ByVal buf_c As Long, buffer() As Integer, ByVal b As Long, ByVal data_c As Long, ByVal d_offset As Long, ByVal len_ As Long)
    Dim i As Long, limit As Long, j As Long
    If buf_c <> data_c And buf_c <= 2 And data_c <= 6 Then
        ' assert(buf_c == 2);
        For i = 0 To buf_c - 1
            compute_stereo_samples f, buffer, b, data_c, d_offset, len_
        Next
    Else
        If buf_c < data_c Then limit = buf_c Else limit = data_c
        If limit = 1 And buf_c = 1 Then
            ' (VB) one channel: the loop below, written as a copy
            copy_samples buffer, b, f.channel_buffers(0).v, d_offset, len_
            Exit Sub
        End If
        For j = 0 To len_ - 1
            For i = 0 To limit - 1
                buffer(b) = float_to_short(f.channel_buffers(i).v(d_offset + j))
                b = b + 1
            Next
            For i = limit To buf_c - 1
                buffer(b) = 0
                b = b + 1
            Next
        Next
    End If
End Sub

Private Function vorbis_get_frame_short_interleaved(f As vorb, ByVal num_c As Long, buffer() As Integer, ByVal b As Long, ByVal num_shorts As Long) As Long
    Dim len_ As Long
    If num_c = 1 Then vorbis_get_frame_short_interleaved = vorbis_get_frame_short(f, num_c, buffer, b, num_shorts): Exit Function
    len_ = vorbis_get_frame_float(f)
    If len_ <> 0 Then
        If len_ * num_c > num_shorts Then len_ = num_shorts \ num_c
        convert_channels_short_interleaved f, num_c, buffer, b, f.channels, f.channel_buffer_start, len_
    End If
    vorbis_get_frame_short_interleaved = len_
End Function

' (VB) buffer + offset: the C's short *buffer
Public Function stb_vorbis_get_samples_short_interleaved(ByVal f As Long, ByVal channels As Long, buffer() As Integer, ByVal offset As Long, ByVal num_shorts As Long) As Long
    stb_vorbis_get_samples_short_interleaved = vorbis_get_samples_short_interleaved(pool(f - 1), channels, buffer, offset, num_shorts)
End Function

Private Function vorbis_get_samples_short_interleaved(f As vorb, ByVal channels As Long, buffer() As Integer, ByVal offset As Long, ByVal num_shorts As Long) As Long
    Dim len_ As Long, n As Long, k As Long
    len_ = num_shorts \ channels
    n = 0
    Do While n < len_
        k = f.channel_buffer_end - f.channel_buffer_start
        If n + k >= len_ Then k = len_ - n
        If k <> 0 Then
            convert_channels_short_interleaved f, channels, buffer, offset, f.channels, f.channel_buffer_start, k
        End If
        offset = offset + k * channels
        n = n + k
        f.channel_buffer_start = f.channel_buffer_start + k
        If n = len_ Then Exit Do
        If vorbis_get_frame_float(f) = 0 Then Exit Do
    Loop
    vorbis_get_samples_short_interleaved = n
End Function

' (VB) output: the decoded samples, interleaved, in output(0 .. return * channels - 1) (the C's malloc'd buffer)
Public Function stb_vorbis_decode_memory(data() As Byte, ByVal len_ As Long, channels As Long, sample_rate As Long, output_() As Integer) As Long
    Dim data_len As Long, offset As Long, total As Long, limit As Long, error_code As Long, v As Long, n As Long
    v = stb_vorbis_open_memory(data, len_, error_code)
    If v = 0 Then stb_vorbis_decode_memory = -1: Exit Function
    limit = pool(v - 1).channels * 4096
    channels = pool(v - 1).channels
    sample_rate = pool(v - 1).sample_rate
    offset = 0: data_len = 0
    total = limit
    ReDim output_(0 To total - 1)
    Do
        n = vorbis_get_frame_short_interleaved(pool(v - 1), pool(v - 1).channels, output_, offset, total - offset)
        If n = 0 Then Exit Do
        data_len = data_len + n
        offset = offset + n * pool(v - 1).channels
        If offset + limit > total Then
            total = total * 2
            ReDim Preserve output_(0 To total - 1)
        End If
    Loop
    stb_vorbis_close v
    stb_vorbis_decode_memory = data_len
End Function
