Attribute VB_Name = "miniz_h"
' miniz.c v1.11's header part (Rich Geldreich's public domain deflate/inflate, May 2011; ref/miniz_v111, the
' release whose inflate-only subset tinfl.c v1.11 is PhysicsFS 3.2.0's physfs_miniz.h): the definitions the
' other modules of this folder share. miniz_tinfl_c is the decompressor (PhysFS's copy, moved here from
' src/framework/physfs_miniz_c), miniz_tdefl_c the compressor and its zlib-style stream functions, miniz_c the
' rest of the zlib-style API and zlib's names. Compiled into WzFramework (PhysFS's zip archives) and WzNet (the
' network streams' zlib compression: 4.7.0 links zlib, whose compressed bytes may differ; any inflate reads
' either). The build is the x86 32-bit one: MINIZ_USE_UNALIGNED_LOADS_AND_STORES and MINIZ_LITTLE_ENDIAN 1,
' MINIZ_HAS_64BIT_REGISTERS not set, TDEFL_LESS_MEMORY 0. (VB) A z_stream's buffers are the caller's arrays,
' passed on each call; next_in / next_out are offsets into them; its state is a handle to an entry of the
' modules' pools (0 = NULL); zalloc / zfree / opaque are not kept (the pools are the allocator).
Option Explicit

' ------------------- zlib-style API Definitions.

Public Const MZ_ADLER32_INIT As Long = 1

Public Const MZ_DEFAULT_STRATEGY As Long = 0
Public Const MZ_FILTERED As Long = 1
Public Const MZ_HUFFMAN_ONLY As Long = 2
Public Const MZ_RLE As Long = 3
Public Const MZ_FIXED As Long = 4

' Method
Public Const MZ_DEFLATED As Long = 8

' (#define MZ_VERSION "9.1.11": mz_version() returns it; VB names ignore case, so the two are one)
Public Const MZ_VERNUM As Long = &H91B0&

' Flush values. For typical usage you only need MZ_NO_FLUSH and MZ_FINISH. The other stuff is for advanced use.
Public Const MZ_NO_FLUSH As Long = 0
Public Const MZ_PARTIAL_FLUSH As Long = 1
Public Const MZ_SYNC_FLUSH As Long = 2
Public Const MZ_FULL_FLUSH As Long = 3
Public Const MZ_FINISH As Long = 4
Public Const MZ_BLOCK As Long = 5

' Return status codes. MZ_PARAM_ERROR is non-standard.
Public Const MZ_OK As Long = 0
Public Const MZ_STREAM_END As Long = 1
Public Const MZ_NEED_DICT As Long = 2
Public Const MZ_ERRNO As Long = -1
Public Const MZ_STREAM_ERROR As Long = -2
Public Const MZ_DATA_ERROR As Long = -3
Public Const MZ_MEM_ERROR As Long = -4
Public Const MZ_BUF_ERROR As Long = -5
Public Const MZ_VERSION_ERROR As Long = -6
Public Const MZ_PARAM_ERROR As Long = -10000

' Compression levels.
Public Const MZ_NO_COMPRESSION As Long = 0
Public Const MZ_BEST_SPEED As Long = 1
Public Const MZ_BEST_COMPRESSION As Long = 9
Public Const MZ_DEFAULT_COMPRESSION As Long = -1

' Window bits
Public Const MZ_DEFAULT_WINDOW_BITS As Long = 15

' (VB) a z_stream's state handle: the decompressors' pool from 1, the compressors' from MZ_DEFLATE_STATE_BASE + 1
Public Const MZ_DEFLATE_STATE_BASE As Long = &H100000

' Compression/decompression stream struct. (mz_stream; zlib's name, as MINIZ_NO_ZLIB_COMPATIBLE_NAMES is not set)
Public Type z_stream
    next_in As Long                                             ' pointer to next byte to read: (VB) an offset into the input array
    avail_in As Long                                            ' number of bytes available at next_in
    total_in As Double                                          ' total number of bytes consumed so far (mz_ulong)
    next_out As Long                                            ' pointer to next byte to write: (VB) an offset into the output array
    avail_out As Long                                           ' number of bytes that can be written to next_out
    total_out As Double                                         ' total number of bytes produced so far (mz_ulong)
    msg As String                                               ' error msg (unused)
    state As Long                                               ' internal state, allocated by zalloc/zfree: (VB) a handle, 0 = NULL
    data_type As Long                                           ' data_type (unused)
    adler As Long                                               ' adler32 of the source or uncompressed data (its bits)
    reserved As Long                                            ' not used
End Type

' zlib-compatible names (MINIZ_NO_ZLIB_COMPATIBLE_NAMES is not set): the constants; the functions are miniz_c's
Public Const Z_NULL As Long = 0
Public Const Z_NO_FLUSH As Long = MZ_NO_FLUSH
Public Const Z_PARTIAL_FLUSH As Long = MZ_PARTIAL_FLUSH
Public Const Z_SYNC_FLUSH As Long = MZ_SYNC_FLUSH
Public Const Z_FULL_FLUSH As Long = MZ_FULL_FLUSH
Public Const Z_FINISH As Long = MZ_FINISH
Public Const Z_BLOCK As Long = MZ_BLOCK
Public Const Z_OK As Long = MZ_OK
Public Const Z_STREAM_END As Long = MZ_STREAM_END
Public Const Z_NEED_DICT As Long = MZ_NEED_DICT
Public Const Z_ERRNO As Long = MZ_ERRNO
Public Const Z_STREAM_ERROR As Long = MZ_STREAM_ERROR
Public Const Z_DATA_ERROR As Long = MZ_DATA_ERROR
Public Const Z_MEM_ERROR As Long = MZ_MEM_ERROR
Public Const Z_BUF_ERROR As Long = MZ_BUF_ERROR
Public Const Z_VERSION_ERROR As Long = MZ_VERSION_ERROR
Public Const Z_PARAM_ERROR As Long = MZ_PARAM_ERROR
Public Const Z_NO_COMPRESSION As Long = MZ_NO_COMPRESSION
Public Const Z_BEST_SPEED As Long = MZ_BEST_SPEED
Public Const Z_BEST_COMPRESSION As Long = MZ_BEST_COMPRESSION
Public Const Z_DEFAULT_COMPRESSION As Long = MZ_DEFAULT_COMPRESSION
Public Const Z_DEFAULT_STRATEGY As Long = MZ_DEFAULT_STRATEGY
Public Const Z_FILTERED As Long = MZ_FILTERED
Public Const Z_HUFFMAN_ONLY As Long = MZ_HUFFMAN_ONLY
Public Const Z_RLE As Long = MZ_RLE
Public Const Z_FIXED As Long = MZ_FIXED
Public Const Z_DEFLATED As Long = MZ_DEFLATED
Public Const Z_DEFAULT_WINDOW_BITS As Long = MZ_DEFAULT_WINDOW_BITS
Public Const MAX_WBITS As Long = 15
Public Const MAX_MEM_LEVEL As Long = 9

' ------------------- Low-level Decompression API Definitions

' Decompression flags.
Public Const TINFL_FLAG_PARSE_ZLIB_HEADER As Long = 1
Public Const TINFL_FLAG_HAS_MORE_INPUT As Long = 2
Public Const TINFL_FLAG_USING_NON_WRAPPING_OUTPUT_BUF As Long = 4
Public Const TINFL_FLAG_COMPUTE_ADLER32 As Long = 8

' Max size of LZ dictionary.
Public Const TINFL_LZ_DICT_SIZE As Long = 32768

' Return status. (tinfl_status)
Public Const TINFL_STATUS_BAD_PARAM As Long = -3
Public Const TINFL_STATUS_ADLER32_MISMATCH As Long = -2
Public Const TINFL_STATUS_FAILED As Long = -1
Public Const TINFL_STATUS_DONE As Long = 0
Public Const TINFL_STATUS_NEEDS_MORE_INPUT As Long = 1
Public Const TINFL_STATUS_HAS_MORE_OUTPUT As Long = 2

Public Type tinfl_huff_table
    m_code_size(0 To 287) As Byte                               ' [TINFL_MAX_HUFF_SYMBOLS_0]
    m_look_up(0 To 1023) As Integer                             ' mz_int16 [TINFL_FAST_LOOKUP_SIZE]
    m_tree(0 To 575) As Integer                                 ' mz_int16 [TINFL_MAX_HUFF_SYMBOLS_0 * 2]
End Type

Public Type tinfl_decompressor
    m_state As Long                                             ' mz_uint32 (all of them: their bits)
    m_num_bits As Long
    m_zhdr0 As Long
    m_zhdr1 As Long
    m_z_adler32 As Long
    m_final As Long
    m_type As Long
    m_check_adler32 As Long
    m_dist As Long
    m_counter As Long
    m_num_extra As Long
    m_table_sizes(0 To 2) As Long                               ' [TINFL_MAX_HUFF_TABLES]
    m_bit_buf As Long                                           ' tinfl_bit_buf_t (the 32-bit one)
    m_dist_from_out_buf_start As Long                           ' size_t
    m_tables(0 To 2) As tinfl_huff_table                        ' [TINFL_MAX_HUFF_TABLES]
    m_raw_header(0 To 3) As Byte
    m_len_codes(0 To 456) As Byte                               ' [TINFL_MAX_HUFF_SYMBOLS_0 + TINFL_MAX_HUFF_SYMBOLS_1 + 137]
End Type

' ------------------- Low-level Compression API Definitions

' Set TDEFL_LESS_MEMORY to 1 to use less memory (compression will be slightly slower, and raw/dynamic blocks will be output more frequently).
' (TDEFL_LESS_MEMORY 0)

' Compression flags logically OR'd together (low 12 bits contain the max. number of probes per dictionary search):
' TDEFL_DEFAULT_MAX_PROBES: The compressor defaults to 128 dictionary probes per dictionary search. 0=Huffman only, 1=Huffman+LZ (fastest/crap compression), 4095=Huffman+LZ (slowest/best compression).
Public Const TDEFL_HUFFMAN_ONLY As Long = 0
Public Const TDEFL_DEFAULT_MAX_PROBES As Long = 128
Public Const TDEFL_MAX_PROBES_MASK As Long = &HFFF&

' TDEFL_WRITE_ZLIB_HEADER: If set, the compressor outputs a zlib header before the deflate data, and the Adler-32 of the source data at the end. Otherwise, you'll get raw deflate data.
' TDEFL_COMPUTE_ADLER32: Always compute the adler-32 of the input data (even when not writing zlib headers).
' TDEFL_GREEDY_PARSING_FLAG: Set to use faster greedy parsing, instead of more efficient lazy parsing.
' TDEFL_NONDETERMINISTIC_PARSING_FLAG: Enable to decrease the compressor's initialization time to the minimum, but the output may vary from run to run given the same input (depending on the contents of memory).
' TDEFL_FORCE_ALL_STATIC_BLOCKS: Disable usage of optimized Huffman tables.
Public Const TDEFL_WRITE_ZLIB_HEADER As Long = &H1000&
Public Const TDEFL_COMPUTE_ADLER32 As Long = &H2000&
Public Const TDEFL_GREEDY_PARSING_FLAG As Long = &H4000&
Public Const TDEFL_NONDETERMINISTIC_PARSING_FLAG As Long = &H8000&
Public Const TDEFL_RLE_MATCHES As Long = &H10000
Public Const TDEFL_FILTER_MATCHES As Long = &H20000
Public Const TDEFL_FORCE_ALL_STATIC_BLOCKS As Long = &H40000
Public Const TDEFL_FORCE_ALL_RAW_BLOCKS As Long = &H80000

Public Const TDEFL_MAX_HUFF_TABLES As Long = 3
Public Const TDEFL_MAX_HUFF_SYMBOLS_0 As Long = 288
Public Const TDEFL_MAX_HUFF_SYMBOLS_1 As Long = 32
Public Const TDEFL_MAX_HUFF_SYMBOLS_2 As Long = 19
Public Const TDEFL_LZ_DICT_SIZE As Long = 32768
Public Const TDEFL_LZ_DICT_SIZE_MASK As Long = TDEFL_LZ_DICT_SIZE - 1
Public Const TDEFL_MIN_MATCH_LEN As Long = 3
Public Const TDEFL_MAX_MATCH_LEN As Long = 258

' TDEFL_OUT_BUF_SIZE MUST be large enough to hold a single entire compressed output block (using static/fixed Huffman codes).
Public Const TDEFL_LZ_CODE_BUF_SIZE As Long = 65536                                     ' 64 * 1024
Public Const TDEFL_OUT_BUF_SIZE As Long = 85196                                         ' (TDEFL_LZ_CODE_BUF_SIZE * 13 ) / 10
Public Const TDEFL_MAX_HUFF_SYMBOLS As Long = 288
Public Const TDEFL_LZ_HASH_BITS As Long = 15
Public Const TDEFL_LEVEL1_HASH_SIZE_MASK As Long = 4095
Public Const TDEFL_LZ_HASH_SHIFT As Long = 5                                            ' (TDEFL_LZ_HASH_BITS + 2) / 3
Public Const TDEFL_LZ_HASH_SIZE As Long = 32768                                         ' 1 << TDEFL_LZ_HASH_BITS

' The low-level tdefl functions below may be used directly if the above helper functions aren't flexible enough. The low-level functions don't make any heap allocations, unlike the above helper functions.
' (tdefl_status)
Public Const TDEFL_STATUS_BAD_PARAM As Long = -2
Public Const TDEFL_STATUS_PUT_BUF_FAILED As Long = -1
Public Const TDEFL_STATUS_OKAY As Long = 0
Public Const TDEFL_STATUS_DONE As Long = 1

' Must map to MZ_NO_FLUSH, MZ_SYNC_FLUSH, etc. enums (tdefl_flush)
Public Const TDEFL_NO_FLUSH As Long = 0
Public Const TDEFL_SYNC_FLUSH As Long = 2
Public Const TDEFL_FULL_FLUSH As Long = 3
Public Const TDEFL_FINISH As Long = 4
