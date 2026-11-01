Attribute VB_Name = "miniz_tdefl_c"
' miniz.c v1.11's compressor (ref/miniz_v111/miniz.c, "Low-level Compression (independent from all decompression
' API's)": tdefl, Rich Geldreich's public domain deflate) and its zlib-style stream functions (mz_deflateInit,
' mz_deflateInit2, mz_deflateReset, mz_deflate, mz_deflateEnd, mz_deflateBound, mz_compress2, mz_compress,
' mz_compressBound), as the x86 32-bit build compiles them: MINIZ_USE_UNALIGNED_LOADS_AND_STORES and
' MINIZ_LITTLE_ENDIAN 1 (tdefl_find_match's 16-bit compares, tdefl_compress_fast, TDEFL_READ_UNALIGNED_WORD as two
' bytes), MINIZ_HAS_64BIT_REGISTERS not set (tdefl_compress_lz_codes' plain TDEFL_PUT_BITS version),
' TDEFL_LESS_MEMORY 0.
' (VB) A pointer into a buffer is an offset into an array: the compressor's own buffers are its arrays, the
' caller's input and output are the arrays tdefl_compress is passed (m_pIn_buf, m_pSrc, m_pOut_buf: offsets into
' them; *m_pIn_buf_size and *m_pOut_buf_size are tdefl_compress's ByRef sizes, m_out_buf_size the output's room
' it read). The compressor a z_stream's zalloc allocates is an entry of a pool: the stream's state is
' MZ_DEFLATE_STATE_BASE + its 1-based index. A block always goes through the compressor's own m_output_buf: the
' C writes it straight into the caller's output when that has room for a whole one (TDEFL_OUT_BUF_SIZE), and
' then hands over at once the same bytes the m_output_buf path copies (with that much room the copy is never
' short). The put-buffer callback (tdefl_init's pPut_buf_func, for tdefl_compress_buffer and
' tdefl_compress_mem_to_output / _heap / _mem and the PNG writer) is not ported: the stream API passes NULL.
' MZ_ASSERT (assert: compiled out in a release build) is not ported. The bit buffer never holds more than 23 bits
' (TDEFL_PUT_BITS writes out each full byte), so it stays a non-negative Long; the shifts are products and
' quotients of pow2(). An mz_uint16 is a Long, masked where the C stores one; the mz_uint counters
' (m_lookahead_pos, ...) wrap as the C's do (the projects remove the overflow checks) and are only masked or
' subtracted from each other.
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal dest As Long, ByVal src As Long, ByVal n As Long)

' Radix sorts tdefl_sym_freq[] array by 16-bit key m_key. Returns ptr to sorted values.
Private Type tdefl_sym_freq
    m_key As Long                       ' mz_uint16
    m_sym_index As Long                 ' mz_uint16
End Type

Private Type tdefl_compressor
    ' (tdefl_put_buf_func_ptr m_pPut_buf_func; void *m_pPut_buf_user;: not ported, the stream API's NULL)
    m_flags As Long                     ' mz_uint
    m_max_probes(0 To 1) As Long        ' mz_uint
    m_greedy_parsing As Long
    m_adler32 As Long                   ' mz_uint (its bits)
    m_lookahead_pos As Long             ' mz_uint
    m_lookahead_size As Long
    m_dict_size As Long
    m_pLZ_code_buf As Long              ' mz_uint8*: an offset into m_lz_code_buf
    m_pLZ_flags As Long                 ' mz_uint8*: an offset into m_lz_code_buf
    m_pOutput_buf As Long               ' mz_uint8*: an offset into m_output_buf
    m_pOutput_buf_end As Long           ' mz_uint8*: an offset into m_output_buf
    m_num_flags_left As Long
    m_total_lz_bytes As Long
    m_lz_code_buf_dict_pos As Long
    m_bits_in As Long
    m_bit_buffer As Long
    m_saved_match_dist As Long
    m_saved_match_len As Long
    m_saved_lit As Long
    m_output_flush_ofs As Long
    m_output_flush_remaining As Long
    m_finished As Long
    m_block_index As Long
    m_wants_to_finish As Long
    m_prev_return_status As Long        ' tdefl_status
    m_pIn_buf As Long                   ' const void*: the input's offset in tdefl_compress's inBuf
    m_pOut_buf As Long                  ' void*: the output's offset in tdefl_compress's outBuf
    m_out_buf_size As Long              ' (VB) *m_pOut_buf_size as tdefl_compress was given it: the output's room
    m_flush As Long                     ' tdefl_flush
    m_pSrc As Long                      ' const mz_uint8*: an offset into tdefl_compress's inBuf
    m_src_buf_left As Long              ' size_t
    m_out_buf_ofs As Long               ' size_t
    m_dict() As Byte                    ' mz_uint8 [TDEFL_LZ_DICT_SIZE + TDEFL_MAX_MATCH_LEN - 1]
    m_huff_count() As Long              ' mz_uint16 [TDEFL_MAX_HUFF_TABLES][TDEFL_MAX_HUFF_SYMBOLS]
    m_huff_codes() As Long              ' mz_uint16 [TDEFL_MAX_HUFF_TABLES][TDEFL_MAX_HUFF_SYMBOLS]
    m_huff_code_sizes() As Byte         ' mz_uint8 [TDEFL_MAX_HUFF_TABLES][TDEFL_MAX_HUFF_SYMBOLS]
    m_lz_code_buf() As Byte             ' mz_uint8 [TDEFL_LZ_CODE_BUF_SIZE]
    m_next() As Long                    ' mz_uint16 [TDEFL_LZ_DICT_SIZE]
    m_hash() As Long                    ' mz_uint16 [TDEFL_LZ_HASH_SIZE]
    m_output_buf() As Byte              ' mz_uint8 [TDEFL_OUT_BUF_SIZE]
    inUse As Boolean                    ' (VB) the pool's: allocated
End Type

' Limits canonical Huffman code table's max code size.
Private Const TDEFL_MAX_SUPPORTED_HUFF_CODESIZE As Long = 32

' (VB) the allocator of the compressors (zalloc / zfree)
Private compressors() As tdefl_compressor
Private compressorsCount As Long

' Purposely making these tables static for faster init and thread safety. (VB) and 1 << n
Private s_tdefl_len_sym(0 To 255) As Long               ' mz_uint16
Private s_tdefl_len_extra(0 To 255) As Long             ' mz_uint8
Private s_tdefl_small_dist_sym(0 To 511) As Long        ' mz_uint8
Private s_tdefl_small_dist_extra(0 To 511) As Long      ' mz_uint8
Private s_tdefl_large_dist_sym(0 To 127) As Long        ' mz_uint8
Private s_tdefl_large_dist_extra(0 To 127) As Long      ' mz_uint8
Private s_tdefl_packed_code_size_syms_swizzle(0 To 18) As Long
Private mz_bitmasks(0 To 16) As Long
Private s_tdefl_num_probes(0 To 10) As Long
Private pow2(0 To 31) As Long
Private tablesReady As Boolean

Private Sub fillTable(t() As Long, ByVal first As Long, ByVal values As String)
    Dim v() As String, i As Long
    v = Split(values, ",")
    For i = 0 To UBound(v)
        t(first + i) = CLng(Trim$(v(i)))
    Next
End Sub

Private Sub tdefl_initTables()
    Dim i As Long
    fillTable s_tdefl_len_sym, 0, "257,258,259,260,261,262,263,264,265,265,266,266,267,267,268,268,269,269,269,269,270,270,270,270,271,271,271,271,272,272,272,272,273,273,273,273,273,273,273,273,274,274,274,274,274,274,274,274,275,275,275,275,275,275,275,275,276,276,276,276,276,276,276,276"
    fillTable s_tdefl_len_sym, 64, "277,277,277,277,277,277,277,277,277,277,277,277,277,277,277,277,278,278,278,278,278,278,278,278,278,278,278,278,278,278,278,278,279,279,279,279,279,279,279,279,279,279,279,279,279,279,279,279,280,280,280,280,280,280,280,280,280,280,280,280,280,280,280,280"
    fillTable s_tdefl_len_sym, 128, "281,281,281,281,281,281,281,281,281,281,281,281,281,281,281,281,281,281,281,281,281,281,281,281,281,281,281,281,281,281,281,281,282,282,282,282,282,282,282,282,282,282,282,282,282,282,282,282,282,282,282,282,282,282,282,282,282,282,282,282,282,282,282,282"
    fillTable s_tdefl_len_sym, 192, "283,283,283,283,283,283,283,283,283,283,283,283,283,283,283,283,283,283,283,283,283,283,283,283,283,283,283,283,283,283,283,283,284,284,284,284,284,284,284,284,284,284,284,284,284,284,284,284,284,284,284,284,284,284,284,284,284,284,284,284,284,284,284,285"
    fillTable s_tdefl_len_extra, 0, "0,0,0,0,0,0,0,0,1,1,1,1,1,1,1,1,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3"
    fillTable s_tdefl_len_extra, 64, "4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4"
    fillTable s_tdefl_len_extra, 128, "5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5"
    fillTable s_tdefl_len_extra, 192, "5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,0"
    fillTable s_tdefl_small_dist_sym, 0, "0,1,2,3,4,4,5,5,6,6,6,6,7,7,7,7,8,8,8,8,8,8,8,8,9,9,9,9,9,9,9,9,10,10,10,10,10,10,10,10,10,10,10,10,10,10,10,10,11,11,11,11,11,11,11,11,11,11,11,11,11,11,11,11"
    fillTable s_tdefl_small_dist_sym, 64, "12,12,12,12,12,12,12,12,12,12,12,12,12,12,12,12,12,12,12,12,12,12,12,12,12,12,12,12,12,12,12,12,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13"
    fillTable s_tdefl_small_dist_sym, 128, "14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14"
    fillTable s_tdefl_small_dist_sym, 192, "15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15"
    fillTable s_tdefl_small_dist_sym, 256, "16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16"
    fillTable s_tdefl_small_dist_sym, 320, "16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16"
    fillTable s_tdefl_small_dist_sym, 384, "17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17"
    fillTable s_tdefl_small_dist_sym, 448, "17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17"
    fillTable s_tdefl_small_dist_extra, 0, "0,0,0,0,1,1,1,1,2,2,2,2,2,2,2,2,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4"
    fillTable s_tdefl_small_dist_extra, 64, "5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5,5"
    fillTable s_tdefl_small_dist_extra, 128, "6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6"
    fillTable s_tdefl_small_dist_extra, 192, "6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6,6"
    fillTable s_tdefl_small_dist_extra, 256, "7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7"
    fillTable s_tdefl_small_dist_extra, 320, "7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7"
    fillTable s_tdefl_small_dist_extra, 384, "7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7"
    fillTable s_tdefl_small_dist_extra, 448, "7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7,7"
    fillTable s_tdefl_large_dist_sym, 0, "0,0,18,19,20,20,21,21,22,22,22,22,23,23,23,23,24,24,24,24,24,24,24,24,25,25,25,25,25,25,25,25,26,26,26,26,26,26,26,26,26,26,26,26,26,26,26,26,27,27,27,27,27,27,27,27,27,27,27,27,27,27,27,27"
    fillTable s_tdefl_large_dist_sym, 64, "28,28,28,28,28,28,28,28,28,28,28,28,28,28,28,28,28,28,28,28,28,28,28,28,28,28,28,28,28,28,28,28,29,29,29,29,29,29,29,29,29,29,29,29,29,29,29,29,29,29,29,29,29,29,29,29,29,29,29,29,29,29,29,29"
    fillTable s_tdefl_large_dist_extra, 0, "0,0,8,8,9,9,9,9,10,10,10,10,10,10,10,10,11,11,11,11,11,11,11,11,11,11,11,11,11,11,11,11,12,12,12,12,12,12,12,12,12,12,12,12,12,12,12,12,12,12,12,12,12,12,12,12,12,12,12,12,12,12,12,12"
    fillTable s_tdefl_large_dist_extra, 64, "13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13,13"
    fillTable s_tdefl_packed_code_size_syms_swizzle, 0, "16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15"
    fillTable mz_bitmasks, 0, "0, 1, 3, 7, 15, 31, 63, 127, 255, 511, 1023, 2047, 4095, 8191, 16383, 32767, 65535"
    fillTable s_tdefl_num_probes, 0, "0, 1, 6, 32,  16, 32, 128, 256,  512, 768, 1500"
    pow2(0) = 1
    For i = 1 To 30
        pow2(i) = pow2(i - 1) * 2
    Next
    pow2(31) = &H80000000
    tablesReady = True
End Sub

' (VB) TDEFL_READ_UNALIGNED_WORD(&d->m_dict[i]): *(const mz_uint16*), little endian
Private Function readWord(d As tdefl_compressor, ByVal i As Long) As Long
    readWord = d.m_dict(i) Or (CLng(d.m_dict(i + 1)) * 256)
End Function

' (VB) the result is in pSyms0 (0) or pSyms1 (1)
Private Function tdefl_radix_sort_syms(ByVal num_syms As Long, pSyms0() As tdefl_sym_freq, pSyms1() As tdefl_sym_freq) As Long
    Dim total_passes As Long, pass_shift As Long, pass As Long, i As Long, hist(0 To 511) As Long
    Dim pCur_syms As Long, freq As Long, offsets(0 To 255) As Long, cur_ofs As Long, k As Long
    total_passes = 2: pCur_syms = 0                     ' pCur_syms = pSyms0, *pNew_syms = pSyms1; MZ_CLEAR_OBJ(hist);
    For i = 0 To num_syms - 1
        freq = pSyms0(i).m_key
        hist(freq And &HFF) = hist(freq And &HFF) + 1
        hist(256 + ((freq \ 256) And &HFF)) = hist(256 + ((freq \ 256) And &HFF)) + 1
    Next
    Do While total_passes > 1
        If num_syms <> hist((total_passes - 1) * 256) Then Exit Do
        total_passes = total_passes - 1
    Loop
    pass_shift = 0
    For pass = 0 To total_passes - 1
        cur_ofs = 0                                     ' const mz_uint32* pHist = &hist[pass << 8];
        For i = 0 To 255
            offsets(i) = cur_ofs: cur_ofs = cur_ofs + hist(pass * 256 + i)
        Next
        If pCur_syms = 0 Then
            For i = 0 To num_syms - 1
                k = (pSyms0(i).m_key \ pow2(pass_shift)) And &HFF
                pSyms1(offsets(k)) = pSyms0(i): offsets(k) = offsets(k) + 1
            Next
        Else
            For i = 0 To num_syms - 1
                k = (pSyms1(i).m_key \ pow2(pass_shift)) And &HFF
                pSyms0(offsets(k)) = pSyms1(i): offsets(k) = offsets(k) + 1
            Next
        End If
        pCur_syms = 1 - pCur_syms                       ' { tdefl_sym_freq* t = pCur_syms; pCur_syms = pNew_syms; pNew_syms = t; }
        pass_shift = pass_shift + 8
    Next
    tdefl_radix_sort_syms = pCur_syms
End Function

' tdefl_calculate_minimum_redundancy() originally written by: Alistair Moffat, alistair@cs.mu.oz.au, Jyrki Katajainen, jyrki@diku.dk, November 1996.
Private Sub tdefl_calculate_minimum_redundancy(a() As tdefl_sym_freq, ByVal n As Long)
    Dim root As Long, leaf As Long, next_ As Long, avbl As Long, used As Long, dpth As Long, takeRoot As Boolean
    If n = 0 Then
        Exit Sub
    ElseIf n = 1 Then
        a(0).m_key = 1
        Exit Sub
    End If
    a(0).m_key = (a(0).m_key + a(1).m_key) And &HFFFF&: root = 0: leaf = 2
    For next_ = 1 To n - 2
        ' if (leaf>=n || A[root].m_key<A[leaf].m_key)
        If leaf >= n Then
            takeRoot = True
        Else
            takeRoot = a(root).m_key < a(leaf).m_key
        End If
        If takeRoot Then
            a(next_).m_key = a(root).m_key: a(root).m_key = next_: root = root + 1
        Else
            a(next_).m_key = a(leaf).m_key: leaf = leaf + 1
        End If
        ' if (leaf>=n || (root<next && A[root].m_key<A[leaf].m_key))
        If leaf >= n Then
            takeRoot = True
        ElseIf root < next_ Then
            takeRoot = a(root).m_key < a(leaf).m_key
        Else
            takeRoot = False
        End If
        If takeRoot Then
            a(next_).m_key = (a(next_).m_key + a(root).m_key) And &HFFFF&: a(root).m_key = next_: root = root + 1
        Else
            a(next_).m_key = (a(next_).m_key + a(leaf).m_key) And &HFFFF&: leaf = leaf + 1
        End If
    Next
    a(n - 2).m_key = 0
    For next_ = n - 3 To 0 Step -1
        a(next_).m_key = (a(a(next_).m_key).m_key + 1) And &HFFFF&
    Next
    avbl = 1: used = 0: dpth = 0: root = n - 2: next_ = n - 1
    Do While avbl > 0
        Do While root >= 0
            If a(root).m_key <> dpth Then Exit Do
            used = used + 1: root = root - 1
        Loop
        Do While avbl > used
            a(next_).m_key = dpth: next_ = next_ - 1: avbl = avbl - 1
        Loop
        avbl = 2 * used: dpth = dpth + 1: used = 0
    Loop
End Sub

Private Sub tdefl_huffman_enforce_max_code_size(pNum_codes() As Long, ByVal code_list_len As Long, ByVal max_code_size As Long)
    Dim i As Long, total As Long
    total = 0
    If code_list_len <= 1 Then Exit Sub
    For i = max_code_size + 1 To TDEFL_MAX_SUPPORTED_HUFF_CODESIZE
        pNum_codes(max_code_size) = pNum_codes(max_code_size) + pNum_codes(i)
    Next
    For i = max_code_size To 1 Step -1
        total = total + pNum_codes(i) * pow2(max_code_size - i)
    Next
    Do While total <> pow2(max_code_size)
        pNum_codes(max_code_size) = pNum_codes(max_code_size) - 1
        For i = max_code_size - 1 To 1 Step -1
            If pNum_codes(i) <> 0 Then
                pNum_codes(i) = pNum_codes(i) - 1: pNum_codes(i + 1) = pNum_codes(i + 1) + 2
                Exit For
            End If
        Next
        total = total - 1
    Loop
End Sub

Private Sub tdefl_optimize_huffman_table(d As tdefl_compressor, ByVal table_num As Long, ByVal table_len As Long, ByVal code_size_limit As Long, ByVal static_table As Boolean)
    Dim i As Long, j As Long, l As Long, num_codes(0 To TDEFL_MAX_SUPPORTED_HUFF_CODESIZE) As Long, next_code(0 To TDEFL_MAX_SUPPORTED_HUFF_CODESIZE) As Long
    Dim syms0(0 To TDEFL_MAX_HUFF_SYMBOLS - 1) As tdefl_sym_freq, syms1(0 To TDEFL_MAX_HUFF_SYMBOLS - 1) As tdefl_sym_freq
    Dim num_used_syms As Long, rev_code As Long, code As Long, code_size As Long
    If static_table Then
        For i = 0 To table_len - 1
            num_codes(d.m_huff_code_sizes(table_num, i)) = num_codes(d.m_huff_code_sizes(table_num, i)) + 1
        Next
    Else
        num_used_syms = 0
        For i = 0 To table_len - 1                      ' const mz_uint16 *pSym_count = &d->m_huff_count[table_num][0];
            If d.m_huff_count(table_num, i) <> 0 Then
                syms0(num_used_syms).m_key = d.m_huff_count(table_num, i): syms0(num_used_syms).m_sym_index = i
                num_used_syms = num_used_syms + 1
            End If
        Next

        ' pSyms = tdefl_radix_sort_syms(num_used_syms, syms0, syms1): (VB) pSyms is syms0 once the result is copied there
        If tdefl_radix_sort_syms(num_used_syms, syms0, syms1) = 1 Then
            For i = 0 To num_used_syms - 1
                syms0(i) = syms1(i)
            Next
        End If
        tdefl_calculate_minimum_redundancy syms0, num_used_syms

        For i = 0 To num_used_syms - 1
            num_codes(syms0(i).m_key) = num_codes(syms0(i).m_key) + 1
        Next

        tdefl_huffman_enforce_max_code_size num_codes, num_used_syms, code_size_limit

        For i = 0 To TDEFL_MAX_HUFF_SYMBOLS - 1         ' MZ_CLEAR_OBJ(d->m_huff_code_sizes[table_num]); MZ_CLEAR_OBJ(d->m_huff_codes[table_num]);
            d.m_huff_code_sizes(table_num, i) = 0: d.m_huff_codes(table_num, i) = 0
        Next
        j = num_used_syms
        For i = 1 To code_size_limit
            For l = num_codes(i) To 1 Step -1
                j = j - 1
                d.m_huff_code_sizes(table_num, syms0(j).m_sym_index) = i
            Next
        Next
    End If

    next_code(1) = 0: j = 0
    For i = 2 To code_size_limit
        j = (j + num_codes(i - 1)) * 2: next_code(i) = j
    Next

    For i = 0 To table_len - 1
        rev_code = 0
        code_size = d.m_huff_code_sizes(table_num, i)
        If code_size <> 0 Then
            code = next_code(code_size): next_code(code_size) = next_code(code_size) + 1
            For l = code_size To 1 Step -1
                rev_code = (rev_code * 2) Or (code And 1): code = code \ 2
            Next
            d.m_huff_codes(table_num, i) = rev_code And &HFFFF&
        End If
    Next
End Sub

' #define TDEFL_PUT_BITS(b, l)
Private Sub tdefl_put_bits(d As tdefl_compressor, ByVal bits As Long, ByVal length As Long)
    d.m_bit_buffer = d.m_bit_buffer Or (bits * pow2(d.m_bits_in)): d.m_bits_in = d.m_bits_in + length
    Do While d.m_bits_in >= 8
        If d.m_pOutput_buf < d.m_pOutput_buf_end Then d.m_output_buf(d.m_pOutput_buf) = d.m_bit_buffer And &HFF: d.m_pOutput_buf = d.m_pOutput_buf + 1
        d.m_bit_buffer = d.m_bit_buffer \ 256
        d.m_bits_in = d.m_bits_in - 8
    Loop
End Sub

' #define TDEFL_RLE_PREV_CODE_SIZE()
Private Sub tdefl_rle_prev_code_size(d As tdefl_compressor, packed_code_sizes() As Byte, num_packed_code_sizes As Long, rle_repeat_count As Long, ByVal prev_code_size As Long)
    If rle_repeat_count <> 0 Then
        If rle_repeat_count < 3 Then
            d.m_huff_count(2, prev_code_size) = (d.m_huff_count(2, prev_code_size) + rle_repeat_count) And &HFFFF&
            Do While rle_repeat_count <> 0                  ' while (rle_repeat_count--)
                rle_repeat_count = rle_repeat_count - 1
                packed_code_sizes(num_packed_code_sizes) = prev_code_size: num_packed_code_sizes = num_packed_code_sizes + 1
            Loop
        Else
            d.m_huff_count(2, 16) = (d.m_huff_count(2, 16) + 1) And &HFFFF&
            packed_code_sizes(num_packed_code_sizes) = 16: num_packed_code_sizes = num_packed_code_sizes + 1
            packed_code_sizes(num_packed_code_sizes) = rle_repeat_count - 3: num_packed_code_sizes = num_packed_code_sizes + 1
        End If
        rle_repeat_count = 0
    End If
End Sub

' #define TDEFL_RLE_ZERO_CODE_SIZE()
Private Sub tdefl_rle_zero_code_size(d As tdefl_compressor, packed_code_sizes() As Byte, num_packed_code_sizes As Long, rle_z_count As Long)
    If rle_z_count <> 0 Then
        If rle_z_count < 3 Then
            d.m_huff_count(2, 0) = (d.m_huff_count(2, 0) + rle_z_count) And &HFFFF&
            Do While rle_z_count <> 0                       ' while (rle_z_count--)
                rle_z_count = rle_z_count - 1
                packed_code_sizes(num_packed_code_sizes) = 0: num_packed_code_sizes = num_packed_code_sizes + 1
            Loop
        ElseIf rle_z_count <= 10 Then
            d.m_huff_count(2, 17) = (d.m_huff_count(2, 17) + 1) And &HFFFF&
            packed_code_sizes(num_packed_code_sizes) = 17: num_packed_code_sizes = num_packed_code_sizes + 1
            packed_code_sizes(num_packed_code_sizes) = rle_z_count - 3: num_packed_code_sizes = num_packed_code_sizes + 1
        Else
            d.m_huff_count(2, 18) = (d.m_huff_count(2, 18) + 1) And &HFFFF&
            packed_code_sizes(num_packed_code_sizes) = 18: num_packed_code_sizes = num_packed_code_sizes + 1
            packed_code_sizes(num_packed_code_sizes) = rle_z_count - 11: num_packed_code_sizes = num_packed_code_sizes + 1
        End If
        rle_z_count = 0
    End If
End Sub

Private Sub tdefl_start_dynamic_block(d As tdefl_compressor)
    Dim num_lit_codes As Long, num_dist_codes As Long, num_bit_lengths As Long, i As Long, total_code_sizes_to_pack As Long
    Dim num_packed_code_sizes As Long, rle_z_count As Long, rle_repeat_count As Long, packed_code_sizes_index As Long
    Dim code_sizes_to_pack(0 To TDEFL_MAX_HUFF_SYMBOLS_0 + TDEFL_MAX_HUFF_SYMBOLS_1 - 1) As Byte
    Dim packed_code_sizes(0 To TDEFL_MAX_HUFF_SYMBOLS_0 + TDEFL_MAX_HUFF_SYMBOLS_1 - 1) As Byte
    Dim prev_code_size As Long, code_size As Long, code As Long, extraBits As Long
    prev_code_size = &HFF

    d.m_huff_count(0, 256) = 1

    tdefl_optimize_huffman_table d, 0, TDEFL_MAX_HUFF_SYMBOLS_0, 15, False
    tdefl_optimize_huffman_table d, 1, TDEFL_MAX_HUFF_SYMBOLS_1, 15, False

    For num_lit_codes = 286 To 258 Step -1
        If d.m_huff_code_sizes(0, num_lit_codes - 1) <> 0 Then Exit For
    Next
    For num_dist_codes = 30 To 2 Step -1
        If d.m_huff_code_sizes(1, num_dist_codes - 1) <> 0 Then Exit For
    Next

    For i = 0 To num_lit_codes - 1                      ' memcpy(code_sizes_to_pack, &d->m_huff_code_sizes[0][0], num_lit_codes);
        code_sizes_to_pack(i) = d.m_huff_code_sizes(0, i)
    Next
    For i = 0 To num_dist_codes - 1                     ' memcpy(code_sizes_to_pack + num_lit_codes, &d->m_huff_code_sizes[1][0], num_dist_codes);
        code_sizes_to_pack(num_lit_codes + i) = d.m_huff_code_sizes(1, i)
    Next
    total_code_sizes_to_pack = num_lit_codes + num_dist_codes: num_packed_code_sizes = 0: rle_z_count = 0: rle_repeat_count = 0

    For i = 0 To TDEFL_MAX_HUFF_SYMBOLS_2 - 1           ' memset(&d->m_huff_count[2][0], 0, sizeof(d->m_huff_count[2][0]) * TDEFL_MAX_HUFF_SYMBOLS_2);
        d.m_huff_count(2, i) = 0
    Next
    For i = 0 To total_code_sizes_to_pack - 1
        code_size = code_sizes_to_pack(i)
        If code_size = 0 Then
            tdefl_rle_prev_code_size d, packed_code_sizes, num_packed_code_sizes, rle_repeat_count, prev_code_size
            rle_z_count = rle_z_count + 1
            If rle_z_count = 138 Then tdefl_rle_zero_code_size d, packed_code_sizes, num_packed_code_sizes, rle_z_count
        Else
            tdefl_rle_zero_code_size d, packed_code_sizes, num_packed_code_sizes, rle_z_count
            If code_size <> prev_code_size Then
                tdefl_rle_prev_code_size d, packed_code_sizes, num_packed_code_sizes, rle_repeat_count, prev_code_size
                d.m_huff_count(2, code_size) = (d.m_huff_count(2, code_size) + 1) And &HFFFF&
                packed_code_sizes(num_packed_code_sizes) = code_size: num_packed_code_sizes = num_packed_code_sizes + 1
            Else
                rle_repeat_count = rle_repeat_count + 1
                If rle_repeat_count = 6 Then tdefl_rle_prev_code_size d, packed_code_sizes, num_packed_code_sizes, rle_repeat_count, prev_code_size
            End If
        End If
        prev_code_size = code_size
    Next
    If rle_repeat_count <> 0 Then
        tdefl_rle_prev_code_size d, packed_code_sizes, num_packed_code_sizes, rle_repeat_count, prev_code_size
    Else
        tdefl_rle_zero_code_size d, packed_code_sizes, num_packed_code_sizes, rle_z_count
    End If

    tdefl_optimize_huffman_table d, 2, TDEFL_MAX_HUFF_SYMBOLS_2, 7, False

    tdefl_put_bits d, 2, 2

    tdefl_put_bits d, num_lit_codes - 257, 5
    tdefl_put_bits d, num_dist_codes - 1, 5

    For num_bit_lengths = 18 To 0 Step -1
        If d.m_huff_code_sizes(2, s_tdefl_packed_code_size_syms_swizzle(num_bit_lengths)) <> 0 Then Exit For
    Next
    num_bit_lengths = num_bit_lengths + 1               ' MZ_MAX(4, (num_bit_lengths + 1))
    If num_bit_lengths < 4 Then num_bit_lengths = 4
    tdefl_put_bits d, num_bit_lengths - 4, 4
    For i = 0 To num_bit_lengths - 1
        tdefl_put_bits d, d.m_huff_code_sizes(2, s_tdefl_packed_code_size_syms_swizzle(i)), 3
    Next

    packed_code_sizes_index = 0
    Do While packed_code_sizes_index < num_packed_code_sizes
        code = packed_code_sizes(packed_code_sizes_index): packed_code_sizes_index = packed_code_sizes_index + 1
        tdefl_put_bits d, d.m_huff_codes(2, code), d.m_huff_code_sizes(2, code)
        If code >= 16 Then
            Select Case code - 16                       ' "\02\03\07"[code - 16]
            Case 0: extraBits = 2
            Case 1: extraBits = 3
            Case Else: extraBits = 7
            End Select
            tdefl_put_bits d, packed_code_sizes(packed_code_sizes_index), extraBits
            packed_code_sizes_index = packed_code_sizes_index + 1
        End If
    Loop
End Sub

Private Sub tdefl_start_static_block(d As tdefl_compressor)
    Dim i As Long
    ' mz_uint8 *p = &d->m_huff_code_sizes[0][0];
    For i = 0 To 143
        d.m_huff_code_sizes(0, i) = 8
    Next
    For i = 144 To 255
        d.m_huff_code_sizes(0, i) = 9
    Next
    For i = 256 To 279
        d.m_huff_code_sizes(0, i) = 7
    Next
    For i = 280 To 287
        d.m_huff_code_sizes(0, i) = 8
    Next

    For i = 0 To 31                                     ' memset(d->m_huff_code_sizes[1], 5, 32);
        d.m_huff_code_sizes(1, i) = 5
    Next

    tdefl_optimize_huffman_table d, 0, 288, 15, True
    tdefl_optimize_huffman_table d, 1, 32, 15, True

    tdefl_put_bits d, 1, 2
End Sub

' (MINIZ_HAS_64BIT_REGISTERS not set: the plain version)
Private Function tdefl_compress_lz_codes(d As tdefl_compressor) As Boolean
    Dim flags As Long, pLZ_codes As Long, sym As Long, num_extra_bits As Long, match_len As Long, match_dist As Long, lit As Long

    flags = 1
    pLZ_codes = 0                                       ' d->m_lz_code_buf
    Do While pLZ_codes < d.m_pLZ_code_buf
        If flags = 1 Then
            flags = d.m_lz_code_buf(pLZ_codes) Or &H100: pLZ_codes = pLZ_codes + 1
        End If
        If (flags And 1) <> 0 Then
            match_len = d.m_lz_code_buf(pLZ_codes): match_dist = d.m_lz_code_buf(pLZ_codes + 1) Or (CLng(d.m_lz_code_buf(pLZ_codes + 2)) * 256): pLZ_codes = pLZ_codes + 3

            tdefl_put_bits d, d.m_huff_codes(0, s_tdefl_len_sym(match_len)), d.m_huff_code_sizes(0, s_tdefl_len_sym(match_len))
            tdefl_put_bits d, match_len And mz_bitmasks(s_tdefl_len_extra(match_len)), s_tdefl_len_extra(match_len)

            If match_dist < 512 Then
                sym = s_tdefl_small_dist_sym(match_dist): num_extra_bits = s_tdefl_small_dist_extra(match_dist)
            Else
                sym = s_tdefl_large_dist_sym(match_dist \ 256): num_extra_bits = s_tdefl_large_dist_extra(match_dist \ 256)
            End If
            tdefl_put_bits d, d.m_huff_codes(1, sym), d.m_huff_code_sizes(1, sym)
            tdefl_put_bits d, match_dist And mz_bitmasks(num_extra_bits), num_extra_bits
        Else
            lit = d.m_lz_code_buf(pLZ_codes): pLZ_codes = pLZ_codes + 1
            tdefl_put_bits d, d.m_huff_codes(0, lit), d.m_huff_code_sizes(0, lit)
        End If
        flags = flags \ 2                               ' the for's flags >>= 1
    Loop

    tdefl_put_bits d, d.m_huff_codes(0, 256), d.m_huff_code_sizes(0, 256)

    tdefl_compress_lz_codes = (d.m_pOutput_buf < d.m_pOutput_buf_end)
End Function

Private Function tdefl_compress_block(d As tdefl_compressor, ByVal static_block As Boolean) As Boolean
    If static_block Then
        tdefl_start_static_block d
    Else
        tdefl_start_dynamic_block d
    End If
    tdefl_compress_block = tdefl_compress_lz_codes(d)
End Function

' (VB) outBuf: tdefl_compress's output array
Private Function tdefl_flush_block(d As tdefl_compressor, outBuf() As Byte, ByVal flush As Long) As Long
    Dim saved_bit_buf As Long, saved_bits_in As Long, pSaved_output_buf As Long, comp_block_succeeded As Boolean
    Dim n As Long, use_raw_block As Boolean, pOutput_buf_start As Long, i As Long, a As Long, z As Long, bytes_to_copy As Long
    comp_block_succeeded = False
    use_raw_block = ((d.m_flags And TDEFL_FORCE_ALL_RAW_BLOCKS) <> 0) And ((d.m_lookahead_pos - d.m_lz_code_buf_dict_pos) <= d.m_dict_size)
    ' (VB) pOutput_buf_start is always d->m_output_buf (see the header)
    pOutput_buf_start = 0

    d.m_pOutput_buf = pOutput_buf_start
    d.m_pOutput_buf_end = d.m_pOutput_buf + TDEFL_OUT_BUF_SIZE - 16

    d.m_output_flush_ofs = 0
    d.m_output_flush_remaining = 0

    d.m_lz_code_buf(d.m_pLZ_flags) = d.m_lz_code_buf(d.m_pLZ_flags) \ pow2(d.m_num_flags_left)
    If d.m_num_flags_left = 8 Then d.m_pLZ_code_buf = d.m_pLZ_code_buf - 1

    If ((d.m_flags And TDEFL_WRITE_ZLIB_HEADER) <> 0) And (d.m_block_index = 0) Then
        tdefl_put_bits d, &H78, 8: tdefl_put_bits d, &H1, 8
    End If

    If flush = TDEFL_FINISH Then tdefl_put_bits d, 1, 1 Else tdefl_put_bits d, 0, 1

    pSaved_output_buf = d.m_pOutput_buf: saved_bit_buf = d.m_bit_buffer: saved_bits_in = d.m_bits_in

    If Not use_raw_block Then
        comp_block_succeeded = tdefl_compress_block(d, ((d.m_flags And TDEFL_FORCE_ALL_STATIC_BLOCKS) <> 0) Or (d.m_total_lz_bytes < 48))
    End If

    ' If the block gets expanded, forget the current contents of the output buffer and send a raw block instead.
    If (use_raw_block Or ((d.m_total_lz_bytes <> 0) And ((d.m_pOutput_buf - pSaved_output_buf + 1) >= d.m_total_lz_bytes))) And _
       ((d.m_lookahead_pos - d.m_lz_code_buf_dict_pos) <= d.m_dict_size) Then
        d.m_pOutput_buf = pSaved_output_buf: d.m_bit_buffer = saved_bit_buf: d.m_bits_in = saved_bits_in
        tdefl_put_bits d, 0, 2
        If d.m_bits_in <> 0 Then tdefl_put_bits d, 0, 8 - d.m_bits_in
        For i = 2 To 1 Step -1                          ' for (i = 2; i; --i, d->m_total_lz_bytes ^= 0xFFFF)
            tdefl_put_bits d, d.m_total_lz_bytes And &HFFFF&, 16
            d.m_total_lz_bytes = d.m_total_lz_bytes Xor &HFFFF&
        Next
        For i = 0 To d.m_total_lz_bytes - 1
            tdefl_put_bits d, d.m_dict((d.m_lz_code_buf_dict_pos + i) And TDEFL_LZ_DICT_SIZE_MASK), 8
        Next
    ' Check for the extremely unlikely (if not impossible) case of the compressed block not fitting into the output buffer when using dynamic codes.
    ElseIf Not comp_block_succeeded Then
        d.m_pOutput_buf = pSaved_output_buf: d.m_bit_buffer = saved_bit_buf: d.m_bits_in = saved_bits_in
        tdefl_compress_block d, True
    End If

    If flush <> 0 Then
        If flush = TDEFL_FINISH Then
            If d.m_bits_in <> 0 Then tdefl_put_bits d, 0, 8 - d.m_bits_in
            If (d.m_flags And TDEFL_WRITE_ZLIB_HEADER) <> 0 Then
                a = d.m_adler32
                For i = 0 To 3
                    tdefl_put_bits d, SHRU(a, 24) And &HFF, 8: a = SHL(a, 8)
                Next
            End If
        Else
            z = 0: tdefl_put_bits d, 0, 3
            If d.m_bits_in <> 0 Then tdefl_put_bits d, 0, 8 - d.m_bits_in
            For i = 2 To 1 Step -1                      ' for (i = 2; i; --i, z ^= 0xFFFF)
                tdefl_put_bits d, z And &HFFFF&, 16
                z = z Xor &HFFFF&
            Next
        End If
    End If

    For i = 0 To TDEFL_MAX_HUFF_SYMBOLS_0 - 1           ' memset(&d->m_huff_count[0][0], 0, ...)
        d.m_huff_count(0, i) = 0
    Next
    For i = 0 To TDEFL_MAX_HUFF_SYMBOLS_1 - 1           ' memset(&d->m_huff_count[1][0], 0, ...)
        d.m_huff_count(1, i) = 0
    Next

    d.m_pLZ_code_buf = 1: d.m_pLZ_flags = 0: d.m_num_flags_left = 8: d.m_lz_code_buf_dict_pos = d.m_lz_code_buf_dict_pos + d.m_total_lz_bytes: d.m_total_lz_bytes = 0: d.m_block_index = d.m_block_index + 1

    n = d.m_pOutput_buf - pOutput_buf_start
    If n <> 0 Then
        ' (if (d->m_pPut_buf_func): not ported)
        ' else if (pOutput_buf_start == d->m_output_buf): (VB) always
        bytes_to_copy = d.m_out_buf_size - d.m_out_buf_ofs  ' (int)MZ_MIN((size_t)n, (size_t)(*d->m_pOut_buf_size - d->m_out_buf_ofs))
        If n < bytes_to_copy Then bytes_to_copy = n
        If bytes_to_copy > 0 Then CopyMemory VarPtr(outBuf(d.m_pOut_buf + d.m_out_buf_ofs)), VarPtr(d.m_output_buf(0)), bytes_to_copy
        d.m_out_buf_ofs = d.m_out_buf_ofs + bytes_to_copy
        n = n - bytes_to_copy
        If n <> 0 Then
            d.m_output_flush_ofs = bytes_to_copy
            d.m_output_flush_remaining = n
        End If
    End If

    tdefl_flush_block = d.m_output_flush_remaining
End Function

' (MINIZ_USE_UNALIGNED_LOADS_AND_STORES: the 16-bit compares)
Private Sub tdefl_find_match(d As tdefl_compressor, ByVal lookahead_pos As Long, ByVal max_dist As Long, ByVal max_match_len As Long, pMatch_dist As Long, pMatch_len As Long)
    Dim dist As Long, pos As Long, match_len As Long, probe_pos As Long, next_probe_pos As Long, probe_len As Long
    Dim num_probes_left As Long, s As Long, p As Long, q As Long, c01 As Long, s01 As Long, k As Long, found As Boolean
    pos = lookahead_pos And TDEFL_LZ_DICT_SIZE_MASK: match_len = pMatch_len: probe_pos = pos
    If match_len >= 32 Then num_probes_left = d.m_max_probes(1) Else num_probes_left = d.m_max_probes(0)
    s = pos                                             ' const mz_uint16 *s = (const mz_uint16*)(d->m_dict + pos): (VB) byte offsets
    c01 = readWord(d, pos + match_len - 1): s01 = readWord(d, s)
    If max_match_len <= match_len Then Exit Sub
    Do
        Do
            num_probes_left = num_probes_left - 1
            If num_probes_left = 0 Then Exit Sub
            found = False
            For k = 1 To 3                              ' TDEFL_PROBE; TDEFL_PROBE; TDEFL_PROBE;
                next_probe_pos = d.m_next(probe_pos)
                If next_probe_pos = 0 Then Exit Sub
                dist = (lookahead_pos - next_probe_pos) And &HFFFF&
                If dist > max_dist Then Exit Sub
                probe_pos = next_probe_pos And TDEFL_LZ_DICT_SIZE_MASK
                If readWord(d, probe_pos + match_len - 1) = c01 Then
                    found = True
                    Exit For
                End If
            Next
            If found Then Exit Do
        Loop
        If dist = 0 Then Exit Do
        q = probe_pos
        If readWord(d, q) <> s01 Then GoTo continue_
        p = s: probe_len = 32
        Do
            p = p + 2: q = q + 2
            If readWord(d, p) <> readWord(d, q) Then Exit Do
            p = p + 2: q = q + 2
            If readWord(d, p) <> readWord(d, q) Then Exit Do
            p = p + 2: q = q + 2
            If readWord(d, p) <> readWord(d, q) Then Exit Do
            p = p + 2: q = q + 2
            If readWord(d, p) <> readWord(d, q) Then Exit Do
            probe_len = probe_len - 1
            If probe_len = 0 Then Exit Do
        Loop
        If probe_len = 0 Then
            pMatch_dist = dist
            If max_match_len < TDEFL_MAX_MATCH_LEN Then pMatch_len = max_match_len Else pMatch_len = TDEFL_MAX_MATCH_LEN
            Exit Do
        Else
            ' probe_len = ((mz_uint)(p - s) * 2) + (mz_uint)(*(const mz_uint8*)p == *(const mz_uint8*)q): (VB) p - s counts bytes
            probe_len = p - s
            If d.m_dict(p) = d.m_dict(q) Then probe_len = probe_len + 1
            If probe_len > match_len Then
                pMatch_dist = dist
                If max_match_len < probe_len Then match_len = max_match_len Else match_len = probe_len
                pMatch_len = match_len
                If match_len = max_match_len Then Exit Do
                c01 = readWord(d, pos + match_len - 1)
            End If
        End If
continue_:
    Loop
End Sub

' (MINIZ_USE_UNALIGNED_LOADS_AND_STORES && MINIZ_LITTLE_ENDIAN) (VB) inBuf, outBuf: tdefl_compress's arrays
Private Function tdefl_compress_fast(d As tdefl_compressor, inBuf() As Byte, outBuf() As Byte) As Boolean
    ' Faster, minimally featured LZRW1-style match+parse loop with better register utilization. Intended for applications where raw throughput is valued more highly than ratio.
    Const TDEFL_COMP_FAST_LOOKAHEAD_SIZE As Long = 4096
    Dim lookahead_pos As Long, lookahead_size As Long, dict_size As Long, total_lz_bytes As Long, num_flags_left As Long
    Dim pLZ_code_buf As Long, pLZ_flags As Long, cur_pos As Long, dst_pos As Long, num_bytes_to_process As Long, n As Long
    Dim cur_match_dist As Long, cur_match_len As Long, pCur_dict As Long, first_trigram As Long, hash As Long, probe_pos As Long
    Dim p As Long, q As Long, probe_len As Long, s0 As Long, s1 As Long, lit As Long, flushed As Long, isMatch As Boolean
    lookahead_pos = d.m_lookahead_pos: lookahead_size = d.m_lookahead_size: dict_size = d.m_dict_size: total_lz_bytes = d.m_total_lz_bytes: num_flags_left = d.m_num_flags_left
    pLZ_code_buf = d.m_pLZ_code_buf: pLZ_flags = d.m_pLZ_flags
    cur_pos = lookahead_pos And TDEFL_LZ_DICT_SIZE_MASK

    Do While (d.m_src_buf_left <> 0) Or ((d.m_flush <> 0) And (lookahead_size <> 0))
        dst_pos = (lookahead_pos + lookahead_size) And TDEFL_LZ_DICT_SIZE_MASK
        num_bytes_to_process = TDEFL_COMP_FAST_LOOKAHEAD_SIZE - lookahead_size      ' (mz_uint)MZ_MIN(d->m_src_buf_left, ...)
        If d.m_src_buf_left < num_bytes_to_process Then num_bytes_to_process = d.m_src_buf_left
        d.m_src_buf_left = d.m_src_buf_left - num_bytes_to_process
        lookahead_size = lookahead_size + num_bytes_to_process

        Do While num_bytes_to_process <> 0
            n = TDEFL_LZ_DICT_SIZE - dst_pos            ' MZ_MIN(TDEFL_LZ_DICT_SIZE - dst_pos, num_bytes_to_process)
            If num_bytes_to_process < n Then n = num_bytes_to_process
            CopyMemory VarPtr(d.m_dict(dst_pos)), VarPtr(inBuf(d.m_pSrc)), n
            If dst_pos < (TDEFL_MAX_MATCH_LEN - 1) Then
                dictCopyToTail d, inBuf, dst_pos, n
            End If
            d.m_pSrc = d.m_pSrc + n
            dst_pos = (dst_pos + n) And TDEFL_LZ_DICT_SIZE_MASK
            num_bytes_to_process = num_bytes_to_process - n
        Loop

        If TDEFL_LZ_DICT_SIZE - lookahead_size < dict_size Then dict_size = TDEFL_LZ_DICT_SIZE - lookahead_size
        If (d.m_flush = 0) And (lookahead_size < TDEFL_COMP_FAST_LOOKAHEAD_SIZE) Then Exit Do

        Do While lookahead_size >= 4
            cur_match_len = 1
            pCur_dict = cur_pos
            ' (*(const mz_uint32 *)pCur_dict) & 0xFFFFFF
            first_trigram = d.m_dict(pCur_dict) Or (CLng(d.m_dict(pCur_dict + 1)) * 256) Or (CLng(d.m_dict(pCur_dict + 2)) * 65536)
            hash = (first_trigram Xor (first_trigram \ pow2(24 - (TDEFL_LZ_HASH_BITS - 8)))) And TDEFL_LEVEL1_HASH_SIZE_MASK
            probe_pos = d.m_hash(hash)
            d.m_hash(hash) = lookahead_pos And &HFFFF&

            ' if (((cur_match_dist = (mz_uint16)(lookahead_pos - probe_pos)) <= dict_size) &&
            '     ((*(const mz_uint32 *)(d->m_dict + (probe_pos &= TDEFL_LZ_DICT_SIZE_MASK)) & 0xFFFFFF) == first_trigram))
            cur_match_dist = (lookahead_pos - probe_pos) And &HFFFF&
            isMatch = False
            If cur_match_dist <= dict_size Then
                probe_pos = probe_pos And TDEFL_LZ_DICT_SIZE_MASK
                isMatch = ((d.m_dict(probe_pos) Or (CLng(d.m_dict(probe_pos + 1)) * 256) Or (CLng(d.m_dict(probe_pos + 2)) * 65536)) = first_trigram)
            End If
            If isMatch Then
                p = pCur_dict
                q = probe_pos
                probe_len = 32
                Do
                    p = p + 2: q = q + 2
                    If readWord(d, p) <> readWord(d, q) Then Exit Do
                    p = p + 2: q = q + 2
                    If readWord(d, p) <> readWord(d, q) Then Exit Do
                    p = p + 2: q = q + 2
                    If readWord(d, p) <> readWord(d, q) Then Exit Do
                    p = p + 2: q = q + 2
                    If readWord(d, p) <> readWord(d, q) Then Exit Do
                    probe_len = probe_len - 1
                    If probe_len = 0 Then Exit Do
                Loop
                ' cur_match_len = ((mz_uint)(p - (const mz_uint16 *)pCur_dict) * 2) + (mz_uint)(*(const mz_uint8 *)p == *(const mz_uint8 *)q): (VB) bytes
                cur_match_len = p - pCur_dict
                If d.m_dict(p) = d.m_dict(q) Then cur_match_len = cur_match_len + 1
                If probe_len = 0 Then
                    If cur_match_dist <> 0 Then cur_match_len = TDEFL_MAX_MATCH_LEN Else cur_match_len = 0
                End If

                If (cur_match_len < TDEFL_MIN_MATCH_LEN) Or ((cur_match_len = TDEFL_MIN_MATCH_LEN) And (cur_match_dist >= 8& * 1024&)) Then
                    cur_match_len = 1
                    d.m_lz_code_buf(pLZ_code_buf) = first_trigram And &HFF: pLZ_code_buf = pLZ_code_buf + 1
                    d.m_lz_code_buf(pLZ_flags) = d.m_lz_code_buf(pLZ_flags) \ 2
                    d.m_huff_count(0, first_trigram And &HFF) = (d.m_huff_count(0, first_trigram And &HFF) + 1) And &HFFFF&
                Else
                    If lookahead_size < cur_match_len Then cur_match_len = lookahead_size

                    cur_match_dist = cur_match_dist - 1

                    d.m_lz_code_buf(pLZ_code_buf) = cur_match_len - TDEFL_MIN_MATCH_LEN
                    d.m_lz_code_buf(pLZ_code_buf + 1) = cur_match_dist And &HFF          ' *(mz_uint16 *)(&pLZ_code_buf[1]) = (mz_uint16)cur_match_dist;
                    d.m_lz_code_buf(pLZ_code_buf + 2) = (cur_match_dist \ 256) And &HFF
                    pLZ_code_buf = pLZ_code_buf + 3
                    d.m_lz_code_buf(pLZ_flags) = (d.m_lz_code_buf(pLZ_flags) \ 2) Or &H80

                    s0 = s_tdefl_small_dist_sym(cur_match_dist And 511)
                    s1 = s_tdefl_large_dist_sym(cur_match_dist \ 256)
                    If cur_match_dist < 512 Then
                        d.m_huff_count(1, s0) = (d.m_huff_count(1, s0) + 1) And &HFFFF&
                    Else
                        d.m_huff_count(1, s1) = (d.m_huff_count(1, s1) + 1) And &HFFFF&
                    End If

                    d.m_huff_count(0, s_tdefl_len_sym(cur_match_len - TDEFL_MIN_MATCH_LEN)) = (d.m_huff_count(0, s_tdefl_len_sym(cur_match_len - TDEFL_MIN_MATCH_LEN)) + 1) And &HFFFF&
                End If
            Else
                d.m_lz_code_buf(pLZ_code_buf) = first_trigram And &HFF: pLZ_code_buf = pLZ_code_buf + 1
                d.m_lz_code_buf(pLZ_flags) = d.m_lz_code_buf(pLZ_flags) \ 2
                d.m_huff_count(0, first_trigram And &HFF) = (d.m_huff_count(0, first_trigram And &HFF) + 1) And &HFFFF&
            End If

            num_flags_left = num_flags_left - 1
            If num_flags_left = 0 Then
                num_flags_left = 8: pLZ_flags = pLZ_code_buf: pLZ_code_buf = pLZ_code_buf + 1
            End If

            total_lz_bytes = total_lz_bytes + cur_match_len
            lookahead_pos = lookahead_pos + cur_match_len
            dict_size = dict_size + cur_match_len       ' MZ_MIN(dict_size + cur_match_len, TDEFL_LZ_DICT_SIZE)
            If dict_size > TDEFL_LZ_DICT_SIZE Then dict_size = TDEFL_LZ_DICT_SIZE
            cur_pos = (cur_pos + cur_match_len) And TDEFL_LZ_DICT_SIZE_MASK
            lookahead_size = lookahead_size - cur_match_len

            If pLZ_code_buf > TDEFL_LZ_CODE_BUF_SIZE - 8 Then
                d.m_lookahead_pos = lookahead_pos: d.m_lookahead_size = lookahead_size: d.m_dict_size = dict_size
                d.m_total_lz_bytes = total_lz_bytes: d.m_pLZ_code_buf = pLZ_code_buf: d.m_pLZ_flags = pLZ_flags: d.m_num_flags_left = num_flags_left
                flushed = tdefl_flush_block(d, outBuf, 0)
                If flushed <> 0 Then
                    tdefl_compress_fast = (flushed >= 0)
                    Exit Function
                End If
                total_lz_bytes = d.m_total_lz_bytes: pLZ_code_buf = d.m_pLZ_code_buf: pLZ_flags = d.m_pLZ_flags: num_flags_left = d.m_num_flags_left
            End If
        Loop

        Do While lookahead_size <> 0
            lit = d.m_dict(cur_pos)

            total_lz_bytes = total_lz_bytes + 1
            d.m_lz_code_buf(pLZ_code_buf) = lit: pLZ_code_buf = pLZ_code_buf + 1
            d.m_lz_code_buf(pLZ_flags) = d.m_lz_code_buf(pLZ_flags) \ 2
            num_flags_left = num_flags_left - 1
            If num_flags_left = 0 Then
                num_flags_left = 8: pLZ_flags = pLZ_code_buf: pLZ_code_buf = pLZ_code_buf + 1
            End If

            d.m_huff_count(0, lit) = (d.m_huff_count(0, lit) + 1) And &HFFFF&

            lookahead_pos = lookahead_pos + 1
            dict_size = dict_size + 1                   ' MZ_MIN(dict_size + 1, TDEFL_LZ_DICT_SIZE)
            If dict_size > TDEFL_LZ_DICT_SIZE Then dict_size = TDEFL_LZ_DICT_SIZE
            cur_pos = (cur_pos + 1) And TDEFL_LZ_DICT_SIZE_MASK
            lookahead_size = lookahead_size - 1

            If pLZ_code_buf > TDEFL_LZ_CODE_BUF_SIZE - 8 Then
                d.m_lookahead_pos = lookahead_pos: d.m_lookahead_size = lookahead_size: d.m_dict_size = dict_size
                d.m_total_lz_bytes = total_lz_bytes: d.m_pLZ_code_buf = pLZ_code_buf: d.m_pLZ_flags = pLZ_flags: d.m_num_flags_left = num_flags_left
                flushed = tdefl_flush_block(d, outBuf, 0)
                If flushed <> 0 Then
                    tdefl_compress_fast = (flushed >= 0)
                    Exit Function
                End If
                total_lz_bytes = d.m_total_lz_bytes: pLZ_code_buf = d.m_pLZ_code_buf: pLZ_flags = d.m_pLZ_flags: num_flags_left = d.m_num_flags_left
            End If
        Loop
    Loop

    d.m_lookahead_pos = lookahead_pos: d.m_lookahead_size = lookahead_size: d.m_dict_size = dict_size
    d.m_total_lz_bytes = total_lz_bytes: d.m_pLZ_code_buf = pLZ_code_buf: d.m_pLZ_flags = pLZ_flags: d.m_num_flags_left = num_flags_left
    tdefl_compress_fast = True
End Function

' (VB) memcpy(d->m_dict + TDEFL_LZ_DICT_SIZE + dst_pos, d->m_pSrc, MZ_MIN(n, (TDEFL_MAX_MATCH_LEN - 1) - dst_pos)):
' the window's first bytes again after its end, so a match may run past it
Private Sub dictCopyToTail(d As tdefl_compressor, inBuf() As Byte, ByVal dst_pos As Long, ByVal n As Long)
    Dim m As Long
    m = (TDEFL_MAX_MATCH_LEN - 1) - dst_pos
    If n < m Then m = n
    If m > 0 Then CopyMemory VarPtr(d.m_dict(TDEFL_LZ_DICT_SIZE + dst_pos)), VarPtr(inBuf(d.m_pSrc)), m
End Sub

Private Sub tdefl_record_literal(d As tdefl_compressor, ByVal lit As Long)
    d.m_total_lz_bytes = d.m_total_lz_bytes + 1
    d.m_lz_code_buf(d.m_pLZ_code_buf) = lit: d.m_pLZ_code_buf = d.m_pLZ_code_buf + 1
    d.m_lz_code_buf(d.m_pLZ_flags) = d.m_lz_code_buf(d.m_pLZ_flags) \ 2
    d.m_num_flags_left = d.m_num_flags_left - 1
    If d.m_num_flags_left = 0 Then
        d.m_num_flags_left = 8: d.m_pLZ_flags = d.m_pLZ_code_buf: d.m_pLZ_code_buf = d.m_pLZ_code_buf + 1
    End If
    d.m_huff_count(0, lit) = (d.m_huff_count(0, lit) + 1) And &HFFFF&
End Sub

Private Sub tdefl_record_match(d As tdefl_compressor, ByVal match_len As Long, ByVal match_dist As Long)
    Dim s0 As Long, s1 As Long

    d.m_total_lz_bytes = d.m_total_lz_bytes + match_len

    d.m_lz_code_buf(d.m_pLZ_code_buf) = match_len - TDEFL_MIN_MATCH_LEN

    match_dist = match_dist - 1
    d.m_lz_code_buf(d.m_pLZ_code_buf + 1) = match_dist And &HFF
    d.m_lz_code_buf(d.m_pLZ_code_buf + 2) = match_dist \ 256: d.m_pLZ_code_buf = d.m_pLZ_code_buf + 3

    d.m_lz_code_buf(d.m_pLZ_flags) = (d.m_lz_code_buf(d.m_pLZ_flags) \ 2) Or &H80
    d.m_num_flags_left = d.m_num_flags_left - 1
    If d.m_num_flags_left = 0 Then
        d.m_num_flags_left = 8: d.m_pLZ_flags = d.m_pLZ_code_buf: d.m_pLZ_code_buf = d.m_pLZ_code_buf + 1
    End If

    s0 = s_tdefl_small_dist_sym(match_dist And 511): s1 = s_tdefl_large_dist_sym(match_dist \ 256)
    If match_dist < 512 Then
        d.m_huff_count(1, s0) = (d.m_huff_count(1, s0) + 1) And &HFFFF&
    Else
        d.m_huff_count(1, s1) = (d.m_huff_count(1, s1) + 1) And &HFFFF&
    End If

    d.m_huff_count(0, s_tdefl_len_sym(match_len - TDEFL_MIN_MATCH_LEN)) = (d.m_huff_count(0, s_tdefl_len_sym(match_len - TDEFL_MIN_MATCH_LEN)) + 1) And &HFFFF&
End Sub

' (VB) inBuf, outBuf: tdefl_compress's arrays
Private Function tdefl_compress_normal(d As tdefl_compressor, inBuf() As Byte, outBuf() As Byte) As Boolean
    Dim pSrc As Long, src_buf_left As Long, flush As Long
    Dim len_to_move As Long, cur_match_dist As Long, cur_match_len As Long, cur_pos As Long
    Dim dst_pos As Long, ins_pos As Long, hash As Long, num_bytes_to_process As Long, pSrc_end As Long, c As Long, n As Long
    pSrc = d.m_pSrc: src_buf_left = d.m_src_buf_left
    flush = d.m_flush

    Do While (src_buf_left <> 0) Or ((flush <> 0) And (d.m_lookahead_size <> 0))
        ' Update dictionary and hash chains. Keeps the lookahead size equal to TDEFL_MAX_MATCH_LEN.
        If (d.m_lookahead_size + d.m_dict_size) >= (TDEFL_MIN_MATCH_LEN - 1) Then
            dst_pos = (d.m_lookahead_pos + d.m_lookahead_size) And TDEFL_LZ_DICT_SIZE_MASK: ins_pos = d.m_lookahead_pos + d.m_lookahead_size - 2
            hash = (CLng(d.m_dict(ins_pos And TDEFL_LZ_DICT_SIZE_MASK)) * pow2(TDEFL_LZ_HASH_SHIFT)) Xor d.m_dict((ins_pos + 1) And TDEFL_LZ_DICT_SIZE_MASK)
            num_bytes_to_process = TDEFL_MAX_MATCH_LEN - d.m_lookahead_size     ' (mz_uint)MZ_MIN(src_buf_left, ...)
            If src_buf_left < num_bytes_to_process Then num_bytes_to_process = src_buf_left
            pSrc_end = pSrc + num_bytes_to_process
            src_buf_left = src_buf_left - num_bytes_to_process
            d.m_lookahead_size = d.m_lookahead_size + num_bytes_to_process
            Do While pSrc <> pSrc_end
                c = inBuf(pSrc): pSrc = pSrc + 1
                d.m_dict(dst_pos) = c
                If dst_pos < (TDEFL_MAX_MATCH_LEN - 1) Then d.m_dict(TDEFL_LZ_DICT_SIZE + dst_pos) = c
                hash = ((hash * pow2(TDEFL_LZ_HASH_SHIFT)) Xor c) And (TDEFL_LZ_HASH_SIZE - 1)
                d.m_next(ins_pos And TDEFL_LZ_DICT_SIZE_MASK) = d.m_hash(hash): d.m_hash(hash) = ins_pos And &HFFFF&
                dst_pos = (dst_pos + 1) And TDEFL_LZ_DICT_SIZE_MASK: ins_pos = ins_pos + 1
            Loop
        Else
            Do While (src_buf_left <> 0) And (d.m_lookahead_size < TDEFL_MAX_MATCH_LEN)
                c = inBuf(pSrc): pSrc = pSrc + 1
                dst_pos = (d.m_lookahead_pos + d.m_lookahead_size) And TDEFL_LZ_DICT_SIZE_MASK
                src_buf_left = src_buf_left - 1
                d.m_dict(dst_pos) = c
                If dst_pos < (TDEFL_MAX_MATCH_LEN - 1) Then
                    d.m_dict(TDEFL_LZ_DICT_SIZE + dst_pos) = c
                End If
                d.m_lookahead_size = d.m_lookahead_size + 1
                If (d.m_lookahead_size + d.m_dict_size) >= TDEFL_MIN_MATCH_LEN Then
                    ins_pos = d.m_lookahead_pos + (d.m_lookahead_size - 1) - 2
                    hash = ((CLng(d.m_dict(ins_pos And TDEFL_LZ_DICT_SIZE_MASK)) * pow2(TDEFL_LZ_HASH_SHIFT * 2)) Xor (CLng(d.m_dict((ins_pos + 1) And TDEFL_LZ_DICT_SIZE_MASK)) * pow2(TDEFL_LZ_HASH_SHIFT)) Xor c) And (TDEFL_LZ_HASH_SIZE - 1)
                    d.m_next(ins_pos And TDEFL_LZ_DICT_SIZE_MASK) = d.m_hash(hash): d.m_hash(hash) = ins_pos And &HFFFF&
                End If
            Loop
        End If
        If TDEFL_LZ_DICT_SIZE - d.m_lookahead_size < d.m_dict_size Then d.m_dict_size = TDEFL_LZ_DICT_SIZE - d.m_lookahead_size
        If (flush = 0) And (d.m_lookahead_size < TDEFL_MAX_MATCH_LEN) Then Exit Do

        ' Simple lazy/greedy parsing state machine.
        len_to_move = 1: cur_match_dist = 0: cur_pos = d.m_lookahead_pos And TDEFL_LZ_DICT_SIZE_MASK
        If d.m_saved_match_len <> 0 Then cur_match_len = d.m_saved_match_len Else cur_match_len = TDEFL_MIN_MATCH_LEN - 1
        If (d.m_flags And (TDEFL_RLE_MATCHES Or TDEFL_FORCE_ALL_RAW_BLOCKS)) <> 0 Then
            If (d.m_dict_size <> 0) And ((d.m_flags And TDEFL_FORCE_ALL_RAW_BLOCKS) = 0) Then
                c = d.m_dict((cur_pos - 1) And TDEFL_LZ_DICT_SIZE_MASK)
                cur_match_len = 0
                Do While cur_match_len < d.m_lookahead_size
                    If d.m_dict(cur_pos + cur_match_len) <> c Then Exit Do
                    cur_match_len = cur_match_len + 1
                Loop
                If cur_match_len < TDEFL_MIN_MATCH_LEN Then cur_match_len = 0 Else cur_match_dist = 1
            End If
        Else
            tdefl_find_match d, d.m_lookahead_pos, d.m_dict_size, d.m_lookahead_size, cur_match_dist, cur_match_len
        End If
        If ((cur_match_len = TDEFL_MIN_MATCH_LEN) And (cur_match_dist >= 8& * 1024&)) Or (cur_pos = cur_match_dist) Or (((d.m_flags And TDEFL_FILTER_MATCHES) <> 0) And (cur_match_len <= 5)) Then
            cur_match_dist = 0: cur_match_len = 0
        End If
        If d.m_saved_match_len <> 0 Then
            If cur_match_len > d.m_saved_match_len Then
                tdefl_record_literal d, d.m_saved_lit And &HFF
                If cur_match_len >= 128 Then
                    tdefl_record_match d, cur_match_len, cur_match_dist
                    d.m_saved_match_len = 0: len_to_move = cur_match_len
                Else
                    d.m_saved_lit = d.m_dict(cur_pos): d.m_saved_match_dist = cur_match_dist: d.m_saved_match_len = cur_match_len
                End If
            Else
                tdefl_record_match d, d.m_saved_match_len, d.m_saved_match_dist
                len_to_move = d.m_saved_match_len - 1: d.m_saved_match_len = 0
            End If
        ElseIf cur_match_dist = 0 Then
            tdefl_record_literal d, d.m_dict(cur_pos)
        ElseIf (d.m_greedy_parsing <> 0) Or ((d.m_flags And TDEFL_RLE_MATCHES) <> 0) Or (cur_match_len >= 128) Then
            tdefl_record_match d, cur_match_len, cur_match_dist
            len_to_move = cur_match_len
        Else
            d.m_saved_lit = d.m_dict(cur_pos): d.m_saved_match_dist = cur_match_dist: d.m_saved_match_len = cur_match_len
        End If
        ' Move the lookahead forward by len_to_move bytes.
        d.m_lookahead_pos = d.m_lookahead_pos + len_to_move
        d.m_lookahead_size = d.m_lookahead_size - len_to_move
        d.m_dict_size = d.m_dict_size + len_to_move     ' MZ_MIN(d->m_dict_size + len_to_move, TDEFL_LZ_DICT_SIZE)
        If d.m_dict_size > TDEFL_LZ_DICT_SIZE Then d.m_dict_size = TDEFL_LZ_DICT_SIZE
        ' Check if it's time to flush the current LZ codes to the internal output buffer.
        If (d.m_pLZ_code_buf > TDEFL_LZ_CODE_BUF_SIZE - 8) Or _
           ((d.m_total_lz_bytes > 31& * 1024&) And ((((d.m_pLZ_code_buf * 115) \ 128) >= d.m_total_lz_bytes) Or ((d.m_flags And TDEFL_FORCE_ALL_RAW_BLOCKS) <> 0))) Then
            d.m_pSrc = pSrc: d.m_src_buf_left = src_buf_left
            n = tdefl_flush_block(d, outBuf, 0)
            If n <> 0 Then
                tdefl_compress_normal = (n >= 0)
                Exit Function
            End If
        End If
    Loop

    d.m_pSrc = pSrc: d.m_src_buf_left = src_buf_left
    tdefl_compress_normal = True
End Function

' (VB) *d->m_pIn_buf_size and *d->m_pOut_buf_size are pIn_buf_size and pOut_buf_size (never NULL here)
Private Function tdefl_flush_output_buffer(d As tdefl_compressor, outBuf() As Byte, pIn_buf_size As Long, pOut_buf_size As Long) As Long
    Dim n As Long
    ' if (d->m_pIn_buf_size)
    pIn_buf_size = d.m_pSrc - d.m_pIn_buf

    ' if (d->m_pOut_buf_size)
    n = d.m_out_buf_size - d.m_out_buf_ofs              ' MZ_MIN(*d->m_pOut_buf_size - d->m_out_buf_ofs, d->m_output_flush_remaining)
    If d.m_output_flush_remaining < n Then n = d.m_output_flush_remaining
    If n > 0 Then CopyMemory VarPtr(outBuf(d.m_pOut_buf + d.m_out_buf_ofs)), VarPtr(d.m_output_buf(d.m_output_flush_ofs)), n
    d.m_output_flush_ofs = d.m_output_flush_ofs + n
    d.m_output_flush_remaining = d.m_output_flush_remaining - n
    d.m_out_buf_ofs = d.m_out_buf_ofs + n

    pOut_buf_size = d.m_out_buf_ofs

    If (d.m_finished <> 0) And (d.m_output_flush_remaining = 0) Then
        tdefl_flush_output_buffer = TDEFL_STATUS_DONE
    Else
        tdefl_flush_output_buffer = TDEFL_STATUS_OKAY
    End If
End Function

' Compresses a block of data, consuming as much of the input as possible, and writing as much compressed data as possible.
' (VB) pIn_buf is inBuf(inOfs ..) (NULL: inBuf unallocated), pOut_buf outBuf(outOfs ..) (NULL: unallocated); their
' sizes are pIn_buf_size and pOut_buf_size (the stream API never passes NULL for them)
Private Function tdefl_compress(d As tdefl_compressor, inBuf() As Byte, ByVal inOfs As Long, pIn_buf_size As Long, outBuf() As Byte, ByVal outOfs As Long, pOut_buf_size As Long, ByVal flush As Long) As Long
    Dim inNull As Boolean, outNull As Boolean
    ' (if (!d): a VB compressor is never NULL)
    inNull = Not mz_bufAllocated(inBuf): outNull = Not mz_bufAllocated(outBuf)

    d.m_pIn_buf = inOfs                                 ' d->m_pIn_buf_size = pIn_buf_size: (VB) the ByRef pIn_buf_size
    d.m_pOut_buf = outOfs: d.m_out_buf_size = pOut_buf_size
    d.m_pSrc = inOfs: d.m_src_buf_left = pIn_buf_size
    d.m_out_buf_ofs = 0
    d.m_flush = flush

    ' ((d->m_pPut_buf_func != NULL) == ((pOut_buf != NULL) || (pOut_buf_size != NULL))): never (no callback, a size pointer)
    If (d.m_prev_return_status <> TDEFL_STATUS_OKAY) Or ((d.m_wants_to_finish <> 0) And (flush <> TDEFL_FINISH)) Or _
       ((pIn_buf_size <> 0) And inNull) Or ((pOut_buf_size <> 0) And outNull) Then
        pIn_buf_size = 0
        pOut_buf_size = 0
        d.m_prev_return_status = TDEFL_STATUS_BAD_PARAM
        tdefl_compress = d.m_prev_return_status
        Exit Function
    End If
    If flush = TDEFL_FINISH Then d.m_wants_to_finish = d.m_wants_to_finish Or 1

    If (d.m_output_flush_remaining <> 0) Or (d.m_finished <> 0) Then
        d.m_prev_return_status = tdefl_flush_output_buffer(d, outBuf, pIn_buf_size, pOut_buf_size)
        tdefl_compress = d.m_prev_return_status
        Exit Function
    End If

    ' (MINIZ_USE_UNALIGNED_LOADS_AND_STORES && MINIZ_LITTLE_ENDIAN)
    If ((d.m_flags And TDEFL_MAX_PROBES_MASK) = 1) And _
       ((d.m_flags And TDEFL_GREEDY_PARSING_FLAG) <> 0) And _
       ((d.m_flags And (TDEFL_FILTER_MATCHES Or TDEFL_FORCE_ALL_RAW_BLOCKS Or TDEFL_RLE_MATCHES)) = 0) Then
        If Not tdefl_compress_fast(d, inBuf, outBuf) Then
            tdefl_compress = d.m_prev_return_status
            Exit Function
        End If
    Else
        If Not tdefl_compress_normal(d, inBuf, outBuf) Then
            tdefl_compress = d.m_prev_return_status
            Exit Function
        End If
    End If

    If ((d.m_flags And (TDEFL_WRITE_ZLIB_HEADER Or TDEFL_COMPUTE_ADLER32)) <> 0) And (Not inNull) Then
        d.m_adler32 = mz_adler32(d.m_adler32, inBuf, inOfs, d.m_pSrc - inOfs)
    End If

    If (flush <> 0) And (d.m_lookahead_size = 0) And (d.m_src_buf_left = 0) And (d.m_output_flush_remaining = 0) Then
        If tdefl_flush_block(d, outBuf, flush) < 0 Then
            tdefl_compress = d.m_prev_return_status
            Exit Function
        End If
        If flush = TDEFL_FINISH Then d.m_finished = 1 Else d.m_finished = 0
        If flush = TDEFL_FULL_FLUSH Then
            Erase d.m_hash: ReDim d.m_hash(0 To TDEFL_LZ_HASH_SIZE - 1)     ' MZ_CLEAR_OBJ(d->m_hash); MZ_CLEAR_OBJ(d->m_next);
            Erase d.m_next: ReDim d.m_next(0 To TDEFL_LZ_DICT_SIZE - 1)
            d.m_dict_size = 0
        End If
    End If

    d.m_prev_return_status = tdefl_flush_output_buffer(d, outBuf, pIn_buf_size, pOut_buf_size)
    tdefl_compress = d.m_prev_return_status
End Function

' Initializes the compressor. (VB) pPut_buf_func and pPut_buf_user are NULL (the stream API's)
Private Function tdefl_init(d As tdefl_compressor, ByVal flags As Long) As Long
    Dim i As Long
    ' d->m_pPut_buf_func = pPut_buf_func; d->m_pPut_buf_user = pPut_buf_user;
    d.m_flags = flags: d.m_max_probes(0) = 1 + ((flags And &HFFF&) + 2) \ 3
    If (flags And TDEFL_GREEDY_PARSING_FLAG) <> 0 Then d.m_greedy_parsing = 1 Else d.m_greedy_parsing = 0
    d.m_max_probes(1) = 1 + (((flags And &HFFF&) \ 4) + 2) \ 3
    If (flags And TDEFL_NONDETERMINISTIC_PARSING_FLAG) = 0 Then
        Erase d.m_hash: ReDim d.m_hash(0 To TDEFL_LZ_HASH_SIZE - 1)         ' MZ_CLEAR_OBJ(d->m_hash);
    End If
    d.m_lookahead_pos = 0: d.m_lookahead_size = 0: d.m_dict_size = 0: d.m_total_lz_bytes = 0: d.m_lz_code_buf_dict_pos = 0: d.m_bits_in = 0
    d.m_output_flush_ofs = 0: d.m_output_flush_remaining = 0: d.m_finished = 0: d.m_block_index = 0: d.m_bit_buffer = 0: d.m_wants_to_finish = 0
    d.m_pLZ_code_buf = 1: d.m_pLZ_flags = 0: d.m_num_flags_left = 8
    d.m_pOutput_buf = 0: d.m_pOutput_buf_end = 0: d.m_prev_return_status = TDEFL_STATUS_OKAY
    d.m_saved_match_dist = 0: d.m_saved_match_len = 0: d.m_saved_lit = 0: d.m_adler32 = 1
    d.m_pIn_buf = 0: d.m_pOut_buf = 0
    d.m_out_buf_size = 0                                ' d->m_pIn_buf_size = NULL; d->m_pOut_buf_size = NULL;
    d.m_flush = TDEFL_NO_FLUSH: d.m_pSrc = 0: d.m_src_buf_left = 0: d.m_out_buf_ofs = 0
    For i = 0 To TDEFL_MAX_HUFF_SYMBOLS_0 - 1
        d.m_huff_count(0, i) = 0
    Next
    For i = 0 To TDEFL_MAX_HUFF_SYMBOLS_1 - 1
        d.m_huff_count(1, i) = 0
    Next
    tdefl_init = TDEFL_STATUS_OKAY
End Function

Private Function tdefl_get_adler32(d As tdefl_compressor) As Long
    tdefl_get_adler32 = d.m_adler32
End Function

' Create tdefl_compress() flags given zlib-style compression parameters.
' level may range from [0,10] (where 10 is absolute max compression, but may be much slower on some files)
' window_bits may be -15 (raw deflate) or 15 (zlib)
' strategy may be either MZ_DEFAULT_STRATEGY, MZ_FILTERED, MZ_HUFFMAN_ONLY, MZ_RLE, or MZ_FIXED
Public Function tdefl_create_comp_flags_from_zip_params(ByVal level As Long, ByVal window_bits As Long, ByVal strategy As Long) As Long
    Dim comp_flags As Long
    If Not tablesReady Then tdefl_initTables
    If level >= 0 Then
        If level < 10 Then comp_flags = s_tdefl_num_probes(level) Else comp_flags = s_tdefl_num_probes(10)
    Else
        comp_flags = s_tdefl_num_probes(6)
    End If
    If level <= 3 Then comp_flags = comp_flags Or TDEFL_GREEDY_PARSING_FLAG
    If window_bits > 0 Then comp_flags = comp_flags Or TDEFL_WRITE_ZLIB_HEADER

    If level = 0 Then
        comp_flags = comp_flags Or TDEFL_FORCE_ALL_RAW_BLOCKS
    ElseIf strategy = MZ_FILTERED Then
        comp_flags = comp_flags Or TDEFL_FILTER_MATCHES
    ElseIf strategy = MZ_HUFFMAN_ONLY Then
        comp_flags = comp_flags And Not TDEFL_MAX_PROBES_MASK
    ElseIf strategy = MZ_FIXED Then
        comp_flags = comp_flags Or TDEFL_FORCE_ALL_STATIC_BLOCKS
    ElseIf strategy = MZ_RLE Then
        comp_flags = comp_flags Or TDEFL_RLE_MATCHES
    End If

    tdefl_create_comp_flags_from_zip_params = comp_flags
End Function

' ------------------- zlib-style API's (the compression half)

' (VB) zalloc(opaque, 1, sizeof(tdefl_compressor)): a free entry of the pool, or a new one (its arrays allocated);
' the handle is MZ_DEFLATE_STATE_BASE + its index
Private Function compressor_alloc() As Long
    Dim i As Long
    For i = 1 To compressorsCount
        If Not compressors(i).inUse Then Exit For
    Next
    If i > compressorsCount Then
        compressorsCount = compressorsCount + 1
        ReDim Preserve compressors(1 To compressorsCount)
        ReDim compressors(i).m_dict(0 To TDEFL_LZ_DICT_SIZE + TDEFL_MAX_MATCH_LEN - 2)
        ReDim compressors(i).m_huff_count(0 To TDEFL_MAX_HUFF_TABLES - 1, 0 To TDEFL_MAX_HUFF_SYMBOLS - 1)
        ReDim compressors(i).m_huff_codes(0 To TDEFL_MAX_HUFF_TABLES - 1, 0 To TDEFL_MAX_HUFF_SYMBOLS - 1)
        ReDim compressors(i).m_huff_code_sizes(0 To TDEFL_MAX_HUFF_TABLES - 1, 0 To TDEFL_MAX_HUFF_SYMBOLS - 1)
        ReDim compressors(i).m_lz_code_buf(0 To TDEFL_LZ_CODE_BUF_SIZE - 1)
        ReDim compressors(i).m_next(0 To TDEFL_LZ_DICT_SIZE - 1)
        ReDim compressors(i).m_hash(0 To TDEFL_LZ_HASH_SIZE - 1)
        ReDim compressors(i).m_output_buf(0 To TDEFL_OUT_BUF_SIZE - 1)
    End If
    compressors(i).inUse = True
    compressor_alloc = MZ_DEFLATE_STATE_BASE + i
End Function

' (VB) def_free_func's half for this pool: the entry's back to free (its arrays kept for the next stream)
Public Sub miniz_deflateStateFree(ByVal handle As Long)
    compressors(handle - MZ_DEFLATE_STATE_BASE).inUse = False
End Sub

' (VB) a stream's compressor, or 0 when its state is none or not a compressor's (where the C would misread memory)
Private Function compressorOf(pStream As z_stream) As Long
    If pStream.state > MZ_DEFLATE_STATE_BASE Then
        If pStream.state - MZ_DEFLATE_STATE_BASE <= compressorsCount Then compressorOf = pStream.state - MZ_DEFLATE_STATE_BASE
    End If
End Function

Public Function mz_deflateInit(pStream As z_stream, ByVal level As Long) As Long
    mz_deflateInit = mz_deflateInit2(pStream, level, MZ_DEFLATED, MZ_DEFAULT_WINDOW_BITS, 9, MZ_DEFAULT_STRATEGY)
End Function

Public Function mz_deflateInit2(pStream As z_stream, ByVal level As Long, ByVal method As Long, ByVal window_bits As Long, ByVal mem_level As Long, ByVal strategy As Long) As Long
    Dim pComp As Long, comp_flags As Long
    comp_flags = TDEFL_COMPUTE_ADLER32 Or tdefl_create_comp_flags_from_zip_params(level, window_bits, strategy)

    ' (if (!pStream) return MZ_STREAM_ERROR;: a VB stream is never NULL)
    If (method <> MZ_DEFLATED) Or ((mem_level < 1) Or (mem_level > 9)) Or ((window_bits <> MZ_DEFAULT_WINDOW_BITS) And (-window_bits <> MZ_DEFAULT_WINDOW_BITS)) Then
        mz_deflateInit2 = MZ_PARAM_ERROR
        Exit Function
    End If

    pStream.data_type = 0
    pStream.adler = MZ_ADLER32_INIT
    pStream.msg = vbNullString
    pStream.reserved = 0
    pStream.total_in = 0
    pStream.total_out = 0
    ' if (!pStream->zalloc) pStream->zalloc = def_alloc_func;
    ' if (!pStream->zfree) pStream->zfree = def_free_func;

    pComp = compressor_alloc()                          ' (tdefl_compressor *)pStream->zalloc(pStream->opaque, 1, sizeof(tdefl_compressor))
    If pComp = 0 Then
        mz_deflateInit2 = MZ_MEM_ERROR
        Exit Function
    End If

    pStream.state = pComp

    If tdefl_init(compressors(pComp - MZ_DEFLATE_STATE_BASE), comp_flags) <> TDEFL_STATUS_OKAY Then
        mz_deflateEnd pStream
        mz_deflateInit2 = MZ_PARAM_ERROR
        Exit Function
    End If

    mz_deflateInit2 = MZ_OK
End Function

Public Function mz_deflateReset(pStream As z_stream) As Long
    Dim c As Long
    c = compressorOf(pStream)
    If c = 0 Then                                       ' (!pStream) || (!pStream->state) || (!pStream->zalloc) || (!pStream->zfree)
        mz_deflateReset = MZ_STREAM_ERROR
        Exit Function
    End If
    pStream.total_in = 0: pStream.total_out = 0
    tdefl_init compressors(c), compressors(c).m_flags
    mz_deflateReset = MZ_OK
End Function

' (VB) next_in / next_out: offsets into inBuf / outBuf (next_in NULL: inBuf unallocated, next_out NULL: outBuf)
Public Function mz_deflate(pStream As z_stream, ByVal flush As Long, inBuf() As Byte, outBuf() As Byte) As Long
    Dim in_bytes As Long, out_bytes As Long, orig_total_in As Double, orig_total_out As Double, mz_status As Long
    Dim defl_status As Long, c As Long
    mz_status = MZ_OK

    c = compressorOf(pStream)
    If (c = 0) Or (flush < 0) Or (flush > MZ_FINISH) Then
        mz_deflate = MZ_STREAM_ERROR
        Exit Function
    End If
    If Not mz_bufAllocated(outBuf) Then                 ' (!pStream->next_out)
        mz_deflate = MZ_STREAM_ERROR
        Exit Function
    End If
    If pStream.avail_out = 0 Then
        mz_deflate = MZ_BUF_ERROR
        Exit Function
    End If

    If flush = MZ_PARTIAL_FLUSH Then flush = MZ_SYNC_FLUSH

    If compressors(c).m_prev_return_status = TDEFL_STATUS_DONE Then
        If flush = MZ_FINISH Then mz_deflate = MZ_STREAM_END Else mz_deflate = MZ_BUF_ERROR
        Exit Function
    End If

    orig_total_in = pStream.total_in: orig_total_out = pStream.total_out
    Do
        in_bytes = pStream.avail_in: out_bytes = pStream.avail_out

        defl_status = tdefl_compress(compressors(c), inBuf, pStream.next_in, in_bytes, outBuf, pStream.next_out, out_bytes, flush)
        pStream.next_in = pStream.next_in + in_bytes: pStream.avail_in = pStream.avail_in - in_bytes
        pStream.total_in = pStream.total_in + in_bytes: pStream.adler = tdefl_get_adler32(compressors(c))

        pStream.next_out = pStream.next_out + out_bytes: pStream.avail_out = pStream.avail_out - out_bytes
        pStream.total_out = pStream.total_out + out_bytes

        If defl_status < 0 Then
            mz_status = MZ_STREAM_ERROR
            Exit Do
        ElseIf defl_status = TDEFL_STATUS_DONE Then
            mz_status = MZ_STREAM_END
            Exit Do
        ElseIf pStream.avail_out = 0 Then
            Exit Do
        ElseIf (pStream.avail_in = 0) And (flush <> MZ_FINISH) Then
            If (flush <> 0) Or (pStream.total_in <> orig_total_in) Or (pStream.total_out <> orig_total_out) Then Exit Do
            mz_deflate = MZ_BUF_ERROR                   ' Can't make forward progress without some input.
            Exit Function
        End If
    Loop
    mz_deflate = mz_status
End Function

Public Function mz_deflateEnd(pStream As z_stream) As Long
    ' (if (!pStream) return MZ_STREAM_ERROR;: a VB stream is never NULL)
    If pStream.state <> 0 Then
        def_free_func pStream.state                     ' pStream->zfree(pStream->opaque, pStream->state)
        pStream.state = 0
    End If
    mz_deflateEnd = MZ_OK
End Function

' (VB) mz_ulong: a Double (pStream is not used)
Public Function mz_deflateBound(pStream As z_stream, ByVal source_len As Double) As Double
    Dim a As Double, b As Double
    ' This is really over conservative. (And lame, but it's actually pretty tricky to compute a true upper bound given the way tdefl's blocking works.)
    a = 128 + Fix((source_len * 110) / 100)
    b = 128 + source_len + (Fix(source_len / (31 * 1024)) + 1) * 5
    If a > b Then mz_deflateBound = a Else mz_deflateBound = b
End Function

' (VB) pDest(0 ..) and pSource(0 ..): whole arrays from their first byte
Public Function mz_compress2(pDest() As Byte, pDest_len As Long, pSource() As Byte, ByVal source_len As Long, ByVal level As Long) As Long
    Dim status As Long, stream As z_stream             ' memset(&stream, 0, sizeof(stream));

    ' In case mz_ulong is 64-bits (argh I hate longs). (VB: Longs)

    stream.next_in = 0
    stream.avail_in = source_len
    stream.next_out = 0
    stream.avail_out = pDest_len

    status = mz_deflateInit(stream, level)
    If status <> MZ_OK Then
        mz_compress2 = status
        Exit Function
    End If

    status = mz_deflate(stream, MZ_FINISH, pSource, pDest)
    If status <> MZ_STREAM_END Then
        mz_deflateEnd stream
        If status = MZ_OK Then mz_compress2 = MZ_BUF_ERROR Else mz_compress2 = status
        Exit Function
    End If

    pDest_len = stream.total_out
    mz_compress2 = mz_deflateEnd(stream)
End Function

Public Function mz_compress(pDest() As Byte, pDest_len As Long, pSource() As Byte, ByVal source_len As Long) As Long
    mz_compress = mz_compress2(pDest, pDest_len, pSource, source_len, MZ_DEFAULT_COMPRESSION)
End Function

Public Function mz_compressBound(ByVal source_len As Double) As Double
    Dim none As z_stream
    mz_compressBound = mz_deflateBound(none, source_len)
End Function
