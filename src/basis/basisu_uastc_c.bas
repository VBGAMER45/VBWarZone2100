Attribute VB_Name = "basisu_uastc_c"
' Basis Universal's transcoder (ref/basis_universal @ df079ec, the commit 4.7.0 pins: transcoder/basisu_transcoder.cpp,
' transcoder/basisu_transcoder_uastc.h), its UASTC LDR 4x4 block part: unpack_uastc, the UASTC -> ASTC 4x4, BC7,
' BC1, BC3, BC4 and BC5 block transcoders with the encoders they use (pack_astc_block, encode_bc7_block, encode_bc1,
' encode_bc4), uastc_init and the BC1 single colour tables basisu_transcoder_init makes. The ETC1 / ETC2 / EAC,
' PVRTC, ATC, FXT1 and HDR transcoders are not ported (basisu_transcoder_c's basis_is_format_supported says so).
' A uastc_block is 16 bytes of a Byte array at an offset; the color32 pixels of a block are px(0 To 63), pixel i's
' r, g, b, a at 4 * i (C's color32 block_pixels[4][4] in memory order); a destination block is bytes of a Byte
' array at an offset. The constant tables are basisu_tables_vb's.
'
' Float arithmetic (the BC7 p-bit choices, BC1's principal axis and least squares) is rounded to float after every
' operation as the C's float code is: fadd_, fsub_, fmul_ and fdiv_ pass their result through r32_'s ByVal Single
' parameter (stored as 4 bytes, read back: one x87 operation on two floats, rounded once to float, is the correctly
' rounded float result SSE gives). A store to a Single variable is not enough: the compiler (unrounded floating point
' allowed) hands the register's unrounded value on to the next use, and Fix(fmul_(...)) gave 1023 for a product that
' is 1024.0 as a float (encode_bc1's principal axis: one block in 17 normal maps came out different from the C's).
' read_bits reads byte by byte, as the C does when BASISD_USE_UNALIGNED_WORD_READS is 0 (the same bits).
Option Explicit

Private Const TOTAL_UASTC_MODES As Long = 19
Private Const UASTC_MODE_INDEX_SOLID_COLOR As Long = 8
Private Const TOTAL_ASTC_BC7_COMMON_PARTITIONS2 As Long = 30
Private Const TOTAL_ASTC_BC7_COMMON_PARTITIONS3 As Long = 11
Private Const TOTAL_BC7_3_ASTC2_COMMON_PARTITIONS As Long = 19
Private Const BC7ENC_TOTAL_ASTC_RANGES As Long = 21
Private Const BC7ENC_MODE_6_OPTIMAL_INDEX As Long = 5
Private Const BC7ENC_MODE_5_OPTIMAL_INDEX As Long = 1
Private Const ASTC_BLOCK_MODE_BITS As Long = 11
Private Const ASTC_PART_BITS As Long = 2
Private Const ASTC_CEM_BITS As Long = 4
Private Const ASTC_PARTITION_INDEX_BITS As Long = 10
Private Const ASTC_CCS_BITS As Long = 2
Private Const cEncodeBC1HighQuality As Long = 1
Private Const cEncodeBC1HigherQuality As Long = 2
Private Const cEncodeBC1UseSelectors As Long = 4

' basisu_transcoder_uastc.h: struct astc_block_desc (the uint8_t arrays are Longs)
Private Type astc_block_desc
    m_weight_range As Long          ' weight BISE range
    m_subsets As Long               ' number of ASTC partitions
    m_partition_seed As Long        ' partition pattern seed
    m_cem As Long                   ' color endpoint mode used by all subsets
    m_ccs As Long                   ' color component selector (dual plane only)
    m_dual_plane As Boolean         ' true if dual plane
    m_endpoints(0 To 17) As Long    ' endpoint values, in RR GG BB etc. order
    m_weights(0 To 63) As Long      ' weight index values, raster order, in P0 P1, P0 P1, etc. or P0, P0, P0, P0, etc. order
End Type

' struct unpacked_uastc_block (color32 m_solid_color: r, g, b, a)
Private Type unpacked_uastc_block
    m_astc As astc_block_desc
    m_mode As Long
    m_common_pattern As Long
    m_solid_color(0 To 3) As Long
    m_bc1_hint0 As Boolean
    m_bc1_hint1 As Boolean
    m_etc1_flip As Boolean
    m_etc1_diff As Boolean
    m_etc1_inten0 As Long
    m_etc1_inten1 As Long
    m_etc1_bias As Long
    m_etc2_hints As Long
    m_etc1_selector As Long
    m_etc1_r As Long
    m_etc1_g As Long
    m_etc1_b As Long
End Type

' struct bc7_optimization_results (color_quad_u8 m_low[3], m_high[3]: [subset * 4 + comp]; uint32_t m_pbits[3][2]:
' [subset * 2 + i])
Private Type bc7_optimization_results
    m_mode As Long
    m_partition As Long
    m_selectors(0 To 15) As Long
    m_alpha_selectors(0 To 15) As Long
    m_low(0 To 11) As Long
    m_high(0 To 11) As Long
    m_pbits(0 To 5) As Long
    m_index_selector As Long
    m_rotation As Long
End Type

' astc_quant_bin g_astc_unquant[BC7ENC_TOTAL_ASTC_RANGES][256] (uastc_init fills it)
Private g_astc_unquant_m_unquant(0 To 20, 0 To 255) As Long
Private g_astc_unquant_m_index(0 To 20, 0 To 255) As Long
' endpoint_err g_bc7_mode_6_optimal_endpoints[256][2] ([c][pbit]), g_bc7_mode_5_optimal_endpoints[256]
Private g_bc7_mode_6_optimal_endpoints_m_error(0 To 255, 0 To 1) As Long
Private g_bc7_mode_6_optimal_endpoints_m_lo(0 To 255, 0 To 1) As Long
Private g_bc7_mode_6_optimal_endpoints_m_hi(0 To 255, 0 To 1) As Long
Private g_bc7_mode_5_optimal_endpoints_m_error(0 To 255) As Long
Private g_bc7_mode_5_optimal_endpoints_m_lo(0 To 255) As Long
Private g_bc7_mode_5_optimal_endpoints_m_hi(0 To 255) As Long
' bc1_match_entry g_bc1_match5_equals_1[256], g_bc1_match6_equals_1[256] (selector 1), ..._equals_0 (selector 0)
Private g_bc1_match5_equals_1_m_hi(0 To 255) As Long
Private g_bc1_match5_equals_1_m_lo(0 To 255) As Long
Private g_bc1_match6_equals_1_m_hi(0 To 255) As Long
Private g_bc1_match6_equals_1_m_lo(0 To 255) As Long
Private g_bc1_match5_equals_0_m_hi(0 To 255) As Long
Private g_bc1_match5_equals_0_m_lo(0 To 255) As Long
Private g_bc1_match6_equals_0_m_hi(0 To 255) As Long
Private g_bc1_match6_equals_0_m_lo(0 To 255) As Long
' g_astc_endpoint_unquant_params[BC7ENC_TOTAL_ASTC_RANGES] (taken right from the ASTC spec)
Private g_astc_endpoint_unquant_params_m_pB_str(0 To 20) As String
Private g_astc_endpoint_unquant_params_m_c(0 To 20) As Long
' (VB) g_astc_weight_tables[6] and s_uastc_to_bc1_weights[6], as [table][index] tables
Private g_astc_weight_tables(0 To 5, 0 To 31) As Long
Private s_uastc_to_bc1_weights(0 To 5, 0 To 31) As Long
' (VB) 1 << n
Private p2(0 To 31) As Long
' (VB) encode_bc1's constant float quotients (C: 1.0f / 255.0f etc., folded to the nearest float)
Private f_1_255 As Single, f_31_255 As Single, f_63_255 As Single
' (VB) the blocks the transcoders work on (C: locals; nothing here is re-entered)
Private ub As unpacked_uastc_block
Private bc7r As bc7_optimization_results
Private zero_bc7 As bc7_optimization_results
Private px(0 To 63) As Long
Private ready As Boolean

'--------------------------------------------------------------------------
' (VB) helpers

' the float v: a ByVal Single parameter is stored as 4 bytes when the call is made, and read back here (see the header)
Private Function r32_(ByVal v As Single) As Single
    r32_ = v
End Function

Private Function fadd_(ByVal a As Single, ByVal b As Single) As Single
    fadd_ = r32_(a + b)
End Function

Private Function fsub_(ByVal a As Single, ByVal b As Single) As Single
    fsub_ = r32_(a - b)
End Function

Private Function fmul_(ByVal a As Single, ByVal b As Single) As Single
    fmul_ = r32_(a * b)
End Function

Private Function fdiv_(ByVal a As Single, ByVal b As Single) As Single
    fdiv_ = r32_(a / b)
End Function

' basisu::clamp / clampi
Private Function clampi(ByVal v As Long, ByVal l As Long, ByVal h As Long) As Long
    If v < l Then
        clampi = l
    ElseIf v > h Then
        clampi = h
    Else
        clampi = v
    End If
End Function

' static inline uint32_t read_bit(const uint8_t* pBuf, uint32_t& bit_offset)
Private Function read_bit(b() As Byte, ByVal o As Long, bit_offset As Long) As Long
    read_bit = (b(o + bit_offset \ 8) \ p2(bit_offset And 7)) And 1
    bit_offset = bit_offset + 1
End Function

' read_bits1_to_9 / read_bits1_to_9_fst: codesize (0 to 9) bits at bit_offset of the block at b(o)
Private Function read_bits(b() As Byte, ByVal o As Long, bit_offset As Long, ByVal codesize As Long) As Long
    Dim p As Long, byte_bit_offset As Long, bits As Long, bits_read As Long
    If codesize = 0 Then Exit Function
    p = o + bit_offset \ 8
    byte_bit_offset = bit_offset And 7
    bits = b(p) \ p2(byte_bit_offset)
    bits_read = 8 - byte_bit_offset
    If codesize > bits_read Then bits = bits Or (b(p + 1) * p2(bits_read))
    bit_offset = bit_offset + codesize
    read_bits = bits And (p2(codesize) - 1)
End Function

'--------------------------------------------------------------------------
' init: uastc_init, the BC1 single colour tables (basisu_transcoder_init), the VB tables

Public Sub basisu_uastc_vbinit()
    Dim i As Long, j As Long
    If ready Then Exit Sub
    p2(0) = 1
    For i = 1 To 30
        p2(i) = p2(i - 1) * 2
    Next
    p2(31) = &H80000000
    basisu_tables_init
    For i = 0 To 20
        g_astc_endpoint_unquant_params_m_pB_str(i) = ""
        g_astc_endpoint_unquant_params_m_c(i) = 0
    Next
    set_unquant_params 4, "000000000", 204      ' 0-5
    set_unquant_params 6, "000000000", 113      ' 0-9
    set_unquant_params 7, "b000b0bb0", 93       ' 0-11
    set_unquant_params 9, "b0000bb00", 54       ' 0-19
    set_unquant_params 10, "cb000cbcb", 44      ' 0-23
    set_unquant_params 12, "cb0000cbc", 26      ' 0-39
    set_unquant_params 13, "dcb000dcb", 22      ' 0-47
    set_unquant_params 15, "dcb0000dc", 13      ' 0-79
    set_unquant_params 16, "edcb000ed", 11      ' 0-95
    set_unquant_params 18, "edcb0000e", 6       ' 0-159
    set_unquant_params 19, "fedcb000f", 5       ' 0-191
    For i = 0 To 31
        If i < 2 Then g_astc_weight_tables(1, i) = g_bc7_weights1(i)
        If i < 4 Then g_astc_weight_tables(2, i) = g_bc7_weights2(i)
        If i < 8 Then g_astc_weight_tables(3, i) = g_bc7_weights3(i)
        If i < 16 Then g_astc_weight_tables(4, i) = g_astc_weights4(i)
        g_astc_weight_tables(5, i) = g_astc_weights5(i)
        If i < 2 Then s_uastc_to_bc1_weights(1, i) = s_uastc1_to_bc1(i)
        If i < 4 Then s_uastc_to_bc1_weights(2, i) = s_uastc2_to_bc1(i)
        If i < 8 Then s_uastc_to_bc1_weights(3, i) = s_uastc3_to_bc1(i)
        If i < 16 Then s_uastc_to_bc1_weights(4, i) = s_uastc4_to_bc1(i)
        s_uastc_to_bc1_weights(5, i) = s_uastc5_to_bc1(i)
    Next
    f_1_255 = fdiv_(1!, 255!)
    f_31_255 = fdiv_(31!, 255!)
    f_63_255 = fdiv_(63!, 255!)
    j = 0
    ready = True
End Sub

Private Sub set_unquant_params(ByVal range As Long, ByVal pB_str As String, ByVal c As Long)
    g_astc_endpoint_unquant_params_m_pB_str(range) = pB_str
    g_astc_endpoint_unquant_params_m_c(range) = c
End Sub

' int astc_get_levels(int range)
Private Function astc_get_levels(ByVal range As Long) As Long
    astc_get_levels = (1 + 2 * g_astc_bise_range_table(range, 1) + 4 * g_astc_bise_range_table(range, 2)) * p2(g_astc_bise_range_table(range, 0))
End Function

' bool astc_is_valid_endpoint_range(uint32_t range)
Private Function astc_is_valid_endpoint_range(ByVal range As Long) As Boolean
    If (g_astc_bise_range_table(range, 1) = 0) And (g_astc_bise_range_table(range, 2) = 0) Then
        astc_is_valid_endpoint_range = True
        Exit Function
    End If
    astc_is_valid_endpoint_range = (g_astc_endpoint_unquant_params_m_c(range) <> 0)
End Function

' uint32_t unquant_astc_endpoint(uint32_t packed_bits, uint32_t packed_trits, uint32_t packed_quints, uint32_t range)
Private Function unquant_astc_endpoint(ByVal packed_bits As Long, ByVal packed_trits As Long, ByVal packed_quints As Long, ByVal range As Long) As Long
    Dim bits As Long, trits As Long, quints As Long, val_ As Long, bits_left As Long, v As Long, n As Long
    Dim a As Long, c As Long, d As Long, b As Long, i As Long, ch As Long
    bits = g_astc_bise_range_table(range, 0)
    trits = g_astc_bise_range_table(range, 1)
    quints = g_astc_bise_range_table(range, 2)

    val_ = 0
    If (trits = 0) And (quints = 0) Then
        bits_left = 8
        Do While bits_left > 0
            v = packed_bits
            n = bits_left
            If bits < n Then n = bits
            If n < bits Then v = v \ p2(bits - n)
            val_ = val_ Or (v * p2(bits_left - n))
            bits_left = bits_left - n
        Loop
    Else
        If (packed_bits And 1) <> 0 Then a = 511 Else a = 0
        c = g_astc_endpoint_unquant_params_m_c(range)
        If trits <> 0 Then d = packed_trits Else d = packed_quints

        b = 0
        For i = 0 To 8
            b = b * 2
            ch = Asc(Mid$(g_astc_endpoint_unquant_params_m_pB_str(range), i + 1, 1))
            If ch <> 48 Then                                ' '0'
                ch = ch - 97                                ' 'a'
                b = b Or ((packed_bits \ p2(ch)) And 1)
            End If
        Next

        val_ = d * c + b
        val_ = val_ Xor a
        val_ = (a And &H80) Or (val_ \ 4)
    End If

    unquant_astc_endpoint = val_
End Function

' uint32_t unquant_astc_endpoint_val(uint32_t packed_val, uint32_t range)
Private Function unquant_astc_endpoint_val(ByVal packed_val As Long, ByVal range As Long) As Long
    Dim bits As Long, trits As Long, quints As Long
    bits = g_astc_bise_range_table(range, 0)
    trits = g_astc_bise_range_table(range, 1)
    quints = g_astc_bise_range_table(range, 2)

    If (trits = 0) And (quints = 0) Then
        unquant_astc_endpoint_val = unquant_astc_endpoint(packed_val, 0, 0, range)
    ElseIf trits <> 0 Then
        unquant_astc_endpoint_val = unquant_astc_endpoint(packed_val And (p2(bits) - 1), packed_val \ p2(bits), 0, range)
    Else
        unquant_astc_endpoint_val = unquant_astc_endpoint(packed_val And (p2(bits) - 1), 0, packed_val \ p2(bits), range)
    End If
End Function

' void uastc_init()
Public Sub uastc_init()
    Dim range As Long, levels As Long, i As Long, j As Long, t As Long, order As Long, unq As Long
    Dim vals(0 To 255) As Long
    Dim c As Long, lp As Long, l As Long, h As Long, low As Long, high As Long, k As Long, er As Long
    Dim best_error As Long, best_lo As Long, best_hi As Long, w As Long

    basisu_uastc_vbinit

    For range = 0 To BC7ENC_TOTAL_ASTC_RANGES - 1
        If Not astc_is_valid_endpoint_range(range) Then GoTo continue_range

        levels = astc_get_levels(range)

        For i = 0 To levels - 1
            vals(i) = (unquant_astc_endpoint_val(i, range) * 256) Or i
        Next

        ' std::sort(vals, vals + levels) (the values are distinct)
        For i = 1 To levels - 1
            t = vals(i)
            j = i - 1
            Do While j >= 0
                If vals(j) <= t Then Exit Do
                vals(j + 1) = vals(j)
                j = j - 1
            Loop
            vals(j + 1) = t
        Next

        For i = 0 To levels - 1
            order = vals(i) And &HFF
            unq = vals(i) \ 256

            g_astc_unquant_m_unquant(range, order) = unq And &HFF
            g_astc_unquant_m_index(range, order) = i And &HFF
        Next
continue_range:
    Next

    ' TODO: Precompute?
    ' BC7 777.1
    w = g_bc7_weights4(BC7ENC_MODE_6_OPTIMAL_INDEX)
    For c = 0 To 255
        For lp = 0 To 1
            best_error = 65535                              ' UINT16_MAX
            best_lo = 0
            best_hi = 0

            For l = 0 To 127
                low = (l * 2) Or lp

                For h = 0 To 127
                    high = (h * 2) Or lp

                    k = (low * (64 - w) + high * w + 32) \ 64

                    er = (k - c) * (k - c)
                    If er < best_error Then
                        best_error = er
                        best_lo = l
                        best_hi = h
                    End If
                Next
            Next

            g_bc7_mode_6_optimal_endpoints_m_error(c, lp) = best_error
            g_bc7_mode_6_optimal_endpoints_m_lo(c, lp) = best_lo
            g_bc7_mode_6_optimal_endpoints_m_hi(c, lp) = best_hi
        Next
    Next

    ' BC7 777
    w = g_bc7_weights2(BC7ENC_MODE_5_OPTIMAL_INDEX)
    For c = 0 To 255
        best_error = 65535
        best_lo = 0
        best_hi = 0

        For l = 0 To 127
            low = (l * 2) Or (l \ 64)

            For h = 0 To 127
                high = (h * 2) Or (h \ 64)

                k = (low * (64 - w) + high * w + 32) \ 64

                er = (k - c) * (k - c)
                If er < best_error Then
                    best_error = er
                    best_lo = l
                    best_hi = h
                End If
            Next
        Next

        g_bc7_mode_5_optimal_endpoints_m_error(c) = best_error
        g_bc7_mode_5_optimal_endpoints_m_lo(c) = best_lo
        g_bc7_mode_5_optimal_endpoints_m_hi(c) = best_hi
    Next
End Sub

' static void prepare_bc1_single_color_table(bc1_match_entry* pTable, const uint8_t* pExpand, int size0, int size1, int sel)
Private Sub prepare_bc1_single_color_table(tbl_hi() As Long, tbl_lo() As Long, pExpand() As Long, ByVal size0 As Long, ByVal size1 As Long, ByVal sel As Long)
    Dim i As Long, lowest_e As Long, lo As Long, hi As Long, lo_e As Long, hi_e As Long, e As Long
    For i = 0 To 255
        lowest_e = 256
        For lo = 0 To size0 - 1
            For hi = 0 To size1 - 1
                lo_e = pExpand(lo)
                hi_e = pExpand(hi)

                If sel = 1 Then
                    ' Selector 1
                    e = Abs(((hi_e * 2 + lo_e) \ 3) - i)
                    e = e + (Abs(hi_e - lo_e) * 3) \ 100
                Else
                    ' Selector 0
                    e = Abs(hi_e - i)
                End If

                If e < lowest_e Then
                    tbl_hi(i) = hi
                    tbl_lo(i) = lo

                    lowest_e = e
                End If
            Next
        Next
    Next
End Sub

' basisu_transcoder_init's BC1 part (#if BASISD_SUPPORT_DXT1 || BASISD_SUPPORT_UASTC)
Public Sub basisu_uastc_bc1_init()
    Dim bc1_expand5(0 To 31) As Long, bc1_expand6(0 To 63) As Long, i As Long
    basisu_uastc_vbinit
    For i = 0 To 31
        bc1_expand5(i) = ((i * 8) Or (i \ 4)) And &HFF
    Next
    prepare_bc1_single_color_table g_bc1_match5_equals_1_m_hi, g_bc1_match5_equals_1_m_lo, bc1_expand5, 32, 32, 1
    prepare_bc1_single_color_table g_bc1_match5_equals_0_m_hi, g_bc1_match5_equals_0_m_lo, bc1_expand5, 1, 32, 0

    For i = 0 To 63
        bc1_expand6(i) = ((i * 4) Or (i \ 16)) And &HFF
    Next
    prepare_bc1_single_color_table g_bc1_match6_equals_1_m_hi, g_bc1_match6_equals_1_m_lo, bc1_expand6, 64, 64, 1
    prepare_bc1_single_color_table g_bc1_match6_equals_0_m_hi, g_bc1_match6_equals_0_m_lo, bc1_expand6, 1, 64, 0
End Sub

'--------------------------------------------------------------------------
' unpack_uastc

' const uint8_t* get_anchor_indices(uint32_t subsets, uint32_t mode, uint32_t common_pattern, const uint8_t*& pPartition_pattern):
' the anchors (3) and the pattern (16) copied out
Private Sub get_anchor_indices(ByVal subsets As Long, ByVal mode As Long, ByVal common_pattern As Long, anchors() As Long, pattern() As Long)
    Dim i As Long
    If subsets >= 2 Then
        If subsets = 3 Then
            For i = 0 To 15
                pattern(i) = g_astc_bc7_patterns3(common_pattern, i)
            Next
            For i = 0 To 2
                anchors(i) = g_astc_bc7_pattern3_anchors(common_pattern, i)
            Next
        ElseIf mode = 7 Then
            For i = 0 To 15
                pattern(i) = g_bc7_3_astc2_patterns2(common_pattern, i)
            Next
            For i = 0 To 2
                anchors(i) = g_bc7_3_astc2_patterns2_anchors(common_pattern, i)
            Next
        Else
            For i = 0 To 15
                pattern(i) = g_astc_bc7_patterns2(common_pattern, i)
            Next
            For i = 0 To 2
                anchors(i) = g_astc_bc7_pattern2_anchors(common_pattern, i)
            Next
        End If
    Else
        ' g_zero_pattern
        For i = 0 To 15
            pattern(i) = 0
        Next
        anchors(0) = 0
        anchors(1) = 0
        anchors(2) = 0
    End If
End Sub

' bool unpack_uastc(const uastc_block& blk, unpacked_uastc_block& unpacked, bool blue_contract_check, bool read_hints)
Private Function unpack_uastc_blk(b() As Byte, ByVal o As Long, unpacked As unpacked_uastc_block, ByVal blue_contract_check As Boolean, ByVal read_hints As Boolean) As Boolean
    Static tq_values(0 To 7) As Long, anchors(0 To 2) As Long, pattern(0 To 15) As Long
    Static invert_subset(0 To 2) As Boolean
    Dim mode As Long, bit_ofs As Long, subsets As Long, part_seed As Long, total_planes As Long
    Dim total_comps As Long, weight_bits As Long, total_values As Long, endpoint_range As Long, cem As Long
    Dim ep_bits As Long, ep_trits As Long, ep_quints As Long, total_tqs As Long, bundle_size As Long, mul As Long
    Dim i As Long, num_bits As Long, num_remaining As Long, accum As Long, accum_remaining As Long, next_tq_index As Long
    Dim value As Long, v As Long, a0 As Long, a1 As Long, a2 As Long, subset_index As Long, s0 As Long, s1 As Long
    Dim c As Long, t As Long, base As Long, any_flag As Boolean, weight_mask As Long, subset As Long

    mode = g_uastc_huff_modes(b(o) And 127)
    If mode >= TOTAL_UASTC_MODES Then Exit Function

    unpacked.m_mode = mode
    unpacked.m_common_pattern = 0

    bit_ofs = g_uastc_mode_huff_codes(mode, 1)

    If mode = UASTC_MODE_INDEX_SOLID_COLOR Then
        unpacked.m_solid_color(0) = read_bits(b, o, bit_ofs, 8)
        unpacked.m_solid_color(1) = read_bits(b, o, bit_ofs, 8)
        unpacked.m_solid_color(2) = read_bits(b, o, bit_ofs, 8)
        unpacked.m_solid_color(3) = read_bits(b, o, bit_ofs, 8)

        If read_hints Then
            unpacked.m_etc1_flip = False
            unpacked.m_etc1_diff = read_bit(b, o, bit_ofs) <> 0
            unpacked.m_etc1_inten0 = read_bits(b, o, bit_ofs, 3)
            unpacked.m_etc1_inten1 = 0
            unpacked.m_etc1_selector = read_bits(b, o, bit_ofs, 2)
            unpacked.m_etc1_r = read_bits(b, o, bit_ofs, 5)
            unpacked.m_etc1_g = read_bits(b, o, bit_ofs, 5)
            unpacked.m_etc1_b = read_bits(b, o, bit_ofs, 5)
            unpacked.m_etc1_bias = 0
            unpacked.m_etc2_hints = 0
        End If

        unpack_uastc_blk = True
        Exit Function
    End If

    If read_hints Then
        If g_uastc_mode_has_bc1_hint0(mode) <> 0 Then
            unpacked.m_bc1_hint0 = read_bit(b, o, bit_ofs) <> 0
        Else
            unpacked.m_bc1_hint0 = False
        End If

        If g_uastc_mode_has_bc1_hint1(mode) <> 0 Then
            unpacked.m_bc1_hint1 = read_bit(b, o, bit_ofs) <> 0
        Else
            unpacked.m_bc1_hint1 = False
        End If

        unpacked.m_etc1_flip = read_bit(b, o, bit_ofs) <> 0
        unpacked.m_etc1_diff = read_bit(b, o, bit_ofs) <> 0
        unpacked.m_etc1_inten0 = read_bits(b, o, bit_ofs, 3)
        unpacked.m_etc1_inten1 = read_bits(b, o, bit_ofs, 3)

        If g_uastc_mode_has_etc1_bias(mode) <> 0 Then
            unpacked.m_etc1_bias = read_bits(b, o, bit_ofs, 5)
        Else
            unpacked.m_etc1_bias = 0
        End If

        If g_uastc_mode_has_alpha(mode) <> 0 Then
            unpacked.m_etc2_hints = read_bits(b, o, bit_ofs, 8)
        Else
            unpacked.m_etc2_hints = 0
        End If
    Else
        bit_ofs = bit_ofs + g_uastc_mode_total_hint_bits(mode)
    End If

    subsets = 1
    Select Case mode
    Case 2, 4, 7, 9, 16
        unpacked.m_common_pattern = read_bits(b, o, bit_ofs, 5)
        subsets = 2
    Case 3
        unpacked.m_common_pattern = read_bits(b, o, bit_ofs, 4)
        subsets = 3
    End Select

    part_seed = 0
    Select Case mode
    Case 2, 4, 9, 16
        If unpacked.m_common_pattern >= TOTAL_ASTC_BC7_COMMON_PARTITIONS2 Then Exit Function

        part_seed = g_astc_bc7_common_partitions2_m_astc(unpacked.m_common_pattern)
    Case 3
        If unpacked.m_common_pattern >= TOTAL_ASTC_BC7_COMMON_PARTITIONS3 Then Exit Function

        part_seed = g_astc_bc7_common_partitions3_m_astc(unpacked.m_common_pattern)
    Case 7
        If unpacked.m_common_pattern >= TOTAL_BC7_3_ASTC2_COMMON_PARTITIONS Then Exit Function

        part_seed = g_bc7_3_astc2_common_partitions_m_astc2(unpacked.m_common_pattern)
    End Select

    total_planes = 1
    Select Case mode
    Case 6, 11, 13
        unpacked.m_astc.m_ccs = read_bits(b, o, bit_ofs, 2)
        total_planes = 2
    Case 17
        unpacked.m_astc.m_ccs = 3
        total_planes = 2
    End Select

    unpacked.m_astc.m_dual_plane = (total_planes = 2)

    unpacked.m_astc.m_subsets = subsets
    unpacked.m_astc.m_partition_seed = part_seed

    total_comps = g_uastc_mode_comps(mode)

    weight_bits = g_uastc_mode_weight_bits(mode)

    unpacked.m_astc.m_weight_range = g_uastc_mode_weight_ranges(mode)

    total_values = total_comps * 2 * subsets
    endpoint_range = g_uastc_mode_endpoint_ranges(mode)

    cem = g_uastc_mode_cem(mode)
    unpacked.m_astc.m_cem = cem

    ep_bits = g_astc_bise_range_table(endpoint_range, 0)
    ep_trits = g_astc_bise_range_table(endpoint_range, 1)
    ep_quints = g_astc_bise_range_table(endpoint_range, 2)

    total_tqs = 0
    bundle_size = 0
    mul = 0
    If ep_trits <> 0 Then
        total_tqs = (total_values + 4) \ 5
        bundle_size = 5
        mul = 3
    ElseIf ep_quints <> 0 Then
        total_tqs = (total_values + 2) \ 3
        bundle_size = 3
        mul = 5
    End If

    For i = 0 To total_tqs - 1
        If ep_trits <> 0 Then num_bits = 8 Else num_bits = 7
        If i = (total_tqs - 1) Then
            num_remaining = total_values - (total_tqs - 1) * bundle_size
            If ep_trits <> 0 Then
                Select Case num_remaining
                Case 1: num_bits = 2
                Case 2: num_bits = 4
                Case 3: num_bits = 5
                Case 4: num_bits = 7
                End Select
            ElseIf ep_quints <> 0 Then
                Select Case num_remaining
                Case 1: num_bits = 3
                Case 2: num_bits = 5
                End Select
            End If
        End If

        tq_values(i) = read_bits(b, o, bit_ofs, num_bits)
    Next

    accum = 0
    accum_remaining = 0
    next_tq_index = 0

    For i = 0 To total_values - 1
        value = read_bits(b, o, bit_ofs, ep_bits)

        If total_tqs <> 0 Then
            If accum_remaining = 0 Then
                accum = tq_values(next_tq_index)
                next_tq_index = next_tq_index + 1
                accum_remaining = bundle_size
            End If

            ' TODO: Optimize with tables
            v = accum Mod mul
            accum = accum \ mul
            accum_remaining = accum_remaining - 1

            value = value Or (v * p2(ep_bits))
        End If

        unpacked.m_astc.m_endpoints(i) = value And &HFF
    Next

    get_anchor_indices subsets, mode, unpacked.m_common_pattern, anchors, pattern

    ' (VB) The C reads the weights' (at most 64) bits into a uint64_t and shifts them out; reading each field at its
    ' place in the block gives the same values.
    If mode = 18 Then
        ' Mode 18 is the only mode with more than 64 weight bits.
        For i = 0 To 15
            If i <> 0 Then
                unpacked.m_astc.m_weights(i) = read_bits(b, o, bit_ofs, weight_bits)
            Else
                unpacked.m_astc.m_weights(i) = read_bits(b, o, bit_ofs, weight_bits - 1)
            End If
        Next
    ElseIf total_planes = 2 Then
        ' Dual plane modes always have a single subset, and the first 2 weights are anchors.
        unpacked.m_astc.m_weights(0) = read_bits(b, o, bit_ofs, weight_bits - 1)
        unpacked.m_astc.m_weights(1) = read_bits(b, o, bit_ofs, weight_bits - 1)

        For i = 2 To 31
            unpacked.m_astc.m_weights(i) = read_bits(b, o, bit_ofs, weight_bits)
        Next
    ElseIf subsets = 1 Then
        ' First weight is always an anchor.
        unpacked.m_astc.m_weights(0) = read_bits(b, o, bit_ofs, weight_bits - 1)

        For i = 1 To 15
            unpacked.m_astc.m_weights(i) = read_bits(b, o, bit_ofs, weight_bits)
        Next
    Else
        a0 = anchors(0)
        a1 = anchors(1)
        a2 = anchors(2)

        For i = 0 To 15
            If (i = a0) Or (i = a1) Or (i = a2) Then
                unpacked.m_astc.m_weights(i) = read_bits(b, o, bit_ofs, weight_bits - 1)
            Else
                unpacked.m_astc.m_weights(i) = read_bits(b, o, bit_ofs, weight_bits)
            End If
        Next
    End If

    If blue_contract_check And (total_comps >= 3) Then
        ' We only need to disable ASTC Blue Contraction when we'll be packing to ASTC. The other transcoders don't care.
        invert_subset(0) = False
        invert_subset(1) = False
        invert_subset(2) = False
        any_flag = False

        For subset_index = 0 To subsets - 1
            base = subset_index * total_comps * 2
            s0 = g_astc_unquant_m_unquant(endpoint_range, unpacked.m_astc.m_endpoints(base + 0)) + _
                g_astc_unquant_m_unquant(endpoint_range, unpacked.m_astc.m_endpoints(base + 2)) + _
                g_astc_unquant_m_unquant(endpoint_range, unpacked.m_astc.m_endpoints(base + 4))

            s1 = g_astc_unquant_m_unquant(endpoint_range, unpacked.m_astc.m_endpoints(base + 1)) + _
                g_astc_unquant_m_unquant(endpoint_range, unpacked.m_astc.m_endpoints(base + 3)) + _
                g_astc_unquant_m_unquant(endpoint_range, unpacked.m_astc.m_endpoints(base + 5))

            If s1 < s0 Then
                For c = 0 To total_comps - 1
                    t = unpacked.m_astc.m_endpoints(base + c * 2 + 0)
                    unpacked.m_astc.m_endpoints(base + c * 2 + 0) = unpacked.m_astc.m_endpoints(base + c * 2 + 1)
                    unpacked.m_astc.m_endpoints(base + c * 2 + 1) = t
                Next

                invert_subset(subset_index) = True
                any_flag = True
            End If
        Next

        If any_flag Then
            weight_mask = p2(weight_bits) - 1

            For i = 0 To 15
                subset = pattern(i)

                If invert_subset(subset) Then
                    unpacked.m_astc.m_weights(i * total_planes) = (weight_mask - unpacked.m_astc.m_weights(i * total_planes)) And &HFF

                    If total_planes = 2 Then
                        unpacked.m_astc.m_weights(i * total_planes + 1) = (weight_mask - unpacked.m_astc.m_weights(i * total_planes + 1)) And &HFF
                    End If
                End If
            Next
        End If
    End If

    unpack_uastc_blk = True
End Function

' inline uint32_t astc_interpolate(uint32_t l, uint32_t h, uint32_t w, bool srgb)
Private Function astc_interpolate(ByVal l As Long, ByVal h As Long, ByVal w As Long, ByVal srgb As Boolean) As Long
    If srgb Then
        l = l * 256 Or &H80
        h = h * 256 Or &H80
    Else
        l = l * 256 Or l
        h = h * 256 Or h
    End If

    astc_interpolate = ((l * (64 - w) + h * w + 32) \ 64) \ 256
End Function

' bool unpack_uastc(uint32_t mode, uint32_t common_pattern, const color32& solid_color, const astc_block_desc& astc,
' color32* pPixels, bool srgb), as the unpacked_uastc_block overload calls it: pixels into pp(0 To 63)
Private Function unpack_uastc_pixels(u As unpacked_uastc_block, pp() As Long, ByVal srgb As Boolean) As Boolean
    Static endpoints(0 To 23) As Long, block_colors(0 To 383) As Long, pattern(0 To 15) As Long
    Dim mode As Long, i As Long, total_subsets As Long, total_comps As Long, endpoint_range As Long
    Dim total_planes As Long, weight_bits As Long, weight_levels As Long, subset_index As Long, comp_index As Long
    Dim ll As Long, lh As Long, al As Long, ah As Long, base As Long, l As Long, w As Long, lc As Long, ac As Long
    Dim e As Long, q As Long, weight_index0 As Long, weight_index1 As Long, comp As Long, ccs As Long, k As Long

    mode = u.m_mode
    If mode = UASTC_MODE_INDEX_SOLID_COLOR Then
        For i = 0 To 15
            pp(i * 4 + 0) = u.m_solid_color(0)
            pp(i * 4 + 1) = u.m_solid_color(1)
            pp(i * 4 + 2) = u.m_solid_color(2)
            pp(i * 4 + 3) = u.m_solid_color(3)
        Next
        unpack_uastc_pixels = True
        Exit Function
    End If

    ' color32 endpoints[3][2]: [(subset * 2 + lh) * 4 + comp]
    total_subsets = g_uastc_mode_subsets(mode)
    total_comps = g_uastc_mode_comps(mode)
    If total_comps > 4 Then total_comps = 4
    endpoint_range = g_uastc_mode_endpoint_ranges(mode)
    total_planes = g_uastc_mode_planes(mode)
    weight_bits = g_uastc_mode_weight_bits(mode)
    weight_levels = p2(weight_bits)

    For subset_index = 0 To total_subsets - 1
        e = subset_index * 8
        base = subset_index * total_comps * 2
        If total_comps = 2 Then
            ll = g_astc_unquant_m_unquant(endpoint_range, u.m_astc.m_endpoints(base + 0 * 2 + 0))
            lh = g_astc_unquant_m_unquant(endpoint_range, u.m_astc.m_endpoints(base + 0 * 2 + 1))

            al = g_astc_unquant_m_unquant(endpoint_range, u.m_astc.m_endpoints(base + 1 * 2 + 0))
            ah = g_astc_unquant_m_unquant(endpoint_range, u.m_astc.m_endpoints(base + 1 * 2 + 1))

            endpoints(e + 0) = ll
            endpoints(e + 1) = ll
            endpoints(e + 2) = ll
            endpoints(e + 3) = al
            endpoints(e + 4) = lh
            endpoints(e + 5) = lh
            endpoints(e + 6) = lh
            endpoints(e + 7) = ah
        Else
            For comp_index = 0 To total_comps - 1
                endpoints(e + comp_index) = g_astc_unquant_m_unquant(endpoint_range, u.m_astc.m_endpoints(base + comp_index * 2 + 0))
                endpoints(e + 4 + comp_index) = g_astc_unquant_m_unquant(endpoint_range, u.m_astc.m_endpoints(base + comp_index * 2 + 1))
            Next
            For comp_index = total_comps To 3
                endpoints(e + comp_index) = 255
                endpoints(e + 4 + comp_index) = 255
            Next
        End If
    Next

    ' color32 block_colors[3][32]: [(subset * 32 + level) * 4 + comp]
    For subset_index = 0 To total_subsets - 1
        e = subset_index * 8
        For l = 0 To weight_levels - 1
            w = g_astc_weight_tables(weight_bits, l)
            q = (subset_index * 32 + l) * 4
            If total_comps = 2 Then
                lc = astc_interpolate(endpoints(e + 0), endpoints(e + 4), w, srgb) And &HFF
                ac = astc_interpolate(endpoints(e + 3), endpoints(e + 7), w, srgb) And &HFF

                block_colors(q + 0) = lc
                block_colors(q + 1) = lc
                block_colors(q + 2) = lc
                block_colors(q + 3) = ac
            Else
                For comp_index = 0 To total_comps - 1
                    block_colors(q + comp_index) = astc_interpolate(endpoints(e + comp_index), endpoints(e + 4 + comp_index), w, srgb) And &HFF
                Next

                For comp_index = total_comps To 3
                    block_colors(q + comp_index) = 255
                Next
            End If
        Next
    Next

    If total_subsets >= 2 Then
        If total_subsets = 3 Then
            For i = 0 To 15
                pattern(i) = g_astc_bc7_patterns3(u.m_common_pattern, i)
            Next
        ElseIf mode = 7 Then
            For i = 0 To 15
                pattern(i) = g_bc7_3_astc2_patterns2(u.m_common_pattern, i)
            Next
        Else
            For i = 0 To 15
                pattern(i) = g_astc_bc7_patterns2(u.m_common_pattern, i)
            Next
        End If
    End If

    If total_planes = 1 Then
        If total_subsets = 1 Then
            For i = 0 To 15
                q = u.m_astc.m_weights(i) * 4
                pp(i * 4 + 0) = block_colors(q + 0)
                pp(i * 4 + 1) = block_colors(q + 1)
                pp(i * 4 + 2) = block_colors(q + 2)
                pp(i * 4 + 3) = block_colors(q + 3)
            Next
        Else
            For i = 0 To 15
                q = (pattern(i) * 32 + u.m_astc.m_weights(i)) * 4
                pp(i * 4 + 0) = block_colors(q + 0)
                pp(i * 4 + 1) = block_colors(q + 1)
                pp(i * 4 + 2) = block_colors(q + 2)
                pp(i * 4 + 3) = block_colors(q + 3)
            Next
        End If
    Else
        ccs = u.m_astc.m_ccs
        For i = 0 To 15
            ' const uint32_t subset_index = 0; // pPartition_pattern[i];
            weight_index0 = u.m_astc.m_weights(i * 2) * 4
            weight_index1 = u.m_astc.m_weights(i * 2 + 1) * 4

            For comp = 0 To 3
                If comp = ccs Then
                    k = block_colors(weight_index1 + comp)
                Else
                    k = block_colors(weight_index0 + comp)
                End If
                pp(i * 4 + comp) = k
            Next
        Next
    End If

    unpack_uastc_pixels = True
End Function

' bool unpack_uastc(const uastc_block& blk, color32* pPixels, bool srgb): the block at src(so) into pp(0 To 63)
Public Function unpack_uastc_rgba(src() As Byte, ByVal so As Long, pp() As Long, ByVal srgb As Boolean) As Boolean
    If Not unpack_uastc_blk(src, so, ub, False, False) Then Exit Function

    unpack_uastc_rgba = unpack_uastc_pixels(ub, pp, srgb)
End Function

'--------------------------------------------------------------------------
' ASTC

' static inline uint32_t astc_extract_bits(uint32_t bits, int low, int high): bits [low,high]
Private Function astc_extract_bits(ByVal bits As Long, ByVal low As Long, ByVal high As Long) As Long
    astc_extract_bits = (bits \ p2(low)) And (p2(high - low + 1) - 1)
End Function

' static inline void astc_set_bits(uint32_t* pOutput, int& bit_pos, uint32_t value, uint32_t total_bits) (value < 2^31)
Private Sub astc_set_bits(d() As Byte, ByVal o As Long, bit_pos As Long, ByVal value As Long, ByVal total_bits As Long)
    Dim bits_to_write As Long, s As Long
    Do While total_bits <> 0
        s = bit_pos And 7
        bits_to_write = 8 - s
        If total_bits < bits_to_write Then bits_to_write = total_bits

        d(o + bit_pos \ 8) = d(o + bit_pos \ 8) Or ((value And (255 \ p2(s))) * p2(s))

        bit_pos = bit_pos + bits_to_write
        total_bits = total_bits - bits_to_write
        value = value \ p2(bits_to_write)
    Loop
End Sub

' static inline void astc_set_bits_1_to_9(uint32_t* pDst, int& bit_offset, uint32_t code, uint32_t codesize)
Private Sub astc_set_bits_1_to_9(d() As Byte, ByVal o As Long, bit_offset As Long, ByVal code As Long, ByVal codesize As Long)
    Dim byte_bit_offset As Long, val_ As Long, index As Long
    If codesize <> 0 Then
        byte_bit_offset = bit_offset And 7
        val_ = code * p2(byte_bit_offset)

        index = bit_offset \ 8
        d(o + index) = d(o + index) Or (val_ And &HFF)

        If codesize > (8 - byte_bit_offset) Then d(o + index + 1) = d(o + index + 1) Or ((val_ \ 256) And &HFF)

        bit_offset = bit_offset + codesize
    End If
End Sub

' static void astc_encode_trits(uint32_t* pOutput, const uint8_t* pValues, int& bit_pos, int n)
Private Sub astc_encode_trits(d() As Byte, ByVal o As Long, pValues() As Long, bit_pos As Long, ByVal n As Long)
    Static bits(0 To 4) As Long, s_muls(0 To 4) As Long
    Dim trits As Long, bit_mask As Long, i As Long, t As Long, tt As Long
    s_muls(0) = 1: s_muls(1) = 3: s_muls(2) = 9: s_muls(3) = 27: s_muls(4) = 81

    ' First extract the trits and the bits from the 5 input values
    trits = 0
    bit_mask = p2(n) - 1
    For i = 0 To 4
        t = pValues(i) \ p2(n)

        trits = trits + t * s_muls(i)
        bits(i) = pValues(i) And bit_mask
    Next

    ' Encode the trits, by inverting the bit manipulations done by the decoder, converting 5 trits into 8-bits.
    tt = g_astc_trit_encode(trits)

    ' Now interleave the 8 encoded trit bits with the bits to form the encoded output. See table 94.
    astc_set_bits d, o, bit_pos, bits(0) Or (astc_extract_bits(tt, 0, 1) * p2(n)) Or (bits(1) * p2(2 + n)), n * 2 + 2

    astc_set_bits d, o, bit_pos, astc_extract_bits(tt, 2, 3) Or (bits(2) * 4) Or (astc_extract_bits(tt, 4, 4) * p2(2 + n)) Or _
        (bits(3) * p2(3 + n)) Or (astc_extract_bits(tt, 5, 6) * p2(3 + n * 2)) Or (bits(4) * p2(5 + n * 2)) Or _
        (astc_extract_bits(tt, 7, 7) * p2(5 + n * 3)), n * 3 + 6
End Sub

' static inline void astc_encode_quints(uint32_t* pOutput, const uint8_t* pValues, int& bit_pos, int n)
Private Sub astc_encode_quints(d() As Byte, ByVal o As Long, pValues() As Long, bit_pos As Long, ByVal n As Long)
    Static bits(0 To 2) As Long, s_muls(0 To 2) As Long
    Dim quints As Long, bit_mask As Long, i As Long, t As Long, tt As Long
    s_muls(0) = 1: s_muls(1) = 5: s_muls(2) = 25

    ' First extract the quints and the bits from the 3 input values
    quints = 0
    bit_mask = p2(n) - 1
    For i = 0 To 2
        t = pValues(i) \ p2(n)

        quints = quints + t * s_muls(i)
        bits(i) = pValues(i) And bit_mask
    Next

    ' Encode the quints, by inverting the bit manipulations done by the decoder, converting 3 quints into 7-bits.
    tt = g_astc_quint_encode(quints)

    ' Now interleave the 7 encoded quint bits with the bits to form the encoded output. See table 95-96.
    astc_set_bits d, o, bit_pos, bits(0) Or (astc_extract_bits(tt, 0, 2) * p2(n)) Or (bits(1) * p2(3 + n)) Or _
        (astc_extract_bits(tt, 3, 4) * p2(3 + n * 2)) Or (bits(2) * p2(5 + n * 2)) Or (astc_extract_bits(tt, 5, 6) * p2(5 + n * 3)), 7 + n * 3
End Sub

' static void astc_pack_bise(uint32_t* pDst, const uint8_t* pSrc_vals, int bit_pos, int num_vals, int range)
Private Sub astc_pack_bise(d() As Byte, ByVal o As Long, pSrc_vals() As Long, ByVal bit_pos As Long, ByVal num_vals As Long, ByVal range As Long)
    Static temp(0 To 19) As Byte, vals(0 To 4) As Long
    Dim num_bits As Long, group_size As Long, total_groups As Long, group_index As Long, limit As Long, i As Long

    ' uint32_t temp[5] = { 0, 0, 0, 0, 0 }
    For i = 0 To 19
        temp(i) = 0
    Next

    num_bits = g_astc_bise_range_table(range, 0)

    group_size = 0
    If g_astc_bise_range_table(range, 1) <> 0 Then
        group_size = 5
    ElseIf g_astc_bise_range_table(range, 2) <> 0 Then
        group_size = 3
    End If

    If group_size <> 0 Then
        ' Range has trits or quints - pack each group of 5 or 3 values
        If group_size = 5 Then total_groups = (num_vals + 4) \ 5 Else total_groups = (num_vals + 2) \ 3

        For group_index = 0 To total_groups - 1
            For i = 0 To 4
                vals(i) = 0
            Next

            limit = num_vals - group_index * group_size
            If group_size < limit Then limit = group_size
            For i = 0 To limit - 1
                vals(i) = pSrc_vals(group_index * group_size + i)
            Next

            If group_size = 5 Then
                astc_encode_trits temp, 0, vals, bit_pos, num_bits
            Else
                astc_encode_quints temp, 0, vals, bit_pos, num_bits
            End If
        Next
    Else
        For i = 0 To num_vals - 1
            astc_set_bits_1_to_9 temp, 0, bit_pos, pSrc_vals(i), num_bits
        Next
    End If

    ' pDst[0] |= temp[0]; ... pDst[3] |= temp[3];
    For i = 0 To 15
        d(o + i) = d(o + i) Or temp(i)
    Next
End Sub

' bool pack_astc_block(uint32_t* pDst, const astc_block_desc* pBlock, uint32_t uastc_mode)
Private Function pack_astc_block(d() As Byte, ByVal o As Long, blk As astc_block_desc, ByVal uastc_mode As Long) As Boolean
    Static s_reverse_bits2(0 To 3) As Long, s_reverse_bits3(0 To 7) As Long, s_reverse_bits4(0 To 15) As Long, s_reverse_bits5(0 To 31) As Long
    Static tables_set As Boolean
    Dim total_weights As Long, mode As Long, i As Long, bit_pos As Long, bits_per_weight As Long
    Dim total_weight_bits As Long, ccs_bit_pos As Long, num_cem_pairs As Long, ofs As Long, rev As Long, index As Long

    If Not tables_set Then
        set_table s_reverse_bits2, "0,2,1,3"
        set_table s_reverse_bits3, "0,4,2,6,1,5,3,7"
        set_table s_reverse_bits4, "0,8,4,12,2,10,6,14,1,9,5,13,3,11,7,15"
        set_table s_reverse_bits5, "0,16,8,24,4,20,12,28,2,18,10,26,6,22,14,30,1,17,9,25,5,21,13,29,3,19,11,27,7,23,15,31"
        tables_set = True
    End If

    If blk.m_dual_plane Then total_weights = 32 Else total_weights = 16

    ' Set mode bits - see Table 146-147
    mode = g_uastc_mode_astc_block_mode(uastc_mode)
    d(o + 0) = mode And &HFF
    d(o + 1) = (mode \ 256) And &HFF

    For i = 2 To 15
        d(o + i) = 0
    Next

    bit_pos = ASTC_BLOCK_MODE_BITS

    ' We only support 1-5 bit weight indices
    bits_per_weight = g_astc_bise_range_table(blk.m_weight_range, 0)

    ' See table 143 - PART
    astc_set_bits_1_to_9 d, o, bit_pos, blk.m_subsets - 1, ASTC_PART_BITS

    If blk.m_subsets = 1 Then
        astc_set_bits_1_to_9 d, o, bit_pos, blk.m_cem, ASTC_CEM_BITS
    Else
        ' See table 145
        astc_set_bits d, o, bit_pos, blk.m_partition_seed, ASTC_PARTITION_INDEX_BITS

        ' Table 150 - we assume all CEM's are equal, so write 2 0's along with the CEM
        astc_set_bits_1_to_9 d, o, bit_pos, (blk.m_cem * 4) And 63, ASTC_CEM_BITS + 2
    End If

    If blk.m_dual_plane Then
        total_weight_bits = total_weights * bits_per_weight

        ' See Illegal Encodings 23.24
        ccs_bit_pos = 128 - total_weight_bits - ASTC_CCS_BITS
        astc_set_bits_1_to_9 d, o, ccs_bit_pos, blk.m_ccs, ASTC_CCS_BITS
    End If

    num_cem_pairs = (1 + blk.m_cem \ 4) * blk.m_subsets

    astc_pack_bise d, o, blk.m_endpoints, bit_pos, num_cem_pairs * 2, g_uastc_mode_endpoint_ranges(uastc_mode)

    ' Write the weight bits in reverse bit order.
    Select Case bits_per_weight
    Case 1
        For i = 0 To total_weights - 1
            ofs = 128 - 1 - i
            d(o + ofs \ 8) = d(o + ofs \ 8) Or ((blk.m_weights(i) * p2(ofs And 7)) And &HFF)
        Next
    Case 2
        For i = 0 To total_weights - 1
            ofs = 128 - 2 - (i * 2)
            d(o + ofs \ 8) = d(o + ofs \ 8) Or ((s_reverse_bits2(blk.m_weights(i)) * p2(ofs And 7)) And &HFF)
        Next
    Case 3
        For i = 0 To total_weights - 1
            ofs = 128 - 3 - (i * 3)
            rev = s_reverse_bits3(blk.m_weights(i)) * p2(ofs And 7)

            index = ofs \ 8
            d(o + index) = d(o + index) Or (rev And &HFF)
            index = index + 1
            If index < 16 Then
                d(o + index) = d(o + index) Or ((rev \ 256) And &HFF)
            End If
        Next
    Case 4
        For i = 0 To total_weights - 1
            ofs = 128 - 4 - (i * 4)
            d(o + ofs \ 8) = d(o + ofs \ 8) Or ((s_reverse_bits4(blk.m_weights(i)) * p2(ofs And 7)) And &HFF)
        Next
    Case 5
        For i = 0 To total_weights - 1
            ofs = 128 - 5 - (i * 5)
            rev = s_reverse_bits5(blk.m_weights(i)) * p2(ofs And 7)

            index = ofs \ 8
            d(o + index) = d(o + index) Or (rev And &HFF)
            index = index + 1
            If index < 16 Then
                d(o + index) = d(o + index) Or ((rev \ 256) And &HFF)
            End If
        Next
    End Select

    pack_astc_block = True
End Function

' (VB) fills a small Static table from its C initialiser's values
Private Sub set_table(t() As Long, ByVal s As String)
    Dim parts() As String, i As Long
    parts = Split(s, ",")
    For i = 0 To UBound(parts)
        t(i) = CLng(parts(i))
    Next
End Sub

' void pack_astc_solid_block(void* pDst_block, const color32& color)
Private Sub pack_astc_solid_block(d() As Byte, ByVal o As Long, color() As Long)
    Dim r As Long, g As Long, b As Long, a As Long, i As Long, bit_pos As Long
    r = color(0)
    g = color(1)
    b = color(2)
    a = color(3)

    d(o + 0) = &HFC
    d(o + 1) = &HFD
    d(o + 2) = &HFF
    d(o + 3) = &HFF

    ' pOutput[1] = 0xffffffff; pOutput[2] = 0; pOutput[3] = 0;
    For i = 4 To 7
        d(o + i) = &HFF
    Next
    For i = 8 To 15
        d(o + i) = 0
    Next

    bit_pos = 64
    astc_set_bits d, o, bit_pos, r Or (r * 256), 16
    astc_set_bits d, o, bit_pos, g Or (g * 256), 16
    astc_set_bits d, o, bit_pos, b Or (b * 256), 16
    astc_set_bits d, o, bit_pos, a Or (a * 256), 16
End Sub

' bool transcode_uastc_to_astc(const uastc_block& src_blk, void* pDst)
Public Function transcode_uastc_to_astc(src() As Byte, ByVal so As Long, d() As Byte, ByVal o As Long) As Boolean
    Dim success As Boolean
    If Not unpack_uastc_blk(src, so, ub, True, False) Then Exit Function

    success = False
    If ub.m_mode = UASTC_MODE_INDEX_SOLID_COLOR Then
        pack_astc_solid_block d, o, ub.m_solid_color
        success = True
    Else
        success = pack_astc_block(d, o, ub.m_astc, ub.m_mode)
    End If

    transcode_uastc_to_astc = success
End Function

'--------------------------------------------------------------------------
' BC7

' static inline void bc7_set_block_bits(uint8_t* pBytes, uint32_t val, uint32_t num_bits, uint32_t* pCur_ofs)
Private Sub bc7_set_block_bits(d() As Byte, ByVal o As Long, ByVal val_ As Long, ByVal num_bits As Long, cur_ofs As Long)
    Dim n As Long, s As Long
    Do While num_bits <> 0
        s = cur_ofs And 7
        n = 8 - s
        If num_bits < n Then n = num_bits
        d(o + cur_ofs \ 8) = d(o + cur_ofs \ 8) Or ((val_ * p2(s)) And &HFF)
        val_ = val_ \ p2(n)
        num_bits = num_bits - n
        cur_ofs = cur_ofs + n
    Loop
End Sub

' void encode_bc7_block(void* pBlock, const bc7_optimization_results* pResults)
Private Sub encode_bc7_block(d() As Byte, ByVal o As Long, r As bc7_optimization_results)
    Static color_selectors(0 To 15) As Long, alpha_selectors(0 To 15) As Long, low(0 To 11) As Long, high(0 To 11) As Long
    Static pbits(0 To 5) As Long, anchor(0 To 2) As Long, part(0 To 15) As Long
    Dim best_mode As Long, total_subsets As Long, total_partitions As Long, i As Long, k As Long, q As Long
    Dim anchor_index As Long, color_index_bits As Long, num_color_indices As Long, alpha_index_bits As Long, num_alpha_indices As Long
    Dim t As Long, sep_alpha As Boolean, cur_bit_ofs As Long, total_comps As Long, comp As Long, subset As Long, prec As Long
    Dim idx As Long, n As Long, color_size As Long, alpha_size As Long

    best_mode = r.m_mode

    total_subsets = g_bc7_num_subsets(best_mode)
    total_partitions = p2(g_bc7_partition_bits(best_mode))
    sep_alpha = (best_mode = 4) Or (best_mode = 5)        ' get_bc7_mode_has_seperate_alpha_selectors

    If total_subsets = 1 Then
        For i = 0 To 15
            part(i) = g_bc7_partition1(i)
        Next
    ElseIf total_subsets = 2 Then
        For i = 0 To 15
            part(i) = g_bc7_partition2(r.m_partition * 16 + i)
        Next
    Else
        For i = 0 To 15
            part(i) = g_bc7_partition3(r.m_partition * 16 + i)
        Next
    End If

    For i = 0 To 15
        color_selectors(i) = r.m_selectors(i)
        alpha_selectors(i) = r.m_alpha_selectors(i)
    Next
    For i = 0 To 11
        low(i) = r.m_low(i)
        high(i) = r.m_high(i)
    Next
    For i = 0 To 5
        pbits(i) = r.m_pbits(i)
    Next

    anchor(0) = -1
    anchor(1) = -1
    anchor(2) = -1

    color_size = g_bc7_color_index_bitcount(best_mode) + r.m_index_selector       ' get_bc7_color_index_size
    alpha_size = g_bc7_alpha_index_bitcount(best_mode) - r.m_index_selector       ' get_bc7_alpha_index_size

    For k = 0 To total_subsets - 1
        anchor_index = 0
        If k <> 0 Then
            If (total_subsets = 3) And (k = 1) Then
                anchor_index = g_bc7_table_anchor_index_third_subset_1(r.m_partition)
            ElseIf (total_subsets = 3) And (k = 2) Then
                anchor_index = g_bc7_table_anchor_index_third_subset_2(r.m_partition)
            Else
                anchor_index = g_bc7_table_anchor_index_second_subset(r.m_partition)
            End If
        End If

        anchor(k) = anchor_index

        color_index_bits = color_size
        num_color_indices = p2(color_index_bits)

        If (color_selectors(anchor_index) And (num_color_indices \ 2)) <> 0 Then
            For i = 0 To 15
                If part(i) = k Then color_selectors(i) = ((num_color_indices - 1) - color_selectors(i)) And &HFF
            Next

            If sep_alpha Then
                For q = 0 To 2
                    t = low(k * 4 + q)
                    low(k * 4 + q) = high(k * 4 + q)
                    high(k * 4 + q) = t
                Next
            Else
                For q = 0 To 3
                    t = low(k * 4 + q)
                    low(k * 4 + q) = high(k * 4 + q)
                    high(k * 4 + q) = t
                Next
            End If

            If g_bc7_mode_has_shared_p_bits(best_mode) = 0 Then
                t = pbits(k * 2 + 0)
                pbits(k * 2 + 0) = pbits(k * 2 + 1)
                pbits(k * 2 + 1) = t
            End If
        End If

        If sep_alpha Then
            alpha_index_bits = alpha_size
            num_alpha_indices = p2(alpha_index_bits)

            If (alpha_selectors(anchor_index) And (num_alpha_indices \ 2)) <> 0 Then
                For i = 0 To 15
                    If part(i) = k Then alpha_selectors(i) = ((num_alpha_indices - 1) - alpha_selectors(i)) And &HFF
                Next

                t = low(k * 4 + 3)
                low(k * 4 + 3) = high(k * 4 + 3)
                high(k * 4 + 3) = t
            End If
        End If
    Next

    For i = 0 To 15
        d(o + i) = 0
    Next

    cur_bit_ofs = 0
    bc7_set_block_bits d, o, p2(best_mode), best_mode + 1, cur_bit_ofs

    If (best_mode = 4) Or (best_mode = 5) Then bc7_set_block_bits d, o, r.m_rotation, 2, cur_bit_ofs

    If best_mode = 4 Then bc7_set_block_bits d, o, r.m_index_selector, 1, cur_bit_ofs

    If total_partitions > 1 Then
        If total_partitions = 64 Then n = 6 Else n = 4
        bc7_set_block_bits d, o, r.m_partition, n, cur_bit_ofs
    End If

    If best_mode >= 4 Then total_comps = 4 Else total_comps = 3
    For comp = 0 To total_comps - 1
        If comp = 3 Then prec = g_bc7_alpha_precision_table(best_mode) Else prec = g_bc7_color_precision_table(best_mode)
        For subset = 0 To total_subsets - 1
            bc7_set_block_bits d, o, low(subset * 4 + comp), prec, cur_bit_ofs
            bc7_set_block_bits d, o, high(subset * 4 + comp), prec, cur_bit_ofs
        Next
    Next

    If g_bc7_mode_has_p_bits(best_mode) <> 0 Then
        For subset = 0 To total_subsets - 1
            bc7_set_block_bits d, o, pbits(subset * 2 + 0), 1, cur_bit_ofs
            If g_bc7_mode_has_shared_p_bits(best_mode) = 0 Then bc7_set_block_bits d, o, pbits(subset * 2 + 1), 1, cur_bit_ofs
        Next
    End If

    For idx = 0 To 15
        If r.m_index_selector <> 0 Then n = alpha_size Else n = color_size

        If (idx = anchor(0)) Or (idx = anchor(1)) Or (idx = anchor(2)) Then n = n - 1

        If r.m_index_selector <> 0 Then
            bc7_set_block_bits d, o, alpha_selectors(idx), n, cur_bit_ofs
        Else
            bc7_set_block_bits d, o, color_selectors(idx), n, cur_bit_ofs
        End If
    Next

    If sep_alpha Then
        For idx = 0 To 15
            If r.m_index_selector <> 0 Then n = color_size Else n = alpha_size

            If (idx = anchor(0)) Or (idx = anchor(1)) Or (idx = anchor(2)) Then n = n - 1

            If r.m_index_selector <> 0 Then
                bc7_set_block_bits d, o, color_selectors(idx), n, cur_bit_ofs
            Else
                bc7_set_block_bits d, o, alpha_selectors(idx), n, cur_bit_ofs
            End If
        Next
    End If
End Sub

' static void determine_shared_pbits(uint32_t total_comps, uint32_t comp_bits, float xl[4], float xh[4],
'     color_quad_u8& bestMinColor, color_quad_u8& bestMaxColor, uint32_t best_pbits[2])
Private Sub determine_shared_pbits(ByVal total_comps As Long, ByVal comp_bits As Long, xl() As Single, xh() As Single, bestMinColor() As Long, bestMaxColor() As Long, best_pbits() As Long)
    Static xMinColor(0 To 3) As Long, xMaxColor(0 To 3) As Long, scaledLow(0 To 3) As Long, scaledHigh(0 To 3) As Long
    Dim total_bits As Long, iscalep As Long, scalep As Single, best_err As Single, p As Long, c As Long, i As Long, j As Long
    Dim err As Single, a As Single, b As Single

    total_bits = comp_bits + 1

    iscalep = p2(total_bits) - 1
    scalep = iscalep

    best_err = 1000000000!

    For p = 0 To 1
        For c = 0 To 3
            xMinColor(c) = clampi(Fix(fadd_(fdiv_(fsub_(fmul_(xl(c), scalep), p), 2!), 0.5!)) * 2 + p, p, iscalep - 1 + p) And &HFF
            xMaxColor(c) = clampi(Fix(fadd_(fdiv_(fsub_(fmul_(xh(c), scalep), p), 2!), 0.5!)) * 2 + p, p, iscalep - 1 + p) And &HFF
        Next

        For i = 0 To 3
            scaledLow(i) = (xMinColor(i) * p2(8 - total_bits)) And &HFF
            scaledLow(i) = scaledLow(i) Or (scaledLow(i) \ p2(total_bits))

            scaledHigh(i) = (xMaxColor(i) * p2(8 - total_bits)) And &HFF
            scaledHigh(i) = scaledHigh(i) Or (scaledHigh(i) \ p2(total_bits))
        Next

        err = 0
        For i = 0 To total_comps - 1
            a = fsub_(fdiv_(scaledLow(i), 255!), xl(i))
            b = fsub_(fdiv_(scaledHigh(i), 255!), xh(i))
            err = fadd_(err, fadd_(fmul_(a, a), fmul_(b, b)))
        Next

        If err < best_err Then
            best_err = err
            best_pbits(0) = p
            best_pbits(1) = p
            For j = 0 To 3
                bestMinColor(j) = xMinColor(j) \ 2
                bestMaxColor(j) = xMaxColor(j) \ 2
            Next
        End If
    Next
End Sub

' static void determine_unique_pbits(uint32_t total_comps, uint32_t comp_bits, float xl[4], float xh[4],
'     color_quad_u8& bestMinColor, color_quad_u8& bestMaxColor, uint32_t best_pbits[2])
Private Sub determine_unique_pbits(ByVal total_comps As Long, ByVal comp_bits As Long, xl() As Single, xh() As Single, bestMinColor() As Long, bestMaxColor() As Long, best_pbits() As Long)
    Static xMinColor(0 To 3) As Long, xMaxColor(0 To 3) As Long, scaledLow(0 To 3) As Long, scaledHigh(0 To 3) As Long
    Dim total_bits As Long, iscalep As Long, scalep As Single, best_err0 As Single, best_err1 As Single, p As Long
    Dim c As Long, i As Long, err0 As Single, err1 As Single, a As Single

    total_bits = comp_bits + 1
    iscalep = p2(total_bits) - 1
    scalep = iscalep

    best_err0 = 1000000000!
    best_err1 = 1000000000!

    For p = 0 To 1
        For c = 0 To 3
            xMinColor(c) = clampi(Fix(fadd_(fdiv_(fsub_(fmul_(xl(c), scalep), p), 2!), 0.5!)) * 2 + p, p, iscalep - 1 + p) And &HFF
            xMaxColor(c) = clampi(Fix(fadd_(fdiv_(fsub_(fmul_(xh(c), scalep), p), 2!), 0.5!)) * 2 + p, p, iscalep - 1 + p) And &HFF
        Next

        For i = 0 To 3
            scaledLow(i) = (xMinColor(i) * p2(8 - total_bits)) And &HFF
            scaledLow(i) = scaledLow(i) Or (scaledLow(i) \ p2(total_bits))

            scaledHigh(i) = (xMaxColor(i) * p2(8 - total_bits)) And &HFF
            scaledHigh(i) = scaledHigh(i) Or (scaledHigh(i) \ p2(total_bits))
        Next

        err0 = 0
        err1 = 0
        For i = 0 To total_comps - 1
            a = fsub_(scaledLow(i), fmul_(xl(i), 255!))
            err0 = fadd_(err0, fmul_(a, a))
            a = fsub_(scaledHigh(i), fmul_(xh(i), 255!))
            err1 = fadd_(err1, fmul_(a, a))
        Next

        If err0 < best_err0 Then
            best_err0 = err0
            best_pbits(0) = p

            bestMinColor(0) = xMinColor(0) \ 2
            bestMinColor(1) = xMinColor(1) \ 2
            bestMinColor(2) = xMinColor(2) \ 2
            bestMinColor(3) = xMinColor(3) \ 2
        End If

        If err1 < best_err1 Then
            best_err1 = err1
            best_pbits(1) = p

            bestMaxColor(0) = xMaxColor(0) \ 2
            bestMaxColor(1) = xMaxColor(1) \ 2
            bestMaxColor(2) = xMaxColor(2) \ 2
            bestMaxColor(3) = xMaxColor(3) \ 2
        End If
    Next
End Sub

' uint32_t bc7_convert_partition_index_3_to_2(uint32_t p, uint32_t k)
Private Function bc7_convert_partition_index_3_to_2(ByVal p As Long, ByVal k As Long) As Long
    Select Case k \ 2
    Case 0
        If p <= 1 Then p = 0 Else p = 1
    Case 1
        If p = 0 Then p = 0 Else p = 1
    Case 2
        If (p = 0) Or (p = 2) Then p = 0 Else p = 1
    End Select
    If (k And 1) <> 0 Then p = 1 - p
    bc7_convert_partition_index_3_to_2 = p
End Function

' (VB) xl/xh as C fills them: g_astc_unquant[endpoint_range][endpoint].m_unquant / 255.0f
Private Function unq_f(ByVal endpoint_range As Long, ByVal e As Long) As Single
    unq_f = fdiv_(g_astc_unquant_m_unquant(endpoint_range, e), 255!)
End Function

' bool transcode_uastc_to_bc7(const unpacked_uastc_block& unpacked_src_blk, bc7_optimization_results& dst_blk)
Private Function transcode_uastc_to_bc7_res(u As unpacked_uastc_block, dst_blk As bc7_optimization_results) As Boolean
    Static xl(0 To 3) As Single, xh(0 To 3) As Single, bestMinColor(0 To 3) As Long, bestMaxColor(0 To 3) As Long
    Static best_pbits(0 To 1) As Long, s_bc7_5_to_4(0 To 31) As Long, s_bc7_2_to_4(0 To 3) As Long, s_bc7_3_to_4(0 To 7) As Long
    Static tables_set As Boolean
    Dim mode As Long, endpoint_range As Long, total_comps As Long, i As Long, subset As Long, comp As Long, invert_partition As Boolean
    Dim v As Long, bc7_subset_index As Long, perm As Long, lo As Long, hi As Long, astc_comp As Long, bc7_comp As Long
    Dim l As Long, h As Long, common_pattern_k As Long, bc7_part As Long, astc_part As Long, c As Long
    Dim best_err0 As Long, best_err1 As Long, best_p As Long, astc_subset As Long, sc As Long, b4 As Long, b8 As Long

    If Not tables_set Then
        set_table s_bc7_5_to_4, "0,0,1,1,2,2,3,3,4,4,5,5,6,6,6,7,8,9,9,9,10,10,11,11,12,12,13,13,14,14,15,15"
        set_table s_bc7_2_to_4, "0,5,10,15"
        set_table s_bc7_3_to_4, "0,2,4,6,9,11,13,15"
        tables_set = True
    End If

    dst_blk = zero_bc7                                  ' memset(&dst_blk, 0, sizeof(dst_blk))

    mode = u.m_mode

    endpoint_range = g_uastc_mode_endpoint_ranges(mode)
    total_comps = g_uastc_mode_comps(mode)

    Select Case mode
    Case 0, 5, 10, 12, 14, 15, 18
        ' MODE 0: DualPlane: 0, WeightRange: 8 (16), Subsets: 1, EndpointRange: 19 (192) - BC7 MODE6 RGB
        ' MODE 5: DualPlane: 0, WeightRange : 5 (8), Subsets : 1, EndpointRange : 20 (256) - BC7 MODE6 RGB
        ' MODE 10 DualPlane: 0, WeightRange: 8 (16), Subsets: 1, EndpointRange: 13 (48) - BC7 MODE6
        ' MODE 12: DualPlane: 0, WeightRange : 5 (8), Subsets : 1, EndpointRange : 19 (192) - BC7 MODE6
        ' MODE 14: DualPlane: 0, WeightRange : 2 (4), Subsets : 1, EndpointRange : 20 (256) - BC7 MODE6
        ' MODE 18: DualPlane: 0, WeightRange : 11 (32), Subsets : 1, CEM : 8, EndpointRange : 11 (32) - BC7 MODE6
        ' MODE 15: DualPlane: 0, WeightRange : 8 (16), Subsets : 1, CEM : 4 (LA Direct), EndpointRange : 20 (256) - BC7 MODE6
        dst_blk.m_mode = 6

        If total_comps = 2 Then
            xl(0) = unq_f(endpoint_range, u.m_astc.m_endpoints(0))
            xh(0) = unq_f(endpoint_range, u.m_astc.m_endpoints(1))

            xl(1) = xl(0)
            xh(1) = xh(0)

            xl(2) = xl(0)
            xh(2) = xh(0)

            xl(3) = unq_f(endpoint_range, u.m_astc.m_endpoints(2))
            xh(3) = unq_f(endpoint_range, u.m_astc.m_endpoints(3))
        Else
            xl(0) = unq_f(endpoint_range, u.m_astc.m_endpoints(0))
            xl(1) = unq_f(endpoint_range, u.m_astc.m_endpoints(2))
            xl(2) = unq_f(endpoint_range, u.m_astc.m_endpoints(4))

            xh(0) = unq_f(endpoint_range, u.m_astc.m_endpoints(1))
            xh(1) = unq_f(endpoint_range, u.m_astc.m_endpoints(3))
            xh(2) = unq_f(endpoint_range, u.m_astc.m_endpoints(5))

            If total_comps = 4 Then
                xl(3) = unq_f(endpoint_range, u.m_astc.m_endpoints(6))
                xh(3) = unq_f(endpoint_range, u.m_astc.m_endpoints(7))
            Else
                xl(3) = 1!
                xh(3) = 1!
            End If
        End If

        If total_comps = 2 Then sc = 4 Else sc = total_comps
        determine_unique_pbits sc, 7, xl, xh, bestMinColor, bestMaxColor, best_pbits

        For c = 0 To 3
            dst_blk.m_low(c) = bestMinColor(c)
            dst_blk.m_high(c) = bestMaxColor(c)
        Next

        If total_comps = 3 Then
            dst_blk.m_low(3) = 127
            dst_blk.m_high(3) = 127
        End If

        dst_blk.m_pbits(0) = best_pbits(0)
        dst_blk.m_pbits(1) = best_pbits(1)

        If mode = 18 Then
            For i = 0 To 15
                dst_blk.m_selectors(i) = s_bc7_5_to_4(u.m_astc.m_weights(i))
            Next
        ElseIf mode = 14 Then
            For i = 0 To 15
                dst_blk.m_selectors(i) = s_bc7_2_to_4(u.m_astc.m_weights(i))
            Next
        ElseIf (mode = 5) Or (mode = 12) Then
            For i = 0 To 15
                dst_blk.m_selectors(i) = s_bc7_3_to_4(u.m_astc.m_weights(i))
            Next
        Else
            For i = 0 To 15
                dst_blk.m_selectors(i) = u.m_astc.m_weights(i)
            Next
        End If

    Case 1
        ' DualPlane: 0, WeightRange : 2 (4), Subsets : 1, EndpointRange : 20 (256) - BC7 MODE3
        ' Mode 1 uses endpoint range 20 - no need to use ASTC dequant tables.
        dst_blk.m_mode = 3

        xl(0) = fdiv_(u.m_astc.m_endpoints(0), 255!)
        xl(1) = fdiv_(u.m_astc.m_endpoints(2), 255!)
        xl(2) = fdiv_(u.m_astc.m_endpoints(4), 255!)
        xl(3) = 1!

        xh(0) = fdiv_(u.m_astc.m_endpoints(1), 255!)
        xh(1) = fdiv_(u.m_astc.m_endpoints(3), 255!)
        xh(2) = fdiv_(u.m_astc.m_endpoints(5), 255!)
        xh(3) = 1!

        For c = 0 To 3
            bestMinColor(c) = 0
            bestMaxColor(c) = 0
        Next
        determine_unique_pbits 3, 7, xl, xh, bestMinColor, bestMaxColor, best_pbits

        For i = 0 To 2
            dst_blk.m_low(0 * 4 + i) = bestMinColor(i)
            dst_blk.m_high(0 * 4 + i) = bestMaxColor(i)
            dst_blk.m_low(1 * 4 + i) = bestMinColor(i)
            dst_blk.m_high(1 * 4 + i) = bestMaxColor(i)
        Next
        dst_blk.m_pbits(0) = best_pbits(0)
        dst_blk.m_pbits(1) = best_pbits(1)
        dst_blk.m_pbits(2) = best_pbits(0)
        dst_blk.m_pbits(3) = best_pbits(1)

        For i = 0 To 15
            dst_blk.m_selectors(i) = u.m_astc.m_weights(i)
        Next

    Case 2
        ' 2. DualPlane: 0, WeightRange : 5 (8), Subsets : 2, EndpointRange : 8 (16) - BC7 MODE1
        dst_blk.m_mode = 1
        dst_blk.m_partition = g_astc_bc7_common_partitions2_m_bc7(u.m_common_pattern)

        invert_partition = g_astc_bc7_common_partitions2_m_invert(u.m_common_pattern) <> 0

        xl(3) = 1!
        xh(3) = 1!

        For subset = 0 To 1
            For i = 0 To 2
                v = u.m_astc.m_endpoints(i * 2 + subset * 6)
                v = (v * 16) Or v
                xl(i) = fdiv_(v, 255!)

                v = u.m_astc.m_endpoints(i * 2 + subset * 6 + 1)
                v = (v * 16) Or v
                xh(i) = fdiv_(v, 255!)
            Next

            best_pbits(0) = 0
            best_pbits(1) = 0
            For c = 0 To 3
                bestMinColor(c) = 0
                bestMaxColor(c) = 0
            Next
            determine_shared_pbits 3, 6, xl, xh, bestMinColor, bestMaxColor, best_pbits

            If invert_partition Then bc7_subset_index = 1 - subset Else bc7_subset_index = subset

            For i = 0 To 2
                dst_blk.m_low(bc7_subset_index * 4 + i) = bestMinColor(i)
                dst_blk.m_high(bc7_subset_index * 4 + i) = bestMaxColor(i)
            Next

            dst_blk.m_pbits(bc7_subset_index * 2 + 0) = best_pbits(0)
        Next

        For i = 0 To 15
            dst_blk.m_selectors(i) = u.m_astc.m_weights(i)
        Next

    Case 3
        ' DualPlane: 0, WeightRange : 2 (4), Subsets : 3, EndpointRange : 7 (12) - BC7 MODE2
        dst_blk.m_mode = 2
        dst_blk.m_partition = g_astc_bc7_common_partitions3_m_bc7(u.m_common_pattern)

        perm = g_astc_bc7_common_partitions3_m_astc_to_bc7_perm(u.m_common_pattern)

        For subset = 0 To 2
            For comp = 0 To 2
                lo = g_astc_unquant_m_unquant(endpoint_range, u.m_astc.m_endpoints(comp * 2 + 0 + subset * 6))
                hi = g_astc_unquant_m_unquant(endpoint_range, u.m_astc.m_endpoints(comp * 2 + 1 + subset * 6))

                ' TODO: I think this can be improved by using tables like Basis Universal does with ETC1S conversion.
                lo = (lo * 31 + 127) \ 255
                hi = (hi * 31 + 127) \ 255

                bc7_subset_index = g_astc_to_bc7_partition_index_perm_tables(perm, subset)

                dst_blk.m_low(bc7_subset_index * 4 + comp) = lo
                dst_blk.m_high(bc7_subset_index * 4 + comp) = hi
            Next
        Next

        For i = 0 To 15
            dst_blk.m_selectors(i) = u.m_astc.m_weights(i)
        Next

    Case 4
        ' 4. DualPlane: 0, WeightRange: 2 (4), Subsets: 2, EndpointRange: 12 (40) - BC7 MODE3
        dst_blk.m_mode = 3
        dst_blk.m_partition = g_astc_bc7_common_partitions2_m_bc7(u.m_common_pattern)

        invert_partition = g_astc_bc7_common_partitions2_m_invert(u.m_common_pattern) <> 0

        xl(3) = 1!
        xh(3) = 1!

        For subset = 0 To 1
            For i = 0 To 2
                xl(i) = unq_f(endpoint_range, u.m_astc.m_endpoints(i * 2 + subset * 6))
                xh(i) = unq_f(endpoint_range, u.m_astc.m_endpoints(i * 2 + subset * 6 + 1))
            Next

            best_pbits(0) = 0
            best_pbits(1) = 0
            For c = 0 To 3
                bestMinColor(c) = 0
                bestMaxColor(c) = 0
            Next
            determine_unique_pbits 3, 7, xl, xh, bestMinColor, bestMaxColor, best_pbits

            If invert_partition Then bc7_subset_index = 1 - subset Else bc7_subset_index = subset

            For i = 0 To 2
                dst_blk.m_low(bc7_subset_index * 4 + i) = bestMinColor(i)
                dst_blk.m_high(bc7_subset_index * 4 + i) = bestMaxColor(i)
            Next
            dst_blk.m_low(bc7_subset_index * 4 + 3) = 127
            dst_blk.m_high(bc7_subset_index * 4 + 3) = 127

            dst_blk.m_pbits(bc7_subset_index * 2 + 0) = best_pbits(0)
            dst_blk.m_pbits(bc7_subset_index * 2 + 1) = best_pbits(1)
        Next

        For i = 0 To 15
            dst_blk.m_selectors(i) = u.m_astc.m_weights(i)
        Next

    Case 6, 11, 13, 17
        ' MODE 6: DualPlane: 1, WeightRange : 2 (4), Subsets : 1, EndpointRange : 18 (160) - BC7 MODE5 RGB
        ' MODE 11: DualPlane: 1, WeightRange: 2 (4), Subsets: 1, EndpointRange: 13 (48) - BC7 MODE5
        ' MODE 13: DualPlane: 1, WeightRange: 0 (2), Subsets : 1, EndpointRange : 20 (256) - BC7 MODE5
        ' MODE 17: DualPlane: 1, WeightRange: 2 (4), Subsets: 1, CEM: 4 (LA Direct), EndpointRange: 20 (256) - BC7 MODE5
        dst_blk.m_mode = 5
        dst_blk.m_rotation = (u.m_astc.m_ccs + 1) And 3

        If total_comps = 2 Then
            dst_blk.m_low(0) = ((g_astc_unquant_m_unquant(endpoint_range, u.m_astc.m_endpoints(0)) * 127 + 127) \ 255) And &HFF
            dst_blk.m_high(0) = ((g_astc_unquant_m_unquant(endpoint_range, u.m_astc.m_endpoints(1)) * 127 + 127) \ 255) And &HFF

            dst_blk.m_low(1) = dst_blk.m_low(0)
            dst_blk.m_high(1) = dst_blk.m_high(0)

            dst_blk.m_low(2) = dst_blk.m_low(0)
            dst_blk.m_high(2) = dst_blk.m_high(0)

            dst_blk.m_low(3) = g_astc_unquant_m_unquant(endpoint_range, u.m_astc.m_endpoints(2))
            dst_blk.m_high(3) = g_astc_unquant_m_unquant(endpoint_range, u.m_astc.m_endpoints(3))
        Else
            For astc_comp = 0 To 3
                bc7_comp = astc_comp
                ' ASTC and BC7 handle dual plane component rotations differently:
                ' ASTC: 2nd plane separately interpolates the CCS channel.
                ' BC7: 2nd plane channel is swapped with alpha, 2nd plane controls alpha interpolation, then we swap alpha with the desired channel.
                If astc_comp = u.m_astc.m_ccs Then
                    bc7_comp = 3
                ElseIf astc_comp = 3 Then
                    bc7_comp = u.m_astc.m_ccs
                End If

                l = 255
                h = 255
                If astc_comp < total_comps Then
                    l = g_astc_unquant_m_unquant(endpoint_range, u.m_astc.m_endpoints(astc_comp * 2 + 0))
                    h = g_astc_unquant_m_unquant(endpoint_range, u.m_astc.m_endpoints(astc_comp * 2 + 1))
                End If

                If bc7_comp < 3 Then
                    l = (l * 127 + 127) \ 255
                    h = (h * 127 + 127) \ 255
                End If

                dst_blk.m_low(bc7_comp) = l And &HFF
                dst_blk.m_high(bc7_comp) = h And &HFF
            Next
        End If

        If mode = 13 Then
            For i = 0 To 15
                If u.m_astc.m_weights(i * 2) <> 0 Then dst_blk.m_selectors(i) = 3 Else dst_blk.m_selectors(i) = 0
                If u.m_astc.m_weights(i * 2 + 1) <> 0 Then dst_blk.m_alpha_selectors(i) = 3 Else dst_blk.m_alpha_selectors(i) = 0
            Next
        Else
            For i = 0 To 15
                dst_blk.m_selectors(i) = u.m_astc.m_weights(i * 2)
                dst_blk.m_alpha_selectors(i) = u.m_astc.m_weights(i * 2 + 1)
            Next
        End If

    Case 7
        ' DualPlane: 0, WeightRange : 2 (4), Subsets : 2, EndpointRange : 12 (40) - BC7 MODE2
        dst_blk.m_mode = 2
        dst_blk.m_partition = g_bc7_3_astc2_common_partitions_m_bc73(u.m_common_pattern)

        common_pattern_k = g_bc7_3_astc2_common_partitions_k(u.m_common_pattern)

        For bc7_part = 0 To 2
            astc_part = bc7_convert_partition_index_3_to_2(bc7_part, common_pattern_k)

            For c = 0 To 2
                dst_blk.m_low(bc7_part * 4 + c) = ((g_astc_unquant_m_unquant(endpoint_range, u.m_astc.m_endpoints(c * 2 + 0 + astc_part * 6)) * 31 + 127) \ 255) And &HFF
                dst_blk.m_high(bc7_part * 4 + c) = ((g_astc_unquant_m_unquant(endpoint_range, u.m_astc.m_endpoints(c * 2 + 1 + astc_part * 6)) * 31 + 127) \ 255) And &HFF
            Next
        Next

        For i = 0 To 15
            dst_blk.m_selectors(i) = u.m_astc.m_weights(i)
        Next

    Case UASTC_MODE_INDEX_SOLID_COLOR
        ' Void-Extent: Solid Color RGBA (BC7 MODE5 or MODE6)
        best_err0 = g_bc7_mode_6_optimal_endpoints_m_error(u.m_solid_color(0), 0) + g_bc7_mode_6_optimal_endpoints_m_error(u.m_solid_color(1), 0) + _
            g_bc7_mode_6_optimal_endpoints_m_error(u.m_solid_color(2), 0) + g_bc7_mode_6_optimal_endpoints_m_error(u.m_solid_color(3), 0)

        best_err1 = g_bc7_mode_6_optimal_endpoints_m_error(u.m_solid_color(0), 1) + g_bc7_mode_6_optimal_endpoints_m_error(u.m_solid_color(1), 1) + _
            g_bc7_mode_6_optimal_endpoints_m_error(u.m_solid_color(2), 1) + g_bc7_mode_6_optimal_endpoints_m_error(u.m_solid_color(3), 1)

        If best_err0 > 0 And best_err1 > 0 Then
            dst_blk.m_mode = 5

            For c = 0 To 2
                dst_blk.m_low(c) = g_bc7_mode_5_optimal_endpoints_m_lo(u.m_solid_color(c))
                dst_blk.m_high(c) = g_bc7_mode_5_optimal_endpoints_m_hi(u.m_solid_color(c))
            Next

            For i = 0 To 15
                dst_blk.m_selectors(i) = BC7ENC_MODE_5_OPTIMAL_INDEX
            Next

            dst_blk.m_low(3) = u.m_solid_color(3)
            dst_blk.m_high(3) = u.m_solid_color(3)
        Else
            dst_blk.m_mode = 6

            best_p = 0
            If best_err1 < best_err0 Then best_p = 1

            For c = 0 To 3
                dst_blk.m_low(c) = g_bc7_mode_6_optimal_endpoints_m_lo(u.m_solid_color(c), best_p)
                dst_blk.m_high(c) = g_bc7_mode_6_optimal_endpoints_m_hi(u.m_solid_color(c), best_p)
            Next

            dst_blk.m_pbits(0) = best_p
            dst_blk.m_pbits(1) = best_p
            For i = 0 To 15
                dst_blk.m_selectors(i) = BC7ENC_MODE_6_OPTIMAL_INDEX
            Next
        End If

    Case 9, 16
        ' 9. DualPlane: 0, WeightRange : 2 (4), Subsets : 2, EndpointRange : 8 (16) - BC7 MODE7
        ' 16. DualPlane: 0, WeightRange: 2 (4), Subsets: 2, CEM: 4 (LA Direct), EndpointRange: 20 (256) - BC7 MODE7
        dst_blk.m_mode = 7
        dst_blk.m_partition = g_astc_bc7_common_partitions2_m_bc7(u.m_common_pattern)

        invert_partition = g_astc_bc7_common_partitions2_m_invert(u.m_common_pattern) <> 0

        For astc_subset = 0 To 1
            If total_comps = 2 Then
                b4 = astc_subset * 4
                xl(0) = unq_f(endpoint_range, u.m_astc.m_endpoints(0 + b4))
                xh(0) = unq_f(endpoint_range, u.m_astc.m_endpoints(1 + b4))

                xl(1) = xl(0)
                xh(1) = xh(0)

                xl(2) = xl(0)
                xh(2) = xh(0)

                xl(3) = unq_f(endpoint_range, u.m_astc.m_endpoints(2 + b4))
                xh(3) = unq_f(endpoint_range, u.m_astc.m_endpoints(3 + b4))
            Else
                b8 = astc_subset * 8
                xl(0) = unq_f(endpoint_range, u.m_astc.m_endpoints(0 + b8))
                xl(1) = unq_f(endpoint_range, u.m_astc.m_endpoints(2 + b8))
                xl(2) = unq_f(endpoint_range, u.m_astc.m_endpoints(4 + b8))
                xl(3) = unq_f(endpoint_range, u.m_astc.m_endpoints(6 + b8))

                xh(0) = unq_f(endpoint_range, u.m_astc.m_endpoints(1 + b8))
                xh(1) = unq_f(endpoint_range, u.m_astc.m_endpoints(3 + b8))
                xh(2) = unq_f(endpoint_range, u.m_astc.m_endpoints(5 + b8))
                xh(3) = unq_f(endpoint_range, u.m_astc.m_endpoints(7 + b8))
            End If

            best_pbits(0) = 0
            best_pbits(1) = 0
            For c = 0 To 3
                bestMinColor(c) = 0
                bestMaxColor(c) = 0
            Next
            determine_unique_pbits 4, 5, xl, xh, bestMinColor, bestMaxColor, best_pbits

            If invert_partition Then bc7_subset_index = 1 - astc_subset Else bc7_subset_index = astc_subset

            For c = 0 To 3
                dst_blk.m_low(bc7_subset_index * 4 + c) = bestMinColor(c)
                dst_blk.m_high(bc7_subset_index * 4 + c) = bestMaxColor(c)
            Next

            dst_blk.m_pbits(bc7_subset_index * 2 + 0) = best_pbits(0)
            dst_blk.m_pbits(bc7_subset_index * 2 + 1) = best_pbits(1)
        Next

        For i = 0 To 15
            dst_blk.m_selectors(i) = u.m_astc.m_weights(i)
        Next

    Case Else
        Exit Function
    End Select

    transcode_uastc_to_bc7_res = True
End Function

' bool transcode_uastc_to_bc7(const uastc_block& src_blk, void* pDst)
Public Function transcode_uastc_to_bc7(src() As Byte, ByVal so As Long, d() As Byte, ByVal o As Long) As Boolean
    ' bool transcode_uastc_to_bc7(const uastc_block& src_blk, bc7_optimization_results& dst_blk)
    If Not unpack_uastc_blk(src, so, ub, False, False) Then Exit Function
    If Not transcode_uastc_to_bc7_res(ub, bc7r) Then Exit Function

    encode_bc7_block d, o, bc7r
    transcode_uastc_to_bc7 = True
End Function

'--------------------------------------------------------------------------
' BC4, BC1

' void encode_bc4(void* pDst, const uint8_t* pPixels, uint32_t stride): pPixels is pp(first), pp(first + stride), ...
Private Sub encode_bc4(d() As Byte, ByVal o As Long, pp() As Long, ByVal first As Long, ByVal stride As Long)
    Static s_tran(0 To 7) As Long, idx(0 To 15) As Long, tables_set As Boolean
    Dim min_v As Long, max_v As Long, i As Long, v As Long, delta As Long
    Dim t0 As Long, t1 As Long, t2 As Long, t3 As Long, t4 As Long, t5 As Long, t6 As Long, bias As Long
    Dim cur As Long, k As Long, b As Long

    If Not tables_set Then
        set_table s_tran, "1,7,6,5,4,3,2,0"             ' s_tran0 (s_tran1..3 are it shifted to the pixel's place)
        tables_set = True
    End If

    ' (the C keeps four running minimums and maximums and combines them: the same minimum and maximum)
    min_v = pp(first)
    max_v = min_v
    For i = 1 To 15
        v = pp(first + i * stride)
        If v < min_v Then min_v = v
        If v > max_v Then max_v = v
    Next

    d(o + 0) = max_v
    d(o + 1) = min_v

    If max_v = min_v Then
        For i = 2 To 7
            d(o + i) = 0
        Next
        Exit Sub
    End If

    delta = max_v - min_v

    ' min_v is now 0. Compute thresholds between values by scaling max_v. It's x14 because we're adding two x7 scale factors.
    t0 = delta * 13
    t1 = delta * 11
    t2 = delta * 9
    t3 = delta * 7
    t4 = delta * 5
    t5 = delta * 3
    t6 = delta * 1

    ' BC4 floors in its divisions, which we compensate for with the 4 bias.
    ' This function is optimal for all possible inputs (i.e. it outputs the same results as checking all 8 values and choosing the closest one).
    bias = 4 - min_v * 14

    For i = 0 To 15
        v = pp(first + i * stride) * 14 + bias
        k = 0
        If v >= t0 Then k = k + 1
        If v >= t1 Then k = k + 1
        If v >= t2 Then k = k + 1
        If v >= t3 Then k = k + 1
        If v >= t4 Then k = k + 1
        If v >= t5 Then k = k + 1
        If v >= t6 Then k = k + 1
        idx(i) = s_tran(k)
    Next

    ' f = a0 | a1 | a2 | a3: pixel i's 3 bits at bit 3 * i, written as bytes 2 to 7
    For i = 2 To 7
        d(o + i) = 0
    Next
    For i = 0 To 15
        cur = i * 3
        b = o + 2 + cur \ 8
        k = idx(i) * p2(cur And 7)
        d(b) = d(b) Or (k And &HFF)
        If k > 255 Then d(b + 1) = d(b + 1) Or (k \ 256)
    Next
End Sub

' static void bc1_find_sels(const color32 *pSrc_pixels, uint32_t lr, uint32_t lg, uint32_t lb, uint32_t hr, uint32_t hg, uint32_t hb, uint8_t sels[16])
Private Sub bc1_find_sels(pp() As Long, ByVal lr As Long, ByVal lg As Long, ByVal lb As Long, ByVal hr As Long, ByVal hg As Long, ByVal hb As Long, sels() As Long)
    Static block_r(0 To 3) As Long, block_g(0 To 3) As Long, block_b(0 To 3) As Long, dots(0 To 3) As Long, s_sels(0 To 3) As Long
    Dim ar As Long, ag As Long, ab As Long, i As Long, t0 As Long, t1 As Long, t2 As Long, d As Long, k As Long

    s_sels(0) = 3: s_sels(1) = 2: s_sels(2) = 1: s_sels(3) = 0

    block_r(0) = (lr * 8) Or (lr \ 4): block_g(0) = (lg * 4) Or (lg \ 16): block_b(0) = (lb * 8) Or (lb \ 4)
    block_r(3) = (hr * 8) Or (hr \ 4): block_g(3) = (hg * 4) Or (hg \ 16): block_b(3) = (hb * 8) Or (hb \ 4)
    block_r(1) = (block_r(0) * 2 + block_r(3)) \ 3: block_g(1) = (block_g(0) * 2 + block_g(3)) \ 3: block_b(1) = (block_b(0) * 2 + block_b(3)) \ 3
    block_r(2) = (block_r(3) * 2 + block_r(0)) \ 3: block_g(2) = (block_g(3) * 2 + block_g(0)) \ 3: block_b(2) = (block_b(3) * 2 + block_b(0)) \ 3

    ar = block_r(3) - block_r(0)
    ag = block_g(3) - block_g(0)
    ab = block_b(3) - block_b(0)

    For i = 0 To 3
        dots(i) = block_r(i) * ar + block_g(i) * ag + block_b(i) * ab
    Next

    t0 = dots(0) + dots(1)
    t1 = dots(1) + dots(2)
    t2 = dots(2) + dots(3)

    ar = ar * 2
    ag = ag * 2
    ab = ab * 2

    For i = 0 To 15
        d = pp(i * 4 + 0) * ar + pp(i * 4 + 1) * ag + pp(i * 4 + 2) * ab

        ' Rounding matters here!
        ' d <= t0: <=, not <, to the later LS step "sees" a wider range of selectors. It matters for quality.
        k = 0
        If d <= t0 Then k = k + 1
        If d < t1 Then k = k + 1
        If d < t2 Then k = k + 1
        sels(i) = s_sels(k)
    Next
End Sub

' static bool compute_least_squares_endpoints_rgb(const color32* pColors, const uint8_t* pSelectors, vec3F* pXl, vec3F* pXh)
Private Function compute_least_squares_endpoints_rgb(pp() As Long, pSelectors() As Long, xl() As Single, xh() As Single) As Boolean
    Static s_weight_vals(0 To 3) As Long
    Dim uq00_r As Long, uq10_r As Long, ut_r As Long, uq00_g As Long, uq10_g As Long, ut_g As Long, uq00_b As Long, uq10_b As Long, ut_b As Long
    Dim weight_accum As Long, i As Long, r As Long, g As Long, b As Long, sel As Long, c As Long, lo_v As Long, hi_v As Long, v As Long
    Dim q00_r As Single, q10_r As Single, t_r As Single, q00_g As Single, q10_g As Single, t_g As Single
    Dim q00_b As Single, q10_b As Single, t_b As Single, z00 As Single, z10 As Single, z11 As Single, z01 As Single
    Dim det As Single, iz00 As Single, iz01 As Single, iz10 As Single, iz11 As Single

    ' Derived from bc7enc16's LS function.
    ' This table is: 9 * (w * w), 9 * ((1.0f - w) * w), 9 * ((1.0f - w) * (1.0f - w))
    ' where w is [0,1/3,2/3,1]. 9 is the perfect multiplier.
    s_weight_vals(0) = &H9&: s_weight_vals(1) = &H10204: s_weight_vals(2) = &H40201: s_weight_vals(3) = &H90000

    For i = 0 To 15
        r = pp(i * 4 + 0)
        g = pp(i * 4 + 1)
        b = pp(i * 4 + 2)
        sel = pSelectors(i)
        ut_r = ut_r + r
        ut_g = ut_g + g
        ut_b = ut_b + b
        weight_accum = weight_accum + s_weight_vals(sel)
        uq00_r = uq00_r + sel * r
        uq00_g = uq00_g + sel * g
        uq00_b = uq00_b + sel * b
    Next

    q00_r = uq00_r: q10_r = uq10_r: t_r = ut_r
    q00_g = uq00_g: q10_g = uq10_g: t_g = ut_g
    q00_b = uq00_b: q10_b = uq10_b: t_b = ut_b

    q10_r = fsub_(fmul_(t_r, 3!), q00_r)
    q10_g = fsub_(fmul_(t_g, 3!), q00_g)
    q10_b = fsub_(fmul_(t_b, 3!), q00_b)

    z00 = (weight_accum \ 65536) And &HFF
    z10 = (weight_accum \ 256) And &HFF
    z11 = weight_accum And &HFF
    z01 = z10

    det = fsub_(fmul_(z00, z11), fmul_(z01, z10))
    If Abs(det) < 1E-08! Then Exit Function

    det = fdiv_(3!, det)

    iz00 = fmul_(z11, det)
    iz01 = fmul_(-z01, det)
    iz10 = fmul_(-z10, det)
    iz11 = fmul_(z00, det)

    xl(0) = fadd_(fmul_(iz00, q00_r), fmul_(iz01, q10_r)): xh(0) = fadd_(fmul_(iz10, q00_r), fmul_(iz11, q10_r))
    xl(1) = fadd_(fmul_(iz00, q00_g), fmul_(iz01, q10_g)): xh(1) = fadd_(fmul_(iz10, q00_g), fmul_(iz11, q10_g))
    xl(2) = fadd_(fmul_(iz00, q00_b), fmul_(iz01, q10_b)): xh(2) = fadd_(fmul_(iz10, q00_b), fmul_(iz11, q10_b))

    ' Check and fix channel singularities - might not be needed, but is in UASTC's encoder.
    For c = 0 To 2
        If (xl(c) < 0!) Or (xh(c) > 255!) Then
            lo_v = 256                                  ' UINT32_MAX (the values are bytes)
            hi_v = 0
            For i = 0 To 15
                v = pp(i * 4 + c)
                If v < lo_v Then lo_v = v
                If v > hi_v Then hi_v = v
            Next

            If lo_v = hi_v Then
                xl(c) = lo_v
                xh(c) = hi_v
            End If
        End If
    Next

    compute_least_squares_endpoints_rgb = True
End Function

' dxt1_block::set_low_color / set_high_color
Private Sub set_low_color(d() As Byte, ByVal o As Long, ByVal c As Long)
    d(o + 0) = c And &HFF
    d(o + 1) = (c \ 256) And &HFF
End Sub

Private Sub set_high_color(d() As Byte, ByVal o As Long, ByVal c As Long)
    d(o + 2) = c And &HFF
    d(o + 3) = (c \ 256) And &HFF
End Sub

' void encode_bc1_solid_block(void* pDst, uint32_t fr, uint32_t fg, uint32_t fb)
Private Sub encode_bc1_solid_block(d() As Byte, ByVal o As Long, ByVal fr As Long, ByVal fg As Long, ByVal fb As Long)
    Dim mask As Long, max16 As Long, min16 As Long, t As Long

    mask = &HAA
    max16 = (g_bc1_match5_equals_1_m_hi(fr) * 2048) Or (g_bc1_match6_equals_1_m_hi(fg) * 32) Or g_bc1_match5_equals_1_m_hi(fb)
    min16 = (g_bc1_match5_equals_1_m_lo(fr) * 2048) Or (g_bc1_match6_equals_1_m_lo(fg) * 32) Or g_bc1_match5_equals_1_m_lo(fb)

    If min16 = max16 Then
        ' Always forbid 3 color blocks
        ' This is to guarantee that BC3 blocks never use punchthrough alpha (3 color) mode, which isn't supported on some (all?) GPU's.
        mask = 0

        ' Make l > h
        If min16 > 0 Then
            min16 = min16 - 1
        Else
            ' l = h = 0
            max16 = 1
            min16 = 0
            mask = &H55
        End If
    End If

    If max16 < min16 Then
        t = max16
        max16 = min16
        min16 = t
        mask = mask Xor &H55
    End If

    set_low_color d, o, max16
    set_high_color d, o, min16
    d(o + 4) = mask
    d(o + 5) = mask
    d(o + 6) = mask
    d(o + 7) = mask
End Sub

' static inline uint8_t to_5(uint32_t v) / to_6
Private Function to_5(ByVal v As Long) As Long
    v = v * 31 + 128
    to_5 = ((v + (v \ 256)) \ 256) And &HFF
End Function

Private Function to_6(ByVal v As Long) As Long
    v = v * 63 + 128
    to_6 = ((v + (v \ 256)) \ 256) And &HFF
End Function

' void encode_bc1(void* pDst, const uint8_t* pPixels, uint32_t flags)
Private Sub encode_bc1(d() As Byte, ByVal o As Long, pp() As Long, ByVal flags As Long)
    Static sels(0 To 15) As Long, s_sel_tran(0 To 3) As Long, s_sel_trans(0 To 3) As Long, icov(0 To 5) As Long
    Static cov(0 To 5) As Single, xl(0 To 2) As Single, xh(0 To 2) As Single
    Dim avg_r As Long, avg_g As Long, avg_b As Long, lr As Long, lg As Long, lb As Long, hr As Long, hg As Long, hb As Long
    Dim use_sels As Boolean, i As Long, j As Long, fr As Long, fg As Long, fb As Long
    Dim total_r As Long, total_g As Long, total_b As Long, max_r As Long, max_g As Long, max_b As Long
    Dim min_r As Long, min_g As Long, min_b As Long, r As Long, g As Long, b As Long
    Dim fxr As Single, fxg As Single, fxb As Single, fr_ As Single, fg_ As Single, fb_ As Single, power_iter As Long
    Dim k As Single, m As Single, saxis_r As Long, saxis_g As Long, saxis_b As Long
    Dim low_dot As Long, high_dot As Long, low_c As Long, high_c As Long, dot As Long
    Dim total_ls_passes As Long, ls_pass As Long, lc16 As Long, hc16 As Long, mask As Long, invert_mask As Long
    Dim t As Long, sb As Long

    s_sel_tran(0) = 0: s_sel_tran(1) = 3: s_sel_tran(2) = 1: s_sel_tran(3) = 2
    s_sel_trans(0) = 0: s_sel_trans(1) = 2: s_sel_trans(2) = 3: s_sel_trans(3) = 1

    avg_r = -1
    avg_g = 0
    avg_b = 0
    lr = 0: lg = 0: lb = 0: hr = 0: hg = 0: hb = 0

    use_sels = (flags And cEncodeBC1UseSelectors) <> 0
    If use_sels Then
        ' Caller is jamming in their own selectors for us to try.
        For i = 0 To 15
            sels(i) = s_sel_tran((d(o + 4 + i \ 4) \ p2((i And 3) * 2)) And 3)
        Next
    Else
        fr = pp(0): fg = pp(1): fb = pp(2)

        For j = 1 To 15
            If (pp(j * 4 + 0) <> fr) Or (pp(j * 4 + 1) <> fg) Or (pp(j * 4 + 2) <> fb) Then Exit For
        Next

        If j = 16 Then
            encode_bc1_solid_block d, o, fr, fg, fb
            Exit Sub
        End If

        ' Select 2 colors along the principle axis. (There must be a faster/simpler way.)
        total_r = fr: total_g = fg: total_b = fb
        max_r = fr: max_g = fg: max_b = fb
        min_r = fr: min_g = fg: min_b = fb
        For i = 1 To 15
            r = pp(i * 4 + 0): g = pp(i * 4 + 1): b = pp(i * 4 + 2)
            If r > max_r Then max_r = r
            If g > max_g Then max_g = g
            If b > max_b Then max_b = b
            If r < min_r Then min_r = r
            If g < min_g Then min_g = g
            If b < min_b Then min_b = b
            total_r = total_r + r: total_g = total_g + g: total_b = total_b + b
        Next

        avg_r = (total_r + 8) \ 16
        avg_g = (total_g + 8) \ 16
        avg_b = (total_b + 8) \ 16

        For i = 0 To 5
            icov(i) = 0
        Next
        For i = 0 To 15
            r = pp(i * 4 + 0) - avg_r
            g = pp(i * 4 + 1) - avg_g
            b = pp(i * 4 + 2) - avg_b
            icov(0) = icov(0) + r * r
            icov(1) = icov(1) + r * g
            icov(2) = icov(2) + r * b
            icov(3) = icov(3) + g * g
            icov(4) = icov(4) + g * b
            icov(5) = icov(5) + b * b
        Next

        For i = 0 To 5
            cov(i) = fmul_(icov(i), f_1_255)
        Next

        fxr = max_r - min_r
        fxg = max_g - min_g
        fxb = max_b - min_b
        For power_iter = 0 To 3
            fr_ = fadd_(fadd_(fmul_(fxr, cov(0)), fmul_(fxg, cov(1))), fmul_(fxb, cov(2)))
            fg_ = fadd_(fadd_(fmul_(fxr, cov(1)), fmul_(fxg, cov(3))), fmul_(fxb, cov(4)))
            fb_ = fadd_(fadd_(fmul_(fxr, cov(2)), fmul_(fxg, cov(4))), fmul_(fxb, cov(5)))
            fxr = fr_: fxg = fg_: fxb = fb_
        Next

        ' basisu::maximum(fabsf(xr), fabsf(xg), fabsf(xb))
        k = Abs(fxr)
        If Abs(fxg) > k Then k = Abs(fxg)
        If Abs(fxb) > k Then k = Abs(fxb)
        saxis_r = 306: saxis_g = 601: saxis_b = 117
        If k >= 2! Then
            m = fdiv_(1024!, k)
            saxis_r = Fix(fmul_(fxr, m))
            saxis_g = Fix(fmul_(fxg, m))
            saxis_b = Fix(fmul_(fxb, m))
        End If

        low_dot = &H7FFFFFFF: high_dot = &H80000000: low_c = 0: high_c = 0
        For i = 0 To 15
            dot = pp(i * 4 + 0) * saxis_r + pp(i * 4 + 1) * saxis_g + pp(i * 4 + 2) * saxis_b
            If dot < low_dot Then
                low_dot = dot
                low_c = i
            End If
            If dot > high_dot Then
                high_dot = dot
                high_c = i
            End If
        Next

        lr = to_5(pp(low_c * 4 + 0))
        lg = to_6(pp(low_c * 4 + 1))
        lb = to_5(pp(low_c * 4 + 2))

        hr = to_5(pp(high_c * 4 + 0))
        hg = to_6(pp(high_c * 4 + 1))
        hb = to_5(pp(high_c * 4 + 2))

        bc1_find_sels pp, lr, lg, lb, hr, hg, hb, sels
    End If

    If (flags And cEncodeBC1HigherQuality) <> 0 Then
        total_ls_passes = 3
    ElseIf (flags And cEncodeBC1HighQuality) <> 0 Then
        total_ls_passes = 2
    Else
        total_ls_passes = 1
    End If
    For ls_pass = 0 To total_ls_passes - 1
        ' This is where the real magic happens. We have an array of candidate selectors, so let's use least squares to compute the optimal low/high endpoint colors.
        If Not compute_least_squares_endpoints_rgb(pp, sels, xl, xh) Then
            If avg_r < 0 Then
                total_r = 0: total_g = 0: total_b = 0
                For i = 0 To 15
                    total_r = total_r + pp(i * 4 + 0)
                    total_g = total_g + pp(i * 4 + 1)
                    total_b = total_b + pp(i * 4 + 2)
                Next

                avg_r = (total_r + 8) \ 16
                avg_g = (total_g + 8) \ 16
                avg_b = (total_b + 8) \ 16
            End If

            ' All selectors equal - treat it as a solid block which should always be equal or better.
            lr = g_bc1_match5_equals_1_m_hi(avg_r)
            lg = g_bc1_match6_equals_1_m_hi(avg_g)
            lb = g_bc1_match5_equals_1_m_hi(avg_b)

            hr = g_bc1_match5_equals_1_m_lo(avg_r)
            hg = g_bc1_match6_equals_1_m_lo(avg_g)
            hb = g_bc1_match5_equals_1_m_lo(avg_b)

            ' In high/higher quality mode, let it try again in case the optimal tables have caused the sels to diverge.
        Else
            lr = clampi(Fix(fadd_(fmul_(xl(0), f_31_255), 0.5!)), 0, 31)
            lg = clampi(Fix(fadd_(fmul_(xl(1), f_63_255), 0.5!)), 0, 63)
            lb = clampi(Fix(fadd_(fmul_(xl(2), f_31_255), 0.5!)), 0, 31)

            hr = clampi(Fix(fadd_(fmul_(xh(0), f_31_255), 0.5!)), 0, 31)
            hg = clampi(Fix(fadd_(fmul_(xh(1), f_63_255), 0.5!)), 0, 63)
            hb = clampi(Fix(fadd_(fmul_(xh(2), f_31_255), 0.5!)), 0, 31)
        End If

        bc1_find_sels pp, lr, lg, lb, hr, hg, hb, sels
    Next

    lc16 = lb Or (lg * 32) Or (lr * 2048)               ' dxt1_block::pack_unscaled_color
    hc16 = hb Or (hg * 32) Or (hr * 2048)

    ' Always forbid 3 color blocks
    If lc16 = hc16 Then
        mask = 0

        ' Make l > h
        If hc16 > 0 Then
            hc16 = hc16 - 1
        Else
            ' lc16 = hc16 = 0
            hc16 = 0
            lc16 = 1
            mask = &H55                                 ' select hc16
        End If

        set_low_color d, o, lc16
        set_high_color d, o, hc16

        d(o + 4) = mask
        d(o + 5) = mask
        d(o + 6) = mask
        d(o + 7) = mask
    Else
        invert_mask = 0
        If lc16 < hc16 Then
            t = lc16
            lc16 = hc16
            hc16 = t
            invert_mask = &H55
        End If

        set_low_color d, o, lc16
        set_high_color d, o, hc16

        ' packed_sels |= s_sel_trans[sels[i]] << (i * 2), written byte by byte
        For j = 0 To 3
            sb = 0
            For i = 3 To 0 Step -1
                sb = sb * 4 Or s_sel_trans(sels(j * 4 + i))
            Next
            d(o + 4 + j) = sb Xor invert_mask
        Next
    End If
End Sub

' dxt1_block::pack_color(color32(r, g, b, 255), true, 127)
Private Function pack_color_scaled(ByVal r As Long, ByVal g As Long, ByVal b As Long) As Long
    r = (r * 31 + 127) \ 255
    g = (g * 63 + 127) \ 255
    b = (b * 31 + 127) \ 255
    If b > 31 Then b = 31
    If g > 63 Then g = 63
    If r > 31 Then r = 31
    pack_color_scaled = b Or (g * 32) Or (r * 2048)
End Function

' void transcode_uastc_to_bc1_hint0(const unpacked_uastc_block& unpacked_src_blk, void* pDst)
' Scale the UASTC first subset endpoints and first plane's weight indices directly to BC1's - fastest.
Private Sub transcode_uastc_to_bc1_hint0(u As unpacked_uastc_block, d() As Byte, ByVal o As Long)
    Dim mode As Long, endpoint_range As Long, total_comps As Long, l As Long, h As Long, lc16 As Long, hc16 As Long
    Dim mask As Long, invert As Boolean, t As Long, plane_shift As Long, wb As Long, i As Long, j As Long, s As Long, sb As Long

    mode = u.m_mode

    endpoint_range = g_uastc_mode_endpoint_ranges(mode)

    total_comps = g_uastc_mode_comps(mode)

    If total_comps = 2 Then
        l = g_astc_unquant_m_unquant(endpoint_range, u.m_astc.m_endpoints(0))
        h = g_astc_unquant_m_unquant(endpoint_range, u.m_astc.m_endpoints(1))

        set_low_color d, o, pack_color_scaled(l, l, l)
        set_high_color d, o, pack_color_scaled(h, h, h)
    Else
        set_low_color d, o, pack_color_scaled( _
            g_astc_unquant_m_unquant(endpoint_range, u.m_astc.m_endpoints(0)), _
            g_astc_unquant_m_unquant(endpoint_range, u.m_astc.m_endpoints(2)), _
            g_astc_unquant_m_unquant(endpoint_range, u.m_astc.m_endpoints(4)))

        set_high_color d, o, pack_color_scaled( _
            g_astc_unquant_m_unquant(endpoint_range, u.m_astc.m_endpoints(1)), _
            g_astc_unquant_m_unquant(endpoint_range, u.m_astc.m_endpoints(3)), _
            g_astc_unquant_m_unquant(endpoint_range, u.m_astc.m_endpoints(5)))
    End If

    lc16 = d(o + 0) Or (d(o + 1) * 256&)                ' get_low_color
    hc16 = d(o + 2) Or (d(o + 3) * 256&)                ' get_high_color
    If lc16 = hc16 Then
        ' Always forbid 3 color blocks
        mask = 0

        ' Make l > h
        If hc16 > 0 Then
            hc16 = hc16 - 1
        Else
            ' lc16 = hc16 = 0
            hc16 = 0
            lc16 = 1
            mask = &H55                                 ' select hc16
        End If

        set_low_color d, o, lc16
        set_high_color d, o, hc16

        d(o + 4) = mask
        d(o + 5) = mask
        d(o + 6) = mask
        d(o + 7) = mask
    Else
        invert = False
        If lc16 < hc16 Then
            t = d(o + 0): d(o + 0) = d(o + 2): d(o + 2) = t
            t = d(o + 1): d(o + 1) = d(o + 3): d(o + 3) = t
            invert = True
        End If

        wb = g_uastc_mode_weight_bits(mode)                 ' pTran = s_uastc_to_bc1_weights[...]

        plane_shift = g_uastc_mode_planes(mode) - 1

        ' sels = (sels << 2) | s for i = 15 down to 0: pixel i's 2 bits at bit 2 * i, written byte by byte
        For j = 0 To 3
            sb = 0
            For i = j * 4 + 3 To j * 4 Step -1
                s = s_uastc_to_bc1_weights(wb, u.m_astc.m_weights(i * p2(plane_shift)))

                If invert Then s = s Xor 1

                sb = sb * 4 Or s
            Next
            d(o + 4 + j) = sb
        Next
    End If
End Sub

' void transcode_uastc_to_bc1_hint1(const unpacked_uastc_block& unpacked_src_blk, const color32 block_pixels[4][4], void* pDst, bool high_quality)
' Scale the UASTC first plane's weight indices to BC1, use 1 or 2 least squares passes to compute endpoints - no PCA needed.
Private Sub transcode_uastc_to_bc1_hint1(u As unpacked_uastc_block, pp() As Long, d() As Byte, ByVal o As Long, ByVal high_quality As Boolean)
    Dim mode As Long, wb As Long, plane_shift As Long, i As Long, j As Long, sb As Long

    mode = u.m_mode

    set_low_color d, o, 1
    set_high_color d, o, 0

    wb = g_uastc_mode_weight_bits(mode)

    plane_shift = g_uastc_mode_planes(mode) - 1

    For j = 0 To 3
        sb = 0
        For i = j * 4 + 3 To j * 4 Step -1
            sb = sb * 4 Or s_uastc_to_bc1_weights(wb, u.m_astc.m_weights(i * p2(plane_shift)))
        Next
        d(o + 4 + j) = sb
    Next

    If high_quality Then
        encode_bc1 d, o, pp, cEncodeBC1HighQuality Or cEncodeBC1UseSelectors
    Else
        encode_bc1 d, o, pp, cEncodeBC1UseSelectors
    End If
End Sub

' bool transcode_uastc_to_bc1(const uastc_block& src_blk, void* pDst, bool high_quality)
Public Function transcode_uastc_to_bc1(src() As Byte, ByVal so As Long, d() As Byte, ByVal o As Long, ByVal high_quality As Boolean) As Boolean
    Dim mode As Long
    If Not unpack_uastc_blk(src, so, ub, False, True) Then Exit Function

    mode = ub.m_mode

    If mode = UASTC_MODE_INDEX_SOLID_COLOR Then
        encode_bc1_solid_block d, o, ub.m_solid_color(0), ub.m_solid_color(1), ub.m_solid_color(2)
        transcode_uastc_to_bc1 = True
        Exit Function
    End If

    If (Not high_quality) And ub.m_bc1_hint0 Then
        transcode_uastc_to_bc1_hint0 ub, d, o
    Else
        ' const bool unpack_srgb = false;
        If Not unpack_uastc_pixels(ub, px, False) Then Exit Function

        If ub.m_bc1_hint1 Then
            transcode_uastc_to_bc1_hint1 ub, px, d, o, high_quality
        Else
            If high_quality Then encode_bc1 d, o, px, cEncodeBC1HighQuality Else encode_bc1 d, o, px, 0
        End If
    End If

    transcode_uastc_to_bc1 = True
End Function

' static void write_bc4_solid_block(uint8_t* pDst, uint32_t a)
Private Sub write_bc4_solid_block(d() As Byte, ByVal o As Long, ByVal a As Long)
    Dim i As Long
    d(o + 0) = a
    d(o + 1) = a
    For i = 2 To 7
        d(o + i) = 0
    Next
End Sub

' bool transcode_uastc_to_bc3(const uastc_block& src_blk, void* pDst, bool high_quality)
Public Function transcode_uastc_to_bc3(src() As Byte, ByVal so As Long, d() As Byte, ByVal o As Long, ByVal high_quality As Boolean) As Boolean
    Dim mode As Long
    If Not unpack_uastc_blk(src, so, ub, False, True) Then Exit Function

    mode = ub.m_mode

    ' pBC4_block = pDst; pBC1_block = &((dxt1_block*)pDst)[1]: o + 8

    If mode = UASTC_MODE_INDEX_SOLID_COLOR Then
        write_bc4_solid_block d, o, ub.m_solid_color(3)
        encode_bc1_solid_block d, o + 8, ub.m_solid_color(0), ub.m_solid_color(1), ub.m_solid_color(2)
        transcode_uastc_to_bc3 = True
        Exit Function
    End If

    ' const bool unpack_srgb = false;
    If Not unpack_uastc_pixels(ub, px, False) Then Exit Function

    encode_bc4 d, o, px, 3, 4

    If (Not high_quality) And ub.m_bc1_hint0 Then
        transcode_uastc_to_bc1_hint0 ub, d, o + 8
    Else
        If ub.m_bc1_hint1 Then
            transcode_uastc_to_bc1_hint1 ub, px, d, o + 8, high_quality
        Else
            If high_quality Then encode_bc1 d, o + 8, px, cEncodeBC1HighQuality Else encode_bc1 d, o + 8, px, 0
        End If
    End If

    transcode_uastc_to_bc3 = True
End Function

' bool transcode_uastc_to_bc4(const uastc_block& src_blk, void* pDst, bool high_quality, uint32_t chan0)
Public Function transcode_uastc_to_bc4(src() As Byte, ByVal so As Long, d() As Byte, ByVal o As Long, ByVal high_quality As Boolean, ByVal chan0 As Long) As Boolean
    Dim mode As Long
    If Not unpack_uastc_blk(src, so, ub, False, True) Then Exit Function

    mode = ub.m_mode

    If mode = UASTC_MODE_INDEX_SOLID_COLOR Then
        write_bc4_solid_block d, o, ub.m_solid_color(chan0)
        transcode_uastc_to_bc4 = True
        Exit Function
    End If

    If Not unpack_uastc_pixels(ub, px, False) Then Exit Function

    encode_bc4 d, o, px, chan0, 4

    transcode_uastc_to_bc4 = True
End Function

' bool transcode_uastc_to_bc5(const uastc_block& src_blk, void* pDst, bool high_quality, uint32_t chan0, uint32_t chan1)
Public Function transcode_uastc_to_bc5(src() As Byte, ByVal so As Long, d() As Byte, ByVal o As Long, ByVal high_quality As Boolean, ByVal chan0 As Long, ByVal chan1 As Long) As Boolean
    Dim mode As Long
    If Not unpack_uastc_blk(src, so, ub, False, True) Then Exit Function

    mode = ub.m_mode

    ' pBC4_block0 = pDst; pBC4_block1 = pDst + 8

    If mode = UASTC_MODE_INDEX_SOLID_COLOR Then
        write_bc4_solid_block d, o, ub.m_solid_color(chan0)
        write_bc4_solid_block d, o + 8, ub.m_solid_color(chan1)
        transcode_uastc_to_bc5 = True
        Exit Function
    End If

    If Not unpack_uastc_pixels(ub, px, False) Then Exit Function

    encode_bc4 d, o, px, chan0, 4
    encode_bc4 d, o + 8, px, chan1, 4

    transcode_uastc_to_bc5 = True
End Function
