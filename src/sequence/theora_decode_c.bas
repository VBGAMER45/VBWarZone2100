Attribute VB_Name = "theora_decode_c"
' ref/theora (libtheora 1.2.0) lib/decode.c: the decoder (frame header, coded flags, modes, motion vectors,
' qis, DCT tokens, DC prediction, reconstruction, the loop filter pipeline) and the out-of-loop
' post-processing (deblocking, deringing). The tables (OC_MODE_ALPHABETS, OC_INTERNAL_DCT_TOKEN_EXTRA_BITS,
' OC_DCT_CODE_WORD, the run, mode and MV trees, OC_HUFF_LIST_MAX, OC_MOD_MAX, OC_MOD_SHIFT) are in the
' generated theora_tables_vb. Not ported: the HAVE_CAIRO telemetry (oc_render_telemetry and the
' TH_DECCTL_SET_TELEMETRY_* controls), OC_DUMP_IMAGES, and the striped decode callback (TH_DECCTL_SET_STRIPE_CB:
' the legacy API cannot pass a function; th_decode_ctl answers TH_EIMPL, as for an unknown request).
' (VB) A th_dec_ctx * is a handle into theora_decint_h's oc_decs, allocated and freed here; the internal
' functions take the pool entry ByRef. The decoder's post-processing buffer (pp_frame_data) is reserved in
' state.ref_frame_handle after the reference frames when the decoder is made (oc_dec_init), so every frame
' pointer is an offset into that one array; "allocating" it hands out that offset. A C pointer into the
' decoder's arrays (variance, dc_qi, coded_fragis, ...) is an index. ptrdiff_t, size_t and the EOB counts are
' Longs: OC_DCT_EOB_FINISH is &H7FFFFFFF (the C's PTRDIFF_MAX on a 32-bit build), and the comparison
' ntoks+_eobs<ntoks_left is written as _eobs<ntoks_left-ntoks so that it cannot overflow.
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (dest As Any, ByVal length As Long, ByVal fill As Byte)

' No post-processing.
Private Const OC_PP_LEVEL_DISABLED As Long = 0
' Keep track of DC qi for each block only.
Private Const OC_PP_LEVEL_TRACKDCQI As Long = 1
' Deblock the luma plane.
Private Const OC_PP_LEVEL_DEBLOCKY As Long = 2
' Dering the luma plane.
Private Const OC_PP_LEVEL_DERINGY As Long = 3
' Stronger luma plane deringing.
Private Const OC_PP_LEVEL_SDERINGY As Long = 4
' Deblock the chroma planes.
Private Const OC_PP_LEVEL_DEBLOCKC As Long = 5
' Dering the chroma planes.
Private Const OC_PP_LEVEL_DERINGC As Long = 6
' Stronger chroma plane deringing.
Private Const OC_PP_LEVEL_SDERINGC As Long = 7
' Maximum valid post-processing level.
Public Const OC_PP_LEVEL_MAX As Long = 7

' The mode alphabets for the various mode coding schemes. Scheme 0 uses a custom alphabet, which is not stored
' in this table. (OC_MODE_ALPHABETS[7][OC_NMODES]: theora_tables_vb)

' The original DCT tokens are extended and reordered during the construction of the Huffman tables. The
' extension means more bits can be read with fewer calls to the bitpacker during the Huffman decoding process
' (at the cost of larger Huffman tables), and fewer tokens require additional extra bits (reducing the average
' storage per decoded token). The revised ordering reveals essential information in the token value itself;
' specifically, whether or not there are additional extra bits to read and the parameter to which those extra
' bits are applied. The token is used to fetch a code word from the OC_DCT_CODE_WORD table below. The extra
' bits are added into code word at the bit position inferred from the token value, giving the final code word
' from which all required parameters are derived. The number of EOBs and the leading zero run length can be
' extracted directly. The coefficient magnitude is optionally negated before extraction, according to a
' 'flip' bit.

' The number of additional extra bits that are decoded with each of the internal DCT tokens.
' (OC_INTERNAL_DCT_TOKEN_EXTRA_BITS[15]: theora_tables_vb)

' Whether or not an internal token needs any additional extra bits: OC_DCT_TOKEN_NEEDS_MORE(token) is
' token < 15 (the size of OC_INTERNAL_DCT_TOKEN_EXTRA_BITS), written out.
Private Const OC_NINTERNAL_DCT_TOKEN_EXTRA_BITS As Long = 15

' This token (OC_DCT_REPEAT_RUN3_TOKEN) requires more than 8 extra bits.
Private Const OC_DCT_TOKEN_FAT_EOB As Long = 0

' The number of EOBs to use for an end-of-frame token. Note: We want to set eobs to PTRDIFF_MAX here, but
' that requires C99, which is not yet available everywhere; this should be equivalent.
Private Const OC_DCT_EOB_FINISH As Long = &H7FFFFFFF

' The location of the (6) run length bits in the code word. These are placed at index 0 and given 8 bits
' (even though 6 would suffice) because it may be faster to extract the lower byte on some platforms.
Private Const OC_DCT_CW_RLEN_SHIFT As Long = 0
' The location of the (12) EOB bits in the code word.
Private Const OC_DCT_CW_EOB_SHIFT As Long = 8
' The location of the (1) flip bit in the code word. This must be right under the magnitude bits.
Private Const OC_DCT_CW_FLIP_BIT As Long = 20
' The location of the (11) token magnitude bits in the code word. These must be last, and rely on a
' sign-extending right shift.
Private Const OC_DCT_CW_MAG_SHIFT As Long = 21

' A special code word value that signals the end of the frame (a long EOB run of zero).
Private Const OC_DCT_CW_FINISH As Long = 0

' (oc_dec_dering_frag_rows)
Private Const OC_DERING_THRESH1 As Long = 384
Private Const OC_DERING_THRESH2 As Long = 4 * OC_DERING_THRESH1
Private Const OC_DERING_THRESH3 As Long = 5 * OC_DERING_THRESH1
Private Const OC_DERING_THRESH4 As Long = 10 * OC_DERING_THRESH1

' The position at which to insert the extra bits in the code word. (OC_DCT_TOKEN_EB_POS(_token): 8 for
' tokens below 2, 21 below 12, 0 otherwise.) (VB) The extra bits are multiplied by 1 << that position:
Private Function OC_DCT_TOKEN_EB_MUL(ByVal token As Long) As Long
    If token < 2 Then
        OC_DCT_TOKEN_EB_MUL = &H100&
    ElseIf token < 12 Then
        OC_DCT_TOKEN_EB_MUL = &H200000
    Else
        OC_DCT_TOKEN_EB_MUL = 1
    End If
End Function

' The code words for each internal token. See the notes at OC_DCT_TOKEN_MAP for the reasons why things are
' out of order. (OC_DCT_CODE_WORD[92]: theora_tables_vb)

Private Function oc_sb_run_unpack(opb As oc_pack_buf) As Long
    ' Coding scheme:
    '    Codeword            Run Length
    '  0                       1
    '  10x                     2-3
    '  110x                    4-5
    '  1110xx                  6-9
    '  11110xxx                10-17
    '  111110xxxx              18-33
    '  111111xxxxxxxxxxxx      34-4129
    ' (OC_SB_RUN_TREE[22]: theora_tables_vb)
    Dim ret As Long
    Dim offs As Long
    ret = oc_huff_token_decode(opb, OC_SB_RUN_TREE, 0)
    If ret >= &H10 Then
        offs = ret And &H1F
        ret = 6 + offs + oc_pack_read(opb, (ret - offs) \ 16)
    End If
    oc_sb_run_unpack = ret
End Function

Private Function oc_block_run_unpack(opb As oc_pack_buf) As Long
    ' Coding scheme:
    '   Codeword             Run Length
    '   0x                      1-2
    '   10x                     3-4
    '   110x                    5-6
    '   1110xx                  7-10
    '   11110xx                 11-14
    '   11111xxxx               15-30
    ' (OC_BLOCK_RUN_TREE[61]: theora_tables_vb)
    oc_block_run_unpack = oc_huff_token_decode(opb, OC_BLOCK_RUN_TREE, 0)
End Function

Public Sub oc_dec_accel_init_c(dec As th_dec_ctx)
End Sub

' (VB) _ogg_malloc of a decoder (oc_aligned_malloc): a free entry of oc_decs, zeroed; its handle
Private Function oc_dec_vbAlloc() As Long
    Dim h As Long
    Dim blank As th_dec_ctx
    For h = 1 To oc_ndecs
        If Not oc_decs(h).vbUsed Then Exit For
    Next
    If h > oc_ndecs Then
        oc_ndecs = oc_ndecs + 1
        ReDim Preserve oc_decs(1 To oc_ndecs)
        h = oc_ndecs
    End If
    oc_decs(h) = blank
    oc_decs(h).vbUsed = True
    oc_dec_vbAlloc = h
End Function

' (VB) oc_aligned_free of a decoder
Private Sub oc_dec_vbFree(ByVal h As Long)
    Dim blank As th_dec_ctx
    oc_decs(h) = blank
End Sub

Private Function oc_dec_init(dec As th_dec_ctx, info As th_info, setup As th_setup_info) As Long
    Dim qti As Long
    Dim pli As Long
    Dim qi As Long
    Dim ret As Long
    Dim qsum As Long
    Dim t As Long
    Dim c_w As Long
    Dim c_h As Long
    ret = oc_state_init(dec.state, info, 3)
    If ret < 0 Then oc_dec_init = ret: Exit Function
    ReDim dec.huff_tables(0 To TH_NHUFFMAN_TABLES - 1)
    ret = oc_huff_trees_copy(dec.huff_tables, dec.vb_huff, dec.vb_nhuff, setup.huff_tables, setup.vb_huff)
    If ret < 0 Then
        oc_state_clear dec.state
        oc_dec_init = ret
        Exit Function
    End If
    ' For each fragment, allocate one byte for every DCT coefficient token, plus one byte for extra-bits for
    ' each token, plus one more byte for the long EOB run, just in case it's the very last token and has a run
    ' length of one.
    ReDim dec.dct_tokens(0 To (64 + 64 + 1) * dec.state.nfrags - 1)
    For qi = 0 To 63
        For pli = 0 To 2
            For qti = 0 To 1
                dec.state.dequant_tables((qi * 3 + pli) * 2 + qti) = ((qi * 3 + pli) * 2 + qti) * 64
            Next
        Next
    Next
    ReDim dec.pp_dc_scale(0 To 63)
    oc_dequant_tables_init dec.state.dequant_tables, dec.state.dequant_table_data, dec.pp_dc_scale, False, setup.qinfo
    For qi = 0 To 63
        qsum = 0
        For qti = 0 To 1
            For pli = 0 To 2
                t = dec.state.dequant_tables((qi * 3 + pli) * 2 + qti)
                ' qsum+=a+b+c+d<<(pli==0)
                If pli = 0 Then
                    qsum = qsum + (dec.state.dequant_table_data(t + 12) + dec.state.dequant_table_data(t + 17) + _
                        dec.state.dequant_table_data(t + 18) + dec.state.dequant_table_data(t + 24)) * 2
                Else
                    qsum = qsum + dec.state.dequant_table_data(t + 12) + dec.state.dequant_table_data(t + 17) + _
                        dec.state.dequant_table_data(t + 18) + dec.state.dequant_table_data(t + 24)
                End If
            Next
        Next
        dec.pp_sharp_mod(qi) = -(qsum \ 2048)
    Next
    For qi = 0 To 63
        dec.state.loop_filter_limits(qi) = setup.qinfo.loop_filter_limits(qi)
    Next
    oc_dec_accel_init_c dec
    dec.pp_level = OC_PP_LEVEL_DISABLED
    dec.vbDcQisAllocated = False
    dec.vbVariancesAllocated = False
    dec.pp_frame_data = 0
    ' _dec->stripe_cb.ctx=NULL; _dec->stripe_cb.stripe_decoded=NULL: (VB) not ported
    ' (VB) the post-processed frame's storage (see oc_dec_postprocess_init) after the reference frames
    If (dec.state.info.pixel_fmt And 1) = 0 Then c_w = dec.state.info.frame_width \ 2 Else c_w = dec.state.info.frame_width
    If (dec.state.info.pixel_fmt And 2) = 0 Then c_h = dec.state.info.frame_height \ 2 Else c_h = dec.state.info.frame_height
    ReDim Preserve dec.state.ref_frame_handle(0 To dec.state.vb_ref_frame_sz + _
        dec.state.info.frame_width * dec.state.info.frame_height + 2 * c_w * c_h - 1)
    ' (VB) the pipeline's arrays (the C's are inside the struct)
    ReDim dec.pipe.dct_coeffs(0 To 127)
    ReDim dec.pipe.bounding_values(0 To 255)
    oc_dec_init = 0
End Function

Private Sub oc_dec_clear(dec As th_dec_ctx)
    ' _ogg_free(_dec->pp_frame_data): (VB) part of state.ref_frame_handle
    dec.pp_frame_data = 0
    Erase dec.variances
    dec.vbVariancesAllocated = False
    Erase dec.dc_qis
    dec.vbDcQisAllocated = False
    Erase dec.dct_tokens
    oc_huff_trees_clear dec.huff_tables, dec.vb_huff, dec.vb_nhuff
    oc_state_clear dec.state
End Sub

Private Function oc_dec_frame_header_unpack(dec As th_dec_ctx) As Long
    Dim val_ As Long
    ' Check to make sure this is a data packet.
    val_ = oc_pack_read1(dec.opb)
    If val_ <> 0 Then oc_dec_frame_header_unpack = TH_EBADPACKET: Exit Function
    ' Read in the frame type (I or P).
    val_ = oc_pack_read1(dec.opb)
    dec.state.frame_type = val_
    ' Read in the qi list.
    val_ = oc_pack_read(dec.opb, 6)
    dec.state.qis(0) = val_
    val_ = oc_pack_read1(dec.opb)
    If val_ = 0 Then
        dec.state.nqis = 1
    Else
        val_ = oc_pack_read(dec.opb, 6)
        dec.state.qis(1) = val_
        val_ = oc_pack_read1(dec.opb)
        If val_ = 0 Then
            dec.state.nqis = 2
        Else
            val_ = oc_pack_read(dec.opb, 6)
            dec.state.qis(2) = val_
            dec.state.nqis = 3
        End If
    End If
    If dec.state.frame_type = OC_INTRA_FRAME Then
        ' Keyframes have 3 unused configuration bits, holdovers from VP3 days. Most of the other unused bits in
        ' the VP3 headers were eliminated. I don't know why these remain.
        ' I wanted to eliminate wasted bits, but not all config wiggle room --Monty.
        val_ = oc_pack_read(dec.opb, 3)
        If val_ <> 0 Then oc_dec_frame_header_unpack = TH_EIMPL: Exit Function
    End If
    oc_dec_frame_header_unpack = 0
End Function

' Mark all fragments as coded and in OC_MODE_INTRA. This also builds up the coded fragment list (in coded
' order), and clears the uncoded fragment list. It does not update the coded macro block list nor the super
' block flags, as those are not used when decoding INTRA frames.
Private Sub oc_dec_mark_all_intra(dec As th_dec_ctx)
    Dim ncoded_fragis As Long
    Dim prev_ncoded_fragis As Long
    Dim nsbs As Long
    Dim sbi As Long
    Dim pli As Long
    Dim quadi As Long
    Dim bi As Long
    Dim fragi As Long
    prev_ncoded_fragis = 0
    ncoded_fragis = 0
    sbi = 0
    nsbs = 0
    For pli = 0 To 2
        nsbs = nsbs + dec.state.fplanes(pli).nsbs
        Do While sbi < nsbs
            For quadi = 0 To 3
                If (dec.state.sb_flags(sbi).quad_valid And OC_PB_POW2(quadi)) <> 0 Then
                    For bi = 0 To 3
                        fragi = dec.state.sb_maps(sbi * 16 + quadi * 4 + bi)
                        If fragi >= 0 Then
                            dec.state.frags(fragi).coded = 1
                            dec.state.frags(fragi).refi = OC_FRAME_SELF
                            dec.state.frags(fragi).mb_mode = OC_MODE_INTRA
                            dec.state.coded_fragis(ncoded_fragis) = fragi
                            ncoded_fragis = ncoded_fragis + 1
                        End If
                    Next
                End If
            Next
            sbi = sbi + 1
        Loop
        dec.state.ncoded_fragis(pli) = ncoded_fragis - prev_ncoded_fragis
        prev_ncoded_fragis = ncoded_fragis
    Next
    dec.state.ntotal_coded_fragis = ncoded_fragis
End Sub

' Decodes the bit flags indicating whether each super block is partially coded or not.
' Return: The number of partially coded super blocks.
Private Function oc_dec_partial_sb_flags_unpack(dec As th_dec_ctx) As Long
    Dim nsbs As Long
    Dim sbi As Long
    Dim npartial As Long
    Dim run_count As Long
    Dim val_ As Long
    Dim flag As Long
    Dim full_run As Boolean
    val_ = oc_pack_read1(dec.opb)
    flag = val_
    nsbs = dec.state.nsbs
    sbi = 0
    npartial = 0
    Do While sbi < nsbs
        run_count = oc_sb_run_unpack(dec.opb)
        full_run = run_count >= 4129
        Do
            dec.state.sb_flags(sbi).coded_partially = flag
            dec.state.sb_flags(sbi).coded_fully = 0
            npartial = npartial + flag
            sbi = sbi + 1
            run_count = run_count - 1
        Loop While run_count > 0 And sbi < nsbs
        If full_run And sbi < nsbs Then
            val_ = oc_pack_read1(dec.opb)
            flag = val_
        Else
            flag = 1 - flag
        End If
    Loop
    ' TODO: run_count should be 0 here. If it's not, we should issue a warning of some kind.
    oc_dec_partial_sb_flags_unpack = npartial
End Function

' Decodes the bit flags for whether or not each non-partially-coded super block is fully coded or not. This
' function should only be called if there is at least one non-partially-coded super block.
' Return: The number of partially coded super blocks.
Private Sub oc_dec_coded_sb_flags_unpack(dec As th_dec_ctx)
    Dim nsbs As Long
    Dim sbi As Long
    Dim run_count As Long
    Dim val_ As Long
    Dim flag As Long
    Dim full_run As Boolean
    nsbs = dec.state.nsbs
    ' Skip partially coded super blocks.
    sbi = 0
    Do While dec.state.sb_flags(sbi).coded_partially <> 0
        sbi = sbi + 1
    Loop
    val_ = oc_pack_read1(dec.opb)
    flag = val_
    Do
        run_count = oc_sb_run_unpack(dec.opb)
        full_run = run_count >= 4129
        Do While sbi < nsbs
            If dec.state.sb_flags(sbi).coded_partially = 0 Then
                ' if(run_count--<=0)break;
                If run_count <= 0 Then Exit Do
                run_count = run_count - 1
                dec.state.sb_flags(sbi).coded_fully = flag
            End If
            sbi = sbi + 1
        Loop
        If full_run And sbi < nsbs Then
            val_ = oc_pack_read1(dec.opb)
            flag = val_
        Else
            flag = 1 - flag
        End If
    Loop While sbi < nsbs
    ' TODO: run_count should be 0 here. If it's not, we should issue a warning of some kind.
End Sub

Private Sub oc_dec_coded_flags_unpack(dec As th_dec_ctx)
    Dim nsbs As Long
    Dim sbi As Long
    Dim npartial As Long
    Dim val_ As Long
    Dim pli As Long
    Dim flag As Long
    Dim run_count As Long
    Dim uncoded_fragis As Long      ' ptrdiff_t *: (VB) an index into state.coded_fragis
    Dim ncoded_fragis As Long
    Dim nuncoded_fragis As Long
    Dim prev_ncoded_fragis As Long
    Dim quadi As Long
    Dim quad_coded As Long
    Dim bi As Long
    Dim fragi As Long
    Dim coded As Long
    npartial = oc_dec_partial_sb_flags_unpack(dec)
    If npartial < dec.state.nsbs Then oc_dec_coded_sb_flags_unpack dec
    If npartial > 0 Then
        val_ = oc_pack_read1(dec.opb)
        flag = 1 - val_
    Else
        flag = 0
    End If
    sbi = 0
    nsbs = 0
    run_count = 0
    uncoded_fragis = dec.state.nfrags
    prev_ncoded_fragis = 0
    ncoded_fragis = 0
    nuncoded_fragis = 0
    For pli = 0 To 2
        nsbs = nsbs + dec.state.fplanes(pli).nsbs
        Do While sbi < nsbs
            For quadi = 0 To 3
                If (dec.state.sb_flags(sbi).quad_valid And OC_PB_POW2(quadi)) <> 0 Then
                    quad_coded = 0
                    For bi = 0 To 3
                        fragi = dec.state.sb_maps(sbi * 16 + quadi * 4 + bi)
                        If fragi >= 0 Then
                            If dec.state.sb_flags(sbi).coded_fully <> 0 Then
                                coded = 1
                            ElseIf dec.state.sb_flags(sbi).coded_partially = 0 Then
                                coded = 0
                            Else
                                If run_count <= 0 Then
                                    run_count = oc_block_run_unpack(dec.opb)
                                    flag = 1 - flag
                                End If
                                run_count = run_count - 1
                                coded = flag
                            End If
                            If coded <> 0 Then
                                dec.state.coded_fragis(ncoded_fragis) = fragi
                                ncoded_fragis = ncoded_fragis + 1
                            Else
                                nuncoded_fragis = nuncoded_fragis + 1
                                dec.state.coded_fragis(uncoded_fragis - nuncoded_fragis) = fragi
                            End If
                            quad_coded = quad_coded Or coded
                            dec.state.frags(fragi).coded = coded
                            dec.state.frags(fragi).refi = OC_FRAME_NONE
                        End If
                    Next
                    ' Remember if there's a coded luma block in this macro block.
                    If pli = 0 Then dec.state.mb_modes(sbi * 4 Or quadi) = quad_coded
                End If
            Next
            sbi = sbi + 1
        Loop
        dec.state.ncoded_fragis(pli) = ncoded_fragis - prev_ncoded_fragis
        prev_ncoded_fragis = ncoded_fragis
    Next
    dec.state.ntotal_coded_fragis = ncoded_fragis
    ' TODO: run_count should be 0 here. If it's not, we should issue a warning of some kind.
End Sub

' Coding scheme:
'  Codeword            Mode Index
'  0                       0
'  10                      1
'  110                     2
'  1110                    3
'  11110                   4
'  111110                  5
'  1111110                 6
'  1111111                 7
' (OC_VLC_MODE_TREE[26], OC_CLC_MODE_TREE[9]: theora_tables_vb)

' Unpacks the list of macro block modes for INTER frames.
Private Sub oc_dec_mb_modes_unpack(dec As th_dec_ctx)
    Dim alphabet(0 To 7) As Long    ' const unsigned char *alphabet: (VB) a copy of the alphabet
    Dim nmbs As Long
    Dim mbi As Long
    Dim val_ As Long
    Dim mode_scheme As Long
    Dim mi As Long
    Dim token As Long
    val_ = oc_pack_read(dec.opb, 3)
    mode_scheme = val_
    If mode_scheme = 0 Then
        ' Just in case, initialize the modes to something. If the bitstream doesn't contain each index exactly
        ' once, it's likely corrupt and the rest of the packet is garbage anyway, but this way we won't crash,
        ' and we'll decode SOMETHING.
        ' LOOP VECTORIZES
        For mi = 0 To OC_NMODES - 1
            alphabet(mi) = OC_MODE_INTER_NOMV
        Next
        For mi = 0 To OC_NMODES - 1
            val_ = oc_pack_read(dec.opb, 3)
            alphabet(val_) = OC_MODE_ALPHABETS(6 * OC_NMODES + mi)
        Next
    Else
        For mi = 0 To OC_NMODES - 1
            alphabet(mi) = OC_MODE_ALPHABETS((mode_scheme - 1) * OC_NMODES + mi)
        Next
    End If
    ' mode_tree=mode_scheme==7?OC_CLC_MODE_TREE:OC_VLC_MODE_TREE
    nmbs = dec.state.nmbs
    For mbi = 0 To nmbs - 1
        If dec.state.mb_modes(mbi) > 0 Then
            ' We have a coded luma block; decode a mode.
            If mode_scheme = 7 Then
                token = oc_huff_token_decode(dec.opb, OC_CLC_MODE_TREE, 0)
            Else
                token = oc_huff_token_decode(dec.opb, OC_VLC_MODE_TREE, 0)
            End If
            dec.state.mb_modes(mbi) = alphabet(token)
        End If
        ' For other valid macro blocks, INTER_NOMV is forced, but we rely on the fact that OC_MODE_INTER_NOMV
        ' is already 0.
    Next
End Sub

' (OC_VLC_MV_COMP_TREE[101], OC_CLC_MV_COMP_TREE[65]: theora_tables_vb)

' (VB) _tree: OC_CLC_MV_COMP_TREE when clc, else OC_VLC_MV_COMP_TREE
Private Function oc_mv_unpack(opb As oc_pack_buf, ByVal clc As Boolean) As Long
    Dim dx As Long
    Dim dy As Long
    If clc Then
        dx = oc_huff_token_decode(opb, OC_CLC_MV_COMP_TREE, 0) - 32
        dy = oc_huff_token_decode(opb, OC_CLC_MV_COMP_TREE, 0) - 32
    Else
        dx = oc_huff_token_decode(opb, OC_VLC_MV_COMP_TREE, 0) - 32
        dy = oc_huff_token_decode(opb, OC_VLC_MV_COMP_TREE, 0) - 32
    End If
    oc_mv_unpack = OC_MV(dx, dy)
End Function

' Unpacks the list of motion vectors for INTER frames, and propagtes the macro block modes and motion vectors
' to the individual fragments.
Private Sub oc_dec_mv_unpack_and_frag_modes_fill(dec As th_dec_ctx)
    Dim pixel_fmt As Long           ' set_chroma_mvs=OC_SET_CHROMA_MVS_TABLE[pixel_fmt]
    Dim clc As Boolean              ' mv_comp_tree
    Dim map_idxs As Long            ' const unsigned char *map_idxs: (VB) OC_MB_MAP_IDXS from map_idxs
    Dim map_nidxs As Long
    Dim last_mv As Long
    Dim prior_mv As Long
    Dim cbmvs(0 To 3) As Long
    Dim lbmvs(0 To 3) As Long
    Dim nmbs As Long
    Dim mbi As Long
    Dim val_ As Long
    Dim mb_mode As Long
    Dim mbmv As Long
    Dim fragi As Long
    Dim mapi As Long
    Dim mapii As Long
    Dim refi As Long
    Dim bi As Long
    pixel_fmt = dec.state.info.pixel_fmt
    val_ = oc_pack_read1(dec.opb)
    clc = val_ <> 0
    map_idxs = pixel_fmt * 12
    map_nidxs = OC_MB_MAP_NIDXS(pixel_fmt)
    last_mv = 0
    prior_mv = 0
    nmbs = dec.state.nmbs
    For mbi = 0 To nmbs - 1
        mb_mode = dec.state.mb_modes(mbi)
        If mb_mode <> OC_MODE_INVALID Then
            If mb_mode = OC_MODE_INTER_MV_FOUR Then
                prior_mv = last_mv
                For bi = 0 To 3
                    fragi = dec.state.mb_maps(mbi * 12 + bi)
                    If dec.state.frags(fragi).coded <> 0 Then
                        dec.state.frags(fragi).refi = OC_FRAME_PREV
                        dec.state.frags(fragi).mb_mode = OC_MODE_INTER_MV_FOUR
                        last_mv = oc_mv_unpack(dec.opb, clc)
                        lbmvs(bi) = last_mv
                        dec.state.frag_mvs(fragi) = lbmvs(bi)
                    Else
                        lbmvs(bi) = 0
                    End If
                Next
                OC_SET_CHROMA_MVS pixel_fmt, cbmvs, lbmvs
                For mapii = 4 To map_nidxs - 1
                    mapi = OC_MB_MAP_IDXS(map_idxs + mapii)
                    bi = mapi And 3
                    fragi = dec.state.mb_maps(mbi * 12 + (mapi \ 4) * 4 + bi)
                    If dec.state.frags(fragi).coded <> 0 Then
                        dec.state.frags(fragi).refi = OC_FRAME_PREV
                        dec.state.frags(fragi).mb_mode = OC_MODE_INTER_MV_FOUR
                        dec.state.frag_mvs(fragi) = cbmvs(bi)
                    End If
                Next
            Else
                Select Case mb_mode
                    Case OC_MODE_INTER_MV
                        prior_mv = last_mv
                        mbmv = oc_mv_unpack(dec.opb, clc)
                        last_mv = mbmv
                    Case OC_MODE_INTER_MV_LAST
                        mbmv = last_mv
                    Case OC_MODE_INTER_MV_LAST2
                        mbmv = prior_mv
                        prior_mv = last_mv
                        last_mv = mbmv
                    Case OC_MODE_GOLDEN_MV
                        mbmv = oc_mv_unpack(dec.opb, clc)
                    Case Else
                        mbmv = 0
                End Select
                ' Fill in the MVs for the fragments.
                refi = OC_FRAME_FOR_MODE(mb_mode)
                mapii = 0
                Do
                    mapi = OC_MB_MAP_IDXS(map_idxs + mapii)
                    fragi = dec.state.mb_maps(mbi * 12 + (mapi \ 4) * 4 + (mapi And 3))
                    If dec.state.frags(fragi).coded <> 0 Then
                        dec.state.frags(fragi).refi = refi
                        dec.state.frags(fragi).mb_mode = mb_mode
                        dec.state.frag_mvs(fragi) = mbmv
                    End If
                    mapii = mapii + 1
                Loop While mapii < map_nidxs
            End If
        End If
    Next
End Sub

Private Sub oc_dec_block_qis_unpack(dec As th_dec_ctx)
    Dim ncoded_fragis As Long
    Dim fragii As Long
    Dim fragi As Long
    Dim val_ As Long
    Dim flag As Long
    Dim nqi1 As Long
    Dim run_count As Long
    Dim full_run As Boolean
    ncoded_fragis = dec.state.ntotal_coded_fragis
    If ncoded_fragis <= 0 Then Exit Sub
    If dec.state.nqis = 1 Then
        ' If this frame has only a single qi value, then just use it for all coded fragments.
        For fragii = 0 To ncoded_fragis - 1
            dec.state.frags(dec.state.coded_fragis(fragii)).qii = 0
        Next
    Else
        ' Otherwise, we decode a qi index for each fragment, using two passes of the same binary RLE scheme used
        ' for super-block coded bits. The first pass marks each fragment as having a qii of 0 or greater than
        ' 0, and the second pass (if necessary), distinguishes between a qii of 1 and 2. At first we just store
        ' the qii in the fragment. After all the qii's are decoded, we make a final pass to replace them with
        ' the corresponding qi's for this frame.
        val_ = oc_pack_read1(dec.opb)
        flag = val_
        nqi1 = 0
        fragii = 0
        Do While fragii < ncoded_fragis
            run_count = oc_sb_run_unpack(dec.opb)
            full_run = run_count >= 4129
            Do
                dec.state.frags(dec.state.coded_fragis(fragii)).qii = flag
                fragii = fragii + 1
                nqi1 = nqi1 + flag
                run_count = run_count - 1
            Loop While run_count > 0 And fragii < ncoded_fragis
            If full_run And fragii < ncoded_fragis Then
                val_ = oc_pack_read1(dec.opb)
                flag = val_
            Else
                flag = 1 - flag
            End If
        Loop
        ' TODO: run_count should be 0 here. If it's not, we should issue a warning of some kind.
        ' If we have 3 different qi's for this frame, and there was at least one fragment with a non-zero qi,
        ' make the second pass.
        If dec.state.nqis = 3 And nqi1 > 0 Then
            ' Skip qii==0 fragments.
            fragii = 0
            Do While dec.state.frags(dec.state.coded_fragis(fragii)).qii = 0
                fragii = fragii + 1
            Loop
            val_ = oc_pack_read1(dec.opb)
            flag = val_
            Do
                run_count = oc_sb_run_unpack(dec.opb)
                full_run = run_count >= 4129
                Do While fragii < ncoded_fragis
                    fragi = dec.state.coded_fragis(fragii)
                    If dec.state.frags(fragi).qii <> 0 Then
                        ' if(run_count--<=0)break;
                        If run_count <= 0 Then Exit Do
                        run_count = run_count - 1
                        dec.state.frags(fragi).qii = dec.state.frags(fragi).qii + flag
                    End If
                    fragii = fragii + 1
                Loop
                If full_run And fragii < ncoded_fragis Then
                    val_ = oc_pack_read1(dec.opb)
                    flag = val_
                Else
                    flag = 1 - flag
                End If
            Loop While fragii < ncoded_fragis
            ' TODO: run_count should be 0 here. If it's not, we should issue a warning of some kind.
        End If
    End If
End Sub

' Unpacks the DC coefficient tokens. Unlike when unpacking the AC coefficient tokens, we actually need to
' decode the DC coefficient values now so that we can do DC prediction.
'  _huff_idx:   The index of the Huffman table to use for each color plane.
'  _ntoks_left: The number of tokens left to be decoded in each color plane for each coefficient. This is
'                updated as EOB tokens and zero run tokens are decoded. (VB) ntoks_left(pli * 64 + zzi)
'  Return: The length of any outstanding EOB run.
Private Function oc_dec_dc_coeff_unpack(dec As th_dec_ctx, huff_idxs() As Long, ntoks_left() As Long) As Long
    Dim ncoded_fragis As Long
    Dim fragii As Long
    Dim eobs As Long
    Dim ti As Long
    Dim pli As Long
    Dim run_counts(0 To 63) As Long
    Dim eob_count As Long
    Dim eobi As Long
    Dim rli As Long
    Dim token As Long
    Dim cw As Long
    Dim eb As Long
    Dim skip As Long
    Dim coeff As Long
    Dim tree As Long
    ncoded_fragis = 0
    fragii = 0
    eobs = 0
    ti = 0
    For pli = 0 To 2
        ncoded_fragis = ncoded_fragis + dec.state.ncoded_fragis(pli)
        For rli = 0 To 63
            run_counts(rli) = 0
        Next
        dec.eob_runs(pli * 64) = eobs
        dec.ti0(pli * 64) = ti
        ' Continue any previous EOB run, if there was one.
        eobi = eobs
        If ncoded_fragis - fragii < eobi Then eobi = ncoded_fragis - fragii
        eob_count = eobi
        eobs = eobs - eobi
        Do While eobi > 0
            eobi = eobi - 1
            dec.state.frags(dec.state.coded_fragis(fragii)).dc = 0
            fragii = fragii + 1
        Loop
        tree = dec.huff_tables(huff_idxs((pli + 1) \ 2))
        Do While fragii < ncoded_fragis
            token = oc_huff_token_decode(dec.opb, dec.vb_huff, tree)
            dec.dct_tokens(ti) = token
            ti = ti + 1
            If token < OC_NINTERNAL_DCT_TOKEN_EXTRA_BITS Then
                eb = oc_pack_read(dec.opb, OC_INTERNAL_DCT_TOKEN_EXTRA_BITS(token))
                dec.dct_tokens(ti) = eb And 255
                ti = ti + 1
                If token = OC_DCT_TOKEN_FAT_EOB Then
                    dec.dct_tokens(ti) = (eb \ 256) And 255
                    ti = ti + 1
                End If
                eb = eb * OC_DCT_TOKEN_EB_MUL(token)
            Else
                eb = 0
            End If
            cw = OC_DCT_CODE_WORD(token) + eb
            eobs = (cw And &HFFF00) \ &H100&       ' cw>>OC_DCT_CW_EOB_SHIFT&0xFFF
            If cw = OC_DCT_CW_FINISH Then eobs = OC_DCT_EOB_FINISH
            If eobs <> 0 Then
                eobi = ncoded_fragis - fragii
                If eobs < eobi Then eobi = eobs
                eob_count = eob_count + eobi
                eobs = eobs - eobi
                Do While eobi > 0
                    eobi = eobi - 1
                    dec.state.frags(dec.state.coded_fragis(fragii)).dc = 0
                    fragii = fragii + 1
                Loop
            Else
                skip = cw And 255                   ' (unsigned char)(cw>>OC_DCT_CW_RLEN_SHIFT)
                cw = cw Xor -(cw And &H100000)      ' cw^=-(cw&1<<OC_DCT_CW_FLIP_BIT)
                coeff = (cw And &HFFE00000) \ &H200000  ' cw>>OC_DCT_CW_MAG_SHIFT
                If skip <> 0 Then coeff = 0
                run_counts(skip) = run_counts(skip) + 1
                dec.state.frags(dec.state.coded_fragis(fragii)).dc = coeff
                fragii = fragii + 1
            End If
        Loop
        ' Add the total EOB count to the longest run length.
        run_counts(63) = run_counts(63) + eob_count
        ' And convert the run_counts array to a moment table.
        For rli = 62 To 0 Step -1
            run_counts(rli) = run_counts(rli) + run_counts(rli + 1)
        Next
        ' Finally, subtract off the number of coefficients that have been accounted for by runs started in this
        ' coefficient.
        For rli = 63 To 0 Step -1
            ntoks_left(pli * 64 + rli) = ntoks_left(pli * 64 + rli) - run_counts(rli)
        Next
    Next
    dec.dct_tokens_count = ti
    oc_dec_dc_coeff_unpack = eobs
End Function

' Unpacks the AC coefficient tokens. This can completely discard coefficient values while unpacking, and so is
' somewhat simpler than unpacking the DC coefficient tokens.
'  _huff_idx:   The index of the Huffman table to use for each color plane.
'  _ntoks_left: The number of tokens left to be decoded in each color plane for each coefficient. This is
'                updated as EOB tokens and zero run tokens are decoded.
'  _eobs:       The length of any outstanding EOB run from previous coefficients.
'  Return: The length of any outstanding EOB run.
Private Function oc_dec_ac_coeff_unpack(dec As th_dec_ctx, ByVal zzi As Long, huff_idxs() As Long, _
    ntoks_left_() As Long, ByVal eobs As Long) As Long
    Dim ti As Long
    Dim pli As Long
    Dim run_counts(0 To 63) As Long
    Dim eob_count As Long
    Dim ntoks_left As Long
    Dim ntoks As Long
    Dim rli As Long
    Dim token As Long
    Dim cw As Long
    Dim eb As Long
    Dim skip As Long
    Dim tree As Long
    ti = dec.dct_tokens_count
    For pli = 0 To 2
        dec.eob_runs(pli * 64 + zzi) = eobs
        dec.ti0(pli * 64 + zzi) = ti
        ntoks_left = ntoks_left_(pli * 64 + zzi)
        For rli = 0 To 63
            run_counts(rli) = 0
        Next
        eob_count = 0
        ntoks = 0
        tree = dec.huff_tables(huff_idxs((pli + 1) \ 2))
        Do While eobs < ntoks_left - ntoks      ' while(ntoks+_eobs<ntoks_left)
            ntoks = ntoks + eobs
            eob_count = eob_count + eobs
            token = oc_huff_token_decode(dec.opb, dec.vb_huff, tree)
            dec.dct_tokens(ti) = token
            ti = ti + 1
            If token < OC_NINTERNAL_DCT_TOKEN_EXTRA_BITS Then
                eb = oc_pack_read(dec.opb, OC_INTERNAL_DCT_TOKEN_EXTRA_BITS(token))
                dec.dct_tokens(ti) = eb And 255
                ti = ti + 1
                If token = OC_DCT_TOKEN_FAT_EOB Then
                    dec.dct_tokens(ti) = (eb \ 256) And 255
                    ti = ti + 1
                End If
                eb = eb * OC_DCT_TOKEN_EB_MUL(token)
            Else
                eb = 0
            End If
            cw = OC_DCT_CODE_WORD(token) + eb
            skip = cw And 255
            eobs = (cw And &HFFF00) \ &H100&
            If cw = OC_DCT_CW_FINISH Then eobs = OC_DCT_EOB_FINISH
            If eobs = 0 Then
                run_counts(skip) = run_counts(skip) + 1
                ntoks = ntoks + 1
            End If
        Loop
        ' Add the portion of the last EOB run actually used by this coefficient.
        eob_count = eob_count + (ntoks_left - ntoks)
        ' And remove it from the remaining EOB count.
        eobs = eobs - (ntoks_left - ntoks)
        ' Add the total EOB count to the longest run length.
        run_counts(63) = run_counts(63) + eob_count
        ' And convert the run_counts array to a moment table.
        For rli = 62 To 0 Step -1
            run_counts(rli) = run_counts(rli) + run_counts(rli + 1)
        Next
        ' Finally, subtract off the number of coefficients that have been accounted for by runs started in this
        ' coefficient.
        For rli = 63 - zzi To 0 Step -1
            ntoks_left_(pli * 64 + zzi + rli) = ntoks_left_(pli * 64 + zzi + rli) - run_counts(rli)
        Next
    Next
    dec.dct_tokens_count = ti
    oc_dec_ac_coeff_unpack = eobs
End Function

' Tokens describing the DCT coefficients that belong to each fragment are stored in the bitstream grouped by
' coefficient, not by fragment.
' This means that we either decode all the tokens in order, building up a separate coefficient list for each
' fragment as we go, and then go back and do the iDCT on each fragment, or we have to create separate lists of
' tokens for each coefficient, so that we can pull the next token required off the head of the appropriate
' list when decoding a specific fragment.
' The former was VP3's choice, and it meant 2*w*h extra storage for all the decoded coefficient values.
' We take the second option, which lets us store just one to three bytes per token (generally far fewer than
' the number of coefficients, due to EOB tokens and zero runs), and which requires us to only maintain a
' counter for each of the 64 coefficients, instead of a counter for every fragment to determine where the next
' token goes.
' We actually use 3 counters per coefficient, one for each color plane, so we can decode all color planes
' simultaneously. This lets color conversion, etc., be done as soon as a full MCU (one or two super block rows)
' is decoded, while the image data is still in cache.
Private Sub oc_dec_residual_tokens_unpack(dec As th_dec_ctx)
    ' (OC_HUFF_LIST_MAX[5]={1,6,15,28,64}: theora_tables_vb)
    Dim ntoks_left(0 To 191) As Long    ' ptrdiff_t ntoks_left[3][64]
    Dim huff_idxs(0 To 1) As Long
    Dim eobs As Long
    Dim val_ As Long
    Dim pli As Long
    Dim zzi As Long
    Dim hgi As Long
    For pli = 0 To 2
        For zzi = 0 To 63
            ntoks_left(pli * 64 + zzi) = dec.state.ncoded_fragis(pli)
        Next
    Next
    val_ = oc_pack_read(dec.opb, 4)
    huff_idxs(0) = val_
    val_ = oc_pack_read(dec.opb, 4)
    huff_idxs(1) = val_
    dec.eob_runs(0) = 0
    eobs = oc_dec_dc_coeff_unpack(dec, huff_idxs, ntoks_left)
    val_ = oc_pack_read(dec.opb, 4)
    huff_idxs(0) = val_
    val_ = oc_pack_read(dec.opb, 4)
    huff_idxs(1) = val_
    zzi = 1
    For hgi = 1 To 4
        huff_idxs(0) = huff_idxs(0) + 16
        huff_idxs(1) = huff_idxs(1) + 16
        Do While zzi < OC_HUFF_LIST_MAX(hgi)
            eobs = oc_dec_ac_coeff_unpack(dec, zzi, huff_idxs, ntoks_left, eobs)
            zzi = zzi + 1
        Loop
    Next
    ' TODO: eobs should be exactly zero, or 4096 or greater. The second case occurs when an EOB run of size zero
    ' is encountered, which gets treated as an infinite EOB run (where infinity is PTRDIFF_MAX). If neither of
    ' these conditions holds, then a warning should be issued.
End Sub

Private Function oc_dec_postprocess_init(dec As th_dec_ctx) As Long
    Dim ncoded_fragis As Long
    Dim fragii As Long
    Dim qi0 As Long
    Dim y_sz As Long
    Dim c_sz As Long
    Dim c_w As Long
    Dim c_h As Long
    Dim ppState As Long
    ' musl libc malloc()/realloc() calls might use floating point, so make sure we've cleared the MMX state for
    ' them.
    oc_restore_fpu_c
    ' pp_level 0: disabled; free any memory used and return
    If dec.pp_level <= OC_PP_LEVEL_DISABLED Then
        If dec.vbDcQisAllocated Then
            Erase dec.dc_qis
            dec.vbDcQisAllocated = False
            Erase dec.variances
            dec.vbVariancesAllocated = False
            dec.pp_frame_data = 0
        End If
        oc_dec_postprocess_init = 1
        Exit Function
    End If
    If Not dec.vbDcQisAllocated Then
        ' If we haven't been tracking DC quantization indices, there's no point in starting now.
        If dec.state.frame_type <> OC_INTRA_FRAME Then oc_dec_postprocess_init = 1: Exit Function
        ReDim dec.dc_qis(0 To dec.state.nfrags - 1)
        dec.vbDcQisAllocated = True
        FillMemory dec.dc_qis(0), dec.state.nfrags, CByte(dec.state.qis(0))
    Else
        ' Update the DC quantization index of each coded block.
        ncoded_fragis = dec.state.ncoded_fragis(0) + dec.state.ncoded_fragis(1) + dec.state.ncoded_fragis(2)
        qi0 = dec.state.qis(0)
        For fragii = 0 To ncoded_fragis - 1
            dec.dc_qis(dec.state.coded_fragis(fragii)) = qi0
        Next
    End If
    ' pp_level 1: Stop after updating DC quantization indices.
    If dec.pp_level <= OC_PP_LEVEL_TRACKDCQI Then
        If dec.vbVariancesAllocated Then
            Erase dec.variances
            dec.vbVariancesAllocated = False
            dec.pp_frame_data = 0
        End If
        oc_dec_postprocess_init = 1
        Exit Function
    End If
    If Not dec.vbVariancesAllocated Then
        ' Allocate space for the chroma planes, even if we're not going to use them; this simplifies allocation
        ' state management, though it may waste memory on the few systems that don't overcommit pages.
        ' (VB) the frame_sz bytes were reserved after the reference frames by oc_dec_init
        dec.pp_frame_data = dec.state.vb_ref_frame_sz
        ReDim dec.variances(0 To dec.state.nfrags - 1)
        dec.vbVariancesAllocated = True
        ' Force an update of the PP buffer pointers.
        dec.pp_frame_state = 0
    End If
    ' Update the PP buffer pointers if necessary.
    If dec.pp_level >= OC_PP_LEVEL_DEBLOCKC Then ppState = 2 Else ppState = 1
    If dec.pp_frame_state <> ppState Then
        If dec.pp_level < OC_PP_LEVEL_DEBLOCKC Then
            ' If chroma processing is disabled, just use the PP luma plane.
            dec.pp_frame_buf.p(0).width = dec.state.info.frame_width
            dec.pp_frame_buf.p(0).height = dec.state.info.frame_height
            dec.pp_frame_buf.p(0).stride = -dec.pp_frame_buf.p(0).width
            dec.pp_frame_buf.p(0).data = dec.pp_frame_data + (1 - dec.pp_frame_buf.p(0).height) * dec.pp_frame_buf.p(0).stride
        Else
            ' Otherwise, set up pointers to all three PP planes.
            y_sz = dec.state.info.frame_width * dec.state.info.frame_height
            If (dec.state.info.pixel_fmt And 1) = 0 Then c_w = dec.state.info.frame_width \ 2 Else c_w = dec.state.info.frame_width
            If (dec.state.info.pixel_fmt And 2) = 0 Then c_h = dec.state.info.frame_height \ 2 Else c_h = dec.state.info.frame_height
            c_sz = c_w * c_h
            dec.pp_frame_buf.p(0).width = dec.state.info.frame_width
            dec.pp_frame_buf.p(0).height = dec.state.info.frame_height
            dec.pp_frame_buf.p(0).stride = dec.pp_frame_buf.p(0).width
            dec.pp_frame_buf.p(0).data = dec.pp_frame_data
            dec.pp_frame_buf.p(1).width = c_w
            dec.pp_frame_buf.p(1).height = c_h
            dec.pp_frame_buf.p(1).stride = dec.pp_frame_buf.p(1).width
            dec.pp_frame_buf.p(1).data = dec.pp_frame_buf.p(0).data + y_sz
            dec.pp_frame_buf.p(2).width = c_w
            dec.pp_frame_buf.p(2).height = c_h
            dec.pp_frame_buf.p(2).stride = dec.pp_frame_buf.p(2).width
            dec.pp_frame_buf.p(2).data = dec.pp_frame_buf.p(1).data + c_sz
            oc_ycbcr_buffer_flip dec.pp_frame_buf, dec.pp_frame_buf
        End If
        dec.pp_frame_state = ppState
    End If
    ' If we're not processing chroma, copy the reference frame's chroma planes.
    If dec.pp_level < OC_PP_LEVEL_DEBLOCKC Then
        dec.pp_frame_buf.p(1) = dec.state.ref_frame_bufs(dec.state.ref_frame_idx(OC_FRAME_SELF)).p(1)
        dec.pp_frame_buf.p(2) = dec.state.ref_frame_bufs(dec.state.ref_frame_idx(OC_FRAME_SELF)).p(2)
    End If
    oc_dec_postprocess_init = 0
End Function

' Initialize the main decoding pipeline.
Private Sub oc_dec_pipeline_init(dec As th_dec_ctx, pipe As oc_dec_pipeline_state)
    Dim coded_fragis As Long        ' const ptrdiff_t *: (VB) indices into state.coded_fragis
    Dim uncoded_fragis As Long
    Dim flimit As Long
    Dim pli As Long
    Dim qii As Long
    Dim qti As Long
    Dim zzi As Long
    Dim ncoded_fragis As Long
    ' If chroma is sub-sampled in the vertical direction, we have to decode two super block rows of Y' for each
    ' super block row of Cb and Cr.
    If (dec.state.info.pixel_fmt And 2) = 0 Then pipe.mcu_nvfrags = 8 Else pipe.mcu_nvfrags = 4
    ' Initialize the token and extra bits indices for each plane and coefficient.
    ' Also copy over the initial the EOB run counts.
    For zzi = 0 To 191
        pipe.ti(zzi) = dec.ti0(zzi)
        pipe.eob_runs(zzi) = dec.eob_runs(zzi)
    Next
    ' Set up per-plane pointers to the coded and uncoded fragments lists.
    coded_fragis = 0
    uncoded_fragis = dec.state.nfrags
    For pli = 0 To 2
        pipe.coded_fragis(pli) = coded_fragis
        pipe.uncoded_fragis(pli) = uncoded_fragis
        ncoded_fragis = dec.state.ncoded_fragis(pli)
        coded_fragis = coded_fragis + ncoded_fragis
        uncoded_fragis = uncoded_fragis + ncoded_fragis - dec.state.fplanes(pli).nfrags
    Next
    ' Set up condensed quantizer tables.
    For pli = 0 To 2
        For qii = 0 To dec.state.nqis - 1
            For qti = 0 To 1
                pipe.dequant((pli * 3 + qii) * 2 + qti) = dec.state.dequant_tables((dec.state.qis(qii) * 3 + pli) * 2 + qti)
            Next
        Next
    Next
    ' Set the previous DC predictor to 0 for all color planes and frame types.
    For pli = 0 To 11
        pipe.pred_last(pli) = 0
    Next
    ' Initialize the bounding value array for the loop filter.
    flimit = dec.state.loop_filter_limits(dec.state.qis(0))
    If flimit <> 0 Then pipe.loop_filter = 1 Else pipe.loop_filter = 0
    If flimit <> 0 Then oc_loop_filter_init_c pipe.bounding_values, flimit
    ' Initialize any buffers needed for post-processing. We also save the current post-processing level, to
    ' guard against the user changing it from a callback.
    If oc_dec_postprocess_init(dec) = 0 Then
        pipe.pp_level = dec.pp_level
    ' If we don't have enough information to post-process, disable it, regardless of the user-requested level.
    Else
        pipe.pp_level = OC_PP_LEVEL_DISABLED
        dec.pp_frame_buf = dec.state.ref_frame_bufs(dec.state.ref_frame_idx(OC_FRAME_SELF))
    End If
    ' Clear down the DCT coefficient buffer for the first block.
    For zzi = 0 To 63
        pipe.dct_coeffs(zzi) = 0
    Next
End Sub

' Undo the DC prediction in a single plane of an MCU (one or two super block rows). As a side effect, the
' number of coded and uncoded fragments in this plane of the MCU is also computed.
' (oc_dec_dc_unpredict_mcu_plane: decint.h maps it to the _c version)
Public Sub oc_dec_dc_unpredict_mcu_plane_c(dec As th_dec_ctx, pipe As oc_dec_pipeline_state, ByVal pli As Long)
    Dim pred_last As Long           ' int *pred_last: (VB) an index into pipe.pred_last
    Dim ncoded_fragis As Long
    Dim fragi As Long
    Dim fragx As Long
    Dim fragy As Long
    Dim fragy0 As Long
    Dim fragy_end As Long
    Dim nhfrags As Long
    Dim l_ref As Long
    Dim ul_ref As Long
    Dim u_ref As Long
    Dim ur_ref As Long
    Dim pred As Long
    Dim refi As Long
    Dim sel As Long
    Dim p0 As Long
    Dim p1 As Long
    Dim p2 As Long
    Dim v As Long
    ' Compute the first and last fragment row of the current MCU for this plane.
    fragy0 = pipe.fragy0(pli)
    fragy_end = pipe.fragy_end(pli)
    nhfrags = dec.state.fplanes(pli).nhfrags
    pred_last = pli * 4
    ncoded_fragis = 0
    fragi = dec.state.fplanes(pli).froffset + fragy0 * nhfrags
    For fragy = fragy0 To fragy_end - 1
        If fragy = 0 Then
            ' For the first row, all of the cases reduce to just using the previous predictor for the same
            ' reference frame.
            For fragx = 0 To nhfrags - 1
                If dec.state.frags(fragi).coded <> 0 Then
                    refi = dec.state.frags(fragi).refi
                    ' pred_last[refi]=frags[fragi].dc+=pred_last[refi] (dc: 16 bits)
                    v = (((dec.state.frags(fragi).dc + pipe.pred_last(pred_last + refi)) + 32768) And &HFFFF&) - 32768
                    dec.state.frags(fragi).dc = v
                    pipe.pred_last(pred_last + refi) = v
                    ncoded_fragis = ncoded_fragis + 1
                End If
                fragi = fragi + 1
            Next
        Else
            ' u_frags=frags-nhfrags: (VB) u_frags[i] is frags(i - nhfrags)
            l_ref = -1
            ul_ref = -1
            u_ref = dec.state.frags(fragi - nhfrags).refi
            For fragx = 0 To nhfrags - 1
                If fragx + 1 >= nhfrags Then
                    ur_ref = -1
                Else
                    ur_ref = dec.state.frags(fragi - nhfrags + 1).refi
                End If
                If dec.state.frags(fragi).coded <> 0 Then
                    refi = dec.state.frags(fragi).refi
                    ' We break out a separate case based on which of our neighbors use the same reference frames.
                    ' This is somewhat faster than trying to make a generic case which handles all of them, since
                    ' it reduces lots of poorly predicted jumps to one switch statement, and also lets a number of
                    ' the multiplications be optimized out by strength reduction.
                    sel = 0
                    If l_ref = refi Then sel = 1
                    If ul_ref = refi Then sel = sel Or 2
                    If u_ref = refi Then sel = sel Or 4
                    If ur_ref = refi Then sel = sel Or 8
                    Select Case sel
                        Case 1, 3
                            pred = dec.state.frags(fragi - 1).dc
                        Case 2
                            pred = dec.state.frags(fragi - nhfrags - 1).dc
                        Case 4, 6, 12
                            pred = dec.state.frags(fragi - nhfrags).dc
                        Case 5
                            pred = (dec.state.frags(fragi - 1).dc + dec.state.frags(fragi - nhfrags).dc) \ 2
                        Case 8
                            pred = dec.state.frags(fragi - nhfrags + 1).dc
                        Case 9, 11, 13
                            ' The TI compiler mis-compiles this line.
                            pred = (75 * dec.state.frags(fragi - 1).dc + 53 * dec.state.frags(fragi - nhfrags + 1).dc) \ 128
                        Case 10
                            pred = (dec.state.frags(fragi - nhfrags - 1).dc + dec.state.frags(fragi - nhfrags + 1).dc) \ 2
                        Case 14
                            pred = (3 * (dec.state.frags(fragi - nhfrags - 1).dc + dec.state.frags(fragi - nhfrags + 1).dc) _
                                + 10 * dec.state.frags(fragi - nhfrags).dc) \ 16
                        Case 7, 15
                            p0 = dec.state.frags(fragi - 1).dc
                            p1 = dec.state.frags(fragi - nhfrags - 1).dc
                            p2 = dec.state.frags(fragi - nhfrags).dc
                            pred = (29 * (p0 + p2) - 26 * p1) \ 32
                            If Abs(pred - p2) > 128 Then
                                pred = p2
                            ElseIf Abs(pred - p0) > 128 Then
                                pred = p0
                            ElseIf Abs(pred - p1) > 128 Then
                                pred = p1
                            End If
                        Case Else
                            pred = pipe.pred_last(pred_last + refi)
                    End Select
                    v = (((dec.state.frags(fragi).dc + pred) + 32768) And &HFFFF&) - 32768
                    dec.state.frags(fragi).dc = v
                    pipe.pred_last(pred_last + refi) = v
                    ncoded_fragis = ncoded_fragis + 1
                    l_ref = refi
                Else
                    l_ref = -1
                End If
                ul_ref = u_ref
                u_ref = ur_ref
                fragi = fragi + 1
            Next
        End If
    Next
    pipe.ncoded_fragis(pli) = ncoded_fragis
    ' Also save the number of uncoded fragments so we know how many to copy.
    pipe.nuncoded_fragis(pli) = (fragy_end - fragy0) * nhfrags - ncoded_fragis
End Sub

' Reconstructs all coded fragments in a single MCU (one or two super block rows). This requires that each coded
' fragment have a proper macro block mode and motion vector (if not in INTRA mode), and have its DC value
' decoded, with the DC prediction process reversed, and the number of coded and uncoded fragments in this
' plane of the MCU be counted. The token lists for each color plane and coefficient should also be filled in,
' along with initial token offsets, extra bits offsets, and EOB run counts.
Private Sub oc_dec_frags_recon_mcu_plane(dec As th_dec_ctx, pipe As oc_dec_pipeline_state, ByVal pli As Long)
    Dim dc_quant(0 To 1) As Long
    Dim coded_fragis As Long        ' const ptrdiff_t *: (VB) an index into state.coded_fragis
    Dim ncoded_fragis As Long
    Dim fragii As Long
    Dim ti As Long                  ' ptrdiff_t *ti: (VB) the plane's first index in pipe.ti and pipe.eob_runs
    Dim qti As Long
    Dim ac_quant As Long            ' const ogg_uint16_t *: (VB) an offset into state.dequant_table_data
    Dim fragi As Long
    Dim last_zzi As Long
    Dim zzi As Long
    Dim token As Long
    Dim eob As Long
    Dim cw As Long
    Dim rlen As Long
    Dim coeff As Long
    Dim lti As Long
    ' dct_fzig_zag=_dec->state.opt_data.dct_fzig_zag: (VB) OC_FZIG_ZAG
    coded_fragis = pipe.coded_fragis(pli)
    ncoded_fragis = pipe.ncoded_fragis(pli)
    ti = pli * 64
    For qti = 0 To 1
        dc_quant(qti) = dec.state.dequant_table_data(pipe.dequant((pli * 3 + 0) * 2 + qti))
    Next
    For fragii = 0 To ncoded_fragis - 1
        fragi = dec.state.coded_fragis(coded_fragis + fragii)
        If dec.state.frags(fragi).mb_mode <> OC_MODE_INTRA Then qti = 1 Else qti = 0
        ac_quant = pipe.dequant((pli * 3 + dec.state.frags(fragi).qii) * 2 + qti)
        ' Decode the AC coefficients.
        zzi = 0
        Do While zzi < 64
            last_zzi = zzi
            If pipe.eob_runs(ti + zzi) <> 0 Then
                pipe.eob_runs(ti + zzi) = pipe.eob_runs(ti + zzi) - 1
                Exit Do
            Else
                lti = pipe.ti(ti + zzi)
                token = dec.dct_tokens(lti)
                lti = lti + 1
                cw = OC_DCT_CODE_WORD(token)
                ' These parts could be done branchless, but the branches are fairly predictable and the C code
                ' translates into more than a few instructions, so it's worth it to avoid them.
                If token < OC_NINTERNAL_DCT_TOKEN_EXTRA_BITS Then
                    cw = cw + dec.dct_tokens(lti) * OC_DCT_TOKEN_EB_MUL(token)
                    lti = lti + 1
                End If
                eob = (cw And &HFFF00) \ &H100&
                If token = OC_DCT_TOKEN_FAT_EOB Then
                    eob = eob + dec.dct_tokens(lti) * &H100&
                    lti = lti + 1
                    If eob = 0 Then eob = OC_DCT_EOB_FINISH
                End If
                rlen = cw And 255
                cw = cw Xor -(cw And &H100000)
                coeff = (cw And &HFFE00000) \ &H200000
                pipe.eob_runs(ti + zzi) = eob
                pipe.ti(ti + zzi) = lti
                zzi = zzi + rlen
                pipe.dct_coeffs(OC_FZIG_ZAG(zzi)) = ((coeff * dec.state.dequant_table_data(ac_quant + zzi) + 32768) And &HFFFF&) - 32768
                If eob = 0 Then zzi = zzi + 1
            End If
        Loop
        ' TODO: zzi should be exactly 64 here. If it's not, we should report some kind of warning.
        If zzi > 64 Then zzi = 64
        pipe.dct_coeffs(0) = dec.state.frags(fragi).dc
        ' last_zzi is always initialized. If your compiler thinks otherwise, it is dumb.
        oc_state_frag_recon_c dec.state, fragi, pli, pipe.dct_coeffs, last_zzi, dc_quant(qti)
    Next
    pipe.coded_fragis(pli) = pipe.coded_fragis(pli) + ncoded_fragis
    ' Right now the reconstructed MCU has only the coded blocks in it.
    ' TODO: We make the decision here to always copy the uncoded blocks into it from the reference frame. We
    ' could also copy the coded blocks back over the reference frame, if we wait for an additional MCU to be
    ' decoded, which might be faster if only a small number of blocks are coded. However, this introduces more
    ' latency, creating a larger cache footprint. It's unknown which decision is better, but this one results
    ' in simpler code, and the hard case (high bitrate, high resolution) is handled correctly.
    ' Copy the uncoded blocks from the previous reference frame.
    If pipe.nuncoded_fragis(pli) > 0 Then
        pipe.uncoded_fragis(pli) = pipe.uncoded_fragis(pli) - pipe.nuncoded_fragis(pli)
        oc_frag_copy_list_c dec.state.ref_frame_handle, dec.state.ref_frame_data(OC_FRAME_SELF), _
            dec.state.ref_frame_data(OC_FRAME_PREV), dec.state.ref_ystride(pli), dec.state.coded_fragis, _
            pipe.uncoded_fragis(pli), pipe.nuncoded_fragis(pli), dec.state.frag_buf_offs
    End If
End Sub

' Filter a horizontal block edge.
' (VB) _dst, _src: offsets into pix; _variance0, _variance1: variances(variance0), variances(variance1)
Private Sub oc_filter_hedge(pix() As Byte, ByVal dst As Long, ByVal dst_ystride As Long, ByVal src As Long, _
    ByVal src_ystride As Long, ByVal qstep As Long, ByVal flimit As Long, variances() As Long, _
    ByVal variance0 As Long, ByVal variance1 As Long)
    Dim rdst As Long
    Dim rsrc As Long
    Dim cdst As Long
    Dim csrc As Long
    Static r(0 To 9) As Long        ' (VB) Static: no array set up per call
    Dim sum0 As Long
    Dim sum1 As Long
    Dim bx As Long
    Dim by As Long
    Dim d As Long
    rdst = dst
    rsrc = src
    For bx = 0 To 7
        cdst = rdst
        csrc = rsrc
        For by = 0 To 9
            r(by) = pix(csrc)
            csrc = csrc + src_ystride
        Next
        sum0 = 0
        sum1 = 0
        For by = 0 To 3
            d = r(by + 1) - r(by)
            If d < 0 Then d = -d
            sum0 = sum0 + d
            d = r(by + 5) - r(by + 6)
            If d < 0 Then d = -d
            sum1 = sum1 + d
        Next
        If sum0 < 255 Then variances(variance0) = variances(variance0) + sum0 Else variances(variance0) = variances(variance0) + 255
        If sum1 < 255 Then variances(variance1) = variances(variance1) + sum1 Else variances(variance1) = variances(variance1) + 255
        If sum0 < flimit And sum1 < flimit And r(5) - r(4) < qstep And r(4) - r(5) < qstep Then
            pix(cdst) = (r(0) * 3 + r(1) * 2 + r(2) + r(3) + r(4) + 4) \ 8
            cdst = cdst + dst_ystride
            pix(cdst) = (r(0) * 2 + r(1) + r(2) * 2 + r(3) + r(4) + r(5) + 4) \ 8
            cdst = cdst + dst_ystride
            For by = 0 To 3
                pix(cdst) = (r(by) + r(by + 1) + r(by + 2) + r(by + 3) * 2 + r(by + 4) + r(by + 5) + r(by + 6) + 4) \ 8
                cdst = cdst + dst_ystride
            Next
            pix(cdst) = (r(4) + r(5) + r(6) + r(7) * 2 + r(8) + r(9) * 2 + 4) \ 8
            cdst = cdst + dst_ystride
            pix(cdst) = (r(5) + r(6) + r(7) + r(8) * 2 + r(9) * 3 + 4) \ 8
        Else
            For by = 1 To 8
                pix(cdst) = r(by)
                cdst = cdst + dst_ystride
            Next
        End If
        rdst = rdst + 1
        rsrc = rsrc + 1
    Next
End Sub

' Filter a vertical block edge.
' (VB) _dst: an offset into pix; _variances: variances(variance)
Private Sub oc_filter_vedge(pix() As Byte, ByVal dst As Long, ByVal dst_ystride As Long, ByVal qstep As Long, _
    ByVal flimit As Long, variances() As Long, ByVal variance As Long)
    Dim rdst As Long
    Dim rsrc As Long
    Dim cdst As Long
    Static r(0 To 9) As Long        ' (VB) Static: no array set up per call
    Dim sum0 As Long
    Dim sum1 As Long
    Dim bx As Long
    Dim by As Long
    Dim d As Long
    cdst = dst
    For by = 0 To 7
        rsrc = cdst - 1
        rdst = cdst
        For bx = 0 To 9
            r(bx) = pix(rsrc)
            rsrc = rsrc + 1
        Next
        sum0 = 0
        sum1 = 0
        For bx = 0 To 3
            d = r(bx + 1) - r(bx)
            If d < 0 Then d = -d
            sum0 = sum0 + d
            d = r(bx + 5) - r(bx + 6)
            If d < 0 Then d = -d
            sum1 = sum1 + d
        Next
        If sum0 < 255 Then variances(variance) = variances(variance) + sum0 Else variances(variance) = variances(variance) + 255
        If sum1 < 255 Then
            variances(variance + 1) = variances(variance + 1) + sum1
        Else
            variances(variance + 1) = variances(variance + 1) + 255
        End If
        If sum0 < flimit And sum1 < flimit And r(5) - r(4) < qstep And r(4) - r(5) < qstep Then
            pix(rdst) = (r(0) * 3 + r(1) * 2 + r(2) + r(3) + r(4) + 4) \ 8
            pix(rdst + 1) = (r(0) * 2 + r(1) + r(2) * 2 + r(3) + r(4) + r(5) + 4) \ 8
            rdst = rdst + 2
            For bx = 0 To 3
                pix(rdst) = (r(bx) + r(bx + 1) + r(bx + 2) + r(bx + 3) * 2 + r(bx + 4) + r(bx + 5) + r(bx + 6) + 4) \ 8
                rdst = rdst + 1
            Next
            pix(rdst) = (r(4) + r(5) + r(6) + r(7) * 2 + r(8) + r(9) * 2 + 4) \ 8
            pix(rdst + 1) = (r(5) + r(6) + r(7) + r(8) * 2 + r(9) * 3 + 4) \ 8
        End If
        cdst = cdst + dst_ystride
    Next
End Sub

' (VB) _dst, _src: the th_ycbcr_buffers (the C passes the planes and adds _pli)
Private Sub oc_dec_deblock_frag_rows(dec As th_dec_ctx, dst_ As th_ycbcr_buffer, src_ As th_ycbcr_buffer, _
    ByVal pli As Long, ByVal fragy0 As Long, ByVal fragy_end As Long)
    Dim variance As Long            ' int *variance: (VB) an index into dec.variances
    Dim dc_qi As Long               ' unsigned char *dc_qi: (VB) an index into dec.dc_qis
    Dim dst As Long
    Dim src As Long
    Dim froffset As Long
    Dim dst_ystride As Long
    Dim src_ystride As Long
    Dim nhfrags As Long
    Dim width_ As Long
    Dim notstart As Long
    Dim notdone As Long
    Dim flimit As Long
    Dim qstep As Long
    Dim y_end As Long
    Dim y As Long
    Dim x As Long
    Dim k As Long
    Dim k_end As Long
    Dim height_ As Long
    nhfrags = dec.state.fplanes(pli).nhfrags
    froffset = dec.state.fplanes(pli).froffset + fragy0 * nhfrags
    variance = froffset
    dc_qi = froffset
    If fragy0 > 0 Then notstart = 1 Else notstart = 0
    If fragy_end < dec.state.fplanes(pli).nvfrags Then notdone = 1 Else notdone = 0
    ' We want to clear an extra row of variances, except at the end.
    k_end = variance + nhfrags * notstart + (fragy_end + notdone - fragy0 - notstart) * nhfrags
    For k = variance + nhfrags * notstart To k_end - 1
        dec.variances(k) = 0
    Next
    ' Except for the first time, we want to point to the middle of the row.
    y = fragy0 * 8 + notstart * 4
    dst_ystride = dst_.p(pli).stride
    src_ystride = src_.p(pli).stride
    dst = dst_.p(pli).data + y * dst_ystride
    src = src_.p(pli).data + y * src_ystride
    width_ = dst_.p(pli).width
    Do While y < 4
        CopyMemory dec.state.ref_frame_handle(dst), dec.state.ref_frame_handle(src), width_
        dst = dst + dst_ystride
        src = src + src_ystride
        y = y + 1
    Loop
    ' We also want to skip the last row in the frame for this loop.
    y_end = (fragy_end - (1 - notdone)) * 8
    Do While y < y_end
        qstep = dec.pp_dc_scale(dec.dc_qis(dc_qi))
        flimit = (qstep * 3) \ 4
        oc_filter_hedge dec.state.ref_frame_handle, dst, dst_ystride, src - src_ystride, src_ystride, _
            qstep, flimit, dec.variances, variance, variance + nhfrags
        variance = variance + 1
        dc_qi = dc_qi + 1
        For x = 8 To width_ - 1 Step 8
            qstep = dec.pp_dc_scale(dec.dc_qis(dc_qi))
            flimit = (qstep * 3) \ 4
            oc_filter_hedge dec.state.ref_frame_handle, dst + x, dst_ystride, src + x - src_ystride, src_ystride, _
                qstep, flimit, dec.variances, variance, variance + nhfrags
            oc_filter_vedge dec.state.ref_frame_handle, dst + x - (dst_ystride * 4) - 4, dst_ystride, _
                qstep, flimit, dec.variances, variance - 1
            variance = variance + 1
            dc_qi = dc_qi + 1
        Next
        dst = dst + dst_ystride * 8
        src = src + src_ystride * 8
        y = y + 8
    Loop
    ' And finally, handle the last row in the frame, if it's in the range.
    If notdone = 0 Then
        height_ = dst_.p(pli).height
        Do While y < height_
            CopyMemory dec.state.ref_frame_handle(dst), dec.state.ref_frame_handle(src), width_
            dst = dst + dst_ystride
            src = src + src_ystride
            y = y + 1
        Loop
        ' Filter the last row of vertical block edges.
        dc_qi = dc_qi + 1
        For x = 8 To width_ - 1 Step 8
            qstep = dec.pp_dc_scale(dec.dc_qis(dc_qi))
            dc_qi = dc_qi + 1
            flimit = (qstep * 3) \ 4
            oc_filter_vedge dec.state.ref_frame_handle, dst + x - (dst_ystride * 8) - 4, dst_ystride, _
                qstep, flimit, dec.variances, variance
            variance = variance + 1
        Next
    End If
End Sub

' (VB) _idata: an offset into pix
Private Sub oc_dering_block(pix() As Byte, ByVal idata As Long, ByVal ystride As Long, ByVal b As Long, _
    ByVal dc_scale As Long, ByVal sharp_mod As Long, ByVal strong As Long)
    ' (OC_MOD_MAX[2]={24,32}; OC_MOD_SHIFT[2]={1,0}: theora_tables_vb)
    Static vmod(0 To 71) As Long
    Static hmod(0 To 71) As Long
    Dim psrc As Long
    Dim src As Long
    Dim nsrc As Long
    Dim dst As Long
    Dim mod_hi As Long
    Dim by As Long
    Dim bx As Long
    Dim mod_ As Long
    Dim modMul As Long          ' (VB) 1 << OC_MOD_SHIFT[_strong]
    Dim d As Long
    Dim a As Long
    Dim bb As Long              ' (VB) the C's local b (_b is b)
    Dim w As Long
    Dim v As Long
    mod_hi = 3 * dc_scale
    If OC_MOD_MAX(strong) < mod_hi Then mod_hi = OC_MOD_MAX(strong)
    If OC_MOD_SHIFT(strong) <> 0 Then modMul = 2 Else modMul = 1
    dst = idata
    src = dst
    psrc = src
    If (b And 4) = 0 Then psrc = src - ystride
    For by = 0 To 8
        For bx = 0 To 7
            d = CLng(pix(src + bx)) - pix(psrc + bx)
            If d < 0 Then d = -d
            mod_ = 32 + dc_scale - d * modMul
            If mod_ < -64 Then
                vmod(by * 8 + bx) = sharp_mod
            ElseIf mod_ < 0 Then
                vmod(by * 8 + bx) = 0
            ElseIf mod_ > mod_hi Then
                vmod(by * 8 + bx) = mod_hi
            Else
                vmod(by * 8 + bx) = mod_
            End If
        Next
        psrc = src
        If (b And 8) = 0 Or by < 7 Then src = src + ystride
    Next
    nsrc = dst
    psrc = dst
    If (b And 1) = 0 Then psrc = dst - 1
    For bx = 0 To 8
        src = nsrc
        For by = 0 To 7
            d = CLng(pix(src)) - pix(psrc)
            If d < 0 Then d = -d
            mod_ = 32 + dc_scale - d * modMul
            If mod_ < -64 Then
                hmod(bx * 8 + by) = sharp_mod
            ElseIf mod_ < 0 Then
                hmod(bx * 8 + by) = 0
            ElseIf mod_ > mod_hi Then
                hmod(bx * 8 + by) = mod_hi
            Else
                hmod(bx * 8 + by) = mod_
            End If
            psrc = psrc + ystride
            src = src + ystride
        Next
        psrc = nsrc
        If (b And 2) = 0 Or bx < 7 Then nsrc = nsrc + 1
    Next
    src = dst
    psrc = src
    If (b And 4) = 0 Then psrc = src - ystride
    nsrc = src + ystride
    For by = 0 To 7
        a = 128
        bb = 64
        w = hmod(by)
        a = a - w
        If (b And 1) = 0 Then bb = bb + w * pix(src - 1) Else bb = bb + w * pix(src)
        w = vmod(by * 8)
        a = a - w
        bb = bb + w * pix(psrc)
        w = vmod((by + 1) * 8)
        a = a - w
        bb = bb + w * pix(nsrc)
        w = hmod(8 + by)
        a = a - w
        bb = bb + w * pix(src + 1)
        v = ((a * pix(src) + bb) And &HFFFFFF80) \ 128
        If v < 0 Then v = 0 Else If v > 255 Then v = 255
        pix(dst) = v
        For bx = 1 To 6
            a = 128
            bb = 64
            w = hmod(bx * 8 + by)
            a = a - w
            bb = bb + w * pix(src + bx - 1)
            w = vmod(by * 8 + bx)
            a = a - w
            bb = bb + w * pix(psrc + bx)
            w = vmod((by + 1) * 8 + bx)
            a = a - w
            bb = bb + w * pix(nsrc + bx)
            w = hmod((bx + 1) * 8 + by)
            a = a - w
            bb = bb + w * pix(src + bx + 1)
            v = ((a * pix(src + bx) + bb) And &HFFFFFF80) \ 128
            If v < 0 Then v = 0 Else If v > 255 Then v = 255
            pix(dst + bx) = v
        Next
        a = 128
        bb = 64
        w = hmod(7 * 8 + by)
        a = a - w
        bb = bb + w * pix(src + 6)
        w = vmod(by * 8 + 7)
        a = a - w
        bb = bb + w * pix(psrc + 7)
        w = vmod((by + 1) * 8 + 7)
        a = a - w
        bb = bb + w * pix(nsrc + 7)
        w = hmod(8 * 8 + by)
        a = a - w
        If (b And 2) = 0 Then bb = bb + w * pix(src + 8) Else bb = bb + w * pix(src + 7)
        v = ((a * pix(src + 7) + bb) And &HFFFFFF80) \ 128
        If v < 0 Then v = 0 Else If v > 255 Then v = 255
        pix(dst + 7) = v
        dst = dst + ystride
        psrc = src
        src = nsrc
        If (b And 8) = 0 Or by < 6 Then nsrc = nsrc + ystride
    Next
End Sub

' (VB) _img: the th_ycbcr_buffer (the C passes the planes and adds _pli)
Private Sub oc_dec_dering_frag_rows(dec As th_dec_ctx, img As th_ycbcr_buffer, ByVal pli As Long, _
    ByVal fragy0 As Long, ByVal fragy_end As Long)
    Dim frag As Long                ' oc_fragment *frag: (VB) an index into state.frags
    Dim variance As Long            ' int *variance: (VB) an index into dec.variances
    Dim idata As Long
    Dim froffset As Long
    Dim ystride As Long
    Dim nhfrags As Long
    Dim sthresh As Long
    Dim strong As Long
    Dim y_end As Long
    Dim width_ As Long
    Dim height_ As Long
    Dim y As Long
    Dim x As Long
    Dim b As Long
    Dim qi As Long
    Dim var_ As Long
    Dim more As Boolean
    nhfrags = dec.state.fplanes(pli).nhfrags
    froffset = dec.state.fplanes(pli).froffset + fragy0 * nhfrags
    variance = froffset
    frag = froffset
    If pli <> 0 Then
        If dec.pp_level >= OC_PP_LEVEL_SDERINGC Then strong = 1 Else strong = 0
        sthresh = OC_DERING_THRESH4
    Else
        If dec.pp_level >= OC_PP_LEVEL_SDERINGY Then strong = 1 Else strong = 0
        sthresh = OC_DERING_THRESH3
    End If
    y = fragy0 * 8
    ystride = img.p(pli).stride
    idata = img.p(pli).data + y * ystride
    y_end = fragy_end * 8
    width_ = img.p(pli).width
    height_ = img.p(pli).height
    Do While y < y_end
        For x = 0 To width_ - 1 Step 8
            qi = dec.state.qis(dec.state.frags(frag).qii)
            var_ = dec.variances(variance)
            b = 0
            If x <= 0 Then b = 1
            If x + 8 >= width_ Then b = b Or 2
            If y <= 0 Then b = b Or 4
            If y + 8 >= height_ Then b = b Or 8
            If strong <> 0 And var_ > sthresh Then
                oc_dering_block dec.state.ref_frame_handle, idata + x, ystride, b, dec.pp_dc_scale(qi), _
                    dec.pp_sharp_mod(qi), 1
                ' _pli||!(b&1)&&*(variance-1)>OC_DERING_THRESH4||!(b&2)&&variance[1]>OC_DERING_THRESH4||
                ' !(b&4)&&*(variance-nhfrags)>OC_DERING_THRESH4||!(b&8)&&variance[nhfrags]>OC_DERING_THRESH4
                more = pli <> 0
                If Not more And (b And 1) = 0 Then more = dec.variances(variance - 1) > OC_DERING_THRESH4
                If Not more And (b And 2) = 0 Then more = dec.variances(variance + 1) > OC_DERING_THRESH4
                If Not more And (b And 4) = 0 Then more = dec.variances(variance - nhfrags) > OC_DERING_THRESH4
                If Not more And (b And 8) = 0 Then more = dec.variances(variance + nhfrags) > OC_DERING_THRESH4
                If more Then
                    oc_dering_block dec.state.ref_frame_handle, idata + x, ystride, b, dec.pp_dc_scale(qi), _
                        dec.pp_sharp_mod(qi), 1
                    oc_dering_block dec.state.ref_frame_handle, idata + x, ystride, b, dec.pp_dc_scale(qi), _
                        dec.pp_sharp_mod(qi), 1
                End If
            ElseIf var_ > OC_DERING_THRESH2 Then
                oc_dering_block dec.state.ref_frame_handle, idata + x, ystride, b, dec.pp_dc_scale(qi), _
                    dec.pp_sharp_mod(qi), 1
            ElseIf var_ > OC_DERING_THRESH1 Then
                oc_dering_block dec.state.ref_frame_handle, idata + x, ystride, b, dec.pp_dc_scale(qi), _
                    dec.pp_sharp_mod(qi), 0
            End If
            frag = frag + 1
            variance = variance + 1
        Next
        idata = idata + ystride * 8
        y = y + 8
    Loop
End Sub

Public Function th_decode_alloc(info As th_info, ByVal setup As Long) As Long
    Dim dec As Long
    ' if(_info==NULL||_setup==NULL)return NULL: (VB) _info is never NULL
    If setup = 0 Then th_decode_alloc = 0: Exit Function
    oc_vb_init
    dec = oc_dec_vbAlloc()
    If oc_dec_init(oc_decs(dec), info, oc_setups(setup)) < 0 Then
        oc_dec_vbFree dec
        th_decode_alloc = 0
        Exit Function
    End If
    oc_decs(dec).state.curframe_num = 0
    th_decode_alloc = dec
End Function

Public Sub th_decode_free(ByVal dec As Long)
    If dec <> 0 Then
        oc_dec_clear oc_decs(dec)
        oc_dec_vbFree dec
    End If
End Sub

' (VB) _buf: the int of TH_DECCTL_GET_PPLEVEL_MAX and TH_DECCTL_SET_PPLEVEL (_buf_sz 4); for
' TH_DECCTL_SET_GRANPOS (_buf_sz 8, sizeof(ogg_int64_t)) the granule position, which must then fit in a Long
Public Function th_decode_ctl(ByVal dec As Long, ByVal req As Long, buf As Long, ByVal buf_sz As Long) As Long
    Dim pp_level As Long
    Dim granpos As Double
    Dim shift As Double
    Select Case req
        Case TH_DECCTL_GET_PPLEVEL_MAX
            If dec = 0 Then th_decode_ctl = TH_EFAULT: Exit Function
            If buf_sz <> 4 Then th_decode_ctl = TH_EINVAL: Exit Function
            buf = OC_PP_LEVEL_MAX
            th_decode_ctl = 0
        Case TH_DECCTL_SET_PPLEVEL
            If dec = 0 Then th_decode_ctl = TH_EFAULT: Exit Function
            If buf_sz <> 4 Then th_decode_ctl = TH_EINVAL: Exit Function
            pp_level = buf
            If pp_level < 0 Or pp_level > OC_PP_LEVEL_MAX Then th_decode_ctl = TH_EINVAL: Exit Function
            oc_decs(dec).pp_level = pp_level
            th_decode_ctl = 0
        Case TH_DECCTL_SET_GRANPOS
            If dec = 0 Then th_decode_ctl = TH_EFAULT: Exit Function
            If buf_sz <> 8 Then th_decode_ctl = TH_EINVAL: Exit Function
            granpos = buf
            If granpos < 0 Then th_decode_ctl = TH_EINVAL: Exit Function
            With oc_decs(dec).state
                .granpos = granpos
                shift = 2# ^ .info.keyframe_granule_shift
                .keyframe_num = Int(granpos / shift) - .granpos_bias
                .curframe_num = .keyframe_num + (granpos - Int(granpos / shift) * shift)
            End With
            th_decode_ctl = 0
        Case Else
            ' (TH_DECCTL_SET_STRIPE_CB and the HAVE_CAIRO telemetry requests: not ported)
            th_decode_ctl = TH_EIMPL
    End Select
End Function

' We're decoding an INTER frame, but have no initialized reference buffers (i.e., decoding did not start on a
' key frame). We initialize them to a solid gray here.
Private Sub oc_dec_init_dummy_frame(dec As th_dec_ctx)
    Dim yplane_sz As Long
    Dim cplane_sz As Long
    Dim yoffset As Long
    Dim yhstride As Long
    Dim yheight As Long
    Dim chstride As Long
    Dim cheight As Long
    dec.state.ref_frame_idx(OC_FRAME_GOLD) = 0
    dec.state.ref_frame_idx(OC_FRAME_PREV) = 0
    dec.state.ref_frame_idx(OC_FRAME_SELF) = 0
    dec.state.ref_frame_data(OC_FRAME_GOLD) = dec.state.ref_frame_bufs(0).p(0).data
    dec.state.ref_frame_data(OC_FRAME_PREV) = dec.state.ref_frame_bufs(0).p(0).data
    dec.state.ref_frame_data(OC_FRAME_SELF) = dec.state.ref_frame_bufs(0).p(0).data
    dec.pp_frame_buf = dec.state.ref_frame_bufs(0)
    yhstride = Abs(dec.state.ref_ystride(0))
    yheight = dec.state.info.frame_height + 2 * OC_UMV_PADDING
    chstride = Abs(dec.state.ref_ystride(1))
    If (dec.state.info.pixel_fmt And 2) = 0 Then cheight = yheight \ 2 Else cheight = yheight
    yplane_sz = yhstride * yheight + 16
    cplane_sz = chstride * cheight
    yoffset = yhstride * (yheight - OC_UMV_PADDING - 1) + OC_UMV_PADDING
    FillMemory dec.state.ref_frame_handle(dec.state.ref_frame_data(0) - yoffset), yplane_sz + 2 * cplane_sz, &H80
End Sub

' (VB) _granpos: never NULL
Public Function th_decode_packetin(ByVal dec As Long, op As ogg_packet, granpos As Double) As Long
    If dec = 0 Then th_decode_packetin = TH_EFAULT: Exit Function
    th_decode_packetin = oc_dec_packetin(oc_decs(dec), op, granpos)
End Function

Private Function oc_dec_packetin(dec As th_dec_ctx, op As ogg_packet, granpos As Double) As Long
    Dim ret As Long
    Dim stripe_buf As th_ycbcr_buffer
    Dim stripe_fragy As Long
    Dim refi As Long
    Dim pli As Long
    Dim notstart As Long
    Dim notdone As Long
    Dim avail_fragy0 As Long
    Dim avail_fragy_end As Long
    Dim frag_shift As Long
    Dim pp_offset As Long
    Dim sdelay As Long
    Dim edelay As Long
    Dim v As Long
    ' A completely empty packet indicates a dropped frame and is treated exactly like an inter frame with no
    ' coded blocks.
    If op.bytes = 0 Then
        dec.state.frame_type = OC_INTER_FRAME
        dec.state.ntotal_coded_fragis = 0
    Else
        oc_pack_readinit dec.opb, op.packet, 0, op.bytes
        ret = oc_dec_frame_header_unpack(dec)
        If ret < 0 Then oc_dec_packetin = ret: Exit Function
        If dec.state.frame_type = OC_INTRA_FRAME Then
            oc_dec_mark_all_intra dec
        Else
            oc_dec_coded_flags_unpack dec
        End If
    End If
    ' If there have been no reference frames, and we need one, initialize one.
    If dec.state.frame_type <> OC_INTRA_FRAME Then
        If dec.state.ref_frame_idx(OC_FRAME_GOLD) < 0 Or dec.state.ref_frame_idx(OC_FRAME_PREV) < 0 Then
            oc_dec_init_dummy_frame dec
        End If
    End If
    ' If this was an inter frame with no coded blocks...
    If dec.state.ntotal_coded_fragis <= 0 Then
        ' Just update the granule position and return.
        dec.state.granpos = (dec.state.keyframe_num + dec.state.granpos_bias) * 2# ^ dec.state.info.keyframe_granule_shift _
            + (dec.state.curframe_num - dec.state.keyframe_num)
        dec.state.curframe_num = dec.state.curframe_num + 1
        granpos = dec.state.granpos
        oc_dec_packetin = TH_DUPFRAME
        Exit Function
    End If
    ' Select a free buffer to use for the reconstructed version of this frame.
    refi = 0
    Do While refi = dec.state.ref_frame_idx(OC_FRAME_GOLD) Or refi = dec.state.ref_frame_idx(OC_FRAME_PREV)
        refi = refi + 1
    Loop
    dec.state.ref_frame_idx(OC_FRAME_SELF) = refi
    dec.state.ref_frame_data(OC_FRAME_SELF) = dec.state.ref_frame_bufs(refi).p(0).data
    If dec.state.frame_type = OC_INTRA_FRAME Then
        dec.state.keyframe_num = dec.state.curframe_num
    Else
        oc_dec_mb_modes_unpack dec
        oc_dec_mv_unpack_and_frag_modes_fill dec
    End If
    oc_dec_block_qis_unpack dec
    oc_dec_residual_tokens_unpack dec
    ' Update granule position. This must be done before the striped decode callbacks so that the application
    ' knows what to do with the frame data.
    dec.state.granpos = (dec.state.keyframe_num + dec.state.granpos_bias) * 2# ^ dec.state.info.keyframe_granule_shift _
        + (dec.state.curframe_num - dec.state.keyframe_num)
    dec.state.curframe_num = dec.state.curframe_num + 1
    granpos = dec.state.granpos
    ' All of the rest of the operations -- DC prediction reversal, reconstructing coded fragments, copying
    ' uncoded fragments, loop filtering, extending borders, and out-of-loop post-processing -- should be
    ' pipelined. I.e., DC prediction reversal, reconstruction, and uncoded fragment copying are done for one or
    ' two super block rows, then loop filtering is run as far as it can, then bordering copying, then
    ' post-processing. For 4:2:0 video a Minimum Codable Unit or MCU contains two luma super block rows, and one
    ' chroma. Otherwise, an MCU consists of one super block row from each plane. Inside each MCU, we perform all
    ' of the steps on one color plane before moving on to the next. After reconstruction, the additional
    ' filtering stages introduce a delay since they need some pixels from the next fragment row. Thus the actual
    ' number of decoded rows available is slightly smaller for the first MCU, and slightly larger for the last.
    ' This entire process allows us to operate on the data while it is still in cache, resulting in big
    ' performance improvements. An application callback allows further application processing (blitting to
    ' video memory, color conversion, etc.) to also use the data while it's in cache.
    oc_dec_pipeline_init dec, dec.pipe
    oc_ycbcr_buffer_flip stripe_buf, dec.pp_frame_buf
    notstart = 0
    notdone = 1
    stripe_fragy = 0
    Do While notdone <> 0
        avail_fragy0 = dec.state.fplanes(0).nvfrags
        avail_fragy_end = avail_fragy0
        If stripe_fragy + dec.pipe.mcu_nvfrags < avail_fragy_end Then notdone = 1 Else notdone = 0
        For pli = 0 To 2
            ' Compute the first and last fragment row of the current MCU for this plane.
            frag_shift = 0
            If pli <> 0 Then
                If (dec.state.info.pixel_fmt And 2) = 0 Then frag_shift = 1
            End If
            dec.pipe.fragy0(pli) = stripe_fragy \ OC_PB_POW2(frag_shift)
            v = dec.pipe.fragy0(pli) + dec.pipe.mcu_nvfrags \ OC_PB_POW2(frag_shift)
            If dec.state.fplanes(pli).nvfrags < v Then v = dec.state.fplanes(pli).nvfrags
            dec.pipe.fragy_end(pli) = v
            oc_dec_dc_unpredict_mcu_plane_c dec, dec.pipe, pli
            oc_dec_frags_recon_mcu_plane dec, dec.pipe, pli
            sdelay = 0
            edelay = 0
            If dec.pipe.loop_filter <> 0 Then
                sdelay = sdelay + notstart
                edelay = edelay + notdone
                oc_state_loop_filter_frag_rows_c dec.state, dec.state.ref_frame_handle, dec.pipe.bounding_values, _
                    OC_FRAME_SELF, pli, dec.pipe.fragy0(pli) - sdelay, dec.pipe.fragy_end(pli) - edelay
            End If
            ' To fill the borders, we have an additional two pixel delay, since a fragment in the next row could
            ' filter its top edge, using two pixels from a fragment in this row. But there's no reason to delay a
            ' full fragment between the two.
            oc_state_borders_fill_rows dec.state, dec.state.ref_frame_handle, refi, pli, _
                (dec.pipe.fragy0(pli) - sdelay) * 8 - sdelay * 2, (dec.pipe.fragy_end(pli) - edelay) * 8 - edelay * 2
            ' Out-of-loop post-processing.
            If pli <> 0 Then pp_offset = 3 Else pp_offset = 0
            If dec.pipe.pp_level >= OC_PP_LEVEL_DEBLOCKY + pp_offset Then
                ' Perform de-blocking in one plane.
                sdelay = sdelay + notstart
                edelay = edelay + notdone
                oc_dec_deblock_frag_rows dec, dec.pp_frame_buf, dec.state.ref_frame_bufs(refi), pli, _
                    dec.pipe.fragy0(pli) - sdelay, dec.pipe.fragy_end(pli) - edelay
                If dec.pipe.pp_level >= OC_PP_LEVEL_DERINGY + pp_offset Then
                    ' Perform de-ringing in one plane.
                    sdelay = sdelay + notstart
                    edelay = edelay + notdone
                    oc_dec_dering_frag_rows dec, dec.pp_frame_buf, pli, dec.pipe.fragy0(pli) - sdelay, _
                        dec.pipe.fragy_end(pli) - edelay
                End If
            ' If no post-processing is done, we still need to delay a row for the loop filter, thanks to the
            ' strange filtering order VP3 chose.
            ElseIf dec.pipe.loop_filter <> 0 Then
                sdelay = sdelay + notstart
                edelay = edelay + notdone
            End If
            ' Compute the intersection of the available rows in all planes. If chroma is sub-sampled, the effect of
            ' each of its delays is doubled, but luma might have more post-processing filters enabled than
            ' chroma, so we don't know up front which one is the limiting factor.
            v = (dec.pipe.fragy0(pli) - sdelay) * OC_PB_POW2(frag_shift)
            If v < avail_fragy0 Then avail_fragy0 = v
            v = (dec.pipe.fragy_end(pli) - edelay) * OC_PB_POW2(frag_shift)
            If v < avail_fragy_end Then avail_fragy_end = v
        Next
        ' if(_dec->stripe_cb.stripe_decoded!=NULL){...}: (VB) no striped decode callback (not ported)
        notstart = 1
        stripe_fragy = stripe_fragy + dec.pipe.mcu_nvfrags
    Loop
    ' Finish filling in the reference frame borders.
    For pli = 0 To 2
        oc_state_borders_fill_caps dec.state, refi, pli
    Next
    ' Update the reference frame indices.
    If dec.state.frame_type = OC_INTRA_FRAME Then
        ' The new frame becomes both the previous and gold reference frames.
        dec.state.ref_frame_idx(OC_FRAME_GOLD) = dec.state.ref_frame_idx(OC_FRAME_SELF)
        dec.state.ref_frame_idx(OC_FRAME_PREV) = dec.state.ref_frame_idx(OC_FRAME_SELF)
        dec.state.ref_frame_data(OC_FRAME_GOLD) = dec.state.ref_frame_data(OC_FRAME_SELF)
        dec.state.ref_frame_data(OC_FRAME_PREV) = dec.state.ref_frame_data(OC_FRAME_SELF)
    Else
        ' Otherwise, just replace the previous reference frame.
        dec.state.ref_frame_idx(OC_FRAME_PREV) = dec.state.ref_frame_idx(OC_FRAME_SELF)
        dec.state.ref_frame_data(OC_FRAME_PREV) = dec.state.ref_frame_data(OC_FRAME_SELF)
    End If
    ' Restore the FPU before dump_frame, since that _does_ use the FPU (for PNG gamma values, if nothing else).
    oc_restore_fpu_c
    oc_dec_packetin = 0
End Function

Public Function th_decode_ycbcr_out(ByVal dec As Long, ycbcr As th_ycbcr_buffer) As Long
    If dec = 0 Then th_decode_ycbcr_out = TH_EFAULT: Exit Function
    oc_ycbcr_buffer_flip ycbcr, oc_decs(dec).pp_frame_buf
    th_decode_ycbcr_out = 0
End Function

' (VB) The pixels a th_img_plane of a decoder's th_decode_ycbcr_out points at, copied into out (plane.height
' rows of plane.width bytes, top row first; out is ReDimmed only when it is too small)
Public Sub th_decode_vbCopyPlane(ByVal dec As Long, plane As th_img_plane, out() As Byte, ByVal outSize As Long)
    Dim row As Long
    Dim src As Long
    Dim need As Long
    need = plane.width * plane.height
    If outSize < need Then ReDim out(0 To need - 1)
    src = plane.data
    For row = 0 To plane.height - 1
        CopyMemory out(row * plane.width), oc_decs(dec).state.ref_frame_handle(src), plane.width
        src = src + plane.stride
    Next
End Sub
