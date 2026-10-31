Attribute VB_Name = "silk_decode_core_c"
' libopus 1.5.2 (ref/opus/silk): decode_core.c.
' (VB) xq is an array and an offset; the C's pointers into the state's arrays are offsets. The stack buffers
' are this module's arrays at their largest sizes. pres_Q14 points at res_Q14 for a voiced subframe and at
' pexc_Q14 for the others, where the excitation is copied into res_Q14 instead (the loop only reads it).
' The LTP and LPC filters' silk_SMLAWB are written out in place (silk_SigProc_FIX_h's formula); the 10 or 16
' unrolled LPC taps are a loop (the same sum: 32-bit additions wrap in any order).
Option Explicit

Private sLTP(0 To 319) As Integer                       ' [ ltp_mem_length ]
Private sLTP_Q15(0 To 639) As Long                      ' [ ltp_mem_length + frame_length ]
Private res_Q14(0 To 79) As Long                        ' [ subfr_length ]
Private sLPC_Q14(0 To 95) As Long                       ' [ subfr_length + MAX_LPC_ORDER ]
Private A_Q12_tmp(0 To 15) As Long                      ' opus_int16 [ MAX_LPC_ORDER ]

'/**********************************************************/
'/* Core decoder. Performs inverse NSQ operation LTP + LPC */
'/**********************************************************/
Public Sub silk_decode_core(psDec As silk_decoder_state, psDecCtrl As silk_decoder_control, xq() As Integer, _
        ByVal xqOffset As Long, pulses() As Integer, ByVal arch As Long)
    Dim i As Long, k As Long, lag As Long, start_idx As Long, sLTP_buf_idx As Long, NLSF_interpolation_flag As Long, signalType As Long
    Dim A_Q12 As Long, B_Q14 As Long, pxq As Long       ' opus_int16 *: offsets into PredCoef_Q12, LTPCoef_Q14, xq
    Dim LTP_pred_Q13 As Long, LPC_pred_Q10 As Long, Gain_Q10 As Long, inv_gain_Q31 As Long, gain_adj_Q16 As Long, rand_seed As Long, offset_Q10 As Long
    Dim pred_lag_ptr As Long, pexc_Q14 As Long          ' opus_int32 *: offsets into sLTP_Q15, psDec->exc_Q14
    Dim j As Long, s As Long, c As Long, x As Long, order As Long
    Dim b0 As Long, b1 As Long, b2 As Long, b3 As Long, b4 As Long

    ' silk_assert( psDec->prev_gain_Q16 != 0 );

    offset_Q10 = silk_rom_i16(silk_Quantization_Offsets_Q10 + silk_RSHIFT(psDec.indices.signalType, 1) * 2 + psDec.indices.quantOffsetType)

    If psDec.indices.NLSFInterpCoef_Q2 < 4 Then
        NLSF_interpolation_flag = 1
    Else
        NLSF_interpolation_flag = 0
    End If

    '/* Decode excitation */
    rand_seed = psDec.indices.Seed
    For i = 0 To psDec.frame_length - 1
        rand_seed = RAND_INCREMENT + rand_seed * RAND_MULTIPLIER    ' silk_RAND( rand_seed )
        x = CLng(pulses(i)) * &H4000&                   ' silk_LSHIFT( (opus_int32)pulses[ i ], 14 )
        If x > 0 Then
            x = x - QUANT_LEVEL_ADJUST_Q10 * 16
        ElseIf x < 0 Then
            x = x + QUANT_LEVEL_ADJUST_Q10 * 16
        End If
        x = x + offset_Q10 * 16
        If rand_seed < 0 Then
           x = -x
        End If
        psDec.exc_Q14(i) = x

        rand_seed = rand_seed + pulses(i)               ' silk_ADD32_ovflw( rand_seed, pulses[ i ] )
    Next

    '/* Copy LPC state */
    silk_memcpy_i32 sLPC_Q14, 0, psDec.sLPC_Q14_buf, 0, MAX_LPC_ORDER

    pexc_Q14 = 0
    pxq = xqOffset
    sLTP_buf_idx = psDec.ltp_mem_length
    order = psDec.LPC_order
    '/* Loop over subframes */
    For k = 0 To psDec.nb_subfr - 1
        A_Q12 = silk_RSHIFT(k, 1) * MAX_LPC_ORDER

        '/* Preload LPC coeficients to array on stack. Gives small performance gain */
        For j = 0 To order - 1
            A_Q12_tmp(j) = psDecCtrl.PredCoef_Q12(A_Q12 + j)
        Next
        B_Q14 = k * LTP_ORDER
        signalType = psDec.indices.signalType

        Gain_Q10 = silk_RSHIFT(psDecCtrl.Gains_Q16(k), 6)
        inv_gain_Q31 = silk_INVERSE32_varQ(psDecCtrl.Gains_Q16(k), 47)

        '/* Calculate gain adjustment factor */
        If psDecCtrl.Gains_Q16(k) <> psDec.prev_gain_Q16 Then
            gain_adj_Q16 = silk_DIV32_varQ(psDec.prev_gain_Q16, psDecCtrl.Gains_Q16(k), 16)

            '/* Scale short term state */
            For i = 0 To MAX_LPC_ORDER - 1
                sLPC_Q14(i) = silk_SMULWW(gain_adj_Q16, sLPC_Q14(i))
            Next
        Else
            gain_adj_Q16 = &H10000
        End If

        '/* Save inv_gain */
        ' silk_assert( inv_gain_Q31 != 0 );
        psDec.prev_gain_Q16 = psDecCtrl.Gains_Q16(k)

        '/* Avoid abrupt transition from voiced PLC to unvoiced normal decoding */
        If psDec.lossCnt <> 0 And psDec.prevSignalType = TYPE_VOICED And _
            psDec.indices.signalType <> TYPE_VOICED And k < MAX_NB_SUBFR \ 2 Then

            silk_memset_i16 psDecCtrl.LTPCoef_Q14, B_Q14, LTP_ORDER
            psDecCtrl.LTPCoef_Q14(B_Q14 + LTP_ORDER \ 2) = 4096         ' SILK_FIX_CONST( 0.25, 14 )

            signalType = TYPE_VOICED
            psDecCtrl.pitchL(k) = psDec.lagPrev
        End If

        If signalType = TYPE_VOICED Then
            '/* Voiced */
            lag = psDecCtrl.pitchL(k)

            '/* Re-whitening */
            If k = 0 Or (k = 2 And NLSF_interpolation_flag <> 0) Then
                '/* Rewhiten with new A coefs */
                start_idx = psDec.ltp_mem_length - lag - psDec.LPC_order - LTP_ORDER \ 2
                ' celt_assert( start_idx > 0 );

                If k = 2 Then
                    silk_memcpy_i16 psDec.outBuf, psDec.ltp_mem_length, xq, xqOffset, 2 * psDec.subfr_length
                End If

                silk_LPC_analysis_filter sLTP, start_idx, psDec.outBuf, start_idx + k * psDec.subfr_length, _
                    psDecCtrl.PredCoef_Q12, A_Q12, psDec.ltp_mem_length - start_idx, psDec.LPC_order, arch

                '/* After rewhitening the LTP state is unscaled */
                If k = 0 Then
                    '/* Do LTP downscaling to reduce inter-packet dependency */
                    inv_gain_Q31 = silk_SMULWB(inv_gain_Q31, psDecCtrl.LTP_scale_Q14) * 4
                End If
                For i = 0 To lag + LTP_ORDER \ 2 - 1
                    sLTP_Q15(sLTP_buf_idx - i - 1) = silk_SMULWB(inv_gain_Q31, sLTP(psDec.ltp_mem_length - i - 1))
                Next
            Else
                '/* Update LTP state when Gain changes */
                If gain_adj_Q16 <> &H10000 Then
                    For i = 0 To lag + LTP_ORDER \ 2 - 1
                        sLTP_Q15(sLTP_buf_idx - i - 1) = silk_SMULWW(gain_adj_Q16, sLTP_Q15(sLTP_buf_idx - i - 1))
                    Next
                End If
            End If
        End If

        '/* Long-term prediction */
        If signalType = TYPE_VOICED Then
            '/* Set up pointer */
            pred_lag_ptr = sLTP_buf_idx - lag + LTP_ORDER \ 2
            b0 = psDecCtrl.LTPCoef_Q14(B_Q14)
            b1 = psDecCtrl.LTPCoef_Q14(B_Q14 + 1)
            b2 = psDecCtrl.LTPCoef_Q14(B_Q14 + 2)
            b3 = psDecCtrl.LTPCoef_Q14(B_Q14 + 3)
            b4 = psDecCtrl.LTPCoef_Q14(B_Q14 + 4)
            For i = 0 To psDec.subfr_length - 1
                '/* Unrolled loop */
                '/* Avoids introducing a bias because silk_SMLAWB() always rounds to -inf */
                LTP_pred_Q13 = 2
                s = sLTP_Q15(pred_lag_ptr)                ' silk_SMLAWB( LTP_pred_Q13, pred_lag_ptr[  0 ], B_Q14[ 0 ] )
                LTP_pred_Q13 = LTP_pred_Q13 + ((s And &HFFFF0000) \ &H10000) * b0 + ((((s And &HFFFF&) * b0) And &HFFFF0000) \ &H10000)
                s = sLTP_Q15(pred_lag_ptr - 1)            ' silk_SMLAWB( LTP_pred_Q13, pred_lag_ptr[ -1 ], B_Q14[ 1 ] )
                LTP_pred_Q13 = LTP_pred_Q13 + ((s And &HFFFF0000) \ &H10000) * b1 + ((((s And &HFFFF&) * b1) And &HFFFF0000) \ &H10000)
                s = sLTP_Q15(pred_lag_ptr - 2)            ' silk_SMLAWB( LTP_pred_Q13, pred_lag_ptr[ -2 ], B_Q14[ 2 ] )
                LTP_pred_Q13 = LTP_pred_Q13 + ((s And &HFFFF0000) \ &H10000) * b2 + ((((s And &HFFFF&) * b2) And &HFFFF0000) \ &H10000)
                s = sLTP_Q15(pred_lag_ptr - 3)            ' silk_SMLAWB( LTP_pred_Q13, pred_lag_ptr[ -3 ], B_Q14[ 3 ] )
                LTP_pred_Q13 = LTP_pred_Q13 + ((s And &HFFFF0000) \ &H10000) * b3 + ((((s And &HFFFF&) * b3) And &HFFFF0000) \ &H10000)
                s = sLTP_Q15(pred_lag_ptr - 4)            ' silk_SMLAWB( LTP_pred_Q13, pred_lag_ptr[ -4 ], B_Q14[ 4 ] )
                LTP_pred_Q13 = LTP_pred_Q13 + ((s And &HFFFF0000) \ &H10000) * b4 + ((((s And &HFFFF&) * b4) And &HFFFF0000) \ &H10000)
                pred_lag_ptr = pred_lag_ptr + 1

                '/* Generate LPC excitation */
                x = psDec.exc_Q14(pexc_Q14 + i) + LTP_pred_Q13 * 2      ' silk_ADD_LSHIFT32( pexc_Q14[ i ], LTP_pred_Q13, 1 )
                res_Q14(i) = x

                '/* Update states */
                sLTP_Q15(sLTP_buf_idx) = x * 2                          ' silk_LSHIFT( pres_Q14[ i ], 1 )
                sLTP_buf_idx = sLTP_buf_idx + 1
            Next
        Else
            ' pres_Q14 = pexc_Q14; (VB) copied
            silk_memcpy_i32 res_Q14, 0, psDec.exc_Q14, pexc_Q14, psDec.subfr_length
        End If

        For i = 0 To psDec.subfr_length - 1
            '/* Short-term prediction */
            ' celt_assert( psDec->LPC_order == 10 || psDec->LPC_order == 16 );
            '/* Avoids introducing a bias because silk_SMLAWB() always rounds to -inf */
            LPC_pred_Q10 = silk_RSHIFT(order, 1)
            For j = 0 To order - 1
                ' silk_SMLAWB( LPC_pred_Q10, sLPC_Q14[ MAX_LPC_ORDER + i - j - 1 ], A_Q12_tmp[ j ] )
                s = sLPC_Q14(MAX_LPC_ORDER + i - j - 1)
                c = A_Q12_tmp(j)
                LPC_pred_Q10 = LPC_pred_Q10 + ((s And &HFFFF0000) \ &H10000) * c + ((((s And &HFFFF&) * c) And &HFFFF0000) \ &H10000)
            Next

            '/* Add prediction to LPC excitation */
            sLPC_Q14(MAX_LPC_ORDER + i) = silk_ADD_SAT32(res_Q14(i), silk_LSHIFT_SAT32(LPC_pred_Q10, 4))

            '/* Scale with gain */
            xq(pxq + i) = silk_SAT16(silk_RSHIFT_ROUND(silk_SMULWW(sLPC_Q14(MAX_LPC_ORDER + i), Gain_Q10), 8))
        Next

        '/* Update LPC filter state */
        silk_memcpy_i32 sLPC_Q14, 0, sLPC_Q14, psDec.subfr_length, MAX_LPC_ORDER
        pexc_Q14 = pexc_Q14 + psDec.subfr_length
        pxq = pxq + psDec.subfr_length
    Next

    '/* Save LPC state */
    silk_memcpy_i32 psDec.sLPC_Q14_buf, 0, sLPC_Q14, 0, MAX_LPC_ORDER
End Sub
