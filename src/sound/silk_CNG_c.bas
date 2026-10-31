Attribute VB_Name = "silk_CNG_c"
' libopus 1.5.2 (ref/opus/silk): CNG.c.
' (VB) frame is an array and an offset; the stack buffers are this module's arrays at their largest sizes. The
' synthesis filter's silk_SMLAWB is written out.
Option Explicit

Private CNG_sig_Q14(0 To 335) As Long                   ' [ length + MAX_LPC_ORDER ]
Private A_Q12(0 To 15) As Integer                       ' [ MAX_LPC_ORDER ]

'/* Generates excitation for CNG LPC synthesis */
Private Sub silk_CNG_exc(exc_Q14() As Long, ByVal excOffset As Long, exc_buf_Q14() As Long, ByVal length As Long, rand_seed As Long)
    Dim seed As Long
    Dim i As Long, idx As Long, exc_mask As Long

    exc_mask = CNG_BUF_MASK_MAX
    Do While exc_mask > length
        exc_mask = silk_RSHIFT(exc_mask, 1)
    Loop

    seed = rand_seed
    For i = 0 To length - 1
        seed = RAND_INCREMENT + seed * RAND_MULTIPLIER  ' silk_RAND( seed )
        idx = silk_RSHIFT(seed, 24) And exc_mask
        ' silk_assert( idx >= 0 );
        ' silk_assert( idx <= CNG_BUF_MASK_MAX );
        exc_Q14(excOffset + i) = exc_buf_Q14(idx)
    Next
    rand_seed = seed
End Sub

Public Sub silk_CNG_Reset(psDec As silk_decoder_state)
    Dim i As Long, NLSF_step_Q15 As Long, NLSF_acc_Q15 As Long

    NLSF_step_Q15 = silk_int16_MAX \ (psDec.LPC_order + 1)
    NLSF_acc_Q15 = 0
    For i = 0 To psDec.LPC_order - 1
        NLSF_acc_Q15 = NLSF_acc_Q15 + NLSF_step_Q15
        psDec.sCNG.CNG_smth_NLSF_Q15(i) = NLSF_acc_Q15
    Next
    psDec.sCNG.CNG_smth_Gain_Q16 = 0
    psDec.sCNG.rand_seed = 3176576
End Sub

'/* Updates CNG estimate, and applies the CNG when packet was lost   */
Public Sub silk_CNG(psDec As silk_decoder_state, psDecCtrl As silk_decoder_control, frame() As Integer, _
        ByVal frameOffset As Long, ByVal length As Long)
    Dim i As Long, subfr As Long
    Dim LPC_pred_Q10 As Long, max_Gain_Q16 As Long, gain_Q16 As Long, gain_Q10 As Long
    Dim j As Long, s As Long, c As Long

    With psDec.sCNG
    If psDec.fs_kHz <> .fs_kHz Then
        '/* Reset state */
        silk_CNG_Reset psDec

        .fs_kHz = psDec.fs_kHz
    End If
    If psDec.lossCnt = 0 And psDec.prevSignalType = TYPE_NO_VOICE_ACTIVITY Then
        '/* Update CNG parameters */

        '/* Smoothing of LSF's  */
        For i = 0 To psDec.LPC_order - 1
            .CNG_smth_NLSF_Q15(i) = CLng(.CNG_smth_NLSF_Q15(i)) + silk_SMULWB(CLng(psDec.prevNLSF_Q15(i)) - .CNG_smth_NLSF_Q15(i), CNG_NLSF_SMTH_Q16)
        Next
        '/* Find the subframe with the highest gain */
        max_Gain_Q16 = 0
        subfr = 0
        For i = 0 To psDec.nb_subfr - 1
            If psDecCtrl.Gains_Q16(i) > max_Gain_Q16 Then
                max_Gain_Q16 = psDecCtrl.Gains_Q16(i)
                subfr = i
            End If
        Next
        '/* Update CNG excitation buffer with excitation from this subframe */
        silk_memcpy_i32 .CNG_exc_buf_Q14, psDec.subfr_length, .CNG_exc_buf_Q14, 0, (psDec.nb_subfr - 1) * psDec.subfr_length
        silk_memcpy_i32 .CNG_exc_buf_Q14, 0, psDec.exc_Q14, subfr * psDec.subfr_length, psDec.subfr_length

        '/* Smooth gains */
        For i = 0 To psDec.nb_subfr - 1
            .CNG_smth_Gain_Q16 = .CNG_smth_Gain_Q16 + silk_SMULWB(psDecCtrl.Gains_Q16(i) - .CNG_smth_Gain_Q16, CNG_GAIN_SMTH_Q16)
            '/* If the smoothed gain is 3 dB greater than this subframe's gain, use this subframe's gain to adapt faster. */
            If silk_SMULWW(.CNG_smth_Gain_Q16, CNG_GAIN_SMTH_THRESHOLD_Q16) > psDecCtrl.Gains_Q16(i) Then
                .CNG_smth_Gain_Q16 = psDecCtrl.Gains_Q16(i)
            End If
        Next
    End If

    '/* Add CNG when packet is lost or during DTX */
    If psDec.lossCnt <> 0 Then
        '/* Generate CNG excitation */
        gain_Q16 = silk_SMULWW(psDec.sPLC.randScale_Q14, psDec.sPLC.prevGain_Q16(1))
        If gain_Q16 >= &H200000 Or .CNG_smth_Gain_Q16 > &H800000 Then
            gain_Q16 = silk_SMULTT(gain_Q16, gain_Q16)
            gain_Q16 = silk_SMULTT(.CNG_smth_Gain_Q16, .CNG_smth_Gain_Q16) - gain_Q16 * 32     ' silk_SUB_LSHIFT32( ..., gain_Q16, 5 )
            gain_Q16 = silk_SQRT_APPROX(gain_Q16) * &H10000                                     ' silk_LSHIFT32( ..., 16 )
        Else
            gain_Q16 = silk_SMULWW(gain_Q16, gain_Q16)
            gain_Q16 = silk_SMULWW(.CNG_smth_Gain_Q16, .CNG_smth_Gain_Q16) - gain_Q16 * 32     ' silk_SUB_LSHIFT32( ..., gain_Q16, 5 )
            gain_Q16 = silk_SQRT_APPROX(gain_Q16) * 256                                         ' silk_LSHIFT32( ..., 8 )
        End If
        gain_Q10 = silk_RSHIFT(gain_Q16, 6)

        silk_CNG_exc CNG_sig_Q14, MAX_LPC_ORDER, .CNG_exc_buf_Q14, length, .rand_seed

        '/* Convert CNG NLSF to filter representation */
        silk_NLSF2A A_Q12, 0, .CNG_smth_NLSF_Q15, psDec.LPC_order, psDec.arch

        '/* Generate CNG signal, by synthesis filtering */
        silk_memcpy_i32 CNG_sig_Q14, 0, .CNG_synth_state, 0, MAX_LPC_ORDER
        ' celt_assert( psDec->LPC_order == 10 || psDec->LPC_order == 16 );
        For i = 0 To length - 1
            '/* Avoids introducing a bias because silk_SMLAWB() always rounds to -inf */
            LPC_pred_Q10 = silk_RSHIFT(psDec.LPC_order, 1)
            For j = 0 To psDec.LPC_order - 1
                ' silk_SMLAWB( LPC_pred_Q10, CNG_sig_Q14[ MAX_LPC_ORDER + i - j - 1 ], A_Q12[ j ] )
                s = CNG_sig_Q14(MAX_LPC_ORDER + i - j - 1)
                c = A_Q12(j)
                LPC_pred_Q10 = LPC_pred_Q10 + ((s And &HFFFF0000) \ &H10000) * c + ((((s And &HFFFF&) * c) And &HFFFF0000) \ &H10000)
            Next

            '/* Update states */
            CNG_sig_Q14(MAX_LPC_ORDER + i) = silk_ADD_SAT32(CNG_sig_Q14(MAX_LPC_ORDER + i), silk_LSHIFT_SAT32(LPC_pred_Q10, 4))

            '/* Scale with Gain and add to input signal */
            frame(frameOffset + i) = silk_ADD_SAT16(frame(frameOffset + i), silk_SAT16(silk_RSHIFT_ROUND(silk_SMULWW(CNG_sig_Q14(MAX_LPC_ORDER + i), gain_Q10), 8)))

        Next
        silk_memcpy_i32 .CNG_synth_state, 0, CNG_sig_Q14, length, MAX_LPC_ORDER
    Else
        silk_memset_i32 .CNG_synth_state, 0, psDec.LPC_order
    End If
    End With
End Sub
