Attribute VB_Name = "silk_PLC_c"
' libopus 1.5.2 (ref/opus/silk): PLC.c and PLC.h (no DEEP_PLC, no SMALL_FOOTPRINT).
' (VB) frame is an array and an offset; the C's pointers into the state are offsets; the stack buffers are this
' module's arrays at their largest sizes. The LTP and LPC synthesis filters' silk_SMLAWB are written out.
Option Explicit

' PLC.h
Public Const BWE_COEF As Double = 0.99
Public Const V_PITCH_GAIN_START_MIN_Q14 As Long = 11469 '/* 0.7 in Q14               */
Public Const V_PITCH_GAIN_START_MAX_Q14 As Long = 15565 '/* 0.95 in Q14              */
Public Const MAX_PITCH_LAG_MS As Long = 18
Public Const RAND_BUF_SIZE As Long = 128
Public Const RAND_BUF_MASK As Long = 127                ' ( RAND_BUF_SIZE - 1 )
Public Const LOG2_INV_LPC_GAIN_HIGH_THRES As Long = 3   '/* 2^3 = 8 dB LPC gain      */
Public Const LOG2_INV_LPC_GAIN_LOW_THRES As Long = 8    '/* 2^8 = 24 dB LPC gain     */
Public Const PITCH_DRIFT_FAC_Q16 As Long = 655          '/* 0.01 in Q16              */

Private Const NB_ATT As Long = 2
' static const opus_int16 HARM_ATT_Q15[NB_ATT]              = { 32440, 31130 }; /* 0.99, 0.95 */
' static const opus_int16 PLC_RAND_ATTENUATE_V_Q15[NB_ATT]  = { 31130, 26214 }; /* 0.95, 0.8 */
' static const opus_int16 PLC_RAND_ATTENUATE_UV_Q15[NB_ATT] = { 32440, 29491 }; /* 0.99, 0.9 */

' (VB) the stack buffers
Private exc_buf(0 To 159) As Integer                    ' silk_PLC_energy: [ 2 * subfr_length ]
Private sLTP(0 To 319) As Integer                       ' silk_PLC_conceal: [ ltp_mem_length ]
Private sLTP_Q14(0 To 639) As Long                      ' silk_PLC_conceal: [ ltp_mem_length + frame_length ]
Private A_Q12(0 To 15) As Integer                       ' silk_PLC_conceal: [ MAX_LPC_ORDER ]

Private Function HARM_ATT_Q15(ByVal i As Long) As Long
    If i = 0 Then HARM_ATT_Q15 = 32440 Else HARM_ATT_Q15 = 31130
End Function

Private Function PLC_RAND_ATTENUATE_V_Q15(ByVal i As Long) As Long
    If i = 0 Then PLC_RAND_ATTENUATE_V_Q15 = 31130 Else PLC_RAND_ATTENUATE_V_Q15 = 26214
End Function

Private Function PLC_RAND_ATTENUATE_UV_Q15(ByVal i As Long) As Long
    If i = 0 Then PLC_RAND_ATTENUATE_UV_Q15 = 32440 Else PLC_RAND_ATTENUATE_UV_Q15 = 29491
End Function

Public Sub silk_PLC_Reset(psDec As silk_decoder_state)
    psDec.sPLC.pitchL_Q8 = psDec.frame_length * 128     ' silk_LSHIFT( psDec->frame_length, 8 - 1 )
    psDec.sPLC.prevGain_Q16(0) = 65536                  ' SILK_FIX_CONST( 1, 16 )
    psDec.sPLC.prevGain_Q16(1) = 65536                  ' SILK_FIX_CONST( 1, 16 )
    psDec.sPLC.subfr_length = 20
    psDec.sPLC.nb_subfr = 2
End Sub

Public Sub silk_PLC(psDec As silk_decoder_state, psDecCtrl As silk_decoder_control, frame() As Integer, _
        ByVal frameOffset As Long, ByVal lost As Long, ByVal arch As Long)
    '/* PLC control function */
    If psDec.fs_kHz <> psDec.sPLC.fs_kHz Then
        silk_PLC_Reset psDec
        psDec.sPLC.fs_kHz = psDec.fs_kHz
    End If

    If lost <> 0 Then
        '/****************************/
        '/* Generate Signal          */
        '/****************************/
        silk_PLC_conceal psDec, psDecCtrl, frame, frameOffset, arch

        psDec.lossCnt = psDec.lossCnt + 1
    Else
        '/****************************/
        '/* Update state             */
        '/****************************/
        silk_PLC_update psDec, psDecCtrl
    End If
End Sub

'/**************************************************/
'/* Update state of PLC                            */
'/**************************************************/
Private Sub silk_PLC_update(psDec As silk_decoder_state, psDecCtrl As silk_decoder_control)
    Dim LTP_Gain_Q14 As Long, temp_LTP_Gain_Q14 As Long
    Dim i As Long, j As Long
    Dim scale_Q10 As Long, scale_Q14 As Long, tmp As Long

    With psDec.sPLC
    '/* Update parameters used in case of packet loss */
    psDec.prevSignalType = psDec.indices.signalType
    LTP_Gain_Q14 = 0
    If psDec.indices.signalType = TYPE_VOICED Then
        '/* Find the parameters for the last subframe which contains a pitch pulse */
        j = 0
        Do While j * psDec.subfr_length < psDecCtrl.pitchL(psDec.nb_subfr - 1)
            If j = psDec.nb_subfr Then
                Exit Do
            End If
            temp_LTP_Gain_Q14 = 0
            For i = 0 To LTP_ORDER - 1
                temp_LTP_Gain_Q14 = temp_LTP_Gain_Q14 + psDecCtrl.LTPCoef_Q14((psDec.nb_subfr - 1 - j) * LTP_ORDER + i)
            Next
            If temp_LTP_Gain_Q14 > LTP_Gain_Q14 Then
                LTP_Gain_Q14 = temp_LTP_Gain_Q14
                silk_memcpy_i16 .LTPCoef_Q14, 0, psDecCtrl.LTPCoef_Q14, silk_SMULBB(psDec.nb_subfr - 1 - j, LTP_ORDER), LTP_ORDER

                .pitchL_Q8 = psDecCtrl.pitchL(psDec.nb_subfr - 1 - j) * 256
            End If
            j = j + 1
        Loop

        silk_memset_i16 .LTPCoef_Q14, 0, LTP_ORDER
        .LTPCoef_Q14(LTP_ORDER \ 2) = LTP_Gain_Q14

        '/* Limit LT coefs */
        If LTP_Gain_Q14 < V_PITCH_GAIN_START_MIN_Q14 Then
            tmp = V_PITCH_GAIN_START_MIN_Q14 * 1024         ' silk_LSHIFT( V_PITCH_GAIN_START_MIN_Q14, 10 )
            scale_Q10 = tmp \ silk_max(LTP_Gain_Q14, 1)
            For i = 0 To LTP_ORDER - 1
                .LTPCoef_Q14(i) = silk_RSHIFT(silk_SMULBB(.LTPCoef_Q14(i), scale_Q10), 10)
            Next
        ElseIf LTP_Gain_Q14 > V_PITCH_GAIN_START_MAX_Q14 Then
            tmp = V_PITCH_GAIN_START_MAX_Q14 * 16384        ' silk_LSHIFT( V_PITCH_GAIN_START_MAX_Q14, 14 )
            scale_Q14 = tmp \ silk_max(LTP_Gain_Q14, 1)
            For i = 0 To LTP_ORDER - 1
                .LTPCoef_Q14(i) = silk_RSHIFT(silk_SMULBB(.LTPCoef_Q14(i), scale_Q14), 14)
            Next
        End If
    Else
        .pitchL_Q8 = silk_SMULBB(psDec.fs_kHz, 18) * 256
        silk_memset_i16 .LTPCoef_Q14, 0, LTP_ORDER
    End If

    '/* Save LPC coeficients */
    silk_memcpy_i16 .prevLPC_Q12, 0, psDecCtrl.PredCoef_Q12, MAX_LPC_ORDER, psDec.LPC_order
    .prevLTP_scale_Q14 = psDecCtrl.LTP_scale_Q14

    '/* Save last two gains */
    silk_memcpy_i32 .prevGain_Q16, 0, psDecCtrl.Gains_Q16, psDec.nb_subfr - 2, 2

    .subfr_length = psDec.subfr_length
    .nb_subfr = psDec.nb_subfr
    End With
End Sub

Private Sub silk_PLC_energy(energy1 As Long, shift1 As Long, energy2 As Long, shift2 As Long, _
        exc_Q14() As Long, prevGain_Q10() As Long, ByVal subfr_length As Long, ByVal nb_subfr As Long)
    Dim i As Long, k As Long
    Dim exc_buf_ptr As Long                             ' opus_int16 *: an offset into exc_buf

    '/* Find random noise component */
    '/* Scale previous excitation signal */
    exc_buf_ptr = 0
    For k = 0 To 1
        For i = 0 To subfr_length - 1
            exc_buf(exc_buf_ptr + i) = silk_SAT16(silk_RSHIFT( _
                silk_SMULWW(exc_Q14(i + (k + nb_subfr - 2) * subfr_length), prevGain_Q10(k)), 8))
        Next
        exc_buf_ptr = exc_buf_ptr + subfr_length
    Next
    '/* Find the subframe with lowest energy of the last two and use that as random noise generator */
    silk_sum_sqr_shift energy1, shift1, exc_buf, 0, subfr_length
    silk_sum_sqr_shift energy2, shift2, exc_buf, subfr_length, subfr_length
End Sub

Private Sub silk_PLC_conceal(psDec As silk_decoder_state, psDecCtrl As silk_decoder_control, frame() As Integer, _
        ByVal frameOffset As Long, ByVal arch As Long)
    Dim i As Long, j As Long, k As Long
    Dim lag As Long, idx As Long, sLTP_buf_idx As Long, shift1 As Long, shift2 As Long
    Dim rand_seed As Long, harm_Gain_Q15 As Long, rand_Gain_Q15 As Long, inv_gain_Q30 As Long
    Dim energy1 As Long, energy2 As Long
    Dim rand_ptr As Long, pred_lag_ptr As Long          ' opus_int32 *: offsets into psDec->exc_Q14, sLTP_Q14
    Dim LPC_pred_Q10 As Long, LTP_pred_Q12 As Long
    Dim rand_scale_Q14 As Integer
    Dim sLPC_Q14_ptr As Long                            ' opus_int32 *: an offset into sLTP_Q14
    Dim prevGain_Q10(0 To 1) As Long
    Dim invGain_Q30 As Long, down_scale_Q30 As Long
    Dim s As Long, c As Long, b0 As Long, b1 As Long, b2 As Long, b3 As Long, b4 As Long, r As Long
    ' B_Q14 = psPLC->LTPCoef_Q14 (used in place)

    prevGain_Q10(0) = silk_RSHIFT(psDec.sPLC.prevGain_Q16(0), 6)
    prevGain_Q10(1) = silk_RSHIFT(psDec.sPLC.prevGain_Q16(1), 6)

    If psDec.first_frame_after_reset <> 0 Then
       silk_memset_i16 psDec.sPLC.prevLPC_Q12, 0, MAX_LPC_ORDER
    End If

    silk_PLC_energy energy1, shift1, energy2, shift2, psDec.exc_Q14, prevGain_Q10, psDec.subfr_length, psDec.nb_subfr

    If silk_RSHIFT(energy1, shift2) < silk_RSHIFT(energy2, shift1) Then
        '/* First sub-frame has lowest energy */
        rand_ptr = silk_max_int(0, (psDec.sPLC.nb_subfr - 1) * psDec.sPLC.subfr_length - RAND_BUF_SIZE)
    Else
        '/* Second sub-frame has lowest energy */
        rand_ptr = silk_max_int(0, psDec.sPLC.nb_subfr * psDec.sPLC.subfr_length - RAND_BUF_SIZE)
    End If

    '/* Set up Gain to random noise component */
    rand_scale_Q14 = psDec.sPLC.randScale_Q14

    '/* Set up attenuation gains */
    harm_Gain_Q15 = HARM_ATT_Q15(silk_min_int(NB_ATT - 1, psDec.lossCnt))
    If psDec.prevSignalType = TYPE_VOICED Then
        rand_Gain_Q15 = PLC_RAND_ATTENUATE_V_Q15(silk_min_int(NB_ATT - 1, psDec.lossCnt))
    Else
        rand_Gain_Q15 = PLC_RAND_ATTENUATE_UV_Q15(silk_min_int(NB_ATT - 1, psDec.lossCnt))
    End If

    '/* LPC concealment. Apply BWE to previous LPC */
    silk_bwexpander psDec.sPLC.prevLPC_Q12, 0, psDec.LPC_order, 64881     ' SILK_FIX_CONST( BWE_COEF, 16 )

    '/* Preload LPC coeficients to array on stack. Gives small performance gain */
    silk_memcpy_i16 A_Q12, 0, psDec.sPLC.prevLPC_Q12, 0, psDec.LPC_order

    '/* First Lost frame */
    If psDec.lossCnt = 0 Then
        rand_scale_Q14 = 16384                          ' 1 << 14

        '/* Reduce random noise Gain for voiced frames */
        If psDec.prevSignalType = TYPE_VOICED Then
            For i = 0 To LTP_ORDER - 1
                rand_scale_Q14 = CLng(rand_scale_Q14) - psDec.sPLC.LTPCoef_Q14(i)
            Next
            rand_scale_Q14 = silk_max_16(3277, rand_scale_Q14)          '/* 0.2 */
            rand_scale_Q14 = silk_RSHIFT(silk_SMULBB(rand_scale_Q14, psDec.sPLC.prevLTP_scale_Q14), 14)
        Else
            '/* Reduce random noise for unvoiced frames with high LPC gain */
            invGain_Q30 = silk_LPC_inverse_pred_gain(psDec.sPLC.prevLPC_Q12, 0, psDec.LPC_order, arch)

            down_scale_Q30 = silk_min_32(silk_RSHIFT(&H40000000, LOG2_INV_LPC_GAIN_HIGH_THRES), invGain_Q30)
            down_scale_Q30 = silk_max_32(silk_RSHIFT(&H40000000, LOG2_INV_LPC_GAIN_LOW_THRES), down_scale_Q30)
            down_scale_Q30 = silk_LSHIFT(down_scale_Q30, LOG2_INV_LPC_GAIN_HIGH_THRES)

            rand_Gain_Q15 = silk_RSHIFT(silk_SMULWB(down_scale_Q30, rand_Gain_Q15), 14)
        End If
    End If

    rand_seed = psDec.sPLC.rand_seed
    lag = silk_RSHIFT_ROUND(psDec.sPLC.pitchL_Q8, 8)
    sLTP_buf_idx = psDec.ltp_mem_length

    '/* Rewhiten LTP state */
    idx = psDec.ltp_mem_length - lag - psDec.LPC_order - LTP_ORDER \ 2
    ' celt_assert( idx > 0 );
    silk_LPC_analysis_filter sLTP, idx, psDec.outBuf, idx, A_Q12, 0, psDec.ltp_mem_length - idx, psDec.LPC_order, arch
    '/* Scale LTP state */
    inv_gain_Q30 = silk_INVERSE32_varQ(psDec.sPLC.prevGain_Q16(1), 46)
    inv_gain_Q30 = silk_min(inv_gain_Q30, silk_int32_MAX \ 2)
    For i = idx + psDec.LPC_order To psDec.ltp_mem_length - 1
        sLTP_Q14(i) = silk_SMULWB(inv_gain_Q30, sLTP(i))
    Next

    '/***************************/
    '/* LTP synthesis filtering */
    '/***************************/
    For k = 0 To psDec.nb_subfr - 1
        '/* Set up pointer */
        pred_lag_ptr = sLTP_buf_idx - lag + LTP_ORDER \ 2
        b0 = psDec.sPLC.LTPCoef_Q14(0)
        b1 = psDec.sPLC.LTPCoef_Q14(1)
        b2 = psDec.sPLC.LTPCoef_Q14(2)
        b3 = psDec.sPLC.LTPCoef_Q14(3)
        b4 = psDec.sPLC.LTPCoef_Q14(4)
        For i = 0 To psDec.subfr_length - 1
            '/* Unrolled loop */
            '/* Avoids introducing a bias because silk_SMLAWB() always rounds to -inf */
            LTP_pred_Q12 = 2
            s = sLTP_Q14(pred_lag_ptr)
            LTP_pred_Q12 = LTP_pred_Q12 + ((s And &HFFFF0000) \ &H10000) * b0 + ((((s And &HFFFF&) * b0) And &HFFFF0000) \ &H10000)
            s = sLTP_Q14(pred_lag_ptr - 1)
            LTP_pred_Q12 = LTP_pred_Q12 + ((s And &HFFFF0000) \ &H10000) * b1 + ((((s And &HFFFF&) * b1) And &HFFFF0000) \ &H10000)
            s = sLTP_Q14(pred_lag_ptr - 2)
            LTP_pred_Q12 = LTP_pred_Q12 + ((s And &HFFFF0000) \ &H10000) * b2 + ((((s And &HFFFF&) * b2) And &HFFFF0000) \ &H10000)
            s = sLTP_Q14(pred_lag_ptr - 3)
            LTP_pred_Q12 = LTP_pred_Q12 + ((s And &HFFFF0000) \ &H10000) * b3 + ((((s And &HFFFF&) * b3) And &HFFFF0000) \ &H10000)
            s = sLTP_Q14(pred_lag_ptr - 4)
            LTP_pred_Q12 = LTP_pred_Q12 + ((s And &HFFFF0000) \ &H10000) * b4 + ((((s And &HFFFF&) * b4) And &HFFFF0000) \ &H10000)
            pred_lag_ptr = pred_lag_ptr + 1

            '/* Generate LPC excitation */
            rand_seed = RAND_INCREMENT + rand_seed * RAND_MULTIPLIER    ' silk_RAND( rand_seed )
            idx = silk_RSHIFT(rand_seed, 25) And RAND_BUF_MASK
            r = psDec.exc_Q14(rand_ptr + idx)
            c = rand_scale_Q14
            ' silk_LSHIFT32( silk_SMLAWB( LTP_pred_Q12, rand_ptr[ idx ], rand_scale_Q14 ), 2 )
            sLTP_Q14(sLTP_buf_idx) = (LTP_pred_Q12 + ((r And &HFFFF0000) \ &H10000) * c + ((((r And &HFFFF&) * c) And &HFFFF0000) \ &H10000)) * 4
            sLTP_buf_idx = sLTP_buf_idx + 1
        Next

        '/* Gradually reduce LTP gain */
        For j = 0 To LTP_ORDER - 1
            psDec.sPLC.LTPCoef_Q14(j) = silk_RSHIFT(silk_SMULBB(harm_Gain_Q15, psDec.sPLC.LTPCoef_Q14(j)), 15)
        Next
        '/* Gradually reduce excitation gain */
        rand_scale_Q14 = silk_RSHIFT(silk_SMULBB(rand_scale_Q14, rand_Gain_Q15), 15)

        '/* Slowly increase pitch lag */
        psDec.sPLC.pitchL_Q8 = silk_SMLAWB(psDec.sPLC.pitchL_Q8, psDec.sPLC.pitchL_Q8, PITCH_DRIFT_FAC_Q16)
        psDec.sPLC.pitchL_Q8 = silk_min_32(psDec.sPLC.pitchL_Q8, silk_SMULBB(MAX_PITCH_LAG_MS, psDec.fs_kHz) * 256)
        lag = silk_RSHIFT_ROUND(psDec.sPLC.pitchL_Q8, 8)
    Next

    '/***************************/
    '/* LPC synthesis filtering */
    '/***************************/
    sLPC_Q14_ptr = psDec.ltp_mem_length - MAX_LPC_ORDER

    '/* Copy LPC state */
    silk_memcpy_i32 sLTP_Q14, sLPC_Q14_ptr, psDec.sLPC_Q14_buf, 0, MAX_LPC_ORDER

    ' celt_assert( psDec->LPC_order >= 10 ); /* check that unrolling works */
    For i = 0 To psDec.frame_length - 1
        '/* partly unrolled */
        '/* Avoids introducing a bias because silk_SMLAWB() always rounds to -inf */
        LPC_pred_Q10 = silk_RSHIFT(psDec.LPC_order, 1)
        For j = 0 To psDec.LPC_order - 1
            ' silk_SMLAWB( LPC_pred_Q10, sLPC_Q14_ptr[ MAX_LPC_ORDER + i - j - 1 ], A_Q12[ j ] )
            s = sLTP_Q14(sLPC_Q14_ptr + MAX_LPC_ORDER + i - j - 1)
            c = A_Q12(j)
            LPC_pred_Q10 = LPC_pred_Q10 + ((s And &HFFFF0000) \ &H10000) * c + ((((s And &HFFFF&) * c) And &HFFFF0000) \ &H10000)
        Next

        '/* Add prediction to LPC excitation */
        sLTP_Q14(sLPC_Q14_ptr + MAX_LPC_ORDER + i) = silk_ADD_SAT32(sLTP_Q14(sLPC_Q14_ptr + MAX_LPC_ORDER + i), _
                                            silk_LSHIFT_SAT32(LPC_pred_Q10, 4))

        '/* Scale with Gain */
        frame(frameOffset + i) = silk_SAT16(silk_SAT16(silk_RSHIFT_ROUND(silk_SMULWW(sLTP_Q14(sLPC_Q14_ptr + MAX_LPC_ORDER + i), prevGain_Q10(1)), 8)))
    Next

    '/* Save LPC state */
    silk_memcpy_i32 psDec.sLPC_Q14_buf, 0, sLTP_Q14, sLPC_Q14_ptr + psDec.frame_length, MAX_LPC_ORDER

    '/**************************************/
    '/* Update states                      */
    '/**************************************/
    psDec.sPLC.rand_seed = rand_seed
    psDec.sPLC.randScale_Q14 = rand_scale_Q14
    For i = 0 To MAX_NB_SUBFR - 1
        psDecCtrl.pitchL(i) = lag
    Next
End Sub

'/* Glues concealed frames with new good received frames */
Public Sub silk_PLC_glue_frames(psDec As silk_decoder_state, frame() As Integer, ByVal frameOffset As Long, ByVal length As Long)
    Dim i As Long, energy_shift As Long
    Dim energy As Long
    Dim frac_Q24 As Long, LZ As Long
    Dim gain_Q16 As Long, slope_Q16 As Long

    With psDec.sPLC
    If psDec.lossCnt <> 0 Then
        '/* Calculate energy in concealed residual */
        silk_sum_sqr_shift .conc_energy, .conc_energy_shift, frame, frameOffset, length

        .last_frame_lost = 1
    Else
        If psDec.sPLC.last_frame_lost <> 0 Then
            '/* Calculate residual in decoded signal if last frame was lost */
            silk_sum_sqr_shift energy, energy_shift, frame, frameOffset, length

            '/* Normalize energies */
            If energy_shift > .conc_energy_shift Then
                .conc_energy = silk_RSHIFT(.conc_energy, energy_shift - .conc_energy_shift)
            ElseIf energy_shift < .conc_energy_shift Then
                energy = silk_RSHIFT(energy, .conc_energy_shift - energy_shift)
            End If

            '/* Fade in the energy difference */
            If energy > .conc_energy Then
                LZ = silk_CLZ32(.conc_energy)
                LZ = LZ - 1
                .conc_energy = silk_LSHIFT(.conc_energy, LZ)
                energy = silk_RSHIFT(energy, silk_max_32(24 - LZ, 0))

                frac_Q24 = .conc_energy \ silk_max(energy, 1)

                gain_Q16 = silk_SQRT_APPROX(frac_Q24) * 16
                slope_Q16 = (&H10000 - gain_Q16) \ length
                '/* Make slope 4x steeper to avoid missing onsets after DTX */
                slope_Q16 = slope_Q16 * 4

                For i = 0 To length - 1
                    frame(frameOffset + i) = silk_SMULWB(gain_Q16, frame(frameOffset + i))
                    gain_Q16 = gain_Q16 + slope_Q16
                    If gain_Q16 > &H10000 Then
                        Exit For
                    End If
                Next
            End If
        End If
        .last_frame_lost = 0
    End If
    End With
End Sub
