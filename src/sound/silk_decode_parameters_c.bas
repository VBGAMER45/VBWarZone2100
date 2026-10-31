Attribute VB_Name = "silk_decode_parameters_c"
' libopus 1.5.2 (ref/opus/silk): decode_parameters.c.
' (VB) psDecCtrl->PredCoef_Q12[ k ] is PredCoef_Q12 from k * MAX_LPC_ORDER.
Option Explicit

'/* Decode parameters from payload */
Public Sub silk_decode_parameters(psDec As silk_decoder_state, psDecCtrl As silk_decoder_control, ByVal condCoding As Long)
    Dim i As Long, k As Long, Ix As Long
    Dim pNLSF_Q15(0 To 15) As Integer                   ' [ MAX_LPC_ORDER ]
    Dim pNLSF0_Q15(0 To 15) As Integer                  ' [ MAX_LPC_ORDER ]
    Dim cbk_ptr_Q7 As Long                              ' const opus_int8 *: an offset into silk_rom_i8
    Dim cond As Long

    '/* Dequant Gains */
    If condCoding = CODE_CONDITIONALLY Then cond = 1
    silk_gains_dequant psDecCtrl.Gains_Q16, psDec.indices.GainsIndices, psDec.LastGainIndex, cond, psDec.nb_subfr

    '/****************/
    '/* Decode NLSFs */
    '/****************/
    silk_NLSF_decode pNLSF_Q15, psDec.indices.NLSFIndices, psDec.psNLSF_CB

    '/* Convert NLSF parameters to AR prediction filter coefficients */
    silk_NLSF2A psDecCtrl.PredCoef_Q12, MAX_LPC_ORDER, pNLSF_Q15, psDec.LPC_order, psDec.arch

    '/* If just reset, e.g., because internal Fs changed, do not allow interpolation */
    '/* improves the case of packet loss in the first frame after a switch           */
    If psDec.first_frame_after_reset = 1 Then
        psDec.indices.NLSFInterpCoef_Q2 = 4
    End If

    If psDec.indices.NLSFInterpCoef_Q2 < 4 Then
        '/* Calculation of the interpolated NLSF0 vector from the interpolation factor, */
        '/* the previous NLSF1, and the current NLSF1                                   */
        For i = 0 To psDec.LPC_order - 1
            pNLSF0_Q15(i) = CLng(psDec.prevNLSF_Q15(i)) + silk_RSHIFT(psDec.indices.NLSFInterpCoef_Q2 * _
                (CLng(pNLSF_Q15(i)) - psDec.prevNLSF_Q15(i)), 2)
        Next

        '/* Convert NLSF parameters to AR prediction filter coefficients */
        silk_NLSF2A psDecCtrl.PredCoef_Q12, 0, pNLSF0_Q15, psDec.LPC_order, psDec.arch
    Else
        '/* Copy LPC coefficients for first half from second half */
        silk_memcpy_i16 psDecCtrl.PredCoef_Q12, 0, psDecCtrl.PredCoef_Q12, MAX_LPC_ORDER, psDec.LPC_order
    End If

    silk_memcpy_i16 psDec.prevNLSF_Q15, 0, pNLSF_Q15, 0, psDec.LPC_order

    '/* After a packet loss do BWE of LPC coefs */
    If psDec.lossCnt <> 0 Then
        silk_bwexpander psDecCtrl.PredCoef_Q12, 0, psDec.LPC_order, BWE_AFTER_LOSS_Q16
        silk_bwexpander psDecCtrl.PredCoef_Q12, MAX_LPC_ORDER, psDec.LPC_order, BWE_AFTER_LOSS_Q16
    End If

    If psDec.indices.signalType = TYPE_VOICED Then
        '/*********************/
        '/* Decode pitch lags */
        '/*********************/

        '/* Decode pitch values */
        silk_decode_pitch psDec.indices.lagIndex, psDec.indices.contourIndex, psDecCtrl.pitchL, psDec.fs_kHz, psDec.nb_subfr

        '/* Decode Codebook Index */
        cbk_ptr_Q7 = silk_LTP_vq_ptrs_Q7(psDec.indices.PERIndex)    '/* set pointer to start of codebook */

        For k = 0 To psDec.nb_subfr - 1
            Ix = psDec.indices.LTPIndex(k)
            For i = 0 To LTP_ORDER - 1
                psDecCtrl.LTPCoef_Q14(k * LTP_ORDER + i) = CLng(silk_rom_i8(cbk_ptr_Q7 + Ix * LTP_ORDER + i)) * 128
            Next
        Next

        '/**********************/
        '/* Decode LTP scaling */
        '/**********************/
        Ix = psDec.indices.LTP_scaleIndex
        psDecCtrl.LTP_scale_Q14 = silk_rom_i16(silk_LTPScales_table_Q14 + Ix)
    Else
        For i = 0 To psDec.nb_subfr - 1
            psDecCtrl.pitchL(i) = 0
        Next
        silk_memset_i16 psDecCtrl.LTPCoef_Q14, 0, LTP_ORDER * psDec.nb_subfr
        psDec.indices.PERIndex = 0
        psDecCtrl.LTP_scale_Q14 = 0
    End If
End Sub
