Attribute VB_Name = "silk_decode_indices_c"
' libopus 1.5.2 (ref/opus/silk): decode_indices.c.
Option Explicit

'/* Decode side-information parameters from payload */
Public Sub silk_decode_indices(psDec As silk_decoder_state, psRangeDec As ec_ctx, ByVal FrameIndex As Long, _
        ByVal decode_LBRR As Long, ByVal condCoding As Long)
    Dim i As Long, k As Long, Ix As Long
    Dim decode_absolute_lagIndex As Long, delta_lagIndex As Long
    Dim ec_ix(0 To 15) As Integer                       ' [ MAX_LPC_ORDER ]
    Dim pred_Q8(0 To 15) As Byte                        ' [ MAX_LPC_ORDER ]
    Dim cb As Long

    cb = psDec.psNLSF_CB
    '/*******************************************/
    '/* Decode signal type and quantizer offset */
    '/*******************************************/
    If decode_LBRR <> 0 Or psDec.VAD_flags(FrameIndex) <> 0 Then
        Ix = ec_dec_icdf(psRangeDec, silk_rom_u8, silk_type_offset_VAD_iCDF, 8) + 2
    Else
        Ix = ec_dec_icdf(psRangeDec, silk_rom_u8, silk_type_offset_no_VAD_iCDF, 8)
    End If
    psDec.indices.signalType = silk_RSHIFT(Ix, 1)
    psDec.indices.quantOffsetType = Ix And 1

    '/****************/
    '/* Decode gains */
    '/****************/
    '/* First subframe */
    If condCoding = CODE_CONDITIONALLY Then
        '/* Conditional coding */
        psDec.indices.GainsIndices(0) = ec_dec_icdf(psRangeDec, silk_rom_u8, silk_delta_gain_iCDF, 8)
    Else
        '/* Independent coding, in two stages: MSB bits followed by 3 LSBs */
        psDec.indices.GainsIndices(0) = silk_LSHIFT(ec_dec_icdf(psRangeDec, silk_rom_u8, silk_gain_iCDF + psDec.indices.signalType * (N_LEVELS_QGAIN \ 8), 8), 3)
        psDec.indices.GainsIndices(0) = psDec.indices.GainsIndices(0) + ec_dec_icdf(psRangeDec, silk_rom_u8, silk_uniform8_iCDF, 8)
    End If

    '/* Remaining subframes */
    For i = 1 To psDec.nb_subfr - 1
        psDec.indices.GainsIndices(i) = ec_dec_icdf(psRangeDec, silk_rom_u8, silk_delta_gain_iCDF, 8)
    Next

    '/**********************/
    '/* Decode LSF Indices */
    '/**********************/
    psDec.indices.NLSFIndices(0) = ec_dec_icdf(psRangeDec, silk_rom_u8, silk_NLSF_CBs(cb).CB1_iCDF + silk_RSHIFT(psDec.indices.signalType, 1) * silk_NLSF_CBs(cb).nVectors, 8)
    silk_NLSF_unpack ec_ix, pred_Q8, cb, psDec.indices.NLSFIndices(0)
    ' celt_assert( psDec->psNLSF_CB->order == psDec->LPC_order );
    For i = 0 To silk_NLSF_CBs(cb).order - 1
        Ix = ec_dec_icdf(psRangeDec, silk_rom_u8, silk_NLSF_CBs(cb).ec_iCDF + ec_ix(i), 8)
        If Ix = 0 Then
            Ix = Ix - ec_dec_icdf(psRangeDec, silk_rom_u8, silk_NLSF_EXT_iCDF, 8)
        ElseIf Ix = 2 * NLSF_QUANT_MAX_AMPLITUDE Then
            Ix = Ix + ec_dec_icdf(psRangeDec, silk_rom_u8, silk_NLSF_EXT_iCDF, 8)
        End If
        psDec.indices.NLSFIndices(i + 1) = Ix - NLSF_QUANT_MAX_AMPLITUDE
    Next

    '/* Decode LSF interpolation factor */
    If psDec.nb_subfr = MAX_NB_SUBFR Then
        psDec.indices.NLSFInterpCoef_Q2 = ec_dec_icdf(psRangeDec, silk_rom_u8, silk_NLSF_interpolation_factor_iCDF, 8)
    Else
        psDec.indices.NLSFInterpCoef_Q2 = 4
    End If

    If psDec.indices.signalType = TYPE_VOICED Then
        '/*********************/
        '/* Decode pitch lags */
        '/*********************/
        '/* Get lag index */
        decode_absolute_lagIndex = 1
        If condCoding = CODE_CONDITIONALLY And psDec.ec_prevSignalType = TYPE_VOICED Then
            '/* Decode Delta index */
            delta_lagIndex = ec_dec_icdf(psRangeDec, silk_rom_u8, silk_pitch_delta_iCDF, 8)
            If delta_lagIndex > 0 Then
                delta_lagIndex = delta_lagIndex - 9
                psDec.indices.lagIndex = CLng(psDec.ec_prevLagIndex) + delta_lagIndex
                decode_absolute_lagIndex = 0
            End If
        End If
        If decode_absolute_lagIndex <> 0 Then
            '/* Absolute decoding */
            psDec.indices.lagIndex = ec_dec_icdf(psRangeDec, silk_rom_u8, silk_pitch_lag_iCDF, 8) * silk_RSHIFT(psDec.fs_kHz, 1)
            psDec.indices.lagIndex = CLng(psDec.indices.lagIndex) + ec_dec_icdf(psRangeDec, silk_rom_u8, psDec.pitch_lag_low_bits_iCDF, 8)
        End If
        psDec.ec_prevLagIndex = psDec.indices.lagIndex

        '/* Get countour index */
        psDec.indices.contourIndex = ec_dec_icdf(psRangeDec, silk_rom_u8, psDec.pitch_contour_iCDF, 8)

        '/********************/
        '/* Decode LTP gains */
        '/********************/
        '/* Decode PERIndex value */
        psDec.indices.PERIndex = ec_dec_icdf(psRangeDec, silk_rom_u8, silk_LTP_per_index_iCDF, 8)

        For k = 0 To psDec.nb_subfr - 1
            psDec.indices.LTPIndex(k) = ec_dec_icdf(psRangeDec, silk_rom_u8, silk_LTP_gain_iCDF_ptrs(psDec.indices.PERIndex), 8)
        Next

        '/**********************/
        '/* Decode LTP scaling */
        '/**********************/
        If condCoding = CODE_INDEPENDENTLY Then
            psDec.indices.LTP_scaleIndex = ec_dec_icdf(psRangeDec, silk_rom_u8, silk_LTPscale_iCDF, 8)
        Else
            psDec.indices.LTP_scaleIndex = 0
        End If
    End If
    psDec.ec_prevSignalType = psDec.indices.signalType

    '/***************/
    '/* Decode seed */
    '/***************/
    psDec.indices.Seed = ec_dec_icdf(psRangeDec, silk_rom_u8, silk_uniform4_iCDF, 8)
End Sub
