Attribute VB_Name = "silk_decoder_set_fs_c"
' libopus 1.5.2 (ref/opus/silk): decoder_set_fs.c.
Option Explicit

'/* Set decoder sampling rate */
Public Function silk_decoder_set_fs(psDec As silk_decoder_state, ByVal fs_kHz As Long, ByVal fs_API_Hz As Long) As Long
    Dim frame_length As Long, ret As Long

    ' celt_assert( fs_kHz == 8 || fs_kHz == 12 || fs_kHz == 16 );
    ' celt_assert( psDec->nb_subfr == MAX_NB_SUBFR || psDec->nb_subfr == MAX_NB_SUBFR/2 );

    '/* New (sub)frame length */
    psDec.subfr_length = silk_SMULBB(SUB_FRAME_LENGTH_MS, fs_kHz)
    frame_length = silk_SMULBB(psDec.nb_subfr, psDec.subfr_length)

    '/* Initialize resampler when switching internal or external sampling frequency */
    If psDec.fs_kHz <> fs_kHz Or psDec.fs_API_hz <> fs_API_Hz Then
        '/* Initialize the resampler for dec_API.c preparing resampling from fs_kHz to API_fs_Hz */
        ret = ret + silk_resampler_init(psDec.resampler_state, silk_SMULBB(fs_kHz, 1000), fs_API_Hz, 0)

        psDec.fs_API_hz = fs_API_Hz
    End If

    If psDec.fs_kHz <> fs_kHz Or frame_length <> psDec.frame_length Then
        If fs_kHz = 8 Then
            If psDec.nb_subfr = MAX_NB_SUBFR Then
                psDec.pitch_contour_iCDF = silk_pitch_contour_NB_iCDF
            Else
                psDec.pitch_contour_iCDF = silk_pitch_contour_10_ms_NB_iCDF
            End If
        Else
            If psDec.nb_subfr = MAX_NB_SUBFR Then
                psDec.pitch_contour_iCDF = silk_pitch_contour_iCDF
            Else
                psDec.pitch_contour_iCDF = silk_pitch_contour_10_ms_iCDF
            End If
        End If
        If psDec.fs_kHz <> fs_kHz Then
            psDec.ltp_mem_length = silk_SMULBB(LTP_MEM_LENGTH_MS, fs_kHz)
            If fs_kHz = 8 Or fs_kHz = 12 Then
                psDec.LPC_order = MIN_LPC_ORDER
                psDec.psNLSF_CB = silk_NLSF_CB_NB_MB
            Else
                psDec.LPC_order = MAX_LPC_ORDER
                psDec.psNLSF_CB = silk_NLSF_CB_WB
            End If
            If fs_kHz = 16 Then
                psDec.pitch_lag_low_bits_iCDF = silk_uniform8_iCDF
            ElseIf fs_kHz = 12 Then
                psDec.pitch_lag_low_bits_iCDF = silk_uniform6_iCDF
            ElseIf fs_kHz = 8 Then
                psDec.pitch_lag_low_bits_iCDF = silk_uniform4_iCDF
            Else
                '/* unsupported sampling rate */
                ' celt_assert( 0 );
            End If
            psDec.first_frame_after_reset = 1
            psDec.lagPrev = 100
            psDec.LastGainIndex = 10
            psDec.prevSignalType = TYPE_NO_VOICE_ACTIVITY
            silk_memset_i16 psDec.outBuf, 0, MAX_FRAME_LENGTH + 2 * MAX_SUB_FRAME_LENGTH
            silk_memset_i32 psDec.sLPC_Q14_buf, 0, MAX_LPC_ORDER
        End If

        psDec.fs_kHz = fs_kHz
        psDec.frame_length = frame_length
    End If

    '/* Check that settings are valid */
    ' celt_assert( psDec->frame_length > 0 && psDec->frame_length <= MAX_FRAME_LENGTH );

    silk_decoder_set_fs = ret
End Function
