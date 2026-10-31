Attribute VB_Name = "silk_resampler_c"
' libopus 1.5.2 (ref/opus/silk): resampler.c, with resampler_private.h's and resampler_rom.h's constants (the
' tables are in silk_tables_c). (VB) out and in are arrays and offsets; S->Coefs is an offset into
' silk_rom_i16.
Option Explicit

'/*
' * Matrix of resampling methods used:
' *                                 Fs_out (kHz)
' *                        8      12     16     24     48
' *
' *               8        C      UF     U      UF     UF
' *              12        AF     C      UF     U      UF
' * Fs_in (kHz)  16        D      AF     C      UF     UF
' *              24        AF     D      AF     C      U
' *              48        AF     AF     AF     D      C
' *
' * C   -> Copy (no resampling)
' * D   -> Allpass-based 2x downsampling
' * U   -> Allpass-based 2x upsampling
' * UF  -> Allpass-based 2x upsampling followed by FIR interpolation
' * AF  -> AR2 filter followed by FIR interpolation
' */

' resampler_private.h
'/* Number of input samples to process in the inner loop */
Public Const RESAMPLER_MAX_BATCH_SIZE_MS As Long = 10
Public Const RESAMPLER_MAX_FS_KHZ As Long = 48
Public Const RESAMPLER_MAX_BATCH_SIZE_IN As Long = 480  ' ( RESAMPLER_MAX_BATCH_SIZE_MS * RESAMPLER_MAX_FS_KHZ )

' resampler_rom.h
Public Const RESAMPLER_DOWN_ORDER_FIR0 As Long = 18
Public Const RESAMPLER_DOWN_ORDER_FIR1 As Long = 24
Public Const RESAMPLER_DOWN_ORDER_FIR2 As Long = 36
Public Const RESAMPLER_ORDER_FIR_12 As Long = 8

'/* Tables with delay compensation values to equalize total delay for different modes */
' static const opus_int8 delay_matrix_enc[ 5 ][ 3 ] = {
' /* in  \ out  8  12  16 */
' /*  8 */   {  6,  0,  3 },
' /* 12 */   {  0,  7,  3 },
' /* 16 */   {  0,  1, 10 },
' /* 24 */   {  0,  2,  6 },
' /* 48 */   { 18, 10, 12 }
' };
'
' static const opus_int8 delay_matrix_dec[ 3 ][ 5 ] = {
' /* in  \ out  8  12  16  24  48 */
' /*  8 */   {  4,  0,  2,  0,  0 },
' /* 12 */   {  0,  9,  4,  7,  4 },
' /* 16 */   {  0,  3, 12,  7,  7 }
' };

Private Const USE_silk_resampler_copy As Long = 0
Private Const USE_silk_resampler_private_up2_HQ_wrapper As Long = 1
Private Const USE_silk_resampler_private_IIR_FIR As Long = 2
Private Const USE_silk_resampler_private_down_FIR As Long = 3

Private Function delay_matrix_enc(ByVal i As Long, ByVal j As Long) As Long
    Static m(0 To 14) As Long, ready As Boolean
    If Not ready Then
        m(0) = 6: m(1) = 0: m(2) = 3
        m(3) = 0: m(4) = 7: m(5) = 3
        m(6) = 0: m(7) = 1: m(8) = 10
        m(9) = 0: m(10) = 2: m(11) = 6
        m(12) = 18: m(13) = 10: m(14) = 12
        ready = True
    End If
    delay_matrix_enc = m(i * 3 + j)
End Function

Private Function delay_matrix_dec(ByVal i As Long, ByVal j As Long) As Long
    Static m(0 To 14) As Long, ready As Boolean
    If Not ready Then
        m(0) = 4: m(1) = 0: m(2) = 2: m(3) = 0: m(4) = 0
        m(5) = 0: m(6) = 9: m(7) = 4: m(8) = 7: m(9) = 4
        m(10) = 0: m(11) = 3: m(12) = 12: m(13) = 7: m(14) = 7
        ready = True
    End If
    delay_matrix_dec = m(i * 5 + j)
End Function

'/* Simple way to make [8000, 12000, 16000, 24000, 48000] to [0, 1, 2, 3, 4] */
' #define rateID(R) ( ( ( ((R)>>12) - ((R)>16000) ) >> ((R)>24000) ) - 1 )
Private Function rateID(ByVal R As Long) As Long
    Dim t As Long
    t = silk_RSHIFT(R, 12)
    If R > 16000 Then t = t - 1
    If R > 24000 Then t = silk_RSHIFT(t, 1)
    rateID = t - 1
End Function

'/* Initialize/reset the resampler state for a given pair of input/output sampling rates */
Public Function silk_resampler_init(S As silk_resampler_state_struct, ByVal Fs_Hz_in As Long, ByVal Fs_Hz_out As Long, _
        ByVal forEnc As Long) As Long
    Static blank As silk_resampler_state_struct
    Dim up2x As Long

    '/* Clear state */
    S = blank

    '/* Input checking */
    If forEnc <> 0 Then
        If (Fs_Hz_in <> 8000 And Fs_Hz_in <> 12000 And Fs_Hz_in <> 16000 And Fs_Hz_in <> 24000 And Fs_Hz_in <> 48000) Or _
            (Fs_Hz_out <> 8000 And Fs_Hz_out <> 12000 And Fs_Hz_out <> 16000) Then
            ' celt_assert( 0 );
            silk_resampler_init = -1
            Exit Function
        End If
        S.inputDelay = delay_matrix_enc(rateID(Fs_Hz_in), rateID(Fs_Hz_out))
    Else
        If (Fs_Hz_in <> 8000 And Fs_Hz_in <> 12000 And Fs_Hz_in <> 16000) Or _
            (Fs_Hz_out <> 8000 And Fs_Hz_out <> 12000 And Fs_Hz_out <> 16000 And Fs_Hz_out <> 24000 And Fs_Hz_out <> 48000) Then
            ' celt_assert( 0 );
            silk_resampler_init = -1
            Exit Function
        End If
        S.inputDelay = delay_matrix_dec(rateID(Fs_Hz_in), rateID(Fs_Hz_out))
    End If

    S.Fs_in_kHz = Fs_Hz_in \ 1000
    S.Fs_out_kHz = Fs_Hz_out \ 1000

    '/* Number of samples processed per batch */
    S.batchSize = S.Fs_in_kHz * RESAMPLER_MAX_BATCH_SIZE_MS

    '/* Find resampler with the right sampling ratio */
    up2x = 0
    If Fs_Hz_out > Fs_Hz_in Then
        '/* Upsample */
        If Fs_Hz_out = Fs_Hz_in * 2 Then                            '/* Fs_out : Fs_in = 2 : 1 */
            '/* Special case: directly use 2x upsampler */
            S.resampler_function = USE_silk_resampler_private_up2_HQ_wrapper
        Else
            '/* Default resampler */
            S.resampler_function = USE_silk_resampler_private_IIR_FIR
            up2x = 1
        End If
    ElseIf Fs_Hz_out < Fs_Hz_in Then
        '/* Downsample */
         S.resampler_function = USE_silk_resampler_private_down_FIR
        If Fs_Hz_out * 4 = Fs_Hz_in * 3 Then                        '/* Fs_out : Fs_in = 3 : 4 */
            S.FIR_Fracs = 3
            S.FIR_Order = RESAMPLER_DOWN_ORDER_FIR0
            S.Coefs = silk_Resampler_3_4_COEFS
        ElseIf Fs_Hz_out * 3 = Fs_Hz_in * 2 Then                    '/* Fs_out : Fs_in = 2 : 3 */
            S.FIR_Fracs = 2
            S.FIR_Order = RESAMPLER_DOWN_ORDER_FIR0
            S.Coefs = silk_Resampler_2_3_COEFS
        ElseIf Fs_Hz_out * 2 = Fs_Hz_in Then                        '/* Fs_out : Fs_in = 1 : 2 */
            S.FIR_Fracs = 1
            S.FIR_Order = RESAMPLER_DOWN_ORDER_FIR1
            S.Coefs = silk_Resampler_1_2_COEFS
        ElseIf Fs_Hz_out * 3 = Fs_Hz_in Then                        '/* Fs_out : Fs_in = 1 : 3 */
            S.FIR_Fracs = 1
            S.FIR_Order = RESAMPLER_DOWN_ORDER_FIR2
            S.Coefs = silk_Resampler_1_3_COEFS
        ElseIf Fs_Hz_out * 4 = Fs_Hz_in Then                        '/* Fs_out : Fs_in = 1 : 4 */
            S.FIR_Fracs = 1
            S.FIR_Order = RESAMPLER_DOWN_ORDER_FIR2
            S.Coefs = silk_Resampler_1_4_COEFS
        ElseIf Fs_Hz_out * 6 = Fs_Hz_in Then                        '/* Fs_out : Fs_in = 1 : 6 */
            S.FIR_Fracs = 1
            S.FIR_Order = RESAMPLER_DOWN_ORDER_FIR2
            S.Coefs = silk_Resampler_1_6_COEFS
        Else
            '/* None available */
            ' celt_assert( 0 );
            silk_resampler_init = -1
            Exit Function
        End If
    Else
        '/* Input and output sampling rates are equal: copy */
        S.resampler_function = USE_silk_resampler_copy
    End If

    '/* Ratio of input/output samples */
    S.invRatio_Q16 = silk_LSHIFT32(silk_LSHIFT32(Fs_Hz_in, 14 + up2x) \ Fs_Hz_out, 2)
    '/* Make sure the ratio is rounded up */
    Do While silk_SMULWW(S.invRatio_Q16, Fs_Hz_out) < silk_LSHIFT32(Fs_Hz_in, up2x)
        S.invRatio_Q16 = S.invRatio_Q16 + 1
    Loop

    silk_resampler_init = 0
End Function

'/* Resampler: convert from one sampling rate to another */
'/* Input and output sampling rate are at most 48000 Hz  */
Public Function silk_resampler(S As silk_resampler_state_struct, out_() As Integer, ByVal outOffset As Long, _
        in_() As Integer, ByVal inOffset As Long, ByVal inLen As Long) As Long
    Dim nSamples As Long

    '/* Need at least 1 ms of input data */
    ' celt_assert( inLen >= S->Fs_in_kHz );
    '/* Delay can't exceed the 1 ms of buffering */
    ' celt_assert( S->inputDelay <= S->Fs_in_kHz );

    nSamples = S.Fs_in_kHz - S.inputDelay

    '/* Copy to delay buffer */
    silk_memcpy_i16 S.delayBuf, S.inputDelay, in_, inOffset, nSamples

    Select Case S.resampler_function
        Case USE_silk_resampler_private_up2_HQ_wrapper
            silk_resampler_private_up2_HQ_wrapper S, out_, outOffset, S.delayBuf, 0, S.Fs_in_kHz
            silk_resampler_private_up2_HQ_wrapper S, out_, outOffset + S.Fs_out_kHz, in_, inOffset + nSamples, inLen - S.Fs_in_kHz
        Case USE_silk_resampler_private_IIR_FIR
            silk_resampler_private_IIR_FIR S, out_, outOffset, S.delayBuf, 0, S.Fs_in_kHz
            silk_resampler_private_IIR_FIR S, out_, outOffset + S.Fs_out_kHz, in_, inOffset + nSamples, inLen - S.Fs_in_kHz
        Case USE_silk_resampler_private_down_FIR
            silk_resampler_private_down_FIR S, out_, outOffset, S.delayBuf, 0, S.Fs_in_kHz
            silk_resampler_private_down_FIR S, out_, outOffset + S.Fs_out_kHz, in_, inOffset + nSamples, inLen - S.Fs_in_kHz
        Case Else
            silk_memcpy_i16 out_, outOffset, S.delayBuf, 0, S.Fs_in_kHz
            silk_memcpy_i16 out_, outOffset + S.Fs_out_kHz, in_, inOffset + nSamples, inLen - S.Fs_in_kHz
    End Select

    '/* Copy to delay buffer */
    silk_memcpy_i16 S.delayBuf, 0, in_, inOffset + inLen - S.inputDelay, S.inputDelay
    silk_resampler = 0
End Function
