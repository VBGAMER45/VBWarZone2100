Attribute VB_Name = "silk_stereo_decode_pred_c"
' libopus 1.5.2 (ref/opus/silk): stereo_decode_pred.c.
Option Explicit

'/* Decode mid/side predictors */
Public Sub silk_stereo_decode_pred(psRangeDec As ec_ctx, pred_Q13() As Long)
    Dim n As Long
    Dim ix(0 To 1, 0 To 2) As Long
    Dim low_Q13 As Long, step_Q13 As Long

    '/* Entropy decoding */
    n = ec_dec_icdf(psRangeDec, silk_rom_u8, silk_stereo_pred_joint_iCDF, 8)
    ix(0, 2) = n \ 5
    ix(1, 2) = n - 5 * ix(0, 2)
    For n = 0 To 1
        ix(n, 0) = ec_dec_icdf(psRangeDec, silk_rom_u8, silk_uniform3_iCDF, 8)
        ix(n, 1) = ec_dec_icdf(psRangeDec, silk_rom_u8, silk_uniform5_iCDF, 8)
    Next

    '/* Dequantize */
    For n = 0 To 1
        ix(n, 0) = ix(n, 0) + 3 * ix(n, 2)
        low_Q13 = silk_rom_i16(silk_stereo_pred_quant_Q13 + ix(n, 0))
        step_Q13 = silk_SMULWB(silk_rom_i16(silk_stereo_pred_quant_Q13 + ix(n, 0) + 1) - low_Q13, _
            6554)                                       ' SILK_FIX_CONST( 0.5 / STEREO_QUANT_SUB_STEPS, 16 )
        pred_Q13(n) = silk_SMLABB(low_Q13, step_Q13, 2 * ix(n, 1) + 1)
    Next

    '/* Subtract second from first predictor (helps when actually applying these) */
    pred_Q13(0) = pred_Q13(0) - pred_Q13(1)
End Sub

'/* Decode mid-only flag */
Public Sub silk_stereo_decode_mid_only(psRangeDec As ec_ctx, decode_only_mid As Long)
    '/* Decode flag that only mid channel is coded */
    decode_only_mid = ec_dec_icdf(psRangeDec, silk_rom_u8, silk_stereo_only_code_mid_iCDF, 8)
End Sub
