Attribute VB_Name = "silk_NLSF_decode_c"
' libopus 1.5.2 (ref/opus/silk): NLSF_decode.c. (VB) psNLSF_CB is an index of silk_NLSF_CBs; the
' codebook's tables are offsets into the ROMs (silk_tables_c).
Option Explicit

'/* Predictive dequantizer for NLSF residuals */
Private Sub silk_NLSF_residual_dequant(x_Q10() As Integer, indices() As Long, ByVal indicesOffset As Long, _
        pred_coef_Q8() As Byte, ByVal quant_step_size_Q16 As Long, ByVal order As Long)
    Dim i As Long, out_Q10 As Long, pred_Q10 As Long

    out_Q10 = 0
    For i = order - 1 To 0 Step -1
        pred_Q10 = silk_RSHIFT(silk_SMULBB(out_Q10, pred_coef_Q8(i)), 8)
        out_Q10 = indices(indicesOffset + i) * 1024     ' silk_LSHIFT( indices[ i ], 10 )
        If out_Q10 > 0 Then
            out_Q10 = out_Q10 - 102                     ' silk_SUB16( out_Q10, SILK_FIX_CONST( NLSF_QUANT_LEVEL_ADJ, 10 ) )
        ElseIf out_Q10 < 0 Then
            out_Q10 = out_Q10 + 102                     ' silk_ADD16( out_Q10, SILK_FIX_CONST( NLSF_QUANT_LEVEL_ADJ, 10 ) )
        End If
        out_Q10 = silk_SMLAWB(pred_Q10, out_Q10, quant_step_size_Q16)
        x_Q10(i) = out_Q10
    Next
End Sub

'/***********************/
'/* NLSF vector decoder */
'/***********************/
Public Sub silk_NLSF_decode(pNLSF_Q15() As Integer, NLSFIndices() As Long, ByVal psNLSF_CB As Long)
    Dim i As Long
    Dim pred_Q8(0 To 15) As Byte                        ' [ MAX_LPC_ORDER ]
    Dim ec_ix(0 To 15) As Integer                       ' [ MAX_LPC_ORDER ]
    Dim res_Q10(0 To 15) As Integer                     ' [ MAX_LPC_ORDER ]
    Dim NLSF_Q15_tmp As Long
    Dim pCB_element As Long                             ' const opus_uint8 *: an offset into silk_rom_u8
    Dim pCB_Wght_Q9 As Long                             ' const opus_int16 *: an offset into silk_rom_i16

    '/* Unpack entropy table indices and predictor for current CB1 index */
    silk_NLSF_unpack ec_ix, pred_Q8, psNLSF_CB, NLSFIndices(0)

    '/* Predictive residual dequantizer */
    silk_NLSF_residual_dequant res_Q10, NLSFIndices, 1, pred_Q8, silk_NLSF_CBs(psNLSF_CB).quantStepSize_Q16, silk_NLSF_CBs(psNLSF_CB).order

    '/* Apply inverse square-rooted weights to first stage and add to output */
    pCB_element = silk_NLSF_CBs(psNLSF_CB).CB1_NLSF_Q8 + NLSFIndices(0) * silk_NLSF_CBs(psNLSF_CB).order
    pCB_Wght_Q9 = silk_NLSF_CBs(psNLSF_CB).CB1_Wght_Q9 + NLSFIndices(0) * silk_NLSF_CBs(psNLSF_CB).order
    For i = 0 To silk_NLSF_CBs(psNLSF_CB).order - 1
        NLSF_Q15_tmp = silk_ADD_LSHIFT32((CLng(res_Q10(i)) * 16384) \ silk_rom_i16(pCB_Wght_Q9 + i), silk_rom_u8(pCB_element + i), 7)
        pNLSF_Q15(i) = silk_LIMIT(NLSF_Q15_tmp, 0, 32767)
    Next

    '/* NLSF stabilization */
    silk_NLSF_stabilize pNLSF_Q15, silk_NLSF_CBs(psNLSF_CB).deltaMin_Q15, silk_NLSF_CBs(psNLSF_CB).order
End Sub
