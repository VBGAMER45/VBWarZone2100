Attribute VB_Name = "silk_stereo_MS_to_LR_c"
' libopus 1.5.2 (ref/opus/silk): stereo_MS_to_LR.c. (VB) x1 and x2 are arrays and offsets.
Option Explicit

'/* Convert adaptive Mid/Side representation to Left/Right stereo signal */
Public Sub silk_stereo_MS_to_LR(state As stereo_dec_state, x1() As Integer, ByVal x1Offset As Long, _
        x2() As Integer, ByVal x2Offset As Long, pred_Q13() As Long, ByVal fs_kHz As Long, ByVal frame_length As Long)
    Dim n As Long, denom_Q16 As Long, delta0_Q13 As Long, delta1_Q13 As Long
    Dim sum As Long, diff As Long, pred0_Q13 As Long, pred1_Q13 As Long

    '/* Buffering */
    x1(x1Offset) = state.sMid(0): x1(x1Offset + 1) = state.sMid(1)
    x2(x2Offset) = state.sSide(0): x2(x2Offset + 1) = state.sSide(1)
    state.sMid(0) = x1(x1Offset + frame_length): state.sMid(1) = x1(x1Offset + frame_length + 1)
    state.sSide(0) = x2(x2Offset + frame_length): state.sSide(1) = x2(x2Offset + frame_length + 1)

    '/* Interpolate predictors and add prediction to side channel */
    pred0_Q13 = state.pred_prev_Q13(0)
    pred1_Q13 = state.pred_prev_Q13(1)
    denom_Q16 = &H10000 \ (STEREO_INTERP_LEN_MS * fs_kHz)
    delta0_Q13 = silk_RSHIFT_ROUND(silk_SMULBB(pred_Q13(0) - state.pred_prev_Q13(0), denom_Q16), 16)
    delta1_Q13 = silk_RSHIFT_ROUND(silk_SMULBB(pred_Q13(1) - state.pred_prev_Q13(1), denom_Q16), 16)
    For n = 0 To STEREO_INTERP_LEN_MS * fs_kHz - 1
        pred0_Q13 = pred0_Q13 + delta0_Q13
        pred1_Q13 = pred1_Q13 + delta1_Q13
        sum = silk_LSHIFT(silk_ADD_LSHIFT32(CLng(x1(x1Offset + n)) + x1(x1Offset + n + 2), x1(x1Offset + n + 1), 1), 9)    '/* Q11 */
        sum = silk_SMLAWB(CLng(x2(x2Offset + n + 1)) * 256, sum, pred0_Q13)         '/* Q8  */
        sum = silk_SMLAWB(sum, CLng(x1(x1Offset + n + 1)) * 2048, pred1_Q13)        '/* Q8  */
        x2(x2Offset + n + 1) = silk_SAT16(silk_RSHIFT_ROUND(sum, 8))
    Next
    pred0_Q13 = pred_Q13(0)
    pred1_Q13 = pred_Q13(1)
    For n = STEREO_INTERP_LEN_MS * fs_kHz To frame_length - 1
        sum = silk_LSHIFT(silk_ADD_LSHIFT32(CLng(x1(x1Offset + n)) + x1(x1Offset + n + 2), x1(x1Offset + n + 1), 1), 9)    '/* Q11 */
        sum = silk_SMLAWB(CLng(x2(x2Offset + n + 1)) * 256, sum, pred0_Q13)         '/* Q8  */
        sum = silk_SMLAWB(sum, CLng(x1(x1Offset + n + 1)) * 2048, pred1_Q13)        '/* Q8  */
        x2(x2Offset + n + 1) = silk_SAT16(silk_RSHIFT_ROUND(sum, 8))
    Next
    state.pred_prev_Q13(0) = pred_Q13(0)
    state.pred_prev_Q13(1) = pred_Q13(1)

    '/* Convert to left/right signals */
    For n = 0 To frame_length - 1
        sum = CLng(x1(x1Offset + n + 1)) + x2(x2Offset + n + 1)
        diff = CLng(x1(x1Offset + n + 1)) - x2(x2Offset + n + 1)
        x1(x1Offset + n + 1) = silk_SAT16(sum)
        x2(x2Offset + n + 1) = silk_SAT16(diff)
    Next
End Sub
