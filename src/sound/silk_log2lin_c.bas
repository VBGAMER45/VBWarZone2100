Attribute VB_Name = "silk_log2lin_c"
' libopus 1.5.2 (ref/opus/silk): log2lin.c.
Option Explicit

'/* Approximation of 2^() (very close inverse of silk_lin2log()) */
'/* Convert input to a linear scale    */
Public Function silk_log2lin(ByVal inLog_Q7 As Long) As Long    '/* I  input on log scale                                            */
    Dim out_ As Long, frac_Q7 As Long

    If inLog_Q7 < 0 Then
        silk_log2lin = 0
        Exit Function
    ElseIf inLog_Q7 >= 3967 Then
        silk_log2lin = silk_int32_MAX
        Exit Function
    End If

    out_ = silk_LSHIFT(1, silk_RSHIFT(inLog_Q7, 7))
    frac_Q7 = inLog_Q7 And &H7F&
    If inLog_Q7 < 2048 Then
        '/* Piece-wise parabolic approximation */
        out_ = out_ + silk_RSHIFT(out_ * silk_SMLAWB(frac_Q7, silk_SMULBB(frac_Q7, 128 - frac_Q7), -174), 7)
    Else
        '/* Piece-wise parabolic approximation */
        out_ = out_ + silk_RSHIFT(out_, 7) * silk_SMLAWB(frac_Q7, silk_SMULBB(frac_Q7, 128 - frac_Q7), -174)
    End If
    silk_log2lin = out_
End Function
