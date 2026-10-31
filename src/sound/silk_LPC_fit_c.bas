Attribute VB_Name = "silk_LPC_fit_c"
' libopus 1.5.2 (ref/opus/silk): LPC_fit.c. (VB) a_QOUT is an array and an offset.
Option Explicit

'/* Convert int32 coefficients to int16 coefs and make sure there's no wrap-around.
'   This logic is reused in _celt_lpc(). Any bug fixes should also be applied there. */
Public Sub silk_LPC_fit(a_QOUT() As Integer, ByVal a_QOUTOffset As Long, a_QIN() As Long, ByVal QOUT As Long, _
        ByVal QIN As Long, ByVal d As Long)
    Dim i As Long, k As Long, idx As Long
    Dim maxabs As Long, absval As Long, chirp_Q16 As Long

    '/* Limit the maximum absolute value of the prediction coefficients, so that they'll fit in int16 */
    For i = 0 To 9
        '/* Find maximum absolute value and its index */
        maxabs = 0
        For k = 0 To d - 1
            absval = silk_abs(a_QIN(k))
            If absval > maxabs Then
                maxabs = absval
                idx = k
            End If
        Next
        maxabs = silk_RSHIFT_ROUND(maxabs, QIN - QOUT)

        If maxabs > silk_int16_MAX Then
            '/* Reduce magnitude of prediction coefficients */
            maxabs = silk_min(maxabs, 163838)           '/* ( silk_int32_MAX >> 14 ) + silk_int16_MAX = 163838 */
            chirp_Q16 = 65470 - silk_LSHIFT(maxabs - silk_int16_MAX, 14) \ silk_RSHIFT(maxabs * (idx + 1), 2)   ' SILK_FIX_CONST( 0.999, 16 ) - silk_DIV32( ... )
            silk_bwexpander_32 a_QIN, d, chirp_Q16
        Else
            Exit For
        End If
    Next

    If i = 10 Then
        '/* Reached the last iteration, clip the coefficients */
        For k = 0 To d - 1
            a_QOUT(a_QOUTOffset + k) = silk_SAT16(silk_RSHIFT_ROUND(a_QIN(k), QIN - QOUT))
            a_QIN(k) = silk_LSHIFT(a_QOUT(a_QOUTOffset + k), QIN - QOUT)
        Next
    Else
        For k = 0 To d - 1
            a_QOUT(a_QOUTOffset + k) = silk_RSHIFT_ROUND(a_QIN(k), QIN - QOUT)
        Next
    End If
End Sub
