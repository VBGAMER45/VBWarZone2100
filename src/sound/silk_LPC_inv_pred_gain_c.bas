Attribute VB_Name = "silk_LPC_inv_pred_gain_c"
' libopus 1.5.2 (ref/opus/silk): LPC_inv_pred_gain.c (silk_LPC_inverse_pred_gain is SigProc_FIX.h's macro
' for silk_LPC_inverse_pred_gain_c: the C path). (VB) A_Q12 is an array and an offset; the 64-bit products are
' silk_SMULL_RSHIFT_ROUND64's exact Doubles, compared with the int32 range as the C's opus_int64 tmp64.
Option Explicit

Private Const QA As Long = 24
Private Const A_LIMIT As Long = 16773022                ' SILK_FIX_CONST( 0.99975, QA )

' #define MUL32_FRAC_Q(a32, b32, Q)   ((opus_int32)(silk_RSHIFT_ROUND64(silk_SMULL(a32, b32), Q)))
Private Function MUL32_FRAC_Q(ByVal a32 As Long, ByVal b32 As Long, ByVal Q As Long) As Long
    MUL32_FRAC_Q = silk_I32(silk_SMULL_RSHIFT_ROUND64(a32, b32, Q))
End Function

'/* Compute inverse of LPC prediction gain, and                          */
'/* test if LPC coefficients are stable (all poles within unit circle)   */
Private Function LPC_inverse_pred_gain_QA_c(A_QA() As Long, ByVal order As Long) As Long '/* O   Returns inverse prediction gain in energy domain, Q30    */
    Dim k As Long, n As Long, mult2Q As Long
    Dim invGain_Q30 As Long, rc_Q31 As Long, rc_mult1_Q30 As Long, rc_mult2 As Long, tmp1 As Long, tmp2 As Long
    Dim tmp64 As Double

    invGain_Q30 = &H40000000                            ' SILK_FIX_CONST( 1, 30 )
    For k = order - 1 To 1 Step -1
        '/* Check for stability */
        If (A_QA(k) > A_LIMIT) Or (A_QA(k) < -A_LIMIT) Then
            Exit Function
        End If

        '/* Set RC equal to negated AR coef */
        rc_Q31 = -silk_LSHIFT(A_QA(k), 31 - QA)

        '/* rc_mult1_Q30 range: [ 1 : 2^30 ] */
        rc_mult1_Q30 = &H40000000 - silk_SMMUL(rc_Q31, rc_Q31)
        ' silk_assert( rc_mult1_Q30 > ( 1 << 15 ) );                   /* reduce A_LIMIT if fails */
        ' silk_assert( rc_mult1_Q30 <= ( 1 << 30 ) );

        '/* Update inverse gain */
        '/* invGain_Q30 range: [ 0 : 2^30 ] */
        invGain_Q30 = silk_SMMUL(invGain_Q30, rc_mult1_Q30) * 4
        ' silk_assert( invGain_Q30 >= 0           );
        ' silk_assert( invGain_Q30 <= ( 1 << 30 ) );
        If invGain_Q30 < 107374 Then                    ' SILK_FIX_CONST( 1.0f / MAX_PREDICTION_POWER_GAIN, 30 )
            Exit Function
        End If

        '/* rc_mult2 range: [ 2^30 : silk_int32_MAX ] */
        mult2Q = 32 - silk_CLZ32(silk_abs(rc_mult1_Q30))
        rc_mult2 = silk_INVERSE32_varQ(rc_mult1_Q30, mult2Q + 30)

        '/* Update AR coefficient */
        For n = 0 To silk_RSHIFT(k + 1, 1) - 1
            tmp1 = A_QA(n)
            tmp2 = A_QA(k - n - 1)
            tmp64 = silk_SMULL_RSHIFT_ROUND64(silk_SUB_SAT32(tmp1, MUL32_FRAC_Q(tmp2, rc_Q31, 31)), rc_mult2, mult2Q)
            If tmp64 > 2147483647# Or tmp64 < -2147483648# Then
               Exit Function
            End If
            A_QA(n) = tmp64
            tmp64 = silk_SMULL_RSHIFT_ROUND64(silk_SUB_SAT32(tmp2, MUL32_FRAC_Q(tmp1, rc_Q31, 31)), rc_mult2, mult2Q)
            If tmp64 > 2147483647# Or tmp64 < -2147483648# Then
               Exit Function
            End If
            A_QA(k - n - 1) = tmp64
        Next
    Next

    '/* Check for stability */
    If (A_QA(k) > A_LIMIT) Or (A_QA(k) < -A_LIMIT) Then
        Exit Function
    End If

    '/* Set RC equal to negated AR coef */
    rc_Q31 = -silk_LSHIFT(A_QA(0), 31 - QA)

    '/* Range: [ 1 : 2^30 ] */
    rc_mult1_Q30 = &H40000000 - silk_SMMUL(rc_Q31, rc_Q31)

    '/* Update inverse gain */
    '/* Range: [ 0 : 2^30 ] */
    invGain_Q30 = silk_SMMUL(invGain_Q30, rc_mult1_Q30) * 4
    ' silk_assert( invGain_Q30 >= 0           );
    ' silk_assert( invGain_Q30 <= ( 1 << 30 ) );
    If invGain_Q30 < 107374 Then                        ' SILK_FIX_CONST( 1.0f / MAX_PREDICTION_POWER_GAIN, 30 )
        Exit Function
    End If

    LPC_inverse_pred_gain_QA_c = invGain_Q30
End Function

'/* For input in Q12 domain */
Public Function silk_LPC_inverse_pred_gain(A_Q12() As Integer, ByVal A_Q12Offset As Long, ByVal order As Long, _
        ByVal arch As Long) As Long                     '/* O   Returns inverse prediction gain in energy domain, Q30        */
    Dim k As Long
    Dim Atmp_QA(0 To 23) As Long                        ' [ SILK_MAX_ORDER_LPC ]
    Dim DC_resp As Long

    '/* Increase Q domain of the AR coefficients */
    For k = 0 To order - 1
        DC_resp = DC_resp + A_Q12(A_Q12Offset + k)
        Atmp_QA(k) = CLng(A_Q12(A_Q12Offset + k)) * 4096    ' silk_LSHIFT32( (opus_int32)A_Q12[ k ], QA - 12 )
    Next
    '/* If the DC is unstable, we don't even need to do the full calculations */
    If DC_resp >= 4096 Then
        Exit Function
    End If
    silk_LPC_inverse_pred_gain = LPC_inverse_pred_gain_QA_c(Atmp_QA, order)
End Function
