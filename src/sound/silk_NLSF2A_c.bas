Attribute VB_Name = "silk_NLSF2A_c"
' libopus 1.5.2 (ref/opus/silk): NLSF2A.c. (VB) a_Q12 is an array and an offset; the 64-bit products of
' silk_NLSF2A_find_poly are silk_SMULL_RSHIFT_ROUND64's exact Doubles.
Option Explicit

'/* conversion between prediction filter coefficients and LSFs   */
'/* order should be even                                         */
'/* a piecewise linear approximation maps LSF <-> cos(LSF)       */
'/* therefore the result is not accurate LSFs, but the two       */
'/* functions are accurate inverses of each other                */

Private Const QA As Long = 16

'/* helper function for NLSF2A(..) */
Private Sub silk_NLSF2A_find_poly(out_() As Long, cLSF() As Long, ByVal cLSFOffset As Long, ByVal dd As Long)
    Dim k As Long, n As Long
    Dim ftmp As Long

    out_(0) = &H10000                                   ' silk_LSHIFT( 1, QA )
    out_(1) = -cLSF(cLSFOffset)
    For k = 1 To dd - 1
        ftmp = cLSF(cLSFOffset + 2 * k)                 '/* QA*/
        out_(k + 1) = out_(k - 1) * 2 - silk_I32(silk_SMULL_RSHIFT_ROUND64(ftmp, out_(k), QA))
        For n = k To 2 Step -1
            out_(n) = out_(n) + (out_(n - 2) - silk_I32(silk_SMULL_RSHIFT_ROUND64(ftmp, out_(n - 1), QA)))
        Next
        out_(1) = out_(1) - ftmp
    Next
End Sub

'/* compute whitening filter coefficients from normalized line spectral frequencies */
Public Sub silk_NLSF2A(a_Q12() As Integer, ByVal a_Q12Offset As Long, NLSF() As Integer, ByVal d As Long, ByVal arch As Long)
    '/* This ordering was found to maximize quality. It improves numerical accuracy of
    '   silk_NLSF2A_find_poly() compared to "standard" ordering. */
    ' static const unsigned char ordering16[16] = { 0, 15, 8, 7, 4, 11, 12, 3, 2, 13, 10, 5, 6, 9, 14, 1 };
    ' static const unsigned char ordering10[10] = { 0, 9, 6, 3, 4, 5, 8, 1, 2, 7 };
    Static ordering16(0 To 15) As Long, ordering10(0 To 9) As Long, orderingReady As Boolean
    Dim k As Long, i As Long, dd As Long
    Dim cos_LSF_QA(0 To 23) As Long                     ' [ SILK_MAX_ORDER_LPC ]
    Dim P(0 To 12) As Long, Q(0 To 12) As Long          ' [ SILK_MAX_ORDER_LPC / 2 + 1 ]
    Dim Ptmp As Long, Qtmp As Long, f_int As Long, f_frac As Long, cos_val As Long, delta As Long
    Dim a32_QA1(0 To 23) As Long                        ' [ SILK_MAX_ORDER_LPC ]
    Dim ord As Long

    If Not orderingReady Then
        ordering16(0) = 0: ordering16(1) = 15: ordering16(2) = 8: ordering16(3) = 7
        ordering16(4) = 4: ordering16(5) = 11: ordering16(6) = 12: ordering16(7) = 3
        ordering16(8) = 2: ordering16(9) = 13: ordering16(10) = 10: ordering16(11) = 5
        ordering16(12) = 6: ordering16(13) = 9: ordering16(14) = 14: ordering16(15) = 1
        ordering10(0) = 0: ordering10(1) = 9: ordering10(2) = 6: ordering10(3) = 3: ordering10(4) = 4
        ordering10(5) = 5: ordering10(6) = 8: ordering10(7) = 1: ordering10(8) = 2: ordering10(9) = 7
        orderingReady = True
    End If

    ' silk_assert( LSF_COS_TAB_SZ_FIX == 128 );
    ' celt_assert( d==10 || d==16 );

    '/* convert LSFs to 2*cos(LSF), using piecewise linear curve from table */
    For k = 0 To d - 1
        ' silk_assert( NLSF[k] >= 0 );

        '/* f_int on a scale 0-127 (rounded down) */
        f_int = silk_RSHIFT(NLSF(k), 15 - 7)

        '/* f_frac, range: 0..255 */
        f_frac = NLSF(k) - f_int * 256                  ' NLSF[k] - silk_LSHIFT( f_int, 15 - 7 )

        ' silk_assert(f_int >= 0);
        ' silk_assert(f_int < LSF_COS_TAB_SZ_FIX );

        '/* Read start and end value from table */
        cos_val = silk_rom_i16(silk_LSFCosTab_FIX_Q12 + f_int)                     '/* Q12 */
        delta = silk_rom_i16(silk_LSFCosTab_FIX_Q12 + f_int + 1) - cos_val         '/* Q12, with a range of 0..200 */

        '/* Linear interpolation */
        If d = 16 Then ord = ordering16(k) Else ord = ordering10(k)
        cos_LSF_QA(ord) = silk_RSHIFT_ROUND(cos_val * 256 + delta * f_frac, 20 - QA)    '/* QA */
    Next

    dd = silk_RSHIFT(d, 1)

    '/* generate even and odd polynomials using convolution */
    silk_NLSF2A_find_poly P, cos_LSF_QA, 0, dd
    silk_NLSF2A_find_poly Q, cos_LSF_QA, 1, dd

    '/* convert even and odd polynomials to opus_int32 Q12 filter coefs */
    For k = 0 To dd - 1
        Ptmp = P(k + 1) + P(k)
        Qtmp = Q(k + 1) - Q(k)

        '/* the Ptmp and Qtmp values at this stage need to fit in int32 */
        a32_QA1(k) = -Qtmp - Ptmp                       '/* QA+1 */
        a32_QA1(d - k - 1) = Qtmp - Ptmp                '/* QA+1 */
    Next

    '/* Convert int32 coefficients to Q12 int16 coefs */
    silk_LPC_fit a_Q12, a_Q12Offset, a32_QA1, 12, QA + 1, d

    i = 0
    Do While i < MAX_LPC_STABILIZE_ITERATIONS
        If silk_LPC_inverse_pred_gain(a_Q12, a_Q12Offset, d, arch) <> 0 Then Exit Do
        '/* Prediction coefficients are (too close to) unstable; apply bandwidth expansion   */
        '/* on the unscaled coefficients, convert to Q12 and measure again                   */
        silk_bwexpander_32 a32_QA1, d, 65536 - silk_LSHIFT(2, i)
        For k = 0 To d - 1
            a_Q12(a_Q12Offset + k) = silk_RSHIFT_ROUND(a32_QA1(k), QA + 1 - 12)       '/* QA+1 -> Q12 */
        Next
        i = i + 1
    Loop
End Sub
