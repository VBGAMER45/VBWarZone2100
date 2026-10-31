Attribute VB_Name = "silk_resampler_private_up2_HQ_c"
' libopus 1.5.2 (ref/opus/silk): resampler_private_up2_HQ.c. (VB) out and in are arrays and offsets; the
' silk_SMULWB / silk_SMLAWB of the all-pass sections are written out (silk_SigProc_FIX_h's formula), with
' resampler_rom.h's silk_resampler_up2_hq_0 = { 1746, 14986, 39083 - 65536 } and
' silk_resampler_up2_hq_1 = { 6854, 25769, 55542 - 65536 } as constants.
Option Explicit

Private Const hq0_0 As Long = 1746                      ' silk_resampler_up2_hq_0[ 0 ]
Private Const hq0_1 As Long = 14986                     ' silk_resampler_up2_hq_0[ 1 ]
Private Const hq0_2 As Long = -26453                    ' silk_resampler_up2_hq_0[ 2 ] = 39083 - 65536
Private Const hq1_0 As Long = 6854                      ' silk_resampler_up2_hq_1[ 0 ]
Private Const hq1_1 As Long = 25769                     ' silk_resampler_up2_hq_1[ 1 ]
Private Const hq1_2 As Long = -9994                     ' silk_resampler_up2_hq_1[ 2 ] = 55542 - 65536

'/* Upsample by a factor 2, high quality */
'/* Uses 2nd order allpass filters for the 2x upsampling, followed by a      */
'/* notch filter just above Nyquist.                                         */
Public Sub silk_resampler_private_up2_HQ(S() As Long, out_() As Integer, ByVal outOffset As Long, _
        in_() As Integer, ByVal inOffset As Long, ByVal len_ As Long)
    Dim k As Long
    Dim in32 As Long, out32_1 As Long, out32_2 As Long, Y As Long, X As Long
    Dim S0 As Long, S1 As Long, S2 As Long, S3 As Long, S4 As Long, S5 As Long

    ' silk_assert( silk_resampler_up2_hq_0[ 0 ] > 0 ); ... ( silk_resampler_up2_hq_1[ 2 ] < 0 );
    S0 = S(0): S1 = S(1): S2 = S(2): S3 = S(3): S4 = S(4): S5 = S(5)

    '/* Internal variables and state are in Q10 format */
    For k = 0 To len_ - 1
        '/* Convert to Q10 */
        in32 = CLng(in_(inOffset + k)) * 1024           ' silk_LSHIFT( (opus_int32)in[ k ], 10 )

        '/* First all-pass section for even output sample */
        Y = in32 - S0
        X = ((Y And &HFFFF0000) \ &H10000) * hq0_0 + ((((Y And &HFFFF&) * hq0_0) And &HFFFF0000) \ &H10000)     ' silk_SMULWB( Y, silk_resampler_up2_hq_0[ 0 ] )
        out32_1 = S0 + X
        S0 = in32 + X

        '/* Second all-pass section for even output sample */
        Y = out32_1 - S1
        X = ((Y And &HFFFF0000) \ &H10000) * hq0_1 + ((((Y And &HFFFF&) * hq0_1) And &HFFFF0000) \ &H10000)     ' silk_SMULWB( Y, silk_resampler_up2_hq_0[ 1 ] )
        out32_2 = S1 + X
        S1 = out32_1 + X

        '/* Third all-pass section for even output sample */
        Y = out32_2 - S2
        X = Y + ((Y And &HFFFF0000) \ &H10000) * hq0_2 + ((((Y And &HFFFF&) * hq0_2) And &HFFFF0000) \ &H10000) ' silk_SMLAWB( Y, Y, silk_resampler_up2_hq_0[ 2 ] )
        out32_1 = S2 + X
        S2 = out32_2 + X

        '/* Apply gain in Q15, convert back to int16 and store to output */
        out_(outOffset + 2 * k) = silk_SAT16(silk_RSHIFT_ROUND(out32_1, 10))

        '/* First all-pass section for odd output sample */
        Y = in32 - S3
        X = ((Y And &HFFFF0000) \ &H10000) * hq1_0 + ((((Y And &HFFFF&) * hq1_0) And &HFFFF0000) \ &H10000)     ' silk_SMULWB( Y, silk_resampler_up2_hq_1[ 0 ] )
        out32_1 = S3 + X
        S3 = in32 + X

        '/* Second all-pass section for odd output sample */
        Y = out32_1 - S4
        X = ((Y And &HFFFF0000) \ &H10000) * hq1_1 + ((((Y And &HFFFF&) * hq1_1) And &HFFFF0000) \ &H10000)     ' silk_SMULWB( Y, silk_resampler_up2_hq_1[ 1 ] )
        out32_2 = S4 + X
        S4 = out32_1 + X

        '/* Third all-pass section for odd output sample */
        Y = out32_2 - S5
        X = Y + ((Y And &HFFFF0000) \ &H10000) * hq1_2 + ((((Y And &HFFFF&) * hq1_2) And &HFFFF0000) \ &H10000) ' silk_SMLAWB( Y, Y, silk_resampler_up2_hq_1[ 2 ] )
        out32_1 = S5 + X
        S5 = out32_2 + X

        '/* Apply gain in Q15, convert back to int16 and store to output */
        out_(outOffset + 2 * k + 1) = silk_SAT16(silk_RSHIFT_ROUND(out32_1, 10))
    Next
    S(0) = S0: S(1) = S1: S(2) = S2: S(3) = S3: S(4) = S4: S(5) = S5
End Sub

Public Sub silk_resampler_private_up2_HQ_wrapper(SS As silk_resampler_state_struct, out_() As Integer, ByVal outOffset As Long, _
        in_() As Integer, ByVal inOffset As Long, ByVal len_ As Long)
    silk_resampler_private_up2_HQ SS.sIIR, out_, outOffset, in_, inOffset, len_
End Sub
