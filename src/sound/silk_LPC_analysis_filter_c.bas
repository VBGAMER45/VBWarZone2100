Attribute VB_Name = "silk_LPC_analysis_filter_c"
' libopus 1.5.2 (ref/opus/silk): LPC_analysis_filter.c (the float build: USE_CELT_FIR applies to FIXED_POINT
' only, so the C loop). (VB) out, in and B are arrays and offsets; the silk_SMLABB_ovflw taps are written out
' (the products of two int16 fit in 32 bits; the sum wraps as the C's unsigned casts do).
Option Explicit

'/*******************************************/
'/* LPC analysis filter                     */
'/* NB! State is kept internally and the    */
'/* filter always starts with zero state    */
'/* first d output samples are set to zero  */
'/*******************************************/

'/* OPT: Using celt_fir() for this function should be faster, but it may cause
'   integer overflows in intermediate values (not final results), which the
'   current implementation silences by casting to unsigned. Enabling
'   this should be safe in pretty much all cases, even though it is not technically
'   C89-compliant. */
' #define USE_CELT_FIR 0

Public Sub silk_LPC_analysis_filter(out_() As Integer, ByVal outOffset As Long, in_() As Integer, ByVal inOffset As Long, _
        B() As Integer, ByVal BOffset As Long, ByVal len_ As Long, ByVal d As Long, ByVal arch As Long)
    Dim j As Long
    Dim ix As Long
    Dim out32_Q12 As Long, out32 As Long
    Dim in_ptr As Long                                  ' const opus_int16 *: an offset into in_

    ' celt_assert( d >= 6 );
    ' celt_assert( (d & 1) == 0 );
    ' celt_assert( d <= len );

    For ix = d To len_ - 1
        in_ptr = inOffset + ix - 1

        out32_Q12 = CLng(in_(in_ptr)) * B(BOffset)
        '/* Allowing wrap around so that two wraps can cancel each other. The rare
        '   cases where the result wraps around can only be triggered by invalid streams*/
        out32_Q12 = out32_Q12 + CLng(in_(in_ptr - 1)) * B(BOffset + 1)
        out32_Q12 = out32_Q12 + CLng(in_(in_ptr - 2)) * B(BOffset + 2)
        out32_Q12 = out32_Q12 + CLng(in_(in_ptr - 3)) * B(BOffset + 3)
        out32_Q12 = out32_Q12 + CLng(in_(in_ptr - 4)) * B(BOffset + 4)
        out32_Q12 = out32_Q12 + CLng(in_(in_ptr - 5)) * B(BOffset + 5)
        For j = 6 To d - 1 Step 2
            out32_Q12 = out32_Q12 + CLng(in_(in_ptr - j)) * B(BOffset + j)
            out32_Q12 = out32_Q12 + CLng(in_(in_ptr - j - 1)) * B(BOffset + j + 1)
        Next

        '/* Subtract prediction */
        out32_Q12 = CLng(in_(in_ptr + 1)) * &H1000& - out32_Q12  ' silk_SUB32_ovflw( silk_LSHIFT( (opus_int32)in_ptr[ 1 ], 12 ), out32_Q12 )

        '/* Scale to Q0 */
        out32 = silk_RSHIFT_ROUND(out32_Q12, 12)

        '/* Saturate output */
        out_(outOffset + ix) = silk_SAT16(out32)
    Next

    '/* Set first d output samples to zero */
    silk_memset_i16 out_, outOffset, d
End Sub
