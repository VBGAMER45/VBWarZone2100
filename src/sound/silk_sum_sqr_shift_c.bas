Attribute VB_Name = "silk_sum_sqr_shift_c"
' libopus 1.5.2 (ref/opus/silk): sum_sqr_shift.c. (VB) x is an array and an offset; nrg_tmp (opus_uint32) is a
' Long holding its bits, shifted logically.
Option Explicit

'/* Compute number of bits to right shift the sum of squares of a vector */
'/* of int16s to make it fit in an int32                                 */
Public Sub silk_sum_sqr_shift(energy As Long, shift As Long, x() As Integer, ByVal xOffset As Long, ByVal len_ As Long)
    Dim i As Long, shft As Long
    Dim nrg_tmp As Long                                 ' opus_uint32
    Dim nrg As Long

    '/* Do a first run with the maximum shift we could have. */
    shft = 31 - silk_CLZ32(len_)
    '/* Let's be conservative with rounding and start with nrg=len. */
    nrg = len_
    i = 0
    Do While i < len_ - 1
        nrg_tmp = CLng(x(xOffset + i)) * x(xOffset + i)
        nrg_tmp = nrg_tmp + CLng(x(xOffset + i + 1)) * x(xOffset + i + 1)      ' silk_SMLABB_ovflw
        nrg = silk_ADD_RSHIFT_uint(nrg, nrg_tmp, shft)
        i = i + 2
    Loop
    If i < len_ Then
        '/* One sample left to process */
        nrg_tmp = CLng(x(xOffset + i)) * x(xOffset + i)
        nrg = silk_ADD_RSHIFT_uint(nrg, nrg_tmp, shft)
    End If
    ' silk_assert( nrg >= 0 );

    '/* Make sure the result will fit in a 32-bit signed integer with two bits
    '   of headroom. */
    shft = silk_max_32(0, shft + 3 - silk_CLZ32(nrg))
    nrg = 0
    i = 0
    Do While i < len_ - 1
        nrg_tmp = CLng(x(xOffset + i)) * x(xOffset + i)
        nrg_tmp = nrg_tmp + CLng(x(xOffset + i + 1)) * x(xOffset + i + 1)      ' silk_SMLABB_ovflw
        nrg = silk_ADD_RSHIFT_uint(nrg, nrg_tmp, shft)
        i = i + 2
    Loop
    If i < len_ Then
        '/* One sample left to process */
        nrg_tmp = CLng(x(xOffset + i)) * x(xOffset + i)
        nrg = silk_ADD_RSHIFT_uint(nrg, nrg_tmp, shft)
    End If

    ' silk_assert( nrg >= 0 );

    '/* Output arguments */
    shift = shft
    energy = nrg
End Sub
