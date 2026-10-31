Attribute VB_Name = "theora_idct_c"
' ref/theora (libtheora 1.2.0) lib/idct.c with dct.h: the inverse DCT (the C version).
' (VB) The coefficient buffers (ogg_int16_t) are Long arrays holding 16-bit values; _y and _x are an array and
' an offset (y(yo), x(xo)); the C's stores to ogg_int16_t are written out as ((v + 32768) And &HFFFF&) -
' 32768, and its arithmetic shifts by 16 as (v And &HFFFF0000) \ &H10000 (exact: the low bits are cleared).
' The scratch block w is a module array.
Option Explicit

' cos(n*pi/16) (resp. sin(m*pi/16)) scaled by 65536.
Public Const OC_C1S7 As Long = 64277
Public Const OC_C2S6 As Long = 60547
Public Const OC_C3S5 As Long = 54491
Public Const OC_C4S4 As Long = 46341
Public Const OC_C5S3 As Long = 36410
Public Const OC_C6S2 As Long = 25080
Public Const OC_C7S1 As Long = 12785

Private w(0 To 63) As Long      ' ogg_int16_t w[64] of the 8x8 transforms

' Performs an inverse 8 point Type-II DCT transform. The output is scaled by a factor of 2 relative to the
' orthonormal version of the transform.
'  _y: The buffer to store the result in. Data will be placed in every 8th entry (e.g., in a column of an 8x8
'       block).
'  _x: The input coefficients. The first 8 entries are used (e.g., from a row of an 8x8 block).
Private Sub idct8(y() As Long, ByVal yo As Long, x() As Long, ByVal xo As Long)
    Dim t0 As Long, t1 As Long, t2 As Long, t3 As Long, t4 As Long, t5 As Long, t6 As Long, t7 As Long
    Dim r As Long
    ' Stage 1:
    ' 0-1 butterfly.
    t0 = (OC_C4S4 * (((x(xo) + x(xo + 4) + 32768) And &HFFFF&) - 32768) And &HFFFF0000) \ &H10000
    t1 = (OC_C4S4 * (((x(xo) - x(xo + 4) + 32768) And &HFFFF&) - 32768) And &HFFFF0000) \ &H10000
    ' 2-3 rotation by 6pi/16.
    t2 = ((OC_C6S2 * x(xo + 2)) And &HFFFF0000) \ &H10000 - ((OC_C2S6 * x(xo + 6)) And &HFFFF0000) \ &H10000
    t3 = ((OC_C2S6 * x(xo + 2)) And &HFFFF0000) \ &H10000 + ((OC_C6S2 * x(xo + 6)) And &HFFFF0000) \ &H10000
    ' 4-7 rotation by 7pi/16.
    t4 = ((OC_C7S1 * x(xo + 1)) And &HFFFF0000) \ &H10000 - ((OC_C1S7 * x(xo + 7)) And &HFFFF0000) \ &H10000
    ' 5-6 rotation by 3pi/16.
    t5 = ((OC_C3S5 * x(xo + 5)) And &HFFFF0000) \ &H10000 - ((OC_C5S3 * x(xo + 3)) And &HFFFF0000) \ &H10000
    t6 = ((OC_C5S3 * x(xo + 5)) And &HFFFF0000) \ &H10000 + ((OC_C3S5 * x(xo + 3)) And &HFFFF0000) \ &H10000
    t7 = ((OC_C1S7 * x(xo + 1)) And &HFFFF0000) \ &H10000 + ((OC_C7S1 * x(xo + 7)) And &HFFFF0000) \ &H10000
    ' Stage 2:
    ' 4-5 butterfly.
    r = t4 + t5
    t5 = (OC_C4S4 * (((t4 - t5 + 32768) And &HFFFF&) - 32768) And &HFFFF0000) \ &H10000
    t4 = r
    ' 7-6 butterfly.
    r = t7 + t6
    t6 = (OC_C4S4 * (((t7 - t6 + 32768) And &HFFFF&) - 32768) And &HFFFF0000) \ &H10000
    t7 = r
    ' Stage 3:
    ' 0-3 butterfly.
    r = t0 + t3
    t3 = t0 - t3
    t0 = r
    ' 1-2 butterfly.
    r = t1 + t2
    t2 = t1 - t2
    t1 = r
    ' 6-5 butterfly.
    r = t6 + t5
    t5 = t6 - t5
    t6 = r
    ' Stage 4:
    ' 0-7 butterfly.
    y(yo) = ((t0 + t7 + 32768) And &HFFFF&) - 32768
    ' 1-6 butterfly.
    y(yo + 8) = ((t1 + t6 + 32768) And &HFFFF&) - 32768
    ' 2-5 butterfly.
    y(yo + 16) = ((t2 + t5 + 32768) And &HFFFF&) - 32768
    ' 3-4 butterfly.
    y(yo + 24) = ((t3 + t4 + 32768) And &HFFFF&) - 32768
    y(yo + 32) = ((t3 - t4 + 32768) And &HFFFF&) - 32768
    y(yo + 40) = ((t2 - t5 + 32768) And &HFFFF&) - 32768
    y(yo + 48) = ((t1 - t6 + 32768) And &HFFFF&) - 32768
    y(yo + 56) = ((t0 - t7 + 32768) And &HFFFF&) - 32768
End Sub

' Performs an inverse 8 point Type-II DCT transform. The output is scaled by a factor of 2 relative to the
' orthonormal version of the transform.
'  _y: The buffer to store the result in. Data will be placed in every 8th entry (e.g., in a column of an 8x8
'       block).
'  _x: The input coefficients. Only the first 4 entries are used. The other 4 are assumed to be 0.
Private Sub idct8_4(y() As Long, ByVal yo As Long, x() As Long, ByVal xo As Long)
    Dim t0 As Long, t1 As Long, t2 As Long, t3 As Long, t4 As Long, t5 As Long, t6 As Long, t7 As Long
    Dim r As Long
    ' Stage 1:
    t0 = ((OC_C4S4 * x(xo)) And &HFFFF0000) \ &H10000
    t2 = ((OC_C6S2 * x(xo + 2)) And &HFFFF0000) \ &H10000
    t3 = ((OC_C2S6 * x(xo + 2)) And &HFFFF0000) \ &H10000
    t4 = ((OC_C7S1 * x(xo + 1)) And &HFFFF0000) \ &H10000
    t5 = -(((OC_C5S3 * x(xo + 3)) And &HFFFF0000) \ &H10000)
    t6 = ((OC_C3S5 * x(xo + 3)) And &HFFFF0000) \ &H10000
    t7 = ((OC_C1S7 * x(xo + 1)) And &HFFFF0000) \ &H10000
    ' Stage 2:
    r = t4 + t5
    t5 = (OC_C4S4 * (((t4 - t5 + 32768) And &HFFFF&) - 32768) And &HFFFF0000) \ &H10000
    t4 = r
    r = t7 + t6
    t6 = (OC_C4S4 * (((t7 - t6 + 32768) And &HFFFF&) - 32768) And &HFFFF0000) \ &H10000
    t7 = r
    ' Stage 3:
    t1 = t0 + t2
    t2 = t0 - t2
    r = t0 + t3
    t3 = t0 - t3
    t0 = r
    r = t6 + t5
    t5 = t6 - t5
    t6 = r
    ' Stage 4:
    y(yo) = ((t0 + t7 + 32768) And &HFFFF&) - 32768
    y(yo + 8) = ((t1 + t6 + 32768) And &HFFFF&) - 32768
    y(yo + 16) = ((t2 + t5 + 32768) And &HFFFF&) - 32768
    y(yo + 24) = ((t3 + t4 + 32768) And &HFFFF&) - 32768
    y(yo + 32) = ((t3 - t4 + 32768) And &HFFFF&) - 32768
    y(yo + 40) = ((t2 - t5 + 32768) And &HFFFF&) - 32768
    y(yo + 48) = ((t1 - t6 + 32768) And &HFFFF&) - 32768
    y(yo + 56) = ((t0 - t7 + 32768) And &HFFFF&) - 32768
End Sub

' Performs an inverse 8 point Type-II DCT transform. The output is scaled by a factor of 2 relative to the
' orthonormal version of the transform.
'  _y: The buffer to store the result in. Data will be placed in every 8th entry (e.g., in a column of an 8x8
'       block).
'  _x: The input coefficients. Only the first 3 entries are used. The other 5 are assumed to be 0.
Private Sub idct8_3(y() As Long, ByVal yo As Long, x() As Long, ByVal xo As Long)
    Dim t0 As Long, t1 As Long, t2 As Long, t3 As Long, t4 As Long, t5 As Long, t6 As Long, t7 As Long
    Dim r As Long
    ' Stage 1:
    t0 = ((OC_C4S4 * x(xo)) And &HFFFF0000) \ &H10000
    t2 = ((OC_C6S2 * x(xo + 2)) And &HFFFF0000) \ &H10000
    t3 = ((OC_C2S6 * x(xo + 2)) And &HFFFF0000) \ &H10000
    t4 = ((OC_C7S1 * x(xo + 1)) And &HFFFF0000) \ &H10000
    t7 = ((OC_C1S7 * x(xo + 1)) And &HFFFF0000) \ &H10000
    ' Stage 2:
    t5 = ((OC_C4S4 * t4) And &HFFFF0000) \ &H10000
    t6 = ((OC_C4S4 * t7) And &HFFFF0000) \ &H10000
    ' Stage 3:
    t1 = t0 + t2
    t2 = t0 - t2
    r = t0 + t3
    t3 = t0 - t3
    t0 = r
    r = t6 + t5
    t5 = t6 - t5
    t6 = r
    ' Stage 4:
    y(yo) = ((t0 + t7 + 32768) And &HFFFF&) - 32768
    y(yo + 8) = ((t1 + t6 + 32768) And &HFFFF&) - 32768
    y(yo + 16) = ((t2 + t5 + 32768) And &HFFFF&) - 32768
    y(yo + 24) = ((t3 + t4 + 32768) And &HFFFF&) - 32768
    y(yo + 32) = ((t3 - t4 + 32768) And &HFFFF&) - 32768
    y(yo + 40) = ((t2 - t5 + 32768) And &HFFFF&) - 32768
    y(yo + 48) = ((t1 - t6 + 32768) And &HFFFF&) - 32768
    y(yo + 56) = ((t0 - t7 + 32768) And &HFFFF&) - 32768
End Sub

' Performs an inverse 8 point Type-II DCT transform. The output is scaled by a factor of 2 relative to the
' orthonormal version of the transform.
'  _y: The buffer to store the result in. Data will be placed in every 8th entry (e.g., in a column of an 8x8
'       block).
'  _x: The input coefficients. Only the first 2 entries are used. The other 6 are assumed to be 0.
Private Sub idct8_2(y() As Long, ByVal yo As Long, x() As Long, ByVal xo As Long)
    Dim t0 As Long, t4 As Long, t5 As Long, t6 As Long, t7 As Long
    Dim r As Long
    ' Stage 1:
    t0 = ((OC_C4S4 * x(xo)) And &HFFFF0000) \ &H10000
    t4 = ((OC_C7S1 * x(xo + 1)) And &HFFFF0000) \ &H10000
    t7 = ((OC_C1S7 * x(xo + 1)) And &HFFFF0000) \ &H10000
    ' Stage 2:
    t5 = ((OC_C4S4 * t4) And &HFFFF0000) \ &H10000
    t6 = ((OC_C4S4 * t7) And &HFFFF0000) \ &H10000
    ' Stage 3:
    r = t6 + t5
    t5 = t6 - t5
    t6 = r
    ' Stage 4:
    y(yo) = ((t0 + t7 + 32768) And &HFFFF&) - 32768
    y(yo + 8) = ((t0 + t6 + 32768) And &HFFFF&) - 32768
    y(yo + 16) = ((t0 + t5 + 32768) And &HFFFF&) - 32768
    y(yo + 24) = ((t0 + t4 + 32768) And &HFFFF&) - 32768
    y(yo + 32) = ((t0 - t4 + 32768) And &HFFFF&) - 32768
    y(yo + 40) = ((t0 - t5 + 32768) And &HFFFF&) - 32768
    y(yo + 48) = ((t0 - t6 + 32768) And &HFFFF&) - 32768
    y(yo + 56) = ((t0 - t7 + 32768) And &HFFFF&) - 32768
End Sub

' Performs an inverse 8 point Type-II DCT transform. The output is scaled by a factor of 2 relative to the
' orthonormal version of the transform.
'  _y: The buffer to store the result in. Data will be placed in every 8th entry (e.g., in a column of an 8x8
'       block).
'  _x: The input coefficients. Only the first entry is used. The other 7 are assumed to be 0.
Private Sub idct8_1(y() As Long, ByVal yo As Long, x() As Long, ByVal xo As Long)
    Dim v As Long
    ' (ogg_int16_t)(OC_C4S4*_x[0]>>16): (VB) |result| < |_x[0]|, no wrap
    v = ((OC_C4S4 * x(xo)) And &HFFFF0000) \ &H10000
    y(yo) = v
    y(yo + 8) = v
    y(yo + 16) = v
    y(yo + 24) = v
    y(yo + 32) = v
    y(yo + 40) = v
    y(yo + 48) = v
    y(yo + 56) = v
End Sub

' Performs an inverse 8x8 Type-II DCT transform. The input is assumed to be scaled by a factor of 4 relative to
' orthonormal version of the transform. All coefficients but the first 3 in zig-zag scan order are assumed to
' be 0:
'  x  x  0  0  0  0  0  0
'  x  0  0  0  0  0  0  0
'  0  0  0  0  0  0  0  0
'  0  0  0  0  0  0  0  0
'  0  0  0  0  0  0  0  0
'  0  0  0  0  0  0  0  0
'  0  0  0  0  0  0  0  0
'  0  0  0  0  0  0  0  0
'  _y: The buffer to store the result in. This may be the same as _x.
'  _x: The input coefficients.
Private Sub oc_idct8x8_3(y() As Long, ByVal yo As Long, x() As Long, ByVal xo As Long)
    Dim i As Long
    ' Transform rows of x into columns of w.
    idct8_2 w, 0, x, xo
    idct8_1 w, 1, x, xo + 8
    ' Transform rows of w into columns of y.
    For i = 0 To 7
        idct8_2 y, yo + i, w, i * 8
    Next
    ' Adjust for the scale factor. (_y[i]+8>>4 fits in 16 bits)
    For i = 0 To 63
        y(yo + i) = ((y(yo + i) + 8) And &HFFFFFFF0) \ 16
    Next
    ' Clear input data for next block.
    x(xo) = 0
    x(xo + 1) = 0
    x(xo + 8) = 0
End Sub

' Performs an inverse 8x8 Type-II DCT transform. The input is assumed to be scaled by a factor of 4 relative to
' orthonormal version of the transform. All coefficients but the first 10 in zig-zag scan order are assumed to
' be 0:
'  x  x  x  x  0  0  0  0
'  x  x  x  0  0  0  0  0
'  x  x  0  0  0  0  0  0
'  x  0  0  0  0  0  0  0
'  0  0  0  0  0  0  0  0
'  0  0  0  0  0  0  0  0
'  0  0  0  0  0  0  0  0
'  0  0  0  0  0  0  0  0
'  _y: The buffer to store the result in. This may be the same as _x.
'  _x: The input coefficients.
Private Sub oc_idct8x8_10(y() As Long, ByVal yo As Long, x() As Long, ByVal xo As Long)
    Dim i As Long
    ' Transform rows of x into columns of w.
    idct8_4 w, 0, x, xo
    idct8_3 w, 1, x, xo + 8
    idct8_2 w, 2, x, xo + 16
    idct8_1 w, 3, x, xo + 24
    ' Transform rows of w into columns of y.
    For i = 0 To 7
        idct8_4 y, yo + i, w, i * 8
    Next
    ' Adjust for the scale factor. (_y[i]+8>>4 fits in 16 bits)
    For i = 0 To 63
        y(yo + i) = ((y(yo + i) + 8) And &HFFFFFFF0) \ 16
    Next
    ' Clear input data for next block.
    x(xo) = 0
    x(xo + 1) = 0
    x(xo + 2) = 0
    x(xo + 3) = 0
    x(xo + 8) = 0
    x(xo + 9) = 0
    x(xo + 10) = 0
    x(xo + 16) = 0
    x(xo + 17) = 0
    x(xo + 24) = 0
End Sub

' Performs an inverse 8x8 Type-II DCT transform. The input is assumed to be scaled by a factor of 4 relative to
' orthonormal version of the transform.
'  _y: The buffer to store the result in. This may be the same as _x.
'  _x: The input coefficients.
Private Sub oc_idct8x8_slow(y() As Long, ByVal yo As Long, x() As Long, ByVal xo As Long)
    Dim i As Long
    ' Transform rows of x into columns of w.
    For i = 0 To 7
        idct8 w, i, x, xo + i * 8
    Next
    ' Transform rows of w into columns of y.
    For i = 0 To 7
        idct8 y, yo + i, w, i * 8
    Next
    ' Adjust for the scale factor. (_y[i]+8>>4 fits in 16 bits)
    For i = 0 To 63
        y(yo + i) = ((y(yo + i) + 8) And &HFFFFFFF0) \ 16
    Next
    ' Clear input data for next block.
    For i = 0 To 63
        x(xo + i) = 0
    Next
End Sub

' Performs an inverse 8x8 Type-II DCT transform. The input is assumed to be scaled by a factor of 4 relative to
' orthonormal version of the transform.
Public Sub oc_idct8x8_c(y() As Long, ByVal yo As Long, x() As Long, ByVal xo As Long, ByVal last_zzi As Long)
    ' _last_zzi is subtly different from an actual count of the number of coefficients we decoded for this
    ' block. It contains the value of zzi BEFORE the final token in the block was decoded. In most cases this
    ' is an EOB token (the continuation of an EOB run from a previous block counts), and so this is the same as
    ' the coefficient count. However, in the case that the last token was NOT an EOB token, but filled the
    ' block up with exactly 64 coefficients, _last_zzi will be less than 64. Provided the last token was not a
    ' pure zero run, the minimum value it can be is 46, and so that doesn't affect any of the cases in this
    ' routine. However, if the last token WAS a pure zero run of length 63, then _last_zzi will be 1 while the
    ' number of coefficients decoded is 64. Thus, we will trigger the following special case, where the real
    ' coefficient count would not. Note also that a zero run of length 64 will give _last_zzi a value of 0, but
    ' we still process the DC coefficient, which might have a non-zero value due to DC prediction. Although
    ' convoluted, this is arguably the correct behavior: it allows us to use a smaller transform when the block
    ' ends with a long zero run instead of a normal EOB token. It could be smarter... multiple separate zero
    ' runs at the end of a block will fool it, but an encoder that generates these really deserves what it
    ' gets. Needless to say we inherited this approach from VP3.
    ' Then perform the iDCT.
    If last_zzi <= 3 Then
        oc_idct8x8_3 y, yo, x, xo
    ElseIf last_zzi <= 10 Then
        oc_idct8x8_10 y, yo, x, xo
    Else
        oc_idct8x8_slow y, yo, x, xo
    End If
End Sub
