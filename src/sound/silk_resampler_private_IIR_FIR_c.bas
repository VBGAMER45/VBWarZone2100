Attribute VB_Name = "silk_resampler_private_IIR_FIR_c"
' libopus 1.5.2 (ref/opus/silk): resampler_private_IIR_FIR.c. (VB) out and in are arrays and offsets; the
' INTERPOL function returns the output offset (C: the pointer); the stack buffer is this module's array at
' its largest size. The FIR's silk_SMULBB / silk_SMLABB (int16 by int16) are written out.
Option Explicit

Private buf(0 To 967) As Integer                        ' [ 2 * S->batchSize + RESAMPLER_ORDER_FIR_12 ] (batchSize <= 480)

Private Function silk_resampler_private_IIR_FIR_INTERPOL(out_() As Integer, ByVal outOffset As Long, _
        ByVal max_index_Q16 As Long, ByVal index_increment_Q16 As Long) As Long
    Dim index_Q16 As Long, res_Q15 As Long
    Dim buf_ptr As Long                                 ' opus_int16 *: an offset into buf
    Dim table_index As Long
    Dim t0 As Long, t1 As Long                          ' silk_resampler_frac_FIR_12[ table_index ], [ 11 - table_index ]: offsets into silk_rom_i16

    '/* Interpolate upsampled signal and store in output array */
    index_Q16 = 0
    Do While index_Q16 < max_index_Q16
        table_index = ((index_Q16 And &HFFFF&) * 12) \ &H10000          ' silk_SMULWB( index_Q16 & 0xFFFF, 12 )
        buf_ptr = (index_Q16 And &HFFFF0000) \ &H10000                  ' index_Q16 >> 16
        t0 = silk_resampler_frac_FIR_12 + table_index * 4
        t1 = silk_resampler_frac_FIR_12 + (11 - table_index) * 4
        res_Q15 = CLng(buf(buf_ptr)) * silk_rom_i16(t0)
        res_Q15 = res_Q15 + CLng(buf(buf_ptr + 1)) * silk_rom_i16(t0 + 1)
        res_Q15 = res_Q15 + CLng(buf(buf_ptr + 2)) * silk_rom_i16(t0 + 2)
        res_Q15 = res_Q15 + CLng(buf(buf_ptr + 3)) * silk_rom_i16(t0 + 3)
        res_Q15 = res_Q15 + CLng(buf(buf_ptr + 4)) * silk_rom_i16(t1 + 3)
        res_Q15 = res_Q15 + CLng(buf(buf_ptr + 5)) * silk_rom_i16(t1 + 2)
        res_Q15 = res_Q15 + CLng(buf(buf_ptr + 6)) * silk_rom_i16(t1 + 1)
        res_Q15 = res_Q15 + CLng(buf(buf_ptr + 7)) * silk_rom_i16(t1)
        ' *out++ = (opus_int16)silk_SAT16( silk_RSHIFT_ROUND( res_Q15, 15 ) );
        res_Q15 = (((res_Q15 And &HFFFFC000) \ &H4000&) + 1)            ' ( res_Q15 >> 14 ) + 1
        res_Q15 = (res_Q15 And &HFFFFFFFE) \ 2                          ' >> 1
        If res_Q15 > 32767 Then
            res_Q15 = 32767
        ElseIf res_Q15 < -32768 Then
            res_Q15 = -32768
        End If
        out_(outOffset) = res_Q15
        outOffset = outOffset + 1
        index_Q16 = index_Q16 + index_increment_Q16
    Loop
    silk_resampler_private_IIR_FIR_INTERPOL = outOffset
End Function

'/* Upsample using a combination of allpass-based 2x upsampling and FIR interpolation */
Public Sub silk_resampler_private_IIR_FIR(SS As silk_resampler_state_struct, out_() As Integer, ByVal outOffset As Long, _
        in_() As Integer, ByVal inOffset As Long, ByVal inLen As Long)
    Dim nSamplesIn As Long
    Dim max_index_Q16 As Long, index_increment_Q16 As Long

    '/* Copy buffered samples to start of buffer */
    silk_memcpy_i16 buf, 0, SS.sFIR_i16, 0, RESAMPLER_ORDER_FIR_12

    '/* Iterate over blocks of frameSizeIn input samples */
    index_increment_Q16 = SS.invRatio_Q16
    Do
        nSamplesIn = silk_min(inLen, SS.batchSize)

        '/* Upsample 2x */
        silk_resampler_private_up2_HQ SS.sIIR, buf, RESAMPLER_ORDER_FIR_12, in_, inOffset, nSamplesIn

        max_index_Q16 = nSamplesIn * &H20000            ' silk_LSHIFT32( nSamplesIn, 16 + 1 ) /* + 1 because 2x upsampling */
        outOffset = silk_resampler_private_IIR_FIR_INTERPOL(out_, outOffset, max_index_Q16, index_increment_Q16)
        inOffset = inOffset + nSamplesIn
        inLen = inLen - nSamplesIn

        If inLen > 0 Then
            '/* More iterations to do; copy last part of filtered signal to beginning of buffer */
            silk_memcpy_i16 buf, 0, buf, nSamplesIn * 2, RESAMPLER_ORDER_FIR_12
        Else
            Exit Do
        End If
    Loop

    '/* Copy last part of filtered signal to the state for the next call */
    silk_memcpy_i16 SS.sFIR_i16, 0, buf, nSamplesIn * 2, RESAMPLER_ORDER_FIR_12
End Sub
