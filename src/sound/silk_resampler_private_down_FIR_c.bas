Attribute VB_Name = "silk_resampler_private_down_FIR_c"
' libopus 1.5.2 (ref/opus/silk): resampler_private_down_FIR.c. (VB) out and in are arrays and offsets;
' FIR_Coefs is an offset into silk_rom_i16; the INTERPOL function returns the output offset (C: the pointer);
' the stack buffer is this module's array at its largest size. The unrolled taps are loops (the same 32-bit
' sums).
Option Explicit

Private buf(0 To 515) As Long                           ' [ S->batchSize + S->FIR_Order ] (batchSize <= 480)

Private Function silk_resampler_private_down_FIR_INTERPOL(out_() As Integer, ByVal outOffset As Long, ByVal FIR_Coefs As Long, _
        ByVal FIR_Order As Long, ByVal FIR_Fracs As Long, ByVal max_index_Q16 As Long, ByVal index_increment_Q16 As Long) As Long
    Dim index_Q16 As Long, res_Q6 As Long
    Dim buf_ptr As Long                                 ' opus_int32 *: an offset into buf
    Dim interpol_ind As Long
    Dim interpol_ptr As Long                            ' const opus_int16 *: an offset into silk_rom_i16
    Dim j As Long

    Select Case FIR_Order
        Case RESAMPLER_DOWN_ORDER_FIR0
            index_Q16 = 0
            Do While index_Q16 < max_index_Q16
                '/* Integer part gives pointer to buffered input */
                buf_ptr = silk_RSHIFT(index_Q16, 16)

                '/* Fractional part gives interpolation coefficients */
                interpol_ind = silk_SMULWB(index_Q16 And &HFFFF&, FIR_Fracs)

                '/* Inner product */
                interpol_ptr = FIR_Coefs + (RESAMPLER_DOWN_ORDER_FIR0 \ 2) * interpol_ind
                res_Q6 = silk_SMULWB(buf(buf_ptr), silk_rom_i16(interpol_ptr))
                For j = 1 To 8
                    res_Q6 = silk_SMLAWB(res_Q6, buf(buf_ptr + j), silk_rom_i16(interpol_ptr + j))
                Next
                interpol_ptr = FIR_Coefs + (RESAMPLER_DOWN_ORDER_FIR0 \ 2) * (FIR_Fracs - 1 - interpol_ind)
                For j = 0 To 8
                    res_Q6 = silk_SMLAWB(res_Q6, buf(buf_ptr + 17 - j), silk_rom_i16(interpol_ptr + j))
                Next

                '/* Scale down, saturate and store in output array */
                out_(outOffset) = silk_SAT16(silk_RSHIFT_ROUND(res_Q6, 6))
                outOffset = outOffset + 1
                index_Q16 = index_Q16 + index_increment_Q16
            Loop
        Case RESAMPLER_DOWN_ORDER_FIR1
            index_Q16 = 0
            Do While index_Q16 < max_index_Q16
                '/* Integer part gives pointer to buffered input */
                buf_ptr = silk_RSHIFT(index_Q16, 16)

                '/* Inner product */
                res_Q6 = silk_SMULWB(buf(buf_ptr) + buf(buf_ptr + 23), silk_rom_i16(FIR_Coefs))
                For j = 1 To 11
                    res_Q6 = silk_SMLAWB(res_Q6, buf(buf_ptr + j) + buf(buf_ptr + 23 - j), silk_rom_i16(FIR_Coefs + j))
                Next

                '/* Scale down, saturate and store in output array */
                out_(outOffset) = silk_SAT16(silk_RSHIFT_ROUND(res_Q6, 6))
                outOffset = outOffset + 1
                index_Q16 = index_Q16 + index_increment_Q16
            Loop
        Case RESAMPLER_DOWN_ORDER_FIR2
            index_Q16 = 0
            Do While index_Q16 < max_index_Q16
                '/* Integer part gives pointer to buffered input */
                buf_ptr = silk_RSHIFT(index_Q16, 16)

                '/* Inner product */
                res_Q6 = silk_SMULWB(buf(buf_ptr) + buf(buf_ptr + 35), silk_rom_i16(FIR_Coefs))
                For j = 1 To 17
                    res_Q6 = silk_SMLAWB(res_Q6, buf(buf_ptr + j) + buf(buf_ptr + 35 - j), silk_rom_i16(FIR_Coefs + j))
                Next

                '/* Scale down, saturate and store in output array */
                out_(outOffset) = silk_SAT16(silk_RSHIFT_ROUND(res_Q6, 6))
                outOffset = outOffset + 1
                index_Q16 = index_Q16 + index_increment_Q16
            Loop
        Case Else
            ' celt_assert( 0 );
    End Select
    silk_resampler_private_down_FIR_INTERPOL = outOffset
End Function

'/* Resample with a 2nd order AR filter followed by FIR interpolation */
Public Sub silk_resampler_private_down_FIR(SS As silk_resampler_state_struct, out_() As Integer, ByVal outOffset As Long, _
        in_() As Integer, ByVal inOffset As Long, ByVal inLen As Long)
    Dim nSamplesIn As Long
    Dim max_index_Q16 As Long, index_increment_Q16 As Long
    Dim FIR_Coefs As Long                               ' const opus_int16 *: an offset into silk_rom_i16

    '/* Copy buffered samples to start of buffer */
    silk_memcpy_i32 buf, 0, SS.sFIR_i32, 0, SS.FIR_Order

    FIR_Coefs = SS.Coefs + 2

    '/* Iterate over blocks of frameSizeIn input samples */
    index_increment_Q16 = SS.invRatio_Q16
    Do
        nSamplesIn = silk_min(inLen, SS.batchSize)

        '/* Second-order AR filter (output in Q8) */
        silk_resampler_private_AR2 SS.sIIR, buf, SS.FIR_Order, in_, inOffset, SS.Coefs, nSamplesIn

        max_index_Q16 = nSamplesIn * &H10000            ' silk_LSHIFT32( nSamplesIn, 16 )

        '/* Interpolate filtered signal */
        outOffset = silk_resampler_private_down_FIR_INTERPOL(out_, outOffset, FIR_Coefs, SS.FIR_Order, _
            SS.FIR_Fracs, max_index_Q16, index_increment_Q16)

        inOffset = inOffset + nSamplesIn
        inLen = inLen - nSamplesIn

        If inLen > 1 Then
            '/* More iterations to do; copy last part of filtered signal to beginning of buffer */
            silk_memcpy_i32 buf, 0, buf, nSamplesIn, SS.FIR_Order
        Else
            Exit Do
        End If
    Loop

    '/* Copy last part of filtered signal to the state for the next call */
    silk_memcpy_i32 SS.sFIR_i32, 0, buf, nSamplesIn, SS.FIR_Order
End Sub
