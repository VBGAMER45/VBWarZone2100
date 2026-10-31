Attribute VB_Name = "silk_code_signs_c"
' libopus 1.5.2 (ref/opus/silk): code_signs.c, the decoder's half (silk_decode_signs; silk_encode_signs is the
' encoder's).
Option Explicit

'/*#define silk_enc_map(a)                ((a) > 0 ? 1 : 0)*/
'/*#define silk_dec_map(a)                ((a) > 0 ? 1 : -1)*/
'/* shifting avoids if-statement */
' #define silk_dec_map(a)                  ( silk_LSHIFT( (a),  1 ) - 1 )

'/* Decodes signs of excitation */
Public Sub silk_decode_signs(psRangeDec As ec_ctx, pulses() As Integer, ByVal length As Long, ByVal signalType As Long, _
        ByVal quantOffsetType As Long, sum_pulses() As Long)
    Dim i As Long, j As Long, p As Long
    Dim icdf(0 To 1) As Byte
    Dim q_ptr As Long                                   ' opus_int16 *: an offset into pulses
    Dim icdf_ptr As Long                                ' const opus_uint8 *: an offset into silk_rom_u8

    icdf(1) = 0
    q_ptr = 0
    i = silk_SMULBB(7, silk_ADD_LSHIFT(quantOffsetType, signalType, 1))
    icdf_ptr = silk_sign_iCDF + i
    length = silk_RSHIFT(length + SHELL_CODEC_FRAME_LENGTH \ 2, LOG2_SHELL_CODEC_FRAME_LENGTH)
    For i = 0 To length - 1
        p = sum_pulses(i)
        If p > 0 Then
            icdf(0) = silk_rom_u8(icdf_ptr + silk_min(p And &H1F, 6))
            For j = 0 To SHELL_CODEC_FRAME_LENGTH - 1
                If pulses(q_ptr + j) > 0 Then
                    '/* attach sign */
                    '/* implementation with shift, subtraction, multiplication */
                    pulses(q_ptr + j) = CLng(pulses(q_ptr + j)) * (ec_dec_icdf(psRangeDec, icdf, 0, 8) * 2 - 1)
                End If
            Next
        End If
        q_ptr = q_ptr + SHELL_CODEC_FRAME_LENGTH
    Next
End Sub
