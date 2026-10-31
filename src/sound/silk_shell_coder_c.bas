Attribute VB_Name = "silk_shell_coder_c"
' libopus 1.5.2 (ref/opus/silk): shell_coder.c, the decoder's half (decode_split, silk_shell_decoder;
' combine_pulses, encode_split and silk_shell_encoder are the encoder's).
' (VB) a pointer into a pulses array is the array and an index.
Option Explicit

Private Sub decode_split(p_child1() As Integer, ByVal i1 As Long, p_child2() As Integer, ByVal i2 As Long, _
        psRangeDec As ec_ctx, ByVal p As Long, ByVal shell_table As Long)
    If p > 0 Then
        p_child1(i1) = ec_dec_icdf(psRangeDec, silk_rom_u8, shell_table + silk_rom_u8(silk_shell_code_table_offsets + p), 8)
        p_child2(i2) = p - p_child1(i1)
    Else
        p_child1(i1) = 0
        p_child2(i2) = 0
    End If
End Sub

'/* Shell decoder, operates on one shell code frame of 16 pulses */
Public Sub silk_shell_decoder(pulses0() As Integer, ByVal pulses0Offset As Long, psRangeDec As ec_ctx, ByVal pulses4 As Long)
    Dim pulses3(0 To 1) As Integer, pulses2(0 To 3) As Integer, pulses1(0 To 7) As Integer
    Dim o As Long

    '/* this function operates on one shell code frame of 16 pulses */
    ' silk_assert( SHELL_CODEC_FRAME_LENGTH == 16 );
    o = pulses0Offset

    decode_split pulses3, 0, pulses3, 1, psRangeDec, pulses4, silk_shell_code_table3
    decode_split pulses2, 0, pulses2, 1, psRangeDec, pulses3(0), silk_shell_code_table2

    decode_split pulses1, 0, pulses1, 1, psRangeDec, pulses2(0), silk_shell_code_table1
    decode_split pulses0, o + 0, pulses0, o + 1, psRangeDec, pulses1(0), silk_shell_code_table0
    decode_split pulses0, o + 2, pulses0, o + 3, psRangeDec, pulses1(1), silk_shell_code_table0

    decode_split pulses1, 2, pulses1, 3, psRangeDec, pulses2(1), silk_shell_code_table1
    decode_split pulses0, o + 4, pulses0, o + 5, psRangeDec, pulses1(2), silk_shell_code_table0
    decode_split pulses0, o + 6, pulses0, o + 7, psRangeDec, pulses1(3), silk_shell_code_table0

    decode_split pulses2, 2, pulses2, 3, psRangeDec, pulses3(1), silk_shell_code_table2

    decode_split pulses1, 4, pulses1, 5, psRangeDec, pulses2(2), silk_shell_code_table1
    decode_split pulses0, o + 8, pulses0, o + 9, psRangeDec, pulses1(4), silk_shell_code_table0
    decode_split pulses0, o + 10, pulses0, o + 11, psRangeDec, pulses1(5), silk_shell_code_table0

    decode_split pulses1, 6, pulses1, 7, psRangeDec, pulses2(3), silk_shell_code_table1
    decode_split pulses0, o + 12, pulses0, o + 13, psRangeDec, pulses1(6), silk_shell_code_table0
    decode_split pulses0, o + 14, pulses0, o + 15, psRangeDec, pulses1(7), silk_shell_code_table0
End Sub
