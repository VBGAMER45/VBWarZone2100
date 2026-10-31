Attribute VB_Name = "silk_bwexpander_c"
' libopus 1.5.2 (ref/opus/silk): bwexpander.c. (VB) ar is an array and an offset.
Option Explicit

'/* Chirp (bandwidth expand) LP AR filter */
Public Sub silk_bwexpander(ar() As Integer, ByVal arOffset As Long, ByVal d As Long, ByVal chirp_Q16 As Long)
    Dim i As Long
    Dim chirp_minus_one_Q16 As Long

    chirp_minus_one_Q16 = chirp_Q16 - 65536

    '/* NB: Dont use silk_SMULWB, instead of silk_RSHIFT_ROUND( silk_MUL(), 16 ), below.  */
    '/* Bias in silk_SMULWB can lead to unstable filters                                */
    For i = 0 To d - 2
        ar(arOffset + i) = silk_RSHIFT_ROUND(chirp_Q16 * ar(arOffset + i), 16)
        chirp_Q16 = chirp_Q16 + silk_RSHIFT_ROUND(chirp_Q16 * chirp_minus_one_Q16, 16)
    Next
    ar(arOffset + d - 1) = silk_RSHIFT_ROUND(chirp_Q16 * ar(arOffset + d - 1), 16)
End Sub
