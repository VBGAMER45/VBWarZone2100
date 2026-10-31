Attribute VB_Name = "silk_bwexpander_32_c"
' libopus 1.5.2 (ref/opus/silk): bwexpander_32.c.
Option Explicit

'/* Chirp (bandwidth expand) LP AR filter.
'   This logic is reused in _celt_lpc(). Any bug fixes should also be applied there. */
Public Sub silk_bwexpander_32(ar() As Long, ByVal d As Long, ByVal chirp_Q16 As Long)
    Dim i As Long
    Dim chirp_minus_one_Q16 As Long

    chirp_minus_one_Q16 = chirp_Q16 - 65536

    For i = 0 To d - 2
        ar(i) = silk_SMULWW(chirp_Q16, ar(i))
        chirp_Q16 = chirp_Q16 + silk_RSHIFT_ROUND(chirp_Q16 * chirp_minus_one_Q16, 16)
    Next
    ar(d - 1) = silk_SMULWW(chirp_Q16, ar(d - 1))
End Sub
