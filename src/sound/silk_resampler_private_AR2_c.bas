Attribute VB_Name = "silk_resampler_private_AR2_c"
' libopus 1.5.2 (ref/opus/silk): resampler_private_AR2.c. (VB) out_Q8 and in are arrays and offsets; A_Q14 is
' an offset into silk_rom_i16.
Option Explicit

'/* Second order AR filter with single delay elements */
Public Sub silk_resampler_private_AR2(S() As Long, out_Q8() As Long, ByVal out_Q8Offset As Long, in_() As Integer, _
        ByVal inOffset As Long, ByVal A_Q14 As Long, ByVal len_ As Long)
    Dim k As Long
    Dim out32 As Long

    For k = 0 To len_ - 1
        out32 = S(0) + CLng(in_(inOffset + k)) * 256    ' silk_ADD_LSHIFT32( S[ 0 ], (opus_int32)in[ k ], 8 )
        out_Q8(out_Q8Offset + k) = out32
        out32 = out32 * 4                               ' silk_LSHIFT( out32, 2 )
        S(0) = silk_SMLAWB(S(1), out32, silk_rom_i16(A_Q14))
        S(1) = silk_SMULWB(out32, silk_rom_i16(A_Q14 + 1))
    Next
End Sub
