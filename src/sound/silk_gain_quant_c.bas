Attribute VB_Name = "silk_gain_quant_c"
' libopus 1.5.2 (ref/opus/silk): gain_quant.c, the decoder's part (silk_gains_dequant; silk_gains_quant and
' silk_gains_ID are the encoder's).
Option Explicit

Private Const OFFSET As Long = 2090                     ' ( ( MIN_QGAIN_DB * 128 ) / 6 + 16 * 128 )
Private Const SCALE_Q16 As Long = 2251                  ' ( ( 65536 * ( N_LEVELS_QGAIN - 1 ) ) / ( ( ( MAX_QGAIN_DB - MIN_QGAIN_DB ) * 128 ) / 6 ) )
Private Const INV_SCALE_Q16 As Long = 1907825           ' ( ( 65536 * ( ( ( MAX_QGAIN_DB - MIN_QGAIN_DB ) * 128 ) / 6 ) ) / ( N_LEVELS_QGAIN - 1 ) )

'/* Gains scalar dequantization, uniform on log scale */
Public Sub silk_gains_dequant(gain_Q16() As Long, ind() As Long, prev_ind As Long, ByVal conditional As Long, ByVal nb_subfr As Long)
    Dim k As Long, ind_tmp As Long, double_step_size_threshold As Long

    For k = 0 To nb_subfr - 1
        If k = 0 And conditional = 0 Then
            '/* Gain index is not allowed to go down more than 16 steps (~21.8 dB) */
            prev_ind = silk_max_int(ind(k), prev_ind - 16)
        Else
            '/* Delta index */
            ind_tmp = ind(k) + MIN_DELTA_GAIN_QUANT

            '/* Accumulate deltas */
            double_step_size_threshold = 2 * MAX_DELTA_GAIN_QUANT - N_LEVELS_QGAIN + prev_ind
            If ind_tmp > double_step_size_threshold Then
                prev_ind = prev_ind + (ind_tmp * 2 - double_step_size_threshold)
            Else
                prev_ind = prev_ind + ind_tmp
            End If
        End If
        prev_ind = silk_LIMIT(prev_ind, 0, N_LEVELS_QGAIN - 1)

        '/* Scale and convert to linear scale */
        gain_Q16(k) = silk_log2lin(silk_min_32(silk_SMULWB(INV_SCALE_Q16, prev_ind) + OFFSET, 3967))   '/* 3967 = 31 in Q7 */
    Next
End Sub
