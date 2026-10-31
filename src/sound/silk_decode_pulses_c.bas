Attribute VB_Name = "silk_decode_pulses_c"
' libopus 1.5.2 (ref/opus/silk): decode_pulses.c.
Option Explicit

'/*********************************************/
'/* Decode quantization indices of excitation */
'/*********************************************/
Public Sub silk_decode_pulses(psRangeDec As ec_ctx, pulses() As Integer, ByVal signalType As Long, _
        ByVal quantOffsetType As Long, ByVal frame_length As Long)
    Dim i As Long, j As Long, k As Long, iter As Long, abs_q As Long, nLS As Long, RateLevelIndex As Long
    Dim sum_pulses(0 To 19) As Long                     ' [ MAX_NB_SHELL_BLOCKS ]
    Dim nLshifts(0 To 19) As Long                       ' [ MAX_NB_SHELL_BLOCKS ]
    Dim pulses_ptr As Long                              ' opus_int16 *: an offset into pulses
    Dim cdf_ptr As Long                                 ' const opus_uint8 *: an offset into silk_rom_u8

    '/*********************/
    '/* Decode rate level */
    '/*********************/
    RateLevelIndex = ec_dec_icdf(psRangeDec, silk_rom_u8, silk_rate_levels_iCDF + silk_RSHIFT(signalType, 1) * (N_RATE_LEVELS - 1), 8)

    '/* Calculate number of shell blocks */
    ' silk_assert( 1 << LOG2_SHELL_CODEC_FRAME_LENGTH == SHELL_CODEC_FRAME_LENGTH );
    iter = silk_RSHIFT(frame_length, LOG2_SHELL_CODEC_FRAME_LENGTH)
    If iter * SHELL_CODEC_FRAME_LENGTH < frame_length Then
        ' celt_assert( frame_length == 12 * 10 ); /* Make sure only happens for 10 ms @ 12 kHz */
        iter = iter + 1
    End If

    '/***************************************************/
    '/* Sum-Weighted-Pulses Decoding                    */
    '/***************************************************/
    cdf_ptr = silk_pulses_per_block_iCDF + RateLevelIndex * (SILK_MAX_PULSES + 2)
    For i = 0 To iter - 1
        nLshifts(i) = 0
        sum_pulses(i) = ec_dec_icdf(psRangeDec, silk_rom_u8, cdf_ptr, 8)

        '/* LSB indication */
        Do While sum_pulses(i) = SILK_MAX_PULSES + 1
            nLshifts(i) = nLshifts(i) + 1
            '/* When we've already got 10 LSBs, we shift the table to not allow (SILK_MAX_PULSES + 1) */
            If nLshifts(i) = 10 Then
                sum_pulses(i) = ec_dec_icdf(psRangeDec, silk_rom_u8, silk_pulses_per_block_iCDF + (N_RATE_LEVELS - 1) * (SILK_MAX_PULSES + 2) + 1, 8)
            Else
                sum_pulses(i) = ec_dec_icdf(psRangeDec, silk_rom_u8, silk_pulses_per_block_iCDF + (N_RATE_LEVELS - 1) * (SILK_MAX_PULSES + 2), 8)
            End If
        Loop
    Next

    '/***************************************************/
    '/* Shell decoding                                  */
    '/***************************************************/
    For i = 0 To iter - 1
        If sum_pulses(i) > 0 Then
            silk_shell_decoder pulses, silk_SMULBB(i, SHELL_CODEC_FRAME_LENGTH), psRangeDec, sum_pulses(i)
        Else
            silk_memset_i16 pulses, silk_SMULBB(i, SHELL_CODEC_FRAME_LENGTH), SHELL_CODEC_FRAME_LENGTH
        End If
    Next

    '/***************************************************/
    '/* LSB Decoding                                    */
    '/***************************************************/
    For i = 0 To iter - 1
        If nLshifts(i) > 0 Then
            nLS = nLshifts(i)
            pulses_ptr = silk_SMULBB(i, SHELL_CODEC_FRAME_LENGTH)
            For k = 0 To SHELL_CODEC_FRAME_LENGTH - 1
                abs_q = pulses(pulses_ptr + k)
                For j = 0 To nLS - 1
                    abs_q = abs_q * 2
                    abs_q = abs_q + ec_dec_icdf(psRangeDec, silk_rom_u8, silk_lsb_iCDF, 8)
                Next
                pulses(pulses_ptr + k) = abs_q
            Next
            '/* Mark the number of pulses non-zero for sign decoding. */
            sum_pulses(i) = sum_pulses(i) Or (nLS * 32)
        End If
    Next

    '/****************************************/
    '/* Decode and add signs to pulse signal */
    '/****************************************/
    silk_decode_signs psRangeDec, pulses, frame_length, signalType, quantOffsetType, sum_pulses
End Sub
