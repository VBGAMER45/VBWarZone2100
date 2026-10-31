Attribute VB_Name = "silk_init_decoder_c"
' libopus 1.5.2 (ref/opus/silk): init_decoder.c (no OSCE).
Option Explicit

'/************************/
'/* Reset Decoder State  */
'/************************/
Public Function silk_reset_decoder(psDec As silk_decoder_state) As Long  '/* I/O  Decoder state pointer                       */
    Static blank As silk_decoder_state
    '/* Clear the entire encoder state, except anything copied */
    ' (VB) from SILK_DECODER_STATE_RESET_START (prev_gain_Q16) to the end: without OSCE, the whole state
    psDec = blank

    '/* Used to deactivate LSF interpolation */
    psDec.first_frame_after_reset = 1
    psDec.prev_gain_Q16 = 65536
    psDec.arch = 0                          ' opus_select_arch(): the C paths only

    '/* Reset CNG state */
    silk_CNG_Reset psDec

    '/* Reset PLC state */
    silk_PLC_Reset psDec

    silk_reset_decoder = 0
End Function

'/************************/
'/* Init Decoder State   */
'/************************/
Public Function silk_init_decoder(psDec As silk_decoder_state) As Long   '/* I/O  Decoder state pointer                       */
    Static blank As silk_decoder_state
    '/* Clear the entire encoder state, except anything copied */
    psDec = blank

    silk_reset_decoder psDec
    silk_init_decoder = 0
End Function
