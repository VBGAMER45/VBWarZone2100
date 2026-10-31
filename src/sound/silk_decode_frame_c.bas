Attribute VB_Name = "silk_decode_frame_c"
' libopus 1.5.2 (ref/opus/silk): decode_frame.c (no DEEP_PLC, no OSCE).
' (VB) pOut is an array and an offset.
Option Explicit

' (VB) the stack buffer of pulses
Private pulses(0 To 319) As Integer                     ' ( L + SHELL_CODEC_FRAME_LENGTH - 1 ) & ~( SHELL_CODEC_FRAME_LENGTH - 1 )

'/****************/
'/* Decode frame */
'/****************/
Public Function silk_decode_frame(psDec As silk_decoder_state, psRangeDec As ec_ctx, pOut() As Integer, _
        ByVal pOutOffset As Long, pN As Long, ByVal lostFlag As Long, ByVal condCoding As Long, ByVal arch As Long) As Long
    Dim psDecCtrl As silk_decoder_control
    Dim L As Long, mv_len As Long, ret As Long
    Dim decodeIt As Boolean

    L = psDec.frame_length
    psDecCtrl.LTP_scale_Q14 = 0

    '/* Safety checks */
    ' celt_assert( L > 0 && L <= MAX_FRAME_LENGTH );

    If lostFlag = FLAG_DECODE_NORMAL Then
        decodeIt = True
    ElseIf lostFlag = FLAG_DECODE_LBRR Then
        decodeIt = (psDec.LBRR_flags(psDec.nFramesDecoded) = 1)
    End If
    If decodeIt Then
        '/*********************************************/
        '/* Decode quantization indices of side info  */
        '/*********************************************/
        silk_decode_indices psDec, psRangeDec, psDec.nFramesDecoded, lostFlag, condCoding

        '/*********************************************/
        '/* Decode quantization indices of excitation */
        '/*********************************************/
        silk_decode_pulses psRangeDec, pulses, psDec.indices.signalType, psDec.indices.quantOffsetType, psDec.frame_length

        '/********************************************/
        '/* Decode parameters and pulse signal       */
        '/********************************************/
        silk_decode_parameters psDec, psDecCtrl, condCoding

        '/********************************************************/
        '/* Run inverse NSQ                                      */
        '/********************************************************/
        silk_decode_core psDec, psDecCtrl, pOut, pOutOffset, pulses, arch

        '/*************************/
        '/* Update output buffer. */
        '/*************************/
        ' celt_assert( psDec->ltp_mem_length >= psDec->frame_length );
        mv_len = psDec.ltp_mem_length - psDec.frame_length
        silk_memcpy_i16 psDec.outBuf, 0, psDec.outBuf, psDec.frame_length, mv_len
        silk_memcpy_i16 psDec.outBuf, mv_len, pOut, pOutOffset, psDec.frame_length

        '/********************************************************/
        '/* Update PLC state                                     */
        '/********************************************************/
        silk_PLC psDec, psDecCtrl, pOut, pOutOffset, 0, arch
        psDec.lossCnt = 0
        psDec.prevSignalType = psDec.indices.signalType
        ' celt_assert( psDec->prevSignalType >= 0 && psDec->prevSignalType <= 2 );

        '/* A frame has been decoded without errors */
        psDec.first_frame_after_reset = 0
    Else
        '/* Handle packet loss by extrapolation */
        silk_PLC psDec, psDecCtrl, pOut, pOutOffset, 1, arch

        '/*************************/
        '/* Update output buffer. */
        '/*************************/
        ' celt_assert( psDec->ltp_mem_length >= psDec->frame_length );
        mv_len = psDec.ltp_mem_length - psDec.frame_length
        silk_memcpy_i16 psDec.outBuf, 0, psDec.outBuf, psDec.frame_length, mv_len
        silk_memcpy_i16 psDec.outBuf, mv_len, pOut, pOutOffset, psDec.frame_length
    End If

    '/************************************************/
    '/* Comfort noise generation / estimation        */
    '/************************************************/
    silk_CNG psDec, psDecCtrl, pOut, pOutOffset, L

    '/****************************************************************/
    '/* Ensure smooth connection of extrapolated and good frames     */
    '/****************************************************************/
    silk_PLC_glue_frames psDec, pOut, pOutOffset, L

    '/* Update some decoder state variables */
    psDec.lagPrev = psDecCtrl.pitchL(psDec.nb_subfr - 1)

    '/* Set output frame length */
    pN = L

    silk_decode_frame = ret
End Function
