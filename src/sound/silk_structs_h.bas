Attribute VB_Name = "silk_structs_h"
' libopus 1.5.2 (ref/opus/silk): structs.h's decoder types (the encoder's silk_nsq_state, silk_VAD_state,
' silk_LP_state, stereo_enc_state and silk_encoder_state are not ported), resampler_structs.h's
' silk_resampler_state_struct and control.h's silk_DecControlStruct. No OSCE (ENABLE_OSCE undefined), no
' DEEP_PLC.
' (VB) opus_int16 fields are Integers (a store keeps the low 16 bits, as C's does); opus_int8 fields are Longs
' (the decoder's int8 values are small: the range decoder's symbols and the gain index, clamped to 0..63).
' A const pointer to a table is an offset into its ROM (silk_tables_c); psNLSF_CB is an index of
' silk_NLSF_CBs (0 = NULL). A 2-D array is one dimension, row by row as C lays it out.
Option Explicit

' resampler_structs.h
Public Const SILK_RESAMPLER_MAX_FIR_ORDER As Long = 36
Public Const SILK_RESAMPLER_MAX_IIR_ORDER As Long = 6

Public Type silk_resampler_state_struct
    sIIR(0 To 5) As Long                    ' [ SILK_RESAMPLER_MAX_IIR_ORDER ] /* this must be the first element of this struct */
    sFIR_i32(0 To 35) As Long               ' union sFIR: i32[ SILK_RESAMPLER_MAX_FIR_ORDER ] (down_FIR's)
    sFIR_i16(0 To 35) As Integer            ' union sFIR: i16[ SILK_RESAMPLER_MAX_FIR_ORDER ] (IIR_FIR's; (VB) beside i32, as one resampler uses one of them)
    delayBuf(0 To 47) As Integer
    resampler_function As Long
    batchSize As Long
    invRatio_Q16 As Long
    FIR_Order As Long
    FIR_Fracs As Long
    Fs_in_kHz As Long
    Fs_out_kHz As Long
    inputDelay As Long
    Coefs As Long                           ' const opus_int16 *: an offset into silk_rom_i16
End Type

' structs.h
'/* Structure containing NLSF codebook */
Public Type silk_NLSF_CB_struct
    nVectors As Long                        ' const opus_int16
    order As Long                           ' const opus_int16
    quantStepSize_Q16 As Long               ' const opus_int16
    invQuantStepSize_Q6 As Long             ' const opus_int16
    CB1_NLSF_Q8 As Long                     ' const opus_uint8 *  (silk_rom_u8)
    CB1_Wght_Q9 As Long                     ' const opus_int16 *  (silk_rom_i16)
    CB1_iCDF As Long                        ' const opus_uint8 *
    pred_Q8 As Long                         ' const opus_uint8 *
    ec_sel As Long                          ' const opus_uint8 *
    ec_iCDF As Long                         ' const opus_uint8 *
    ec_Rates_Q5 As Long                     ' const opus_uint8 *
    deltaMin_Q15 As Long                    ' const opus_int16 *  (silk_rom_i16)
End Type

Public Type stereo_dec_state
    pred_prev_Q13(0 To 1) As Integer
    sMid(0 To 1) As Integer
    sSide(0 To 1) As Integer
End Type

Public Type SideInfoIndices
    GainsIndices(0 To 3) As Long            ' opus_int8 [ MAX_NB_SUBFR ]
    LTPIndex(0 To 3) As Long                ' opus_int8 [ MAX_NB_SUBFR ]
    NLSFIndices(0 To 16) As Long            ' opus_int8 [ MAX_LPC_ORDER + 1 ]
    lagIndex As Integer                     ' opus_int16
    contourIndex As Long                    ' opus_int8
    signalType As Long                      ' opus_int8
    quantOffsetType As Long                 ' opus_int8
    NLSFInterpCoef_Q2 As Long               ' opus_int8
    PERIndex As Long                        ' opus_int8
    LTP_scaleIndex As Long                  ' opus_int8
    Seed As Long                            ' opus_int8
End Type

'/* Struct for Packet Loss Concealment */
Public Type silk_PLC_struct
    pitchL_Q8 As Long                       '/* Pitch lag to use for voiced concealment                          */
    LTPCoef_Q14(0 To 4) As Integer          '/* LTP coeficients to use for voiced concealment                    */
    prevLPC_Q12(0 To 15) As Integer
    last_frame_lost As Long                 '/* Was previous frame lost                                          */
    rand_seed As Long                       '/* Seed for unvoiced signal generation                              */
    randScale_Q14 As Integer                '/* Scaling of unvoiced random signal                                */
    conc_energy As Long
    conc_energy_shift As Long
    prevLTP_scale_Q14 As Integer
    prevGain_Q16(0 To 1) As Long
    fs_kHz As Long
    nb_subfr As Long
    subfr_length As Long
    enable_deep_plc As Long
End Type

'/* Struct for CNG */
Public Type silk_CNG_struct
    CNG_exc_buf_Q14(0 To 319) As Long       ' [ MAX_FRAME_LENGTH ]
    CNG_smth_NLSF_Q15(0 To 15) As Integer   ' [ MAX_LPC_ORDER ]
    CNG_synth_state(0 To 15) As Long        ' [ MAX_LPC_ORDER ]
    CNG_smth_Gain_Q16 As Long
    rand_seed As Long
    fs_kHz As Long
End Type

'/********************************/
'/* Decoder state                */
'/********************************/
Public Type silk_decoder_state
    ' #define SILK_DECODER_STATE_RESET_START prev_gain_Q16 (the first field: no OSCE)
    prev_gain_Q16 As Long
    exc_Q14(0 To 319) As Long               ' [ MAX_FRAME_LENGTH ]
    sLPC_Q14_buf(0 To 15) As Long           ' [ MAX_LPC_ORDER ]
    outBuf(0 To 479) As Integer             ' [ MAX_FRAME_LENGTH + 2 * MAX_SUB_FRAME_LENGTH ] /* Buffer for output signal                     */
    lagPrev As Long                         '/* Previous Lag                                                     */
    LastGainIndex As Long                   ' opus_int8 /* Previous gain index                                  */
    fs_kHz As Long                          '/* Sampling frequency in kHz                                        */
    fs_API_hz As Long                       '/* API sample frequency (Hz)                                        */
    nb_subfr As Long                        '/* Number of 5 ms subframes in a frame                              */
    frame_length As Long                    '/* Frame length (samples)                                           */
    subfr_length As Long                    '/* Subframe length (samples)                                        */
    ltp_mem_length As Long                  '/* Length of LTP memory                                             */
    LPC_order As Long                       '/* LPC order                                                        */
    prevNLSF_Q15(0 To 15) As Integer        '/* Used to interpolate LSFs                                         */
    first_frame_after_reset As Long         '/* Flag for deactivating NLSF interpolation                         */
    pitch_lag_low_bits_iCDF As Long         ' const opus_uint8 * (silk_rom_u8) /* Pointer to iCDF table for low bits of pitch lag index */
    pitch_contour_iCDF As Long              ' const opus_uint8 * (silk_rom_u8) /* Pointer to iCDF table for pitch contour index */

    '/* For buffering payload in case of more frames per packet */
    nFramesDecoded As Long
    nFramesPerPacket As Long

    '/* Specifically for entropy coding */
    ec_prevSignalType As Long
    ec_prevLagIndex As Integer

    VAD_flags(0 To 2) As Long               ' [ MAX_FRAMES_PER_PACKET ]
    LBRR_flag As Long
    LBRR_flags(0 To 2) As Long              ' [ MAX_FRAMES_PER_PACKET ]

    resampler_state As silk_resampler_state_struct

    psNLSF_CB As Long                       ' const silk_NLSF_CB_struct *: an index of silk_NLSF_CBs /* Pointer to NLSF codebook */

    '/* Quantization indices */
    indices As SideInfoIndices

    '/* CNG state */
    sCNG As silk_CNG_struct

    '/* Stuff used for PLC */
    lossCnt As Long
    prevSignalType As Long
    arch As Long

    sPLC As silk_PLC_struct
End Type

'/************************/
'/* Decoder control      */
'/************************/
Public Type silk_decoder_control
    '/* Prediction and coding parameters */
    pitchL(0 To 3) As Long                  ' [ MAX_NB_SUBFR ]
    Gains_Q16(0 To 3) As Long               ' [ MAX_NB_SUBFR ]
    '/* Holds interpolated and final coefficients, 4-byte aligned */
    PredCoef_Q12(0 To 31) As Integer        ' [ 2 ][ MAX_LPC_ORDER ]: row k from k * MAX_LPC_ORDER
    LTPCoef_Q14(0 To 19) As Integer         ' [ LTP_ORDER * MAX_NB_SUBFR ]
    LTP_scale_Q14 As Long
End Type

' control.h
'/**************************************************************************/
'/* Structure for controlling decoder operation and reading decoder status */
'/**************************************************************************/
Public Type silk_DecControlStruct
    '/* I:   Number of channels; 1/2                                                         */
    nChannelsAPI As Long
    '/* I:   Number of channels; 1/2                                                         */
    nChannelsInternal As Long
    '/* I:   Output signal sampling rate in Hertz; 8000/12000/16000/24000/32000/44100/48000  */
    API_sampleRate As Long
    '/* I:   Internal sampling rate used, in Hertz; 8000/12000/16000                         */
    internalSampleRate As Long
    '/* I:   Number of samples per packet in milliseconds; 10/20/40/60                       */
    payloadSize_ms As Long
    '/* O:   Pitch lag of previous frame (0 if unvoiced), measured in samples at 48 kHz      */
    prevPitchLag As Long
    '/* I:   Enable Deep PLC                                                                 */
    enable_deep_plc As Long
End Type
