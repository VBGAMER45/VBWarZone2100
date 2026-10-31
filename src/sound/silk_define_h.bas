Attribute VB_Name = "silk_define_h"
' libopus 1.5.2 (ref/opus/silk): define.h, errors.h and control.h's decoder flags (constants only; control.h's
' silk_DecControlStruct is in silk_structs_h). The floating-point constants are Doubles.
Option Explicit

' errors.h
'/**************************/
'/* Error messages         */
'/**************************/
Public Const SILK_NO_ERROR As Long = 0

'/**************************/
'/* Encoder error messages */
'/**************************/
'/* Input length is not a multiple of 10 ms, or length is longer than the packet length */
Public Const SILK_ENC_INPUT_INVALID_NO_OF_SAMPLES As Long = -101
'/* Sampling frequency not 8000, 12000 or 16000 Hertz */
Public Const SILK_ENC_FS_NOT_SUPPORTED As Long = -102
'/* Packet size not 10, 20, 40, or 60 ms */
Public Const SILK_ENC_PACKET_SIZE_NOT_SUPPORTED As Long = -103
'/* Allocated payload buffer too short */
Public Const SILK_ENC_PAYLOAD_BUF_TOO_SHORT As Long = -104
'/* Loss rate not between 0 and 100 percent */
Public Const SILK_ENC_INVALID_LOSS_RATE As Long = -105
'/* Complexity setting not valid, use 0...10 */
Public Const SILK_ENC_INVALID_COMPLEXITY_SETTING As Long = -106
'/* Inband FEC setting not valid, use 0 or 1 */
Public Const SILK_ENC_INVALID_INBAND_FEC_SETTING As Long = -107
'/* DTX setting not valid, use 0 or 1 */
Public Const SILK_ENC_INVALID_DTX_SETTING As Long = -108
'/* CBR setting not valid, use 0 or 1 */
Public Const SILK_ENC_INVALID_CBR_SETTING As Long = -109
'/* Internal encoder error */
Public Const SILK_ENC_INTERNAL_ERROR As Long = -110
'/* Internal encoder error */
Public Const SILK_ENC_INVALID_NUMBER_OF_CHANNELS_ERROR As Long = -111

'/**************************/
'/* Decoder error messages */
'/**************************/
'/* Output sampling frequency lower than internal decoded sampling frequency */
Public Const SILK_DEC_INVALID_SAMPLING_FREQUENCY As Long = -200
'/* Payload size exceeded the maximum allowed 1024 bytes */
Public Const SILK_DEC_PAYLOAD_TOO_LARGE As Long = -201
'/* Payload has bit errors */
Public Const SILK_DEC_PAYLOAD_ERROR As Long = -202
'/* Payload has bit errors */
Public Const SILK_DEC_INVALID_FRAME_SIZE As Long = -203

' control.h
'/* Decoder API flags */
Public Const FLAG_DECODE_NORMAL As Long = 0
Public Const FLAG_PACKET_LOST As Long = 1
Public Const FLAG_DECODE_LBRR As Long = 2

' define.h
'/* Max number of encoder channels (1/2) */
Public Const ENCODER_NUM_CHANNELS As Long = 2
'/* Number of decoder channels (1/2) */
Public Const DECODER_NUM_CHANNELS As Long = 2

Public Const MAX_FRAMES_PER_PACKET As Long = 3

'/* Limits on bitrate */
Public Const MIN_TARGET_RATE_BPS As Long = 5000
Public Const MAX_TARGET_RATE_BPS As Long = 80000

'/* LBRR thresholds */
Public Const LBRR_NB_MIN_RATE_BPS As Long = 12000
Public Const LBRR_MB_MIN_RATE_BPS As Long = 14000
Public Const LBRR_WB_MIN_RATE_BPS As Long = 16000

'/* DTX settings */
Public Const NB_SPEECH_FRAMES_BEFORE_DTX As Long = 10       '/* eq 200 ms */
Public Const MAX_CONSECUTIVE_DTX As Long = 20               '/* eq 400 ms */
Public Const DTX_ACTIVITY_THRESHOLD As Double = 0.1

'/* VAD decision */
Public Const VAD_NO_DECISION As Long = -1
Public Const VAD_NO_ACTIVITY As Long = 0
Public Const VAD_ACTIVITY As Long = 1

'/* Maximum sampling frequency */
Public Const MAX_FS_KHZ As Long = 16
Public Const MAX_API_FS_KHZ As Long = 48

'/* Signal types */
Public Const TYPE_NO_VOICE_ACTIVITY As Long = 0
Public Const TYPE_UNVOICED As Long = 1
Public Const TYPE_VOICED As Long = 2

'/* Conditional coding types */
Public Const CODE_INDEPENDENTLY As Long = 0
Public Const CODE_INDEPENDENTLY_NO_LTP_SCALING As Long = 1
Public Const CODE_CONDITIONALLY As Long = 2

'/* Settings for stereo processing */
Public Const STEREO_QUANT_TAB_SIZE As Long = 16
Public Const STEREO_QUANT_SUB_STEPS As Long = 5
Public Const STEREO_INTERP_LEN_MS As Long = 8               '/* must be even */
Public Const STEREO_RATIO_SMOOTH_COEF As Double = 0.01      '/* smoothing coef for signal norms and stereo width */

'/* Range of pitch lag estimates */
Public Const PITCH_EST_MIN_LAG_MS As Long = 2               '/* 2 ms -> 500 Hz */
Public Const PITCH_EST_MAX_LAG_MS As Long = 18              '/* 18 ms -> 56 Hz */

'/* Maximum number of subframes */
Public Const MAX_NB_SUBFR As Long = 4

'/* Number of samples per frame */
Public Const LTP_MEM_LENGTH_MS As Long = 20
Public Const SUB_FRAME_LENGTH_MS As Long = 5
Public Const MAX_SUB_FRAME_LENGTH As Long = 80              ' ( SUB_FRAME_LENGTH_MS * MAX_FS_KHZ )
Public Const MAX_FRAME_LENGTH_MS As Long = 20               ' ( SUB_FRAME_LENGTH_MS * MAX_NB_SUBFR )
Public Const MAX_FRAME_LENGTH As Long = 320                 ' ( MAX_FRAME_LENGTH_MS * MAX_FS_KHZ )

'/* Milliseconds of lookahead for pitch analysis */
Public Const LA_PITCH_MS As Long = 2
Public Const LA_PITCH_MAX As Long = 32                      ' ( LA_PITCH_MS * MAX_FS_KHZ )

'/* Order of LPC used in find pitch */
Public Const MAX_FIND_PITCH_LPC_ORDER As Long = 16

'/* Length of LPC window used in find pitch */
Public Const FIND_PITCH_LPC_WIN_MS As Long = 24             ' ( 20 + (LA_PITCH_MS << 1) )
Public Const FIND_PITCH_LPC_WIN_MS_2_SF As Long = 14        ' ( 10 + (LA_PITCH_MS << 1) )
Public Const FIND_PITCH_LPC_WIN_MAX As Long = 384           ' ( FIND_PITCH_LPC_WIN_MS * MAX_FS_KHZ )

'/* Milliseconds of lookahead for noise shape analysis */
Public Const LA_SHAPE_MS As Long = 5
Public Const LA_SHAPE_MAX As Long = 80                      ' ( LA_SHAPE_MS * MAX_FS_KHZ )

'/* Maximum length of LPC window used in noise shape analysis */
Public Const SHAPE_LPC_WIN_MAX As Long = 240                ' ( 15 * MAX_FS_KHZ )

'/* dB level of lowest gain quantization level */
Public Const MIN_QGAIN_DB As Long = 2
'/* dB level of highest gain quantization level */
Public Const MAX_QGAIN_DB As Long = 88
'/* Number of gain quantization levels */
Public Const N_LEVELS_QGAIN As Long = 64
'/* Max increase in gain quantization index */
Public Const MAX_DELTA_GAIN_QUANT As Long = 36
'/* Max decrease in gain quantization index */
Public Const MIN_DELTA_GAIN_QUANT As Long = -4

'/* Quantization offsets (multiples of 4) */
Public Const OFFSET_VL_Q10 As Long = 32
Public Const OFFSET_VH_Q10 As Long = 100
Public Const OFFSET_UVL_Q10 As Long = 100
Public Const OFFSET_UVH_Q10 As Long = 240

Public Const QUANT_LEVEL_ADJUST_Q10 As Long = 80

'/* Maximum numbers of iterations used to stabilize an LPC vector */
Public Const MAX_LPC_STABILIZE_ITERATIONS As Long = 16
Public Const MAX_PREDICTION_POWER_GAIN As Double = 10000#
Public Const MAX_PREDICTION_POWER_GAIN_AFTER_RESET As Double = 100#

Public Const MAX_LPC_ORDER As Long = 16
Public Const MIN_LPC_ORDER As Long = 10

'/* Find Pred Coef defines */
Public Const LTP_ORDER As Long = 5

'/* LTP quantization settings */
Public Const NB_LTP_CBKS As Long = 3

'/* Flag to use harmonic noise shaping */
Public Const USE_HARM_SHAPING As Long = 1

'/* Max LPC order of noise shaping filters */
Public Const MAX_SHAPE_LPC_ORDER As Long = 24

Public Const HARM_SHAPE_FIR_TAPS As Long = 3

'/* Maximum number of delayed decision states */
Public Const MAX_DEL_DEC_STATES As Long = 4

Public Const LTP_BUF_LENGTH As Long = 512
Public Const LTP_MASK As Long = 511                         ' ( LTP_BUF_LENGTH - 1 )

Public Const DECISION_DELAY As Long = 40

'/* Number of subframes for excitation entropy coding */
Public Const SHELL_CODEC_FRAME_LENGTH As Long = 16
Public Const LOG2_SHELL_CODEC_FRAME_LENGTH As Long = 4
Public Const MAX_NB_SHELL_BLOCKS As Long = 20               ' ( MAX_FRAME_LENGTH / SHELL_CODEC_FRAME_LENGTH )

'/* Number of rate levels, for entropy coding of excitation */
Public Const N_RATE_LEVELS As Long = 10

'/* Maximum sum of pulses per shell coding frame */
Public Const SILK_MAX_PULSES As Long = 16

Public Const MAX_MATRIX_SIZE As Long = 16                   ' MAX_LPC_ORDER /* Max of LPC Order and LTP order */

Public Const NSQ_LPC_BUF_LENGTH As Long = 16                ' MAX_LPC_ORDER

'/***************************/
'/* Voice activity detector */
'/***************************/
Public Const VAD_N_BANDS As Long = 4

Public Const VAD_INTERNAL_SUBFRAMES_LOG2 As Long = 2
Public Const VAD_INTERNAL_SUBFRAMES As Long = 4             ' ( 1 << VAD_INTERNAL_SUBFRAMES_LOG2 )

Public Const VAD_NOISE_LEVEL_SMOOTH_COEF_Q16 As Long = 1024 '/* Must be <  4096 */
Public Const VAD_NOISE_LEVELS_BIAS As Long = 50

'/* Sigmoid settings */
Public Const VAD_NEGATIVE_OFFSET_Q5 As Long = 128           '/* sigmoid is 0 at -128 */
Public Const VAD_SNR_FACTOR_Q16 As Long = 45000

'/* smoothing for SNR measurement */
Public Const VAD_SNR_SMOOTH_COEF_Q18 As Long = 4096

'/* Size of the piecewise linear cosine approximation table for the LSFs */
Public Const LSF_COS_TAB_SZ_FIX As Long = 128

'/******************/
'/* NLSF quantizer */
'/******************/
Public Const NLSF_W_Q As Long = 2
Public Const NLSF_VQ_MAX_VECTORS As Long = 32
Public Const NLSF_QUANT_MAX_AMPLITUDE As Long = 4
Public Const NLSF_QUANT_MAX_AMPLITUDE_EXT As Long = 10
Public Const NLSF_QUANT_LEVEL_ADJ As Double = 0.1
Public Const NLSF_QUANT_DEL_DEC_STATES_LOG2 As Long = 2
Public Const NLSF_QUANT_DEL_DEC_STATES As Long = 4          ' ( 1 << NLSF_QUANT_DEL_DEC_STATES_LOG2 )

'/* Transition filtering for mode switching */
Public Const TRANSITION_TIME_MS As Long = 5120              '/* 5120 = 64 * FRAME_LENGTH_MS * ( TRANSITION_INT_NUM - 1 ) = 64*(20*4)*/
Public Const TRANSITION_NB As Long = 3                      '/* Hardcoded in tables */
Public Const TRANSITION_NA As Long = 2                      '/* Hardcoded in tables */
Public Const TRANSITION_INT_NUM As Long = 5                 '/* Hardcoded in tables */
Public Const TRANSITION_FRAMES As Long = 256                ' ( TRANSITION_TIME_MS / MAX_FRAME_LENGTH_MS )
Public Const TRANSITION_INT_STEPS As Long = 64              ' ( TRANSITION_FRAMES  / ( TRANSITION_INT_NUM - 1 ) )

'/* BWE factors to apply after packet loss */
Public Const BWE_AFTER_LOSS_Q16 As Long = 63570

'/* Defines for CN generation */
Public Const CNG_BUF_MASK_MAX As Long = 255                 '/* 2^floor(log2(MAX_FRAME_LENGTH))-1    */
Public Const CNG_GAIN_SMTH_Q16 As Long = 4634               '/* 0.25^(1/4)                           */
Public Const CNG_GAIN_SMTH_THRESHOLD_Q16 As Long = 46396    '/* -3 dB                                */
Public Const CNG_NLSF_SMTH_Q16 As Long = 16348              '/* 0.25                                 */
